"""Offline parsing, reconciliation and safe Excel export. No authentication code."""
from __future__ import annotations
import io
import os
from pathlib import Path
import re
import tempfile
from dataclasses import dataclass, field
import pandas as pd
from lxml import html as lxml_html
from openpyxl.styles import Font, PatternFill

KEYS = ['Parent', 'Child']
ERROR_COLUMNS = ['Source', 'Input ASIN', 'Page', 'Error']
ALIASES = {'Parent Item': 'Parent', 'Parent ASIN': 'Parent', 'Child Item': 'Child', 'Child ASIN': 'Child'}
ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')

class DataError(ValueError):
    pass

class Cancelled(Exception):
    pass

@dataclass
class Result:
    csi: pd.DataFrame = field(default_factory=pd.DataFrame)
    rtip: pd.DataFrame = field(default_factory=pd.DataFrame)
    errors: list[dict] = field(default_factory=list)
    metadata: dict = field(default_factory=dict)

    def error(self, source, asin, page, message):
        self.errors.append(dict(zip(ERROR_COLUMNS, [source, asin, page, str(message)])))


def parse_asins(text):
    tokens = [s.upper() for s in re.split(r'[\s,;]+', text.strip()) if s]
    if not tokens:
        raise DataError('Enter at least one 10-character ASIN.')
    bad = [s for s in tokens if not re.fullmatch(r'[A-Z0-9]{10}', s)]
    if bad:
        raise DataError('Invalid ASIN(s): ' + ', '.join(bad[:8]))
    return list(dict.fromkeys(tokens))


def clean_frame(frame):
    frame = frame.copy().fillna('')
    frame.columns = [' '.join(str(c).split()) for c in frame.columns]
    frame = frame.rename(columns=ALIASES)
    if frame.columns.duplicated().any():
        raise DataError('Duplicate column names after normalization.')
    for col in frame:
        frame[col] = frame[col].map(lambda x: ILLEGAL.sub('', str(x)).replace('more tools', '').strip())
    return frame


def normalize(frame, source):
    if frame.empty and not len(frame.columns):
        return pd.DataFrame(columns=KEYS)
    frame = clean_frame(frame)
    missing = [c for c in KEYS if c not in frame]
    if missing:
        raise DataError(f'{source}: missing columns: {", ".join(missing)}')
    # Remove repeated header rows; keep malformed records visible as an error.
    headers = frame['Parent'].isin(['Parent', 'Parent Item', 'Parent ASIN'])
    frame = frame.loc[~headers].copy()
    frame = frame.loc[~frame.eq('').all(axis=1)].copy()
    for key in KEYS:
        frame[key] = frame[key].str.upper()
        if not frame[key].str.fullmatch(r'[A-Z0-9]{10}').all():
            raise DataError(f'{source}: blank or invalid {key} identifier in table.')
    return frame.drop(columns=['#', 'index'], errors='ignore').drop_duplicates().reset_index(drop=True)


def parse_table(html, table_id, source):
    try:
        root = lxml_html.fromstring(html)
    except (ValueError, lxml_html.etree.ParserError) as exc:
        raise DataError(f'{source}: invalid or empty HTML.') from exc
    tables = root.xpath('//table[@id=$table_id]', table_id=table_id)
    if not tables:
        raise DataError(f'{source}: expected table {table_id!r} not found; check login or page format.')
    try:
        serialized = lxml_html.tostring(tables[0], encoding='unicode')
        schema = pd.read_html(io.StringIO(serialized), keep_default_na=False, flavor='lxml')[0]
        frames = pd.read_html(io.StringIO(serialized), keep_default_na=False,
                              converters={i: str for i in range(len(schema.columns))}, flavor='lxml')
    except (ValueError, IndexError) as exc:
        raise DataError(f'{source}: table could not be read.') from exc
    return normalize(frames[0], source)


def load_table(path, source):
    path = Path(path)
    if path.suffix.lower() == '.csv':
        frame = pd.read_csv(path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
    elif path.suffix.lower() == '.xlsx':
        frame = pd.read_excel(path, dtype=str, keep_default_na=False)
    elif path.suffix.lower() in ('.html', '.htm'):
        return parse_table(path.read_text(encoding='utf-8'), 'productdata' if source == 'CSI' else 'render', source)
    else:
        raise DataError('Supported input types: CSV, XLSX, HTML.')
    return normalize(frame, source)


def summarize(frame, source):
    frame = normalize(frame, source)
    columns = [c for c in frame if c not in KEYS]
    if frame.empty:
        return pd.DataFrame(columns=KEYS + [f'{source}_{c}' for c in columns] + [f'{source}_Rows'])
    grouped = frame.groupby(KEYS, sort=False, dropna=False)
    counts = grouped.size().rename(f'{source}_Rows')
    if columns:
        # Keep every distinct value while preventing a many-to-many row explosion.
        values = grouped[columns].agg(lambda s: ' | '.join(dict.fromkeys(v for v in s if v)))
        values.columns = [f'{source}_{c}' for c in columns]
        return values.join(counts).reset_index()
    return counts.reset_index()


def reconcile(csi, rtip):
    left, right = summarize(csi, 'CSI'), summarize(rtip, 'RTIP')
    merged = left.merge(right, on=KEYS, how='outer', validate='one_to_one', indicator=True)
    merged['Match Status'] = merged.pop('_merge').map({'both': 'Both sources', 'left_only': 'CSI only', 'right_only': 'RTIP only'}).astype(str)
    for source in ('CSI', 'RTIP'):
        merged[f'{source}_Rows'] = pd.to_numeric(merged[f'{source}_Rows'], errors='raise').fillna(0).astype(int)
    merged['Review Required'] = (merged['CSI_Rows'].gt(1) | merged['RTIP_Rows'].gt(1)).map({True: 'Multiple source rows', False: ''})
    return merged.fillna('').sort_values(KEYS).reset_index(drop=True)


def export_result(result, output, overwrite=False):
    output = Path(output).expanduser().resolve()
    if output.suffix.lower() != '.xlsx':
        raise DataError('Output must have the .xlsx extension.')
    if output.exists() and not overwrite:
        raise DataError(f'Output already exists: {output}. Choose another path or allow overwrite.')
    frames = {'CSI Output': normalize(result.csi, 'CSI'), 'RTIP_Output': normalize(result.rtip, 'RTIP'),
              'Final Data': reconcile(result.csi, result.rtip),
              'Errors': pd.DataFrame(result.errors, columns=ERROR_COLUMNS),
              'Run Info': pd.DataFrame(list(result.metadata.items()) + [('Error count', len(result.errors)),
                 ('Interpretation', 'Source-only means absent from retrieved data, not proof of a missing live relationship.')], columns=['Setting', 'Value'])}
    for name, frame in frames.items():
        if len(frame) > 1_048_575 or len(frame.columns) > 16_384:
            raise DataError(f'{name} exceeds Excel worksheet limits. Split the input batch.')
        if any(frame[c].map(lambda v: len(str(v)) > 32767).any() for c in frame):
            raise DataError(f'{name} contains a value longer than Excel allows (32767 characters).')
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix='.variation-', suffix='.xlsx', dir=output.parent)
    os.close(fd)
    try:
        with pd.ExcelWriter(temporary, engine='openpyxl') as writer:
            for name, frame in frames.items():
                frame.to_excel(writer, sheet_name=name, index=False)
                sheet = writer.sheets[name]
                sheet.freeze_panes = 'A2'
                sheet.auto_filter.ref = sheet.dimensions
                for row in sheet:
                    for cell in row:
                        if isinstance(cell.value, str):
                            cell.value = ILLEGAL.sub('', cell.value)
                            cell.data_type = 's'  # Untrusted text is never an Excel formula.
                for cell in sheet[1]:
                    cell.fill = PatternFill('solid', fgColor='17365D')
                    cell.font = Font(color='FFFFFF', bold=True)
                for column in sheet.columns:
                    width = max(len(str(c.value or '')) for c in list(column)[:101])
                    sheet.column_dimensions[column[0].column_letter].width = min(60, max(14, width + 2))
        if overwrite:
            os.replace(temporary, output)
        else:
            # Atomic creation without silently replacing a file created during export.
            os.link(temporary, output)
            os.unlink(temporary)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
    return output
