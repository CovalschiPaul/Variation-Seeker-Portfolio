# Variation Seeker · Portfolio Demo

**Turn two catalog relationship exports into one auditable Excel workbook.**

A Python data-quality tool. This repository demonstrates the parsing, validation and reconciliation portion of a catalog-analysis workflow using **entirely synthetic data**. It works offline and contains no production service connections, authentication code or company records.

## Try it in three commands

Python 3.10–3.13. On Windows, Python's standard installer includes Tkinter for the desktop window.

```bash
python -m pip install -r requirements.txt
python Variation_Seeker.py --demo --output demo_output.xlsx
python Variation_Seeker.py
```

The second command runs without a display. The third opens the desktop app with the sample files preselected. On Linux, the desktop window may require your distribution's `python3-tk` package.

## What the sample demonstrates

| Scenario | Expected result |
|---|---|
| Same parent and child in both sources | Both sources |
| Relationship present only in the first export | CSI only |
| Relationship present only in the second export | RTIP only |
| Repeated identical source row | Deduplicated |
| Same key with conflicting attributes | One comparison row, joined distinct values, review flag |
| Numeric identifier beginning with zeros | Preserved as text |

The sample produces **six relationship pairs**: three in both sources, two CSI-only, and one RTIP-only. One pair is flagged for conflicting source rows. CSI and RTIP are labels for the two input formats here, not live integrations.

## Workbook

- **CSI Output / RTIP_Output** — normalized source records, retaining conflicting rows.
- **Final Data** — explicit parent-and-child outer join, source-prefixed attributes, match status, source row counts and review flags.
- **Errors** — consistent error-report schema (empty for successful offline runs).
- **Run Info** — mode and interpretation notes.

Source-only means absent from the supplied export, not proof of a missing relationship in a live system. `Both sources` means the keys match; it does not assert that every attribute agrees. Distinct values are joined with ` | `; the raw sheets remain authoritative.

## Your own local files

```bash
python Variation_Seeker.py --csi source_a.csv --rtip source_b.xlsx --output comparison.xlsx --log run.log
```

Accepted inputs: UTF-8 CSV, XLSX (first sheet), or HTML with a `productdata` / `render` table. Each input needs parent and child identifiers. Accepted headers include `Parent Item`, `Parent ASIN`, `Parent`, and corresponding child forms. Identifiers must contain exactly ten ASCII letters/digits. Store spreadsheet identifier cells as **text**; zeros already lost by Excel cannot be recovered.

Existing output files are protected unless `--overwrite` is supplied (the GUI asks before replacing). Input files cannot be used as the output. Exported text is stored as text, including formula-like values. Excel row and cell-length limits are checked before writing.

## Engineering choices

Small, readable modules: `variation_core.py` for pure data logic and safe workbook writing; `Variation_Seeker.py` for the CLI and responsive Tkinter window. Duplicate keys are summarized before a validated one-to-one merge, preventing accidental Cartesian expansion. Exports are written to a temporary file before finalization.

```bash
python -m unittest discover -s tests -v
```

Tests cover malformed and empty inputs, duplicate keys, leading zeros, formula-like cells, overwrite protection, HTML parsing and a 10,000-record comparison. This is a portfolio demonstration, not a claim of production validation or a measured business-impact benchmark.

## Privacy

Only fictional examples belong in this repository. Local reports, logs, environments and configuration files are ignored by Git. The private application is maintained separately; this project has an independent history.
