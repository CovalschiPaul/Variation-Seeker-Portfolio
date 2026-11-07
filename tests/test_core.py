import tempfile
import unittest
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from variation_core import DataError, Result, export_result, load_table, normalize, parse_asins, parse_table, reconcile

P='0000000001'
C='0000000002'

def table(rows):
    return pd.DataFrame(rows,columns=['Parent','Child','Value'])

class CoreTests(unittest.TestCase):
    def test_asins(self):
        self.assertEqual(parse_asins(' 0000000001, d000000001;D000000001\n'),[P,'D000000001'])
    def test_empty_asins(self):
        with self.assertRaises(DataError): parse_asins('  ')
    def test_bad_asins(self):
        with self.assertRaises(DataError): parse_asins('not-an-asin')
    def test_all_empty(self):
        self.assertEqual(len(reconcile(pd.DataFrame(),pd.DataFrame())),0)
    def test_one_source(self):
        result=reconcile(table([[P,C,'a']]),pd.DataFrame())
        self.assertEqual(result.iloc[0]['Match Status'],'CSI only')
    def test_only_right(self):
        result=reconcile(pd.DataFrame(),table([[P,C,'a']]))
        self.assertEqual(result.iloc[0]['Match Status'],'RTIP only')
    def test_matching_uses_keys_only(self):
        result=reconcile(table([[P,C,'left']]),table([[P,C,'right']]))
        self.assertEqual(len(result),1)
        self.assertEqual(result.iloc[0]['Match Status'],'Both sources')
        self.assertEqual(result.iloc[0]['CSI_Value'],'left')
    def test_exact_duplicates(self):
        result=reconcile(table([[P,C,'a']]*3),table([[P,C,'b']]*4))
        self.assertEqual(result.iloc[0]['CSI_Rows'],1)
        self.assertEqual(len(result),1)
    def test_conflicting_duplicates(self):
        result=reconcile(table([[P,C,'a'],[P,C,'b']]),table([[P,C,'c'],[P,C,'d']]))
        self.assertEqual(len(result),1)
        self.assertEqual(result.iloc[0]['CSI_Value'],'a | b')
        self.assertEqual(result.iloc[0]['Review Required'],'Multiple source rows')
    def test_missing_columns(self):
        with self.assertRaises(DataError): normalize(pd.DataFrame({'wrong':[1]}),'CSI')
    def test_invalid_keys(self):
        with self.assertRaises(DataError): normalize(table([['',C,'a']]),'CSI')
    def test_missing_html_table(self):
        with self.assertRaises(DataError): parse_table('<p>Sign in</p>','productdata','CSI')
    def test_html_preserves_numeric_ids(self):
        source='<table id="productdata"><tr><th>Parent Item</th><th>Child Item</th></tr><tr><td>0000000001</td><td>0000000002</td></tr></table>'
        result=parse_table(source,'productdata','CSI')
        self.assertEqual(result.iloc[0]['Parent'],P)
    def test_duplicate_column_alias(self):
        with self.assertRaises(DataError): normalize(pd.DataFrame(columns=['Parent','Parent Item','Child']),'CSI')
    def test_csv_preserves_numeric_ids(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'data.csv'
            path.write_text('Parent,Child\n0000000001,0000000002\n')
            self.assertEqual(load_table(path,'CSI').iloc[0]['Parent'],P)
    def test_excel_roundtrip_and_formula_safety(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'out.xlsx'
            export_result(Result(csi=table([[P,C,'=1+1']])),path)
            wb=load_workbook(path)
            self.assertEqual(wb.sheetnames,['CSI Output','RTIP_Output','Final Data','Errors','Run Info'])
            self.assertEqual(wb['CSI Output']['C2'].data_type,'s')
            self.assertEqual(wb['CSI Output']['A2'].value,P)
            wb.close()
            self.assertEqual(load_table(path,'CSI').iloc[0]['Parent'],P)
    def test_overwrite_protection(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'out.xlsx'
            path.write_bytes(b'keep me')
            with self.assertRaises(DataError): export_result(Result(),path)
            self.assertEqual(path.read_bytes(),b'keep me')
            export_result(Result(),path,overwrite=True)
            wb=load_workbook(path); wb.close()
    def test_overlong_cell_does_not_replace_output(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'out.xlsx'; path.write_bytes(b'keep me')
            with self.assertRaises(DataError): export_result(Result(csi=table([[P,C,'x'*32768]])),path,True)
            self.assertEqual(path.read_bytes(),b'keep me')
    def test_large_batch(self):
        rows=[[f'{i:010}',f'{i+100000:010}','value'] for i in range(10000)]
        result=reconcile(table(rows+rows),table(rows))
        self.assertEqual(len(result),10000)
        self.assertTrue(result['Match Status'].eq('Both sources').all())
    def test_header_only(self):
        self.assertTrue(normalize(pd.DataFrame(columns=['Parent Item','Child Item']),'CSI').empty)
    def test_repeated_header_and_tools(self):
        result=normalize(pd.DataFrame([['Parent Item','Child Item','Value'],[P,C,'more tools Blue']],columns=['Parent Item','Child Item','Value']),'CSI')
        self.assertEqual(len(result),1)
        self.assertEqual(result.iloc[0]['Value'],'Blue')
    def test_unsupported_extension(self):
        with self.assertRaises(DataError): load_table('data.txt','CSI')

if __name__=='__main__': unittest.main()
