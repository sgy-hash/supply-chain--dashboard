import unittest
from io import BytesIO
import pandas as pd
from core import *
class Tests(unittest.TestCase):
 def data(self,**v):
  r={'SKU':'A','当前库存':10,'近30天出库量':300,'安全库存':20,'补货提前期(天)':7,'最小起订量':12};r.update(v);return pd.DataFrame([r])
 def test_moq(self):
  d,_=analyze_inventory(self.data(),AnalysisParameters());self.assertEqual(d.iloc[0]['建议补货量'],300)
 def test_short_target(self):
  d,_=analyze_inventory(self.data(),AnalysisParameters(target_cover_days=1));self.assertEqual(d.iloc[0]['建议补货量'],84)
 def test_no_sales(self):
  d,_=analyze_inventory(self.data(**{'近30天出库量':0}),AnalysisParameters());self.assertEqual(d.iloc[0]['建议补货量'],0)
 def test_invalid(self):
  for v in [-1,float('inf'),'abc']:
   with self.assertRaises(ValueError):analyze_inventory(self.data(**{'当前库存':v}),AnalysisParameters())
 def test_duplicate(self):
  with self.assertRaises(ValueError):analyze_inventory(pd.concat([self.data()]*2),AnalysisParameters())
 def test_delay(self):
  d,_=analyze_logistics(pd.DataFrame({'运单号':['A','B','C'],'承运商':['X']*3,'计划运输天数':[2]*3,'实际运输天数':[3,4,None]}),AnalysisParameters());self.assertEqual(d['运输状态'].tolist(),['准时','延误','运输中'])
 def test_export(self):
  d,_=analyze_inventory(self.data(),AnalysisParameters())
  for frame in [d,d.iloc[:0]]:
   b=export_analysis_excel(frame,pd.DataFrame(),pd.DataFrame(),pd.DataFrame());self.assertEqual(pd.ExcelFile(BytesIO(b)).sheet_names,['库存分析','补货建议'])
 def test_interface(self):
  from streamlit.testing.v1 import AppTest
  a=AppTest.from_file('streamlit_app.py',default_timeout=40).run();self.assertFalse(a.exception,str(a.exception))
  a.multiselect[1].select(a.multiselect[1].options[0]).run();self.assertFalse(a.exception,str(a.exception))
  a.text_input[0].set_value('[').run();self.assertFalse(a.exception,str(a.exception));self.assertIn('0',a.metric[0].value)
  a.radio[0].set_value('上传自己的Excel').run();self.assertFalse(a.exception,str(a.exception))
if __name__=='__main__':unittest.main()
