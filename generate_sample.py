"""Generate fictional demo data. No company records or personal data."""
from pathlib import Path
import numpy as np
import pandas as pd
rng=np.random.default_rng(42)
rows=[]
for i in range(60):
 rows.append({'SKU':f'DEMO-{i:03}','货品名称':f'模拟商品{i:02}','品类':['生鲜','饮料','日用品'][i%3],'仓库':['上海仓','深圳仓'][i%2],'期初库存':int(rng.integers(0,1200)),'当前库存':int(rng.integers(0,1800)),'在途库存':int(rng.integers(0,100)),'安全库存':50,'近30天出库量':0 if i%10==0 else int(rng.integers(10,1600)),'近90天补货次数':int(rng.integers(0,15)),'补货提前期(天)':7,'单位成本':float(rng.integers(2,100)),'最长库龄(天)':int(rng.integers(1,150)),'最小起订量':12})
shipments=[]
for i in range(120):
 j=i%60
 shipments.append({'运单号':f'DEMO-SHIP-{i:04}','SKU':f'DEMO-{j:03}','发运日期':pd.Timestamp('2026-09-01')+pd.Timedelta(days=i%30),'起始仓':['上海仓','深圳仓'][j%2],'目的地':['北京','杭州','广州'][i%3],'承运商':['模拟承运商A','模拟承运商B'][i%2],'发运数量':100,'计划运输天数':3,'实际运输天数':None if i%7==0 else int(rng.integers(1,8)),'运输成本':float(rng.integers(100,500))})
with pd.ExcelWriter(Path(__file__).with_name('供应链示例数据.xlsx')) as w:
 pd.DataFrame(rows).to_excel(w,sheet_name='库存明细',index=False)
 pd.DataFrame(shipments).to_excel(w,sheet_name='物流明细',index=False)
