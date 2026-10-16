from pathlib import Path
from core import (
    AnalysisParameters,
    analyze_inventory,
    analyze_logistics,
    inventory_kpis,
    load_excel_data,
    logistics_kpis,
)

sample = Path(__file__).with_name("供应链示例数据.xlsx")
inventory_raw, logistics_raw, warnings = load_excel_data(sample)
params = AnalysisParameters()
inventory, inv_warnings = analyze_inventory(inventory_raw, params)
logistics, log_warnings = analyze_logistics(logistics_raw, params)

assert len(inventory) > 0
assert "风险状态" in inventory.columns
assert "建议补货量" in inventory.columns
assert len(logistics) > 0
assert "运输状态" in logistics.columns

print("库存记录:", len(inventory))
print("物流记录:", len(logistics))
print("库存KPI:", inventory_kpis(inventory))
print("物流KPI:", logistics_kpis(logistics))
print("验证成功")
