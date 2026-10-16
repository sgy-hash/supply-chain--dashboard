from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO

import numpy as np
import pandas as pd


INVENTORY_SHEET_CANDIDATES = ["库存明细", "库存", "Inventory", "inventory"]
LOGISTICS_SHEET_CANDIDATES = ["物流明细", "物流", "Logistics", "logistics"]


@dataclass(frozen=True)
class AnalysisParameters:
    analysis_days: int = 30
    stagnant_days: int = 90
    target_cover_days: int = 30
    high_cover_days: int = 60
    delay_threshold_days: int = 1


INVENTORY_ALIASES: dict[str, list[str]] = {
    "SKU": ["SKU", "sku", "商品编码", "货号", "物料编码"],
    "货品名称": ["货品名称", "商品名称", "产品名称", "品名", "name"],
    "品类": ["品类", "商品品类", "类别", "category"],
    "仓库": ["仓库", "仓库名称", "warehouse"],
    "期初库存": ["期初库存", "期初数量", "opening_stock"],
    "当前库存": ["当前库存", "现有库存", "库存数量", "on_hand"],
    "在途库存": ["在途库存", "在途数量", "in_transit"],
    "安全库存": ["安全库存", "safety_stock"],
    "近30天出库量": ["近30天出库量", "30天出库量", "近30天销量", "近30天销售量", "outbound_30d"],
    "近90天补货次数": ["近90天补货次数", "90天补货次数", "补货次数", "replenishment_count_90d"],
    "补货提前期(天)": ["补货提前期(天)", "补货提前期", "采购提前期(天)", "lead_time_days"],
    "单位成本": ["单位成本", "采购成本", "成本单价", "unit_cost"],
    "最长库龄(天)": ["最长库龄(天)", "库龄(天)", "最大库龄", "age_days"],
    "最小起订量": ["最小起订量", "MOQ", "moq"],
}

LOGISTICS_ALIASES: dict[str, list[str]] = {
    "运单号": ["运单号", "物流单号", "shipment_id", "waybill"],
    "发运日期": ["发运日期", "发货日期", "ship_date"],
    "SKU": ["SKU", "sku", "商品编码", "货号"],
    "起始仓": ["起始仓", "发货仓", "始发仓", "origin"],
    "目的地": ["目的地", "收货地", "destination"],
    "承运商": ["承运商", "物流商", "carrier"],
    "发运数量": ["发运数量", "运输数量", "发货数量", "quantity"],
    "计划运输天数": ["计划运输天数", "承诺时效(天)", "计划时长", "planned_days"],
    "实际运输天数": ["实际运输天数", "实际时长", "actual_days"],
    "运输成本": ["运输成本", "物流成本", "运费", "freight_cost"],
}


def _find_sheet_name(sheet_names: list[str], candidates: list[str]) -> str | None:
    normalized = {str(name).strip().lower(): str(name) for name in sheet_names}
    for candidate in candidates:
        exact = normalized.get(candidate.strip().lower())
        if exact:
            return exact
    for name in sheet_names:
        lowered = str(name).lower()
        if any(candidate.lower() in lowered for candidate in candidates):
            return str(name)
    return None


def load_excel_data(file: str | Path | BinaryIO) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Read an Excel workbook and find inventory/logistics sheets."""
    excel = pd.ExcelFile(file, engine="openpyxl")
    warnings: list[str] = []

    inventory_sheet = _find_sheet_name(excel.sheet_names, INVENTORY_SHEET_CANDIDATES)
    logistics_sheet = _find_sheet_name(excel.sheet_names, LOGISTICS_SHEET_CANDIDATES)

    if inventory_sheet is None:
        raise ValueError(
            "未找到库存工作表。请将工作表命名为“库存明细”或“Inventory”。"
        )
    inventory = pd.read_excel(excel, sheet_name=inventory_sheet)

    if logistics_sheet is None:
        warnings.append("未找到物流工作表，物流分析页将暂时为空。")
        logistics = pd.DataFrame()
    else:
        logistics = pd.read_excel(excel, sheet_name=logistics_sheet)

    return inventory, logistics, warnings


def _standardize_columns(
    df: pd.DataFrame,
    aliases: dict[str, list[str]],
) -> tuple[pd.DataFrame, list[str]]:
    result = df.copy()
    result.columns = [str(column).strip() for column in result.columns]
    lower_to_original = {str(column).strip().lower(): column for column in result.columns}

    rename_map: dict[str, str] = {}
    for canonical, options in aliases.items():
        if canonical in result.columns:
            continue
        for option in options:
            original = lower_to_original.get(option.strip().lower())
            if original is not None:
                rename_map[original] = canonical
                break

    result = result.rename(columns=rename_map)
    missing = [column for column in aliases if column not in result.columns]
    return result, missing


def _number(series: pd.Series, default: float = 0.0) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(default)


def _ceil_to_moq(quantity: pd.Series, moq: pd.Series) -> pd.Series:
    safe_moq = moq.where(moq > 0, 1)
    return np.ceil(quantity.clip(lower=0) / safe_moq) * safe_moq


def analyze_inventory(
    raw: pd.DataFrame,
    parameters: AnalysisParameters,
) -> tuple[pd.DataFrame, list[str]]:
    df, missing = _standardize_columns(raw, INVENTORY_ALIASES)
    warnings: list[str] = []

    essential = ["SKU", "当前库存", "近30天出库量"]
    missing_essential = [column for column in essential if column not in df.columns]
    if missing_essential:
        raise ValueError(
            "库存数据缺少必要字段：" + "、".join(missing_essential)
        )

    defaults: dict[str, Any] = {
        "货品名称": "",
        "品类": "未分类",
        "仓库": "未指定",
        "期初库存": 0,
        "在途库存": 0,
        "安全库存": 0,
        "近90天补货次数": 0,
        "补货提前期(天)": 7,
        "单位成本": 0,
        "最长库龄(天)": 0,
        "最小起订量": 1,
    }
    for column, default in defaults.items():
        if column not in df.columns:
            df[column] = default
            warnings.append(f"库存数据缺少“{column}”，系统已使用默认值 {default}。")

    df = df.dropna(how="all").copy()
    df["SKU"] = df["SKU"].astype(str).str.strip()
    df = df[df["SKU"].ne("") & df["SKU"].ne("nan")].copy()

    numeric_columns = [
        "期初库存", "当前库存", "在途库存", "安全库存", "近30天出库量",
        "近90天补货次数", "补货提前期(天)", "单位成本", "最长库龄(天)", "最小起订量",
    ]
    for column in numeric_columns:
        values = pd.to_numeric(df[column], errors="coerce")
        if values.isna().any() or not np.isfinite(values).all() or (values < 0).any():
            raise ValueError(f"库存字段“{column}”包含缺失、非数值、无穷或负数，请修正后上传。")
        df[column] = values

    for column in ["品类", "仓库", "货品名称"]:
        df[column] = df[column].fillna(defaults[column]).astype(str)
    if df.duplicated(["SKU", "仓库"]).any():
        raise ValueError("同一SKU与仓库存在重复记录，请先汇总为一条库存快照。")

    df["日均出库"] = df["近30天出库量"] / max(parameters.analysis_days, 1)
    df["平均库存"] = (df["期初库存"] + df["当前库存"]) / 2
    df["可用库存"] = df["当前库存"] + df["在途库存"]

    positive_demand = df["日均出库"] > 0
    df["周转天数"] = np.where(
        positive_demand,
        df["平均库存"] / df["日均出库"],
        np.nan,
    )
    df["库存覆盖天数"] = np.where(
        positive_demand,
        df["可用库存"] / df["日均出库"],
        np.nan,
    )
    df["月均补货频次"] = df["近90天补货次数"] / 3
    df["订货点"] = (
        df["日均出库"] * df["补货提前期(天)"] + df["安全库存"]
    )

    desired_stock = np.maximum(
        df["日均出库"] * parameters.target_cover_days, df["订货点"]
    )
    raw_order = desired_stock - df["可用库存"]
    need_order = positive_demand & (df["可用库存"] < df["订货点"])
    df["建议补货量"] = np.where(
        need_order,
        _ceil_to_moq(raw_order, df["最小起订量"]),
        0,
    )
    df["库存金额"] = df["当前库存"] * df["单位成本"]
    df["预计缺货天数"] = np.where(
        positive_demand,
        df["当前库存"] / df["日均出库"],
        np.nan,
    )

    no_sales = ~positive_demand
    shortage = positive_demand & (df["可用库存"] < df["订货点"])
    stagnant = positive_demand & (
        (df["周转天数"] > parameters.stagnant_days)
        | (df["最长库龄(天)"] > parameters.stagnant_days)
    )
    high_stock = positive_demand & (
        df["库存覆盖天数"] > parameters.high_cover_days
    )

    df["风险状态"] = np.select(
        [no_sales, shortage, stagnant, high_stock],
        ["无销量", "缺货风险", "呆滞库存", "高库存"],
        default="正常",
    )

    df["建议动作"] = np.select(
        [
            df["风险状态"].eq("缺货风险"),
            df["风险状态"].eq("无销量"),
            df["风险状态"].eq("呆滞库存"),
            df["风险状态"].eq("高库存"),
        ],
        [
            "优先补货并确认供应商交期；必要时从其他仓库调拨",
            "核查是否下架、断货或销量数据缺失；暂停自动补货",
            "暂停补货，考虑促销、退供或跨仓调拨",
            "降低补货频次，优先消化现有库存或调拨至缺货仓",
        ],
        default="维持当前补货策略并持续监控",
    )

    priority_map = {
        "缺货风险": 1,
        "无销量": 2,
        "呆滞库存": 3,
        "高库存": 4,
        "正常": 5,
    }
    df["处理优先级"] = df["风险状态"].map(priority_map).fillna(9).astype(int)

    ordered_columns = [
        "SKU", "货品名称", "品类", "仓库", "期初库存", "当前库存", "在途库存",
        "可用库存", "安全库存", "近30天出库量", "日均出库", "近90天补货次数",
        "月均补货频次", "补货提前期(天)", "订货点", "建议补货量", "单位成本",
        "库存金额", "最长库龄(天)", "周转天数", "库存覆盖天数", "预计缺货天数",
        "风险状态", "建议动作", "处理优先级",
    ]
    extra_columns = [column for column in df.columns if column not in ordered_columns]
    df = df[[column for column in ordered_columns if column in df.columns] + extra_columns]
    return df, warnings


def analyze_logistics(
    raw: pd.DataFrame,
    parameters: AnalysisParameters,
) -> tuple[pd.DataFrame, list[str]]:
    if raw is None or raw.empty:
        return pd.DataFrame(), ["暂无物流数据。"]

    df, missing = _standardize_columns(raw, LOGISTICS_ALIASES)
    warnings: list[str] = []

    essential = ["运单号", "承运商", "计划运输天数"]
    missing_essential = [column for column in essential if column not in df.columns]
    if missing_essential:
        raise ValueError(
            "物流数据缺少必要字段：" + "、".join(missing_essential)
        )

    defaults: dict[str, Any] = {
        "发运日期": pd.NaT,
        "SKU": "",
        "起始仓": "未指定",
        "目的地": "未指定",
        "发运数量": 0,
        "实际运输天数": np.nan,
        "运输成本": 0,
    }
    for column, default in defaults.items():
        if column not in df.columns:
            df[column] = default
            warnings.append(f"物流数据缺少“{column}”，系统已使用默认值。")

    df = df.dropna(how="all").copy()
    df["运单号"] = df["运单号"].astype(str).str.strip()
    df = df[df["运单号"].ne("") & df["运单号"].ne("nan")].copy()

    for column in ["发运数量", "计划运输天数", "实际运输天数", "运输成本"]:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    for column in ["发运数量", "计划运输天数", "运输成本"]:
        if df[column].isna().any() or not np.isfinite(df[column]).all() or (df[column] < 0).any():
            raise ValueError(f"物流字段“{column}”包含无效值，请提供非负数值。")
    actual = df["实际运输天数"]
    if (actual.dropna() < 0).any() or not np.isfinite(actual.dropna()).all():
        raise ValueError("实际运输天数必须为空（运输中）或非负有限数值。")
    df["发运数量"] = df["发运数量"].fillna(0)
    df["计划运输天数"] = df["计划运输天数"].fillna(0)
    df["运输成本"] = df["运输成本"].fillna(0)
    df["发运日期"] = pd.to_datetime(df["发运日期"], errors="coerce")

    completed = df["实际运输天数"].notna()
    df["延误天数"] = np.where(
        completed,
        (df["实际运输天数"] - df["计划运输天数"]).clip(lower=0),
        np.nan,
    )
    df["单位运输成本"] = np.where(
        df["发运数量"] > 0,
        df["运输成本"] / df["发运数量"],
        np.nan,
    )
    df["计划到达日期"] = df["发运日期"] + pd.to_timedelta(
        df["计划运输天数"], unit="D"
    )
    df["实际到达日期"] = df["发运日期"] + pd.to_timedelta(
        df["实际运输天数"], unit="D"
    )

    df["运输状态"] = np.select(
        [
            ~completed,
            completed & (df["延误天数"] > parameters.delay_threshold_days),
        ],
        ["运输中", "延误"],
        default="准时",
    )
    df["调度建议"] = np.select(
        [
            df["运输状态"].eq("延误"),
            df["运输状态"].eq("运输中"),
        ],
        [
            "复盘承运商和路线；连续延误时考虑切换承运商或提前发运",
            "跟踪预计到达日期，并关注可能影响缺货的在途订单",
        ],
        default="保持当前路线并持续监控成本",
    )
    return df, warnings


def carrier_summary(logistics: pd.DataFrame) -> pd.DataFrame:
    if logistics.empty:
        return pd.DataFrame(
            columns=["承运商", "已完成运单", "准时运单", "准时率", "平均运输天数", "平均单位运输成本"]
        )

    completed = logistics[logistics["运输状态"].ne("运输中")].copy()
    if completed.empty:
        return pd.DataFrame(
            columns=["承运商", "已完成运单", "准时运单", "准时率", "平均运输天数", "平均单位运输成本"]
        )

    grouped = completed.groupby("承运商", dropna=False)
    summary = grouped.agg(
        已完成运单=("运单号", "count"),
        准时运单=("运输状态", lambda s: int((s == "准时").sum())),
        平均运输天数=("实际运输天数", "mean"),
        平均单位运输成本=("单位运输成本", "mean"),
    ).reset_index()
    summary["准时率"] = summary["准时运单"] / summary["已完成运单"].replace(0, np.nan)
    return summary[
        ["承运商", "已完成运单", "准时运单", "准时率", "平均运输天数", "平均单位运输成本"]
    ].sort_values(["准时率", "平均运输天数"], ascending=[False, True])


def route_summary(logistics: pd.DataFrame) -> pd.DataFrame:
    if logistics.empty:
        return pd.DataFrame()

    completed = logistics[logistics["运输状态"].ne("运输中")].copy()
    if completed.empty:
        return pd.DataFrame()

    route = completed.groupby(["起始仓", "目的地"], dropna=False).agg(
        运单数=("运单号", "count"),
        延误运单=("运输状态", lambda s: int((s == "延误").sum())),
        平均运输天数=("实际运输天数", "mean"),
        平均运输成本=("运输成本", "mean"),
        平均单位运输成本=("单位运输成本", "mean"),
    ).reset_index()
    route["延误率"] = route["延误运单"] / route["运单数"].replace(0, np.nan)
    return route.sort_values(["延误率", "平均单位运输成本"], ascending=[False, False])


def inventory_kpis(inventory: pd.DataFrame) -> dict[str, float]:
    finite_turnover = inventory["周转天数"].replace([np.inf, -np.inf], np.nan)
    return {
        "库存总金额": float(inventory["库存金额"].sum()),
        "当前库存量": float(inventory["当前库存"].sum()),
        "在途库存量": float(inventory["在途库存"].sum()),
        "缺货风险SKU": int((inventory["风险状态"] == "缺货风险").sum()),
        "呆滞或无销量SKU": int(
            inventory["风险状态"].isin(["呆滞库存", "无销量"]).sum()
        ),
        "高库存SKU": int((inventory["风险状态"] == "高库存").sum()),
        "平均周转天数": float(finite_turnover.mean()) if finite_turnover.notna().any() else float("nan"),
        "建议补货总量": float(inventory["建议补货量"].sum()),
    }


def logistics_kpis(logistics: pd.DataFrame) -> dict[str, float]:
    if logistics.empty:
        return {
            "已完成运单": 0,
            "运输中运单": 0,
            "延误运单": 0,
            "准时率": float("nan"),
            "平均运输天数": float("nan"),
            "物流总成本": 0.0,
        }

    completed = logistics[logistics["运输状态"].ne("运输中")]
    on_time = (completed["运输状态"] == "准时").sum()
    return {
        "已完成运单": int(len(completed)),
        "运输中运单": int((logistics["运输状态"] == "运输中").sum()),
        "延误运单": int((logistics["运输状态"] == "延误").sum()),
        "准时率": float(on_time / len(completed)) if len(completed) else float("nan"),
        "平均运输天数": float(completed["实际运输天数"].mean()) if len(completed) else float("nan"),
        "物流总成本": float(logistics["运输成本"].sum()),
    }


def generate_management_insights(
    inventory: pd.DataFrame,
    logistics: pd.DataFrame,
) -> list[str]:
    insights: list[str] = []
    if inventory.empty and logistics.empty:
        return ["当前筛选没有匹配记录，请调整筛选后再查看诊断。"]

    shortage = inventory[inventory["风险状态"] == "缺货风险"]
    if not shortage.empty:
        top = shortage.sort_values("建议补货量", ascending=False).head(3)
        sku_text = "、".join(top["SKU"].astype(str).tolist())
        insights.append(
            f"共有 {len(shortage)} 个SKU存在缺货风险，其中建议优先处理：{sku_text}。"
        )

    slow = inventory[inventory["风险状态"].isin(["呆滞库存", "无销量"])]
    if not slow.empty:
        value = slow["库存金额"].sum()
        insights.append(
            f"呆滞或无销量库存涉及 {len(slow)} 个SKU，占用库存金额约 ¥{value:,.0f}；应优先暂停补货并安排促销、退供或跨仓调拨。"
        )

    high = inventory[inventory["风险状态"] == "高库存"]
    if not high.empty:
        insights.append(
            f"共有 {len(high)} 个SKU库存覆盖过高，建议降低补货频次，并与缺货仓库进行调拨匹配。"
        )

    if not logistics.empty:
        carrier = carrier_summary(logistics)
        if not carrier.empty:
            worst = carrier.sort_values(["准时率", "平均运输天数"], ascending=[True, False]).iloc[0]
            insights.append(
                f"承运商“{worst['承运商']}”的准时率相对较低（{worst['准时率']:.1%}），建议复盘其高延误路线并评估替代承运商。"
            )

    if not insights:
        insights.append("目前未识别出明显异常，建议继续按周监控需求、库存和运输时效变化。")
    return insights


def export_analysis_excel(
    inventory: pd.DataFrame,
    logistics: pd.DataFrame,
    carrier: pd.DataFrame,
    routes: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        inventory.to_excel(writer, sheet_name="库存分析", index=False)
        inventory.sort_values(
            ["处理优先级", "建议补货量"], ascending=[True, False]
        )[
            [
                "SKU", "货品名称", "品类", "仓库", "风险状态", "可用库存",
                "订货点", "建议补货量", "周转天数", "库存覆盖天数", "建议动作",
            ]
        ].to_excel(
            writer, sheet_name="补货建议", index=False
        )
        if not logistics.empty:
            logistics.to_excel(writer, sheet_name="物流分析", index=False)
        if not carrier.empty:
            carrier.to_excel(writer, sheet_name="承运商绩效", index=False)
        if not routes.empty:
            routes.to_excel(writer, sheet_name="路线分析", index=False)
    return output.getvalue()
