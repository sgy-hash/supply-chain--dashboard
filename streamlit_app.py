from __future__ import annotations

from pathlib import Path
import math

import pandas as pd
import plotly.express as px
import streamlit as st

from core import (
    AnalysisParameters,
    analyze_inventory,
    analyze_logistics,
    carrier_summary,
    export_analysis_excel,
    generate_management_insights,
    inventory_kpis,
    load_excel_data,
    logistics_kpis,
    route_summary,
)


st.set_page_config(
    page_title="供应链库存与物流看板",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 3rem;}
    div[data-testid="stMetric"] {
        background: var(--secondary-background-color);
        border: 1px solid #e6e9ef;
        border-radius: 12px;
        padding: 14px;
        box-shadow: 0 2px 8px rgba(20, 30, 55, 0.04);
    }
    .small-note {
        color: #64748b;
        font-size: 0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.caption("SUPPLY CHAIN · OPERATIONS INTELLIGENCE")
st.title("供应链运营工作台")
st.caption("上传库存与物流Excel，自动分析库存结构、周转效率、补货风险与运输表现。")


@st.cache_data(show_spinner=False)
def load_and_analyze(source, params: AnalysisParameters):
    raw_inventory, raw_logistics, load_warnings = load_excel_data(source)
    inventory, inv_warnings = analyze_inventory(raw_inventory, params)
    logistics, log_warnings = analyze_logistics(raw_logistics, params)
    return (
        inventory,
        logistics,
        load_warnings + inv_warnings + log_warnings,
    )


def money(value: float) -> str:
    return f"¥{value:,.0f}"


def number(value: float, digits: int = 1) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    return f"{value:,.{digits}f}"


def percentage(value: float) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    return f"{value:.1%}"


with st.sidebar:
    st.header("数据与参数")

    mode = st.radio(
        "数据来源",
        ["使用内置示例数据", "上传自己的Excel"],
        help="第一次使用建议先选择示例数据，确认网页能够正常运行。",
    )

    uploaded = None
    if mode == "上传自己的Excel":
        uploaded = st.file_uploader(
            "上传Excel文件",
            type=["xlsx", "xlsm"],
            help="工作簿需至少包含“库存明细”工作表；物流工作表可命名为“物流明细”。",
        )

    st.divider()
    st.subheader("分析参数")
    analysis_days = st.number_input("出库数据对应天数", 7, 365, 30, help="仅修改分母；不会生成历史数据或改变Excel中的出库量。默认字段是近30天出库量，请保持30，除非你已更换对应周期的数据。")
    stagnant_days = st.number_input("呆滞库存阈值（天）", 30, 365, 90)
    target_cover_days = st.number_input("目标库存覆盖天数", 7, 180, 30)
    high_cover_days = st.number_input("高库存覆盖阈值（天）", 14, 365, 60)
    delay_threshold = st.number_input("物流延误阈值（天）", 0, 30, 1)

    parameters = AnalysisParameters(
        analysis_days=int(analysis_days),
        stagnant_days=int(stagnant_days),
        target_cover_days=int(target_cover_days),
        high_cover_days=int(high_cover_days),
        delay_threshold_days=int(delay_threshold),
    )

    st.divider()
    sample_path = Path(__file__).with_name("供应链示例数据.xlsx")
    if sample_path.exists():
        st.download_button(
            "下载示例Excel",
            data=sample_path.read_bytes(),
            file_name="供应链示例数据.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

if mode == "上传自己的Excel" and uploaded is None:
    st.info("请在左侧上传Excel。你也可以先切换到“使用内置示例数据”查看完整效果。")
    st.stop()

data_source = uploaded if uploaded is not None else sample_path

try:
    with st.spinner("正在读取并分析数据……"):
        inventory, logistics, warnings = load_and_analyze(data_source, parameters)
except Exception as exc:
    st.error(f"数据读取或分析失败：{exc}")
    with st.expander("检查Excel格式"):
        st.markdown(
            """
            库存工作表建议至少包含：`SKU`、`当前库存`、`近30天出库量`。  
            建议同时提供：货品名称、品类、仓库、期初库存、在途库存、安全库存、
            近90天补货次数、补货提前期(天)、单位成本、最长库龄(天)、最小起订量。

            物流工作表建议包含：`运单号`、`承运商`、`计划运输天数`，以及实际运输天数、
            发运数量和运输成本等字段。
            """
        )
    st.stop()

if warnings:
    with st.expander(f"数据质量提示 · {len(warnings)} 项"):
        for warning in warnings:
            st.warning(warning)

# Global filters
st.subheader("筛选")
filter_cols = st.columns(4)
category_options = sorted(inventory["品类"].dropna().astype(str).unique().tolist())
warehouse_options = sorted(inventory["仓库"].dropna().astype(str).unique().tolist())
status_options = ["缺货风险", "呆滞库存", "高库存", "无销量", "正常"]

with filter_cols[0]:
    selected_categories = st.multiselect("品类", category_options)
with filter_cols[1]:
    selected_warehouses = st.multiselect("仓库", warehouse_options)
with filter_cols[2]:
    selected_statuses = st.multiselect("库存状态", status_options)
with filter_cols[3]:
    sku_keyword = st.text_input("SKU或货品名称搜索")

filtered = inventory.copy()
if selected_categories:
    filtered = filtered[filtered["品类"].astype(str).isin(selected_categories)]
if selected_warehouses:
    filtered = filtered[filtered["仓库"].astype(str).isin(selected_warehouses)]
if selected_statuses:
    filtered = filtered[filtered["风险状态"].isin(selected_statuses)]
if sku_keyword.strip():
    keyword = sku_keyword.strip()
    filtered = filtered[
        filtered["SKU"].astype(str).str.contains(keyword, case=False, na=False, regex=False)
        | filtered["货品名称"].astype(str).str.contains(keyword, case=False, na=False, regex=False)
    ]

# All views and downloads share the same inventory scope. Logistics follows SKU
# and warehouse when filters are active; dates apply to logistics only.
if selected_categories or selected_warehouses or selected_statuses or sku_keyword.strip():
    if not logistics.empty:
        logistics = logistics[logistics["SKU"].astype(str).isin(filtered["SKU"].astype(str))]
        if selected_warehouses:
            logistics = logistics[logistics["起始仓"].isin(selected_warehouses)]
if not logistics.empty:
    with st.expander("物流筛选 · 承运商与发运日期"):
        selected_carriers = st.multiselect("承运商", sorted(logistics["承运商"].dropna().astype(str).unique()))
        dates = logistics["发运日期"].dropna()
        if not dates.empty:
            date_range = st.date_input("发运日期范围（仅影响物流）", (dates.min().date(), dates.max().date()))
            if len(date_range) == 2:
                logistics = logistics[logistics["发运日期"].dt.date.between(*date_range)]
        if selected_carriers:
            logistics = logistics[logistics["承运商"].isin(selected_carriers)]
inv_kpi = inventory_kpis(filtered)
log_kpi = logistics_kpis(logistics)
carrier = carrier_summary(logistics)
routes = route_summary(logistics)
st.caption(f"当前范围：{len(filtered):,} 条SKU×仓库记录 / 全部 {len(inventory):,} 条；{len(logistics):,} 条运单。库存为快照，非历史趋势。")
if filtered.empty:
    st.info("没有匹配的库存记录，请清除或调整筛选条件。")

tabs = st.tabs(["经营总览", "库存与周转", "补货建议", "物流效率", "数据说明"])

with tabs[0]:
    st.subheader("核心指标")
    row1 = st.columns(4)
    row1[0].metric("库存总金额", money(inv_kpi["库存总金额"]))
    row1[1].metric("缺货风险记录", f"{inv_kpi['缺货风险SKU']:,}")
    row1[2].metric("呆滞/无销量记录", f"{inv_kpi['呆滞或无销量SKU']:,}")
    row1[3].metric("平均周转天数", number(inv_kpi["平均周转天数"]))

    row2 = st.columns(4)
    row2[0].metric("在途库存量", f"{inv_kpi['在途库存量']:,.0f}")
    row2[1].metric("建议补货总量", f"{inv_kpi['建议补货总量']:,.0f}")
    row2[2].metric("物流准时率", percentage(log_kpi["准时率"]))
    row2[3].metric("延误运单", f"{log_kpi['延误运单']:,}")

    st.subheader("风险诊断与行动建议")
    st.caption("基于明确规则生成，未接入大语言模型；建议需结合实际业务判断。")
    for insight in generate_management_insights(filtered, logistics):
        st.markdown(f"- {insight}")

    left, right = st.columns(2)
    status_summary = (
        filtered.groupby("风险状态", as_index=False)
        .agg(SKU数量=("SKU", "count"), 库存金额=("库存金额", "sum"))
    )
    with left:
        fig = px.pie(
            status_summary,
            names="风险状态",
            values="SKU数量",
            hole=0.48,
            title="库存状态结构",
        )
        st.plotly_chart(fig, width="stretch")

    category_summary = (
        filtered.groupby("品类", as_index=False)
        .agg(库存金额=("库存金额", "sum"), 当前库存=("当前库存", "sum"))
        .sort_values("库存金额", ascending=False)
    )
    with right:
        fig = px.bar(
            category_summary,
            x="品类",
            y="库存金额",
            text_auto=".2s",
            title="各品类库存金额",
        )
        st.plotly_chart(fig, width="stretch")

    col1, col2 = st.columns(2)
    turnover_summary = (
        filtered.dropna(subset=["周转天数"])
        .groupby("品类", as_index=False)["周转天数"]
        .mean()
        .sort_values("周转天数", ascending=False)
    )
    with col1:
        fig = px.bar(
            turnover_summary,
            x="周转天数",
            y="品类",
            orientation="h",
            title="各品类平均周转天数",
        )
        st.plotly_chart(fig, width="stretch")

    if not carrier.empty:
        with col2:
            fig = px.bar(
                carrier,
                x="承运商",
                y="准时率",
                text=carrier["准时率"].map(lambda value: f"{value:.0%}"),
                title="承运商准时率",
                range_y=[0, 1],
            )
            st.plotly_chart(fig, width="stretch")
    else:
        col2.info("暂无足够的已完成运单用于承运商分析。")

with tabs[1]:
    st.subheader("库存结构与周转效率")
    st.caption(f"当前筛选后共有 {len(filtered):,} 个SKU记录。")

    scatter_data = filtered.dropna(subset=["库存覆盖天数", "周转天数"]).copy()
    scatter_data = scatter_data[
        (scatter_data["库存覆盖天数"] < 500)
        & (scatter_data["周转天数"] < 500)
    ]
    if not scatter_data.empty:
        fig = px.scatter(
            scatter_data,
            x="库存覆盖天数",
            y="周转天数",
            size="库存金额",
            color="风险状态",
            hover_data=["SKU", "货品名称", "品类", "仓库"],
            title="库存覆盖天数与周转天数",
        )
        fig.add_vline(
            x=parameters.high_cover_days,
            line_dash="dash",
            annotation_text="高库存阈值",
        )
        fig.add_hline(
            y=parameters.stagnant_days,
            line_dash="dash",
            annotation_text="呆滞阈值",
        )
        st.plotly_chart(fig, width="stretch")

    display_columns = [
        "SKU", "货品名称", "品类", "仓库", "当前库存", "在途库存", "库存金额",
        "日均出库", "周转天数", "库存覆盖天数", "最长库龄(天)", "风险状态", "建议动作",
    ]
    st.dataframe(
        filtered.sort_values(
            ["处理优先级", "库存金额"], ascending=[True, False]
        )[display_columns],
        width="stretch",
        hide_index=True,
        column_config={
            "库存金额": st.column_config.NumberColumn(format="¥%.2f"),
            "日均出库": st.column_config.NumberColumn(format="%.2f"),
            "周转天数": st.column_config.NumberColumn(format="%.1f"),
            "库存覆盖天数": st.column_config.NumberColumn(format="%.1f"),
        },
    )

with tabs[2]:
    st.subheader("补货与库存优化建议")
    recommendation = filtered[
        [
            "SKU", "货品名称", "品类", "仓库", "风险状态", "可用库存", "订货点",
            "建议补货量", "预计缺货天数", "周转天数", "库存覆盖天数", "建议动作",
            "处理优先级",
        ]
    ].sort_values(["处理优先级", "建议补货量"], ascending=[True, False])

    priority_only = st.checkbox("只显示需要处理的SKU", value=True)
    if priority_only:
        recommendation = recommendation[recommendation["风险状态"] != "正常"]

    st.dataframe(
        recommendation.drop(columns=["处理优先级"]),
        width="stretch",
        hide_index=True,
        column_config={
            "建议补货量": st.column_config.NumberColumn(format="%.0f"),
            "订货点": st.column_config.NumberColumn(format="%.1f"),
            "预计缺货天数": st.column_config.NumberColumn(format="%.1f"),
            "周转天数": st.column_config.NumberColumn(format="%.1f"),
            "库存覆盖天数": st.column_config.NumberColumn(format="%.1f"),
        },
    )

    export_bytes = export_analysis_excel(filtered, logistics, carrier, routes)
    st.download_button(
        "下载当前筛选分析结果",
        data=export_bytes,
        file_name="供应链分析结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

with tabs[3]:
    st.subheader("物流时效与成本")
    if logistics.empty:
        st.info("当前文件没有可用的物流数据。")
    else:
        log_metrics = st.columns(5)
        log_metrics[0].metric("已完成运单", log_kpi["已完成运单"])
        log_metrics[1].metric("运输中运单", log_kpi["运输中运单"])
        log_metrics[2].metric("延误运单", log_kpi["延误运单"])
        log_metrics[3].metric("平均运输天数", number(log_kpi["平均运输天数"]))
        log_metrics[4].metric("物流总成本", money(log_kpi["物流总成本"]))

        c1, c2 = st.columns(2)
        with c1:
            if not carrier.empty:
                st.markdown("#### 承运商绩效")
                st.dataframe(
                    carrier,
                    width="stretch",
                    hide_index=True,
                    column_config={
                        "准时率": st.column_config.ProgressColumn(
                            min_value=0, max_value=1, format="%.1%%"
                        ),
                        "平均运输天数": st.column_config.NumberColumn(format="%.1f"),
                        "平均单位运输成本": st.column_config.NumberColumn(format="¥%.2f"),
                    },
                )
        with c2:
            if not routes.empty:
                st.markdown("#### 高延误路线")
                st.dataframe(
                    routes.head(15),
                    width="stretch",
                    hide_index=True,
                    column_config={
                        "延误率": st.column_config.ProgressColumn(
                            min_value=0, max_value=1, format="%.1%%"
                        ),
                        "平均运输成本": st.column_config.NumberColumn(format="¥%.2f"),
                        "平均单位运输成本": st.column_config.NumberColumn(format="¥%.2f"),
                    },
                )

        st.markdown("#### 运单明细")
        st.dataframe(
            logistics.sort_values(
                ["运输状态", "延误天数"], ascending=[True, False]
            ),
            width="stretch",
            hide_index=True,
            column_config={
                "发运日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "计划到达日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "实际到达日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                "运输成本": st.column_config.NumberColumn(format="¥%.2f"),
                "单位运输成本": st.column_config.NumberColumn(format="¥%.2f"),
            },
        )

with tabs[4]:
    st.subheader("数据口径与使用说明")
    st.markdown(
        f"""
        **主要计算口径**

        - 日均出库量 = 近{parameters.analysis_days}天出库量 ÷ {parameters.analysis_days}
        - 库存覆盖天数 =（当前库存 + 在途库存）÷ 日均出库量
        - 周转天数 = 平均库存 ÷ 日均出库量
        - 订货点 = 日均出库量 × 补货提前期 + 安全库存
        - 建议补货量：当可用库存低于订货点时，补至目标覆盖库存与订货点中的较大值，并按最小起订量向上取整
        - 呆滞库存：周转天数或最长库龄超过 {parameters.stagnant_days} 天
        - 高库存：库存覆盖天数超过 {parameters.high_cover_days} 天
        - 物流延误：实际运输天数减去计划运输天数，严格大于 {parameters.delay_threshold_days} 天

        **重要说明**

        此版本使用静态库存快照，周转天数为各记录简单平均，金额使用统一人民币口径。在途库存不代表可以立即使用，尚未按预计到货日分配。未纳入保质期、批次和损耗率，因此不能据此直接执行香蕉等生鲜品补货。真实企业使用前，应结合业务特点进一步校准：
        季节性需求、促销计划、供应商可靠性、采购批量、服务水平目标、退货和跨仓调拨等因素。
        """
    )

    st.markdown("#### 系统识别到的库存字段")
    st.write(list(inventory.columns))
    if not logistics.empty:
        st.markdown("#### 系统识别到的物流字段")
        st.write(list(logistics.columns))
