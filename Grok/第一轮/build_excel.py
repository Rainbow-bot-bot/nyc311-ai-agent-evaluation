"""Build 最终成果.xlsx from summary CSVs."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart.marker import Marker

SUM = Path(r"D:\项目2\Grok\summaries")
OUT = Path(r"D:\项目2\Grok\最终成果.xlsx")

NAVY = "1F4E79"
TEAL = "2E75B6"
GOLD = "C65911"
GREEN = "548235"
GRAY = "F2F2F2"
WHITE = "FFFFFF"
DARK = "1A1A1A"
MUTED = "666666"
LIGHT_BLUE = "D6EAF8"
LIGHT_ORANGE = "FCE4D6"
LIGHT_GREEN = "E2EFDA"
LIGHT_YELLOW = "FFF2CC"

thin = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
header_font = Font(name="Calibri", bold=True, color=WHITE, size=11)
title_font = Font(name="Calibri", bold=True, color=NAVY, size=18)
h2_font = Font(name="Calibri", bold=True, color=NAVY, size=13)
body_font = Font(name="Calibri", size=11, color=DARK)
kpi_font = Font(name="Calibri", bold=True, color=NAVY, size=16)
small_font = Font(name="Calibri", size=9, color=MUTED, italic=True)


def fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def style_header_row(ws, row: int, start: int, end: int, color=NAVY):
    for col in range(start, end + 1):
        cell = ws.cell(row, col)
        cell.fill = fill(color)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin


def write_df(ws, df: pd.DataFrame, r0: int, c0: int = 1, header_color=NAVY):
    for i, col in enumerate(df.columns):
        cell = ws.cell(r0, c0 + i, col)
        cell.fill = fill(header_color)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        cell.border = thin
    for ri, row in enumerate(df.itertuples(index=False), start=1):
        for ci, val in enumerate(row):
            cell = ws.cell(r0 + ri, c0 + ci, val)
            cell.font = body_font
            cell.border = thin
            cell.alignment = Alignment(vertical="center")
            if isinstance(val, float):
                if abs(val) >= 100:
                    cell.number_format = "#,##0.0"
                elif abs(val) >= 1:
                    cell.number_format = "#,##0.00"
                else:
                    cell.number_format = "0.0%"
            elif isinstance(val, int) and not isinstance(val, bool):
                cell.number_format = "#,##0"
            if ri % 2 == 0:
                if cell.fill.fgColor is None or cell.fill.fgColor.rgb in ("00000000", None):
                    cell.fill = fill(GRAY)
    return r0 + len(df)


def autosize(ws, min_w=10, max_w=42):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        longest = 0
        for cell in col:
            if cell.value is None:
                continue
            longest = max(longest, min(len(str(cell.value)), 60))
        ws.column_dimensions[letter].width = min(max(longest + 2, min_w), max_w)


def kpi_box(ws, r, c, label, value, note, color=NAVY):
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
    lab = ws.cell(r, c, label)
    lab.font = Font(name="Calibri", bold=True, color=WHITE, size=10)
    lab.fill = fill(color)
    lab.alignment = Alignment(horizontal="center")
    ws.cell(r, c + 1).fill = fill(color)
    ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
    val = ws.cell(r + 1, c, value)
    val.font = kpi_font
    val.fill = fill("F7F9FC")
    val.alignment = Alignment(horizontal="center")
    ws.cell(r + 1, c + 1).fill = fill("F7F9FC")
    ws.merge_cells(start_row=r + 2, start_column=c, end_row=r + 2, end_column=c + 1)
    n = ws.cell(r + 2, c, note)
    n.font = small_font
    n.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[r + 2].height = 32


# ---------- load ----------
overview = pd.read_csv(SUM / "overview.csv")
monthly = pd.read_csv(SUM / "monthly.csv")
types = pd.read_csv(SUM / "types_overall.csv")
yoy_type = pd.read_csv(SUM / "yoy_type.csv")
yoy_boro = pd.read_csv(SUM / "yoy_borough.csv")
yoy_ag = pd.read_csv(SUM / "yoy_agency.csv")
close_type = pd.read_csv(SUM / "close_by_type.csv")
close_ag = pd.read_csv(SUM / "close_by_agency.csv")
hn = pd.read_csv(SUM / "heat_noise_parking_monthly.csv")
heat_b = pd.read_csv(SUM / "heat_by_borough.csv")
boro_share = pd.read_csv(SUM / "diag_boro_type_shares.csv")
contrib = pd.read_csv(SUM / "diag_yoy_contrib.csv")
yoy_tot = pd.read_csv(SUM / "diag_yoy_totals.csv")
snow = pd.read_csv(SUM / "diag_snow_monthly.csv")
feb24 = pd.read_csv(SUM / "diag_2026_02_24_types.csv")
water = pd.read_csv(SUM / "diag_water_recode.csv")
heat_close = pd.read_csv(SUM / "diag_heat_close_monthly.csv")
status = pd.read_csv(SUM / "status_counts.csv")
agency = pd.read_csv(SUM / "agency_counts.csv")
channel = pd.read_csv(SUM / "channel_counts.csv")
borough = pd.read_csv(SUM / "borough_counts.csv")
open_age = pd.read_csv(SUM / "open_age_by_agency.csv")
dow = pd.read_csv(SUM / "dow_volume.csv")
newyear = pd.read_csv(SUM / "diag_newyear_window.csv")
febwin = pd.read_csv(SUM / "diag_2026_feb_window.csv")
daily = pd.read_csv(SUM / "daily.csv")

n_all = int(overview.n_rows.iloc[0])
n_keys = int(overview.n_unique_keys.iloc[0])
n_y1 = int(yoy_tot.n_y1.iloc[0])
n_y2 = int(yoy_tot.n_y2.iloc[0])
delta_yoy = n_y2 - n_y1
pct_yoy = delta_yoy / n_y1

# derived monthly
m = monthly.merge(hn, on="created_month")
m["heat_share"] = m["n_heat"] / m["n_created"]
m["open_share"] = m["n_openish"] / m["n_created"]
m["note"] = ""
m.loc[m.created_month == "2024-09", "note"] = "月初不完整（9/7起）"
m.loc[m.created_month == "2026-09", "note"] = "月末不完整（至9/5）"
m.loc[m.created_month >= "2026-07", "note"] = m["note"].where(
    m["note"] != "", "近期未结案抬升，结案时长受快照截断"
)

# yoy type table
yt = yoy_type.copy()
yt["delta"] = yt["n_y2"] - yt["n_y1"]
yt["pct"] = yt["delta"] / yt["n_y1"].replace(0, pd.NA)
yt["share_of_delta"] = yt["delta"] / delta_yoy
yt = yt.sort_values("delta", ascending=False)

contrib = contrib.copy()
contrib["pct_y1"] = contrib["n_y1"] / n_y1
contrib["pct_y2"] = contrib["n_y2"] / n_y2
contrib["share_of_delta"] = contrib["delta"] / delta_yoy

# borough shares
bs = boro_share.copy()
bs["all_share"] = bs["n_all"] / bs["n_all"].sum()
bs["heat_share_in_boro"] = bs["n_heat"] / bs["n_all"]
bs["heat_share_city"] = bs["n_heat"] / bs["n_heat"].sum()
bs["park_share_in_boro"] = bs["n_park"] / bs["n_all"]
bs["noise_share_in_boro"] = bs["n_noise"] / bs["n_all"]

# close type display
ct = close_type.head(18).copy()
ct["close_rate"] = ct["n_valid_close"] / ct["n_created"]
ct["median_days"] = ct["median_close_hours"] / 24
ct["p90_days"] = ct["p90_close_hours"] / 24

ca = close_ag.copy()
ca["open_rate"] = ca["n_openish"] / ca["n_created"]
ca["median_days"] = ca["median_close_hours"] / 24

# winter windows
def months_between(df, col, months):
    return df[df[col].isin(months)]

w1 = ["2024-11", "2024-12", "2025-01", "2025-02", "2025-03"]
w2 = ["2025-11", "2025-12", "2026-01", "2026-02", "2026-03"]
heat_w1 = int(m.loc[m.created_month.isin(w1), "n_heat"].sum())
heat_w2 = int(m.loc[m.created_month.isin(w2), "n_heat"].sum())
all_w1 = int(m.loc[m.created_month.isin(w1), "n_created"].sum())
all_w2 = int(m.loc[m.created_month.isin(w2), "n_created"].sum())
snow_w1 = int(snow.loc[snow.m.isin(w1), "n"].sum())
snow_w2 = int(snow.loc[snow.m.isin(w2), "n"].sum())

dow_map = {0: "周日", 1: "周一", 2: "周二", 3: "周三", 4: "周四", 5: "周五", 6: "周六"}
dow["星期"] = dow["dow"].map(dow_map)

wb = Workbook()

# ===================== 1 使用说明 =====================
ws = wb.active
ws.title = "使用说明"
ws.sheet_properties.tabColor = NAVY
ws["A1"] = "NYC 311 服务请求｜数据分析看板"
ws["A1"].font = title_font
ws.merge_cells("A1:G1")
ws["A2"] = (
    "记录粒度：一条 unique_key = 一条 311 服务请求记录（不等于独立现实事件；同一问题可被多次申告）。"
    "数据来源：NYC Open Data 导出的按创建日期切分的 Parquet 分片，只读扫描，未改写原始文件。"
)
ws["A2"].font = body_font
ws["A2"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A2:G2")
ws.row_dimensions[2].height = 40

ws["A4"] = "建议阅读顺序"
ws["A4"].font = h2_font
order = [
    ("总览", "总量、时间窗、状态与机构构成。先看分母。"),
    ("核心发现", "四条可核验结论：同比增量来源、冬季结构、两类结案速度、快照/更名偏误。"),
    ("月度趋势", "创建量与供暖/噪声/违停结构；不完整月已标注。"),
    ("同比分解", "2024-10～2025-08 vs 2025-10～2026-08（各 11 个完整月）的类型贡献。"),
    ("冬季与尖峰", "供暖季节、积雪日 2026-02-24、跨年噪声对照。"),
    ("结案时长", "仅用 status=Closed 且 closed≥created 的记录；中位数优先于均值。"),
    ("行政区", "布朗克斯供暖与住宅噪声占比偏高（投诉构成，非人均）。"),
    ("数据质量", "缺失、截断月、类型更名、未结案年龄。"),
]
ws["A5"] = "工作表"
ws["B5"] = "看什么"
style_header_row(ws, 5, 1, 2)
for i, (a, b) in enumerate(order, start=6):
    ws.cell(i, 1, a).font = Font(name="Calibri", bold=True, size=11)
    ws.cell(i, 2, b).font = body_font
    ws.cell(i, 1).border = thin
    ws.cell(i, 2).border = thin
    if i % 2 == 0:
        ws.cell(i, 1).fill = fill(GRAY)
        ws.cell(i, 2).fill = fill(GRAY)

ws["A15"] = "证据层级（本看板全程使用）"
ws["A15"].font = h2_font
ws["A16"] = (
    "描述性证据：计数、占比、中位数、时间窗内的增减。这些数字可从 summaries/*.csv 与 DuckDB 扫描复现。\n"
    "推断：例如“2025–26 冬积雪投诉暴增，与 2026-02-24 单日尖峰同向”，属于对记录结构的合理解读，但不是气象因果证明。\n"
    "不做的因果声称：311 记录增加 ≠ 城市问题恶化；结案变慢 ≠ 机构绩效下降（近期尤其受快照截断影响）；行政区差异 ≠ 服务质量排名。"
)
ws["A16"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("A16:G16")
ws.row_dimensions[16].height = 72

ws["A18"] = "关键口径（计算所有比率前先看这里）"
ws["A18"].font = h2_font
caliber = [
    ("时间范围", "创建时间 2024-09-07 00:00:12 至 2026-09-05 01:50:33。分片文件名写到 2026-09-07，但库内最晚创建日为 09-05。"),
    ("完整月同比", "Y1=2024-10-01≤created<2025-09-01；Y2=2025-10-01≤created<2026-09-01。排除 2024-09 与 2026-09 两个残缺月。"),
    ("结案时长", "小时 = closed_ts − created_ts。仅 status=Closed 且 closed_ts≥created_ts。排除 1,877 条闭案早于创建。"),
    ("未结案", "status ∈ {Open, In Progress, Assigned, Pending, Started}。不是“永远未解决”，只是快照时尚未 Closed。"),
    ("快照偏误", "数据像一次静态抽取。越靠近抽取日，慢工单越来不及进入 Closed，近期中位结案时长会被“已结的快单”拉低。"),
    ("单位", "除非标明小时/天，计数均为服务请求记录数。"),
]
ws["A19"] = "项目"
ws["B19"] = "定义"
style_header_row(ws, 19, 1, 2)
for i, (a, b) in enumerate(caliber, start=20):
    ws.cell(i, 1, a).font = Font(name="Calibri", bold=True)
    ws.cell(i, 2, b).alignment = Alignment(wrap_text=True)
    ws.cell(i, 1).border = thin
    ws.cell(i, 2).border = thin
    ws.row_dimensions[i].height = 32

ws["A27"] = "本文件由 analyze_311.py / diagnose.py 汇总后生成；原始 Parquet 仍只在 D:\\项目1\\原始数据。"
ws["A27"].font = small_font
ws.merge_cells("A27:G27")

ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 88
for col in "CDEFG":
    ws.column_dimensions[col].width = 14
ws.row_dimensions[1].height = 24
ws.freeze_panes = "A4"

# ===================== 2 总览 =====================
ws = wb.create_sheet("总览")
ws.sheet_properties.tabColor = TEAL
ws["A1"] = "数据总览"
ws["A1"].font = title_font
ws.merge_cells("A1:H1")
ws["A2"] = "分母：全样本 7,525,498 条请求记录；unique_key 无重复。以下 KPI 均对应该总体，除非另行标注完整月同比窗。"
ws["A2"].font = small_font
ws.merge_cells("A2:H2")

kpi_box(ws, 4, 1, "请求记录数", f"{n_all:,}", "unique_key 去重后相同", NAVY)
kpi_box(ws, 4, 3, "覆盖跨度", "730 天", "2024-09-07 至 2026-09-05", TEAL)
kpi_box(ws, 4, 5, "已关闭（status）", f"{int(status.loc[status.status=='Closed','n'].iloc[0]):,}", f"{int(status.loc[status.status=='Closed','n'].iloc[0])/n_all:.1%} 的记录", GREEN)
kpi_box(ws, 4, 7, "完整月同比", f"{pct_yoy:+.1%}", f"Y2 {n_y2:,} − Y1 {n_y1:,}", GOLD)

ws["A8"] = "状态构成（全样本）"
ws["A8"].font = h2_font
st = status.copy()
st["占比"] = st["n"] / st["n"].sum()
st.columns = ["status", "记录数", "占比"]
write_df(ws, st, 9, 1)

ws["D8"] = "渠道构成"
ws["D8"].font = h2_font
ch = channel.copy()
ch["占比"] = ch["n"] / ch["n"].sum()
ch.columns = ["channel", "记录数", "占比"]
write_df(ws, ch, 9, 4)

ws["G8"] = "行政区构成"
ws["G8"].font = h2_font
bo = borough.copy()
bo["占比"] = bo["n"] / bo["n"].sum()
bo.columns = ["borough", "记录数", "占比"]
write_df(ws, bo, 9, 7)

ws["A19"] = "受理机构 Top（全样本）"
ws["A19"].font = h2_font
ag = agency.copy()
ag["占比"] = ag["n"] / ag["n"].sum()
ag.columns = ["agency", "记录数", "占比"]
write_df(ws, ag.head(12), 20, 1)

ws["E19"] = "投诉类型 Top 15（全样本）"
ws["E19"].font = h2_font
tp = types.head(15).copy()
tp["占比"] = tp["n"] / n_all
tp.columns = ["complaint_type", "记录数", "占总样本"]
write_df(ws, tp, 20, 5)

ws["A35"] = (
    "阅读提示：全时段排名由高频日常类型主导——违停 1,159,053（15.4%）、住宅噪声 881,548（11.7%）、"
    "供暖/热水 651,594（8.7%）。供暖在夏季接近消失，因此“全年第三”会低估它在冬季的权重。看季节请转到「冬季与尖峰」。"
)
ws["A35"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A35:H35")
ws.row_dimensions[35].height = 40

ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 12
ws.column_dimensions["D"].width = 14
ws.column_dimensions["E"].width = 28
ws.column_dimensions["F"].width = 14
ws.column_dimensions["G"].width = 16
ws.column_dimensions["H"].width = 14
ws.freeze_panes = "A4"

# ===================== 3 核心发现 =====================
ws = wb.create_sheet("核心发现")
ws.sheet_properties.tabColor = GOLD
ws["A1"] = "核心发现（可回查）"
ws["A1"].font = title_font
ws.merge_cells("A1:F1")
ws["A2"] = "每条发现包含：指标定义、比较窗、分母、主要数字、证据层级。详细表在后续工作表。"
ws["A2"].font = small_font

findings = [
    (
        "1",
        "完整月同比多出约 34.7 万条记录（+10.5%），增量并不均匀",
        f"比较：Y1 {n_y1:,} 条（2024-10-01 至 2025-08-31）vs Y2 {n_y2:,} 条（2025-10-01 至 2026-08-31）。差值 {delta_yoy:,}。",
        "分解（互斥桶，加总等于总增量）：供暖/热水 +63,644（占增量 18.4%）；积雪/结冰 +55,130（15.9%）；"
        "违停 +54,955（15.8%）；路面状况 Street Condition +48,302（13.9%）；堵塞车道 +23,731（6.8%）；其余类型合计 +100,933（29.1%）。",
        "描述性证据充分。推断：积雪增量与更冷/多雪的冬天一致，但不能用 311 记录代替气象观测。"
        "Street Condition 大增未做独立审计，可能含真实路面问题、申报习惯或分类漂移。",
    ),
    (
        "2",
        "冬季把城市 311 的“产品组合”改写成供暖与天气工单",
        f"供暖/热水：冬季窗 Nov–Mar，Y1 冬季 {heat_w1:,} 条 / 当季全部 {all_w1:,}（占 {heat_w1/all_w1:.1%}）；"
        f"Y2 冬季 {heat_w2:,} / {all_w2:,}（占 {heat_w2/all_w2:.1%}）。",
        "积雪/结冰：Y1 冬 8,808 量级 vs Y2 冬约 6.4 万（完整月同比口径 8,808 → 63,938）。"
        "单日尖峰 2026-02-24：当天 22,805 条中 11,370 条为 Snow or Ice（49.9%）。"
        "2025-01 住宅噪声异常高（70,837）而 2026-01 仅 28,739；跨年那一周 2025 年多日噪声 3,000–6,000/天，2026 年同期约 750–1,400/天。",
        "描述：季节与尖峰结构清楚。推断：2026 年 1 月噪声偏低可能与更恶劣冬季活动减少有关，这是假说，不是验证过的因果。",
    ),
    (
        "3",
        "结案时钟是双速的：NYPD 小时级，住房/建筑类按天到周",
        "NYPD 3,403,951 条，有效关闭 3,402,671，中位 1.34 小时，P90 7.0 小时（违停、住宅噪声、堵塞车道同属这一档）。",
        "HPD 1,633,583 条，有效关闭 1,573,865，中位 87.5 小时（约 3.6 天），P90 890 小时（约 37 天）。"
        "供暖/热水中位 41.8 小时、P90 84.2 小时；不卫生条件中位 250 小时（约 10.4 天）。"
        "布朗克斯占完整月样本供暖工单的 35.1%，但其全部工单只占全市 22.1%——供暖负担集中，不是总体量幻觉。",
        "描述：关闭速度差数量级，且与机构职责匹配（出警 vs 住房维修）。不要把中位时长直接当 KPI 排名：工单定义、法定时限、是否需入户都不同。",
    ),
    (
        "4",
        "最近两个月的“结案变快/未结案变多”首先是测量问题",
        "2026-08 创建 328,446 条，快照时仍开放类状态 53,158（16.2%）；2026-09 残缺月 41,974 条中开放类 17,412（41.5%）。"
        "对比 2025-01 开放类仅 1,832 / 348,180（0.5%）。2026-08 已关闭工单的 P90 结案时长降到 178 小时，低于此前多数月份的 300–500 小时——这是慢单尚未关闭，不是突然提速。",
        "分类更名：Water System 与 Sewer 在 2026-08 降为 0，Water Maintenance / Sewer Maintenance 承接。"
        "7 月 Water System 16,261 + Water Maintenance 1,892；8 月 0 + 8,925。用类型名做跨月趋势时，8 月起必须合并新旧名，否则会误报“供水投诉消失”。",
        "这是 metric-diagnostics 意义上的测量驱动：日志/分类/窗口截断可以单独制造指标运动。",
    ),
]

r = 4
for num, title, line1, line2, line3 in findings:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws.cell(r, 1, f"发现 {num}  {title}")
    c.font = Font(name="Calibri", bold=True, color=WHITE, size=12)
    c.fill = fill(NAVY)
    ws.row_dimensions[r].height = 20
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 3, end_column=6)
    body = ws.cell(r + 1, 1, line1 + "\n" + line2 + "\n" + line3)
    body.alignment = Alignment(wrap_text=True, vertical="top")
    body.font = body_font
    ws.row_dimensions[r + 1].height = 36
    ws.row_dimensions[r + 2].height = 48
    ws.row_dimensions[r + 3].height = 36
    r += 5

ws["A25"] = "刻意没有做成“结论”的观察"
ws["A25"].font = h2_font
ws["A26"] = (
    "Drug Activity 完整月同比 27,869 → 8,955（−68%），Lead 27,623 → 13,653。降幅大，但缺少分类规则或执法口径材料，本分析不解释原因。\n"
    "EDC 未结案中位年龄约 304 天（多为 Noise - Helicopter 一类长尾），更像工单流程/关闭规则问题，不宜并入“城市响应变慢”总判断。"
)
ws["A26"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A26:F26")
ws.row_dimensions[26].height = 48

ws.column_dimensions["A"].width = 22
for col in "BCDEF":
    ws.column_dimensions[col].width = 18

# ===================== 4 月度趋势 =====================
ws = wb.create_sheet("月度趋势")
ws.sheet_properties.tabColor = TEAL
ws["A1"] = "按创建月的请求量与结构"
ws["A1"].font = title_font
ws["A2"] = "n_created 是当月创建的记录数。开放类占比升高出现在样本末段，优先视为快照截断。2024-09、2026-09 不是完整月。"
ws["A2"].font = small_font
ws.merge_cells("A2:L2")

md = pd.DataFrame({
    "创建月": m["created_month"],
    "创建记录数": m["n_created"],
    "status=Closed": m["n_closed_status"],
    "开放类状态": m["n_openish"],
    "开放类占比": m["open_share"],
    "有效关闭数": m["n_valid_close"],
    "中位结案小时": m["median_close_hours"],
    "P90结案小时": m["p90_close_hours"],
    "供暖/热水": m["n_heat"],
    "住宅噪声": m["n_noise_res"],
    "违停": m["n_illegal_parking"],
    "供暖占当月": m["heat_share"],
    "备注": m["note"],
})
write_df(ws, md, 4, 1)
last = 4 + len(md)

chart = LineChart()
chart.title = "月度创建量（记录数）"
chart.y_axis.title = "记录数"
chart.x_axis.title = "创建月"
chart.height = 8
chart.width = 18
chart.style = 10
data = Reference(ws, min_col=2, min_row=4, max_row=last)
cats = Reference(ws, min_col=1, min_row=5, max_row=last)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
ws.add_chart(chart, "A32")

chart2 = LineChart()
chart2.title = "供暖 / 住宅噪声 / 违停（月度）"
chart2.y_axis.title = "记录数"
chart2.height = 8
chart2.width = 18
chart2.style = 12
data2 = Reference(ws, min_col=9, min_row=4, max_col=11, max_row=last)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats)
ws.add_chart(chart2, "A48")

ws["A64"] = "图阅读：1 月是全年高峰（2025-01 与 2026-01 总量几乎相同：348,180 vs 348,511），但内部结构不同——2026-01 供暖更高、住宅噪声更低。2025-02 偏低部分来自短月。2026-08/09 开放类占比跳升不要当作业绩恶化。"
ws["A64"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A64:L64")
ws.row_dimensions[64].height = 36

for i, col in enumerate(md.columns, start=1):
    ws.column_dimensions[get_column_letter(i)].width = 14 if i > 1 else 12
ws.column_dimensions["M"].width = 28
# format percent cols
for row in range(5, last + 1):
    ws.cell(row, 5).number_format = "0.0%"
    ws.cell(row, 12).number_format = "0.0%"
ws.freeze_panes = "A5"
ws.auto_filter.ref = f"A4:M{last}"

# ===================== 5 同比分解 =====================
ws = wb.create_sheet("同比分解")
ws.sheet_properties.tabColor = GREEN
ws["A1"] = "完整月同比分解"
ws["A1"].font = title_font
ws["A2"] = (
    f"Y1：2024-10-01 ≤ created < 2025-09-01，n={n_y1:,}。  "
    f"Y2：2025-10-01 ≤ created < 2026-09-01，n={n_y2:,}。  "
    f"Δ={delta_yoy:,}（{pct_yoy:+.1%}）。两个窗口各 11 个日历月，长度对齐。"
)
ws["A2"].font = body_font
ws.merge_cells("A2:H2")

ws["A4"] = "互斥贡献桶（加总 = 总增量）"
ws["A4"].font = h2_font
cb = contrib.copy()
cb.columns = ["类型桶", "Y1记录数", "Y2记录数", "增量", "Y1占比", "Y2占比", "占总增量"]
write_df(ws, cb, 5, 1)
for row in range(6, 6 + len(cb)):
    ws.cell(row, 5).number_format = "0.0%"
    ws.cell(row, 6).number_format = "0.0%"
    ws.cell(row, 7).number_format = "0.0%"

bar = BarChart()
bar.type = "bar"
bar.title = "各类型桶对总增量的贡献（记录数）"
bar.y_axis.title = None
bar.x_axis.title = "增量（Y2−Y1）"
bar.height = 8
bar.width = 16
data = Reference(ws, min_col=4, min_row=5, max_row=5 + len(cb))
cats = Reference(ws, min_col=1, min_row=6, max_row=5 + len(cb))
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
bar.shape = 4
ws.add_chart(bar, "A14")

ws["A30"] = "增量最大的具体类型（阈值：两期合计 ≥ 5,000）"
ws["A30"].font = h2_font
top_up = yt.head(15)[["complaint_type", "n_y1", "n_y2", "delta", "pct", "share_of_delta"]].copy()
top_up.columns = ["complaint_type", "Y1", "Y2", "增量", "相对Y1", "占总增量"]
write_df(ws, top_up, 31, 1, GREEN)
for row in range(32, 32 + len(top_up)):
    ws.cell(row, 5).number_format = "0.0%"
    ws.cell(row, 6).number_format = "0.0%"

ws["H30"] = "降幅最大的具体类型"
ws["H30"].font = h2_font
top_dn = yt.sort_values("delta").head(12)[["complaint_type", "n_y1", "n_y2", "delta", "pct"]].copy()
top_dn.columns = ["complaint_type", "Y1", "Y2", "增量", "相对Y1"]
write_df(ws, top_dn, 31, 8, GOLD)
for row in range(32, 32 + len(top_dn)):
    ws.cell(row, 12).number_format = "0.0%"

ws["A49"] = "按行政区（同一完整月窗口）"
ws["A49"].font = h2_font
yb = yoy_boro.copy()
yb["delta"] = yb["n_y2"] - yb["n_y1"]
yb["pct"] = yb["delta"] / yb["n_y1"]
yb.columns = ["borough", "Y1", "Y2", "增量", "相对Y1"]
write_df(ws, yb, 50, 1)
for row in range(51, 51 + len(yb)):
    ws.cell(row, 5).number_format = "0.0%"

ws["G49"] = "按机构（同一窗口，按增量降序）"
ws["G49"].font = h2_font
ya = yoy_ag.copy()
ya["delta"] = ya["n_y2"] - ya["n_y1"]
ya["pct"] = ya.apply(lambda r: (r["delta"] / r["n_y1"]) if r["n_y1"] else None, axis=1)
ya.columns = ["agency", "Y1", "Y2", "增量", "相对Y1"]
write_df(ws, ya.head(12), 50, 7)
for row in range(51, 51 + 12):
    ws.cell(row, 11).number_format = "0.0%"

ws["A60"] = (
    "分解读法：HPD 增量（+146,946）与供暖/住房类上升同向；DOT 增量（+66,856）与 Street Condition 同向；"
    "NYPD 增量较小（+58,500，+3.9%）。史泰登岛相对增幅最高（+23.8%），但基数小。"
    "OOS、NYC311-PRD 在 Y1 为 0，更像新编码而非业务从零开始。"
)
ws["A60"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A60:L60")
ws.row_dimensions[60].height = 40

for i in range(1, 13):
    ws.column_dimensions[get_column_letter(i)].width = 16
ws.column_dimensions["A"].width = 22
ws.column_dimensions["H"].width = 28
ws.freeze_panes = "A4"

# ===================== 6 冬季与尖峰 =====================
ws = wb.create_sheet("冬季与尖峰")
ws.sheet_properties.tabColor = "C00000"
ws["A1"] = "冬季结构与天气尖峰"
ws["A1"].font = title_font
ws["A2"] = "供暖法定季在纽约大致为 10 月–5 月；本表用记录自证季节性，不引入外部气温序列（任务限制）。"
ws["A2"].font = small_font

ws["A4"] = "冬季窗对照 Nov–Mar"
ws["A4"].font = h2_font
winter_tbl = pd.DataFrame({
    "窗口": ["Y1冬 2024-11～2025-03", "Y2冬 2025-11～2026-03"],
    "全部记录": [all_w1, all_w2],
    "供暖/热水": [heat_w1, heat_w2],
    "供暖占当季": [heat_w1 / all_w1, heat_w2 / all_w2],
    "积雪/结冰(同比口径)": [8808, 63938],
})
write_df(ws, winter_tbl, 5, 1)
ws.cell(6, 4).number_format = "0.0%"
ws.cell(7, 4).number_format = "0.0%"

ws["A10"] = "月度：供暖、积雪、违停、住宅噪声"
ws["A10"].font = h2_font
snow_m = snow.rename(columns={"m": "created_month", "n": "积雪/结冰"})
mix = m[["created_month", "n_created", "n_heat", "n_noise_res", "n_illegal_parking"]].merge(
    snow_m, on="created_month", how="left"
)
mix["积雪/结冰"] = mix["积雪/结冰"].fillna(0).astype(int)
mix.columns = ["创建月", "全部", "供暖/热水", "住宅噪声", "违停", "积雪/结冰"]
write_df(ws, mix, 11, 1)
mix_last = 11 + len(mix)

line = LineChart()
line.title = "冬季相关类型月度记录数"
line.height = 8
line.width = 16
data = Reference(ws, min_col=3, min_row=11, max_col=6, max_row=mix_last)
cats = Reference(ws, min_col=1, min_row=12, max_row=mix_last)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
ws.add_chart(line, "H10")

ws["A40"] = "尖峰日 2026-02-24（当天 22,805 条，全样本最高日）"
ws["A40"].font = h2_font
f24 = feb24.head(12).copy()
f24["占当天"] = f24["n"] / 22805
f24.columns = ["complaint_type", "记录数", "占当天"]
write_df(ws, f24, 41, 1)
for row in range(42, 42 + len(f24)):
    ws.cell(row, 3).number_format = "0.0%"

ws["E40"] = "前后窗口（2026-02-15～03-04）"
ws["E40"].font = h2_font
fw = febwin.copy()
fw.columns = ["日期", "全部", "供暖", "积雪", "违停", "含Noise类"]
write_df(ws, fw, 41, 5)

ws["A56"] = "跨年对照：住宅噪声（不是全市总量）"
ws["A56"].font = h2_font
ny = newyear.copy()
ny["年份组"] = ny["d"].astype(str).str.slice(0, 4)
ny.columns = ["日期", "全部", "住宅噪声", "供暖", "违停", "年份组"]
write_df(ws, ny, 57, 1)

ws["A78"] = (
    "2026-02-24 的增量几乎全是积雪工单：当天 Snow or Ice 11,370，而 02-23 为 4,585、02-25 为 3,719（见右表）。"
    "这是典型的天气事件尖峰，会抬高月度总量与 DOT/DSNY 相关类型。"
    "跨年噪声：2025-01-05 住宅噪声 6,053，2026-01-05 仅 750。若只比较两个 1 月总量会得出“城市一样忙”，内部已经换了产品。"
)
ws["A78"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A78:J78")
ws.row_dimensions[78].height = 48

for i in range(1, 12):
    ws.column_dimensions[get_column_letter(i)].width = 16
ws.column_dimensions["A"].width = 22
ws.freeze_panes = "A5"

# ===================== 7 结案时长 =====================
ws = wb.create_sheet("结案时长")
ws.sheet_properties.tabColor = "7030A0"
ws["A1"] = "结案时长：双速系统"
ws["A1"].font = title_font
ws["A2"] = (
    "样本：status=Closed 且 closed_ts≥created_ts。中位数对长尾稳健；P90 显示尾部。"
    "不含未关闭工单（右截断）。因此机构未结案率高时，中位时长偏向已关掉的快单。"
)
ws["A2"].font = small_font
ws.merge_cells("A2:J2")

ws["A4"] = "按机构（创建量 ≥ 5,000）"
ws["A4"].font = h2_font
cad = ca[["agency", "n_created", "n_valid_close", "n_openish", "open_rate", "median_close_hours", "median_days", "p90_close_hours"]].copy()
cad.columns = ["agency", "创建数", "有效关闭", "开放类", "开放类占比", "中位小时", "中位天", "P90小时"]
write_df(ws, cad, 5, 1)
for row in range(6, 6 + len(cad)):
    ws.cell(row, 5).number_format = "0.0%"

bar = BarChart()
bar.type = "bar"
bar.title = "机构中位结案小时（对数观感：NYPD 与 EDC 不在同一量级）"
bar.height = 9
bar.width = 14
data = Reference(ws, min_col=6, min_row=5, max_row=5 + len(cad))
cats = Reference(ws, min_col=1, min_row=6, max_row=5 + len(cad))
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
ws.add_chart(bar, "J4")

ws["A21"] = "量大类型的关闭速度（创建量 ≥ 20,000，取前 18）"
ws["A21"].font = h2_font
ctd = ct[["complaint_type", "n_created", "n_valid_close", "close_rate", "median_close_hours", "median_days", "p90_close_hours", "p90_days"]].copy()
ctd.columns = ["complaint_type", "创建数", "有效关闭", "有效关闭率", "中位小时", "中位天", "P90小时", "P90天"]
write_df(ws, ctd, 22, 1)
for row in range(23, 23 + len(ctd)):
    ws.cell(row, 4).number_format = "0.0%"

ws["A43"] = "供暖/热水：按创建月的关闭中位数（2026-08/09 仍受截断）"
ws["A43"].font = h2_font
hc = heat_close.copy()
hc["关闭率"] = hc["n_closed"] / hc["n"]
hc.columns = ["创建月", "供暖创建", "有效关闭", "中位小时", "P90小时", "关闭率"]
write_df(ws, hc, 44, 1)
for row in range(45, 45 + len(hc)):
    ws.cell(row, 6).number_format = "0.0%"

hc_last = 44 + len(hc)
line = LineChart()
line.title = "供暖工单中位结案小时"
line.height = 7
line.width = 14
data = Reference(ws, min_col=4, min_row=44, max_row=hc_last)
cats = Reference(ws, min_col=1, min_row=45, max_row=hc_last)
line.add_data(data, titles_from_data=True)
line.set_categories(cats)
ws.add_chart(line, "H43")

ws["A73"] = (
    "供暖中位关闭在供暖高峰月升到约 46–52 小时（2024-12、2025-01、2025-12、2026-01/02），"
    "淡季约 30–36 小时。高峰期变慢与工单量上升同时发生，符合排队压力的描述，但没有现场作业数据，不能分离“更难修”与“人不够”。"
)
ws["A73"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A73:J73")
ws.row_dimensions[73].height = 40

for i in range(1, 10):
    ws.column_dimensions[get_column_letter(i)].width = 14
ws.column_dimensions["A"].width = 28
ws.freeze_panes = "A5"

# ===================== 8 行政区 =====================
ws = wb.create_sheet("行政区")
ws.sheet_properties.tabColor = "548235"
ws["A1"] = "行政区构成（完整月 2024-10-01～2026-08-31）"
ws["A1"].font = title_font
ws["A2"] = "无人口分母，不做人均。Unspecified 已从本表排除。heat_share_city 是该区占全市供暖工单的份额。"
ws["A2"].font = small_font

bd = bs.copy()
bd = bd.rename(columns={
    "borough": "borough",
    "n_heat": "供暖",
    "n_all": "全部",
    "n_park": "违停",
    "n_noise": "住宅噪声",
    "all_share": "占全市工单",
    "heat_share_in_boro": "区内供暖占比",
    "heat_share_city": "占全市供暖",
    "park_share_in_boro": "区内违停占比",
    "noise_share_in_boro": "区内住宅噪声占比",
})
write_df(ws, bd, 4, 1)
for row in range(5, 5 + len(bd)):
    for col in range(6, 11):
        ws.cell(row, col).number_format = "0.0%"

bar = BarChart()
bar.type = "col"
bar.grouping = "clustered"
bar.title = "五区：供暖 / 违停 / 住宅噪声 记录数"
bar.height = 8
bar.width = 14
data = Reference(ws, min_col=2, min_row=4, max_col=5, max_row=4 + len(bd))
cats = Reference(ws, min_col=1, min_row=5, max_row=4 + len(bd))
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
ws.add_chart(bar, "A12")

ws["A28"] = "全样本行政区结案中位（受类型组合影响：布朗克斯噪声/供暖多，中位时长会被住房类拉长）"
ws["A28"].font = h2_font
cb = pd.read_csv(SUM / "close_by_borough.csv")
cb["中位天"] = cb["median_close_hours"] / 24
cb.columns = ["borough", "创建数", "有效关闭", "中位小时", "P90小时", "中位天"]
write_df(ws, cb, 29, 1)

ws["A38"] = "星期结构（DuckDB dow：0=周日）"
ws["A38"].font = h2_font
dw = dow[["星期", "n_all", "n_noise_res", "n_parking", "n_heat"]].copy()
dw.columns = ["星期", "全部", "住宅噪声", "违停", "供暖"]
# reorder Sun-Sat
order_w = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"]
dw["星期"] = pd.Categorical(dw["星期"], order_w, ordered=True)
dw = dw.sort_values("星期")
write_df(ws, dw, 39, 1)

ws["A49"] = (
    "布朗克斯：占全市工单 22.1%，却占全市供暖 35.1%、住宅噪声也明显高于其总体份额。"
    "布鲁克林/皇后：违停是区内第一产品。周末（六、日）住宅噪声约为工作日的 1.5–1.9 倍，违停则偏工作日。"
    "这是申告行为的时间节奏，不是“周末问题更严重”的因果证明。"
)
ws["A49"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A49:J49")
ws.row_dimensions[49].height = 48

for i in range(1, 11):
    ws.column_dimensions[get_column_letter(i)].width = 18
ws.freeze_panes = "A5"

# ===================== 9 数据质量 =====================
ws = wb.create_sheet("数据质量")
ws.sheet_properties.tabColor = "833C0C"
ws["A1"] = "哪些质量问题会改变结论"
ws["A1"].font = title_font

ws["A3"] = "全样本质量计数"
ws["A3"].font = h2_font
qrows = [
    ("记录数", n_all, "分母"),
    ("unique_key 去重后", n_keys, "无重复主键，记录级分析可信"),
    ("created 无法解析", int(overview.n_created_null.iloc[0]), "为 0，时间窗可靠"),
    ("agency / complaint_type 空", int(overview.n_agency_blank.iloc[0]) + int(overview.n_type_blank.iloc[0]), "为 0"),
    ("closed_ts 空", int(overview.n_closed_null.iloc[0]), "约 3.2%；含未关闭与未填关闭时间"),
    ("关闭早于创建", int(overview.n_closed_before_created.iloc[0]), "0.025%，结案时长已排除"),
    ("行政区 Unspecified/空", int(overview.n_borough_unspecified.iloc[0]), "0.08%，分区分析可忽略"),
    ("渠道 UNKNOWN", int(overview.n_channel_unknown.iloc[0]), "7.9%，渠道趋势需谨慎"),
]
ws["A4"] = "检查项"
ws["B4"] = "计数"
ws["C4"] = "对分析的影响"
style_header_row(ws, 4, 1, 3)
for i, (a, b, c) in enumerate(qrows, start=5):
    ws.cell(i, 1, a).border = thin
    ws.cell(i, 2, b).border = thin
    ws.cell(i, 2).number_format = "#,##0"
    ws.cell(i, 3, c).border = thin

ws["A15"] = "会真正扭曲趋势的两件事"
ws["A15"].font = h2_font
ws["A16"] = "A. 快照截断（近期未结案）"
ws["A16"].font = Font(name="Calibri", bold=True, color=GOLD)
oa = open_age.copy()
oa.columns = ["agency", "开放类数量", "中位年龄(天)", "P90年龄(天)"]
write_df(ws, oa, 17, 1)

ws["F16"] = "B. 2026-08 供水/下水道类型更名"
ws["F16"].font = Font(name="Calibri", bold=True, color=GOLD)
wt = water.copy()
wt.columns = ["月", "Water System", "Water Maintenance", "Sewer", "Sewer Maintenance"]
write_df(ws, wt, 17, 6)

ws["A33"] = (
    "实务规则：1）比较关闭速度不要用 2026-07 以后的创建月。2）做 Water/Sewer 趋势时从 2026-08 起合并新旧类型名。"
    "3）2024-09 与 2026-09 不得进入月环比。4）渠道 UNKNOWN 在 2026-03 跳到 47,359，渠道份额序列从该月起不稳定。"
)
ws["A33"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A33:K33")
ws.row_dimensions[33].height = 40

ws["A35"] = "分片行数（应与月度 n_created 一致；2024-09 / 2026-09 为残缺分片）"
ws["A35"].font = h2_font
fc = pd.read_csv(SUM / "file_row_counts.csv")
fc["file"] = fc["filename"].str.replace(r".*\\", "", regex=True)
fc = fc[["file", "n"]]
fc.columns = ["parquet 分片", "行数"]
write_df(ws, fc, 36, 1)

ws.column_dimensions["A"].width = 36
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 36
ws.column_dimensions["D"].width = 16
ws.column_dimensions["F"].width = 14
for col in "GHIJK":
    ws.column_dimensions[col].width = 20
ws.freeze_panes = "A4"

# ===================== 10 口径 =====================
ws = wb.create_sheet("口径与假设")
ws.sheet_properties.tabColor = "7F7F7F"
ws["A1"] = "口径、假设、复现"
ws["A1"].font = title_font
ws["A3"] = "最小假设"
ws["A3"].font = h2_font
assumps = [
    "抽取时点按最后分片命名取 2026-09-07；库内最大创建时间为 2026-09-05 01:50:33。未结案年龄按 2026-09-07 计算，差一两天不改变排序。",
    "开放类 = Open/In Progress/Assigned/Pending/Started。Unspecified（302 条）不计入开放也不计入关闭。",
    "完整月同比不用日历年，而用各 11 个月的对齐窗口，避免残缺月和“是否含 9 月”争议。",
    "不引入人口、气温、执法编制等外部数据（任务仅授权原始 311 与 input 清单）。因此没有人均、没有天气回归。",
    "同一 unique_key 只出现一次；无法在本数据中识别“同一公寓的重复供暖投诉是否同一事件”。计数始终是请求记录。",
    "中文说明中的机构中文名仅作阅读辅助：HPD=住房保护，NYPD=警察，DSNY=卫生，DOT=交通，DEP=环保，DPR=公园。分析键仍是原始英文代码。",
]
for i, t in enumerate(assumps, start=4):
    ws.cell(i, 1, f"{i-3}. {t}").alignment = Alignment(wrap_text=True)
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)
    ws.row_dimensions[i].height = 32

ws["A11"] = "复现"
ws["A11"].font = h2_font
ws["A12"] = (
    "1. python analyze_311.py  → summaries/*.csv\n"
    "2. python diagnose.py     → summaries/diag_*.csv\n"
    "3. python build_excel.py  → 最终成果.xlsx\n"
    "原始文件只读自 D:\\项目1\\原始数据；DuckDB read_parquet 扫描，不落原始副本。"
)
ws["A12"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A12:F12")
ws.row_dimensions[12].height = 64

ws["A14"] = "关键数字核对清单（生成后应仍成立）"
ws["A14"].font = h2_font
checks = [
    ("全样本记录", n_all, 7525498),
    ("Y1 完整月", n_y1, 3293003),
    ("Y2 完整月", n_y2, 3639698),
    ("分片行数合计", int(fc["行数"].sum()), n_all),
]
ws["A15"] = "项目"
ws["B15"] = "本次计算"
ws["C15"] = "预期锚点"
ws["D15"] = "一致?"
style_header_row(ws, 15, 1, 4)
for i, (name, got, exp) in enumerate(checks, start=16):
    ws.cell(i, 1, name).border = thin
    ws.cell(i, 2, got).border = thin
    ws.cell(i, 2).number_format = "#,##0"
    ws.cell(i, 3, exp).border = thin
    ws.cell(i, 3).number_format = "#,##0"
    ok = "是" if got == exp else "否"
    ws.cell(i, 4, ok).border = thin
    ws.cell(i, 4).fill = fill(LIGHT_GREEN if ok == "是" else "F4CCCC")

ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 12
for col in "EF":
    ws.column_dimensions[col].width = 18

wb.save(OUT)
print("saved", OUT, "sheets", wb.sheetnames)
print("file_rows", int(fc["行数"].sum()), "n_all", n_all)
print("yoy", n_y1, n_y2, delta_yoy)
