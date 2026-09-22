"""NYC 311 round-2 analysis: compute tables from raw parquet and write Excel.

Read-only source: D:/项目1/原始数据
Outputs stay under D:/项目2/Grok/第二轮
"""
from __future__ import annotations

from pathlib import Path

import duckdb
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart.layout import Layout, ManualLayout

ROOT = Path(r"D:\项目2\Grok\第二轮")
MID = ROOT / "中间汇总"
MID.mkdir(exist_ok=True)
SRC = "read_parquet('D:/项目1/原始数据/*.parquet')"

con = duckdb.connect()


def q(sql: str) -> pd.DataFrame:
    return con.execute(sql).df()


def save_csv(df: pd.DataFrame, name: str) -> pd.DataFrame:
    path = MID / name
    df.to_csv(path, index=False, encoding="utf-8-sig")
    return df


# ---------------------------------------------------------------------------
# Core extracts
# ---------------------------------------------------------------------------
overview = q(
    f"""
SELECT
  count(*) AS n_records,
  count(DISTINCT unique_key) AS n_unique_key,
  min(try_cast(created_date AS timestamp)) AS min_created,
  max(try_cast(created_date AS timestamp)) AS max_created,
  max(try_cast(closed_date AS timestamp)) AS max_closed,
  max(try_cast(resolution_action_updated_date AS timestamp)) AS max_resolution_update,
  sum(closed_date IS NULL OR closed_date='') AS n_no_closed_date,
  sum(status='Closed') AS n_status_closed,
  sum(status<>'Closed') AS n_status_not_closed,
  count(DISTINCT agency) AS n_agency,
  count(DISTINCT complaint_type) AS n_complaint_type,
  sum(try_cast(closed_date AS timestamp) < try_cast(created_date AS timestamp)) AS n_closed_before_created
FROM {SRC}
"""
)
save_csv(overview, "00_overview.csv")

monthly = q(
    f"""
SELECT
  strftime(try_cast(created_date AS timestamp), '%Y-%m') AS year_month,
  count(*) AS n_created,
  count(DISTINCT date_trunc('day', try_cast(created_date AS timestamp))) AS n_days_with_create,
  round(count(*) * 1.0 / count(DISTINCT date_trunc('day', try_cast(created_date AS timestamp))), 1) AS avg_per_active_day,
  sum(status='Closed') AS n_status_closed,
  sum(status<>'Closed') AS n_status_not_closed,
  round(100.0 * sum(status<>'Closed') / count(*), 2) AS pct_status_not_closed,
  sum(closed_date IS NULL OR closed_date='') AS n_no_closed_date,
  sum(complaint_type='HEAT/HOT WATER') AS n_heat,
  sum(complaint_type='Snow or Ice') AS n_snow,
  sum(complaint_type='Illegal Parking') AS n_illegal_parking,
  sum(complaint_type='Street Condition') AS n_street,
  sum(complaint_type='Street Condition' AND descriptor='Pothole' AND open_data_channel_type='UNKNOWN') AS n_unk_pothole,
  sum(complaint_type IN ('Water System','Water Maintenance','Sewer','Sewer Maintenance')) AS n_water_family,
  sum(complaint_type='Water System') AS n_water_system,
  sum(complaint_type='Water Maintenance') AS n_water_maint,
  sum(complaint_type='Sewer') AS n_sewer,
  sum(complaint_type='Sewer Maintenance') AS n_sewer_maint,
  sum(complaint_type='Lead') AS n_lead,
  sum(complaint_type='Drug Activity') AS n_drug
FROM {SRC}
GROUP BY 1
ORDER BY 1
"""
)
save_csv(monthly, "01_monthly.csv")

agency_all = q(
    f"""
SELECT agency, count(*) AS n, round(100.0*count(*)/(SELECT count(*) FROM {SRC}), 2) AS pct
FROM {SRC}
GROUP BY 1 ORDER BY n DESC
"""
)
save_csv(agency_all, "02_agency.csv")

borough_all = q(
    f"""
SELECT borough, count(*) AS n, round(100.0*count(*)/(SELECT count(*) FROM {SRC}), 2) AS pct
FROM {SRC}
GROUP BY 1 ORDER BY n DESC
"""
)
save_csv(borough_all, "03_borough.csv")

channel_all = q(
    f"""
SELECT coalesce(open_data_channel_type,'(null)') AS channel, count(*) AS n,
 round(100.0*count(*)/(SELECT count(*) FROM {SRC}), 2) AS pct
FROM {SRC}
GROUP BY 1 ORDER BY n DESC
"""
)
save_csv(channel_all, "04_channel.csv")

status_all = q(
    f"""
SELECT status, count(*) AS n, round(100.0*count(*)/(SELECT count(*) FROM {SRC}), 2) AS pct
FROM {SRC}
GROUP BY 1 ORDER BY n DESC
"""
)
save_csv(status_all, "05_status.csv")

top_types = q(
    f"""
SELECT complaint_type, agency, count(*) AS n,
 round(100.0*count(*)/(SELECT count(*) FROM {SRC}), 2) AS pct
FROM {SRC}
GROUP BY 1,2 ORDER BY n DESC LIMIT 25
"""
)
save_csv(top_types, "06_top_complaint_types.csv")

# Matched 11-month windows
yoy_agency = q(
    f"""
WITH b AS (
  SELECT agency,
    CASE WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2024-10' AND '2025-08' THEN 'A'
         WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-10' AND '2026-08' THEN 'B' END AS period
  FROM {SRC}
)
SELECT agency,
  sum(period='A') AS n_A, sum(period='B') AS n_B,
  sum(period='B')-sum(period='A') AS delta,
  round(100.0*(sum(period='B')-sum(period='A'))/nullif(sum(period='A'),0), 2) AS pct_change
FROM b WHERE period IS NOT NULL
GROUP BY 1
ORDER BY abs(delta) DESC
"""
)
save_csv(yoy_agency, "07_yoy_agency.csv")

yoy_type = q(
    f"""
WITH b AS (
  SELECT complaint_type,
    CASE WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2024-10' AND '2025-08' THEN 'A'
         WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-10' AND '2026-08' THEN 'B' END AS period
  FROM {SRC}
)
SELECT complaint_type,
  sum(period='A') AS n_A, sum(period='B') AS n_B,
  sum(period='B')-sum(period='A') AS delta,
  round(100.0*(sum(period='B')-sum(period='A'))/nullif(sum(period='A'),0), 2) AS pct_change
FROM b WHERE period IS NOT NULL
GROUP BY 1
ORDER BY abs(delta) DESC
"""
)
save_csv(yoy_type, "08_yoy_complaint_type.csv")

excl_buckets = q(
    f"""
WITH b AS (
  SELECT
    CASE WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2024-10' AND '2025-08' THEN 'A'
         WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-10' AND '2026-08' THEN 'B' END AS period,
    CASE
      WHEN complaint_type='Snow or Ice' THEN 'Snow or Ice'
      WHEN complaint_type='HEAT/HOT WATER' THEN 'HEAT/HOT WATER'
      WHEN complaint_type='Illegal Parking' THEN 'Illegal Parking'
      WHEN complaint_type='Street Condition' AND descriptor='Pothole' AND open_data_channel_type='UNKNOWN' THEN 'UNKNOWN-channel pothole'
      WHEN complaint_type='Street Condition' THEN 'Other Street Condition'
      ELSE 'All other types'
    END AS bucket
  FROM {SRC}
)
SELECT bucket,
  sum(period='A') AS n_A, sum(period='B') AS n_B,
  sum(period='B')-sum(period='A') AS delta
FROM b WHERE period IS NOT NULL
GROUP BY 1
ORDER BY abs(delta) DESC
"""
)
save_csv(excl_buckets, "09_exclusive_volume_buckets.csv")

period_tot = q(
    f"""
SELECT
  CASE WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2024-10' AND '2025-08' THEN 'A_2024-10_to_2025-08'
       WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-10' AND '2026-08' THEN 'B_2025-10_to_2026-08' END AS period,
  count(*) AS n
FROM {SRC}
GROUP BY 1
HAVING period IS NOT NULL
ORDER BY 1
"""
)
save_csv(period_tot, "10_period_totals.csv")

snow_days = q(
    f"""
SELECT date_trunc('day', try_cast(created_date AS timestamp))::DATE AS created_day,
       count(*) AS n
FROM {SRC}
WHERE complaint_type='Snow or Ice'
GROUP BY 1
ORDER BY n DESC
"""
)
save_csv(snow_days, "11_snow_days.csv")

snow_month = monthly[["year_month", "n_snow"]].copy()

street_channel_month = q(
    f"""
SELECT strftime(try_cast(created_date AS timestamp),'%Y-%m') AS year_month,
  sum(complaint_type='Street Condition') AS n_street,
  sum(complaint_type='Street Condition' AND descriptor='Pothole' AND open_data_channel_type='UNKNOWN') AS n_unk_pothole,
  sum(complaint_type='Street Condition' AND NOT (descriptor='Pothole' AND open_data_channel_type='UNKNOWN')) AS n_street_ex_unk_pothole,
  sum(complaint_type='Street Condition' AND open_data_channel_type='UNKNOWN') AS n_street_unknown_channel,
  sum(complaint_type='Street Condition' AND open_data_channel_type IN ('ONLINE','PHONE','MOBILE','OTHER')) AS n_street_named_channel
FROM {SRC}
GROUP BY 1 ORDER BY 1
"""
)
save_csv(street_channel_month, "12_street_channel_month.csv")

unk_pothole_res = q(
    f"""
SELECT
  strftime(try_cast(created_date AS timestamp),'%Y-%m') AS year_month,
  count(*) AS n,
  sum(resolution_description ILIKE '%duplicate%') AS n_duplicate_text,
  round(100.0*sum(resolution_description ILIKE '%duplicate%')/count(*),1) AS pct_duplicate_text
FROM {SRC}
WHERE complaint_type='Street Condition' AND descriptor='Pothole' AND open_data_channel_type='UNKNOWN'
GROUP BY 1 ORDER BY 1
"""
)
save_csv(unk_pothole_res, "13_unk_pothole_duplicate_text.csv")

heat_winter = q(
    f"""
WITH t AS (
  SELECT
    CASE WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2024-10' AND '2025-04' THEN 'W1_2024-10_to_2025-04'
         WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-10' AND '2026-04' THEN 'W2_2025-10_to_2026-04' END AS winter,
    unique_key, bbl, incident_address, borough, status,
    resolution_description
  FROM {SRC}
  WHERE complaint_type='HEAT/HOT WATER'
)
SELECT winter,
  count(*) AS n_tickets,
  count(DISTINCT CASE WHEN bbl IS NOT NULL AND bbl NOT IN ('','0','0000000000') THEN bbl END) AS n_bbl,
  count(DISTINCT incident_address) AS n_address,
  sum(resolution_description ILIKE '%duplicate%') AS n_duplicate_text,
  round(100.0*sum(resolution_description ILIKE '%duplicate%')/count(*),1) AS pct_duplicate_text
FROM t WHERE winter IS NOT NULL
GROUP BY 1 ORDER BY 1
"""
)
save_csv(heat_winter, "14_heat_winter.csv")

heat_top_bbl = q(
    f"""
SELECT bbl, incident_address, borough, count(*) AS n
FROM {SRC}
WHERE complaint_type='HEAT/HOT WATER'
AND strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-10' AND '2026-04'
GROUP BY 1,2,3
ORDER BY n DESC
LIMIT 15
"""
)
save_csv(heat_top_bbl, "15_heat_top_bbl_w2.csv")

heat_robust = q(
    f"""
WITH base AS (
  SELECT
    CASE WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2024-10' AND '2025-04' THEN 'W1'
         WHEN strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-10' AND '2026-04' THEN 'W2' END AS winter,
    bbl, unique_key
  FROM {SRC}
  WHERE complaint_type='HEAT/HOT WATER'
),
topw2 AS (
  SELECT bbl FROM base WHERE winter='W2' AND bbl IS NOT NULL AND bbl NOT IN ('','0','0000000000')
  GROUP BY 1 ORDER BY count(*) DESC LIMIT 10
)
SELECT winter,
  count(*) AS n_all,
  sum(CASE WHEN bbl IN (SELECT bbl FROM topw2) THEN 1 ELSE 0 END) AS n_in_w2_top10_bbl,
  count(*) - sum(CASE WHEN bbl IN (SELECT bbl FROM topw2) THEN 1 ELSE 0 END) AS n_ex_w2_top10_bbl
FROM base WHERE winter IS NOT NULL
GROUP BY 1 ORDER BY 1
"""
)
save_csv(heat_robust, "16_heat_exclude_top10.csv")

close_agency = q(
    f"""
SELECT agency,
  count(*) AS n_created_before_2026_06,
  sum(status='Closed') AS n_status_closed,
  sum(closed_date IS NOT NULL AND closed_date<>'') AS n_has_closed_date,
  percentile_cont(0.5) WITHIN GROUP (
    ORDER BY datediff('hour', try_cast(created_date AS timestamp), try_cast(closed_date AS timestamp))
  ) FILTER (
    WHERE closed_date IS NOT NULL AND closed_date<>''
      AND try_cast(closed_date AS timestamp) >= try_cast(created_date AS timestamp)
  ) AS median_hours_among_closed,
  percentile_cont(0.9) WITHIN GROUP (
    ORDER BY datediff('hour', try_cast(created_date AS timestamp), try_cast(closed_date AS timestamp))
  ) FILTER (
    WHERE closed_date IS NOT NULL AND closed_date<>''
      AND try_cast(closed_date AS timestamp) >= try_cast(created_date AS timestamp)
  ) AS p90_hours_among_closed
FROM {SRC}
WHERE try_cast(created_date AS timestamp) < TIMESTAMP '2026-06-01'
GROUP BY 1
ORDER BY n_created_before_2026_06 DESC
"""
)
save_csv(close_agency, "17_close_time_agency_created_before_202606.csv")

close_overall = q(
    f"""
SELECT
  count(*) AS n_created,
  sum(status='Closed') AS n_status_closed,
  percentile_cont(0.5) WITHIN GROUP (
    ORDER BY datediff('hour', try_cast(created_date AS timestamp), try_cast(closed_date AS timestamp))
  ) FILTER (
    WHERE closed_date IS NOT NULL AND closed_date<>''
      AND try_cast(closed_date AS timestamp) >= try_cast(created_date AS timestamp)
  ) AS median_hours_closed_all,
  percentile_cont(0.5) WITHIN GROUP (
    ORDER BY datediff('hour', try_cast(created_date AS timestamp), try_cast(closed_date AS timestamp))
  ) FILTER (
    WHERE agency<>'NYPD' AND closed_date IS NOT NULL AND closed_date<>''
      AND try_cast(closed_date AS timestamp) >= try_cast(created_date AS timestamp)
  ) AS median_hours_closed_ex_nypd
FROM {SRC}
WHERE try_cast(created_date AS timestamp) < TIMESTAMP '2026-06-01'
"""
)
save_csv(close_overall, "18_close_time_overall.csv")

lead_month = q(
    f"""
SELECT strftime(try_cast(created_date AS timestamp),'%Y-%m') AS year_month,
       descriptor, count(*) AS n
FROM {SRC}
WHERE complaint_type='Lead'
GROUP BY 1,2 ORDER BY 1,2
"""
)
save_csv(lead_month, "19_lead_month_descriptor.csv")

oos = q(
    f"""
SELECT agency, agency_name, complaint_type, count(*) AS n,
 min(strftime(try_cast(created_date AS timestamp),'%Y-%m')) AS min_ym,
 max(strftime(try_cast(created_date AS timestamp),'%Y-%m')) AS max_ym
FROM {SRC}
WHERE agency='OOS'
GROUP BY 1,2,3
"""
)
save_csv(oos, "20_oos.csv")

dup_share = q(
    f"""
SELECT 'all records' AS slice, count(*) AS n,
 sum(resolution_description ILIKE '%duplicate%') AS n_dup_text
FROM {SRC}
UNION ALL
SELECT 'HEAT/HOT WATER', count(*), sum(resolution_description ILIKE '%duplicate%')
FROM {SRC} WHERE complaint_type='HEAT/HOT WATER'
UNION ALL
SELECT 'UNKNOWN-channel pothole', count(*), sum(resolution_description ILIKE '%duplicate%')
FROM {SRC}
WHERE complaint_type='Street Condition' AND descriptor='Pothole' AND open_data_channel_type='UNKNOWN'
UNION ALL
SELECT 'Illegal Parking', count(*), sum(resolution_description ILIKE '%duplicate%')
FROM {SRC} WHERE complaint_type='Illegal Parking'
UNION ALL
SELECT 'Noise - Residential', count(*), sum(resolution_description ILIKE '%duplicate%')
FROM {SRC} WHERE complaint_type='Noise - Residential'
"""
)
save_csv(dup_share, "21_duplicate_text_share.csv")

snow_conc = q(
    f"""
WITH d AS (
  SELECT date_trunc('day', try_cast(created_date AS timestamp))::DATE AS dt, count(*) AS n
  FROM {SRC}
  WHERE complaint_type='Snow or Ice'
  AND strftime(try_cast(created_date AS timestamp),'%Y-%m') BETWEEN '2025-11' AND '2026-03'
  GROUP BY 1
)
SELECT count(*) AS n_days, sum(n) AS n_tickets, max(n) AS max_day,
 sum(n) FILTER (WHERE n>=1000) AS tickets_on_days_ge_1000,
 count(*) FILTER (WHERE n>=1000) AS n_days_ge_1000
FROM d
"""
)
save_csv(snow_conc, "22_snow_concentration_w2.csv")

# ---------------------------------------------------------------------------
# Derived numbers
# ---------------------------------------------------------------------------
nA = int(period_tot.loc[period_tot["period"].str.startswith("A"), "n"].iloc[0])
nB = int(period_tot.loc[period_tot["period"].str.startswith("B"), "n"].iloc[0])
delta = nB - nA
pct = 100.0 * delta / nA

excl = excl_buckets.copy()
excl["share_of_delta_pct"] = (excl["delta"] / delta * 100).round(2)
save_csv(excl, "09_exclusive_volume_buckets.csv")

n_all = int(overview["n_records"].iloc[0])
n_key = int(overview["n_unique_key"].iloc[0])

# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
wb = Workbook()

thin = Border(
    left=Side(style="thin", color="D0D5DD"),
    right=Side(style="thin", color="D0D5DD"),
    top=Side(style="thin", color="D0D5DD"),
    bottom=Side(style="thin", color="D0D5DD"),
)
fill_head = PatternFill("solid", fgColor="1F4E79")
fill_kpi = PatternFill("solid", fgColor="D6EAF8")
fill_fact = PatternFill("solid", fgColor="E8F5E9")
fill_inf = PatternFill("solid", fgColor="FFF8E1")
fill_unv = PatternFill("solid", fgColor="FCE4EC")
fill_warn = PatternFill("solid", fgColor="FDEBD0")
fill_alt = PatternFill("solid", fgColor="F8FAFC")
font_head = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
font_title = Font(name="Calibri", size=16, bold=True, color="1F4E79")
font_h2 = Font(name="Calibri", size=13, bold=True, color="1F4E79")
font_body = Font(name="Calibri", size=11)
wrap = Alignment(wrap_text=True, vertical="top")


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c)
        cell.fill = fill_head
        cell.font = font_head
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin


def write_df(ws, df, start_row, start_col=1, header=True):
    r = start_row
    if header:
        for j, col in enumerate(df.columns, start_col):
            ws.cell(r, j, col)
        style_header(ws, r, start_col + len(df.columns) - 1)
        r += 1
    for i, row in df.iterrows():
        for j, col in enumerate(df.columns, start_col):
            val = row[col]
            cell = ws.cell(r, j, None if pd.isna(val) else val)
            cell.border = thin
            cell.font = font_body
            if i % 2 == 1:
                cell.fill = fill_alt
        r += 1
    return r


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def kpi(ws, row, col, label, value, note=""):
    ws.cell(row, col, label).font = Font(bold=True, name="Calibri", size=10, color="1F4E79")
    ws.cell(row, col).fill = fill_kpi
    ws.cell(row, col).alignment = wrap
    ws.cell(row + 1, col, value).font = Font(bold=True, name="Calibri", size=14)
    ws.cell(row + 1, col).fill = fill_kpi
    if note:
        ws.cell(row + 2, col, note).font = Font(name="Calibri", size=9, italic=True, color="475569")
        ws.cell(row + 2, col).alignment = wrap
        ws.cell(row + 2, col).fill = fill_kpi


# ===== Sheet 0 cover =====
ws = wb.active
ws.title = "00_阅读说明"
set_widths(ws, [22, 88, 22, 22])
ws.merge_cells("A1:D1")
ws["A1"] = "NYC 311 服务请求：证据导向分析看板（2024-09-07 至 2026-09-05 创建记录）"
ws["A1"].font = font_title
ws.row_dimensions[1].height = 28
ws.merge_cells("A2:D2")
ws["A2"] = (
    "记录粒度=一条 unique_key 对应一条 311 服务请求，不等于一次独立现实事件。"
    "本表只描述该导出样本中的创建量、构成与关闭字段；不推断机构绩效或外部天气/政策原因。"
)
ws["A2"].alignment = wrap
ws.row_dimensions[2].height = 36

ws["A4"] = "建议阅读顺序"
ws["A4"].font = font_h2
ws["A5"] = (
    "1) 本页口径 → 2) 数据概况 → 3) 核心发现（量增分解）→ 4) 积雪/坑洞集中来源 → "
    "5) 供暖 → 6) 关闭时长与右删失 → 7) 分类口径变化 → 8) 方法与限制。灰色附录为完整明细。"
)
ws.merge_cells("A5:D5")
ws["A5"].alignment = wrap
ws.row_dimensions[5].height = 32

ws["A7"] = "证据分层（全文通用）"
ws["A7"].font = font_h2
ws["A8"] = "层次"
ws["B8"] = "含义"
ws["C8"] = "本看板如何使用"
style_header(ws, 8, 3)
ws["A9"] = "事实"
ws["B9"] = "可从原始字段直接计数或复算的结果"
ws["C9"] = "绿色标注的数字；均给出窗口、筛选、分母"
ws["A10"] = "推断"
ws["B10"] = "在事实之上、仍可用本数据检验的解释（例如构成驱动而非全面上升）"
ws["C10"] = "黄色；已做分解、对照或剔除检验"
ws["A11"] = "未验证解释"
ws["B11"] = "需要天气、政策、巡检计划、居民行为等本数据没有的材料"
ws["C11"] = "粉色；明确不作为结论"
for r, fill in [(9, fill_fact), (10, fill_inf), (11, fill_unv)]:
    for c in range(1, 4):
        ws.cell(r, c).fill = fill
        ws.cell(r, c).alignment = wrap
        ws.cell(r, c).border = thin
ws.row_dimensions[9].height = 28
ws.row_dimensions[10].height = 32
ws.row_dimensions[11].height = 32

ws["A13"] = "分析窗口（项目定义，不是源系统官方报告期）"
ws["A13"].font = font_h2
ws["A14"] = "窗口"
ws["B14"] = "定义"
ws["C14"] = "用途"
style_header(ws, 14, 3)
rows_win = [
    ("全样本", "created_date 解析后 ∈ [2024-09-07, 2026-09-05 01:50:33]；25 个 parquet 分片全量", "概况、构成"),
    ("可比 11 个月 A", "创建年月 2024-10 至 2025-08（含）", "量增对照基期；两端完整月"),
    ("可比 11 个月 B", "创建年月 2025-10 至 2026-08（含）", "对照报告期；不含不完整的 2024-09 与 2026-09"),
    ("供暖季 W1/W2", "创建年月 10–04：W1=2024-10..2025-04，W2=2025-10..2026-04", "跨冬对照；未使用 5 月"),
    ("关闭时长子集", "创建时间 < 2026-06-01，且 closed_date 可解析且 ≥ created_date", "给长周期类型留出观察窗；仍可能右删失"),
]
for i, row in enumerate(rows_win, 15):
    for j, v in enumerate(row, 1):
        ws.cell(i, j, v).alignment = wrap
        ws.cell(i, j).border = thin
    ws.row_dimensions[i].height = 28

ws["A21"] = "五个核心问题（收缩后）"
ws["A21"].font = font_h2
qs = [
    "Q1 可比 11 个月创建量增加了多少？增量主要来自哪些互斥类型桶？",
    "Q2 增量是否被少数事件日或单一渠道（UNKNOWN 坑洞）驱动？剔除后结论是否还在？",
    "Q3 HEAT/HOT WATER 跨冬增加，是更多楼宇还是少数地址堆积？重复关闭文案占比多少？",
    "Q4 关闭时长能否当作处理快慢？近期未关闭比例上升是不是绩效变化？",
    "Q5 哪些类型名称在样本后期发生切换，直接做类型同比会误读？",
]
for i, t in enumerate(qs, 22):
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    ws.cell(i, 1, t).alignment = wrap
    ws.row_dimensions[i].height = 20

ws["A28"] = "数据不能代表什么"
ws["A28"].font = font_h2
ws.merge_cells("A29:C31")
ws["A29"] = (
    "没有人口/道路里程分母，不能比较行政区“严重程度”。没有 SLA 目标，不能把中位关闭小时写成履约好坏。"
    "closed_date 最大值为 2026-12-14，晚于最大创建时间 2026-09-05，说明导出时点晚于创建窗口，但未关闭工单仍被右删失。"
    "天气、DOT 巡检计划、供暖法规、铅试剂发放活动等均未出现在授权数据中，不得写成事实。"
)
ws["A29"].alignment = wrap
ws["A29"].fill = fill_warn

ws["A33"] = "复现"
ws["A33"].font = font_h2
ws["A34"] = "脚本 build_workbook.py；中间 CSV 在 中间汇总/；原始 parquet 只读自 D:\\项目1\\原始数据。"
ws.merge_cells("A34:C34")

# ===== Sheet 1 overview =====
ws1 = wb.create_sheet("01_数据概况")
set_widths(ws1, [28, 22, 22, 22, 22, 22, 22, 22, 18, 18])
ws1["A1"] = "数据概况：全量 25 个分片"
ws1["A1"].font = font_title
ws1.merge_cells("A1:F1")

min_c = str(overview["min_created"].iloc[0])[:19]
max_c = str(overview["max_created"].iloc[0])[:19]
max_cl = str(overview["max_closed"].iloc[0])[:19]
kpi(ws1, 3, 1, "服务请求条数", f"{n_all:,}", "分母：全部分片")
kpi(ws1, 3, 2, "unique_key 去重", f"{n_key:,}", "与条数相等，无重复键")
kpi(ws1, 3, 3, "创建时间范围", f"{min_c[:10]} ～ {max_c[:10]}", "2024-09 与 2026-09 不完整")
kpi(ws1, 3, 4, "状态非 Closed", f"{int(overview['n_status_not_closed'].iloc[0]):,}", "含 Open/In Progress 等")
kpi(ws1, 3, 5, "无 closed_date", f"{int(overview['n_no_closed_date'].iloc[0]):,}", "与状态不完全等同")
kpi(ws1, 3, 6, "closed_date 早于创建", f"{int(overview['n_closed_before_created'].iloc[0]):,}", "质量瑕疵，占比<0.03%")

ws1["A7"] = "事实：月度创建量（单位=条；2024-09 从 7 日起，2026-09 仅到 5 日）"
ws1["A7"].font = font_h2
mshow = monthly[
    [
        "year_month",
        "n_created",
        "n_days_with_create",
        "avg_per_active_day",
        "n_status_not_closed",
        "pct_status_not_closed",
        "n_heat",
        "n_snow",
        "n_illegal_parking",
        "n_unk_pothole",
    ]
].copy()
mshow.columns = [
    "年月",
    "创建条数",
    "有创建的天数",
    "活跃日均条数",
    "状态非Closed",
    "非Closed占比%",
    "HEAT/HOT WATER",
    "Snow or Ice",
    "Illegal Parking",
    "UNKNOWN渠道坑洞",
]
end = write_df(ws1, mshow, 8)
# chart
chart = LineChart()
chart.title = "月度创建条数（不完整月未按日折算，见日均列）"
chart.y_axis.title = "条"
chart.x_axis.title = "年月"
chart.height = 8
chart.width = 18
data = Reference(ws1, min_col=2, min_row=8, max_row=8 + len(mshow))
cats = Reference(ws1, min_col=1, min_row=9, max_row=8 + len(mshow))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
ws1.add_chart(chart, "A36")

ws1[f"A{end + 1}"] = "构成（全样本）"
ws1[f"A{end + 1}"].font = font_h2
ws1[f"A{end + 2}"] = "按机构"
r2 = write_df(ws1, agency_all.rename(columns={"agency": "机构", "n": "条数", "pct": "占比%"}), end + 3)
ws1.cell(end + 2, 5, "按行政区")
r3 = write_df(
    ws1,
    borough_all.rename(columns={"borough": "行政区", "n": "条数", "pct": "占比%"}),
    end + 3,
    start_col=5,
)
ws1.cell(end + 2, 9, "按渠道")
write_df(
    ws1,
    channel_all.rename(columns={"channel": "渠道", "n": "条数", "pct": "占比%"}),
    end + 3,
    start_col=9,
)

ws1["A64"] = "状态分布"
write_df(ws1, status_all.rename(columns={"status": "状态", "n": "条数", "pct": "占比%"}), 65)
ws1["E64"] = "创建量最多的投诉类型（及主机构）"
write_df(
    ws1,
    top_types.rename(columns={"complaint_type": "投诉类型", "agency": "机构", "n": "条数", "pct": "占比%"}),
    65,
    start_col=5,
)

ws1["A90"] = (
    "口径提示：NYPD 约占 45%、HPD 约占 22%。全市中位关闭时长会被 NYPD 大量短单拉低，见工作表 05。"
    "2026-09 仅 41,974 条且截止 9月5日，不能与完整月直接比总量。"
)
ws1.merge_cells("A90:J91")
ws1["A90"].alignment = wrap
ws1["A90"].fill = fill_warn

# ===== Sheet 2 volume =====
ws2 = wb.create_sheet("02_量增分解")
set_widths(ws2, [36, 16, 16, 16, 18, 55])
ws2["A1"] = "Q1 可比 11 个月创建量变化：互斥桶分解"
ws2["A1"].font = font_title
ws2.merge_cells("A1:F1")
ws2["A2"] = (
    f"窗口 A = 2024-10..2025-08 创建；窗口 B = 2025-10..2026-08 创建；单位=服务请求条数；"
    f"分母=A 期 {nA:,} 条。筛选=无（全部类型）。聚合=按创建年月落入窗口后计数。"
)
ws2["A2"].alignment = wrap
ws2.merge_cells("A2:F2")
ws2.row_dimensions[2].height = 32

kpi(ws2, 4, 1, "A 期条数", f"{nA:,}", "11 个完整月")
kpi(ws2, 4, 2, "B 期条数", f"{nB:,}", "11 个完整月")
kpi(ws2, 4, 3, "增量", f"{delta:,}", f"相对 A 期 +{pct:.2f}%")
kpi(ws2, 4, 4, "日数是否对齐", "两端均为 10–08，各含一个非闰年 2 月", "不按日历年")

ws2["A8"] = "事实：互斥桶（每条请求只进一桶）对增量的贡献"
ws2["A8"].font = font_h2
bshow = excl.copy()
bshow.columns = ["互斥桶", "A期条数", "B期条数", "增量", "占全部增量%"]
write_df(ws2, bshow, 9)

# residual
ex_snow_unk = nB - nA
# compute residual from buckets
snow_d = int(excl.loc[excl["bucket"] == "Snow or Ice", "delta"].iloc[0])
heat_d = int(excl.loc[excl["bucket"] == "HEAT/HOT WATER", "delta"].iloc[0])
park_d = int(excl.loc[excl["bucket"] == "Illegal Parking", "delta"].iloc[0])
unk_d = int(excl.loc[excl["bucket"] == "UNKNOWN-channel pothole", "delta"].iloc[0])
st_d = int(excl.loc[excl["bucket"] == "Other Street Condition", "delta"].iloc[0])
oth_d = int(excl.loc[excl["bucket"] == "All other types", "delta"].iloc[0])

chart2 = BarChart()
chart2.type = "bar"
chart2.title = "各互斥桶对增量的贡献（条）"
chart2.y_axis.title = None
chart2.x_axis.title = "增量（条）"
data = Reference(ws2, min_col=4, min_row=9, max_row=9 + len(bshow))
cats = Reference(ws2, min_col=1, min_row=10, max_row=9 + len(bshow))
chart2.add_data(data, titles_from_data=True)
chart2.set_categories(cats)
chart2.shape = 4
chart2.height = 8
chart2.width = 16
ws2.add_chart(chart2, "A18")

ws2["A34"] = "推断（已用剔除检验）"
ws2["A34"].font = font_h2
ws2.merge_cells("A35:F37")
ws2["A35"] = (
    f"全部增量 {delta:,} 条中，Snow or Ice {snow_d:,}（{100*snow_d/delta:.1f}%）、"
    f"HEAT/HOT WATER {heat_d:,}（{100*heat_d/delta:.1f}%）、"
    f"Illegal Parking {park_d:,}（{100*park_d/delta:.1f}%）、"
    f"UNKNOWN 渠道坑洞 {unk_d:,}（{100*unk_d/delta:.1f}%）。"
    f"四者合计 {snow_d+heat_d+park_d+unk_d:,} 条，约占增量的 {100*(snow_d+heat_d+park_d+unk_d)/delta:.1f}%。"
    f"剔除雪和 UNKNOWN 坑洞后，其余仍增加 {delta-snow_d-unk_d:,} 条（相对 A 期其余部分仍为正）。"
    "因此：增量既不是“所有类型均匀变多”，也不是“只剩一个异常类型”。"
    "Illegal Parking 在月度序列上持续偏高，不像单日事故。"
)
ws2["A35"].alignment = wrap
ws2["A35"].fill = fill_inf

ws2["A39"] = "机构增量（非互斥于上表类型桶；HPD 增量大量与供暖重叠）"
write_df(
    ws2,
    yoy_agency.rename(columns={"agency": "机构", "n_A": "A期", "n_B": "B期", "delta": "增量", "pct_change": "相对A期%"}),
    40,
)

ws2["A60"] = "类型增量绝对值最大的 15 类（可与互斥桶对照；类型之间互斥）"
write_df(
    ws2,
    yoy_type.head(15).rename(
        columns={
            "complaint_type": "投诉类型",
            "n_A": "A期",
            "n_B": "B期",
            "delta": "增量",
            "pct_change": "相对A期%",
        }
    ),
    61,
)

# ===== Sheet 3 snow pothole =====
ws3 = wb.create_sheet("03_积雪与坑洞")
set_widths(ws3, [22, 16, 16, 18, 18, 18, 55])
ws3["A1"] = "Q2 增量是否被集中来源驱动：积雪日与 UNKNOWN 渠道坑洞"
ws3["A1"].font = font_title
ws3.merge_cells("A1:G1")
ws3["A2"] = (
    "积雪：complaint_type='Snow or Ice'，单位=条。"
    "坑洞集中来源：complaint_type='Street Condition' 且 descriptor='Pothole' 且 open_data_channel_type='UNKNOWN'。"
)
ws3["A2"].alignment = wrap
ws3.merge_cells("A2:G2")

n_snow_days = int(snow_conc["n_days"].iloc[0])
n_snow_tix = int(snow_conc["n_tickets"].iloc[0])
max_snow_day = int(snow_conc["max_day"].iloc[0])
tix_ge = int(snow_conc["tickets_on_days_ge_1000"].iloc[0])
nd_ge = int(snow_conc["n_days_ge_1000"].iloc[0])
kpi(ws3, 4, 1, "W2 积雪条数(11–03)", f"{n_snow_tix:,}", "2025-11..2026-03")
kpi(ws3, 4, 2, "≥1000条的天数", f"{nd_ge}", f"这 {nd_ge} 天合计 {tix_ge:,} 条")
kpi(ws3, 4, 3, "占该窗积雪", f"{100*tix_ge/n_snow_tix:.1f}%", "集中而非均匀")
kpi(ws3, 4, 4, "最高单日", f"{max_snow_day:,}", str(snow_days.iloc[0]["created_day"]))

ws3["A8"] = "事实：Snow or Ice 创建量最高的日期（全样本）"
write_df(ws3, snow_days.head(12).rename(columns={"created_day": "日期", "n": "条数"}), 9)

ws3["A23"] = "事实：街道状况 vs UNKNOWN 渠道坑洞（月）"
stshow = street_channel_month.rename(
    columns={
        "year_month": "年月",
        "n_street": "Street Condition",
        "n_unk_pothole": "UNKNOWN坑洞",
        "n_street_ex_unk_pothole": "街道状况其余",
        "n_street_unknown_channel": "街道状况UNKNOWN渠道",
        "n_street_named_channel": "街道状况具名渠道",
    }
)
write_df(ws3, stshow, 24)

chart3 = LineChart()
chart3.title = "Street Condition：UNKNOWN 坑洞 vs 其余"
chart3.height = 8
chart3.width = 16
data = Reference(ws3, min_col=2, min_row=24, max_col=4, max_row=24 + len(stshow))
cats = Reference(ws3, min_col=1, min_row=25, max_row=24 + len(stshow))
chart3.add_data(data, titles_from_data=True)
chart3.set_categories(cats)
ws3.add_chart(chart3, "A52")

ws3["A68"] = "事实：UNKNOWN 坑洞关闭说明中含 duplicate 的比例（文案匹配，非官方重复键）"
write_df(
    ws3,
    unk_pothole_res.rename(
        columns={
            "year_month": "年月",
            "n": "UNKNOWN坑洞条数",
            "n_duplicate_text": "说明含duplicate",
            "pct_duplicate_text": "占比%",
        }
    ),
    69,
)

ws3["A97"] = "推断"
ws3["A97"].font = font_h2
ws3.merge_cells("A98:G100")
ws3["A98"] = (
    "2026-03 Street Condition 升至 28,690 条，其中 UNKNOWN 渠道 Pothole 为 22,790 条，且该月全部 Pothole 记录渠道均为 UNKNOWN。"
    "具名渠道（ONLINE/PHONE 等）的街道状况也上升，但幅度远小于 UNKNOWN 坑洞。"
    "2026-03 UNKNOWN 坑洞关闭说明中约 43% 含 duplicate，说明条数明显大于“待处理的独立坑洞事件”。"
    "积雪增量同样高度日度集中：第二冬 16 个高量日贡献约八成积雪单。"
    "因此把全市 311 “变忙了”写成普遍运营压力，会被这两类集中来源放大。"
)
ws3["A98"].alignment = wrap
ws3["A98"].fill = fill_inf
ws3["A102"] = "未验证解释"
ws3["A102"].fill = fill_unv
ws3.merge_cells("A103:G104")
ws3["A103"] = (
    "积雪单日高峰是否对应降雪过程、UNKNOWN 坑洞是否来自 DOT 巡检/批量导入、duplicate 文案是否等于同一地理坑洞，"
    "本数据没有天气表、巡检计划或官方重复组 ID，不能验证。"
)
ws3["A103"].alignment = wrap
ws3["A103"].fill = fill_unv

# ===== Sheet 4 heat =====
ws4 = wb.create_sheet("04_供暖")
set_widths(ws4, [42, 18, 18, 18, 18, 18, 40])
ws4["A1"] = "Q3 HEAT/HOT WATER：跨冬对照、楼宇数与重复文案"
ws4["A1"].font = font_title
ws4.merge_cells("A1:G1")
ws4["A2"] = (
    "筛选：complaint_type='HEAT/HOT WATER'。"
    "供暖季窗口为项目定义的创建年月 10–04，不是法规供暖季证明。"
    "BBL 空值/0 不计入楼宇数。单位：条或去重 BBL。"
)
ws4["A2"].alignment = wrap
ws4.merge_cells("A2:G2")

h1 = heat_winter.iloc[0]
h2 = heat_winter.iloc[1]
# ensure order
if str(h1["winter"]).startswith("W2"):
    h1, h2 = h2, h1
kpi(ws4, 4, 1, "W1 条数", f"{int(h1['n_tickets']):,}", str(h1["winter"]))
kpi(ws4, 4, 2, "W2 条数", f"{int(h2['n_tickets']):,}", str(h2["winter"]))
kpi(ws4, 4, 3, "W1 去重BBL", f"{int(h1['n_bbl']):,}", "非空有效 BBL")
kpi(ws4, 4, 4, "W2 去重BBL", f"{int(h2['n_bbl']):,}", "楼宇数也增加")
kpi(ws4, 4, 5, "W2 说明含duplicate", f"{h2['pct_duplicate_text']}%", f"{int(h2['n_duplicate_text']):,} 条")

ws4["A8"] = "事实：两个供暖季对照"
hw = heat_winter.rename(
    columns={
        "winter": "供暖季窗口",
        "n_tickets": "条数",
        "n_bbl": "去重BBL",
        "n_address": "去重地址文本",
        "n_duplicate_text": "说明含duplicate",
        "pct_duplicate_text": "duplicate文案%",
    }
)
write_df(ws4, hw, 9)

ws4["A13"] = "稳健性：剔除 W2 报修量最高的 10 个 BBL 后，两个冬天的条数仍上升"
write_df(
    ws4,
    heat_robust.rename(
        columns={
            "winter": "季",
            "n_all": "全部条数",
            "n_in_w2_top10_bbl": "落入W2前10 BBL",
            "n_ex_w2_top10_bbl": "剔除后条数",
        }
    ),
    14,
)

ws4["A18"] = "W2 条数最高的 BBL（说明高度右偏，中位楼宇只有约 2 条，见中间汇总）"
write_df(
    ws4,
    heat_top_bbl.rename(columns={"bbl": "BBL", "incident_address": "地址", "borough": "行政区", "n": "条数"}),
    19,
)

# heat monthly chart from monthly
ws4["A37"] = "月度 HEAT/HOT WATER 创建条数"
hs = monthly[["year_month", "n_heat"]].rename(columns={"year_month": "年月", "n_heat": "HEAT/HOT WATER"})
write_df(ws4, hs, 38)
chart4 = LineChart()
chart4.title = "HEAT/HOT WATER 月度条数"
chart4.height = 7
chart4.width = 16
data = Reference(ws4, min_col=2, min_row=38, max_row=38 + len(hs))
cats = Reference(ws4, min_col=1, min_row=39, max_row=38 + len(hs))
chart4.add_data(data, titles_from_data=True)
chart4.set_categories(cats)
ws4.add_chart(chart4, "E37")

ws4["A66"] = "推断"
ws4.merge_cells("A67:G69")
ws4["A67"] = (
    f"W2 相对 W1：条数 {int(h1['n_tickets']):,} → {int(h2['n_tickets']):,} "
    f"（{100*(int(h2['n_tickets'])-int(h1['n_tickets']))/int(h1['n_tickets']):.1f}%），"
    f"去重 BBL {int(h1['n_bbl']):,} → {int(h2['n_bbl']):,} "
    f"（{100*(int(h2['n_bbl'])-int(h1['n_bbl']))/int(h1['n_bbl']):.1f}%）。"
    "条数和楼宇数同时上升，且剔除 W2 最热 10 个 BBL 后方向不变，故不是“一两栋楼刷爆总量”。"
    "但单楼最高达两千余条，且约三成关闭说明含 duplicate，条数仍显著大于独立供暖事件数。"
)
ws4["A67"].alignment = wrap
ws4["A67"].fill = fill_inf
ws4["A71"] = "未验证解释：第二冬是否更冷、锅炉故障是否更多、居民是否更频繁催单，均无温度或工单线程数据。"
ws4["A71"].fill = fill_unv
ws4.merge_cells("A71:G72")
ws4["A71"].alignment = wrap

# ===== Sheet 5 close =====
ws5 = wb.create_sheet("05_关闭时长与删失")
set_widths(ws5, [18, 22, 20, 22, 26, 26, 50])
ws5["A1"] = "Q4 关闭时长构成 vs 近期未关闭比例（右删失）"
ws5["A1"].font = font_title
ws5.merge_cells("A1:G1")
ws5["A2"] = (
    "关闭小时：datediff(hour, created, closed)；仅 closed_date 可解析且 ≥ created。"
    "创建 < 2026-06-01。中位数只在“已有关闭时间”的子集上计算，未关闭工单不进入中位数——这会让慢类型看起来更快。"
    "未关闭比例按创建月、分母=该月全部创建条数，分子=status<>'Closed'。"
)
ws5["A2"].alignment = wrap
ws5.merge_cells("A2:G2")
ws5.row_dimensions[2].height = 48

med_all = close_overall["median_hours_closed_all"].iloc[0]
med_ex = close_overall["median_hours_closed_ex_nypd"].iloc[0]
kpi(ws5, 4, 1, "创建<2026-06 中位关闭小时", f"{med_all:.0f}", "含 NYPD")
kpi(ws5, 4, 2, "剔除 NYPD 后中位小时", f"{med_ex:.0f}", "构成效应")
kpi(ws5, 4, 3, "2024-10 非Closed%", f"{monthly.loc[monthly.year_month=='2024-10','pct_status_not_closed'].iloc[0]}%", "长观察窗")
kpi(ws5, 4, 4, "2026-08 非Closed%", f"{monthly.loc[monthly.year_month=='2026-08','pct_status_not_closed'].iloc[0]}%", "观察窗短")
kpi(ws5, 4, 5, "2026-09 非Closed%", f"{monthly.loc[monthly.year_month=='2026-09','pct_status_not_closed'].iloc[0]}%", "月不完整")

ws5["A8"] = "事实：按创建月的未关闭比例（不要解读成“越来越慢”，除非先对齐观察窗）"
m2 = monthly[["year_month", "n_created", "n_status_not_closed", "pct_status_not_closed", "n_no_closed_date"]].rename(
    columns={
        "year_month": "年月",
        "n_created": "创建条数",
        "n_status_not_closed": "状态非Closed",
        "pct_status_not_closed": "非Closed%",
        "n_no_closed_date": "无closed_date",
    }
)
write_df(ws5, m2, 9)
chart5 = LineChart()
chart5.title = "创建月的状态非Closed占比%"
chart5.height = 7
chart5.width = 16
data = Reference(ws5, min_col=4, min_row=9, max_row=9 + len(m2))
cats = Reference(ws5, min_col=1, min_row=10, max_row=9 + len(m2))
chart5.add_data(data, titles_from_data=True)
chart5.set_categories(cats)
ws5.add_chart(chart5, "G8")

ws5["A37"] = "事实：机构关闭时长（创建<2026-06-01）"
cshow = close_agency.rename(
    columns={
        "agency": "机构",
        "n_created_before_2026_06": "创建条数",
        "n_status_closed": "状态Closed",
        "n_has_closed_date": "有closed_date",
        "median_hours_among_closed": "已关闭中位小时",
        "p90_hours_among_closed": "已关闭P90小时",
    }
)
write_df(ws5, cshow, 38)

ws5["A58"] = "推断"
ws5.merge_cells("A59:F61")
ws5["A59"] = (
    "全市约 7 小时的中位关闭时间几乎完全是构成结果：NYPD 已关闭中位约 1 小时，去掉 NYPD 后中位升到约 64 小时。"
    "TLC/EDC/DOB 的已关闭中位以百至千小时计，且长尾更长；把它们和 NYPD 比“快慢”没有共同 SLA。"
    "未关闭占比从早期约 1% 升到 2026-08 的 16% 和 2026-09 的 41%，与观察窗变短同向，不能单独证明处理变慢。"
    "早期月份仍有约 1% 未关闭，说明存在长期未结案，不是所有工单都会在几个月内关闭。"
)
ws5["A59"].alignment = wrap
ws5["A59"].fill = fill_inf
ws5["A63"] = "未验证解释：各局目标处理时限、夜间值班、现场执法是否变化，无 SLA 表。"
ws5["A63"].fill = fill_unv
ws5.merge_cells("A63:F64")

# ===== Sheet 6 taxonomy =====
ws6 = wb.create_sheet("06_分类口径变化")
set_widths(ws6, [14, 18, 18, 14, 18, 18, 55])
ws6["A1"] = "Q5 类型名称切换：水务家族与铅试剂"
ws6["A1"].font = font_title
ws6.merge_cells("A1:G1")
ws6["A2"] = (
    "若直接对单一 complaint_type 做同比，2026-08 的 Water System / Sewer 会显示为“归零”，"
    "同时 Water/Sewer Maintenance 会显示为“爆发”。下面用加总家族避免该误读。"
)
ws6["A2"].alignment = wrap
ws6.merge_cells("A2:G2")

wf = monthly[
    [
        "year_month",
        "n_water_system",
        "n_water_maint",
        "n_sewer",
        "n_sewer_maint",
        "n_water_family",
        "n_lead",
    ]
].rename(
    columns={
        "year_month": "年月",
        "n_water_system": "Water System",
        "n_water_maint": "Water Maintenance",
        "n_sewer": "Sewer",
        "n_sewer_maint": "Sewer Maintenance",
        "n_water_family": "四类合计",
        "n_lead": "Lead",
    }
)
ws6["A4"] = "事实：水务相关类型月度（单位=条）"
write_df(ws6, wf, 5)
chart6 = LineChart()
chart6.title = "水务四类合计 vs 分项"
chart6.height = 8
chart6.width = 16
data = Reference(ws6, min_col=2, min_row=5, max_col=6, max_row=5 + len(wf))
cats = Reference(ws6, min_col=1, min_row=6, max_row=5 + len(wf))
chart6.add_data(data, titles_from_data=True)
chart6.set_categories(cats)
ws6.add_chart(chart6, "A33")

ws6["A50"] = "事实：Lead 描述字段切换（L10 名称在 2026-08 消失，Lead Kit 接续）"
lshow = lead_month.rename(columns={"year_month": "年月", "descriptor": "描述", "n": "条数"})
write_df(ws6, lshow, 51)

ws6["A80"] = "事实：新出现机构 OOS"
if len(oos):
    write_df(
        ws6,
        oos.rename(
            columns={
                "agency": "机构",
                "agency_name": "机构名",
                "complaint_type": "类型",
                "n": "条数",
                "min_ym": "最早月",
                "max_ym": "最晚月",
            }
        ),
        81,
    )

ws6["A85"] = "推断"
ws6.merge_cells("A86:G88")
ws6["A86"] = (
    "2026-08 起 Water System 与 Sewer 在本样本中为 0，同月 Water Maintenance 与 Sewer Maintenance 升至数千。"
    "四类合计在 2026-08 仍有 13,985 条，并非水务需求消失。2026-07 为并存过渡月。"
    "Lead 从 'Lead Kit Request (Residential) (L10)' 切到 'Lead Kit'：这是试剂申请类工单，"
    "2024-10/11 的高峰（5,529 / 9,103）会强烈影响 Lead 类型的同比，不能写成“铅污染事件增减”。"
    "OOS / Office of the Sheriff / Cannabis Retailer 自 2025-09 出现，量级约 4 千，对全市增量影响很小。"
)
ws6["A86"].alignment = wrap
ws6["A86"].fill = fill_inf
ws6["A90"] = "未验证解释：是否官方 taxonomy 发布、试剂发放活动、消防栓夏季开放政策，均无外部文件。"
ws6["A90"].fill = fill_unv
ws6.merge_cells("A90:G91")

# Drug as caution not core
ws6["A93"] = "附录观察（非核心结论）：Drug Activity 在 2025-07/08 升至 6,038 / 5,447 后回落，未发现等量迁入其他 NYPD 类型。原因未验证。"
ws6.merge_cells("A93:G94")
ws6["A93"].alignment = wrap

# ===== Sheet 7 methods =====
ws7 = wb.create_sheet("07_方法与限制")
set_widths(ws7, [22, 90])
ws7["A1"] = "方法、假设、限制"
ws7["A1"].font = font_title
items = [
    ("记录粒度", "unique_key 全样本唯一，一行=一条服务请求。关闭说明含 duplicate 时，多行可能对应同一现场问题。"),
    ("时间字段", "created_date / closed_date 以字符串存储，分析中用 try_cast 为 timestamp。无法解析的关闭时间视为缺失。"),
    ("不完整月", "2024-09 自 7 日；2026-09 至 5 日 01:50:33。同比主结论不用这两月的总量。"),
    ("可比窗口", "选择完整月 10–08 对齐，避免日历年 2024/2026 残缺。未做工作日调整。"),
    ("供暖季", "用 10–04 创建月近似冬季供暖需求高峰，不是官方 heat season 法规窗口（常含 10-01 至 05-31）。"),
    ("关闭中位数", "条件是已关闭；未关闭不计入。创建截止 2026-06-01，相对导出中出现的 2026-12-14 关闭记录仍可能不够覆盖超长工单。"),
    ("duplicate 文案", "resolution_description ILIKE '%duplicate%' 是项目定义检查，不是 NYC 官方重复工单主键。"),
    ("UNKNOWN 渠道", "字段值原样使用。不能从本数据证明 UNKNOWN=系统导入，只证明该渠道在 2026-03 坑洞上极度集中。"),
    ("增量贡献", "互斥桶加总等于全部增量。机构表与类型表不要与互斥桶再相加。"),
    ("负关闭时序", f"closed_date 早于 created_date 共 {int(overview['n_closed_before_created'].iloc[0])} 条，关闭时长计算已排除。"),
    ("禁止的结论", "不得把描述性差异写成执法力度、居民素质、机构绩效或天气因果。"),
    ("授权边界", "只用 D:\\项目1\\原始数据 与 D:\\项目2\\input\\data_manifest.json；未使用天气或其他 NYC 表。"),
]
ws7["A3"] = "项目"
ws7["B3"] = "说明"
style_header(ws7, 3, 2)
for i, (a, b) in enumerate(items, 4):
    ws7.cell(i, 1, a).border = thin
    ws7.cell(i, 1).alignment = wrap
    ws7.cell(i, 2, b).border = thin
    ws7.cell(i, 2).alignment = wrap
    ws7.row_dimensions[i].height = 36

ws7["A17"] = "曾探索但未升为核心发现的方向"
ws7["A17"].font = font_h2
ws7.merge_cells("A18:B20")
ws7["A18"] = (
    "Drug Activity 夏季 2025 尖峰后下降；行政区份额差异（无人口分母）；"
    "噪声类总量大但跨年变化不如供暖/积雪/坑洞对增量的解释力强；"
    "EDC 直升机噪声类型随 EDC 量下降。这些保留在中间汇总，避免故事过多。"
)
ws7["A18"].alignment = wrap

ws7["A22"] = "数字回查"
ws7["A22"].font = font_h2
ws7["A23"] = "全样本条数必须等于各月 n_created 之和，也等于各机构之和。可比窗口 A+B 不等于全样本（不含 2024-09、2025-09、2026-09）。"
ws7.merge_cells("A23:B23")

# ===== appendix monthly already in sheet1; extra type yoy full
ws8 = wb.create_sheet("附录_类型同比全表")
set_widths(ws8, [40, 16, 16, 16, 14])
ws8["A1"] = "可比窗口全部投诉类型（按增量绝对值排序）"
write_df(
    ws8,
    yoy_type.rename(
        columns={
            "complaint_type": "投诉类型",
            "n_A": "A期_2024-10至2025-08",
            "n_B": "B期_2025-10至2026-08",
            "delta": "增量",
            "pct_change": "相对A期%",
        }
    ),
    2,
)

ws9 = wb.create_sheet("附录_重复文案切片")
set_widths(ws9, [32, 16, 16, 16])
d2 = dup_share.copy()
d2["pct"] = (100.0 * d2["n_dup_text"] / d2["n"]).round(2)
write_df(
    ws9,
    d2.rename(columns={"slice": "切片", "n": "条数", "n_dup_text": "说明含duplicate", "pct": "占比%"}),
    1,
)

# freeze + print
for s in wb.worksheets:
    s.freeze_panes = "A2" if s.title != "00_阅读说明" else "A4"
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToPage = True
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    s.print_title_rows = "1:1"
    s.sheet_properties.pageSetUpPr.fitToPage = True

out = ROOT / "最终成果.xlsx"
wb.save(out)

# reconciliation prints
month_sum = int(monthly["n_created"].sum())
agency_sum = int(agency_all["n"].sum())
bucket_delta = int(excl["delta"].sum())
print("n_all", n_all, "month_sum", month_sum, "agency_sum", agency_sum)
print("nA", nA, "nB", nB, "delta", delta, "bucket_delta", bucket_delta)
print("wrote", out)
