# -*- coding: utf-8 -*-
"""从 中间汇总/ CSV 构建 最终成果.xlsx 看板（含原生 Excel 图表）。"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter

O = r"D:\项目2\GLM\中间汇总"
OUTX = r"D:\项目2\GLM\最终成果.xlsx"

# ---------- 样式 ----------
H_FILL = PatternFill('solid', fgColor='1F4E79')      # 深蓝表头
H_FONT = Font(bold=True, color='FFFFFF', size=11)
T_FONT = Font(bold=True, size=14, color='1F4E79')
S_FONT = Font(bold=True, size=12, color='1F4E79')
THIN = Border(bottom=Side(style='thin', color='BFBFBF'))
ALT = PatternFill('solid', fgColor='F2F7FB')

def style_header(ws, row, ncols, start=1):
    for j in range(start, start + ncols):
        c = ws.cell(row=row, column=j)
        c.fill = H_FILL; c.font = H_FONT
        c.alignment = Alignment(horizontal='center', vertical='center')

def write_table(ws, df, start_row, start_col=1, num_fmt=None, pct_cols=()):
    for j, col in enumerate(df.columns):
        ws.cell(row=start_row, column=start_col + j, value=str(col))
    style_header(ws, start_row, len(df.columns), start_col)
    for i, (_, r) in enumerate(df.iterrows()):
        for j, col in enumerate(df.columns):
            v = r[col]
            c = ws.cell(row=start_row + 1 + i, column=start_col + j)
            if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
                c.value = float(v) if isinstance(v, (np.floating, float)) else int(v)
                if col in pct_cols: c.number_format = '0.0"%"'
                elif num_fmt: c.number_format = num_fmt
            else:
                c.value = str(v)
            if i % 2 == 1: c.fill = ALT
            c.border = THIN
    return start_row + len(df)

def set_widths(ws, widths):
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

wb = Workbook()

# ============ 数据准备 ============
mc = pd.read_csv(O + "/月度x投诉类型.csv")
mp = mc.pivot(index='月份', columns='投诉类型', values='记录数').fillna(0)
months = list(mp.index)
mp['总计'] = mp.sum(axis=1)

complaint = pd.read_csv(O + "/投诉类型总量.csv")
complaint.columns = ['投诉类型', '记录数']
complaint = complaint.sort_values('记录数', ascending=False).reset_index(drop=True)
complaint['占比%'] = complaint['记录数'] / complaint['记录数'].sum() * 100

agency = pd.read_csv(O + "/机构总量.csv")
agency.columns = ['机构', '记录数']
agency = agency.sort_values('记录数', ascending=False).reset_index(drop=True)
agency['占比%'] = agency['记录数'] / agency['记录数'].sum() * 100
agency_dur = pd.read_csv(O + "/机构处理时长.csv")
agency = agency.merge(agency_dur[['机构', '中位小时', 'P75小时', 'P90小时']], on='机构', how='left')
nc_ag = pd.read_csv(O + "/未关闭_按机构.csv")
agency = agency.merge(nc_ag, on='机构', how='left').fillna({'未关闭数': 0})

borough = pd.read_csv(O + "/区域总量.csv")
borough.columns = ['borough', '记录数']
borough = borough.sort_values('记录数', ascending=False).reset_index(drop=True)
borough['占比%'] = borough['记录数'] / borough['记录数'].sum() * 100

status = pd.read_csv(O + "/状态总量.csv")
status.columns = ['状态', '记录数']
status['占比%'] = status['记录数'] / status['记录数'].sum() * 100

channel = pd.read_csv(O + "/渠道总量.csv")
channel.columns = ['渠道', '记录数']
channel['占比%'] = channel['记录数'] / channel['记录数'].sum() * 100

cd = pd.read_csv(O + "/投诉类型处理时长.csv")

# 季节性系列（月度）
seasonal_cols = ['HEAT/HOT WATER', 'Noise - Residential', 'Noise - Street/Sidewalk', 'Illegal Parking', 'Street Condition']

# 区域构成
bc = pd.read_csv(O + "/区域x投诉类型.csv")
bp = bc.pivot(index='区域', columns='投诉类型', values='记录数').fillna(0)
mix_cols = ['Illegal Parking', 'Noise - Residential', 'HEAT/HOT WATER', 'Blocked Driveway',
            'Noise - Street/Sidewalk', 'Street Condition', 'Encampment']
bshare = (bp[mix_cols].div(bp.sum(axis=1), axis=0) * 100).round(1)
bshare = bshare.loc[['BROOKLYN', 'QUEENS', 'BRONX', 'MANHATTAN', 'STATEN ISLAND']]

TOTAL = 7525498

# ============ Sheet 1: 总览 ============
ws = wb.active; ws.title = '1-总览'
ws['A1'] = 'NYC 311 服务请求分析看板（2024-09-07 ～ 2026-09-07）'; ws['A1'].font = T_FONT
ws['A2'] = '数据源：D:\\项目1\\原始数据 25 个 Parquet 分片，全量处理；单位：条（服务请求记录，非独立现实事件）'
ws['A3'] = '本看板无需运行代码即可阅读；核心结论见「2-核心发现」，支撑明细见后续各表。'
for r in (2, 3): ws.cell(row=r, column=1).font = Font(size=10, color='595959')

ov = pd.DataFrame({
    '指标': ['总记录数', '时间范围（created_date）', '月份数', '机构（agency）数', '投诉类型数',
             '最大机构', '最大机构占比', '最大投诉类型', '最大投诉类型占比',
             '已关闭(Closed)记录', '已关闭占比', '快照时未关闭记录', '未关闭占比',
             '记录数最多区域', '该区域占比', '主要受理渠道', '记录数最多月份'],
    '值': [f'{TOTAL:,}', '2024-09-07 ～ 2026-09-05', '25（其中 2024-09 与 2026-09 为不完整月）', f"{agency['机构'].nunique()}",
           f"{complaint['投诉类型'].nunique()}", 'NYPD（警察局）', '45.2%', 'Illegal Parking（违章停车）', '15.4%',
           '7,275,349', '96.7%', '250,149', '3.3%', 'BROOKLYN（布鲁克林）', '30.0%', 'ONLINE（线上）/ PHONE（电话）/ MOBILE（App）', '2026-01（348,511 条）']
})
write_table(ws, ov, 5)
ws['A24'] = '数据质量摘要'; ws['A24'].font = S_FONT
dq = pd.DataFrame({
    '检查项': ['unique_key 重复', 'created_date 缺失', 'closed_date 缺失（与未结案对应）',
               'closed_date 早于 created_date（已从时效计算中剔除）', '投诉类型口径变更',
               '不完整月份'],
    '结果': ['0 条', '0 条', '240,304 条（3.2%）', '1,877 条（0.025%）',
             '2026-07 起 DEP 的 Water System 逐步停用，2026-08 完全拆分为 Water Maintenance / Sewer Maintenance',
             '2024-09（仅 9/7 起）与 2026-09（至 9/7），月度对比仅用完整月']
})
write_table(ws, dq, 25)
set_widths(ws, [52, 62])

# ============ Sheet 2: 核心发现 ============
ws = wb.create_sheet('2-核心发现')
ws['A1'] = '核心发现（全部数字可回查「3~6」各支撑表）'; ws['A1'].font = T_FONT
findings = [
    ('发现1｜总量上升：完整月同比 +10.5%',
     '对比两个可比的 11 个月完整区间：2024-10～2025-08 共 3,293,003 条，2025-10～2026-08 共 3,639,698 条，+10.5%。\n'
     '增长主要来自：Street Condition（+75.1%）、HEAT/HOT WATER（+21.9%）、Illegal Parking（+10.9%）；\n'
     '噪声类（Noise-Residential、Noise-Street/Sidewalk）同期基本持平（约 0%）。'),
    ('发现2｜强季节性：供暖投诉冬夏相差约 20 倍',
     'HEAT/HOT WATER 冬季爆发、夏季几乎消失：2026-01 达 79,928 条，而 2025-08 仅 2,929 条（约 27 倍）。\n'
     'Noise-Residential 在冬季深夜供暖锅炉季与夏季均偏高（2025-01：70,837 为全期峰值）；\n'
     'Noise-Street/Sidewalk 明显夏季高（2025-08：25,205 vs 2025-02：3,881）。\n'
     '总量随之冬季高、春季低：1 月（约 34.9 万）为全年峰值，2-4 月为低谷。'),
    ('发现3｜2026 年春 Street Condition（路面破损）异常激增',
     '2026-03 Street Condition 共 28,690 条，约为 2025-03（7,035 条）的 4.1 倍；其中 DOT 受理 28,629 条，\n'
     'Pothole（坑洼）22,790 条、Cave-in（塌陷）2,419 条；激增遍布全月每一天和全部五个区，4-5 月仍偏高（15,850 / 11,777）。\n'
     '这是描述性事实；原因（如恶劣冬季后的路况恶化、或录入口径变化）无法仅凭本数据验证。'),
    ('发现4｜处理时效按机构分层悬殊，且"关闭"含义不同',
     'NYPD 中位 1.3 小时——其结案描述多为"到场后未发现违法"（快速处置/结案，不等于问题修复）；\n'
     'HPD 中位 87.5 小时、DOB 158.9 小时、DPR 165.3 小时；DOHMH 呈现 60 天（约 1440 小时）自动结案特征。\n'
     '全市月度中位时效也随季节波动：夏季约 4-6 小时，冬季 15-26 小时（投诉构成变化所致）。\n'
     '快照时点（2026-09-07）仍有 250,149 条（3.3%）未关闭，集中在 HPD（59,718）、DPR（57,639）、DOT（28,715）。'),
    ('发现5｜区域画像差异明显',
     '总量：Brooklyn 30.0% > Queens 24.2% > Bronx 22.1% > Manhattan 19.8% > Staten Island 3.8%（绝对记录数，未经人口归一）。\n'
     '构成差异：Bronx 以噪声-住宅 21.9% + HEAT/HOT WATER 13.7%（五区最高）呈住房质量压力型；\n'
     'Manhattan 的 Encampment（流浪营地）占 4.6% 为五区最高、Blocked Driveway 仅 0.6% 为最低；\n'
     'Queens 的 Street Condition 占 4.1% 偏高。'),
    ('发现6｜受理渠道：线上为主',
     'ONLINE 44.3%、PHONE 25.7%、MOBILE 22.1%、UNKNOWN 7.9%。线上+App 合计约 2/3。')
]
r = 3
for title, body in findings:
    ws.cell(row=r, column=1, value=title).font = S_FONT
    ws.cell(row=r, column=1).fill = PatternFill('solid', fgColor='DDEBF7')
    c = ws.cell(row=r + 1, column=1, value=body)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    c.font = Font(size=11)
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=8)
    ws.row_dimensions[r + 1].height = 78
    r += 3
set_widths(ws, [100])

# ============ Sheet 3: 月度趋势 ============
ws = wb.create_sheet('3-月度趋势')
ws['A1'] = '月度趋势（条数；2024-09 与 2026-09 为不完整月）'; ws['A1'].font = T_FONT
trend = pd.DataFrame({'月份': months})
for col in ['总计'] + seasonal_cols:
    trend[col] = [mp.loc[m, col] for m in months]
end_row = write_table(ws, trend, 3, num_fmt='#,##0')

ch = LineChart(); ch.title = '月度总量（不完整月已标注）'; ch.height = 8; ch.width = 22
data = Reference(ws, min_col=2, min_row=3, max_row=end_row)
cats = Reference(ws, min_col=1, min_row=4, max_row=end_row)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.y_axis.title = '条数'
ws.add_chart(ch, 'A32')

ch2 = LineChart(); ch2.title = '主要投诉类型月度走势'; ch2.height = 8; ch2.width = 22
data = Reference(ws, min_col=3, max_col=2 + len(seasonal_cols), min_row=3, max_row=end_row)
ch2.add_data(data, titles_from_data=True); ch2.set_categories(cats)
ch2.y_axis.title = '条数'
ws.add_chart(ch2, 'A52')
set_widths(ws, [12, 12, 16, 16, 18, 16, 16])

# ============ Sheet 4: 投诉类型 ============
ws = wb.create_sheet('4-投诉类型')
ws['A1'] = '投诉类型 Top 20（全期 2024-09-07 ~ 2026-09-07）'; ws['A1'].font = T_FONT
top20 = complaint.head(20).merge(cd[['投诉类型', '中位小时']], on='投诉类型', how='left')
tbl = top20[['投诉类型', '记录数', '占比%', '中位小时']].copy()
tbl['中位小时'] = tbl['中位小时'].round(1)
end_row = write_table(ws, tbl, 3, num_fmt='#,##0', pct_cols=('占比%',))
ws.cell(row=end_row + 2, column=1,
        value='注：Water System 在 2026-07 后被拆分为 Water Maintenance / Sewer Maintenance，故其全期计数不含 2026-08 之后。').font = Font(size=10, color='595959')

ch = BarChart(); ch.type = 'bar'; ch.title = '投诉类型 Top 15（记录数）'; ch.height = 12; ch.width = 20
data = Reference(ws, min_col=2, min_row=3, max_row=3 + 15)
cats = Reference(ws, min_col=1, min_row=4, max_row=3 + 15)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ws.add_chart(ch, 'F3')
set_widths(ws, [28, 14, 10, 12])

# ============ Sheet 5: 处理时效 ============
ws = wb.create_sheet('5-处理时效')
ws['A1'] = '机构处理时效（closed_date - created_date，小时；仅含正常关闭记录）'; ws['A1'].font = T_FONT
dur = agency[['机构', '记录数', '中位小时', 'P75小时', 'P90小时', '未关闭数']].copy()
for c_ in ['中位小时', 'P75小时', 'P90小时']:
    dur[c_] = dur[c_].round(1)
end_row = write_table(ws, dur, 3, num_fmt='#,##0')
notes = [
    '阅读提示：',
    '1) NYPD 的"关闭"多为到场处置后未发现违法即结案（见 resolution_description），中位 1.3 小时不代表问题被修复；',
    '2) HPD/DOB/DPR 涉及房屋检查与整改，耗时长是业务性质差异，不能直接比较优劣；',
    '3) DOHMH 的 P75≈P90≈1440 小时（60 天），呈自动到期结案特征；',
    '4) 快照时点未关闭记录存在右删失：靠近 2026-09 的请求尚未到关闭时间，不代表服务变慢；',
    '5) 已剔除 closed<created 的 1,877 条及晚于快照的 1 条异常；时效分母为时间正常的关闭记录 7,283,316 条（略多于 status=Closed 的 7,275,349 条，因少量非 Closed 状态记录亦有 closed_date）。',
]
r = end_row + 2
for n in notes:
    ws.cell(row=r, column=1, value=n).font = Font(size=10, color='595959'); r += 1

ch = BarChart(); ch.title = '各机构中位处理时长（小时，对数轴）'; ch.height = 10; ch.width = 20
data = Reference(ws, min_col=3, min_row=3, max_row=end_row)
cats = Reference(ws, min_col=1, min_row=4, max_row=end_row)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.y_axis.scaling.logBase = 10
ws.add_chart(ch, 'H3')

# 月度中位时效
mt = pd.read_csv(O + "/月度处理时长.csv").sort_values('月份')
r0 = end_row + 9
ws.cell(row=r0 - 1, column=1, value='月度中位处理时长（小时）').font = S_FONT
end_row2 = write_table(ws, mt[['月份', '关闭请求数', '中位小时', 'P75小时', 'P90小时']].round(1), r0, num_fmt='#,##0.0')
set_widths(ws, [14, 14, 12, 12, 12])

# ============ Sheet 6: 区域分布 ============
ws = wb.create_sheet('6-区域分布')
ws['A1'] = '区域（borough）总量与投诉构成'; ws['A1'].font = T_FONT
bor = borough[['borough', '记录数', '占比%']].copy()
bor = bor.rename(columns={'borough': '区域'})
end_row = write_table(ws, bor, 3, num_fmt='#,##0', pct_cols=('占比%',))

r0 = end_row + 2
ws.cell(row=r0 - 1, column=1, value='各区投诉类型构成（占该区总记录 %）').font = S_FONT
bmix = bshare.reset_index().rename(columns={'区域': 'borough'})
end_row2 = write_table(ws, bmix, r0, pct_cols=tuple(bmix.columns[1:]))
ws.cell(row=end_row2 + 2, column=1,
        value='注：Unspecified（6,331 条，0.1%）未列入构成表；记录数为绝对量，未经各区人口/户数归一，不代表人均强度。').font = Font(size=10, color='595959')

ch = BarChart(); ch.type = 'col'; ch.title = '各区记录数'; ch.height = 8; ch.width = 16
data = Reference(ws, min_col=2, min_row=3, max_row=end_row)
cats = Reference(ws, min_col=1, min_row=4, max_row=end_row)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ws.add_chart(ch, 'F3')
set_widths(ws, [16, 14, 10, 14, 16, 18, 16, 16, 14])

# ============ Sheet 7: 附-状态与渠道 ============
ws = wb.create_sheet('7-附-状态与渠道')
ws['A1'] = '状态与受理渠道构成（全期）'; ws['A1'].font = T_FONT
r = 3
ws.cell(row=r - 1, column=1, value='状态（status）').font = S_FONT
end_row = write_table(ws, status, r, num_fmt='#,##0', pct_cols=('占比%',))
r = end_row + 3
ws.cell(row=r - 1, column=1, value='受理渠道（open_data_channel_type）').font = S_FONT
end_row = write_table(ws, channel, r, num_fmt='#,##0', pct_cols=('占比%',))
r = end_row + 3
ws.cell(row=r - 1, column=1, value='未关闭记录 Top（快照 2026-09-07）').font = S_FONT
ncc = pd.read_csv(O + "/未关闭_按投诉类型.csv").head(10)
end_row = write_table(ws, ncc, r, num_fmt='#,##0')
set_widths(ws, [30, 14, 10])

wb.save(OUTX)
print("saved:", OUTX)
print("sheets:", wb.sheetnames)
