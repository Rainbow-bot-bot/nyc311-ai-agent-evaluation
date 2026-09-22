import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference, Series
import pandas as pd
import duckdb
import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

print("Starting Excel workbook generation...")

wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Palette
NAVY_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
STEEL_SUB_FILL = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
ICE_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
ALERT_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
KPI_BG_FILL = PatternFill(start_color="EAEEF7", end_color="EAEEF7", fill_type="solid")

WHITE_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=16, bold=True, color="1F4E78")
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color="1F4E78")
BOLD_FONT = Font(name="Calibri", size=10, bold=True, color="000000")
REG_FONT = Font(name="Calibri", size=10, color="000000")
NOTE_FONT = Font(name="Calibri", size=9, italic=True, color="595959")
KPI_VAL_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")
KPI_LBL_FONT = Font(name="Calibri", size=9, bold=True, color="595959")

THIN_BORDER_SIDE = Side(border_style="thin", color="D9D9D9")
BORDER_DATA = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE)
BORDER_HEADER = Border(left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=Side(border_style="medium", color="1F4E78"))
BORDER_TOTAL = Border(top=Side(border_style="thin", color="1F4E78"), bottom=Side(border_style="double", color="1F4E78"))

ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

def setup_sheet(ws, title, tab_color="1F4E78"):
    ws.title = title
    ws.sheet_properties.tabColor = tab_color
    ws.views.sheetView[0].showGridLines = True

def autofit_columns(ws, max_cols=20, min_len=12):
    for col in range(1, max_cols + 1):
        col_letter = get_column_letter(col)
        max_width = 0
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row=row, column=col)
            # ignore merged or very long title rows
            if row in [1, 2, 3] and cell.value and len(str(cell.value)) > 30:
                continue
            if cell.value:
                val_str = str(cell.value)
                max_width = max(max_width, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_width + 4, min_len)

# -------------------------------------------------------------
# SHEET 1: 执行摘要与看板导航
# -------------------------------------------------------------
print("Building Sheet 1: 执行摘要与看板导航...")
ws1 = wb.create_sheet(title="执行摘要与看板导航")
setup_sheet(ws1, "执行摘要与看板导航", "1F4E78")

ws1['A1'] = "纽约市 311 城市服务请求多维运营与异常诊断看板 (2024.09 - 2026.09)"
ws1['A1'].font = TITLE_FONT
ws1['A2'] = "数据口径：原始全量数据 N = 7,525,498 条 | 11个月同比期 (2024.10-2025.08 vs 2025.10-2026.08) | 独立证据导向分析"
ws1['A2'].font = NOTE_FONT

# KPI Blocks
kpi_data = [
    ("全量诉求总规模", "7,525,498 件", "覆盖729天，日均10,323件", "B4", "C5"),
    ("11个月同比增量", "+346,695 件 (+10.5%)", "Y1: 3.29M -> Y2: 3.64M", "D4", "E5"),
    ("线上诉求占比(Y2)", "46.8% (增量贡献94.1%)", "线上+23.7%，电话-7.9%", "F4", "G5"),
    ("首要承办机构", "NYPD 纽约市警局", "占全量45.2%，以违泊与噪音为主", "H4", "I5"),
    ("首要增长驱动项", "HPD 房屋局 (+42.4%)", "供暖/热水诉求激增6.36万件", "J4", "K5"),
]

for title, val, sub, top_left, bot_right in kpi_data:
    ws1.merge_cells(f"{top_left}:{bot_right}")
    tl_cell = ws1[top_left]
    tl_cell.fill = KPI_BG_FILL
    tl_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    tl_cell.value = f"{title}\n{val}\n{sub}"
    tl_cell.font = BOLD_FONT
    # border around merged range
    min_c, min_r = openpyxl.utils.coordinate_to_tuple(top_left)
    max_c, max_r = openpyxl.utils.coordinate_to_tuple(bot_right)
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            cell = ws1.cell(row=r, column=c)
            cell.fill = KPI_BG_FILL
            cell.border = BORDER_DATA

# Core Analytical Findings Table
ws1['B7'] = "六大核心分析发现与证据回查索引"
ws1['B7'].font = SECTION_FONT

headers_s1 = ["编号", "核心分析发现与业务结论", "量化证据摘要", "因果属性与证据边界", "对应分析工作表"]
for col_idx, h in enumerate(headers_s1, start=2):
    c = ws1.cell(row=8, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

findings_rows = [
    ("发现 1", "年度诉求稳步增长，线上数字化渠道成为绝对驱动力", "11个月同比净增346,695件(+10.5%)；线上渠道净增326,146件(+23.7%)，贡献整体增量的94.07%，电话呼叫下降-7.9%。", "描述性事实；反映市民报案渠道便利化迁移与使用黏性提升，非单纯城市问题恶化。", "02_年度增长与归因分解"),
    ("发现 2", "增长呈现四核驱动特征，近三分之二增量由四大问题集中产生", "供暖热水(+6.36万)、冰雪(+5.51万)、违章停车(+5.50万)、道路损坏(+4.80万)合计贡献增量63.95%；HPD贡献42.38%。", "描述性事实；增长并非全城各部门普涨，而是高度集中于供暖、交通执法与市政基建。", "02_年度增长与归因分解"),
    ("发现 3", "2026年早春极端暴风雪与寒潮触发链式市政危机", "2026年2月暴雪导致冰雪诉求单日破1.13万件；随后3月气温回升融雪诱发道路冻融，DOT坑洼诉求暴增428%(从4.3千飙升至2.28万件)。", "因果推断：气象冻融物理机理与日频时序前后继起完全咬合，经详细描述词核查证实。", "03_极端气候冲击与连锁反应"),
    ("发现 4", "DEP水务与排污分类在2026年8月发生系统级口径平移", "旧分类Water System与Sewer在2026年8月归零，新分类Water Maintenance与Sewer Maintenance接管，同比增幅逾1000倍属于统计假象。", "行政与IT系统重分类事实；排除了真实管道故障爆发假说，避免了口径漂移导致的严重误判。", "04_分类口径变更核查"),
    ("发现 5", "供暖问题呈现显著二八结构集聚与重复报案放大效应", "全市前1%地址(580栋建筑)集中了25.4%的供暖诉求；前5%集中51.5%；29.1%的供暖工单因同楼已有报案直接结案。", "结构事实；311工单量反映的是租户催促频次与建筑密度，并不等同于物理供热系统故障台数。", "05_结构集聚与二八法则"),
    ("发现 6", "行政区诉求极化显著，警务噪音执法面临高响应与低处罚落差", "布朗克斯承载41.3%的居住噪音和35.1%的供暖诉求；NYPD中位响应虽仅1.3小时，但44.1%噪音到场未见违规，传票率仅0.36%。", "执法运营事实；揭示了311作为民意宣泄窗口与现场执法标准之间的结构性错配。", "06_空间分布与响应执法")
]

for row_idx, r_data in enumerate(findings_rows, start=9):
    bg = ZEBRA_FILL if row_idx % 2 == 1 else PatternFill(fill_type=None)
    for col_idx, val in enumerate(r_data, start=2):
        c = ws1.cell(row=row_idx, column=col_idx, value=val)
        c.fill = bg
        c.font = REG_FONT
        c.border = BORDER_DATA
        if col_idx == 2:
            c.alignment = ALIGN_CENTER
            c.font = BOLD_FONT
        elif col_idx in (3, 4, 5):
            c.alignment = ALIGN_WRAP
        else:
            c.alignment = ALIGN_CENTER

# Evidence boundary and methodology notes
ws1['B16'] = "证据边界与分析规范原则 (严格执行 Evidence-Led 实验标准)"
ws1['B16'].font = SECTION_FONT

boundary_notes = [
    ("1. 观察期与右删失控制", "数据集起止为 2024-09-07 至 2026-09-05。9月均为破月，同比分析采用完整的11个月可比周期(10月1日-8月31日)。处置时长分析避开2026年7-9月未结案截断期，确保时效指标真实反映效率。"),
    ("2. 记录粒度与物理现实区分", "311数据最小粒度为服务请求记录(Service Request)，代表市民报告行为，非物理故障台数。多租户大楼单次锅炉故障会产生数百条工单，分析已通过重复工单描述符进行解耦。"),
    ("3. 口径变更与异常值下钻", "对出现千倍激增的指标必须下钻到底层分类与日频趋势，区分管理IT口径调整与真实事件，杜绝假增长。"),
    ("4. 因果声明审慎性", "严格区分描述性相关与因果关系。物理机理(冻融造成道路坑洼)有确凿时间序列与描述符证据支撑；其他宏观推论一律标记为未验证解释。")
]

for r_idx, (b_title, b_desc) in enumerate(boundary_notes, start=17):
    ws1.cell(row=r_idx, column=2, value=b_title).font = BOLD_FONT
    ws1.cell(row=r_idx, column=2).border = BORDER_DATA
    ws1.cell(row=r_idx, column=2).fill = ICE_FILL
    ws1.cell(row=r_idx, column=2).alignment = ALIGN_LEFT
    
    ws1.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=6)
    c_desc = ws1.cell(row=r_idx, column=3, value=b_desc)
    c_desc.font = REG_FONT
    c_desc.alignment = ALIGN_WRAP
    for c in range(3, 7):
        ws1.cell(row=r_idx, column=c).border = BORDER_DATA

ws1.column_dimensions['A'].width = 4
ws1.column_dimensions['B'].width = 16
ws1.column_dimensions['C'].width = 38
ws1.column_dimensions['D'].width = 46
ws1.column_dimensions['E'].width = 38
ws1.column_dimensions['F'].width = 28
for r in range(9, 15):
    ws1.row_dimensions[r].height = 42
for r in range(17, 21):
    ws1.row_dimensions[r].height = 36

# -------------------------------------------------------------
# SHEET 2: 数据质量与证据边界
# -------------------------------------------------------------
print("Building Sheet 2: 数据质量与证据边界...")
ws2 = wb.create_sheet(title="01_数据质量与证据边界")
setup_sheet(ws2, "01_数据质量与证据边界", "4B6F96")

ws2['A1'] = "数据完整性、逻辑一致性与右删失 (Right-Censoring) 检验"
ws2['A1'].font = TITLE_FONT
ws2['A2'] = "数据总量：7,525,498 条 | 字段数：44 个 | 主键重复数：0 (粒度唯一)"
ws2['A2'].font = NOTE_FONT

# Missingness table
ws2['A4'] = "表 1.1 核心业务字段缺失率核查"
ws2['A4'].font = SECTION_FONT

headers_miss = ["字段英文名", "业务含义", "非空记录数", "缺失记录数", "缺失率 (%)", "质量评估与业务影响"]
for col_idx, h in enumerate(headers_miss, start=1):
    c = ws2.cell(row=5, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

miss_data = [
    ("unique_key", "工单唯一键", 7525498, 0, 0.0, "100% 完整，全表主键，无任何重复，可作为唯一样本粒度"),
    ("created_date", "诉求创建时间", 7525498, 0, 0.0, "100% 完整，时间戳精确到秒，时间序列分析基准字段"),
    ("agency", "承办部门代码", 7525498, 0, 0.0, "100% 完整，共16个部门，无未识别或无效代码"),
    ("complaint_type", "诉求一级类型", 7525498, 0, 0.0, "100% 完整，共300+分类，全量诉求分类依据"),
    ("status", "工单流转状态", 7525498, 0, 0.0, "100% 完整，闭环状态占比96.7%，流转标识清晰"),
    ("borough", "所属行政区", 7525498, 0, 0.0, "100% 填写，仅0.08%为Unspecified，区域分析高度可靠"),
    ("open_data_channel_type", "报案接入渠道", 7525498, 0, 0.0, "100% 完整，含ONLINE/PHONE/MOBILE等，渠道演进分析核心"),
    ("descriptor", "诉求二级细分描述", 7503670, 21828, 0.29, "极低缺失，用于坑洼、噪音明细场景的深入下钻"),
    ("incident_zip", "事发地邮编", 7459939, 65559, 0.87, "缺失不足1%，可精确映射到街区/社区级别微观分析"),
    ("resolution_description", "部门结案处理说明", 7413511, 111987, 1.49, "低缺失，包含重复报案标识、现场执法结论，价值极高"),
    ("latitude / longitude", "地理空间坐标", 7395802, 129696, 1.72, "有效空间坐标占比98.3%，可支撑高精度空间聚类"),
    ("location_type", "场所类型", 6600325, 925173, 12.29, "缺失约12.3%，住宅类较全，公共街道部分未强制填写"),
    ("closed_date", "结案归档时间", 7285194, 240304, 3.19, "未闭环工单自然缺失，需结合右删失窗口进行时效分析")
]

for r_idx, row_vals in enumerate(miss_data, start=6):
    bg = ZEBRA_FILL if r_idx % 2 == 1 else PatternFill(fill_type=None)
    for c_idx, val in enumerate(row_vals, start=1):
        cell = ws2.cell(row=r_idx, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx in (1, 2):
            cell.alignment = ALIGN_LEFT
        elif c_idx in (3, 4):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 5:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
            cell.value = val / 100.0
        else:
            cell.alignment = ALIGN_LEFT

# Table 1.2: Right-Censoring Demonstration
ws2['A21'] = "表 1.2 月度队列结案率与右删失效应检验 (证明时效分析必须隔离末期破月)"
ws2['A21'].font = SECTION_FONT

headers_cohort = ["创建月份队列", "创建工单量", "已结案工单量", "结案率 (%)", "中位闭环耗时 (小时)", "删失状态与分析指引"]
for col_idx, h in enumerate(headers_cohort, start=1):
    c = ws2.cell(row=22, column=col_idx, value=h)
    c.fill = STEEL_SUB_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_cohort = pd.read_csv("data/monthly_cohort_resolution.csv")
for r_idx, r in df_cohort.iterrows():
    row_num = 23 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    
    ym = r['cohort_month']
    c_cnt = int(r['created_cnt'])
    cls_cnt = int(r['closed_cnt'])
    rate = float(r['closed_rate_pct']) / 100.0
    med_h = float(r['median_hours_closed']) if pd.notnull(r['median_hours_closed']) else 0.0
    
    if ym in ['2026-08', '2026-09']:
        status_txt = "严重右删失：距离数据截止不足30天，大量工单仍处于在办状态，不可作为时效评价基准"
        bg = ALERT_FILL
    elif ym in ['2026-06', '2026-07']:
        status_txt = "轻微右删失：复杂工单(如建筑/树木)尚未完全闭环"
    elif ym == '2024-09':
        status_txt = "期初破月：从9月7日开始，不影响结案观测"
    else:
        status_txt = "成熟观测期：结案率稳定在 96% - 99.5%，可可靠计算处置耗时"
        
    vals = [ym, c_cnt, cls_cnt, rate, med_h, status_txt]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws2.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_CENTER
        elif c_idx in (2, 3):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 4:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        elif c_idx == 5:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0.0"
        else:
            cell.alignment = ALIGN_LEFT

autofit_columns(ws2, max_cols=6, min_len=14)
ws2.column_dimensions['F'].width = 56

# -------------------------------------------------------------
# SHEET 3: 年度增长与归因分解
# -------------------------------------------------------------
print("Building Sheet 3: 年度增长与归因分解...")
ws3 = wb.create_sheet(title="02_年度增长与归因分解")
setup_sheet(ws3, "02_年度增长与归因分解", "2F5597")

ws3['A1'] = "纽约市 311 诉求年度增长动力学与多维分解 (11个月同期对比)"
ws3['A1'].font = TITLE_FONT
ws3['A2'] = "对比区间：基期 Year 1 (2024.10 - 2025.08) 对比 报告期 Year 2 (2025.10 - 2026.08) | 净增 +346,695 件 (+10.53%)"
ws3['A2'].font = NOTE_FONT

# Table 2.1: Monthly Progression Table
ws3['A4'] = "表 2.1 诉求量月度同比走势与突增月甄别"
ws3['A4'].font = SECTION_FONT

headers_m_prog = ["月份序号", "基期月份", "基期诉求量", "报告期月份", "报告期诉求量", "同比净增量", "同比增速 (%)", "业务异动定位"]
for col_idx, h in enumerate(headers_m_prog, start=1):
    c = ws3.cell(row=5, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_m_comp = pd.read_csv("data/month_yoy_comparison.csv")
# order: 10, 11, 12, 01, 02, 03, 04, 05, 06, 07, 08
month_order = ['10', '11', '12', '01', '02', '03', '04', '05', '06', '07', '08']
df_m_comp['sort_key'] = df_m_comp['month_num'].astype(str).str.zfill(2).apply(lambda x: month_order.index(x) if x in month_order else 99)
df_m_comp = df_m_comp.sort_values('sort_key').reset_index(drop=True)

for r_idx, r in df_m_comp.iterrows():
    row_num = 6 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    m_num = str(r['month_num']).zfill(2)
    y1_m = r['y1_month']
    y1_c = int(r['y1_count'])
    y2_m = r['y2_month']
    y2_c = int(r['y2_count'])
    diff = int(r['diff'])
    pct = float(r['growth_pct']) / 100.0
    
    if m_num == '02':
        anom = "★ 全年最大异动峰值：增量+79,326件(+31.1%)，暴雪与严寒双重叠加"
        bg = ALERT_FILL
    elif m_num == '03':
        anom = "★ 次大异动峰值：增量+61,166件(+21.8%)，雪后融雪引发道路坑洼大爆发"
        bg = ALERT_FILL
    elif m_num in ['05', '04']:
        anom = "春季平稳增长：气温回升，户外施工与停车执法诉求常态回升"
    else:
        anom = "常态波动区间：增幅在 0.1% - 9.7% 之间"
        
    vals = [r_idx + 1, y1_m, y1_c, y2_m, y2_c, diff, pct, anom]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws3.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx in (1, 2, 4):
            cell.alignment = ALIGN_CENTER
        elif c_idx in (3, 5, 6):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 7:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        else:
            cell.alignment = ALIGN_LEFT

# Summary row for Table 2.1
sum_row = 6 + len(df_m_comp)
ws3.cell(row=sum_row, column=1, value="合计/均值").font = BOLD_FONT
ws3.cell(row=sum_row, column=2, value="11个月累计").font = BOLD_FONT
ws3.cell(row=sum_row, column=3, value="=SUM(C6:C16)").font = BOLD_FONT
ws3.cell(row=sum_row, column=3).number_format = "#,##0"
ws3.cell(row=sum_row, column=4, value="11个月累计").font = BOLD_FONT
ws3.cell(row=sum_row, column=5, value="=SUM(E6:E16)").font = BOLD_FONT
ws3.cell(row=sum_row, column=5).number_format = "#,##0"
ws3.cell(row=sum_row, column=6, value="=E17-C17").font = BOLD_FONT
ws3.cell(row=sum_row, column=6).number_format = "#,##0"
ws3.cell(row=sum_row, column=7, value="=F17/C17").font = BOLD_FONT
ws3.cell(row=sum_row, column=7).number_format = "0.00%"
ws3.cell(row=sum_row, column=8, value="全城年度总增长基准线").font = BOLD_FONT
for c in range(1, 9):
    ws3.cell(row=sum_row, column=c).border = BORDER_TOTAL
    ws3.cell(row=sum_row, column=c).fill = ICE_FILL

# Table 2.2: Agency Growth Contribution
ws3['A19'] = "表 2.2 承办机构增长归因分解 (排名前列及异动机构)"
ws3['A19'].font = SECTION_FONT

headers_agency = ["机构代码", "主要职责领域", "基期件数 (Y1)", "报告期件数 (Y2)", "增减量", "机构同比增速", "对全城增量贡献率", "归因与业务角色"]
for col_idx, h in enumerate(headers_agency, start=1):
    c = ws3.cell(row=20, column=col_idx, value=h)
    c.fill = STEEL_SUB_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_agency = pd.read_csv("data/agency_growth.csv")
agency_role_map = {
    'HPD': ('房屋修缮与供暖保障', '首要增长源：贡献全城42.38%的净增量，供暖热水诉求全面爆发'),
    'DOT': ('交通运输与道路维护', '次要增长源：贡献19.28%净增量，雪后坑洼导致道路维修激增34.7%'),
    'NYPD': ('警务执法与社区秩序', '基础大盘：贡献16.87%净增量，违泊与占用车道诉求持续上升'),
    'DSNY': ('环境卫生与除雪作业', '气候敏感源：贡献9.56%净增量，2026年2月暴风雪除雪需求暴增'),
    'DPR': ('公园管理与倒木清理', '暴风雨与大风天气导致受损树木投诉增长24.8%'),
    'DOB': ('楼宇安全与施工管理', '脚手架与非法施工投诉常态增长18.2%'),
    'DEP': ('环境保护与水务管道', '稳定大盘：微增4.88%，底层发生分类替换(见Sheet 04)'),
    'OOS': ('外部流转服务机构', '新设/变更路由机构代码，基期无记录'),
    'TLC': ('出租车与网约车管理', '网约车服务与拒载投诉小幅回升7.6%'),
    'DCWP': ('消费者与劳工保护', '消费维权与商家违规举报稳定'),
    'NYC311-PRD': ('311内部测试或流转', '系统内部运维记录，极小量'),
    'OTI': ('城市技术与创新局', '体量极小，技术设施类'),
    'DOE': ('教育局学校相关', '学校周边问题投诉下降38.4%'),
    'DOHMH': ('卫生与精神健康局', '老鼠灭治与食品卫生投诉小幅收缩5.36%'),
    'DHS': ('无家可归者服务局', '收容所外流浪人员协助请求同比下降9.19%'),
    'EDC': ('经济发展署相关设施', '轮渡与码头设施投诉大幅减少42.98%')
}

for r_idx, r in df_agency.iterrows():
    row_num = 21 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    code = str(r['agency'])
    desc_area, role_desc = agency_role_map.get(code, ("综合行政", "其他常规机构"))
    y1_c = int(r['cnt_y1'])
    y2_c = int(r['cnt_y2'])
    diff = int(r['diff'])
    pct = float(r['growth_pct']) / 100.0 if pd.notnull(r['growth_pct']) else 0.0
    contrib = float(r['contribution_pct']) / 100.0 if pd.notnull(r['contribution_pct']) else 0.0
    
    if code in ['HPD', 'DOT', 'NYPD', 'DSNY']:
        bg = ICE_FILL
        
    vals = [code, desc_area, y1_c, y2_c, diff, pct, contrib, role_desc]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws3.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_CENTER
            cell.font = BOLD_FONT
        elif c_idx == 2:
            cell.alignment = ALIGN_LEFT
        elif c_idx in (3, 4, 5):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx in (6, 7):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        else:
            cell.alignment = ALIGN_LEFT

# Table 2.3: Channel Shift
ws3['A39'] = "表 2.3 市民报案渠道结构性跃迁 (数字化替代效应)"
ws3['A39'].font = SECTION_FONT

headers_chan = ["报案渠道", "基期件数 (Y1)", "基期份额 (%)", "报告期件数 (Y2)", "报告期份额 (%)", "渠道净增量", "渠道自身增速", "增量贡献解释"]
for col_idx, h in enumerate(headers_chan, start=1):
    c = ws3.cell(row=40, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_chan = pd.read_csv("data/channel_shift_yoy.csv")
chan_notes = {
    'ONLINE': '数字化核心引擎：净增32.6万件，单渠道解释了全城94.07%的增量，反映Web端低门槛报案普及',
    'PHONE': '传统热线萎缩：呼叫量减少7.33万件(-7.92%)，份额从28.1%跌至23.4%，市民报案习惯显著线上化',
    'MOBILE': '移动App保持稳定：诉求小幅微增3.11%，与智能手机客户端渗透率保持同步',
    'UNKNOWN': '系统未知渠道小幅增加，部分来自第三方集成接口或批量数据导入',
    'OTHER': '体量极小，其他辅助测试渠道'
}

for r_idx, r in df_chan.iterrows():
    row_num = 41 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    ch = str(r['channel'])
    y1_c = int(r['y1_cnt'])
    y1_s = float(r['y1_share_pct']) / 100.0
    y2_c = int(r['y2_cnt'])
    y2_s = float(r['y2_share_pct']) / 100.0
    diff = int(r['diff'])
    pct = float(r['growth_pct']) / 100.0
    c_note = chan_notes.get(ch, "常规渠道")
    
    if ch == 'ONLINE':
        bg = ICE_FILL
        
    vals = [ch, y1_c, y1_s, y2_c, y2_s, diff, pct, c_note]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws3.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_CENTER
            cell.font = BOLD_FONT
        elif c_idx in (2, 4, 6):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx in (3, 5, 7):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        else:
            cell.alignment = ALIGN_LEFT

# Add Chart for Channel Shift
chart_chan = BarChart()
chart_chan.type = "col"
chart_chan.style = 10
chart_chan.title = "诉求报案渠道同比结构变动 (Y1 vs Y2)"
chart_chan.y_axis.title = "工单量 (件)"
chart_chan.x_axis.title = "报案渠道"
chart_chan.width = 16
chart_chan.height = 10

data_ref = Reference(ws3, min_col=2, min_row=40, max_col=4, max_row=43) # ONLINE, PHONE, MOBILE
cats_ref = Reference(ws3, min_col=1, min_row=41, max_row=43)
chart_chan.add_data(data_ref, titles_from_data=True)
chart_chan.set_categories(cats_ref)
ws3.add_chart(chart_chan, "I40")

autofit_columns(ws3, max_cols=8, min_len=14)
ws3.column_dimensions['H'].width = 52

# -------------------------------------------------------------
# SHEET 4: 极端气候冲击与连锁反应
# -------------------------------------------------------------
print("Building Sheet 4: 极端气候冲击与连锁反应...")
ws4 = wb.create_sheet(title="03_极端气候冲击与连锁反应")
setup_sheet(ws4, "03_极端气候冲击与连锁反应", "1F4E78")

ws4['A1'] = "2026年早春极端气候冲击与链式基建损毁传导机制 (从暴雪到坑洼)"
ws4['A1'].font = TITLE_FONT
ws4['A2'] = "时序机制核验：2026年2月中下旬暴风雪与寒潮峰值 -> 3月气温回升融雪诱发道路严重冻融破损"
ws4['A2'].font = NOTE_FONT

# Table 3.1: February 2026 Daily Timeline
ws4['A4'] = "表 3.1 2026年2月暴风雪全生命周期日度追踪 (DSNY除雪 vs HPD供暖 vs NYPD堵路)"
ws4['A4'].font = SECTION_FONT

headers_feb = ["事发日期", "全日总诉求", "DSNY 冰雪诉求", "HPD 供暖热水诉求", "NYPD 占用车道/堵塞", "气候事件特征与应急阶段"]
for col_idx, h in enumerate(headers_feb, start=1):
    c = ws4.cell(row=5, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_feb_daily = pd.read_csv("data/feb26_daily_timeline.csv")
for r_idx, r in df_feb_daily.iterrows():
    row_num = 6 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    dt = str(r['c_date'])
    tot = int(r['total_day_cnt'])
    snow = int(r['snow_cnt'])
    heat = int(r['heat_cnt'])
    block = int(r['blocked_cnt'])
    
    if dt in ['2026-02-23', '2026-02-24']:
        bg = ALERT_FILL
        stage = "★ 特大暴风雪主峰：积雪严重瘫痪道路，24日单日冰雪报案破1.13万件，创历史极值"
    elif dt in ['2026-02-25', '2026-02-26']:
        bg = ICE_FILL
        stage = "暴雪清理次生期：铲雪车作业后堆雪阻断私人车道与人行道，车道投诉持续"
    elif dt in ['2026-02-07', '2026-02-08', '2026-02-09']:
        bg = ICE_FILL
        stage = "极寒寒潮脉冲：气温骤降造成多栋建筑锅炉过载停机，8日供暖投诉飙至6,118件"
    else:
        stage = "冬季常规常态"
        
    vals = [dt, tot, snow, heat, block, stage]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws4.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_CENTER
        elif c_idx in (2, 3, 4, 5):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        else:
            cell.alignment = ALIGN_LEFT

# Table 3.2: March 2026 DOT Road Condition Burst
ws4['H4'] = "表 3.2 雪后融雪道路损毁分解 (DOT 道路状况明细)"
ws4['H4'].font = SECTION_FONT

headers_road = ["道路损毁具体描述符", "2025年3月件数", "2026年3月件数", "3月同比净增", "倍数增长", "物理机理与因果支撑"]
for col_idx, h in enumerate(headers_road, start=8):
    c = ws4.cell(row=5, column=col_idx, value=h)
    c.fill = STEEL_SUB_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_street_desc = pd.read_csv("data/street_condition_desc.csv")
pothole_mech = {
    'Pothole': '冻融循环核心产物：融雪水渗入沥青裂缝，经历再结冰膨胀与重载碾压，坑洼激增428%',
    'Cave-in': '深层地基塌陷：融水流失侵蚀管线回填土方，导致路面深坑塌陷激增240%',
    'Defective Hardware': '融雪盐腐蚀与除雪机物理刮碰导致井盖/铁件松动损坏',
    'Failed Street Repair': '前期修补材料在低温冻融及除雪机械高频碾压下发生二次脱落失效',
    'Blocked - Construction': '常规施工阻断，受气温影响较小',
    'Rough, Pitted or Cracked Roads': '路面大面积网状开裂麻面，属于坑洼形成前期的典型病害',
    'Plate Condition - Noisy': '临时施工钢板受损或移位',
    'Line/Marking - Faded': '除雪作业机械刮擦与化雪盐腐蚀导致标线磨损',
    'Line/Marking - After Repaving': '重新铺路后标线涂覆诉求',
    'Dumpster - Construction Waste': '施工垃圾清运'
}

for r_idx, r in df_street_desc.head(8).iterrows():
    row_num = 6 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    desc = str(r['descriptor'])
    m25 = int(r['mar25_cnt'])
    m26 = int(r['mar26_cnt'])
    diff = m26 - m25
    mult = f"{m26 / max(m25, 1):.1f}x"
    mech = pothole_mech.get(desc, "常规道路损坏")
    
    if desc == 'Pothole':
        bg = ALERT_FILL
        
    vals = [desc, m25, m26, diff, mult, mech]
    for c_idx, val in enumerate(vals, start=8):
        cell = ws4.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 8:
            cell.alignment = ALIGN_LEFT
            cell.font = BOLD_FONT
        elif c_idx in (9, 10, 11):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 12:
            cell.alignment = ALIGN_CENTER
            cell.font = BOLD_FONT
        else:
            cell.alignment = ALIGN_LEFT

# Add Line Chart for February Snow vs Heat
chart_feb = LineChart()
chart_feb.title = "2026年2月暴雪与极寒日度诉求脉冲追踪"
chart_feb.style = 13
chart_feb.y_axis.title = "日诉求件数"
chart_feb.x_axis.title = "日期"
chart_feb.width = 18
chart_feb.height = 11

data_feb_ref = Reference(ws4, min_col=3, min_row=5, max_col=4, max_row=33) # Snow & Heat
cats_feb_ref = Reference(ws4, min_col=1, min_row=6, max_row=33)
chart_feb.add_data(data_feb_ref, titles_from_data=True)
chart_feb.set_categories(cats_feb_ref)
ws4.add_chart(chart_feb, "H16")

autofit_columns(ws4, max_cols=13, min_len=14)
ws4.column_dimensions['F'].width = 46
ws4.column_dimensions['M'].width = 50

# -------------------------------------------------------------
# SHEET 5: 分类口径变更核查
# -------------------------------------------------------------
print("Building Sheet 5: 分类口径变更核查...")
ws5 = wb.create_sheet(title="04_分类口径变更核查")
setup_sheet(ws5, "04_分类口径变更核查", "2F5597")

ws5['A1'] = "环境保护局 (DEP) 水务与排污系统分类口径平移溯源 (Taxonomy Drift Case)"
ws5['A1'].font = TITLE_FONT
ws5['A2'] = "方法规范核验：对数据中数万百分比激增的指标，必须下钻验证是否存在系统管理重命名，防止将口径假象误读为真实危机"
ws5['A2'].font = NOTE_FONT

ws5['A4'] = "表 4.1 DEP 水务与排水核心分类月度全历史演变 (证明2026年8月分类系统无缝交接)"
ws5['A4'].font = SECTION_FONT

headers_dep = ["业务月份", "旧: Water System", "新: Water Maintenance", "旧: Sewer", "新: Sewer Maintenance", "口径平移状态与管理审计诊断"]
for col_idx, h in enumerate(headers_dep, start=1):
    c = ws5.cell(row=5, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_dep = pd.read_csv("data/dep_taxonomy_shift.csv")
for r_idx, r in df_dep.iterrows():
    row_num = 6 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    ym = str(r['ym'])
    ws_c = int(r['water_system'])
    wm_c = int(r['water_maint'])
    se_c = int(r['sewer'])
    sm_c = int(r['sewer_maint'])
    
    if ym == '2026-08':
        bg = ALERT_FILL
        audit = "★ 正式彻底交接：旧分类(Water System/Sewer)突降为0，新分类(Maintenance)全额承接"
    elif ym == '2026-07':
        bg = ICE_FILL
        audit = "双轨并行过渡期：新分类开始出现千件级录入，开展系统割接联调"
    elif ym in ['2026-04', '2026-05', '2026-06']:
        audit = "系统试点小批量测试：新分类偶发数件至百件测试单"
    else:
        audit = "历史单一分类运行期：完全使用旧分类，新分类工单为零或个位数"
        
    vals = [ym, ws_c, wm_c, se_c, sm_c, audit]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws5.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_CENTER
        elif c_idx in (2, 3, 4, 5):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        else:
            cell.alignment = ALIGN_LEFT

# Synthesis Explanation Table
ws5['H4'] = "表 4.2 口径假象与真实业务判别对照表"
ws5['H4'].font = SECTION_FONT

headers_synth = ["分析维度", "未经验证的初级误判 (表象陷阱)", "深度核查后的事实结论 (证据支撑)"]
for col_idx, h in enumerate(headers_synth, start=8):
    c = ws5.cell(row=5, column=col_idx, value=h)
    c.fill = STEEL_SUB_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

synth_rows = [
    ("指标变动表象", "Water Maintenance 同比暴增 125,422%，Sewer Maintenance 暴增 73,233%", "旧分类与新分类在2026年7-8月发生一对一替换，总量平稳"),
    ("业务实质判断", "误判为纽约市供水与排污管道在2026年遭遇突发性、毁灭性系统坍塌", "实为DEP在2026年夏季对其工单管理系统或OpenData分类字典进行正规化升级"),
    ("真实运维负荷", "误判为市政维修工作量膨胀千倍，需要紧急巨额预算干预", "水务与下水道总诉求量同比仅温和变动，反映正常季节性波动"),
    ("数据分析戒律", "直接将分组聚合的Top增量列表当作业务结论输出，未做时序溯源", "落实Evidence-Led准则，逢异动必查时序连续性，查明分类交替节点")
]

for r_idx, (dim, trap, fact) in enumerate(synth_rows, start=6):
    row_num = r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    vals = [dim, trap, fact]
    for c_idx, val in enumerate(vals, start=8):
        cell = ws5.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 8:
            cell.alignment = ALIGN_CENTER
            cell.font = BOLD_FONT
        else:
            cell.alignment = ALIGN_WRAP

autofit_columns(ws5, max_cols=10, min_len=14)
ws5.column_dimensions['F'].width = 54
ws5.column_dimensions['I'].width = 40
ws5.column_dimensions['J'].width = 46
for r in range(6, 10):
    ws5.row_dimensions[r].height = 42

# -------------------------------------------------------------
# SHEET 6: 结构集聚与二八法则
# -------------------------------------------------------------
print("Building Sheet 6: 结构集聚与二八法则...")
ws6 = wb.create_sheet(title="05_结构集聚与二八法则")
setup_sheet(ws6, "05_结构集聚与二八法则", "4B6F96")

ws6['A1'] = "供暖与热水诉求的微观二八法则分布与重复报案解耦"
ws6['A1'].font = TITLE_FONT
ws6['A2'] = "分析样本：HPD HEAT/HOT WATER 共 651,594 条 | 涵盖全市 58,083 个独立建筑地址 | 深度剖析结构集聚与治理抓手"
ws6['A2'].font = NOTE_FONT

# Table 5.1: Pareto Distribution Table
ws6['A4'] = "表 5.1 供暖诉求地址级帕累托 (Pareto) 集中度分层"
ws6['A4'].font = SECTION_FONT

headers_pareto = ["建筑地址分层", "涉及建筑数", "占全市建筑比 (%)", "累计诉求件数", "占全市诉求比 (%)", "栋均诉求量", "治理特征与政策启示"]
for col_idx, h in enumerate(headers_pareto, start=1):
    c = ws6.cell(row=5, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_pareto = pd.read_csv("data/heat_pareto_analysis.csv")
pareto_notes = {
    '1. Top 1% Addresses': '极度重度问题建筑：仅580栋大楼，霸占全城超过四分之一诉求，属于慢性劣质老旧住宅',
    '2. Top 1-5% Addresses': '重度问题建筑：2,324栋大楼，前5%大楼合计吸纳全城 51.47% 的供暖投诉，集中治理抓手极清晰',
    '3. Top 5-10% Addresses': '中度问题建筑：多为采暖季短期故障或个别单元供暖管网水力失调',
    '4. Top 10-20% Addresses': '轻度偶发问题建筑：采暖季偶发1-2次报修',
    '5. Remaining 80% Addresses': '广泛长尾散户：全市4.6万栋大楼仅产生19.57%投诉，平均每栋大楼不足3件'
}

for r_idx, r in df_pareto.iterrows():
    row_num = 6 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    tier = str(r['tier'])
    a_cnt = int(r['addr_count'])
    a_pct = float(r['addr_share_pct']) / 100.0
    c_cnt = int(r['complaint_count'])
    c_pct = float(r['complaint_share_pct']) / 100.0
    avg_per = c_cnt / a_cnt
    p_note = pareto_notes.get(tier, "常规分布")
    
    if "Top 1%" in tier:
        bg = ALERT_FILL
    elif "Top 1-5%" in tier:
        bg = ICE_FILL
        
    vals = [tier, a_cnt, a_pct, c_cnt, c_pct, avg_per, p_note]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws6.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_LEFT
            cell.font = BOLD_FONT
        elif c_idx in (2, 4):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx in (3, 5):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        elif c_idx == 6:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0.0"
        else:
            cell.alignment = ALIGN_LEFT

# Table 5.2: Top 10 Chronic Problem Buildings
ws6['A13'] = "表 5.2 全市供暖投诉量最高的前10栋慢性问题住宅"
ws6['A13'].font = SECTION_FONT

headers_top10 = ["排名", "涉案住宅详细地址", "所属行政区", "两年累计供暖诉求", "日均诉求量 (全期)", "重点监管与修缮建议"]
for col_idx, h in enumerate(headers_top10, start=1):
    c = ws6.cell(row=14, column=col_idx, value=h)
    c.fill = STEEL_SUB_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_heat_addr = pd.read_csv("data/heat_top_addresses.csv")
for r_idx, r in df_heat_addr.iterrows():
    row_num = 15 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    addr = str(r['incident_address'])
    boro = str(r['borough'])
    cnt = int(r['heat_complaints'])
    daily_avg = cnt / 729.0
    
    if r_idx == 0:
        advice = "全城之冠：两年超6,000次投诉，平均每天8次以上，属于严重的房东失职或中央锅炉全面瘫痪"
        bg = ALERT_FILL
    elif r_idx < 4:
        advice = "重点稽查对象：两年超1,800次投诉，应纳入HPD替代执法项目(AEP)实施强制性政府代修"
    else:
        advice = "重点督办物业：高频重复投诉集中，建议联合房管局下达行政纠正令与罚单"
        
    vals = [r_idx + 1, addr, boro, cnt, daily_avg, advice]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws6.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_CENTER
            cell.font = BOLD_FONT
        elif c_idx in (2, 3):
            cell.alignment = ALIGN_LEFT
        elif c_idx == 4:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 5:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.0"
        else:
            cell.alignment = ALIGN_LEFT

# Table 5.3: HPD Resolution Outcome & Duplicate Clumping
ws6['H4'] = "表 5.3 HPD 供暖诉求结案处置方式与重复报案实证"
ws6['H4'].font = SECTION_FONT

headers_hpd_res = ["结案处置结论类别", "涉及工单数", "占比 (%)", "管理含义与业务解读"]
for col_idx, h in enumerate(headers_hpd_res, start=8):
    c = ws6.cell(row=5, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

hpd_res_data = [
    ("整栋已有报案(重复单)", 189512, 29.08, "近三成工单为同栋楼其他租户对同一事件的重复提交，表明真实停暖事件数需除以约 1.41"),
    ("无法进入建筑/单元核查", 95048, 14.59, "租户不在家或物业拒绝配合开门，制约现场执法规制落地"),
    ("到场核实已恢复供暖", 90110, 13.83, "房东在检查员抵达前临时点火加温，或阶段性恢复"),
    ("电话回访确认已修复", 60802, 9.33, "通过电话远程回访闭环，未实际出动现场执法人员"),
    ("已开展检查待出结果", 59081, 9.07, "现场勘验已完成，流转至文书与违规核定流程"),
    ("与租户核实已恢复", 58784, 9.02, "与报案租户直接沟通确认温度达标"),
    ("现场勘验未见违规", 38597, 5.92, "实测室内温度符合法定采暖期最低标准(白天68度/夜间62度)"),
    ("核实违规并下达罚单", 25187, 3.87, "仅 3.87% 的工单直接闭环于行政处罚，反映从诉求到实质处罚的漫长门槛"),
    ("其他情形与气温不满足", 34473, 5.29, "室外气温高于法定采暖要求(>55华氏度)等其他结案情形")
]

for r_idx, (cat, cnt, pct, meaning) in enumerate(hpd_res_data, start=6):
    row_num = r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    if "重复单" in cat:
        bg = ALERT_FILL
    vals = [cat, cnt, pct / 100.0, meaning]
    for c_idx, val in enumerate(vals, start=8):
        cell = ws6.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 8:
            cell.alignment = ALIGN_LEFT
            cell.font = BOLD_FONT
        elif c_idx == 9:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 10:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        else:
            cell.alignment = ALIGN_LEFT

autofit_columns(ws6, max_cols=11, min_len=14)
ws6.column_dimensions['G'].width = 46
ws6.column_dimensions['K'].width = 50

# -------------------------------------------------------------
# SHEET 7: 空间分布与响应执法
# -------------------------------------------------------------
print("Building Sheet 7: 空间分布与响应执法...")
ws7 = wb.create_sheet(title="06_空间分布与响应执法")
setup_sheet(ws7, "06_空间分布与响应执法", "1F4E78")

ws7['A1'] = "行政区诉求极化特征与跨部门处置时效实证 (无偏观测窗口)"
ws7['A1'].font = TITLE_FONT
ws7['A2'] = "分析基准：固定观测期 (剔除未成熟右删失月) | NYPD响应速度与执法实质性落地差异"
ws7['A2'].font = NOTE_FONT

# Table 6.1: Borough Complaint Composition
ws7['A4'] = "表 6.1 五大行政区重点问题结构与极化特征 (全期横截面)"
ws7['A4'].font = SECTION_FONT

headers_boro = ["行政区", "全量诉求总数", "全城诉求占比", "违章停车件数", "居住噪音件数", "供暖热水件数", "占用车道件数", "核心矛盾与治理画像"]
for col_idx, h in enumerate(headers_boro, start=1):
    c = ws7.cell(row=5, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

boro_portrait = {
    'BROOKLYN': '全城体量之首：占全城30.0%，违章停车(45.7万)与占道(13.1万)规模庞大，空间资源争夺激烈',
    'QUEENS': '私家车密集与路况痛点：占全城24.2%，堵占车道(15.4万)与道路坑洼(7.37万)均为全城最高',
    'BRONX': '人居环境与供暖危机集中地：仅占人口17%，却产生全城 41.3% 的噪音与 35.1% 的供暖问题，民生痛点最重',
    'MANHATTAN': '高密度租住与商住混杂：供暖投诉达15.6万件，人行道与商业噪音突出，因私家车少而堵塞车道极少(仅9.2千)',
    'STATEN ISLAND': '低密度独立住宅区：诉求总量仅占全城3.85%，主要以道路设施与路灯维护为主'
}

df_boro_ct = pd.read_csv("data/borough_by_complaint_type.csv")
for r_idx, r in df_boro_ct.iterrows():
    row_num = 6 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    boro = str(r['borough'])
    tot = int(r['total_complaints'])
    tot_pct = tot / 7525498.0
    park = int(r['illegal_parking'])
    noise = int(r['noise_residential'])
    heat = int(r['heat_hot_water'])
    blk = int(r['blocked_driveway'])
    port = boro_portrait.get(boro, "常规区域")
    
    if boro == 'BRONX':
        bg = ALERT_FILL
    elif boro in ['BROOKLYN', 'QUEENS']:
        bg = ICE_FILL
        
    vals = [boro, tot, tot_pct, park, noise, heat, blk, port]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws7.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_CENTER
            cell.font = BOLD_FONT
        elif c_idx in (2, 4, 5, 6, 7):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 3:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        else:
            cell.alignment = ALIGN_LEFT

# Table 6.2: Agency Resolution Time in Fixed Window
ws7['A13'] = "表 6.2 主要承办部门处理时效同期对比 (固定窗口: 10月至次年4月，避免右删失)"
ws7['A13'].font = SECTION_FONT

headers_resp = ["承办部门", "对比周期", "总纳管件数", "已结案件数", "结案率 (%)", "中位闭环耗时 (小时)", "平均闭环耗时 (小时)", "时效波动与管理归因"]
for col_idx, h in enumerate(headers_resp, start=1):
    c = ws7.cell(row=14, column=col_idx, value=h)
    c.fill = STEEL_SUB_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

df_resp = pd.read_csv("data/agency_resolution_time.csv")
resp_notes = {
    ('NYPD', 'Y1_Winter_Spring'): '警务极速响应基准：巡逻车就近处置，中位耗时仅 1.64 小时',
    ('NYPD', 'Y2_Winter_Spring'): '响应进一步加快至 1.31 小时，数字化派单效率持续提高',
    ('DEP', 'Y1_Winter_Spring'): '水务管道常规处置，中位 34.67 小时',
    ('DEP', 'Y2_Winter_Spring'): '响应时效提速至 25.05 小时，现场抢修与水质检测加快',
    ('DSNY', 'Y1_Winter_Spring'): '环卫常规作业，中位 30.48 小时',
    ('DSNY', 'Y2_Winter_Spring'): '受2026年2月特大暴雪影响，清雪与清运承压，中位耗时拉长至 38.77 小时',
    ('DOT', 'Y1_Winter_Spring'): '道路维护常规周期，中位 40.91 小时',
    ('DOT', 'Y2_Winter_Spring'): '雪后坑洼暴增近4倍，沥青填补维修排期积压，中位耗时延长至 48.47 小时',
    ('HPD', 'Y1_Winter_Spring'): '房屋暖气检查入户，中位 69.74 小时 (~2.9天)',
    ('HPD', 'Y2_Winter_Spring'): '供暖负荷虽大增，中位保持在 71.46 小时 (~3.0天)，运转体系相对稳健',
    ('DPR', 'Y1_Winter_Spring'): '树木修剪与隐患排查，中位 212 小时 (~8.8天)',
    ('DPR', 'Y2_Winter_Spring'): '重点险情树木优先排查，中位耗时缩短至 96 小时',
    ('DOB', 'Y1_Winter_Spring'): '建筑工程检查涉及法定义务与工程师复核，中位 271 小时 (~11.3天)',
    ('DOB', 'Y2_Winter_Spring'): '常规巡查流程优化，中位耗时缩短至 175 小时 (~7.3天)'
}

for r_idx, r in df_resp.iterrows():
    row_num = 15 + r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    ag = str(r['agency'])
    per = "基期 (24.10-25.04)" if r['period'] == 'Y1_Winter_Spring' else "报告期 (25.10-26.04)"
    tot_c = int(r['total_cases'])
    cls_c = int(r['closed_valid_cases'])
    rate = float(r['close_rate_pct']) / 100.0
    med_h = float(r['median_hours'])
    avg_h = float(r['mean_hours'])
    r_note = resp_notes.get((ag, r['period']), "常规时效记录")
    
    if ag in ['DOT', 'DSNY'] and 'Y2' in r['period']:
        bg = ALERT_FILL
    elif ag == 'NYPD':
        bg = ICE_FILL
        
    vals = [ag, per, tot_c, cls_c, rate, med_h, avg_h, r_note]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws7.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx in (1, 2):
            cell.alignment = ALIGN_CENTER
            if c_idx == 1:
                cell.font = BOLD_FONT
        elif c_idx in (3, 4):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx == 5:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        elif c_idx in (6, 7):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0.0"
        else:
            cell.alignment = ALIGN_LEFT

# Table 6.3: NYPD Enforcement Reality Table
ws7['A31'] = "表 6.3 NYPD 警务执法落地真实落差 (违章停车 vs 居住噪音处置结果对比)"
ws7['A31'].font = SECTION_FONT

headers_nypd_act = ["执法处置结果", "违章停车件数", "违章停车占比 (%)", "居住噪音件数", "居住噪音占比 (%)", "执法有效性与制度性反思"]
for col_idx, h in enumerate(headers_nypd_act, start=1):
    c = ws7.cell(row=32, column=col_idx, value=h)
    c.fill = NAVY_HEADER_FILL
    c.font = WHITE_FONT
    c.alignment = ALIGN_CENTER
    c.border = BORDER_HEADER

nypd_action_comp = [
    ("现场开具传票 (Summons)", 260746, 22.50, 3148, 0.36, "违泊执法具有静态物证，传票率达22.5%；而噪音稍纵即逝，传票率仅0.36%，相差超60倍"),
    ("到场未见违规/当事人已离开", 173985, 15.01, 388452, 44.06, "噪音投诉近半数在警察抵达时源头已停止，或属于主观生活声响未达到法定噪音超标红线"),
    ("判定无需采取警察行动", 187558, 16.18, 113279, 12.85, "日常邻里琐事或合法泊车被误报，警方判定不构成治安或行政违规"),
    ("采取整改/劝导消除措施", 132649, 11.44, 152371, 17.28, "现场口头劝导调低音量、要求车主当场驶离，实现非惩罚性矛盾化解"),
    ("民警无法进入建筑物/单元", 680, 0.06, 86011, 9.76, "住宅楼门禁或租户不开门导致近10%的居住噪音根本无法入户调查，形成执法死角"),
    ("实施现场拘捕 (Arrest)", 95943, 8.28, 107147, 12.15, "涉案伴随更严重的治安冲突、通缉犯识别或拒绝配合等升级警情"),
    ("其他未细分/说明已更新", 307492, 26.53, 31140, 3.53, "系统流转与常规记录归档")
]

for r_idx, (act, p_cnt, p_pct, n_cnt, n_pct, act_exp) in enumerate(nypd_action_comp, start=33):
    row_num = r_idx
    bg = ZEBRA_FILL if row_num % 2 == 1 else PatternFill(fill_type=None)
    if "传票" in act:
        bg = ICE_FILL
    elif "未见违规" in act or "无法进入" in act:
        bg = ALERT_FILL
        
    vals = [act, p_cnt, p_pct / 100.0, n_cnt, n_pct / 100.0, act_exp]
    for c_idx, val in enumerate(vals, start=1):
        cell = ws7.cell(row=row_num, column=c_idx, value=val)
        cell.fill = bg
        cell.font = REG_FONT
        cell.border = BORDER_DATA
        if c_idx == 1:
            cell.alignment = ALIGN_LEFT
            cell.font = BOLD_FONT
        elif c_idx in (2, 4):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "#,##0"
        elif c_idx in (3, 5):
            cell.alignment = ALIGN_RIGHT
            cell.number_format = "0.00%"
        else:
            cell.alignment = ALIGN_LEFT

autofit_columns(ws7, max_cols=8, min_len=14)
ws7.column_dimensions['H'].width = 54
ws7.column_dimensions['F'].width = 54

# Save workbook
output_path = r"D:\项目2\Gemini\第二轮\最终成果.xlsx"
wb.save(output_path)
print(f"Final Excel workbook successfully created and saved at: {output_path}")
