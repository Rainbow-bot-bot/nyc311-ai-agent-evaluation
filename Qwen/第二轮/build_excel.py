"""
生成 最终成果.xlsx —— 普通读者可直接阅读的看板。
所有数字来自 work/tables/*.csv（由 compute_final.py 产出），保证与 分析说明.md 一致、可回查。
图表用 openpyxl 原生图表（Excel 内可交互）。
"""
import pandas as pd, numpy as np, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

T = r"D:\项目2\增强实验\Qwen\work\tables"
OUTX = r"D:\项目2\增强实验\Qwen\最终成果.xlsx"
def L(n): return pd.read_csv(os.path.join(T, n))

# ---- 样式常量 ----
TITLE = Font(name="Microsoft YaHei", size=15, bold=True, color="FFFFFF")
H1    = Font(name="Microsoft YaHei", size=12, bold=True, color="1F4E78")
HDR   = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
BODY  = Font(name="Microsoft YaHei", size=10)
SMALL = Font(name="Microsoft YaHei", size=9, color="595959")
FILL_TITLE = PatternFill("solid", fgColor="1F4E78")
FILL_HDR   = PatternFill("solid", fgColor="2E75B6")
FILL_SEC   = PatternFill("solid", fgColor="DDEBF7")
FILL_NOTE  = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CTR  = Alignment(horizontal="center", vertical="center")

wb = Workbook()

def set_widths(ws, widths):
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width = w

def put_title(ws, text, ncol, sub=None):
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncol)
    c=ws.cell(1,1,text); c.font=TITLE; c.fill=FILL_TITLE; c.alignment=Alignment(vertical="center",horizontal="left")
    ws.row_dimensions[1].height=26
    r=2
    if sub:
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncol)
        c=ws.cell(2,1,sub); c.font=SMALL; c.alignment=WRAP; ws.row_dimensions[2].height=30
        r=3
    return r+1

def put_section(ws, row, text, ncol):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncol)
    c=ws.cell(row,1,text); c.font=H1; c.fill=FILL_SEC; c.alignment=Alignment(vertical="center")
    ws.row_dimensions[row].height=20
    return row+1

def put_note(ws, row, text, ncol, height=None):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncol)
    c=ws.cell(row,1,text); c.font=SMALL; c.fill=FILL_NOTE; c.alignment=WRAP
    ws.row_dimensions[row].height=height or (15*max(1,(len(text)//(ncol*14)+1)))
    return row+1

def write_df(ws, df, start_row, start_col=1, num_fmt=None, header=True):
    """写 DataFrame，返回 (end_row, end_col)。num_fmt: {col_name: fmt}"""
    num_fmt = num_fmt or {}
    cols = list(df.columns)
    r = start_row
    if header:
        for j,cname in enumerate(cols, start_col):
            c=ws.cell(r,j,str(cname)); c.font=HDR; c.fill=FILL_HDR; c.alignment=CTR; c.border=BORDER
        r+=1
    for _,rowv in df.iterrows():
        for j,cname in enumerate(cols, start_col):
            v=rowv[cname]
            if pd.isna(v): v=None
            elif isinstance(v,(np.integer,)): v=int(v)
            elif isinstance(v,(np.floating,)): v=float(v)
            c=ws.cell(r,j,v); c.font=BODY; c.border=BORDER
            if cname in num_fmt and v is not None: c.number_format=num_fmt[cname]
            if isinstance(v,str) and len(v)>18: c.alignment=WRAP
        r+=1
    return r-1, start_col+len(cols)-1

# ============================================================
# Sheet 1: 导读与核心发现
# ============================================================
ws = wb.active; ws.title = "导读与核心发现"
set_widths(ws, [4,26,60,20,16])
r = put_title(ws, "NYC 311 服务请求分析看板（2024-09 至 2026-09）", 5,
    "数据源: D:\\项目1\\原始数据 (25个只读Parquet分片, 7,525,498条请求) | 记录粒度: 一行=一个311请求(unique_key全局唯一) | "
    "所有数字可在各表页与工作表 work/tables/*.csv 回查")

r = put_section(ws, r, "一、这个看板回答什么", 5)
intro = ("从陌生的 NYC 311 原始数据出发，独立确认数据口径后，回答四个问题：(1) 两年里市民反映的问题总量与结构如何变化？"
         "(2) 变化的主要驱动是什么？(3) 311 的响应/结案速度如何、是否随负载变化？(4) 哪些数字是真实信号、哪些是口径/删失造成的假象？"
         "下面每条发现都标注了【事实】【推断】或【未验证】。")
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=5); ws.cell(r,1,intro).font=BODY
ws.cell(r,1).alignment=WRAP; ws.row_dimensions[r].height=58; r+=2

r = put_section(ws, r, "二、五个核心发现", 5)
hdr=["#","发现","关键数字与说明","性质","证据表页"]
for j,h in enumerate(hdr,1):
    c=ws.cell(r,j,h); c.font=HDR; c.fill=FILL_HDR; c.alignment=CTR; c.border=BORDER
r+=1
finds=[
 ["1","总量两年增长约一成，但主要由天气驱动",
  "Oct–Aug 对齐窗口 W1(329.3万)→W2(364.0万)，净增 34.7万(+10.5%)。其中天气敏感6类(雪冰/供暖/道路/受损树木/管道/漏水)贡献 59.7%；剔除后基础增速仅 +5.0%。",
  "【事实】+【推断】","YoY增长分解"],
 ["2","2025–26 冬季异常严酷，并留下春季道路损坏",
  "Snow or Ice 冬三月合计(12–2月) 8,794→63,722(≈7×)；HEAT/HOT WATER 2026-01 峰值 79,928(全期单月最高)；Street Condition 2026-03 冲到 28,690(常态4–7k)且五个区普涨1.9–2.8×，持续到夏季。",
  "【事实】现象；【推断】更冷冬季","冬季与道路"],
 ["3","响应速度高度分层，但主要类别总体稳定",
  "成熟队列全市中位 7.7h、24h内结案 60.9%、7天内 83.7%。NYPD类(违停/噪声)中位≈1.3h；HPD住房类以天计(不卫生250h、门窗291h、漏水247h)。HEAT/HOT WATER 中位 41.8→42.0h，负载上升下仍稳定。",
  "【事实】","响应时长"],
 ["4","受理渠道持续向线上迁移",
  "ONLINE 份额 41.8%→46.8%(+326k)；PHONE 28.1%→23.4%(绝对 −73k, −7.9%)；MOBILE 22.7%→21.2%。电话是唯一绝对下降的渠道。",
  "【事实】","渠道与行政区"],
 ["5","若干'暴涨暴跌'是口径/删失/集中事件，非真实趋势",
  "Water Maintenance YoY'+125422%'=DEP水系2026-07重分类；Drug Activity'−68%'=2025夏Queens尖峰后崩塌；2025-01噪声尖峰=单一ZIP 10466事件；近3月closed_rate骤降=右删失(尚未到结案时间)。",
  "【事实】+【未验证】原因","数据质量与口径"],
]
for f in finds:
    for j,v in enumerate(f,1):
        c=ws.cell(r,j,v); c.font=BODY; c.alignment=WRAP; c.border=BORDER
        if j==1 or j==4: c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[r].height=64; r+=1
r+=1

r = put_section(ws, r, "三、各表页怎么看", 5)
guide=[
 ["数据概况","总量/时间范围/粒度/缺失率/QA剔除，确认这份数据能代表什么、不能代表什么。"],
 ["月度趋势","25个月的请求量、结案率、成熟队列处理时长中位数(含折线图)；两端为部分月，近3月结案率受右删失影响。"],
 ["YoY增长分解","把 +10.5% 拆成'天气'与'基础增长'，并列出增长/下降最多的投诉类型。"],
 ["冬季与道路","发现2的证据：雪冰/供暖/道路逐月曲线 + 道路损坏的分区对比。"],
 ["响应时长","发现3的证据：按机构与投诉类型的处理时长，及供暖类逐月负载vs时长。"],
 ["渠道与行政区","发现4的证据：渠道份额迁移 + 行政区分布与增速。"],
 ["数据质量与口径","发现5的证据：所有需要谨慎解读的口径变更、集中事件、删失，逐条列出处理方式。"],
 ["TOP类型明细","全期请求量前20的投诉类型：占比、结案率、处理时长。"],
]
for g in guide:
    ws.cell(r,2,g[0]).font=Font(name="Microsoft YaHei",size=10,bold=True); ws.cell(r,2).border=BORDER
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=5)
    c=ws.cell(r,3,g[1]); c.font=BODY; c.alignment=WRAP; c.border=BORDER
    ws.row_dimensions[r].height=30; r+=1
r+=1
disc=("事实=可从原始数据直接计算得到；推断=基于数据模式的最可能解释(如'更冷冬季')，未引入外部气象数据；"
      "未验证=数据不足以确定原因(如某些类别骤变的政策/执法动因)。原始输入只读，未做任何改写。")
r=put_note(ws,r,disc,5,42)

# ============================================================
# Sheet 2: 数据概况
# ============================================================
ws = wb.create_sheet("数据概况")
set_widths(ws,[30,30,58])
r = put_title(ws,"数据概况与证据边界",3,"记录粒度、时间覆盖、缺失与质量剔除。对应 work/tables/t_overview.csv")
r = put_section(ws,r,"关键口径指标",3)
ov=L("t_overview.csv")
er,_=write_df(ws,ov,r); r=er+2
r = put_note(ws,r,"说明：closed_dt 缺失=未结案(右删失)，集中在近3个月，属正常；负处理时长/未来结案日期等极少数异常已在处理时长统计中按规则剔除(见 t_overview 末两行)。borough 'Unspecified'(0.08%)与 channel 'UNKNOWN'(7.85%)保留为独立类别，不臆测归属。",3,60)

# ============================================================
# Sheet 3: 月度趋势
# ============================================================
ws = wb.create_sheet("月度趋势")
set_widths(ws,[12,14,14,20,12,30,14])
r = put_title(ws,"月度趋势（25个月）",7,"对应 work/tables/t_monthly.csv。2024-09与2026-09为部分月；近3个月结案率受右删失影响")
mo=L("t_monthly.csv")
tbl_start=r
er,_=write_df(ws,mo,r,num_fmt={"closed_rate_pct":"0.0","median_res_h_matured":"0.0","yoy_pct_aligned":"+0.0"})
tbl_end=er; r=er+2
# 折线图: 请求量
ch=LineChart(); ch.title="每月请求量(两端为部分月)"; ch.height=8; ch.width=20; ch.style=12
data=Reference(ws,min_col=2,min_row=tbl_start,max_row=tbl_end)
cats=Reference(ws,min_col=1,min_row=tbl_start+1,max_row=tbl_end)
ch.add_data(data,titles_from_data=True); ch.set_categories(cats)
ch.y_axis.title="请求数"; ch.x_axis.title="月份"; ch.x_axis.delete=False; ch.y_axis.delete=False
ws.add_chart(ch,f"A{r}")
# 折线图: 处理时长中位数(成熟队列)
ch2=LineChart(); ch2.title="处理时长中位数(小时, 成熟队列; 空缺=右删失未计)"; ch2.height=8; ch2.width=20; ch2.style=13
d2=Reference(ws,min_col=4,min_row=tbl_start,max_row=tbl_end)
ch2.add_data(d2,titles_from_data=True); ch2.set_categories(cats)
ch2.y_axis.title="中位小时"; ch2.x_axis.title="月份"; ch2.x_axis.delete=False; ch2.y_axis.delete=False
ws.add_chart(ch2,f"A{r+16}")
r+=32
r=put_note(ws,r,"读图要点：请求量在冬季(12–2月)走高，2026-02达334,690(YoY+31%)。处理时长中位数在冬季(供暖季)升高、夏季回落，体现负载对时长的影响；2026-07起为空，因这些请求多数尚未结案(右删失)，计入会低估时长。",7,58)

# ============================================================
# Sheet 4: YoY增长分解
# ============================================================
ws = wb.create_sheet("YoY增长分解")
set_widths(ws,[42,20,18,60])
r = put_title(ws,"YoY 增长分解（Oct–Aug 对齐窗口）",4,
   "W1=2024-10..2025-08, W2=2025-10..2026-08, 各11个完整月。对应 t_decomp.csv / t_yoy_types.csv")
r = put_section(ws,r,"把 +10.5% 拆开",4)
dec=L("t_decomp.csv")
er,_=write_df(ws,dec,r); r=er+2
r = put_note(ws,r,"结论：总量增长真实存在，但约六成来自天气敏感类别在严冬的放量；剔除天气后的基础增速约 +5.0%(与夏季月 YoY +8~12.5% 方向一致)。因此不能把 +10.5% 简单读成'市民诉求普遍上升一成'。",4,46)

# 增长/下降 TOP 类型
r = put_section(ws,r,"投诉类型 YoY 变化（按净增量排序）",4)
yoy=L("t_yoy_types.csv")
set_widths(ws,[42,18,18,14,14,18])
top=yoy.head(12).copy(); bot=yoy.tail(10).copy()
r = put_note(ws,r,"增长 TOP12（净增量最大）",6,16)
er,_=write_df(ws,top,r,num_fmt={"pct_change":"+0.0","contrib_pct_of_net":"0.0"}); r=er+1
# 条形图: 增长TOP12 净增
bc=BarChart(); bc.type="bar"; bc.title="YoY 净增量 TOP12(条)"; bc.height=9; bc.width=18
dref=Reference(ws,min_col=4,min_row=er-len(top)+1,max_row=er)  # delta col(数据行)
cref=Reference(ws,min_col=1,min_row=er-len(top)+1,max_row=er)
bc.add_data(dref,titles_from_data=False); bc.set_categories(cref); bc.legend=None
bc.x_axis.delete=False; bc.y_axis.delete=False
ws.add_chart(bc,f"H{er-len(top)}")
r+=1
r = put_note(ws,r,"下降 TOP10（净减量最大；注意其中多项为口径/集中事件，见'数据质量与口径'）",6,16)
er2,_=write_df(ws,bot,r,num_fmt={"pct_change":"+0.0","contrib_pct_of_net":"0.0"}); r=er2+2

# ============================================================
# Sheet 5: 冬季与道路
# ============================================================
ws = wb.create_sheet("冬季与道路")
set_widths(ws,[12,14,18,16,14,14,16])
r = put_title(ws,"发现2：严酷冬季与春季道路损坏",7,"对应 t_winter_monthly.csv / t_street_borough.csv")
wm=L("t_winter_monthly.csv")
r = put_section(ws,r,"天气敏感类别逐月",7)
ws_start=r
er,_=write_df(ws,wm,r); ws_end=er; r=er+2
# Snow or Ice 折线
c1=LineChart(); c1.title="Snow or Ice 逐月(两个冬季对比)"; c1.height=8; c1.width=18; c1.style=12
col_snow=list(wm.columns).index("Snow or Ice")+1
d=Reference(ws,min_col=col_snow,min_row=ws_start,max_row=ws_end)
cats=Reference(ws,min_col=1,min_row=ws_start+1,max_row=ws_end)
c1.add_data(d,titles_from_data=True); c1.set_categories(cats); c1.x_axis.delete=False; c1.y_axis.delete=False
ws.add_chart(c1,f"A{r}")
# HEAT/HOT WATER 折线
c2=LineChart(); c2.title="HEAT/HOT WATER 逐月(供暖季)"; c2.height=8; c2.width=18; c2.style=13
col_heat=list(wm.columns).index("HEAT/HOT WATER")+1
d=Reference(ws,min_col=col_heat,min_row=ws_start,max_row=ws_end)
c2.add_data(d,titles_from_data=True); c2.set_categories(cats); c2.x_axis.delete=False; c2.y_axis.delete=False
ws.add_chart(c2,f"J{r}")
r+=16
# Street Condition 折线
c3=LineChart(); c3.title="Street Condition 逐月(2026-03 骤升并持续)"; c3.height=8; c3.width=18; c3.style=14
col_st=list(wm.columns).index("Street Condition")+1
d=Reference(ws,min_col=col_st,min_row=ws_start,max_row=ws_end)
c3.add_data(d,titles_from_data=True); c3.set_categories(cats); c3.x_axis.delete=False; c3.y_axis.delete=False
ws.add_chart(c3,f"A{r}")
r+=17
r = put_section(ws,r,"道路损坏的分区对比（Mar–Jun, 2026 vs 2025）",7)
sb=L("t_street_borough.csv")
sb_start=r
er,_=write_df(ws,sb,r,num_fmt={"倍数":"0.00"}); sb_end=er; r=er+2
c4=BarChart(); c4.type="col"; c4.title="Street Condition: 2026 vs 2025 (Mar–Jun, 分区)"; c4.height=8; c4.width=18
d=Reference(ws,min_col=2,max_col=3,min_row=sb_start,max_row=sb_end)
cats=Reference(ws,min_col=1,min_row=sb_start+1,max_row=sb_end)
c4.add_data(d,titles_from_data=True); c4.set_categories(cats); c4.x_axis.delete=False; c4.y_axis.delete=False
ws.add_chart(c4,f"A{r}")
r+=17
r=put_note(ws,r,"读图要点：Snow or Ice 在2025-26冬季(12–2月)量级约为上一年冬季的7倍；供暖投诉2026-01达峰79,928。Street Condition 在2026-03(严冬后冻融)骤升至28,690并维持高位，且五个区普涨——说明是广泛的路面损坏，而非单一来源。'更冷冬季'为基于投诉模式的推断，未引入外部气象数据。",7,60)

# ============================================================
# Sheet 6: 响应时长
# ============================================================
ws = wb.create_sheet("响应时长")
set_widths(ws,[26,14,12,12,12,12])
r = put_title(ws,"发现3：响应/结案时长（成熟队列）",6,
   "成熟队列=created≤2026-06-30(结案率≥95%)，规避右删失。对应 t_res_agency.csv / t_res_type.csv / t_heat_monthly.csv")
r = put_section(ws,r,"按机构",6)
ra=L("t_res_agency.csv")
ra_start=r
er,_=write_df(ws,ra,r,num_fmt={"median_h":"0.0","p90_h":"0.0","mean_h":"0.0","pct_24h":"0.0"}); ra_end=er; r=er+2
c1=BarChart(); c1.type="bar"; c1.title="各机构处理时长中位数(小时, 对数刻度:跨1.3h~6296h)"; c1.height=10; c1.width=18
d=Reference(ws,min_col=4,min_row=ra_start,max_row=ra_end)  # median_h
cats=Reference(ws,min_col=1,min_row=ra_start+1,max_row=ra_end)
c1.add_data(d,titles_from_data=True); c1.set_categories(cats); c1.legend=None; c1.x_axis.delete=False; c1.y_axis.delete=False
c1.y_axis.scaling.logBase=10; c1.y_axis.title="中位小时(对数)"
ws.add_chart(c1,f"A{r}")
r+=20
r = put_section(ws,r,"按投诉类型 TOP15（W1 vs W2 中位数）",6)
rt=L("t_res_type.csv")
rt_start=r
er,_=write_df(ws,rt,r,num_fmt={"median_W1_h":"0.0","median_W2_h":"0.0","delta_h":"+0.0"}); rt_end=er; r=er+2
c2=BarChart(); c2.type="col"; c2.title="TOP类型处理时长中位数: W1 vs W2(小时, 对数刻度)"; c2.height=10; c2.width=20
d=Reference(ws,min_col=3,max_col=4,min_row=rt_start,max_row=rt_end)
cats=Reference(ws,min_col=1,min_row=rt_start+1,max_row=rt_end)
c2.add_data(d,titles_from_data=True); c2.set_categories(cats); c2.x_axis.delete=False; c2.y_axis.delete=False
c2.y_axis.scaling.logBase=10; c2.y_axis.title="中位小时(对数)"
ws.add_chart(c2,f"A{r}")
r+=20
r = put_section(ws,r,"HEAT/HOT WATER：负载 vs 时长（逐月）",6)
hh=L("t_heat_monthly.csv")
hh_start=r
er,_=write_df(ws,hh,r,num_fmt={"median_res_h":"0.0"}); hh_end=er; r=er+2
c3=LineChart(); c3.title="供暖类逐月请求量(柱)与中位时长(线)"; c3.height=9; c3.width=20
d=Reference(ws,min_col=2,min_row=hh_start,max_row=hh_end)
cats=Reference(ws,min_col=1,min_row=hh_start+1,max_row=hh_end)
c3.add_data(d,titles_from_data=True); c3.set_categories(cats); c3.x_axis.delete=False; c3.y_axis.delete=False
c3b=LineChart(); d2=Reference(ws,min_col=3,min_row=hh_start,max_row=hh_end)
c3b.add_data(d2,titles_from_data=True); c3b.y_axis.axId=200; c3b.y_axis.title="中位小时"; c3b.y_axis.crosses="max"
c3+=c3b
ws.add_chart(c3,f"A{r}")
r+=19
r=put_note(ws,r,"读图要点：NYPD受理的违停/噪声类中位约1.3小时(接警即处置)；HPD住房类(不卫生/门窗/漏水/管道)以天计，反映需入户整改的执法流程。主要类别 W1→W2 中位数基本平稳(供暖41.8→42.0h)，说明在总量上升与严冬负载下，响应速度总体维持。",6,58)

# ============================================================
# Sheet 7: 渠道与行政区
# ============================================================
ws = wb.create_sheet("渠道与行政区")
set_widths(ws,[16,14,14,14,14,14,14])
r = put_title(ws,"发现4：渠道迁移 与 行政区分布",7,"对应 t_channel.csv / t_channel_monthly.csv / t_borough.csv")
r = put_section(ws,r,"渠道 W1 vs W2",7)
ch=L("t_channel.csv")
if ch.columns[0]!="channel": ch=ch.rename(columns={ch.columns[0]:"channel"})
ch_start=r
er,_=write_df(ws,ch,r,num_fmt={"W1_share_pct":"0.0","W2_share_pct":"0.0","pct_change":"+0.0"}); ch_end=er; r=er+2
c1=BarChart(); c1.type="col"; c1.title="渠道份额%(W1 vs W2)"; c1.height=8; c1.width=16
d=Reference(ws,min_col=4,max_col=5,min_row=ch_start,max_row=ch_end)
cats=Reference(ws,min_col=1,min_row=ch_start+1,max_row=ch_end)
c1.add_data(d,titles_from_data=True); c1.set_categories(cats); c1.x_axis.delete=False; c1.y_axis.delete=False
ws.add_chart(c1,f"A{r}")
r+=17
r = put_section(ws,r,"渠道逐月（份额迁移可视化）",7)
cm=L("t_channel_monthly.csv")
cm_start=r
er,_=write_df(ws,cm,r); cm_end=er; r=er+2
c2=LineChart(); c2.title="ONLINE vs PHONE 逐月量"; c2.height=8; c2.width=20
cols=list(cm.columns)
for name in ["ONLINE","PHONE"]:
    if name in cols:
        ci=cols.index(name)+1
        c2.add_data(Reference(ws,min_col=ci,min_row=cm_start,max_row=cm_end),titles_from_data=True)
c2.set_categories(Reference(ws,min_col=1,min_row=cm_start+1,max_row=cm_end)); c2.x_axis.delete=False; c2.y_axis.delete=False
ws.add_chart(c2,f"A{r}")
r+=17
r = put_section(ws,r,"行政区（全期分布 + YoY）",7)
bo=L("t_borough.csv")
if bo.columns[0]!="borough": bo=bo.rename(columns={bo.columns[0]:"borough"})
bo_start=r
er,_=write_df(ws,bo,r,num_fmt={"share_pct":"0.0","pct_change":"+0.0"}); bo_end=er; r=er+2
c3=BarChart(); c3.type="col"; c3.title="各区请求量(全期)"; c3.height=8; c3.width=16
d=Reference(ws,min_col=2,min_row=bo_start,max_row=bo_end)
cats=Reference(ws,min_col=1,min_row=bo_start+1,max_row=bo_end)
c3.add_data(d,titles_from_data=True); c3.set_categories(cats); c3.legend=None; c3.x_axis.delete=False; c3.y_axis.delete=False
ws.add_chart(c3,f"A{r}")
r+=17
r=put_note(ws,r,"读图要点：ONLINE 是唯一份额与绝对量都上升的主力渠道(+326k)，PHONE 绝对量下降(-73k)——受理在数字化迁移。行政区中 Brooklyn 量最大(占30%)，Staten Island 增速最高(+23.8%)但其净增主要来自 Snow or Ice(见分解)，属天气驱动。",7,50)

# ============================================================
# Sheet 8: 数据质量与口径
# ============================================================
ws = wb.create_sheet("数据质量与口径")
set_widths(ws,[22,46,40,52])
r = put_title(ws,"数据质量与口径护栏",4,"逐条列出会误导解读的口径变更、集中事件与删失，及其处理方式。对应 t_quality_events.csv / t_drug_monthly.csv")
r = put_section(ws,r,"需谨慎解读的事件",4)
qe=L("t_quality_events.csv")
er,_=write_df(ws,qe,r); r=er+2
r = put_section(ws,r,"Drug Activity 逐月（尖峰后崩塌，YoY −68% 有误导性）",4)
da=L("t_drug_monthly.csv")
da_start=r
er,_=write_df(ws,da,r); da_end=er; r=er+2
c1=LineChart(); c1.title="Drug Activity 逐月"; c1.height=8; c1.width=20
d=Reference(ws,min_col=2,min_row=da_start,max_row=da_end)
cats=Reference(ws,min_col=1,min_row=da_start+1,max_row=da_end)
c1.add_data(d,titles_from_data=True); c1.set_categories(cats); c1.legend=None; c1.x_axis.delete=False; c1.y_axis.delete=False
ws.add_chart(c1,f"A{r}")
r+=17
r=put_note(ws,r,"原则：这些'暴涨暴跌'不应直接当作真实趋势。分类变更需合并新旧类别后再比；集中事件(单一ZIP/区)应识别为局部现象；近月结案率下降是删失而非变慢。原因层面(政策/执法/报送口径)标注为未验证，不臆测。",4,50)

# ============================================================
# Sheet 9: TOP类型明细
# ============================================================
ws = wb.create_sheet("TOP类型明细")
set_widths(ws,[28,14,12,14,14,12])
r = put_title(ws,"全期请求量 TOP20 投诉类型",6,"对应 t_top_types.csv。share_pct=占全期7,525,498条的比重；处理时长为成熟队列口径")
tt=L("t_top_types.csv")
tt_start=r
er,_=write_df(ws,tt,r,num_fmt={"share_pct":"0.00","closed_rate_pct":"0.0","median_res_h":"0.0","p90_res_h":"0.0"}); tt_end=er; r=er+2
c1=BarChart(); c1.type="bar"; c1.title="TOP20 投诉类型请求量"; c1.height=13; c1.width=18
d=Reference(ws,min_col=2,min_row=tt_start,max_row=tt_end)
cats=Reference(ws,min_col=1,min_row=tt_start+1,max_row=tt_end)
c1.add_data(d,titles_from_data=True); c1.set_categories(cats); c1.legend=None; c1.x_axis.delete=False; c1.y_axis.delete=False
ws.add_chart(c1,f"A{r}")

# 冻结首行(各数据表页)
for s in wb.worksheets:
    s.sheet_view.showGridLines=False

wb.save(OUTX)
print("saved:", OUTX)
print("sheets:", wb.sheetnames)
