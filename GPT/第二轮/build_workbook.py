"""Create the requested workbook from verified aggregates; never writes to inputs.
Artifact-tool loader is not exposed in this session. Its external dependency cache
is outside the experiment's authorized paths. Use the installed XlsxWriter fallback,
then validate/recalculate/render in installed Microsoft Excel via verify_workbook.py.
"""
from pathlib import Path
import json,datetime as dt
import pandas as pd,numpy as np,xlsxwriter
R=Path(__file__).parent;E=R/'evidence'
def read(n):return pd.read_csv(E/(n+'.csv')).replace({np.nan:None})
quality=read('quality').iloc[0];std=json.loads((E/'standardization_summary.json').read_text())
base=std['base_total'];cur=std['current_total'];delta=cur-base
w=xlsxwriter.Workbook(R/'最终成果.xlsx')
w.set_properties({'title':'NYC 311 服务请求：增长、集中与渠道变化','subject':'2025 与 2026 年 1—8 月同窗比较','author':'GPT','comments':'仅使用指定 Parquet 与输入清单；静态源汇总配合 Excel 计算公式。'})
font='Microsoft YaHei'
def fmt(**kw):return w.add_format({'font_name':font,'font_size':11,'font_color':'#253247','valign':'vcenter',**kw})
f={'text':fmt(),'title':fmt(font_size=17,bold=True),'head':fmt(bold=True,bg_color='#263F63',font_color='white',text_wrap=True,align='center'),'note':fmt(font_color='#526176',font_size=10),'wrap':fmt(text_wrap=True),'num':fmt(num_format='#,##0'),'pct':fmt(num_format='0.0%'),'pp':fmt(num_format='0.00" pp"'),'date':fmt(num_format='yyyy-mm-dd'),'total':fmt(bold=True,top=1,num_format='#,##0'),'section':fmt(bold=True,font_color='#263F63'),'link':fmt(font_color='#2459A6',underline=1),'small':fmt(font_size=10,text_wrap=True)}
sheets={n:w.add_worksheet(n) for n in ['总览','月度趋势','同比拆解','峰值核查','渠道变化','数据质量','口径与复现']}
for name,s in sheets.items():
    s.hide_gridlines(2);s.set_default_row(23);s.set_zoom(85);s.set_landscape();s.set_paper(9);s.fit_to_pages(1,0);s.set_margins(.3,.3,.4,.4);s.set_footer('&LNYC 311 数据分析&R&P / &N');s.set_column('A:A',30,f['text']);s.set_column('B:L',16,f['text'])
    s.write(1,0,{'总览':'NYC 311 服务请求看板','月度趋势':'请求量增长是否持续？','同比拆解':'哪些类别与机构贡献了增量？','峰值核查':'雪冰尖峰与道路持续增长','渠道变化':'ONLINE 占比的变化能否由构成解释？','数据质量':'数据边界与影响结论的质量问题','口径与复现':'口径、字段解释与复现入口'}[name],f['title'])
    s.set_row(1,29)
    if name not in ['总览','口径与复现']:s.freeze_panes(5,1)
    s.repeat_rows(0,4)
sheets['总览'].set_tab_color('#263F63')
def val(s,r,c,v,form=None):
    if v is None:s.write_blank(r,c,None,form or f['text'])
    elif isinstance(v,(int,float,np.number)):s.write_number(r,c,float(v),form or f['num'])
    else:s.write(r,c,v,form or f['text'])
def headers(s,row,values):
    s.write_row(row,0,values,f['head']);s.set_row(row,34)
def formula(s,r,c,text,cached,style='num'):s.write_formula(r,c,text,f[style],cached)
def addtable(s,r,headers_,rows,filter_=True):
    headers(s,r,headers_)
    for i,row in enumerate(rows,r+1):
        for j,v in enumerate(row):val(s,i,j,v)
    if filter_:s.autofilter(r,0,r+len(rows),len(headers_)-1)
    return r+len(rows)

# Monthly totals and exact matched-month comparison.
s=sheets['月度趋势'];s.write(3,0,'同比主窗：两年各 1 月 1 日至 9 月 1 日前，均为 243 天；单位：请求记录。',f['note'])
t=read('compare_month').sort_values('month');headers(s,4,['月份','2025年记录','2026年记录','增加记录','同比增幅'])
for r,row in enumerate(t.itertuples(),5):
    s.write_number(r,0,int(row.month),f['num']);val(s,r,1,row.base);val(s,r,2,row.current)
    formula(s,r,3,f'=C{r+1}-B{r+1}',row.delta);formula(s,r,4,f'=D{r+1}/B{r+1}',row.growth,'pct')
s.write(13,0,'1—8月合计',f['section'])
for c_,value in [(1,base),(2,cur),(3,delta)]:formula(s,13,c_,f'=SUM({chr(65+c_)}6:{chr(65+c_)}13)',value)
formula(s,13,4,'=D14/B14',delta/base,'pct')
s.write(16,0,'全量覆盖：首尾月份只展示，不纳入完整月份同比',f['section'])
headers(s,18,['创建月份','记录数','有记录天数','该月最早时间','该月最晚时间','覆盖判断'])
monthly=read('monthly')
for r,row in enumerate(monthly.itertuples(),19):
    s.write_datetime(r,0,dt.datetime.fromisoformat(row.mth),f['date']);val(s,r,1,row.n);val(s,r,2,row.day_count);s.write(r,3,row.first_ts,f['small']);s.write(r,4,row.last_ts,f['small']);s.write(r,5,'部分月份' if r in [19,43] else '日历日齐全',f['text']);s.set_row(r,32)
s.write(45,0,'合计',f['section']);formula(s,45,1,'=SUM(B20:B44)',int(quality.n))
s.set_column('D:E',23);s.set_column('F:F',18)
s.set_column('G:G',3);s.set_column('H:H',29);s.set_column('I:L',16)
s.write_row(4,7,['敏感性检查','2025记录','2026记录','增加记录','同比增幅'],f['head'])
for r,row in enumerate(read('sensitivity').itertuples(),5):
    s.write(r,7,row.check);val(s,r,8,row.base);val(s,r,9,row.current);formula(s,r,10,f'=J{r+1}-I{r+1}',row.delta);formula(s,r,11,f'=K{r+1}/I{r+1}',row.growth,'pct')
s.write(11,7,'同一排除规则应用于两年；仍保留全部状态和渠道。',f['note'])
s.print_area(0,0,45,11)

# Additive category decomposition, with exact denominators.
s=sheets['同比拆解'];s.write(3,0,'2026年1—8月 对比 2025年同期；贡献率 = 类别增量 / 全部净增量 287,342。负贡献会抵消增长。',f['note'])
cat=read('compare_complaint_type');headers(s,4,['原始请求类别','2025年记录','2026年记录','增加记录','同比增幅','净增量贡献'])
for r,row in enumerate(cat.itertuples(),5):
    s.write(r,0,row.complaint_type);val(s,r,1,row.base);val(s,r,2,row.current);formula(s,r,3,f'=C{r+1}-B{r+1}',row.delta)
    formula(s,r,4,f'=IF(B{r+1}=0,"无基期",D{r+1}/B{r+1})',row.growth if row.base else '无基期','pct')
    formula(s,r,5,f"=D{r+1}/'月度趋势'!$D$14",row.delta/delta,'pct')
end=5+len(cat);s.write(end,0,'全部类别合计',f['section'])
for col_,v in [(1,base),(2,cur),(3,delta)]:formula(s,end,col_,f'=SUM({chr(65+col_)}6:{chr(65+col_)}{end})',v)
formula(s,end,5,f'=SUM(F6:F{end})',1,'pct');s.autofilter(4,0,end-1,5)
s.set_column('A:A',44);s.set_column('G:G',3);s.set_column('H:H',20);s.set_column('I:L',15)
for start,name,colname in [(4,'compare_agency','agency'),(25,'compare_borough','borough')]:
    x=read(name);s.write_row(start,7,['机构' if colname=='agency' else '行政区','2025年','2026年','增加记录','贡献率'],f['head']);s.set_row(start,34)
    for r,row in enumerate(x.itertuples(),start+1):
        s.write(r,7,getattr(row,colname));val(s,r,8,row.base);val(s,r,9,row.current);formula(s,r,10,f'=J{r+1}-I{r+1}',row.delta);formula(s,r,11,f"=K{r+1}/'月度趋势'!$D$14",row.delta/delta,'pct')
s.write(34,7,'行政区比较未按人口归一化。',f['note']);s.write(36,7,'机构、类别、行政区是独立拆解，不可跨表相加。',f['note'])
s.conditional_format(5,3,end-1,3,{'type':'cell','criteria':'<','value':0,'format':fmt(font_color='#AA4444')});s.print_area(0,0,end,11)

# Peak and description evidence.
s=sheets['峰值核查'];s.write(3,0,'同窗记录数；“剔除最高5天”仅削减2026年，用作保守压力测试，不是等长窗口同比。',f['note']);s.set_column('A:A',27);s.set_column('B:B',30);s.set_column('C:H',16)
peak=read('peak_sensitivity');headers(s,4,['类别','2025年记录','2026年记录','2026最高5天','最高5天占比','剔除后2026','相对2025增幅','同比上升月数'])
for r,row in enumerate(peak.itertuples(),5):
    s.write(r,0,row.category);val(s,r,1,row.base);val(s,r,2,row.current);val(s,r,3,row.top5);formula(s,r,4,f'=D{r+1}/C{r+1}',row.top5_share,'pct');formula(s,r,5,f'=C{r+1}-D{r+1}',row.current_without_top5);formula(s,r,6,f'=F{r+1}/B{r+1}-1',row.growth_without_top5,'pct');val(s,r,7,row.months_up)
s.write(10,0,'描述层核查：坑洞增加 35,310 条，占道路状况净增量的 74.0%。',f['section'])
des=read('descriptor_comparison').head(12);headers(s,12,['类别','原始描述','2025年记录','2026年记录','增加记录'])
for r,row in enumerate(des.itertuples(),13):
    s.write(r,0,row.complaint_type);s.write(r,1,row.descriptor,f['wrap']);val(s,r,2,row.base);val(s,r,3,row.current);formula(s,r,4,f'=D{r+1}-C{r+1}',row.delta);s.set_row(r,34)
s.write(27,0,'雪冰集中度不能证明天气原因；缺少天气、人员和独立事件编号。',f['note'])
s.write(28,0,'排除 Snow Tracking 描述后：雪冰仍从 7,723 增至 40,651 条（+426.4%）。',f['note'])
df=read('daily_focus');top=df[(df.complaint_type=='Snow or Ice')&df.dt.str.startswith('2026')].nlargest(5,'n')
headers(s,31,['2026雪冰最高5天','雪冰记录数'])
for r,row in enumerate(top.itertuples(),32):s.write_datetime(r,0,dt.datetime.fromisoformat(row.dt),f['date']);val(s,r,1,row.n)
pc=read('peak_control').iloc[0];s.write(39,0,f'最大单日 2026-02-24：11,370 条；涉及 {int(pc.address_n):,} 个非空原始地址值。',f['note']);s.write(40,0,'地址数仅检验集中性，同址可能多次请求；地址文本未标准化，不等于独立事件。',f['note'])
# Native line chart of daily snow requests, supported by full daily table below.
daily=df[df.complaint_type=='Snow or Ice'].copy();daily['day']=daily.dt.str[5:];daily['yr']=daily.dt.str[:4]
grid=daily.pivot_table(index='day',columns='yr',values='n',aggfunc='sum',fill_value=0)
all_days=pd.date_range('2025-01-01','2025-08-31').strftime('%m-%d');grid=grid.reindex(all_days,fill_value=0)
headers(s,43,['日历月日','2025雪冰','2026雪冰'])
for r,(day,row) in enumerate(grid.iterrows(),44):s.write(r,0,day);val(s,r,1,row['2025']);val(s,r,2,row['2026'])
s.print_area(0,0,40,7)

# Channel standardization and raw channel contributions.
s=sheets['渠道变化'];s.write(3,0,'ONLINE 保留源系统标签；不假定等于所有数字渠道，也不据此识别用户迁移。',f['note']);s.set_column('A:A',32)
ch=read('compare_channel');headers(s,4,['原始渠道','2025年记录','2026年记录','增加记录','2025占比','2026占比','占比变化(pp)'])
for r,row in enumerate(ch.itertuples(),5):
    s.write(r,0,row.channel);val(s,r,1,row.base);val(s,r,2,row.current);formula(s,r,3,f'=C{r+1}-B{r+1}',row.delta);formula(s,r,4,f"=B{r+1}/'月度趋势'!$B$14",row.base/base,'pct');formula(s,r,5,f"=C{r+1}/'月度趋势'!$C$14",row.current/cur,'pct');formula(s,r,6,f'=(F{r+1}-E{r+1})*100',(row.current/cur-row.base/base)*100,'pp')
s.write(12,0,'固定构成检查：共同存在的“月份 × 类别 × 机构”单元，固定2025年记录权重。',f['section'])
headers(s,14,['指标','数值','解释'])
items=[('共同单元2025记录',std['common_base'],'作为标准化权重分母'),('共同单元2026记录',std['common_current'],'2026年在相同单元内的记录'),('2025覆盖率',std['common_base']/base,'共同单元记录 / 全部2025同窗记录'),('2026覆盖率',std['common_current']/cur,'共同单元记录 / 全部2026同窗记录'),('2025 ONLINE占比',std['base_common_rate'],'共同单元，按2025年构成'),('固定构成2026 ONLINE占比',std['fixed_mix_current_rate'],'使用同一2025年构成权重'),('固定构成占比变化(pp)',(std['fixed_mix_current_rate']-std['base_common_rate'])*100,'同单元内的渠道结构仍有变化')]
for r,(label,num,note) in enumerate(items,15):s.write(r,0,label);val(s,r,1,num,f['pct'] if r in [17,18,19,20] else f['pp'] if r==21 else f['num']);s.write(r,2,note,f['note'])
s.write(24,0,'结果支持渠道记录结构改变；无法确认是用户选择、接入方式或编码规则变化。',f['note'])
s.write(25,0,'共同单元以外记录不参与标准化，覆盖率在上表披露；不把标准化比例当作实际比例。',f['note'])
st=read('channel_standardization');headers(s,28,['月份','原始类别','机构','2025总数','2026总数','2025 ONLINE','2026 ONLINE','2025权重','2025占比','2026占比','基期加权项','当期加权项'])
s.set_column('B:B',38);s.set_column('C:C',19)
for r,row in enumerate(st.itertuples(index=False),29):
    mo,cat_,agency,t0,t1,o0,o1,weight=row
    for j,v in enumerate([mo,cat_,agency,t0,t1,o0,o1]):val(s,r,j,v)
    ex=r+1;formula(s,r,7,f'=D{ex}/$B$16',weight,'pct');formula(s,r,8,f'=F{ex}/D{ex}',o0/t0,'pct');formula(s,r,9,f'=G{ex}/E{ex}',o1/t1,'pct');formula(s,r,10,f'=H{ex}*I{ex}',weight*o0/t0,'pct');formula(s,r,11,f'=H{ex}*J{ex}',weight*o1/t1,'pct')
stend=29+len(st)
formula(s,15,1,f'=SUM(D30:D{stend})',std['common_base']);formula(s,16,1,f'=SUM(E30:E{stend})',std['common_current'])
formula(s,17,1,"=B16/'月度趋势'!B14",std['common_base']/base,'pct');formula(s,18,1,"=B17/'月度趋势'!C14",std['common_current']/cur,'pct')
formula(s,19,1,f'=SUM(K30:K{stend})',std['base_common_rate'],'pct');formula(s,20,1,f'=SUM(L30:L{stend})',std['fixed_mix_current_rate'],'pct');formula(s,21,1,'=(B21-B20)*100',(std['fixed_mix_current_rate']-std['base_common_rate'])*100,'pp')
s.autofilter(28,0,stend-1,11);s.print_area(0,0,25,6)

# Quality facts, field completeness, provenance.
s=sheets['数据质量'];s.set_column('A:A',43);s.set_column('B:B',18);s.set_column('C:C',15);s.set_column('D:D',83)
s.write(3,0,'全量 7,525,498 条；缺失包括 NULL 和去除首尾空白后的空字符串，保留原始值。',f['note'])
checks=[('全量请求记录',int(quality.n),'全量分母；无原始行删除'),('唯一键数量',int(quality.unique_keys),'等于行数，未做去重'),('缺失唯一键',int(quality.missing_key),'无需补值'),('创建日期解析失败',int(quality.invalid_created),'创建时间可用于同窗筛选'),('关闭日期解析失败（非空）',int(quality.invalid_closed),'解析成功不代表业务有效'),('关闭日期缺失',240304,'不会影响创建请求量；不以已关闭样本替代全部请求'),('关闭早于创建',int(quality.negative_duration),'不进入有效关闭时长统计；请求量保留'),('关闭等于创建',int(quality.zero_duration),'未默认删除；可能含占位或行政动作，不能直接视为即时解决'),('Closed 但无关闭日期',int(quality.closed_missing),'状态与日期不等价，不做机构效率排名'),('非Closed 但有关闭日期',int(quality.nonclosed_with_date),'部分来自DOB；不能只凭关闭日期断言已完成'),('关闭晚于2026-09-07',int(quality.closed_after_end),'记录67996481为2026-12-14；快照日期未提供，单独标记'),('行政区 Unspecified',6331,'无人口分母，行政区记录量不代表人均问题率'),('渠道 UNKNOWN',590822,'保留为渠道类别，不填入ONLINE或PHONE')]
headers(s,4,['检查项','记录数/数量','占全量比例','对本项目的处理'])
for r,(label,n,note) in enumerate(checks,5):s.write(r,0,label);val(s,r,1,n);formula(s,r,2,f'=B{r+1}/$B$6',n/int(quality.n),'pct');s.write(r,3,note,f['wrap']);s.set_row(r,35)
s.write(20,0,'类别边界风险：DEP 的 Water System 在2026年8月降为0，同时维护类标签增多。',f['section'])
s.write(21,0,'这与分类切换相容，但缺少映射文档，不能把 Water Maintenance 的倍增直接解读为需求爆发。',f['note'])
s.write(22,0,'近月关闭观察期更短；本项目不发布平均结案时长排名或“处理效率提升”的结论。',f['note'])
headers(s,25,['原始字段','缺失数量','缺失比例','本项目中的含义 / 用途'])
missing=read('missing').iloc[0]
meaning={'unique_key':'请求标识；实测全量唯一，不代表独立现实事件','created_date':'创建时间；按源值解析，不自行换时区','closed_date':'关闭日期；与状态不完全一致','complaint_type':'原始请求类别；不擅自合并或重命名','descriptor':'类别下的问题描述；用于钻取坑洞、雪冰等','agency':'受理机构代码；不是人员或工时','agency_name':'机构名称；描述性属性','borough':'行政区标签；缺少人口分母','open_data_channel_type':'源系统渠道标签；分类定义未另行提供','status':'快照状态；不是状态事件历史','due_date':'截止时间字段几乎全缺，不做超时率','incident_address':'地址文本；未标准化，仅检验集中性','incident_zip':'邮编标识；保留字符串','latitude':'纬度文本；本项目不进行地理编码或空间推算','longitude':'经度文本；本项目不进行地理编码或空间推算'}
for r,(name,n) in enumerate(missing.items(),26):s.write(r,0,name);val(s,r,1,n);formula(s,r,2,f'=B{r+1}/$B$6',n/int(quality.n),'pct');s.write(r,3,meaning.get(name,'未纳入核心计算；缺失不统一补0或删行'),f['wrap']);s.set_row(r,29)
s.autofilter(25,0,69,3);s.print_area(0,0,69,3)

# Reader method/source section.
s=sheets['口径与复现'];s.set_column('A:A',27);s.set_column('B:B',111)
notes=[('数据来源',r'D:\项目1\原始数据 中25个Parquet；配套 D:\项目2\input\data_manifest.json。只读原件，未复制原始数据。'),('清单一致性','25个文件的字节数、SHA-256、行数与结构已核对；44个字符串字段，各分片结构一致，日期均处于文件名界限内。'),('实测时间范围','2024-09-07 00:00:00 至 2026-09-05 01:50:33；2024年9月及2026年9月不完整。日历日齐全仅证明有记录，不证明外部系统没有漏采。'),('记录粒度','一行一条 unique_key 请求记录。键唯一不证明请求对应不同事件、不同人员或不同地址，重复投诉仍可能存在。'),('比较窗口','2025-01-01 ≤ created < 2025-09-01，对比2026年相同月日，两窗各243天；不限制状态、类别、机构、行政区或渠道。'),('字段语义边界','未提供数据字典、分类变更说明或采集快照时点。字段解释依名称及实际值作最低限度操作定义，不当作经外部核实的业务规则。'),('计数和比率','记录数为COUNT(*)，核实unique_key全量唯一；增幅=(当期−基期)/基期，贡献率=分组净增/全部净增，零基期显示“无基期”。'),('预处理','日期用TRY_CAST解析；空值审计包含空白。类别与渠道保留原始标签，不更改大小写、不合并近似标签，不删除异常时长行。'),('标准化方法','仅共同存在的月份×类别×机构单元；每单元2025总数/共同单元2025总数为权重，分别加权两年ONLINE占比。明细及Excel公式见渠道变化。'),('尖峰测试','2026各类别自身最高5个自然日从当期剔除，基期保持完整，属于刻意更苛刻的单侧敏感性检查，不作等长同比指标。'),('事实与解释','增长、描述结构与渠道占比是数据内事实。天气、报送规则、真实问题数量和个体迁移均未被识别；仅可提出后续验证方向。'),('建议用途','按道路坑洞持续增量与雪冰集中日期分别审视日常和峰值需求；先核查渠道/分类规则，再研究报送变化。不据请求数直接配置精确工时或评定效率。'),('Excel阅读顺序','总览 → 月度趋势/同比拆解 → 峰值核查/渠道变化 → 数据质量。表格可筛选，渠道标准化明细位于第30行以下；图表为原生Excel对象。'),('源汇总刷新','Excel保存静态源汇总，公式计算增幅和占比；打开无需Python。换原始数据后需重跑脚本，修改单元格不会重新扫描Parquet，文字结论也不会自动改写。'),('复现命令','在本目录依次运行 python profile_data.py、python investigate.py、python deep_checks.py、python finalize_analysis.py、python build_workbook.py、python verify_workbook.py。'),('依赖','Python现有 duckdb / pyarrow / pandas / numpy / xlsxwriter；复核用 openpyxl / pywin32 / PyMuPDF 与本机 Microsoft Excel 16.0。不访问外部网页。'),('证据入口','evidence/source_verification.json、numerical_checks.json，及与各脚本同名逻辑生成的CSV；SQL在脚本及deep_queries.json/final_queries.json。'),('原始记录回查','evidence/peak_records.csv 与 anomaly_records.csv 保留请求键，可在原始Parquet定位；不输出完整地址明细。')]
headers(s,4,['项目','定义与说明'])
for r,(a,b) in enumerate(notes,5):s.write(r,0,a,f['section']);s.write(r,1,b,f['wrap']);s.set_row(r,48 if r in [5,7,9,13,17,19,20,21,22] else 40)
s.print_area(0,0,22,1)

# Front dashboard: only finished results link here.
s=sheets['总览'];s.set_column('A:A',28);s.set_column('B:F',17);s.set_column('G:L',12)
s.write(3,0,'全量覆盖 2024-09-07 至 2026-09-05；主要比较 2025 与 2026 年 1—8 月。单位：请求记录。',f['note'])
headers(s,5,['核心指标','2025年1—8月','2026年1—8月','增加记录','同比增幅'])
s.write(6,0,'请求记录量',f['section']);formula(s,6,1,"='月度趋势'!B14",base);formula(s,6,2,"='月度趋势'!C14",cur);formula(s,6,3,"='月度趋势'!D14",delta);formula(s,6,4,"='月度趋势'!E14",delta/base,'pct')
s.write(8,0,'1  增长有集中来源，也有广泛基础',f['section'])
s.write(9,0,'道路、雪冰、供暖合计贡献47.6%的净增量；排除这三类后，其余请求仍增长7.0%。',f['text'])
s.write(11,0,'2  雪冰尖峰与道路持续增长需要分别看待',f['section'])
s.write(12,0,'雪冰最高5天占59.4%；道路8个月中7个月上升，坑洞贡献道路净增量的74.0%。',f['text'])
s.write(14,0,'3  ONLINE 记录占比上升，固定构成后仍成立',f['section'])
actual0=std['base_online']/base;actual1=std['current_online']/cur
s.write(15,0,f'实际占比 {actual0:.1%} → {actual1:.1%}；固定月份、类别、机构构成后提高6.74个百分点。',f['text'])
chart=w.add_chart({'type':'line'})
for col_,name,color in [(1,'2025年','#8097B0'),(2,'2026年','#275A91')]:
    chart.add_series({'name':name,'categories':['月度趋势',5,0,12,0],'values':['月度趋势',5,col_,12,col_],'line':{'color':color,'width':2.25},'marker':{'type':'circle','size':4,'border':{'color':color},'fill':{'color':color}}})
chart.set_title({'name':'同月份请求量（条）','name_font':{'name':font,'size':12}});chart.set_x_axis({'name':'月份','num_font':{'name':font,'size':10}});chart.set_y_axis({'min':0,'num_format':'#,##0','num_font':{'name':font,'size':10},'major_gridlines':{'visible':True,'line':{'color':'#E7EBF0'}}});chart.set_legend({'position':'bottom','font':{'name':font,'size':10}});chart.set_chartarea({'border':{'none':True}});chart.set_size({'width':605,'height':320});s.insert_chart('A18',chart)
chart=w.add_chart({'type':'bar'});chart.add_series({'name':'增加记录','categories':['同比拆解',5,0,9,0],'values':['同比拆解',5,3,9,3],'fill':{'color':'#356EA0'},'border':{'none':True}});chart.set_title({'name':'主要类别净增量（条）','name_font':{'name':font,'size':12}});chart.set_x_axis({'min':0,'num_format':'#,##0','num_font':{'name':font,'size':10}});chart.set_y_axis({'reverse':True,'num_font':{'name':font,'size':10}});chart.set_legend({'none':True});chart.set_chartarea({'border':{'none':True}});chart.set_size({'width':610,'height':320});s.insert_chart('F18',chart)
s.write(33,0,'使用边界',f['section']);s.write(34,0,'请求记录不等于独立事件或居民数。数据无法证明天气原因、用户迁移或服务绩效变化。',f['text'])
s.write(35,0,'首尾不完整月份不参与同比；状态与关闭日期存在不一致，近期结案时长受观察期截断影响。',f['text'])
for c_,name in [(0,'月度趋势'),(2,'同比拆解'),(4,'峰值核查'),(6,'渠道变化'),(8,'数据质量'),(10,'口径与复现')]:s.write_url(38,c_,f"internal:'{name}'!A1",f['link'],name)
s.print_area(0,0,39,11);s.fit_to_pages(1,1);s.activate();s.set_first_sheet()
w.close()

# Human explanation stays synchronized with the verified aggregates.
text=f'''# NYC 311 服务请求分析说明

本项目全量读取25个Parquet，共 **{int(quality.n):,}条请求记录**、44个字段。25个SHA-256与配套清单一致；unique_key全量唯一且无缺失。实际创建时间为2024-09-07 00:00:00至2026-09-05 01:50:33，不能把末尾分片名称当作完整覆盖日期。

## 主要发现

1. **2026年1—8月请求量同比增加12.1%。** 从{base:,}增至{cur:,}条，净增{delta:,}条。两窗均为243天、全部8个月均增加。道路状况、雪冰、供暖/热水分别增加47,709、47,240、41,874条，合计占净增量47.6%。排除这三类后仍增长7.0%；只看4—8月仍增长9.8%。这是请求记录量变化，不等于真实问题发生率变化。
2. **道路增长更持续，雪冰增长更集中。** 道路状况8个月中7个月同比增加，坑洞描述增加35,310条，占道路净增量74.0%。2026雪冰记录最高5天共32,819条，占55,230条的59.4%；最大单日2026-02-24有11,370条，涉及{int(pc.address_n):,}个非空原始地址值。删除2026最高5天后，雪冰剩余记录仍比完整2025基期高180.5%，道路仍高81.0%。此单侧删除是压力测试，不是等长窗口同比。排除Snow Tracking描述后雪冰仍从7,723增至40,651条（+426.4%），因此该描述扩张不能单独解释增长。
3. **ONLINE占比提高，构成变化不能解释全部差异。** 实际占比从{actual0:.2%}升至{actual1:.2%}（+{(actual1-actual0)*100:.2f}个百分点）；ONLINE增加272,718条，占全部净增量94.9%，其他渠道存在正负抵消。固定“月份×原始类别×机构”的2025权重后，共同单元ONLINE占比从42.56%升至49.30%（+6.74个百分点）。共同单元覆盖两年记录的{std['common_base']/base:.2%}、{std['common_current']/cur:.2%}。这支持渠道记录结构发生变化，但无法确认个体用户迁移、接入方式或编码规则的作用。

## 口径与限制

- 主比较窗口：2025-01-01 ≤ 创建时间 < 2025-09-01，对比2026年相同月日。日历日均有记录；无法从给定材料验证对外部系统的采集完整性。时间按源字符串解析，不自行变换时区。
- 分母：增幅以对应基期记录数为分母；净增量贡献以全部净增287,342条为分母；份额以对应年份全部同窗记录为分母。机构、行政区、类别拆解各自闭合，不能跨维相加。全量状态不限，不删除异常关闭日期记录。
- 未提供业务数据字典、分类变更映射、人口、独立事件/个人标识、工时、天气或快照时点。字段含义依据名称与实际值作最低限度定义；地址未经标准化，地址数量不等于事件数；按行政区记录量无法推出人均发生率或绩效。
- 关闭日期缺失240,304条；负时长1,877条；零时长135,189条；Closed但无关闭日期15,543条；非Closed但有关闭日期25,388条。另有一条关闭日期为2026-12-14（键67996481）。这些问题不影响创建请求量，但阻止把简单关闭时长比较当作效率证据。近期请求观察期更短，本成品不做机构效率排名。
- DEP的Water System在2026年8月为0，同时Water Maintenance和Sewer Maintenance分别达8,925和5,060条。该现象与分类变化相容，但未经映射核实；因此未把新类别倍增作为核心需求结论，也未擅自合并类别。
- 所有观察为给定数据内的描述性事实。天气、真实事件增多及报送机制变化属于未验证解释。

## 如何使用 Excel

先看“总览”，再用“月度趋势”核实同窗和完整覆盖，“同比拆解”查看全部原始类别、机构、行政区的贡献。“峰值核查”提供描述拆解与尖峰压力测试；“渠道变化”提供实际份额和标准化公式，明细从第30行开始。“数据质量”包含44字段缺失率和异常处理，“口径与复现”列出来源与方法。原生图表可编辑，支持表格筛选和冻结标题。

建议按道路坑洞的持续增量与雪冰的集中日期分别审视日常与峰值需求；在解释ONLINE增长前先核实渠道规则。建议不等同于已识别因果，也不据此配置精确人力。

工作簿无需运行代码即可读取。源汇总是固定快照，Excel公式可重算占比、差额及标准化；改动汇总不会重新读取Parquet，文字发现不会自动更新。刷新数据需重跑以下脚本并复核叙述。

## 复现与核验

在本目录依次执行：

```powershell
python profile_data.py
python investigate.py
python deep_checks.py
python finalize_analysis.py
python build_workbook.py
python verify_workbook.py
```

原始输入始终只读。profile_data.py保存字段、缺失、键、日期和月份检查；investigate.py生成互斥维度拆解；deep_checks.py和finalize_analysis.py执行描述、日期、分类和标准化验证；build_workbook.py生成唯一最终工作簿与本文。现有Python依赖包含duckdb、pyarrow、pandas、numpy、xlsxwriter；验收使用openpyxl、pywin32、PyMuPDF及Microsoft Excel。Artifact Tool加载工具未提供且外部依赖目录超出本实验授权路径，因此采用现有XlsxWriter，并在本机Excel复算和导出检查。

关键证据位于evidence目录：source_verification.json（原件哈希/结构）、numerical_checks.json（总量/分母/分片核对）、compare_*.csv（同比）、descriptor_comparison.csv/peak_sensitivity.csv（钻取）、channel_standardization.csv（固定构成）、peak_records.csv/anomaly_records.csv（可回查键）。SQL直接保存在实际执行的Python脚本及deep_queries.json/final_queries.json。全部主维度拆解、月度合计、Parquet元数据行数与直接COUNT(*)精确对齐。
'''
(R/'分析说明.md').write_text(text,encoding='utf-8')
print('Created',R/'最终成果.xlsx')

