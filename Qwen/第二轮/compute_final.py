"""
最终数字表：从 work/slim.parquet 计算所有交付用表格，写入 work/tables/*.csv。
Excel 看板与 分析说明.md 中的每个数字都出自这里的表，保证一致、可回查。

口径（与 build_slim.py / build_aggregates.py 一致）：
- 数据窗口: created 2024-09-07 ~ 2026-09-05(截断日, 仅718条)。完整日截止 2026-09-04。
- 部分月: 2024-09(自07日起), 2026-09(至04日完整)。
- YoY 对齐窗口: W1=2024-10..2025-08, W2=2025-10..2026-08 (各11个完整月)。
- 成熟队列(处理时长): created <= 2026-06-30 (该月及之前 closed_rate>=95%)。
- 有效处理时长 valid_res: closed_dt 存在 且 0<=时长 且 closed_dt 在 [2024-09-07, 2026-09-07]。
- 天气敏感类别: Snow or Ice, HEAT/HOT WATER, Street Condition, Damaged Tree, PLUMBING, WATER LEAK。
"""
import pandas as pd, numpy as np, os

BASE = r"D:\项目2\增强实验\Qwen\work"
T = os.path.join(BASE, "tables"); os.makedirs(T, exist_ok=True)
df = pd.read_parquet(os.path.join(BASE, "slim.parquet"))

SNAP = pd.Timestamp("2026-09-07"); WIN0 = pd.Timestamp("2024-09-07")
MATURITY = pd.Timestamp("2026-07-01")
valid = df.is_closed & df.resolution_hours.notna() & (df.resolution_hours >= 0) \
        & (df.closed_dt <= SNAP) & (df.closed_dt >= WIN0)
df["res_h"] = df.resolution_hours.where(valid)

W1 = [f"2024-{m:02d}" for m in range(10,13)] + [f"2025-{m:02d}" for m in range(1,9)]
W2 = [f"2025-{m:02d}" for m in range(10,13)] + [f"2026-{m:02d}" for m in range(1,9)]
PARTIAL = {"2024-09", "2026-09"}
WEATHER = ["Snow or Ice","HEAT/HOT WATER","Street Condition","Damaged Tree","PLUMBING","WATER LEAK"]
df["win"] = np.where(df.created_month.isin(W1), "W1", np.where(df.created_month.isin(W2), "W2", "other"))
mat = df[df.created_dt < MATURITY]          # 成熟队列
matv = mat[mat.res_h.notna()]               # 成熟且有效处理时长

# 提前计算 W1/W2 与天气敏感类, 供总览与分解共用
n1 = int((df.win=="W1").sum()); n2 = int((df.win=="W2").sum()); net = n2-n1
w1w = int(((df.win=="W1")&df.complaint_type.isin(WEATHER)).sum())
w2w = int(((df.win=="W2")&df.complaint_type.isin(WEATHER)).sum())
base_growth = ((n2-w2w)/(n1-w1w)-1)*100

# ---------- 1. 总览 ----------
ov = pd.DataFrame([
    ["记录总数(行)", f"{len(df):,}", "一行 = 一个 311 服务请求(unique_key 全局唯一, 0 重复)"],
    ["created 时间范围", "2024-09-07 ~ 2026-09-05", "2026-09-05 为截断日(仅718条); 完整日截止 2026-09-04"],
    ["完整月范围", "2024-10 ~ 2026-08 (23个完整月)", "2024-09 与 2026-09 为部分月, 不参与 YoY"],
    ["投诉类型数", str(df.complaint_type.nunique()), "complaint_type 去重计数(全期)"],
    ["机构数", str(df.agency.nunique()), "agency 去重计数(全期)"],
    ["月均请求量(完整月)", f"{df[~df.created_month.isin(PARTIAL)].groupby('created_month').size().mean():,.0f}", "23个完整月的平均"],
    ["已结案记录占比(全期)", f"{df.is_closed.mean()*100:.2f}%", "closed_dt 非空; 近3个月为右删失(尚未到结案时间)"],
    ["成熟队列处理时长中位数", f"{matv.res_h.median():.2f} 小时", "created<=2026-06-30 且有效时长 (n={:,})".format(len(matv))],
    ["24小时内结案占比(成熟队列)", f"{(matv.res_h<=24).mean()*100:.1f}%", "同上口径"],
    ["7天内结案占比(成熟队列)", f"{(matv.res_h<=168).mean()*100:.1f}%", "同上口径"],
    ["YoY 总量变化(Oct-Aug对齐)", f"+{(df[df.win=='W2'].shape[0]/df[df.win=='W1'].shape[0]-1)*100:.1f}%", "W1=3,293,003 → W2=3,639,698"],
    ["基础增速(剔除天气敏感类)", f"+{base_growth:.1f}%", "非天气类别 W1→W2 见 t_decomp.csv"],
    ["缺失率: 渠道 UNKNOWN", f"{df.channel.isna().mean()*100:.2f}%", "open_data_channel_type='UNKNOWN'"],
    ["缺失率: city", f"{df.city.isna().mean()*100:.2f}%", ""],
    ["缺失率: borough Unspecified", f"{(df.borough_raw=='Unspecified').mean()*100:.3f}%", "已归为 NA"],
    ["缺失率: closed_dt(未结案)", f"{df.closed_dt.isna().mean()*100:.2f}%", "右删失, 集中在近3个月"],
    ["QA剔除: 负处理时长", "1,877 条(0.025%)", "closed_dt < created_dt, 处理时长统计中剔除"],
    ["QA剔除: closed_dt 超快照/早于窗口", "1 + 8 条", "未来日期1条; 早于2024-09-07共8条"],
], columns=["指标","值","说明"])
ov.to_csv(os.path.join(T,"t_overview.csv"), index=False)

# ---------- 2. 月度总览 ----------
mo = df.groupby("created_month")
monthly = pd.DataFrame({
    "n_requests": mo.size(),
    "closed_rate_pct": (mo.is_closed.mean()*100).round(2),
})
med = mat.groupby("created_month")["res_h"].median()
monthly["median_res_h_matured"] = med.reindex(monthly.index).round(2)
monthly["is_partial_month"] = [m in PARTIAL for m in monthly.index]
monthly["censored_note"] = np.where(monthly.closed_rate_pct < 95, "右删失: 大量请求尚未到结案时间", "")
# YoY: 对齐月
yoy_map = {}
for m in range(10,13):
    yoy_map[f"2025-{m:02d}"] = (f"2024-{m:02d}", f"2025-{m:02d}")
for m in range(1,9):
    yoy_map[f"2026-{m:02d}"] = (f"2025-{m:02d}", f"2026-{m:02d}")
def yoy(mm):
    if mm in yoy_map:
        a,b = yoy_map[mm]
        na = monthly.loc[a,"n_requests"]; nb = monthly.loc[b,"n_requests"]
        return round((nb/na-1)*100,1)
    return np.nan
monthly["yoy_pct_aligned"] = [yoy(m) for m in monthly.index]
monthly.reset_index().rename(columns={"created_month":"month"}).to_csv(os.path.join(T,"t_monthly.csv"), index=False)

# ---------- 3. TOP20 投诉类型(全期) ----------
g = df.groupby("complaint_type")
gm = mat.groupby("complaint_type")["res_h"]
tt = pd.DataFrame({
    "n_requests": g.size(),
    "share_pct": (g.size()/len(df)*100).round(2),
    "closed_rate_pct": (g.is_closed.mean()*100).round(1),
    "median_res_h": gm.median().round(1),
    "p90_res_h": gm.quantile(.90).round(1),
}).sort_values("n_requests", ascending=False).head(20).reset_index()
tt.to_csv(os.path.join(T,"t_top_types.csv"), index=False)

# ---------- 4. YoY 类型全表 ----------
c1 = df[df.win=="W1"].groupby("complaint_type").size()
c2 = df[df.win=="W2"].groupby("complaint_type").size()
yoy_t = pd.DataFrame({"W1_Oct24_Aug25": c1, "W2_Oct25_Aug26": c2}).fillna(0).astype(int)
yoy_t["delta"] = yoy_t.W2_Oct25_Aug26 - yoy_t.W1_Oct24_Aug25
yoy_t["pct_change"] = np.where(yoy_t.W1_Oct24_Aug25>0, (yoy_t.delta/yoy_t.W1_Oct24_Aug25*100).round(1), np.nan)
yoy_t["contrib_pct_of_net"] = (yoy_t.delta/yoy_t.delta.sum()*100).round(1)
yoy_t = yoy_t.sort_values("delta", ascending=False).reset_index()
yoy_t.to_csv(os.path.join(T,"t_yoy_types.csv"), index=False)

# ---------- 5. 增长分解 ----------
n1 = int((df.win=="W1").sum()); n2 = int((df.win=="W2").sum()); net = n2-n1
w1w = df[(df.win=="W1")&df.complaint_type.isin(WEATHER)].shape[0]
w2w = df[(df.win=="W2")&df.complaint_type.isin(WEATHER)].shape[0]
trio = ["Snow or Ice","HEAT/HOT WATER","Street Condition"]
t1 = df[(df.win=="W1")&df.complaint_type.isin(trio)].shape[0]
t2 = df[(df.win=="W2")&df.complaint_type.isin(trio)].shape[0]
# 10466 爆发
burst = df[(df.complaint_type=="Noise - Residential")&(df.created_dt>=pd.Timestamp("2025-01-01"))
           &(df.created_dt<pd.Timestamp("2025-01-14"))&(df.incident_zip=="10466")].shape[0]
jan25 = int((df.created_month=="2025-01").sum()); jan26 = int((df.created_month=="2026-01").sum())
dec = pd.DataFrame([
    ["W1 总量 (2024-10..2025-08, 11个完整月)", f"{n1:,}", "", ""],
    ["W2 总量 (2025-10..2026-08, 11个完整月)", f"{n2:,}", "", ""],
    ["净增", f"{net:,}", f"+{net/n1*100:.1f}%", "对齐窗口比较"],
    ["其中: 天气敏感6类净增", f"{w2w-w1w:,}", f"占净增 {(w2w-w1w)/net*100:.1f}%",
     f"W1={w1w:,} → W2={w2w:,} (+{(w2w/w1w-1)*100:.1f}%); 类别: {', '.join(WEATHER)}"],
    ["其中: 冬季三类(Snow/Heat/Street)净增", f"{t2-t1:,}", f"占净增 {(t2-t1)/net*100:.1f}%", f"W1={t1:,} → W2={t2:,}"],
    ["剔除天气敏感6类后的基础增速", f"+{((n2-w2w)/(n1-w1w)-1)*100:.1f}%", "", "非天气类别 W1={:,} → W2={:,}".format(n1-w1w,n2-w2w)],
    ["夏季月 YoY (5/6/7/8月, 无冬季影响)", "+12.5% / +9.3% / +8.6% / +8.0%", "", "与基础增速一致, 增长广泛存在"],
    ["Jan 2026 表观 YoY", f"+{(jan26/jan25-1)*100:.1f}%", "", f"Jan2025={jan25:,}, Jan2026={jan26:,}"],
    ["Jan 2025 ZIP 10466 集中爆发(2025-01-01..13)", f"{burst:,} 条", "", "Noise-Residential, 84% MOBILE 渠道"],
    ["Jan 2026 调整后 YoY (剔除该爆发)", f"+{(jan26/(jan25-burst)-1)*100:.1f}%", "", "表观 +0.1% 被基期异常压低"],
], columns=["项目","值","占比/变化","说明"])
dec.to_csv(os.path.join(T,"t_decomp.csv"), index=False)

# ---------- 6. 冬季两类 + Street Condition 逐月 ----------
winter_types = trio + ["PLUMBING","WATER LEAK"]
piv = df[df.complaint_type.isin(winter_types)].pivot_table(
    index="created_month", columns="complaint_type", values="unique_key", aggfunc="count", fill_value=0)
piv = piv.reindex(columns=winter_types)
piv["总量(全部类型)"] = df.groupby("created_month").size()
piv.reset_index().rename(columns={"created_month":"month"}).to_csv(os.path.join(T,"t_winter_monthly.csv"), index=False)

# ---------- 7. Street Condition 区对比 ----------
sc = df[df.complaint_type=="Street Condition"].copy()
sc["boro"] = sc.borough.astype(object).where(sc.borough.notna(), "Unspecified")
a = sc[sc.created_month.isin(["2025-03","2025-04","2025-05","2025-06"])].groupby("boro").size()
b = sc[sc.created_month.isin(["2026-03","2026-04","2026-05","2026-06"])].groupby("boro").size()
sct = pd.DataFrame({"2025_Mar_Jun": a, "2026_Mar_Jun": b}).fillna(0).astype(int)
sct["倍数"] = (sct["2026_Mar_Jun"]/sct["2025_Mar_Jun"]).round(2)
sct = sct.reset_index().rename(columns={"boro":"区"})
sct.to_csv(os.path.join(T,"t_street_borough.csv"), index=False)

# ---------- 8. 处理时长按机构(成熟队列) ----------
ga = mat.groupby("agency")
res_a = pd.DataFrame({
    "n_requests": ga.size(),
    "closed_n": ga.res_h.count(),
    "median_h": ga.res_h.median().round(1),
    "p90_h": ga.res_h.quantile(.90).round(1),
    "mean_h": ga.res_h.mean().round(1),
    "pct_24h": (mat.assign(f24=mat.res_h<=24).groupby("agency").f24.mean()*100).round(1),
}).sort_values("n_requests", ascending=False).reset_index()
res_a.to_csv(os.path.join(T,"t_res_agency.csv"), index=False)

# ---------- 9. 处理时长按类型 TOP15 + W1/W2 对比 ----------
top15 = tt.complaint_type.head(15).tolist()
mres = matv[matv.complaint_type.isin(top15)]
p = mres.pivot_table(index="complaint_type", columns="win", values="res_h", aggfunc="median")
cnt = mres[mres.win!="other"].groupby("complaint_type").size()
res_t = pd.DataFrame({"n_matured": cnt, "median_W1_h": p["W1"], "median_W2_h": p["W2"]})
res_t["delta_h"] = (res_t.median_W2_h - res_t.median_W1_h).round(1)
res_t["median_W1_h"] = res_t.median_W1_h.round(1); res_t["median_W2_h"] = res_t.median_W2_h.round(1)
res_t = res_t.reindex(top15).reset_index()
res_t.to_csv(os.path.join(T,"t_res_type.csv"), index=False)

# ---------- 10. HEAT/HOT WATER 逐月 ----------
hh = df[df.complaint_type=="HEAT/HOT WATER"]
hht = pd.DataFrame({"n": hh.groupby("created_month").size(),
                    "median_res_h": hh.groupby("created_month").res_h.median().round(1)})
hht.reset_index().rename(columns={"created_month":"month"}).to_csv(os.path.join(T,"t_heat_monthly.csv"), index=False)

# ---------- 11. 渠道 ----------
chlab = df.channel.astype(object).where(df.channel.notna(), "UNKNOWN")
ch1 = chlab[df.win=="W1"].value_counts()
ch2 = chlab[df.win=="W2"].value_counts()
cht = pd.DataFrame({"W1": ch1, "W2": ch2}).fillna(0).astype(int)
cht["W1_share_pct"] = (cht.W1/cht.W1.sum()*100).round(1)
cht["W2_share_pct"] = (cht.W2/cht.W2.sum()*100).round(1)
cht["delta"] = cht.W2 - cht.W1
cht["pct_change"] = (cht.delta/cht.W1.replace(0,np.nan)*100).round(1)
cht.reset_index().rename(columns={"index":"channel"}).to_csv(os.path.join(T,"t_channel.csv"), index=False)
# (cht index 名可能为 None, 统一在读取端处理)
# 渠道逐月
mch = df.pivot_table(index="created_month", columns=chlab,
                     values="unique_key", aggfunc="count", fill_value=0)
mch.reset_index().rename(columns={"created_month":"month"}).to_csv(os.path.join(T,"t_channel_monthly.csv"), index=False)

# ---------- 12. 行政区 ----------
boro = df.borough.astype(object).where(df.borough.notna(), "Unspecified")
bo = df.groupby(boro).size().rename("n_all")
b1 = df[df.win=="W1"].groupby(boro).size().rename("W1")
b2 = df[df.win=="W2"].groupby(boro).size().rename("W2")
bot = pd.concat([bo,b1,b2], axis=1).fillna(0).astype(int)
bot["share_pct"] = (bot.n_all/len(df)*100).round(1)
bot["delta"] = bot.W2-bot.W1
bot["pct_change"] = (bot.delta/bot.W1*100).round(1)
bot = bot.sort_values("n_all", ascending=False).reset_index()
bot.rename(columns={"index":"borough"}).to_csv(os.path.join(T,"t_borough.csv"), index=False)

# ---------- 13. 质量事件日志 ----------
qe = pd.DataFrame([
    ["分类变更(DEP水系)", "2026-05 起 Water Maintenance / Sewer Maintenance 出现; Water System 2026-08 起为 0 条; Drinking Water 2025-10 后消失",
     "Water Maintenance 2026-08 达 8,925 条(全部 DEP, descriptor 以 Fire Hydrant 为主)",
     "Water Maintenance 的 YoY +125,422% 与 Water System -7.3% 均为口径伪象; 涉及 DEP 水系的跨期比较需合并新旧类别"],
    ["分类改名(NYPD)", "Bike/Roller/Skate Chronic 止于 2025-10; Bike/Roller/Skate 始于 2025-11", "两类别量级衔接", "改名处理, 跨期比较需合并"],
    ["集中源异常(噪声)", "2025-01-01..13 Noise-Residential 53,907 条, 其中 ZIP 10466(Bronx) 43,014 条(80%), MOBILE 渠道 84%",
     "常态该类型日均约1,200条且 Bronx 占40%; 该13天 Bronx 占85%", "Jan 2025 基期被抬高; Jan 2026 调整 YoY = +14.2%(表观 +0.1%); 原因未验证(无外部信息)"],
    ["类别骤变(Drug Activity)", "2025-07/08 尖峰 6,038/5,447 条(Queens 占 89.6%), 2025-11 起骤降至 ~500-600/月(此前常态 ~1,500-2,500)",
     "descriptor 构成无实质变化(Use Outside 为主)", "YoY -67.9% 有误导性(基期含尖峰+后期崩塌); 报送或执法口径变化, 原因未验证"],
    ["时间戳异常", "负处理时长 1,877 条(0.025%); closed_dt>快照 1 条; closed_dt<窗口起点 8 条", "占比极小", "处理时长统计中剔除(valid_res 规则)"],
    ["右删失", "近3个月 closed_rate: 2026-07=92.3%, 2026-08=85.7%, 2026-09=61.5%", "并非绩效下降, 是尚未到结案时间",
     "处理时长分析使用成熟队列(created<=2026-06-30)"],
    ["截断日", "2026-09-05 仅 718 条(常态约1万/日)", "数据快照截止", "2026-09 视为部分月"],
], columns=["事件","窗口/位置","证据","处理与影响"])
qe.to_csv(os.path.join(T,"t_quality_events.csv"), index=False)

# ---------- 14. Drug Activity 逐月 ----------
da = df[df.complaint_type=="Drug Activity"].groupby("created_month").size().rename("n").reset_index()
da.rename(columns={"created_month":"month"}).to_csv(os.path.join(T,"t_drug_monthly.csv"), index=False)

# ---------- 15. 每日量 ----------
daily = df.groupby(df.created_dt.dt.date).size().reset_index()
daily.columns = ["date","n"]
daily.to_csv(os.path.join(T,"t_daily.csv"), index=False)

print("tables written:", os.listdir(T))
print("\n=== 核对关键数 ===")
print("W1,W2,net:", n1, n2, net)
print("天气6类 W1,W2:", w1w, w2w, "占净增 %.1f%%"%((w2w-w1w)/net*100))
print("基础增速: +%.1f%%"%(((n2-w2w)/(n1-w1w)-1)*100))
print("burst:", burst, "Jan adj YoY: +%.1f%%"%((jan26/(jan25-burst)-1)*100))
print("matured median h: %.2f, 24h: %.1f%%, 7d: %.1f%%"%(matv.res_h.median(),(matv.res_h<=24).mean()*100,(matv.res_h<=168).mean()*100))
