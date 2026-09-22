# -*- coding: utf-8 -*-
"""全量扫描 25 个 NYC 311 parquet 分片，产出中间汇总 CSV。
原始数据只读；所有输出写入 D:/项目2/GLM/中间汇总/。
逐文件读取、只取分析所需列，控制内存。
"""
import sys, io, os, glob
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
import numpy as np

RAW = r"D:\项目1\原始数据"
OUT = r"D:\项目2\GLM\中间汇总"
os.makedirs(OUT, exist_ok=True)

COLS = ['unique_key', 'created_date', 'closed_date', 'agency', 'complaint_type',
        'descriptor', 'borough', 'status', 'open_data_channel_type']

files = sorted(glob.glob(os.path.join(RAW, "*.parquet")))
assert len(files) == 25, len(files)

# 累积容器
frames = []          # 逐月轻量聚合结果
rows_total = 0
dup_keys = 0
created_min, created_max = None, None
closed_lt_created = 0
closed_after_pull = 0      # closed_date 晚于数据快照上限 2026-09-07 之后的记录数
closed_null = 0
created_null = 0
status_ct = {}
channel_ct = {}
agency_ct = pd.Series(dtype='int64')
complaint_ct = pd.Series(dtype='int64')
borough_ct = pd.Series(dtype='int64')
borough_complaint = {}     # (borough, complaint) -> count
agency_complaint = {}      # (agency, complaint) -> count
closure_by_agency = {}     # agency -> list of hours
closure_by_complaint = {}  # complaint -> list of hours
monthly_closure_median = {}  # month -> list of hours
monthly_complaint = {}     # (month, complaint) -> count
monthly_agency = {}        # (month, agency) -> count
monthly_borough = {}       # (month, borough) -> count
not_closed_by_agency = {}
not_closed_by_complaint = {}
descriptor_ct = pd.Series(dtype='int64')

PULL_END = pd.Timestamp('2026-09-08')

for i, p in enumerate(files):
    df = pd.read_parquet(p, columns=COLS)
    rows_total += len(df)

    # 唯一键查重（全量级）
    dup_keys += int(df['unique_key'].duplicated().sum())

    c = pd.to_datetime(df['created_date'], errors='coerce')
    cl = pd.to_datetime(df['closed_date'], errors='coerce')
    created_null += int(c.isna().sum())
    closed_null += int(cl.isna().sum())
    if created_min is None or c.min() < created_min: created_min = c.min()
    if created_max is None or c.max() > created_max: created_max = c.max()

    valid_pair = c.notna() & cl.notna()
    closed_lt_created += int((cl < c).sum())
    closed_after_pull += int((cl > PULL_END).sum())

    month = c.dt.to_period('M').astype(str)
    df['_month'] = month

    # 计数聚合
    agency_ct = agency_ct.add(df['agency'].value_counts(), fill_value=0)
    complaint_ct = complaint_ct.add(df['complaint_type'].value_counts(), fill_value=0)
    borough_ct = borough_ct.add(df['borough'].value_counts(), fill_value=0)
    for k, v in df['status'].value_counts().items(): status_ct[k] = status_ct.get(k, 0) + int(v)
    for k, v in df['open_data_channel_type'].value_counts().items(): channel_ct[k] = channel_ct.get(k, 0) + int(v)
    descriptor_ct = descriptor_ct.add(df['descriptor'].value_counts(), fill_value=0)

    bc = df.groupby(['borough', 'complaint_type']).size()
    for (b, cp), v in bc.items(): borough_complaint[(b, cp)] = borough_complaint.get((b, cp), 0) + int(v)
    ac = df.groupby(['agency', 'complaint_type']).size()
    for (a, cp), v in ac.items(): agency_complaint[(a, cp)] = agency_complaint.get((a, cp), 0) + int(v)

    mc = df.groupby(['_month', 'complaint_type']).size()
    for (m, cp), v in mc.items(): monthly_complaint[(m, cp)] = monthly_complaint.get((m, cp), 0) + int(v)
    ma = df.groupby(['_month', 'agency']).size()
    for (m, a), v in ma.items(): monthly_agency[(m, a)] = monthly_agency.get((m, a), 0) + int(v)
    mb = df.groupby(['_month', 'borough']).size()
    for (m, b), v in mb.items(): monthly_borough[(m, b)] = monthly_borough.get((m, b), 0) + int(v)

    # 处理时长（小时）：仅正常关闭（closed>=created 且 <= 快照上限）
    hrs = ((cl - c).dt.total_seconds() / 3600)
    ok = valid_pair & (cl >= c) & (cl <= PULL_END)
    for a, v in hrs[ok].groupby(df['agency'][ok]).agg(list).items():
        closure_by_agency.setdefault(a, []).extend(v)
    for cp, v in hrs[ok].groupby(df['complaint_type'][ok]).agg(list).items():
        closure_by_complaint.setdefault(cp, []).extend(v)
    mcl = hrs[ok].groupby(month[ok])
    for m, v in mcl.agg(list).items():
        monthly_closure_median.setdefault(m, []).extend(v)

    # 快照时仍未关闭（status 非 Closed）的请求
    nc = df[df['status'] != 'Closed']
    for a, v in nc['agency'].value_counts().items():
        not_closed_by_agency[a] = not_closed_by_agency.get(a, 0) + int(v)
    for cp, v in nc['complaint_type'].value_counts().items():
        not_closed_by_complaint[cp] = not_closed_by_complaint.get(cp, 0) + int(v)

    print(f"[{i+1}/25] {os.path.basename(p)} rows={len(df)}", flush=True)

# ---- 输出 ----
pd.DataFrame([{
    '指标': ['总记录数', 'unique_key重复数', 'created_date缺失', 'closed_date缺失',
             'closed早于created', 'closed晚于快照上限(2026-09-08)',
             'created最早', 'created最晚'],
    '值': [rows_total, dup_keys, created_null, closed_null,
           closed_lt_created, closed_after_pull,
           str(created_min), str(created_max)]
}]).to_csv(os.path.join(OUT, '数据质量总览.csv'), index=False, encoding='utf-8-sig')

agency_ct.sort_values(ascending=False).to_csv(os.path.join(OUT, '机构总量.csv'), encoding='utf-8-sig', header=['记录数'])
complaint_ct.sort_values(ascending=False).to_csv(os.path.join(OUT, '投诉类型总量.csv'), encoding='utf-8-sig', header=['记录数'])
borough_ct.sort_values(ascending=False).to_csv(os.path.join(OUT, '区域总量.csv'), encoding='utf-8-sig', header=['记录数'])
descriptor_ct.sort_values(ascending=False).head(40).to_csv(os.path.join(OUT, '描述词Top40.csv'), encoding='utf-8-sig', header=['记录数'])
pd.Series(status_ct).sort_values(ascending=False).to_csv(os.path.join(OUT, '状态总量.csv'), encoding='utf-8-sig', header=['记录数'])
pd.Series(channel_ct).sort_values(ascending=False).to_csv(os.path.join(OUT, '渠道总量.csv'), encoding='utf-8-sig', header=['记录数'])

pd.DataFrame([{'机构': a, '投诉类型': cp, '记录数': v}
              for (a, cp), v in agency_complaint.items()]).to_csv(
    os.path.join(OUT, '机构x投诉类型.csv'), index=False, encoding='utf-8-sig')
pd.DataFrame([{'区域': b, '投诉类型': cp, '记录数': v}
              for (b, cp), v in borough_complaint.items()]).to_csv(
    os.path.join(OUT, '区域x投诉类型.csv'), index=False, encoding='utf-8-sig')
pd.DataFrame([{'月份': m, '投诉类型': cp, '记录数': v}
              for (m, cp), v in monthly_complaint.items()]).to_csv(
    os.path.join(OUT, '月度x投诉类型.csv'), index=False, encoding='utf-8-sig')
pd.DataFrame([{'月份': m, '机构': a, '记录数': v}
              for (m, a), v in monthly_agency.items()]).to_csv(
    os.path.join(OUT, '月度x机构.csv'), index=False, encoding='utf-8-sig')
pd.DataFrame([{'月份': m, '区域': b, '记录数': v}
              for (m, b), v in monthly_borough.items()]).to_csv(
    os.path.join(OUT, '月度x区域.csv'), index=False, encoding='utf-8-sig')

# 处理时长统计
def stat_table(d, keyname):
    recs = []
    for k, lst in d.items():
        arr = np.asarray(lst, dtype='float64')
        recs.append({keyname: k, '关闭请求数': len(arr),
                     '中位小时': np.median(arr), '平均小时': arr.mean(),
                     'P25小时': np.percentile(arr, 25), 'P75小时': np.percentile(arr, 75),
                     'P90小时': np.percentile(arr, 90)})
    t = pd.DataFrame(recs).sort_values('关闭请求数', ascending=False)
    return t

stat_table(closure_by_agency, '机构').to_csv(os.path.join(OUT, '机构处理时长.csv'), index=False, encoding='utf-8-sig')
stat_table(closure_by_complaint, '投诉类型').to_csv(os.path.join(OUT, '投诉类型处理时长.csv'), index=False, encoding='utf-8-sig')
stat_table(monthly_closure_median, '月份').to_csv(os.path.join(OUT, '月度处理时长.csv'), index=False, encoding='utf-8-sig')

pd.DataFrame([{'机构': k, '未关闭数': v} for k, v in
              sorted(not_closed_by_agency.items(), key=lambda x: -x[1])]).to_csv(
    os.path.join(OUT, '未关闭_按机构.csv'), index=False, encoding='utf-8-sig')
pd.DataFrame([{'投诉类型': k, '未关闭数': v} for k, v in
              sorted(not_closed_by_complaint.items(), key=lambda x: -x[1])]).to_csv(
    os.path.join(OUT, '未关闭_按投诉类型.csv'), index=False, encoding='utf-8-sig')

print("DONE rows_total=", rows_total)
