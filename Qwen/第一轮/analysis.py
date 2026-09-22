# -*- coding: utf-8 -*-
"""
NYC 311 数据分析主管线（Qwen）
只读 D:\项目1\原始数据 下的 Parquet 分片，输出汇总 CSV 到 汇总/ 目录。
运行: python analysis.py
"""
import duckdb, os, json

RAW = r"D:\项目1\原始数据\created_date=*.parquet"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "汇总")
os.makedirs(OUT, exist_ok=True)
con = duckdb.connect()

def save(name, df):
    path = os.path.join(OUT, name)
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"saved {name}: {len(df)} rows")

def q(s):
    return con.sql(s).df()

# ---------- 1. 数据概况 ----------
overview = q(f"""
SELECT count(*) AS 总记录数,
 count(DISTINCT unique_key) AS 唯一键数,
 min(CAST(created_date AS TIMESTAMP)) AS 最早创建时间,
 max(CAST(created_date AS TIMESTAMP)) AS 最晚创建时间,
 count(DISTINCT complaint_type) AS 投诉类型数,
 count(DISTINCT borough) AS 行政区数
FROM read_parquet('{RAW}')""")
save("概况_总量.csv", overview)

miss = q(f"""
SELECT 'closed_date(关闭时间)' AS 字段, round(100.0*sum(CASE WHEN closed_date IS NULL OR closed_date='' THEN 1 ELSE 0 END)/count(*),2) AS 缺失率pct FROM read_parquet('{RAW}')
UNION ALL SELECT 'descriptor(细分描述)', round(100.0*sum(CASE WHEN descriptor IS NULL OR descriptor='' THEN 1 ELSE 0 END)/count(*),2) FROM read_parquet('{RAW}')
UNION ALL SELECT 'location_type(位置类型)', round(100.0*sum(CASE WHEN location_type IS NULL OR location_type='' THEN 1 ELSE 0 END)/count(*),2) FROM read_parquet('{RAW}')
UNION ALL SELECT 'incident_zip(邮编)', round(100.0*sum(CASE WHEN incident_zip IS NULL OR incident_zip='' THEN 1 ELSE 0 END)/count(*),2) FROM read_parquet('{RAW}')
UNION ALL SELECT 'city(城市)', round(100.0*sum(CASE WHEN city IS NULL OR city='' THEN 1 ELSE 0 END)/count(*),2) FROM read_parquet('{RAW}')
UNION ALL SELECT 'due_date(截止时间)', round(100.0*sum(CASE WHEN due_date IS NULL OR due_date='' THEN 1 ELSE 0 END)/count(*),2) FROM read_parquet('{RAW}')
UNION ALL SELECT 'latitude(纬度)', round(100.0*sum(CASE WHEN latitude IS NULL OR latitude='' THEN 1 ELSE 0 END)/count(*),2) FROM read_parquet('{RAW}')""")
save("概况_缺失率.csv", miss)

save("概况_状态.csv", q(f"SELECT status AS 状态, count(*) AS 记录数, round(100.0*count(*)/7525498,2) AS 占比pct FROM read_parquet('{RAW}') GROUP BY 1 ORDER BY 2 DESC"))
save("概况_渠道.csv", q(f"SELECT open_data_channel_type AS 渠道, count(*) AS 记录数, round(100.0*count(*)/7525498,2) AS 占比pct FROM read_parquet('{RAW}') GROUP BY 1 ORDER BY 2 DESC"))
save("概况_机构.csv", q(f"SELECT agency AS 机构代码, any_value(agency_name) AS 机构名称, count(*) AS 记录数, round(100.0*count(*)/7525498,2) AS 占比pct FROM read_parquet('{RAW}') GROUP BY 1 ORDER BY 3 DESC LIMIT 12"))

# ---------- 2. 月度趋势 ----------
save("月度_总量.csv", q(f"""
SELECT strftime(CAST(created_date AS TIMESTAMP),'%Y-%m') AS 月份, count(*) AS 记录数
FROM read_parquet('{RAW}') GROUP BY 1 ORDER BY 1"""))

KEY_TYPES = ['HEAT/HOT WATER','Snow or Ice','Street Condition','Illegal Parking','Noise - Residential','Drug Activity','Noise - Helicopter']
tl = ", ".join([f"sum(CASE WHEN complaint_type='{t}' THEN 1 ELSE 0 END) AS \"{t}\"" for t in KEY_TYPES])
save("月度_关键类型.csv", q(f"""
SELECT strftime(CAST(created_date AS TIMESTAMP),'%Y-%m') AS 月份, {tl}
FROM read_parquet('{RAW}') GROUP BY 1 ORDER BY 1"""))

# 坑洞细分
save("月度_坑洞.csv", q(f"""
SELECT strftime(CAST(created_date AS TIMESTAMP),'%Y-%m') AS 月份, count(*) AS 坑洞投诉数
FROM read_parquet('{RAW}') WHERE complaint_type='Street Condition' AND descriptor='Pothole' GROUP BY 1 ORDER BY 1"""))

# ---------- 3. 投诉类型 TOP / 行政区 ----------
save("类型_TOP20.csv", q(f"""
SELECT complaint_type AS 投诉类型, any_value(agency) AS 主要机构, count(*) AS 记录数,
 round(100.0*count(*)/7525498,2) AS 占比pct
FROM read_parquet('{RAW}') GROUP BY 1 ORDER BY 3 DESC LIMIT 20"""))

save("行政区_总量.csv", q(f"""
SELECT borough AS 行政区, count(*) AS 记录数, round(100.0*count(*)/7525498,2) AS 占比pct
FROM read_parquet('{RAW}') GROUP BY 1 ORDER BY 2 DESC"""))

# 行政区 x 前6类型
save("行政区_类型交叉.csv", q(f"""
SELECT borough AS 行政区, complaint_type AS 投诉类型, count(*) AS 记录数
FROM read_parquet('{RAW}')
WHERE complaint_type IN ('Illegal Parking','Noise - Residential','HEAT/HOT WATER','Blocked Driveway','UNSANITARY CONDITION','Street Condition')
GROUP BY 1,2 ORDER BY 1,3 DESC"""))

# 每区记录占比 / 每千人口口径不可得，只做记录数对比（在说明中注明）

# ---------- 4. 同比（口径：两年各自完整的 10月–8月 共11个月窗口） ----------
save("同比_类型变化.csv", q(f"""
WITH d AS (
 SELECT complaint_type,
  sum(CASE WHEN CAST(created_date AS TIMESTAMP) < TIMESTAMP '2025-09-01' THEN 1 ELSE 0 END) y1,
  sum(CASE WHEN CAST(created_date AS TIMESTAMP) >= TIMESTAMP '2025-09-01' THEN 1 ELSE 0 END) y2
 FROM read_parquet('{RAW}')
 WHERE (CAST(created_date AS TIMESTAMP) >= TIMESTAMP '2024-10-01' AND CAST(created_date AS TIMESTAMP) < TIMESTAMP '2025-09-01')
    OR (CAST(created_date AS TIMESTAMP) >= TIMESTAMP '2025-10-01' AND CAST(created_date AS TIMESTAMP) < TIMESTAMP '2026-09-01')
 GROUP BY 1)
SELECT complaint_type AS 投诉类型, y1 AS 前一年_2410至2508, y2 AS 后一年_2510至2608,
 y2-y1 AS 变化量, round(100.0*(y2-y1)/y1,1) AS 变化率pct
FROM d WHERE y1+y2 > 20000 ORDER BY 变化率pct DESC"""))

# ---------- 5. 响应时长（已关闭记录，created→closed） ----------
save("响应_类型.csv", q(f"""
SELECT complaint_type AS 投诉类型, count(*) AS 已关闭记录数,
 round(median(date_diff('minute', CAST(created_date AS TIMESTAMP), CAST(closed_date AS TIMESTAMP)))/60.0,1) AS 中位小时,
 round(quantile_cont(date_diff('minute', CAST(created_date AS TIMESTAMP), CAST(closed_date AS TIMESTAMP)),0.9)/60.0,1) AS p90小时
FROM read_parquet('{RAW}')
WHERE closed_date IS NOT NULL AND closed_date != ''
  AND date_diff('minute', CAST(created_date AS TIMESTAMP), CAST(closed_date AS TIMESTAMP)) >= 0
GROUP BY 1 HAVING count(*) > 50000 ORDER BY 中位小时"""))

# 关闭时长为负（closed<created）的异常量
neg = q(f"""SELECT count(*) AS n FROM read_parquet('{RAW}')
WHERE closed_date IS NOT NULL AND closed_date != '' AND date_diff('minute', CAST(created_date AS TIMESTAMP), CAST(closed_date AS TIMESTAMP)) < 0""")
save("响应_负时长异常.csv", neg)

# ---------- 6. Illegal Parking 处置结构 ----------
save("违停_处置结构.csv", q(f"""
SELECT CASE
  WHEN resolution_description ILIKE '%issued a summons%' THEN '开具传票'
  WHEN resolution_description ILIKE '%took action to fix%' THEN '到场采取行动'
  WHEN resolution_description ILIKE '%determined that a violation%' THEN '认定存在违规'
  WHEN resolution_description ILIKE '%were gone%' OR resolution_description ILIKE '%unable to locate%' THEN '到场时车辆已离开'
  WHEN resolution_description ILIKE '%no criminal violation%' OR resolution_description ILIKE '%no evidence%' OR resolution_description ILIKE '%no violation%' THEN '到场未见违规'
  WHEN resolution_description ILIKE '%not necessary%' OR resolution_description ILIKE '%no further action%' OR resolution_description ILIKE '%no police action%' THEN '判定无需警察行动'
  ELSE '其他/未说明' END AS 处置结果,
 count(*) AS 记录数, round(100.0*count(*)/sum(count(*)) OVER (),1) AS 占比pct
FROM read_parquet('{RAW}') WHERE complaint_type='Illegal Parking' GROUP BY 1 ORDER BY 2 DESC"""))

save("违停_行政区.csv", q(f"""
SELECT borough AS 行政区, count(*) AS 记录数, round(100.0*count(*)/1159053,1) AS 占比pct
FROM read_parquet('{RAW}') WHERE complaint_type='Illegal Parking' GROUP BY 1 ORDER BY 2 DESC"""))

# ---------- 7. 供暖季对比（10月–3月） ----------
save("供暖_两季对比.csv", q(f"""
SELECT borough AS 行政区,
 sum(CASE WHEN CAST(created_date AS TIMESTAMP) BETWEEN TIMESTAMP '2024-10-01' AND TIMESTAMP '2025-04-01' THEN 1 ELSE 0 END) AS 供暖季2024_25,
 sum(CASE WHEN CAST(created_date AS TIMESTAMP) BETWEEN TIMESTAMP '2025-10-01' AND TIMESTAMP '2026-04-01' THEN 1 ELSE 0 END) AS 供暖季2025_26
FROM read_parquet('{RAW}') WHERE complaint_type='HEAT/HOT WATER' GROUP BY 1 ORDER BY 3 DESC"""))

# ---------- 8. 冬季除雪对比（12–2月 Snow or Ice） ----------
save("冬季_两季对比.csv", q(f"""
SELECT 'Snow or Ice' AS 类型,
 sum(CASE WHEN CAST(created_date AS TIMESTAMP) BETWEEN TIMESTAMP '2024-12-01' AND TIMESTAMP '2025-03-01' THEN 1 ELSE 0 END) AS 冬2024_25,
 sum(CASE WHEN CAST(created_date AS TIMESTAMP) BETWEEN TIMESTAMP '2025-12-01' AND TIMESTAMP '2026-03-01' THEN 1 ELSE 0 END) AS 冬2025_26
FROM read_parquet('{RAW}') WHERE complaint_type='Snow or Ice'"""))

print("analysis.py 完成")
