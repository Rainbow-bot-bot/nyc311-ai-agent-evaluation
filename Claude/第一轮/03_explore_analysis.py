"""步骤3：围绕候选分析问题做证据探索（同比、渠道、时效、季节性、地理、重复上报）。"""
import duckdb, os

RAW = r"D:\项目1\原始数据"
GLOB = os.path.join(RAW, "*.parquet").replace("\\", "/")
con = duckdb.connect()
con.execute("PRAGMA threads=8")
con.execute(f"""
CREATE VIEW sr AS
SELECT unique_key,
       strptime(created_date, '%Y-%m-%dT%H:%M:%S.%g') AS created_ts,
       strptime(closed_date,  '%Y-%m-%dT%H:%M:%S.%g') AS closed_ts,
       agency, complaint_type, descriptor, status, borough, incident_zip,
       open_data_channel_type AS channel, location_type, incident_address,
       resolution_description
FROM read_parquet('{GLOB}')
""")

def show(t, q, n=None):
    print(f"\n=== {t} ===")
    print(con.execute(q).fetchdf().to_string(index=False, max_rows=n))

# --- 同比：两个近似完整年窗口 ---
Y1 = "created_ts >= '2024-09-07' AND created_ts < '2025-09-07'"
Y2 = "created_ts >= '2025-09-07' AND created_ts < '2026-09-06'"
show("两年窗口总量与天数", f"""
SELECT 'Y1 2024-09-07~2025-09-06' AS win, COUNT(*) n, COUNT(DISTINCT created_ts::DATE) AS n_days,
       ROUND(COUNT(*)*1.0/COUNT(DISTINCT created_ts::DATE),1) per_day FROM sr WHERE {Y1}
UNION ALL
SELECT 'Y2 2025-09-07~2026-09-05', COUNT(*), COUNT(DISTINCT created_ts::DATE),
       ROUND(COUNT(*)*1.0/COUNT(DISTINCT created_ts::DATE),1) FROM sr WHERE {Y2}
""")

show("同比变化最大的 complaint_type（按绝对增量）", f"""
WITH a AS (SELECT complaint_type, COUNT(*) n1 FROM sr WHERE {Y1} GROUP BY 1),
     b AS (SELECT complaint_type, COUNT(*) n2 FROM sr WHERE {Y2} GROUP BY 1)
SELECT COALESCE(a.complaint_type,b.complaint_type) AS complaint_type,
       COALESCE(n1,0) n1, COALESCE(n2,0) n2, COALESCE(n2,0)-COALESCE(n1,0) AS delta,
       CASE WHEN COALESCE(n1,0)>0 THEN ROUND(100.0*(COALESCE(n2,0)-n1)/n1,1) END AS pct_chg
FROM a FULL JOIN b USING (complaint_type)
ORDER BY ABS(COALESCE(n2,0)-COALESCE(n1,0)) DESC LIMIT 25
""", n=30)

show("同比：按 agency", f"""
WITH a AS (SELECT agency, COUNT(*) n1 FROM sr WHERE {Y1} GROUP BY 1),
     b AS (SELECT agency, COUNT(*) n2 FROM sr WHERE {Y2} GROUP BY 1)
SELECT agency, n1, n2, n2-n1 AS delta, ROUND(100.0*(n2-n1)/n1,1) pct_chg
FROM a JOIN b USING (agency) ORDER BY ABS(n2-n1) DESC
""", n=20)

show("渠道结构按月", """
SELECT date_trunc('month', created_ts)::DATE mon,
       ROUND(100.0*SUM(channel='ONLINE')/COUNT(*),1) online,
       ROUND(100.0*SUM(channel='MOBILE')/COUNT(*),1) mobile,
       ROUND(100.0*SUM(channel='PHONE')/COUNT(*),1) phone,
       ROUND(100.0*SUM(channel='UNKNOWN')/COUNT(*),1) unk,
       COUNT(*) n
FROM sr GROUP BY 1 ORDER BY 1
""", n=40)

show("UNKNOWN 渠道集中在哪些 agency/complaint_type", """
SELECT agency, complaint_type, COUNT(*) n FROM sr WHERE channel='UNKNOWN'
GROUP BY 1,2 ORDER BY n DESC LIMIT 12
""")

# --- 时效 ---
COH = "created_ts >= '2024-10-01' AND created_ts < '2026-06-01'"  # 留 >=3 月观察期
show("时效队列的完成情况（判断右截断影响）", f"""
SELECT COUNT(*) n, ROUND(100.0*SUM(closed_ts IS NOT NULL)/COUNT(*),2) pct_closed
FROM sr WHERE {COH}
""")

show("按 agency 的处理时长（小时，仅已关闭且时长>=0）", f"""
SELECT agency, COUNT(*) n_closed,
  ROUND(median(date_diff('minute',created_ts,closed_ts))/60.0,2) med_h,
  ROUND(quantile_cont(date_diff('minute',created_ts,closed_ts),0.9)/60.0,1) p90_h,
  ROUND(100.0*SUM(date_diff('minute',created_ts,closed_ts)=0)/COUNT(*),2) pct_zero_min,
  ROUND(100.0*SUM(date_diff('hour',created_ts,closed_ts)<=24)/COUNT(*),1) pct_le_24h,
  ROUND(100.0*SUM(date_diff('day',created_ts,closed_ts)>30)/COUNT(*),1) pct_gt_30d
FROM sr WHERE {COH} AND closed_ts IS NOT NULL AND closed_ts >= created_ts
GROUP BY 1 ORDER BY n_closed DESC
""", n=20)

show("Top20 complaint_type 的处理时长", f"""
WITH t AS (SELECT complaint_type, COUNT(*) n FROM sr WHERE {COH} GROUP BY 1 ORDER BY n DESC LIMIT 20)
SELECT s.complaint_type, any_value(s.agency) agency, COUNT(*) n_closed,
  ROUND(median(date_diff('minute',created_ts,closed_ts))/60.0,2) med_h,
  ROUND(quantile_cont(date_diff('minute',created_ts,closed_ts),0.9)/60.0,1) p90_h,
  ROUND(100.0*SUM(date_diff('minute',created_ts,closed_ts)=0)/COUNT(*),2) pct_zero_min
FROM sr s JOIN t USING (complaint_type)
WHERE {COH} AND closed_ts IS NOT NULL AND closed_ts >= created_ts
GROUP BY 1 ORDER BY med_h
""", n=25)

show("0 分钟关闭集中在哪里", """
SELECT agency, complaint_type, COUNT(*) n FROM sr
WHERE closed_ts IS NOT NULL AND date_diff('minute',created_ts,closed_ts)=0
GROUP BY 1,2 ORDER BY n DESC LIMIT 15
""")

show("NYPD 噪音类 resolution_description Top", """
SELECT left(resolution_description, 90) AS res, COUNT(*) n FROM sr
WHERE agency='NYPD' AND complaint_type LIKE 'Noise%'
GROUP BY 1 ORDER BY n DESC LIMIT 10
""")

# --- 季节性 ---
show("Top10 complaint_type 的月度分布（用于季节性）", """
WITH t AS (SELECT complaint_type FROM sr GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 10)
SELECT month(created_ts) m, s.complaint_type, COUNT(*) n
FROM sr s JOIN t USING (complaint_type) GROUP BY 1,2 ORDER BY 2,1
""", n=200)

# --- 地理组成 ---
show("borough × Top8 complaint_type 组成占比", """
WITH t AS (SELECT complaint_type FROM sr GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 8)
SELECT borough, s.complaint_type, COUNT(*) n,
       ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (PARTITION BY borough),2) pct_in_boro
FROM sr s JOIN t USING (complaint_type)
WHERE borough <> 'Unspecified' GROUP BY 1,2 ORDER BY 1,3 DESC
""", n=100)

# --- 重复上报 ---
show("同址同类同日多条记录（记录数 vs 现实事件）", """
WITH g AS (
  SELECT created_ts::DATE d, complaint_type, incident_address, COUNT(*) c
  FROM sr WHERE incident_address IS NOT NULL GROUP BY 1,2,3)
SELECT SUM(c) AS rows_with_addr,
       SUM(CASE WHEN c>1 THEN c ELSE 0 END) AS rows_in_dup_groups,
       ROUND(100.0*SUM(CASE WHEN c>1 THEN c ELSE 0 END)/SUM(c),2) AS pct,
       SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END) AS excess_rows
FROM g
""")

show("重复上报最集中的 complaint_type", """
WITH g AS (
  SELECT created_ts::DATE d, complaint_type, incident_address, COUNT(*) c
  FROM sr WHERE incident_address IS NOT NULL GROUP BY 1,2,3)
SELECT complaint_type, SUM(c) rows, SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END) excess,
       ROUND(100.0*SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END)/SUM(c),2) pct_excess
FROM g GROUP BY 1 HAVING SUM(c)>50000 ORDER BY pct_excess DESC LIMIT 15
""")
