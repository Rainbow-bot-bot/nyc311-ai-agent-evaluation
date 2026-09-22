"""步骤5：把同比增长拆成可加的贡献份额，并核对若干口径细节。"""
import duckdb, os
RAW = r"D:\项目1\原始数据"
GLOB = os.path.join(RAW, "*.parquet").replace("\\", "/")
con = duckdb.connect(); con.execute("PRAGMA threads=8")
con.execute(f"""
CREATE VIEW sr AS
SELECT unique_key,
       strptime(created_date, '%Y-%m-%dT%H:%M:%S.%g') AS created_ts,
       strptime(closed_date,  '%Y-%m-%dT%H:%M:%S.%g') AS closed_ts,
       agency, complaint_type, descriptor, status, borough,
       open_data_channel_type AS channel, resolution_description
FROM read_parquet('{GLOB}')
""")
def show(t,q,n=None):
    print(f"\n=== {t} ===")
    print(con.execute(q).fetchdf().to_string(index=False, max_rows=n))

Y1 = "created_ts >= '2024-09-07' AND created_ts < '2025-09-07'"
Y2 = "created_ts >= '2025-09-07' AND created_ts < '2026-09-06'"

show("同比增量的贡献份额（Top15 + 其余合计，校验加总）", f"""
WITH a AS (SELECT complaint_type, COUNT(*) n1 FROM sr WHERE {Y1} GROUP BY 1),
     b AS (SELECT complaint_type, COUNT(*) n2 FROM sr WHERE {Y2} GROUP BY 1),
j AS (SELECT COALESCE(a.complaint_type,b.complaint_type) ct, COALESCE(n1,0) n1, COALESCE(n2,0) n2,
             COALESCE(n2,0)-COALESCE(n1,0) d FROM a FULL JOIN b USING (complaint_type)),
tot AS (SELECT SUM(d) total_delta FROM j)
SELECT ct, n1, n2, d, ROUND(100.0*d/(SELECT total_delta FROM tot),1) pct_of_net_growth
FROM j ORDER BY ABS(d) DESC LIMIT 15
""", n=20)

show("净增量校验", f"""
SELECT (SELECT COUNT(*) FROM sr WHERE {Y2}) - (SELECT COUNT(*) FROM sr WHERE {Y1}) AS net_delta
""")

show("冬季口径对比（12-01~次年03-31）", """
SELECT CASE WHEN created_ts >= '2024-12-01' AND created_ts < '2025-04-01' THEN 'Winter 2024-25'
            WHEN created_ts >= '2025-12-01' AND created_ts < '2026-04-01' THEN 'Winter 2025-26' END AS winter,
  COUNT(*) all_rows,
  SUM(complaint_type='Snow or Ice') snow_ice,
  SUM(complaint_type='HEAT/HOT WATER') heat,
  SUM(complaint_type='Street Condition') street_cond
FROM sr WHERE created_ts >= '2024-12-01' AND created_ts < '2026-04-01'
  AND (created_ts < '2025-04-01' OR created_ts >= '2025-12-01')
GROUP BY 1 ORDER BY 1
""")

show("三类天气敏感型合计对净增长的贡献", f"""
WITH w AS (SELECT 'weather3' k,
  SUM(CASE WHEN {Y2} THEN 1 ELSE 0 END) - SUM(CASE WHEN {Y1} THEN 1 ELSE 0 END) d
  FROM sr WHERE complaint_type IN ('Snow or Ice','HEAT/HOT WATER','Street Condition'))
SELECT d AS weather3_delta, 325070 AS net_delta, ROUND(100.0*d/325070,1) AS pct_of_net FROM w
""")

show("DEP 口径迁移核对：2026-07/08 逐日", """
SELECT created_ts::DATE d, SUM(complaint_type='Water System') water_system,
       SUM(complaint_type='Water Maintenance') water_maint,
       SUM(complaint_type='Sewer') sewer, SUM(complaint_type='Sewer Maintenance') sewer_maint
FROM sr WHERE created_ts >= '2026-07-20' AND created_ts < '2026-08-10' GROUP BY 1 ORDER BY 1
""", n=30)

show("NYPD 全部结案描述 Top20（用于结案性质分类）", """
SELECT left(resolution_description,110) res, COUNT(*) n,
       ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) pct
FROM sr WHERE agency='NYPD' AND status='Closed' GROUP BY 1 ORDER BY n DESC LIMIT 20
""", n=25)

show("HPD HEAT/HOT WATER 结案描述 Top8", """
SELECT left(resolution_description,110) res, COUNT(*) n,
       ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) pct
FROM sr WHERE complaint_type='HEAT/HOT WATER' AND status='Closed' GROUP BY 1 ORDER BY n DESC LIMIT 8
""", n=10)

show("按 agency 的时长在两年间是否变化（同队列口径）", """
SELECT agency,
  ROUND(median(CASE WHEN created_ts<'2025-06-01' THEN date_diff('minute',created_ts,closed_ts) END)/60.0,2) med_h_y1,
  ROUND(median(CASE WHEN created_ts>='2025-06-01' THEN date_diff('minute',created_ts,closed_ts) END)/60.0,2) med_h_y2
FROM sr WHERE created_ts>='2024-10-01' AND created_ts<'2026-06-01'
  AND closed_ts IS NOT NULL AND closed_ts>=created_ts
GROUP BY 1 ORDER BY 1
""", n=20)
