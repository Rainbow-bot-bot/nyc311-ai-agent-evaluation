"""步骤2：探索可分析维度的取值分布，以及会真正影响分析的质量问题。"""
import duckdb, os

RAW = r"D:\项目1\原始数据"
GLOB = os.path.join(RAW, "*.parquet").replace("\\", "/")
con = duckdb.connect()
con.execute("PRAGMA threads=8")

# 统一视图：解析时间戳
con.execute(f"""
CREATE VIEW sr AS
SELECT unique_key,
       strptime(created_date, '%Y-%m-%dT%H:%M:%S.%g') AS created_ts,
       strptime(closed_date,  '%Y-%m-%dT%H:%M:%S.%g') AS closed_ts,
       agency, agency_name, complaint_type, descriptor, status, borough,
       incident_zip, open_data_channel_type, location_type, community_board,
       council_district, police_precinct, resolution_description
FROM read_parquet('{GLOB}')
""")

def show(title, q, n=None):
    print(f"\n=== {title} ===")
    df = con.execute(q).fetchdf()
    print(df.to_string(index=False, max_rows=n))

show("时间戳解析失败数", """
SELECT COUNT(*) AS total,
       SUM(created_ts IS NULL) AS created_parse_fail,
       SUM(closed_date IS NOT NULL) AS closed_present
FROM sr, (SELECT 1) WHERE TRUE
""".replace("closed_date IS NOT NULL", "closed_ts IS NOT NULL"))

show("agency 分布（全部16个）", """
SELECT agency, any_value(agency_name) AS agency_name, COUNT(*) AS n,
       ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS pct
FROM sr GROUP BY agency ORDER BY n DESC
""")

show("status 分布", """
SELECT status, COUNT(*) AS n, ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS pct,
       SUM(closed_ts IS NULL) AS no_closed_date
FROM sr GROUP BY status ORDER BY n DESC
""")

show("borough 分布", """
SELECT borough, COUNT(*) AS n, ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS pct
FROM sr GROUP BY borough ORDER BY n DESC
""")

show("open_data_channel_type 分布", """
SELECT open_data_channel_type, COUNT(*) AS n, ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS pct
FROM sr GROUP BY 1 ORDER BY n DESC
""")

show("complaint_type Top 30", """
SELECT complaint_type, COUNT(*) AS n, ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS pct,
       COUNT(DISTINCT agency) AS n_agency
FROM sr GROUP BY 1 ORDER BY n DESC LIMIT 30
""", n=40)

show("complaint_type 集中度", """
WITH t AS (SELECT complaint_type, COUNT(*) n FROM sr GROUP BY 1),
r AS (SELECT *, ROW_NUMBER() OVER (ORDER BY n DESC) rk, SUM(n) OVER () tot FROM t)
SELECT MAX(CASE WHEN rk<=10 THEN cum END) AS top10_share,
       MAX(CASE WHEN rk<=20 THEN cum END) AS top20_share,
       MAX(CASE WHEN rk<=50 THEN cum END) AS top50_share
FROM (SELECT rk, ROUND(100.0*SUM(n) OVER (ORDER BY rk)/tot,2) cum FROM r)
""")

show("质量问题：closed_ts < created_ts / 未来关闭 / 超长时长", """
SELECT COUNT(*) AS closed_rows,
       SUM(closed_ts < created_ts)::BIGINT AS closed_before_created,
       SUM(closed_ts > TIMESTAMP '2026-09-07')::BIGINT AS closed_after_data_end,
       SUM(date_diff('day', created_ts, closed_ts) > 365)::BIGINT AS over_365_days,
       SUM(date_diff('second', created_ts, closed_ts) = 0)::BIGINT AS zero_seconds
FROM sr WHERE closed_ts IS NOT NULL
""")

show("每月记录数与关闭率（含右截断观察）", """
SELECT date_trunc('month', created_ts)::DATE AS mon, COUNT(*) AS n,
       ROUND(100.0*SUM(CASE WHEN status IN ('Closed') THEN 1 ELSE 0 END)/COUNT(*),2) AS pct_status_closed,
       ROUND(100.0*SUM(closed_ts IS NOT NULL)/COUNT(*),2) AS pct_has_closed_date
FROM sr GROUP BY 1 ORDER BY 1
""", n=40)

show("incident_zip 异常值样例", """
SELECT incident_zip, COUNT(*) n FROM sr
WHERE incident_zip IS NOT NULL AND NOT regexp_matches(incident_zip, '^[0-9]{5}$')
GROUP BY 1 ORDER BY n DESC LIMIT 15
""")

show("location_type Top 15", """
SELECT location_type, COUNT(*) n FROM sr GROUP BY 1 ORDER BY n DESC LIMIT 15
""")
