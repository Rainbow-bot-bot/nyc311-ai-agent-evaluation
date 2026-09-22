"""步骤4：验证候选发现，重点区分"口径/上报变化"与"现实变化"。"""
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
       open_data_channel_type AS channel, incident_address, resolution_description
FROM read_parquet('{GLOB}')
""")
def show(t,q,n=None):
    print(f"\n=== {t} ===")
    print(con.execute(q).fetchdf().to_string(index=False, max_rows=n))

show("Snow or Ice 月度（判断是否单次天气事件）", """
SELECT date_trunc('month',created_ts)::DATE mon, COUNT(*) n FROM sr
WHERE complaint_type='Snow or Ice' GROUP BY 1 ORDER BY 1
""", n=40)

show("Street Condition / Damaged Tree 月度", """
SELECT date_trunc('month',created_ts)::DATE mon,
  SUM(complaint_type='Street Condition') street_cond,
  SUM(complaint_type='Damaged Tree') damaged_tree,
  SUM(complaint_type='Illegal Parking') illegal_parking,
  SUM(complaint_type='HEAT/HOT WATER') heat
FROM sr GROUP BY 1 ORDER BY 1
""", n=40)

show("疑似口径变更：DEP 相关类型的首次/末次出现月", """
SELECT complaint_type, MIN(date_trunc('month',created_ts))::DATE first_mon,
       MAX(date_trunc('month',created_ts))::DATE last_mon, COUNT(*) n
FROM sr WHERE agency='DEP' GROUP BY 1 ORDER BY n DESC
""", n=40)

show("疑似口径变更：Water/Sewer Maintenance 与 Water System 月度", """
SELECT date_trunc('month',created_ts)::DATE mon,
  SUM(complaint_type='Water System') water_system,
  SUM(complaint_type='Water Maintenance') water_maint,
  SUM(complaint_type='Sewer Maintenance') sewer_maint,
  SUM(complaint_type='Sewer') sewer,
  SUM(agency='DEP') dep_total
FROM sr GROUP BY 1 ORDER BY 1
""", n=40)

show("下降类型的月度：Drug Activity / Lead / Rodent / Noise-Helicopter", """
SELECT date_trunc('month',created_ts)::DATE mon,
  SUM(complaint_type='Drug Activity') drug,
  SUM(complaint_type='Lead') lead_c,
  SUM(complaint_type='Rodent') rodent,
  SUM(complaint_type='Noise - Helicopter') heli
FROM sr GROUP BY 1 ORDER BY 1
""", n=40)

show("所有 complaint_type 的首次/末次出现月（找口径切换）", """
SELECT complaint_type, MIN(date_trunc('month',created_ts))::DATE first_mon,
       MAX(date_trunc('month',created_ts))::DATE last_mon, COUNT(*) n
FROM sr GROUP BY 1
HAVING (MIN(created_ts) >= '2024-11-01' OR MAX(created_ts) < '2026-07-01') AND COUNT(*) >= 2000
ORDER BY n DESC
""", n=60)

show("UNKNOWN 渠道：是否为整类缺失", """
SELECT complaint_type, COUNT(*) n, ROUND(100.0*SUM(channel='UNKNOWN')/COUNT(*),1) pct_unknown
FROM sr GROUP BY 1 HAVING COUNT(*)>20000 AND SUM(channel='UNKNOWN')>0
ORDER BY pct_unknown DESC LIMIT 25
""", n=30)

show("排除 UNKNOWN 后的渠道结构按月", """
SELECT date_trunc('month',created_ts)::DATE mon, COUNT(*) n_known,
  ROUND(100.0*SUM(channel='ONLINE')/COUNT(*),1) online,
  ROUND(100.0*SUM(channel='MOBILE')/COUNT(*),1) mobile,
  ROUND(100.0*SUM(channel='PHONE')/COUNT(*),1) phone
FROM sr WHERE channel IN ('ONLINE','MOBILE','PHONE') GROUP BY 1 ORDER BY 1
""", n=40)

show("NYPD 结案性质分类（是否真正处置）", """
SELECT CASE
  WHEN resolution_description ILIKE '%took action to fix%' THEN 'A 采取行动处置'
  WHEN resolution_description ILIKE '%unable to gain entry%' OR resolution_description ILIKE '%unable to%' THEN 'C 无法进入/无法处理'
  WHEN resolution_description ILIKE '%no evidence%' OR resolution_description ILIKE '%observed no%'
       OR resolution_description ILIKE '%police action was not necessary%'
       OR resolution_description ILIKE '%determined that%' THEN 'B 到场未发现/无需行动'
  WHEN resolution_description ILIKE '%information available%' THEN 'B 到场未发现/无需行动'
  ELSE 'D 其他/未分类' END AS outcome,
  COUNT(*) n, ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),1) pct
FROM sr WHERE agency='NYPD' AND status='Closed'
GROUP BY 1 ORDER BY n DESC
""")

show("重复上报最集中的 complaint_type", """
WITH g AS (SELECT created_ts::DATE d, complaint_type, incident_address, COUNT(*) c
  FROM sr WHERE incident_address IS NOT NULL GROUP BY 1,2,3)
SELECT complaint_type, SUM(c) n_rows, SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END) excess,
       ROUND(100.0*SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END)/SUM(c),2) pct_excess
FROM g GROUP BY 1 HAVING SUM(c)>50000 ORDER BY pct_excess DESC LIMIT 15
""")

show("按周的总量（看是否有系统性断档）", """
SELECT date_trunc('week',created_ts)::DATE wk, COUNT(*) n FROM sr
GROUP BY 1 ORDER BY n LIMIT 10
""")
