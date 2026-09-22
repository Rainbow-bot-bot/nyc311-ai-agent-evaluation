"""步骤6：生成最终看板所需的全部汇总表，输出到 中间汇总/*.csv（可复现、可回查）。

统一口径（全脚本一致）：
- 记录粒度：1 行 = 1 条 311 服务请求（unique_key 唯一，7,525,498 条，无重复主键）
- 时间基准：created_date（请求创建时间），解析格式 %Y-%m-%dT%H:%M:%S.%g
- 全量窗口：2024-09-07 ~ 2026-09-05（首尾月份不完整）
- 同比窗口：Y1 = 2024-09-07~2025-09-06（365 天）；Y2 = 2025-09-07~2026-09-05（364 天）
- 完整月窗口：2024-10-01 ~ 2026-08-31（23 个完整自然月）
- 季节性窗口：2025-09-01 ~ 2026-08-31（最近 12 个完整自然月）
- 时效队列：created 在 2024-10-01~2026-05-31（留 >=3 个月观察期），仅取已有 closed_date
  且 closed_date >= created_date 的记录

分析口径的一处剔除（见 t28/t29）：
  规则 = 剔除任一 (incident_address, complaint_type) 组合，若其"单日记录数最大值 >= 200"。
  该阈值落在数据的自然断点上：符合条件的组合只有 4 个（单日峰值 4,978 / 2,281 / 1,109 / 272），
  紧邻其下的最大单日簇为 184，且住宅噪音类"单地址两年总量"的 99 分位数只有 65。
  这 4 个组合全部是布朗克斯东 230/231 街同一街区的 Noise - Residential 记录，共 185,791 条。
  这种量级不可能与现实事件一一对应，因此从分析口径剔除；数据集原始总量仍按 7,525,498 报告。
"""
import duckdb, os, pathlib

RAW = r"D:\项目1\原始数据"
GLOB = os.path.join(RAW, "*.parquet").replace("\\", "/")
OUT = pathlib.Path(r"D:\项目2\Claude\中间汇总")
OUT.mkdir(exist_ok=True)

con = duckdb.connect(); con.execute("PRAGMA threads=8")

# sr_all = 数据集原样（用于报告原始总量与记录异常本身）
con.execute(f"""
CREATE VIEW sr_all AS
SELECT unique_key,
       strptime(created_date, '%Y-%m-%dT%H:%M:%S.%g') AS created_ts,
       strptime(closed_date,  '%Y-%m-%dT%H:%M:%S.%g') AS closed_ts,
       agency, agency_name, complaint_type, descriptor, status, borough,
       incident_zip, incident_address, open_data_channel_type AS channel,
       location_type, resolution_description
FROM read_parquet('{GLOB}')
""")

# 异常聚集识别：单日单址单类型记录数 >= 200 的 (地址, 类型) 组合
con.execute("""
CREATE TABLE excluded_pairs AS
WITH g AS (SELECT incident_address, complaint_type, created_ts::DATE d, COUNT(*) c
           FROM sr_all WHERE incident_address IS NOT NULL GROUP BY 1,2,3)
SELECT incident_address, complaint_type, MAX(c) AS peak_day,
       SUM(c) AS total_rows, COUNT(*) AS n_days
FROM g GROUP BY 1,2 HAVING MAX(c) >= 200
""")

# sr = 分析口径（剔除上述异常聚集）
con.execute("""
CREATE VIEW sr AS
SELECT s.* FROM sr_all s
LEFT JOIN excluded_pairs e
  ON s.incident_address = e.incident_address AND s.complaint_type = e.complaint_type
WHERE e.incident_address IS NULL
""")

N_ALL = con.execute("SELECT COUNT(*) FROM sr_all").fetchone()[0]
N_EXCL = con.execute("SELECT COALESCE(SUM(total_rows),0) FROM excluded_pairs").fetchone()[0]
N_ANA = con.execute("SELECT COUNT(*) FROM sr").fetchone()[0]
print(f"原始总量={N_ALL:,}  剔除={N_EXCL:,}  分析口径={N_ANA:,}")
assert N_ALL - N_EXCL == N_ANA, "剔除量与分析口径不自洽"

Y1 = "created_ts >= '2024-09-07' AND created_ts < '2025-09-07'"
Y2 = "created_ts >= '2025-09-07' AND created_ts < '2026-09-06'"
FULLMON = "created_ts >= '2024-10-01' AND created_ts < '2026-09-01'"
SEASON = "created_ts >= '2025-09-01' AND created_ts < '2026-09-01'"
COH = "created_ts >= '2024-10-01' AND created_ts < '2026-06-01'"
DUR = "date_diff('minute', created_ts, closed_ts)"

def thousands(v):
    """把纯整数字符串加千分位，便于人工阅读；非整数原样返回。"""
    s = str(v)
    return f"{int(s):,}" if s.isdigit() else s

NET = (con.execute(f"SELECT COUNT(*) FROM sr WHERE {Y2}").fetchone()[0]
       - con.execute(f"SELECT COUNT(*) FROM sr WHERE {Y1}").fetchone()[0])
print(f"分析口径同比净增量 NET={NET:,}")

def save(name, q, comma_cols=()):
    df = con.execute(q).fetchdf()
    for c in comma_cols:
        df[c] = df[c].map(thousands)
    df.to_csv(OUT / f"{name}.csv", index=False, encoding="utf-8-sig")
    print(f"[{name}] {df.shape[0]} 行 x {df.shape[1]} 列")
    return df

# ---------- 1. 数据概况 ----------
save("t01_overview", f"""
SELECT '数据集原始记录数' AS 指标, COUNT(*)::VARCHAR AS 值 FROM sr_all
UNION ALL SELECT '唯一 unique_key 数', COUNT(DISTINCT unique_key)::VARCHAR FROM sr_all
UNION ALL SELECT '剔除的异常聚集记录数（见 3.2）', {N_EXCL}::VARCHAR
UNION ALL SELECT '本看板分析口径记录数', {N_ANA}::VARCHAR
UNION ALL SELECT '源文件数（Parquet 分片）', '25'
UNION ALL SELECT '字段数', '44'
UNION ALL SELECT 'created_date 最早', MIN(created_ts)::VARCHAR FROM sr_all
UNION ALL SELECT 'created_date 最晚', MAX(created_ts)::VARCHAR FROM sr_all
UNION ALL SELECT '覆盖天数', COUNT(DISTINCT created_ts::DATE)::VARCHAR FROM sr_all
UNION ALL SELECT '受理机构数 (agency)', COUNT(DISTINCT agency)::VARCHAR FROM sr_all
UNION ALL SELECT '问题类型数 (complaint_type)', COUNT(DISTINCT complaint_type)::VARCHAR FROM sr_all
UNION ALL SELECT '细分描述数 (descriptor)', COUNT(DISTINCT descriptor)::VARCHAR FROM sr_all
UNION ALL SELECT '行政区数（含 Unspecified）', COUNT(DISTINCT borough)::VARCHAR FROM sr_all
UNION ALL SELECT '邮编数 (incident_zip)', COUNT(DISTINCT incident_zip)::VARCHAR FROM sr_all
UNION ALL SELECT 'status = Closed 占比（分析口径）',
  ROUND(100.0*SUM(status='Closed')/COUNT(*),2)::VARCHAR || '%' FROM sr
UNION ALL SELECT '有 closed_date 占比（分析口径）',
  ROUND(100.0*SUM(closed_ts IS NOT NULL)/COUNT(*),2)::VARCHAR || '%' FROM sr
""", comma_cols=("值",))

# ---------- 异常聚集的证据表 ----------
save("t28_excluded_pairs", """
SELECT incident_address AS 地址原文, complaint_type AS 问题类型,
       peak_day AS 单日最大记录数, total_rows AS 两年合计记录数, n_days AS 出现天数,
       ROUND(1.0*total_rows/n_days,0) AS 有记录日的日均
FROM excluded_pairs ORDER BY total_rows DESC
""")

save("t29_excluded_monthly", """
WITH ex AS (
  SELECT date_trunc('month', s.created_ts)::DATE mon, COUNT(*) n_ex
  FROM sr_all s JOIN excluded_pairs e
    ON s.incident_address=e.incident_address AND s.complaint_type=e.complaint_type
  GROUP BY 1),
nr AS (
  SELECT date_trunc('month', created_ts)::DATE mon, COUNT(*) n_nr
  FROM sr_all WHERE complaint_type='Noise - Residential' GROUP BY 1)
SELECT nr.mon AS 月份, COALESCE(n_ex,0) AS "东230街聚集记录数",
       n_nr AS "全市住宅噪音记录数（含聚集）",
       n_nr - COALESCE(n_ex,0) AS "住宅噪音（分析口径）",
       ROUND(100.0*COALESCE(n_ex,0)/n_nr,1) AS "聚集占该月住宅噪音%"
FROM nr LEFT JOIN ex USING (mon) ORDER BY 1
""")

save("t30_noise_yoy_effect", f"""
SELECT CASE WHEN created_ts >= '2024-09-07' AND created_ts < '2025-09-07' THEN 'Y1'
            WHEN created_ts >= '2025-09-07' AND created_ts < '2026-09-06' THEN 'Y2' END AS 窗口,
  SUM(complaint_type='Noise - Residential') AS "住宅噪音（含异常聚集）",
  SUM(complaint_type='Noise - Residential' AND incident_address IN
      (SELECT incident_address FROM excluded_pairs)) AS "其中异常聚集",
  SUM(complaint_type='Noise - Residential') - SUM(complaint_type='Noise - Residential'
      AND incident_address IN (SELECT incident_address FROM excluded_pairs))
      AS "住宅噪音（分析口径）"
FROM sr_all GROUP BY 1 HAVING 窗口 IS NOT NULL ORDER BY 1
""")

save("t02_quality", f"""
SELECT '单址单类型记录异常聚集' AS 质量项, {N_EXCL}::VARCHAR AS 记录数,
  ROUND(100.0*{N_EXCL}/COUNT(*),2)::VARCHAR||'%' AS 占全量,
  '影响最大的质量问题：4 个(地址,类型)组合单日峰值达 4,978 条，占全部住宅噪音记录 21.1%，'
  '会同时扭曲同比、行政区结构、季节性与重复上报结论，已从分析口径整体剔除（见 3.2）' AS 对分析的影响
  FROM sr_all
UNION ALL SELECT '主键重复（unique_key）', (COUNT(*)-COUNT(DISTINCT unique_key))::VARCHAR,
  ROUND(100.0*(COUNT(*)-COUNT(DISTINCT unique_key))/COUNT(*),3)::VARCHAR||'%' AS 占全量,
  '无问题，可用 unique_key 作为唯一记录标识' AS 对分析的影响 FROM sr
UNION ALL SELECT 'created_date 解析失败', SUM(created_ts IS NULL)::VARCHAR,
  ROUND(100.0*SUM(created_ts IS NULL)/COUNT(*),3)::VARCHAR||'%',
  '无问题，全部时间戳格式一致' FROM sr
UNION ALL SELECT '缺 closed_date（未结案或未回填）', SUM(closed_ts IS NULL)::VARCHAR,
  ROUND(100.0*SUM(closed_ts IS NULL)/COUNT(*),2)::VARCHAR||'%',
  '影响时效分析：近月比例升高（右截断），时效仅用留足观察期的队列' FROM sr
UNION ALL SELECT 'closed_date 早于 created_date', SUM(closed_ts < created_ts)::VARCHAR,
  ROUND(100.0*SUM(closed_ts < created_ts)/COUNT(*),3)::VARCHAR||'%',
  '逻辑不可能，时效分析中已剔除' FROM sr
UNION ALL SELECT '处理时长恰为 0 分钟', SUM({DUR}=0)::VARCHAR,
  ROUND(100.0*SUM({DUR}=0)/COUNT(*),2)::VARCHAR||'%',
  '多为系统批量结案（DEP/DOHMH/DOB 集中），会把中位时长压低，已单列比例' FROM sr
UNION ALL SELECT '处理时长 > 365 天', SUM(date_diff('day',created_ts,closed_ts)>365)::VARCHAR,
  ROUND(100.0*SUM(date_diff('day',created_ts,closed_ts)>365)/COUNT(*),3)::VARCHAR||'%',
  '长尾真实存在（TLC/EDC/DOB），用中位数与 P90 而非均值' FROM sr
UNION ALL SELECT 'borough = Unspecified', SUM(borough='Unspecified')::VARCHAR,
  ROUND(100.0*SUM(borough='Unspecified')/COUNT(*),3)::VARCHAR||'%',
  '占比极低，行政区分析中已排除' FROM sr
UNION ALL SELECT '渠道 = UNKNOWN', SUM(channel='UNKNOWN')::VARCHAR,
  ROUND(100.0*SUM(channel='UNKNOWN')/COUNT(*),2)::VARCHAR||'%',
  '并非随机缺失：DOT/DOB 若干类型 100% 为 UNKNOWN；渠道分析已整体排除' FROM sr
UNION ALL SELECT 'incident_zip 非 5 位数字', SUM(incident_zip IS NOT NULL AND NOT regexp_matches(incident_zip,'^[0-9]{{5}}$'))::VARCHAR,
  '<0.001%', '数量可忽略，不影响结论' FROM sr
UNION ALL SELECT '缺 latitude/longitude', '129696', '1.72%',
  '未做点位地图分析，不影响本看板结论' 
""", comma_cols=("记录数",))

save("t03_file_inventory", f"""
SELECT regexp_extract(filename, '[^/\\\\]+$') AS 文件名, COUNT(*) AS 记录数,
       MIN(strptime(created_date,'%Y-%m-%dT%H:%M:%S.%g'))::DATE AS 最早创建日,
       MAX(strptime(created_date,'%Y-%m-%dT%H:%M:%S.%g'))::DATE AS 最晚创建日
FROM read_parquet('{GLOB}', filename=true) GROUP BY 1 ORDER BY 1
""")

# ---------- 2. 月度趋势 ----------
save("t04_monthly", f"""
SELECT date_trunc('month',created_ts)::DATE AS 月份, COUNT(*) AS 记录数,
  ROUND(COUNT(*)*1.0/COUNT(DISTINCT created_ts::DATE),0) AS 日均记录数,
  ROUND(100.0*SUM(status='Closed')/COUNT(*),2) AS "已结案占比%",
  CASE WHEN MIN(created_ts) < '2024-10-01' OR MIN(created_ts) >= '2026-09-01'
       THEN '不完整月（首/末月）' ELSE '完整月' END AS 月份完整性
FROM sr GROUP BY 1 ORDER BY 1
""")

save("t05_yoy_window", f"""
SELECT 'Y1 (2024-09-07~2025-09-06)' AS 窗口, COUNT(*) AS 记录数,
  COUNT(DISTINCT created_ts::DATE) AS 天数,
  ROUND(COUNT(*)*1.0/COUNT(DISTINCT created_ts::DATE),1) AS 日均记录数 FROM sr WHERE {Y1}
UNION ALL SELECT 'Y2 (2025-09-07~2026-09-05)', COUNT(*), COUNT(DISTINCT created_ts::DATE),
  ROUND(COUNT(*)*1.0/COUNT(DISTINCT created_ts::DATE),1) FROM sr WHERE {Y2}
""")

# ---------- 3. 同比归因 ----------
save("t06_yoy_by_type", f"""
WITH a AS (SELECT complaint_type, COUNT(*) n1 FROM sr WHERE {Y1} GROUP BY 1),
     b AS (SELECT complaint_type, COUNT(*) n2 FROM sr WHERE {Y2} GROUP BY 1),
j AS (SELECT COALESCE(a.complaint_type,b.complaint_type) ct, COALESCE(n1,0) n1, COALESCE(n2,0) n2
      FROM a FULL JOIN b USING (complaint_type))
SELECT ct AS 问题类型, n1 AS "Y1记录数", n2 AS "Y2记录数", n2-n1 AS 增减量,
  CASE WHEN n1>=100 THEN ROUND(100.0*(n2-n1)/n1,1) END AS "同比%",
  ROUND(100.0*(n2-n1)/{NET}.0,1) AS "占净增长%"
FROM j ORDER BY ABS(n2-n1) DESC LIMIT 20
""")

# 用 FULL JOIN，保证只在单个窗口出现的机构（如 OOS）不被丢掉，增减量之和才等于净增长
save("t07_yoy_by_agency", f"""
WITH a AS (SELECT agency, COUNT(*) n1 FROM sr WHERE {Y1} GROUP BY 1),
     b AS (SELECT agency, COUNT(*) n2 FROM sr WHERE {Y2} GROUP BY 1),
     nm AS (SELECT agency, any_value(agency_name) AS agname FROM sr GROUP BY 1),
j AS (SELECT COALESCE(a.agency,b.agency) ag, COALESCE(n1,0) n1, COALESCE(n2,0) n2
      FROM a FULL JOIN b USING (agency))
SELECT ag AS 机构代码, nm.agname AS 机构名称, n1 AS "Y1记录数", n2 AS "Y2记录数",
  n2-n1 AS 增减量,
  CASE WHEN n1>=100 THEN ROUND(100.0*(n2-n1)/n1,1) END AS "同比%",
  ROUND(100.0*(n2-n1)/{NET}.0,1) AS "占净增长%"
FROM j LEFT JOIN nm ON j.ag = nm.agency ORDER BY ABS(n2-n1) DESC
""")

save("t08_winter", """
SELECT CASE WHEN created_ts < '2025-04-01' THEN '2024-25 冬季 (12/01~03/31)'
            ELSE '2025-26 冬季 (12/01~03/31)' END AS 冬季,
  COUNT(*) AS 全部记录数,
  SUM(complaint_type='Snow or Ice') AS "冰雪 Snow or Ice",
  SUM(complaint_type='HEAT/HOT WATER') AS "供暖热水 HEAT/HOT WATER",
  SUM(complaint_type='Street Condition') AS "路面状况 Street Condition"
FROM sr
WHERE (created_ts >= '2024-12-01' AND created_ts < '2025-04-01')
   OR (created_ts >= '2025-12-01' AND created_ts < '2026-04-01')
GROUP BY 1 ORDER BY 1
""")

save("t09_weather_monthly", """
SELECT date_trunc('month',created_ts)::DATE AS 月份,
  SUM(complaint_type='Snow or Ice') AS "冰雪",
  SUM(complaint_type='HEAT/HOT WATER') AS "供暖热水",
  SUM(complaint_type='Street Condition') AS "路面状况",
  SUM(complaint_type='Damaged Tree') AS "树木受损"
FROM sr GROUP BY 1 ORDER BY 1
""")

# ---------- 4. 口径变更 ----------
save("t10_taxonomy_switch", """
SELECT date_trunc('month',created_ts)::DATE AS 月份,
  SUM(complaint_type='Water System') AS "Water System（旧）",
  SUM(complaint_type='Water Maintenance') AS "Water Maintenance（新）",
  SUM(complaint_type='Sewer') AS "Sewer（旧）",
  SUM(complaint_type='Sewer Maintenance') AS "Sewer Maintenance（新）",
  SUM(complaint_type IN ('Water System','Water Maintenance','Sewer','Sewer Maintenance')) AS "四类合计",
  SUM(agency='DEP') AS "DEP 全部"
FROM sr GROUP BY 1 ORDER BY 1
""")

save("t11_taxonomy_daily", """
SELECT created_ts::DATE AS 日期,
  SUM(complaint_type='Water System') AS "Water System（旧）",
  SUM(complaint_type='Water Maintenance') AS "Water Maintenance（新）",
  SUM(complaint_type='Sewer') AS "Sewer（旧）",
  SUM(complaint_type='Sewer Maintenance') AS "Sewer Maintenance（新）"
FROM sr WHERE created_ts >= '2026-07-20' AND created_ts < '2026-08-11'
GROUP BY 1 ORDER BY 1
""")

save("t12_type_lifespan", """
SELECT complaint_type AS 问题类型, any_value(agency) AS 机构,
  MIN(date_trunc('month',created_ts))::DATE AS 首次出现月,
  MAX(date_trunc('month',created_ts))::DATE AS 最后出现月, COUNT(*) AS 记录数,
  CASE WHEN MIN(created_ts) >= '2024-11-01' AND MAX(created_ts) >= '2026-08-01' THEN '窗口内新增'
       WHEN MAX(created_ts) < '2026-07-01' AND MIN(created_ts) < '2024-11-01' THEN '窗口内停用'
       ELSE '季节性/其他' END AS 判断
FROM sr GROUP BY 1
HAVING (MIN(created_ts) >= '2024-11-01' OR MAX(created_ts) < '2026-07-01') AND COUNT(*) >= 2000
ORDER BY 记录数 DESC
""")

# ---------- 5. 处理时效 ----------
save("t13_duration_by_agency", f"""
SELECT agency AS 机构代码, any_value(agency_name) AS 机构名称, COUNT(*) AS 已结案记录数,
  ROUND(median({DUR})/60.0,2) AS "中位时长(小时)",
  ROUND(quantile_cont({DUR},0.9)/60.0,1) AS "P90时长(小时)",
  ROUND(median({DUR})/1440.0,2) AS "中位时长(天)",
  ROUND(100.0*SUM({DUR}<=1440)/COUNT(*),1) AS "24小时内结案%",
  ROUND(100.0*SUM({DUR}>43200)/COUNT(*),1) AS "超30天结案%",
  ROUND(100.0*SUM({DUR}=0)/COUNT(*),2) AS "0分钟结案%"
FROM sr WHERE {COH} AND closed_ts IS NOT NULL AND closed_ts >= created_ts
GROUP BY 1 ORDER BY 已结案记录数 DESC
""")

save("t14_duration_by_type", f"""
WITH t AS (SELECT complaint_type FROM sr WHERE {COH} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 20)
SELECT s.complaint_type AS 问题类型, any_value(s.agency) AS 机构, COUNT(*) AS 已结案记录数,
  ROUND(median({DUR})/60.0,2) AS "中位时长(小时)",
  ROUND(quantile_cont({DUR},0.9)/60.0,1) AS "P90时长(小时)",
  ROUND(100.0*SUM({DUR}<=1440)/COUNT(*),1) AS "24小时内结案%",
  ROUND(100.0*SUM({DUR}=0)/COUNT(*),2) AS "0分钟结案%"
FROM sr s JOIN t USING (complaint_type)
WHERE {COH} AND closed_ts IS NOT NULL AND closed_ts >= created_ts
GROUP BY 1 ORDER BY "中位时长(小时)"
""")

save("t15_censoring", f"""
SELECT date_trunc('month',created_ts)::DATE AS 月份, COUNT(*) AS 记录数,
  ROUND(100.0*SUM(closed_ts IS NOT NULL)/COUNT(*),2) AS "已有结案时间占比%"
FROM sr GROUP BY 1 ORDER BY 1
""")

# ---------- 6. 结案性质 ----------
OUTCOME = """CASE
 WHEN resolution_description IS NULL THEN 'G 无结案说明'
 WHEN resolution_description ILIKE '%took action to fix%'
   OR resolution_description ILIKE '%issued a summons%'
   OR resolution_description ILIKE '%a violation was issued%'
   OR resolution_description ILIKE '%a report was prepared%' THEN 'A 采取行动/开票/立案'
 WHEN resolution_description ILIKE '%no evidence%'
   OR resolution_description ILIKE '%observed no%'
   OR resolution_description ILIKE '%no criminal violation%'
   OR resolution_description ILIKE '%police action was not necessary%'
   OR resolution_description ILIKE '%were gone%'
   OR resolution_description ILIKE '%no encampment was found%' THEN 'B 到场未发现问题/无需行动'
 WHEN resolution_description ILIKE '%unable to gain entry%'
   OR resolution_description ILIKE '%unable to%' THEN 'C 无法进入/无法核实'
 WHEN resolution_description ILIKE '%jurisdiction%' THEN 'D 不属本机构管辖'
 WHEN resolution_description ILIKE '%referred the complaint%' THEN 'E 转交其他机构'
 ELSE 'F 其他/仅提供信息' END"""

save("t16_nypd_outcome", f"""
SELECT {OUTCOME} AS 结案性质, COUNT(*) AS 记录数,
  ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS "占比%"
FROM sr WHERE agency='NYPD' AND status='Closed'
GROUP BY 1 ORDER BY 1
""")

save("t17_nypd_outcome_by_type", f"""
WITH t AS (SELECT complaint_type FROM sr WHERE agency='NYPD' AND status='Closed'
           GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 8)
SELECT s.complaint_type AS 问题类型, COUNT(*) AS 已结案记录数,
  ROUND(100.0*SUM(CASE WHEN {OUTCOME} = 'A 采取行动/开票/立案' THEN 1 ELSE 0 END)/COUNT(*),1) AS "A 采取行动%",
  ROUND(100.0*SUM(CASE WHEN {OUTCOME} = 'B 到场未发现问题/无需行动' THEN 1 ELSE 0 END)/COUNT(*),1) AS "B 未发现问题%",
  ROUND(median({DUR})/60.0,2) AS "中位时长(小时)"
FROM sr s JOIN t USING (complaint_type)
WHERE agency='NYPD' AND status='Closed'
GROUP BY 1 ORDER BY 已结案记录数 DESC
""")

save("t18_hpd_heat_reason", """
WITH g AS (
  SELECT resolution_description AS d, COUNT(*) n
  FROM sr WHERE complaint_type='HEAT/HOT WATER' AND status='Closed' AND resolution_description IS NOT NULL
  GROUP BY 1)
SELECT CASE WHEN length(d) > 84 THEN left(d, 84) || ' ...' ELSE d END AS 结案说明摘要,
  n AS 记录数, ROUND(100.0*n/SUM(n) OVER (),2) AS "占比%"
FROM g ORDER BY n DESC LIMIT 8
""")

# ---------- 7. 重复上报 ----------
save("t19_dup_overall", f"""
WITH g AS (SELECT created_ts::DATE d, complaint_type, incident_address, COUNT(*) c
           FROM sr WHERE incident_address IS NOT NULL GROUP BY 1,2,3)
SELECT SUM(c) AS "可判定记录数(有地址)", COUNT(*) AS "去重后事件簇数",
  SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END) AS 冗余记录数,
  ROUND(100.0*SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END)/SUM(c),2) AS "冗余占可判定记录%",
  ROUND(100.0*SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END)/{N_ANA}.0,2) AS "冗余占分析口径全量%"
FROM g
""")

save("t20_dup_by_type", """
WITH g AS (SELECT created_ts::DATE d, complaint_type, incident_address, COUNT(*) c
           FROM sr WHERE incident_address IS NOT NULL GROUP BY 1,2,3)
SELECT complaint_type AS 问题类型, SUM(c) AS "记录数(有地址)",
  COUNT(*) AS 事件簇数, SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END) AS 冗余记录数,
  ROUND(100.0*SUM(CASE WHEN c>1 THEN c-1 ELSE 0 END)/SUM(c),2) AS "冗余占比%",
  MAX(c) AS 单簇最大记录数
FROM g GROUP BY 1 HAVING SUM(c) > 50000 ORDER BY "冗余占比%" DESC LIMIT 15
""")

# ---------- 8. 季节性 ----------
save("t21_seasonality", f"""
WITH t AS (SELECT complaint_type FROM sr WHERE {SEASON}
             AND complaint_type NOT IN ('Water System','Sewer','Water Maintenance','Sewer Maintenance')
           GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 10),
m AS (SELECT s.complaint_type ct, month(created_ts) mm, COUNT(*) n
      FROM sr s JOIN t USING (complaint_type) WHERE {SEASON} GROUP BY 1,2)
SELECT ct AS 问题类型,
  MAX(CASE WHEN mm=1 THEN n END) AS "1月", MAX(CASE WHEN mm=2 THEN n END) AS "2月",
  MAX(CASE WHEN mm=3 THEN n END) AS "3月", MAX(CASE WHEN mm=4 THEN n END) AS "4月",
  MAX(CASE WHEN mm=5 THEN n END) AS "5月", MAX(CASE WHEN mm=6 THEN n END) AS "6月",
  MAX(CASE WHEN mm=7 THEN n END) AS "7月", MAX(CASE WHEN mm=8 THEN n END) AS "8月",
  MAX(CASE WHEN mm=9 THEN n END) AS "9月", MAX(CASE WHEN mm=10 THEN n END) AS "10月",
  MAX(CASE WHEN mm=11 THEN n END) AS "11月", MAX(CASE WHEN mm=12 THEN n END) AS "12月",
  SUM(n) AS 全年合计,
  ROUND(1.0*MAX(n)/NULLIF(MIN(n),0),1) AS "旺淡月倍数"
FROM m GROUP BY 1 ORDER BY 全年合计 DESC
""")

save("t22_hourly", f"""
SELECT hour(created_ts) AS 小时, COUNT(*) AS 记录数,
  ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS "占比%"
FROM sr WHERE {FULLMON} GROUP BY 1 ORDER BY 1
""")

# ---------- 9. 行政区 ----------
save("t23_borough", f"""
SELECT borough AS 行政区, COUNT(*) AS 记录数,
  ROUND(100.0*COUNT(*)/SUM(COUNT(*)) OVER (),2) AS "占全市%",
  ROUND(median({DUR})/60.0,2) AS "中位时长(小时,已结案)"
FROM sr WHERE borough <> 'Unspecified' GROUP BY 1 ORDER BY 记录数 DESC
""")

save("t24_borough_mix", f"""
WITH t AS (SELECT complaint_type FROM sr GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 10),
c AS (SELECT s.complaint_type ct, borough, COUNT(*) n FROM sr s JOIN t USING (complaint_type)
      WHERE borough <> 'Unspecified' GROUP BY 1,2),
tot AS (SELECT borough, COUNT(*) bn FROM sr WHERE borough<>'Unspecified' GROUP BY 1),
city AS (SELECT ct, SUM(n) cn, (SELECT SUM(bn) FROM tot) ctot FROM c GROUP BY 1)
SELECT c.ct AS 问题类型, c.borough AS 行政区, c.n AS 记录数,
  ROUND(100.0*c.n/tot.bn,2) AS "占该区全部记录%",
  ROUND(100.0*city.cn/city.ctot,2) AS "全市基准%",
  ROUND((1.0*c.n/tot.bn)/(1.0*city.cn/city.ctot),2) AS 集中度指数
FROM c JOIN tot ON c.borough=tot.borough JOIN city ON c.ct=city.ct
ORDER BY 问题类型, 集中度指数 DESC
""")

# ---------- 10. 渠道 ----------
save("t25_channel_monthly", """
SELECT date_trunc('month',created_ts)::DATE AS 月份, COUNT(*) AS "已知渠道记录数",
  ROUND(100.0*SUM(channel='ONLINE')/COUNT(*),1) AS "网页 ONLINE%",
  ROUND(100.0*SUM(channel='MOBILE')/COUNT(*),1) AS "移动端 MOBILE%",
  ROUND(100.0*SUM(channel='PHONE')/COUNT(*),1) AS "电话 PHONE%"
FROM sr WHERE channel IN ('ONLINE','MOBILE','PHONE') GROUP BY 1 ORDER BY 1
""")

save("t26_channel_unknown", """
SELECT complaint_type AS 问题类型, any_value(agency) AS 机构, COUNT(*) AS 记录数,
  ROUND(100.0*SUM(channel='UNKNOWN')/COUNT(*),1) AS "UNKNOWN 占比%"
FROM sr GROUP BY 1 HAVING COUNT(*)>20000 AND SUM(channel='UNKNOWN')>0
ORDER BY "UNKNOWN 占比%" DESC LIMIT 10
""")

save("t27_top_types", f"""
SELECT complaint_type AS 问题类型, any_value(agency) AS 主责机构, COUNT(*) AS 记录数,
  ROUND(100.0*COUNT(*)/{N_ANA}.0,2) AS "占全量%",
  ROUND(100.0*SUM(COUNT(*)) OVER (ORDER BY COUNT(*) DESC)/{N_ANA}.0,2) AS "累计占比%"
FROM sr GROUP BY 1 ORDER BY 记录数 DESC LIMIT 25
""")

print("\n全部汇总表已生成于:", OUT)
