import duckdb
import pandas as pd

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

# Define periods
p1_filter = "created_date >= '2024-10-01' AND created_date < '2025-09-01'"
p2_filter = "created_date >= '2025-10-01' AND created_date < '2026-09-01'"

print("--- 1. Growth Contribution by Agency ---")
res_agency_growth = db.execute(f"""
    WITH p1 AS (
        SELECT agency, COUNT(*) as cnt_p1
        FROM read_parquet('{parquet_path}')
        WHERE {p1_filter}
        GROUP BY agency
    ),
    p2 AS (
        SELECT agency, COUNT(*) as cnt_p2
        FROM read_parquet('{parquet_path}')
        WHERE {p2_filter}
        GROUP BY agency
    )
    SELECT 
        COALESCE(p1.agency, p2.agency) as agency,
        COALESCE(p1.cnt_p1, 0) as cnt_y1,
        COALESCE(p2.cnt_p2, 0) as cnt_y2,
        COALESCE(p2.cnt_p2, 0) - COALESCE(p1.cnt_p1, 0) as diff,
        ROUND((COALESCE(p2.cnt_p2, 0) - COALESCE(p1.cnt_p1, 0)) * 100.0 / NULLIF(p1.cnt_p1, 0), 2) as growth_pct,
        ROUND((COALESCE(p2.cnt_p2, 0) - COALESCE(p1.cnt_p1, 0)) * 100.0 / 346695.0, 2) as contribution_pct
    FROM p1
    FULL OUTER JOIN p2 ON p1.agency = p2.agency
    ORDER BY diff DESC
""").df()
print(res_agency_growth.to_string(index=False))

print("\n--- 2. Growth Contribution by Top Complaint Types ---")
res_ct_growth = db.execute(f"""
    WITH p1 AS (
        SELECT complaint_type, agency, COUNT(*) as cnt_p1
        FROM read_parquet('{parquet_path}')
        WHERE {p1_filter}
        GROUP BY complaint_type, agency
    ),
    p2 AS (
        SELECT complaint_type, agency, COUNT(*) as cnt_p2
        FROM read_parquet('{parquet_path}')
        WHERE {p2_filter}
        GROUP BY complaint_type, agency
    )
    SELECT 
        COALESCE(p1.complaint_type, p2.complaint_type) as complaint_type,
        COALESCE(p1.agency, p2.agency) as agency,
        COALESCE(p1.cnt_p1, 0) as cnt_y1,
        COALESCE(p2.cnt_p2, 0) as cnt_y2,
        COALESCE(p2.cnt_p2, 0) - COALESCE(p1.cnt_p1, 0) as diff,
        ROUND((COALESCE(p2.cnt_p2, 0) - COALESCE(p1.cnt_p1, 0)) * 100.0 / NULLIF(p1.cnt_p1, 0), 2) as growth_pct,
        ROUND((COALESCE(p2.cnt_p2, 0) - COALESCE(p1.cnt_p1, 0)) * 100.0 / 346695.0, 2) as contribution_pct
    FROM p1
    FULL OUTER JOIN p2 ON p1.complaint_type = p2.complaint_type AND p1.agency = p2.agency
    ORDER BY diff DESC
    LIMIT 20
""").df()
print(res_ct_growth.to_string(index=False))

print("\n--- 3. What happened in Feb 2026 vs Feb 2025? ---")
res_feb = db.execute(f"""
    WITH feb25 AS (
        SELECT complaint_type, agency, COUNT(*) as cnt_25
        FROM read_parquet('{parquet_path}')
        WHERE created_date >= '2025-02-01' AND created_date < '2025-03-01'
        GROUP BY complaint_type, agency
    ),
    feb26 AS (
        SELECT complaint_type, agency, COUNT(*) as cnt_26
        FROM read_parquet('{parquet_path}')
        WHERE created_date >= '2026-02-01' AND created_date < '2026-03-01'
        GROUP BY complaint_type, agency
    )
    SELECT 
        COALESCE(f25.complaint_type, f26.complaint_type) as complaint_type,
        COALESCE(f25.agency, f26.agency) as agency,
        COALESCE(f25.cnt_25, 0) as feb25,
        COALESCE(f26.cnt_26, 0) as feb26,
        COALESCE(f26.cnt_26, 0) - COALESCE(f25.cnt_25, 0) as diff,
        ROUND((COALESCE(f26.cnt_26, 0) - COALESCE(f25.cnt_25, 0)) * 100.0 / 79326.0, 2) as contribution_to_feb_spike_pct
    FROM feb25 f25
    FULL OUTER JOIN feb26 f26 ON f25.complaint_type = f26.complaint_type AND f25.agency = f26.agency
    ORDER BY diff DESC
    LIMIT 15
""").df()
print(res_feb.to_string(index=False))

print("\n--- 4. What happened in March 2026 vs March 2025? ---")
res_mar = db.execute(f"""
    WITH mar25 AS (
        SELECT complaint_type, agency, COUNT(*) as cnt_25
        FROM read_parquet('{parquet_path}')
        WHERE created_date >= '2025-03-01' AND created_date < '2025-04-01'
        GROUP BY complaint_type, agency
    ),
    mar26 AS (
        SELECT complaint_type, agency, COUNT(*) as cnt_26
        FROM read_parquet('{parquet_path}')
        WHERE created_date >= '2026-03-01' AND created_date < '2026-04-01'
        GROUP BY complaint_type, agency
    )
    SELECT 
        COALESCE(m25.complaint_type, m26.complaint_type) as complaint_type,
        COALESCE(m25.agency, m26.agency) as agency,
        COALESCE(m25.cnt_25, 0) as mar25,
        COALESCE(m26.cnt_26, 0) as mar26,
        COALESCE(m26.cnt_26, 0) - COALESCE(m25.cnt_25, 0) as diff,
        ROUND((COALESCE(m26.cnt_26, 0) - COALESCE(m25.cnt_25, 0)) * 100.0 / 61166.0, 2) as contribution_to_mar_spike_pct
    FROM mar25 m25
    FULL OUTER JOIN mar26 m26 ON m25.complaint_type = m26.complaint_type AND m25.agency = m26.agency
    ORDER BY diff DESC
    LIMIT 15
""").df()
print(res_mar.to_string(index=False))

res_agency_growth.to_csv("data/agency_growth.csv", index=False)
res_ct_growth.to_csv("data/complaint_type_growth.csv", index=False)
res_feb.to_csv("data/feb_spike_decomp.csv", index=False)
res_mar.to_csv("data/mar_spike_decomp.csv", index=False)
