import duckdb
import pandas as pd

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

p1_filter = "created_date >= '2024-10-01' AND created_date < '2025-09-01'"
p2_filter = "created_date >= '2025-10-01' AND created_date < '2026-09-01'"

print("--- 1. Borough Growth YoY ---")
res_borough_yoy = db.execute(f"""
    WITH p1 AS (
        SELECT borough, COUNT(*) as y1_cnt
        FROM read_parquet('{parquet_path}')
        WHERE {p1_filter}
        GROUP BY 1
    ),
    p2 AS (
        SELECT borough, COUNT(*) as y2_cnt
        FROM read_parquet('{parquet_path}')
        WHERE {p2_filter}
        GROUP BY 1
    )
    SELECT 
        COALESCE(p1.borough, p2.borough) as borough,
        y1_cnt,
        ROUND(y1_cnt * 100.0 / 3293003, 2) as y1_share_pct,
        y2_cnt,
        ROUND(y2_cnt * 100.0 / 3639698, 2) as y2_share_pct,
        y2_cnt - y1_cnt as diff,
        ROUND((y2_cnt - y1_cnt) * 100.0 / y1_cnt, 2) as growth_pct,
        ROUND((y2_cnt - y1_cnt) * 100.0 / 346695, 2) as contribution_pct
    FROM p1 FULL OUTER JOIN p2 ON p1.borough = p2.borough
    ORDER BY y2_cnt DESC
""").df()
print(res_borough_yoy.to_string(index=False))

print("\n--- 2. Resolution Time (Hours) by Agency (Excluding right-censored recent month, look at Oct-Apr to avoid truncation) ---")
# To avoid right censoring, let's look at records created between 2024-10-01 and 2025-04-30 (Year 1)
# vs 2025-10-01 and 2026-04-30 (Year 2), giving at least 4-5 months of observation window for closing!
# Also filter out closed_date < created_date anomalies
res_resp_time = db.execute(f"""
    WITH clean_cases AS (
        SELECT 
            agency,
            CASE 
                WHEN created_date >= '2024-10-01' AND created_date < '2025-05-01' THEN 'Y1_Winter_Spring'
                WHEN created_date >= '2025-10-01' AND created_date < '2026-05-01' THEN 'Y2_Winter_Spring'
            END as period,
            epoch(TRY_CAST(closed_date AS TIMESTAMP)) - epoch(TRY_CAST(created_date AS TIMESTAMP)) as duration_sec,
            status
        FROM read_parquet('{parquet_path}')
        WHERE (
            (created_date >= '2024-10-01' AND created_date < '2025-05-01') OR
            (created_date >= '2025-10-01' AND created_date < '2026-05-01')
        )
    )
    SELECT 
        agency,
        period,
        COUNT(*) as total_cases,
        COUNT(CASE WHEN status = 'Closed' AND duration_sec >= 0 THEN 1 END) as closed_valid_cases,
        ROUND(COUNT(CASE WHEN status = 'Closed' AND duration_sec >= 0 THEN 1 END) * 100.0 / COUNT(*), 2) as close_rate_pct,
        ROUND(MEDIAN(CASE WHEN status = 'Closed' AND duration_sec >= 0 THEN duration_sec / 3600.0 END), 2) as median_hours,
        ROUND(AVG(CASE WHEN status = 'Closed' AND duration_sec >= 0 THEN duration_sec / 3600.0 END), 2) as mean_hours
    FROM clean_cases
    WHERE agency IN ('NYPD', 'HPD', 'DSNY', 'DOT', 'DEP', 'DPR', 'DOB')
    GROUP BY agency, period
    ORDER BY agency, period
""").df()
print(res_resp_time.to_string(index=False))

print("\n--- 3. Top Address Concentration in HPD Heat/Hot Water ---")
res_heat_conc = db.execute(f"""
    SELECT 
        COUNT(*) as total_heat_records,
        COUNT(DISTINCT incident_address) as distinct_addresses,
        ROUND(COUNT(*) * 1.0 / COUNT(DISTINCT incident_address), 2) as avg_complaints_per_addr
    FROM read_parquet('{parquet_path}')
    WHERE complaint_type = 'HEAT/HOT WATER' AND incident_address IS NOT NULL
""").df()
print(res_heat_conc)

res_heat_top_addr = db.execute(f"""
    SELECT 
        incident_address,
        borough,
        COUNT(*) as heat_complaints
    FROM read_parquet('{parquet_path}')
    WHERE complaint_type = 'HEAT/HOT WATER' AND incident_address IS NOT NULL
    GROUP BY 1, 2
    ORDER BY heat_complaints DESC
    LIMIT 10
""").df()
print("\nTop 10 heat complaint addresses:")
print(res_heat_top_addr.to_string(index=False))

res_borough_yoy.to_csv("data/borough_growth_yoy.csv", index=False)
res_resp_time.to_csv("data/agency_resolution_time.csv", index=False)
res_heat_top_addr.to_csv("data/heat_top_addresses.csv", index=False)
