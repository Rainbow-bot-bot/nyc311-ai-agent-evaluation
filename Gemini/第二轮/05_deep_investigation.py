import duckdb
import pandas as pd

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

print("--- 1. Street Condition Descriptors (DOT) YoY & March 2026 ---")
res_street_desc = db.execute(f"""
    SELECT 
        descriptor,
        COUNT(CASE WHEN created_date >= '2024-10-01' AND created_date < '2025-09-01' THEN 1 END) as y1_cnt,
        COUNT(CASE WHEN created_date >= '2025-10-01' AND created_date < '2026-09-01' THEN 1 END) as y2_cnt,
        COUNT(CASE WHEN created_date >= '2026-03-01' AND created_date < '2026-04-01' THEN 1 END) as mar26_cnt,
        COUNT(CASE WHEN created_date >= '2025-03-01' AND created_date < '2025-04-01' THEN 1 END) as mar25_cnt
    FROM read_parquet('{parquet_path}')
    WHERE complaint_type = 'Street Condition'
    GROUP BY descriptor
    ORDER BY y2_cnt DESC
    LIMIT 10
""").df()
print(res_street_desc.to_string(index=False))

print("\n--- 2. DEP Taxonomy Shift: Water System vs Water Maintenance, Sewer vs Sewer Maintenance ---")
res_dep_taxonomy = db.execute(f"""
    SELECT 
        SUBSTRING(created_date, 1, 7) as ym,
        COUNT(CASE WHEN complaint_type = 'Water System' THEN 1 END) as water_system,
        COUNT(CASE WHEN complaint_type = 'Water Maintenance' THEN 1 END) as water_maint,
        COUNT(CASE WHEN complaint_type = 'Sewer' THEN 1 END) as sewer,
        COUNT(CASE WHEN complaint_type = 'Sewer Maintenance' THEN 1 END) as sewer_maint
    FROM read_parquet('{parquet_path}')
    WHERE agency = 'DEP'
    GROUP BY 1
    ORDER BY 1
""").df()
print(res_dep_taxonomy.to_string(index=False))

print("\n--- 3. February 2026 Daily Timeline of Snow and Heat ---")
res_feb_daily = db.execute(f"""
    SELECT 
        CAST(TRY_CAST(created_date AS TIMESTAMP) AS DATE) as c_date,
        COUNT(CASE WHEN complaint_type = 'Snow or Ice' THEN 1 END) as snow_cnt,
        COUNT(CASE WHEN complaint_type = 'HEAT/HOT WATER' THEN 1 END) as heat_cnt,
        COUNT(CASE WHEN complaint_type = 'Blocked Driveway' THEN 1 END) as blocked_cnt,
        COUNT(*) as total_day_cnt
    FROM read_parquet('{parquet_path}')
    WHERE created_date >= '2026-02-01' AND created_date < '2026-03-01'
    GROUP BY 1
    ORDER BY 1
""").df()
print(res_feb_daily.to_string(index=False))

print("\n--- 4. Channel Shift YoY ---")
p1_filter = "created_date >= '2024-10-01' AND created_date < '2025-09-01'"
p2_filter = "created_date >= '2025-10-01' AND created_date < '2026-09-01'"
res_channel_yoy = db.execute(f"""
    WITH p1 AS (
        SELECT open_data_channel_type, COUNT(*) as y1_cnt
        FROM read_parquet('{parquet_path}')
        WHERE {p1_filter}
        GROUP BY 1
    ),
    p2 AS (
        SELECT open_data_channel_type, COUNT(*) as y2_cnt
        FROM read_parquet('{parquet_path}')
        WHERE {p2_filter}
        GROUP BY 1
    )
    SELECT 
        COALESCE(p1.open_data_channel_type, p2.open_data_channel_type) as channel,
        y1_cnt,
        ROUND(y1_cnt * 100.0 / 3293003, 2) as y1_share_pct,
        y2_cnt,
        ROUND(y2_cnt * 100.0 / 3639698, 2) as y2_share_pct,
        y2_cnt - y1_cnt as diff,
        ROUND((y2_cnt - y1_cnt) * 100.0 / y1_cnt, 2) as growth_pct
    FROM p1 FULL OUTER JOIN p2 ON p1.open_data_channel_type = p2.open_data_channel_type
    ORDER BY y2_cnt DESC
""").df()
print(res_channel_yoy.to_string(index=False))

res_street_desc.to_csv("data/street_condition_desc.csv", index=False)
res_dep_taxonomy.to_csv("data/dep_taxonomy_shift.csv", index=False)
res_feb_daily.to_csv("data/feb26_daily_timeline.csv", index=False)
res_channel_yoy.to_csv("data/channel_shift_yoy.csv", index=False)
