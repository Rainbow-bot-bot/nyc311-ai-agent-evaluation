import duckdb
import pandas as pd
import numpy as np
import sys
sys.stdout.reconfigure(encoding='utf-8')

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

# 1. Day of week and Hour of Day distribution
print("--- 1. Day of Week & Hour Distribution ---")
res_hourly = db.execute(f"""
    SELECT 
        EXTRACT(hour FROM TRY_CAST(created_date AS TIMESTAMP)) as hour_of_day,
        COUNT(*) as total_cnt,
        COUNT(CASE WHEN complaint_type ILIKE '%Noise%' THEN 1 END) as noise_cnt,
        COUNT(CASE WHEN complaint_type = 'Illegal Parking' THEN 1 END) as parking_cnt,
        COUNT(CASE WHEN complaint_type = 'HEAT/HOT WATER' THEN 1 END) as heat_cnt
    FROM read_parquet('{parquet_path}')
    GROUP BY 1
    ORDER BY 1
""").df()
res_hourly.to_csv("data/hourly_distribution.csv", index=False)

res_dow = db.execute(f"""
    SELECT 
        DAYNAME(TRY_CAST(created_date AS TIMESTAMP)) as day_of_week,
        EXTRACT(dow FROM TRY_CAST(created_date AS TIMESTAMP)) as dow_num,
        COUNT(*) as total_cnt,
        COUNT(CASE WHEN complaint_type ILIKE '%Noise%' THEN 1 END) as noise_cnt,
        COUNT(CASE WHEN complaint_type = 'Illegal Parking' THEN 1 END) as parking_cnt
    FROM read_parquet('{parquet_path}')
    GROUP BY 1, 2
    ORDER BY dow_num
""").df()
res_dow.to_csv("data/dow_distribution.csv", index=False)

# 2. Borough breakdown by Top Complaint Types
print("--- 2. Borough Breakdown by Complaint Type ---")
res_borough_ct = db.execute(f"""
    SELECT 
        borough,
        COUNT(*) as total_complaints,
        COUNT(CASE WHEN complaint_type = 'Illegal Parking' THEN 1 END) as illegal_parking,
        COUNT(CASE WHEN complaint_type = 'Noise - Residential' THEN 1 END) as noise_residential,
        COUNT(CASE WHEN complaint_type = 'HEAT/HOT WATER' THEN 1 END) as heat_hot_water,
        COUNT(CASE WHEN complaint_type = 'Blocked Driveway' THEN 1 END) as blocked_driveway,
        COUNT(CASE WHEN complaint_type = 'Street Condition' THEN 1 END) as street_condition
    FROM read_parquet('{parquet_path}')
    WHERE borough != 'Unspecified'
    GROUP BY 1
    ORDER BY total_complaints DESC
""").df()
print(res_borough_ct)
res_borough_ct.to_csv("data/borough_by_complaint_type.csv", index=False)

# 3. Zip code level analysis for Heat/Hot Water
print("--- 3. Top 15 Zip Codes for Heat/Hot Water ---")
res_zip_heat = db.execute(f"""
    SELECT 
        incident_zip,
        borough,
        COUNT(*) as heat_cnt,
        COUNT(DISTINCT incident_address) as distinct_addrs,
        ROUND(COUNT(*) * 1.0 / COUNT(DISTINCT incident_address), 1) as complaints_per_addr
    FROM read_parquet('{parquet_path}')
    WHERE complaint_type = 'HEAT/HOT WATER' AND incident_zip IS NOT NULL
    GROUP BY 1, 2
    ORDER BY heat_cnt DESC
    LIMIT 15
""").df()
print(res_zip_heat)
res_zip_heat.to_csv("data/zip_heat_top15.csv", index=False)

# 4. Check right-censoring across monthly cohorts
print("--- 4. Resolution Rate by Monthly Cohort ---")
res_cohort = db.execute(f"""
    SELECT 
        SUBSTRING(created_date, 1, 7) as cohort_month,
        COUNT(*) as created_cnt,
        COUNT(CASE WHEN status = 'Closed' THEN 1 END) as closed_cnt,
        ROUND(COUNT(CASE WHEN status = 'Closed' THEN 1 END) * 100.0 / COUNT(*), 2) as closed_rate_pct,
        ROUND(MEDIAN(CASE WHEN status = 'Closed' AND TRY_CAST(closed_date AS TIMESTAMP) >= TRY_CAST(created_date AS TIMESTAMP) 
            THEN (epoch(TRY_CAST(closed_date AS TIMESTAMP)) - epoch(TRY_CAST(created_date AS TIMESTAMP))) / 3600.0 END), 2) as median_hours_closed
    FROM read_parquet('{parquet_path}')
    GROUP BY 1
    ORDER BY 1
""").df()
print(res_cohort.to_string(index=False))
res_cohort.to_csv("data/monthly_cohort_resolution.csv", index=False)
