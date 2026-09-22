import duckdb
import pandas as pd
import json
import os

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

print("--- 1. Unique Key & Record Grain ---")
res_grain = db.execute(f"""
    SELECT 
        COUNT(*) as total_rows,
        COUNT(DISTINCT unique_key) as unique_keys,
        COUNT(*) - COUNT(DISTINCT unique_key) as duplicate_keys
    FROM read_parquet('{parquet_path}')
""").df()
print(res_grain)

print("\n--- 2. Missing Value Analysis ---")
res_missing = db.execute(f"""
    SELECT
        COUNT(*) as total,
        COUNT(unique_key) as count_unique_key,
        COUNT(created_date) as count_created,
        COUNT(closed_date) as count_closed,
        COUNT(agency) as count_agency,
        COUNT(complaint_type) as count_complaint_type,
        COUNT(descriptor) as count_descriptor,
        COUNT(location_type) as count_location_type,
        COUNT(incident_zip) as count_zip,
        COUNT(status) as count_status,
        COUNT(borough) as count_borough,
        COUNT(open_data_channel_type) as count_channel,
        COUNT(latitude) as count_lat,
        COUNT(longitude) as count_lon,
        COUNT(resolution_description) as count_resolution
    FROM read_parquet('{parquet_path}')
""").df()
missing_df = pd.DataFrame({
    'field': res_missing.columns[1:],
    'non_null_count': res_missing.iloc[0, 1:].values,
})
missing_df['missing_count'] = res_missing.iloc[0, 0] - missing_df['non_null_count']
missing_df['missing_pct'] = (missing_df['missing_count'] / res_missing.iloc[0, 0]) * 100
print(missing_df.to_string(index=False))

print("\n--- 3. Monthly Volume & Trend ---")
res_monthly = db.execute(f"""
    SELECT 
        SUBSTRING(created_date, 1, 7) as year_month,
        COUNT(*) as record_count,
        COUNT(closed_date) as closed_count,
        COUNT(CASE WHEN status = 'Closed' THEN 1 END) as status_closed_count
    FROM read_parquet('{parquet_path}')
    GROUP BY 1
    ORDER BY 1
""").df()
print(res_monthly.to_string(index=False))

print("\n--- 4. Date Anomalies & Right Censoring ---")
res_anomalies = db.execute(f"""
    SELECT 
        COUNT(CASE WHEN closed_date < created_date THEN 1 END) as closed_before_created,
        COUNT(CASE WHEN TRY_CAST(closed_date AS TIMESTAMP) > TIMESTAMP '2026-09-07 23:59:59' THEN 1 END) as future_closed,
        COUNT(CASE WHEN TRY_CAST(created_date AS TIMESTAMP) > TIMESTAMP '2026-09-07 23:59:59' THEN 1 END) as future_created,
        COUNT(CASE WHEN status = 'Closed' AND closed_date IS NULL THEN 1 END) as closed_status_null_date,
        COUNT(CASE WHEN status != 'Closed' AND closed_date IS NOT NULL THEN 1 END) as non_closed_has_date
    FROM read_parquet('{parquet_path}')
""").df()
print(res_anomalies)

print("\n--- 5. Borough Distribution ---")
res_borough = db.execute(f"""
    SELECT 
        COALESCE(borough, 'NULL') as borough,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as pct
    FROM read_parquet('{parquet_path}')
    GROUP BY 1
    ORDER BY cnt DESC
""").df()
print(res_borough.to_string(index=False))

print("\n--- 6. Status Distribution ---")
res_status = db.execute(f"""
    SELECT 
        COALESCE(status, 'NULL') as status,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as pct
    FROM read_parquet('{parquet_path}')
    GROUP BY 1
    ORDER BY cnt DESC
""").df()
print(res_status.to_string(index=False))

os.makedirs("data", exist_ok=True)
res_monthly.to_csv("data/monthly_volume.csv", index=False)
missing_df.to_csv("data/missing_analysis.csv", index=False)
res_borough.to_csv("data/borough_dist.csv", index=False)
res_status.to_csv("data/status_dist.csv", index=False)
