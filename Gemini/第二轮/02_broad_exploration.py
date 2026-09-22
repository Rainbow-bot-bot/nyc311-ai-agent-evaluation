import duckdb
import pandas as pd

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

print("--- 1. Agency Distribution ---")
res_agency = db.execute(f"""
    SELECT 
        agency,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / 7525498, 2) as pct,
        COUNT(DISTINCT complaint_type) as distinct_complaint_types
    FROM read_parquet('{parquet_path}')
    GROUP BY agency
    ORDER BY cnt DESC
""").df()
print(res_agency.to_string(index=False))

print("\n--- 2. Top 25 Complaint Types Overall ---")
res_ct = db.execute(f"""
    SELECT 
        complaint_type,
        agency,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / 7525498, 2) as pct
    FROM read_parquet('{parquet_path}')
    GROUP BY complaint_type, agency
    ORDER BY cnt DESC
    LIMIT 25
""").df()
print(res_ct.to_string(index=False))

print("\n--- 3. Channel Type Distribution ---")
res_channel = db.execute(f"""
    SELECT 
        open_data_channel_type,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / 7525498, 2) as pct
    FROM read_parquet('{parquet_path}')
    GROUP BY open_data_channel_type
    ORDER BY cnt DESC
""").df()
print(res_channel.to_string(index=False))

res_agency.to_csv("data/agency_dist.csv", index=False)
res_ct.to_csv("data/top25_complaint_types.csv", index=False)
res_channel.to_csv("data/channel_dist.csv", index=False)
