import duckdb
import os

db = duckdb.connect()

data_path = "D:/项目1/原始数据/*.parquet"

print("Querying total count and date ranges...")
res = db.execute(f"""
    SELECT 
        COUNT(*) as total_records,
        MIN(Created_Date) as min_created,
        MAX(Created_Date) as max_created,
        MIN(Closed_Date) as min_closed,
        MAX(Closed_Date) as max_closed
    FROM read_parquet('{data_path}')
""").df()
print(res)

print("\nColumns and Types:")
cols = db.execute(f"DESCRIBE SELECT * FROM read_parquet('{data_path}') LIMIT 1").df()
for idx, row in cols.iterrows():
    print(f"{row['column_name']}: {row['column_type']}")
