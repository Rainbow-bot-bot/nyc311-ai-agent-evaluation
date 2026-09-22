import duckdb
import pandas as pd

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

# Let's inspect daily record count distribution
print("--- Daily volume summary ---")
res_daily = db.execute(f"""
    SELECT 
        CAST(TRY_CAST(created_date AS TIMESTAMP) AS DATE) as c_date,
        COUNT(*) as cnt
    FROM read_parquet('{parquet_path}')
    GROUP BY 1
    ORDER BY 1
""").df()

print(f"Total days: {len(res_daily)}")
print(f"First 5 days:\n{res_daily.head(5)}")
print(f"Last 5 days:\n{res_daily.tail(5)}")

# Let's check Year 1 vs Year 2 on a clean 11-month basis (Oct to Aug)
# and also on full matched 364-day basis
res_11m = db.execute(f"""
    SELECT 
        CASE 
            WHEN created_date >= '2024-10-01' AND created_date < '2025-09-01' THEN 'Year1 (2024-10 to 2025-08)'
            WHEN created_date >= '2025-10-01' AND created_date < '2026-09-01' THEN 'Year2 (2025-10 to 2026-08)'
            ELSE 'Other'
        END as period,
        COUNT(*) as cnt
    FROM read_parquet('{parquet_path}')
    GROUP BY 1
    ORDER BY 1
""").df()
print("\n--- 11-Month Matched Period Volume ---")
print(res_11m)

# Month-by-month comparison (e.g. Oct 24 vs Oct 25, Nov 24 vs Nov 25, ..., Aug 24 vs Aug 25)
res_m_comp = db.execute(f"""
    WITH m AS (
        SELECT 
            SUBSTRING(created_date, 1, 7) as ym,
            SUBSTRING(created_date, 6, 2) as m,
            SUBSTRING(created_date, 1, 4) as y,
            COUNT(*) as cnt
        FROM read_parquet('{parquet_path}')
        WHERE created_date >= '2024-10-01' AND created_date < '2026-09-01'
        GROUP BY 1, 2, 3
    )
    SELECT 
        y1.m as month_num,
        y1.ym as y1_month,
        y1.cnt as y1_count,
        y2.ym as y2_month,
        y2.cnt as y2_count,
        y2.cnt - y1.cnt as diff,
        ROUND((y2.cnt - y1.cnt) * 100.0 / y1.cnt, 2) as growth_pct
    FROM m y1
    JOIN m y2 ON y1.m = y2.m AND y1.y = '2024' AND y2.y = '2025'
    UNION ALL
    SELECT 
        y1.m as month_num,
        y1.ym as y1_month,
        y1.cnt as y1_count,
        y2.ym as y2_month,
        y2.cnt as y2_count,
        y2.cnt - y1.cnt as diff,
        ROUND((y2.cnt - y1.cnt) * 100.0 / y1.cnt, 2) as growth_pct
    FROM m y1
    JOIN m y2 ON y1.m = y2.m AND y1.y = '2025' AND y2.y = '2026'
    ORDER BY month_num
""").df()
print("\n--- Month-over-Month Comparison ---")
print(res_m_comp.to_string(index=False))

# Let's save daily and monthly comparison
res_daily.to_csv("data/daily_volume.csv", index=False)
res_m_comp.to_csv("data/month_yoy_comparison.csv", index=False)
