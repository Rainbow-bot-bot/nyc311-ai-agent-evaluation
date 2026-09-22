"""步骤1：摸底 NYC 311 原始 Parquet 的字段、粒度、时间范围与缺失情况。只读访问原始数据。"""
import duckdb, os, sys

RAW = r"D:\项目1\原始数据"
GLOB = os.path.join(RAW, "*.parquet").replace("\\", "/")

con = duckdb.connect()
con.execute("PRAGMA threads=8")

print("=== 1. 字段与类型 ===")
schema = con.execute(f"DESCRIBE SELECT * FROM read_parquet('{GLOB}')").fetchall()
for i, r in enumerate(schema, 1):
    print(f"{i:3d}. {r[0]:<40s} {r[1]}")

print("\n=== 2. 总行数 / unique_key 去重 ===")
print(con.execute(f"""
SELECT COUNT(*) AS rows,
       COUNT(DISTINCT unique_key) AS distinct_unique_key
FROM read_parquet('{GLOB}')
""").fetchdf().to_string(index=False))

print("\n=== 3. created_date 时间范围 ===")
print(con.execute(f"""
SELECT MIN(created_date) AS min_created, MAX(created_date) AS max_created,
       MIN(closed_date) AS min_closed, MAX(closed_date) AS max_closed
FROM read_parquet('{GLOB}')
""").fetchdf().to_string(index=False))

print("\n=== 4. 每个分片的行数与实际时间范围 ===")
print(con.execute(f"""
SELECT regexp_extract(filename, '[^/\\\\]+$') AS f,
       COUNT(*) AS n,
       MIN(created_date)::DATE AS d_min, MAX(created_date)::DATE AS d_max
FROM read_parquet('{GLOB}', filename=true)
GROUP BY 1 ORDER BY 1
""").fetchdf().to_string(index=False))

print("\n=== 5. 各字段缺失率 ===")
cols = [r[0] for r in schema]
parts = []
for c in cols:
    parts.append(f"""SELECT '{c}' AS col, COUNT(*) - COUNT("{c}") AS nulls,
       ROUND(100.0*(COUNT(*)-COUNT("{c}"))/COUNT(*),2) AS null_pct,
       COUNT(DISTINCT "{c}") AS n_distinct
FROM read_parquet('{GLOB}')""")
q = "\nUNION ALL\n".join(parts) + "\nORDER BY null_pct DESC"
print(con.execute(q).fetchdf().to_string(index=False))
