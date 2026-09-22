"""Follow-up diagnostics: mix shift, spikes, recodes."""
from pathlib import Path
import duckdb

RAW = Path(r"D:\项目1\原始数据")
OUT = Path(r"D:\项目2\Grok\summaries")
GLOB = str(RAW / "*.parquet").replace("\\", "/")

con = duckdb.connect()
con.execute("SET threads TO 4")
con.execute(f"""
CREATE OR REPLACE VIEW src AS
SELECT
  unique_key,
  try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f') AS created_ts,
  try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f') AS closed_ts,
  agency, complaint_type, descriptor, status, borough,
  open_data_channel_type AS channel
FROM read_parquet('{GLOB}')
""")

# Jan 1-10 2025 vs 2026 daily noise vs heat vs all
jan_early = con.execute("""
SELECT CAST(created_ts AS DATE) AS d,
  COUNT(*) AS n_all,
  COUNT(*) FILTER (WHERE complaint_type = 'Noise - Residential') AS n_noise,
  COUNT(*) FILTER (WHERE complaint_type = 'HEAT/HOT WATER') AS n_heat,
  COUNT(*) FILTER (WHERE complaint_type = 'Illegal Parking') AS n_park
FROM src
WHERE (created_ts >= TIMESTAMP '2024-12-28' AND created_ts < TIMESTAMP '2025-01-15')
   OR (created_ts >= TIMESTAMP '2025-12-28' AND created_ts < TIMESTAMP '2026-01-15')
GROUP BY 1 ORDER BY 1
""").df()
jan_early.to_csv(OUT / "diag_newyear_window.csv", index=False)

# Feb 24 2026 composition
feb24 = con.execute("""
SELECT COALESCE(complaint_type,'(blank)') AS complaint_type, COUNT(*) AS n
FROM src
WHERE CAST(created_ts AS DATE) = DATE '2026-02-24'
GROUP BY 1 ORDER BY n DESC LIMIT 20
""").df()
feb24.to_csv(OUT / "diag_2026_02_24_types.csv", index=False)

# nearby days for context
feb_week = con.execute("""
SELECT CAST(created_ts AS DATE) AS d,
  COUNT(*) AS n_all,
  COUNT(*) FILTER (WHERE complaint_type = 'HEAT/HOT WATER') AS n_heat,
  COUNT(*) FILTER (WHERE complaint_type = 'Snow or Ice') AS n_snow,
  COUNT(*) FILTER (WHERE complaint_type = 'Illegal Parking') AS n_park,
  COUNT(*) FILTER (WHERE complaint_type ILIKE '%Noise%') AS n_noise
FROM src
WHERE created_ts >= TIMESTAMP '2026-02-15' AND created_ts < TIMESTAMP '2026-03-05'
GROUP BY 1 ORDER BY 1
""").df()
feb_week.to_csv(OUT / "diag_2026_feb_window.csv", index=False)

# snow/ice monthly
snow = con.execute("""
SELECT strftime(created_ts,'%Y-%m') AS m, COUNT(*) AS n
FROM src WHERE complaint_type = 'Snow or Ice'
GROUP BY 1 ORDER BY 1
""").df()
snow.to_csv(OUT / "diag_snow_monthly.csv", index=False)

# YoY totals
yoy_tot = con.execute("""
SELECT
  COUNT(*) FILTER (WHERE created_ts >= TIMESTAMP '2024-10-01' AND created_ts < TIMESTAMP '2025-09-01') AS n_y1,
  COUNT(*) FILTER (WHERE created_ts >= TIMESTAMP '2025-10-01' AND created_ts < TIMESTAMP '2026-09-01') AS n_y2
FROM src
""").df()
yoy_tot.to_csv(OUT / "diag_yoy_totals.csv", index=False)

# contribution of heat+snow+parking+street to YoY delta
contrib = con.execute("""
WITH t AS (
  SELECT
    CASE
      WHEN complaint_type IN ('HEAT/HOT WATER','Snow or Ice','Illegal Parking','Street Condition','Blocked Driveway')
        THEN complaint_type
      ELSE 'Other'
    END AS grp,
    CASE
      WHEN created_ts >= TIMESTAMP '2024-10-01' AND created_ts < TIMESTAMP '2025-09-01' THEN 'y1'
      WHEN created_ts >= TIMESTAMP '2025-10-01' AND created_ts < TIMESTAMP '2026-09-01' THEN 'y2'
    END AS period
  FROM src
)
SELECT grp,
  COUNT(*) FILTER (WHERE period='y1') AS n_y1,
  COUNT(*) FILTER (WHERE period='y2') AS n_y2,
  COUNT(*) FILTER (WHERE period='y2') - COUNT(*) FILTER (WHERE period='y1') AS delta
FROM t WHERE period IS NOT NULL
GROUP BY 1 ORDER BY delta DESC
""").df()
contrib.to_csv(OUT / "diag_yoy_contrib.csv", index=False)

# Water System vs Water Maintenance monthly - recode check
water = con.execute("""
SELECT strftime(created_ts,'%Y-%m') AS m,
  COUNT(*) FILTER (WHERE complaint_type = 'Water System') AS n_water_system,
  COUNT(*) FILTER (WHERE complaint_type = 'Water Maintenance') AS n_water_maint,
  COUNT(*) FILTER (WHERE complaint_type = 'Sewer') AS n_sewer,
  COUNT(*) FILTER (WHERE complaint_type = 'Sewer Maintenance') AS n_sewer_maint
FROM src GROUP BY 1 ORDER BY 1
""").df()
water.to_csv(OUT / "diag_water_recode.csv", index=False)

# HPD heat close hours by winter month (closed valid), exclude last 2 months for censoring
heat_close_m = con.execute("""
SELECT strftime(created_ts,'%Y-%m') AS m,
  COUNT(*) AS n,
  COUNT(*) FILTER (WHERE status='Closed' AND closed_ts>=created_ts) AS n_closed,
  MEDIAN(epoch(closed_ts-created_ts)/3600.0) FILTER (WHERE status='Closed' AND closed_ts>=created_ts) AS med_h,
  QUANTILE_CONT(epoch(closed_ts-created_ts)/3600.0, 0.9) FILTER (WHERE status='Closed' AND closed_ts>=created_ts) AS p90_h
FROM src
WHERE complaint_type='HEAT/HOT WATER'
GROUP BY 1 ORDER BY 1
""").df()
heat_close_m.to_csv(OUT / "diag_heat_close_monthly.csv", index=False)

# borough share of heat vs all, complete months Oct24-Aug26
heat_boro_share = con.execute("""
SELECT COALESCE(borough,'(blank)') AS borough,
  COUNT(*) FILTER (WHERE complaint_type='HEAT/HOT WATER') AS n_heat,
  COUNT(*) AS n_all,
  COUNT(*) FILTER (WHERE complaint_type='Illegal Parking') AS n_park,
  COUNT(*) FILTER (WHERE complaint_type='Noise - Residential') AS n_noise
FROM src
WHERE created_ts >= TIMESTAMP '2024-10-01' AND created_ts < TIMESTAMP '2026-09-01'
  AND borough NOT IN ('Unspecified')
GROUP BY 1 ORDER BY n_all DESC
""").df()
heat_boro_share.to_csv(OUT / "diag_boro_type_shares.csv", index=False)

# channel monthly share
ch = con.execute("""
SELECT strftime(created_ts,'%Y-%m') AS m,
  COUNT(*) FILTER (WHERE channel='ONLINE') AS n_online,
  COUNT(*) FILTER (WHERE channel='MOBILE') AS n_mobile,
  COUNT(*) FILTER (WHERE channel='PHONE') AS n_phone,
  COUNT(*) FILTER (WHERE channel='UNKNOWN' OR channel IS NULL) AS n_unknown,
  COUNT(*) AS n
FROM src GROUP BY 1 ORDER BY 1
""").df()
ch.to_csv(OUT / "diag_channel_monthly.csv", index=False)

# complete-month daily median for spike z
print("Yoy totals")
print(yoy_tot)
print("\nContrib")
print(contrib)
print("\nFeb24")
print(feb24.head(12))
print("\nSnow")
print(snow)
print("\nWater recode")
print(water.to_string(index=False))
print("\nJan window last rows")
print(jan_early.to_string(index=False))
