"""NYC 311 exploratory aggregation. Reads parquet in-place; writes only summaries."""
from __future__ import annotations

from pathlib import Path

import duckdb
import pandas as pd

RAW = Path(r"D:\项目1\原始数据")
OUT = Path(r"D:\项目2\Grok\summaries")
OUT.mkdir(parents=True, exist_ok=True)

GLOB = str(RAW / "*.parquet").replace("\\", "/")

con = duckdb.connect()
con.execute("SET threads TO 4")
con.execute(f"""
CREATE OR REPLACE VIEW src AS
SELECT
  unique_key,
  try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f') AS created_ts,
  try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f') AS closed_ts,
  agency,
  complaint_type,
  descriptor,
  status,
  borough,
  open_data_channel_type AS channel,
  incident_zip
FROM read_parquet('{GLOB}')
""")

# --- overview / quality ---
overview = con.execute("""
SELECT
  COUNT(*) AS n_rows,
  COUNT(DISTINCT unique_key) AS n_unique_keys,
  MIN(created_ts) AS min_created,
  MAX(created_ts) AS max_created,
  COUNT(*) FILTER (WHERE created_ts IS NULL) AS n_created_null,
  COUNT(*) FILTER (WHERE unique_key IS NULL OR unique_key = '') AS n_key_null,
  COUNT(*) FILTER (WHERE closed_ts IS NULL) AS n_closed_null,
  COUNT(*) FILTER (WHERE closed_ts IS NOT NULL AND closed_ts < created_ts) AS n_closed_before_created,
  COUNT(*) FILTER (WHERE closed_ts IS NOT NULL AND closed_ts > TIMESTAMP '2026-09-08') AS n_closed_after_snapshot,
  COUNT(*) FILTER (WHERE agency IS NULL OR agency = '') AS n_agency_blank,
  COUNT(*) FILTER (WHERE complaint_type IS NULL OR complaint_type = '') AS n_type_blank,
  COUNT(*) FILTER (WHERE borough IS NULL OR borough = '' OR borough = 'Unspecified') AS n_borough_unspecified,
  COUNT(*) FILTER (WHERE channel IS NULL OR channel = '' OR channel = 'UNKNOWN') AS n_channel_unknown
FROM src
""").df()
overview.to_csv(OUT / "overview.csv", index=False)

status_counts = con.execute("""
SELECT COALESCE(status, '(null)') AS status, COUNT(*) AS n
FROM src GROUP BY 1 ORDER BY n DESC
""").df()
status_counts.to_csv(OUT / "status_counts.csv", index=False)

borough_counts = con.execute("""
SELECT COALESCE(borough, '(null)') AS borough, COUNT(*) AS n
FROM src GROUP BY 1 ORDER BY n DESC
""").df()
borough_counts.to_csv(OUT / "borough_counts.csv", index=False)

channel_counts = con.execute("""
SELECT COALESCE(channel, '(null)') AS channel, COUNT(*) AS n
FROM src GROUP BY 1 ORDER BY n DESC
""").df()
channel_counts.to_csv(OUT / "channel_counts.csv", index=False)

agency_counts = con.execute("""
SELECT COALESCE(agency, '(null)') AS agency, COUNT(*) AS n
FROM src GROUP BY 1 ORDER BY n DESC
""").df()
agency_counts.to_csv(OUT / "agency_counts.csv", index=False)

# monthly created
monthly = con.execute("""
SELECT
  strftime(created_ts, '%Y-%m') AS created_month,
  COUNT(*) AS n_created,
  COUNT(*) FILTER (WHERE status = 'Closed') AS n_closed_status,
  COUNT(*) FILTER (WHERE closed_ts IS NOT NULL) AS n_has_closed_ts,
  COUNT(*) FILTER (WHERE status IN ('Open','In Progress','Assigned','Pending','Started')) AS n_openish,
  COUNT(*) FILTER (
    WHERE status = 'Closed'
      AND closed_ts IS NOT NULL
      AND closed_ts >= created_ts
  ) AS n_valid_close,
  MEDIAN(epoch(closed_ts - created_ts) / 3600.0) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS median_close_hours,
  QUANTILE_CONT(epoch(closed_ts - created_ts) / 3600.0, 0.9) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS p90_close_hours
FROM src
WHERE created_ts IS NOT NULL
GROUP BY 1
ORDER BY 1
""").df()
monthly.to_csv(OUT / "monthly.csv", index=False)

# daily for spike checks
daily = con.execute("""
SELECT CAST(created_ts AS DATE) AS created_day, COUNT(*) AS n
FROM src
WHERE created_ts IS NOT NULL
GROUP BY 1
ORDER BY 1
""").df()
daily.to_csv(OUT / "daily.csv", index=False)

# top complaint types overall
types_overall = con.execute("""
SELECT COALESCE(complaint_type, '(blank)') AS complaint_type, COUNT(*) AS n
FROM src GROUP BY 1 ORDER BY n DESC LIMIT 40
""").df()
types_overall.to_csv(OUT / "types_overall.csv", index=False)

# monthly x top types: first get top 15 types
top_types = types_overall["complaint_type"].head(15).tolist()
con.execute("CREATE OR REPLACE TABLE top_types AS SELECT * FROM types_overall LIMIT 15")

monthly_type = con.execute("""
SELECT
  strftime(s.created_ts, '%Y-%m') AS created_month,
  COALESCE(s.complaint_type, '(blank)') AS complaint_type,
  COUNT(*) AS n
FROM src s
WHERE s.created_ts IS NOT NULL
  AND COALESCE(s.complaint_type, '(blank)') IN (SELECT complaint_type FROM top_types)
GROUP BY 1, 2
ORDER BY 1, 3 DESC
""").df()
monthly_type.to_csv(OUT / "monthly_type.csv", index=False)

monthly_borough = con.execute("""
SELECT
  strftime(created_ts, '%Y-%m') AS created_month,
  COALESCE(borough, '(null)') AS borough,
  COUNT(*) AS n
FROM src
WHERE created_ts IS NOT NULL
GROUP BY 1, 2
ORDER BY 1, 3 DESC
""").df()
monthly_borough.to_csv(OUT / "monthly_borough.csv", index=False)

monthly_agency = con.execute("""
SELECT
  strftime(created_ts, '%Y-%m') AS created_month,
  COALESCE(agency, '(null)') AS agency,
  COUNT(*) AS n
FROM src
WHERE created_ts IS NOT NULL
GROUP BY 1, 2
ORDER BY 1, 3 DESC
""").df()
monthly_agency.to_csv(OUT / "monthly_agency.csv", index=False)

# close time by type for closed valid
close_by_type = con.execute("""
SELECT
  COALESCE(complaint_type, '(blank)') AS complaint_type,
  COUNT(*) AS n_created,
  COUNT(*) FILTER (WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts) AS n_valid_close,
  MEDIAN(epoch(closed_ts - created_ts) / 3600.0) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS median_close_hours,
  QUANTILE_CONT(epoch(closed_ts - created_ts) / 3600.0, 0.9) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS p90_close_hours
FROM src
GROUP BY 1
HAVING COUNT(*) >= 20000
ORDER BY n_created DESC
""").df()
close_by_type.to_csv(OUT / "close_by_type.csv", index=False)

close_by_agency = con.execute("""
SELECT
  COALESCE(agency, '(blank)') AS agency,
  COUNT(*) AS n_created,
  COUNT(*) FILTER (WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts) AS n_valid_close,
  COUNT(*) FILTER (WHERE status IN ('Open','In Progress','Assigned','Pending','Started')) AS n_openish,
  MEDIAN(epoch(closed_ts - created_ts) / 3600.0) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS median_close_hours,
  QUANTILE_CONT(epoch(closed_ts - created_ts) / 3600.0, 0.9) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS p90_close_hours
FROM src
GROUP BY 1
HAVING COUNT(*) >= 5000
ORDER BY n_created DESC
""").df()
close_by_agency.to_csv(OUT / "close_by_agency.csv", index=False)

close_by_borough = con.execute("""
SELECT
  COALESCE(borough, '(blank)') AS borough,
  COUNT(*) AS n_created,
  COUNT(*) FILTER (WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts) AS n_valid_close,
  MEDIAN(epoch(closed_ts - created_ts) / 3600.0) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS median_close_hours,
  QUANTILE_CONT(epoch(closed_ts - created_ts) / 3600.0, 0.9) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS p90_close_hours
FROM src
WHERE borough IS NOT NULL AND borough <> '' AND borough <> 'Unspecified'
GROUP BY 1
ORDER BY n_created DESC
""").df()
close_by_borough.to_csv(OUT / "close_by_borough.csv", index=False)

# heat/hot water monthly
heat_monthly = con.execute("""
SELECT
  strftime(created_ts, '%Y-%m') AS created_month,
  COUNT(*) FILTER (WHERE complaint_type = 'HEAT/HOT WATER') AS n_heat,
  COUNT(*) AS n_all,
  COUNT(*) FILTER (WHERE complaint_type = 'Noise - Residential') AS n_noise_res,
  COUNT(*) FILTER (WHERE complaint_type = 'Illegal Parking') AS n_illegal_parking,
  COUNT(*) FILTER (WHERE complaint_type = 'Blocked Driveway') AS n_blocked,
  COUNT(*) FILTER (WHERE complaint_type ILIKE '%Noise%') AS n_any_noise
FROM src
WHERE created_ts IS NOT NULL
GROUP BY 1
ORDER BY 1
""").df()
heat_monthly.to_csv(OUT / "heat_noise_parking_monthly.csv", index=False)

# heat by borough and winter windows
heat_borough = con.execute("""
SELECT
  COALESCE(borough, '(blank)') AS borough,
  COUNT(*) AS n_heat,
  MEDIAN(epoch(closed_ts - created_ts) / 3600.0) FILTER (
    WHERE status = 'Closed' AND closed_ts IS NOT NULL AND closed_ts >= created_ts
  ) AS median_close_hours
FROM src
WHERE complaint_type = 'HEAT/HOT WATER'
GROUP BY 1
ORDER BY n_heat DESC
""").df()
heat_borough.to_csv(OUT / "heat_by_borough.csv", index=False)

# noise hour-of-week pattern
noise_hod = con.execute("""
SELECT
  EXTRACT(dow FROM created_ts) AS dow,
  EXTRACT(hour FROM created_ts) AS hour,
  COUNT(*) AS n
FROM src
WHERE complaint_type = 'Noise - Residential' AND created_ts IS NOT NULL
GROUP BY 1, 2
ORDER BY 1, 2
""").df()
noise_hod.to_csv(OUT / "noise_residential_dow_hour.csv", index=False)

# YoY complete months: Oct-Aug 11 months
yoy_type = con.execute("""
WITH tagged AS (
  SELECT
    complaint_type,
    CASE
      WHEN created_ts >= TIMESTAMP '2024-10-01' AND created_ts < TIMESTAMP '2025-09-01' THEN 'y1_oct24_aug25'
      WHEN created_ts >= TIMESTAMP '2025-10-01' AND created_ts < TIMESTAMP '2026-09-01' THEN 'y2_oct25_aug26'
    END AS period
  FROM src
)
SELECT COALESCE(complaint_type, '(blank)') AS complaint_type,
  COUNT(*) FILTER (WHERE period = 'y1_oct24_aug25') AS n_y1,
  COUNT(*) FILTER (WHERE period = 'y2_oct25_aug26') AS n_y2
FROM tagged
WHERE period IS NOT NULL
GROUP BY 1
HAVING n_y1 + n_y2 >= 5000
ORDER BY (n_y2 - n_y1) DESC
""").df()
yoy_type.to_csv(OUT / "yoy_type.csv", index=False)

yoy_borough = con.execute("""
WITH tagged AS (
  SELECT
    borough,
    CASE
      WHEN created_ts >= TIMESTAMP '2024-10-01' AND created_ts < TIMESTAMP '2025-09-01' THEN 'y1'
      WHEN created_ts >= TIMESTAMP '2025-10-01' AND created_ts < TIMESTAMP '2026-09-01' THEN 'y2'
    END AS period
  FROM src
)
SELECT COALESCE(borough, '(blank)') AS borough,
  COUNT(*) FILTER (WHERE period = 'y1') AS n_y1,
  COUNT(*) FILTER (WHERE period = 'y2') AS n_y2
FROM tagged
WHERE period IS NOT NULL
GROUP BY 1
ORDER BY n_y2 DESC
""").df()
yoy_borough.to_csv(OUT / "yoy_borough.csv", index=False)

yoy_agency = con.execute("""
WITH tagged AS (
  SELECT
    agency,
    CASE
      WHEN created_ts >= TIMESTAMP '2024-10-01' AND created_ts < TIMESTAMP '2025-09-01' THEN 'y1'
      WHEN created_ts >= TIMESTAMP '2025-10-01' AND created_ts < TIMESTAMP '2026-09-01' THEN 'y2'
    END AS period
  FROM src
)
SELECT COALESCE(agency, '(blank)') AS agency,
  COUNT(*) FILTER (WHERE period = 'y1') AS n_y1,
  COUNT(*) FILTER (WHERE period = 'y2') AS n_y2
FROM tagged
WHERE period IS NOT NULL
GROUP BY 1
ORDER BY (n_y2 - n_y1) DESC
""").df()
yoy_agency.to_csv(OUT / "yoy_agency.csv", index=False)

# top daily spikes
daily_spikes = daily.nlargest(20, "n").sort_values("n", ascending=False)
daily_spikes.to_csv(OUT / "daily_spikes.csv", index=False)

# descriptors for heat
heat_desc = con.execute("""
SELECT COALESCE(descriptor, '(blank)') AS descriptor, COUNT(*) AS n
FROM src
WHERE complaint_type = 'HEAT/HOT WATER'
GROUP BY 1 ORDER BY n DESC
""").df()
heat_desc.to_csv(OUT / "heat_descriptors.csv", index=False)

# open backlog by created month and type (top)
open_by_month_agency = con.execute("""
SELECT
  strftime(created_ts, '%Y-%m') AS created_month,
  COALESCE(agency, '(blank)') AS agency,
  COUNT(*) FILTER (WHERE status IN ('Open','In Progress','Assigned','Pending','Started')) AS n_openish,
  COUNT(*) AS n
FROM src
WHERE created_ts IS NOT NULL
GROUP BY 1, 2
HAVING n_openish >= 50
ORDER BY 1, 3 DESC
""").df()
open_by_month_agency.to_csv(OUT / "open_by_month_agency.csv", index=False)

# weekday vs weekend volume
dow_vol = con.execute("""
SELECT
  EXTRACT(dow FROM created_ts) AS dow,
  COUNT(*) AS n_all,
  COUNT(*) FILTER (WHERE complaint_type = 'Noise - Residential') AS n_noise_res,
  COUNT(*) FILTER (WHERE complaint_type = 'Illegal Parking') AS n_parking,
  COUNT(*) FILTER (WHERE complaint_type = 'HEAT/HOT WATER') AS n_heat
FROM src
WHERE created_ts IS NOT NULL
GROUP BY 1
ORDER BY 1
""").df()
dow_vol.to_csv(OUT / "dow_volume.csv", index=False)

# snapshot: still-open tickets created before last 30 days of window
open_age = con.execute("""
SELECT
  COALESCE(agency, '(blank)') AS agency,
  COUNT(*) AS n_openish,
  MEDIAN(epoch(TIMESTAMP '2026-09-07' - created_ts) / 86400.0) AS median_age_days,
  QUANTILE_CONT(epoch(TIMESTAMP '2026-09-07' - created_ts) / 86400.0, 0.9) AS p90_age_days
FROM src
WHERE status IN ('Open','In Progress','Assigned','Pending','Started')
GROUP BY 1
HAVING COUNT(*) >= 100
ORDER BY n_openish DESC
""").df()
open_age.to_csv(OUT / "open_age_by_agency.csv", index=False)

# file-level row counts to reconcile with monthly
file_check = con.execute(f"""
SELECT filename, COUNT(*) AS n
FROM read_parquet('{GLOB}', filename=true)
GROUP BY 1
ORDER BY 1
""").df()
file_check.to_csv(OUT / "file_row_counts.csv", index=False)

print("overview")
print(overview.to_string(index=False))
print("\nmonthly")
print(monthly.to_string(index=False))
print("\nstatus")
print(status_counts.to_string(index=False))
print("wrote summaries to", OUT)
