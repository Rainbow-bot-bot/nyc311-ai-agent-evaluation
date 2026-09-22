import duckdb
import pandas as pd
import sys

# Ensure UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

db = duckdb.connect()
parquet_path = "D:/项目1/原始数据/*.parquet"

print("--- 1. NYPD Resolution Descriptions for Illegal Parking ---")
res_nypd_parking = db.execute(f"""
    SELECT 
        resolution_description,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as pct
    FROM read_parquet('{parquet_path}')
    WHERE agency = 'NYPD' AND complaint_type = 'Illegal Parking'
    GROUP BY 1
    ORDER BY cnt DESC
    LIMIT 10
""").df()
res_nypd_parking.to_csv("data/nypd_parking_resolutions.csv", index=False, encoding='utf-8-sig')

print("--- 2. NYPD Resolution Descriptions for Noise - Residential ---")
res_nypd_noise = db.execute(f"""
    SELECT 
        resolution_description,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as pct
    FROM read_parquet('{parquet_path}')
    WHERE agency = 'NYPD' AND complaint_type = 'Noise - Residential'
    GROUP BY 1
    ORDER BY cnt DESC
    LIMIT 10
""").df()
res_nypd_noise.to_csv("data/nypd_noise_resolutions.csv", index=False, encoding='utf-8-sig')

print("--- 3. HPD Resolution Descriptions for HEAT/HOT WATER ---")
res_hpd_heat = db.execute(f"""
    SELECT 
        resolution_description,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as pct
    FROM read_parquet('{parquet_path}')
    WHERE agency = 'HPD' AND complaint_type = 'HEAT/HOT WATER'
    GROUP BY 1
    ORDER BY cnt DESC
    LIMIT 10
""").df()
res_hpd_heat.to_csv("data/hpd_heat_resolutions.csv", index=False, encoding='utf-8-sig')

print("--- 4. Pareto Analysis for Heat/Hot Water Addresses ---")
res_heat_pareto = db.execute(f"""
    WITH addr_counts AS (
        SELECT 
            incident_address,
            COUNT(*) as cnt
        FROM read_parquet('{parquet_path}')
        WHERE complaint_type = 'HEAT/HOT WATER' AND incident_address IS NOT NULL
        GROUP BY 1
    ),
    ranked AS (
        SELECT 
            cnt,
            ROW_NUMBER() OVER (ORDER BY cnt DESC) as rnk,
            COUNT(*) OVER () as total_addrs,
            SUM(cnt) OVER () as total_heat
        FROM addr_counts
    )
    SELECT 
        CASE 
            WHEN rnk <= total_addrs * 0.01 THEN '1. Top 1% Addresses'
            WHEN rnk <= total_addrs * 0.05 THEN '2. Top 1-5% Addresses'
            WHEN rnk <= total_addrs * 0.10 THEN '3. Top 5-10% Addresses'
            WHEN rnk <= total_addrs * 0.20 THEN '4. Top 10-20% Addresses'
            ELSE '5. Remaining 80% Addresses'
        END as tier,
        COUNT(*) as addr_count,
        ROUND(COUNT(*) * 100.0 / MAX(total_addrs), 2) as addr_share_pct,
        SUM(cnt) as complaint_count,
        ROUND(SUM(cnt) * 100.0 / MAX(total_heat), 2) as complaint_share_pct
    FROM ranked
    GROUP BY 1
    ORDER BY 1
""").df()
print(res_heat_pareto.to_string(index=False))
res_heat_pareto.to_csv("data/heat_pareto_analysis.csv", index=False, encoding='utf-8-sig')

print("\n--- 5. NYPD Illegal Parking Action Rate ---")
# Check how often NYPD took action vs no action/no violation
res_action_parking = db.execute(f"""
    SELECT 
        CASE 
            WHEN resolution_description ILIKE '%issued a summons%' THEN 'Summons Issued'
            WHEN resolution_description ILIKE '%took action to fix the condition%' THEN 'Action Taken'
            WHEN resolution_description ILIKE '%observed no evidence%' THEN 'No Evidence / Gone'
            WHEN resolution_description ILIKE '%police action was not necessary%' THEN 'Action Not Necessary'
            WHEN resolution_description ILIKE '%unable to gain entry%' THEN 'Unable to Gain Entry'
            WHEN resolution_description ILIKE '%arrest%' THEN 'Arrest Made'
            ELSE 'Other / Unspecified'
        END as outcome_category,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as pct
    FROM read_parquet('{parquet_path}')
    WHERE agency = 'NYPD' AND complaint_type = 'Illegal Parking'
    GROUP BY 1
    ORDER BY cnt DESC
""").df()
print(res_action_parking.to_string(index=False))
res_action_parking.to_csv("data/nypd_parking_outcomes.csv", index=False, encoding='utf-8-sig')

print("\n--- 6. NYPD Residential Noise Action Rate ---")
res_action_noise = db.execute(f"""
    SELECT 
        CASE 
            WHEN resolution_description ILIKE '%issued a summons%' THEN 'Summons Issued'
            WHEN resolution_description ILIKE '%took action to fix the condition%' THEN 'Action Taken'
            WHEN resolution_description ILIKE '%observed no evidence%' THEN 'No Evidence / Gone'
            WHEN resolution_description ILIKE '%police action was not necessary%' THEN 'Action Not Necessary'
            WHEN resolution_description ILIKE '%unable to gain entry%' THEN 'Unable to Gain Entry'
            WHEN resolution_description ILIKE '%arrest%' THEN 'Arrest Made'
            ELSE 'Other / Unspecified'
        END as outcome_category,
        COUNT(*) as cnt,
        ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER (), 2) as pct
    FROM read_parquet('{parquet_path}')
    WHERE agency = 'NYPD' AND complaint_type = 'Noise - Residential'
    GROUP BY 1
    ORDER BY cnt DESC
""").df()
print(res_action_noise.to_string(index=False))
res_action_noise.to_csv("data/nypd_noise_outcomes.csv", index=False, encoding='utf-8-sig')
