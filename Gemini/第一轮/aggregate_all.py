import os
import duckdb
import pandas as pd

DATA_DIR = "D:/项目1/原始数据"
pattern = f"{DATA_DIR}/*.parquet"
con = duckdb.connect()

print("Extracting all core aggregated datasets...")

monthly_df = con.execute(f'''
    SELECT 
        strftime(try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), '%Y-%m') as year_month,
        count(*) as total_requests,
        count(case when complaint_type like '%Noise%' then 1 end) as noise_requests,
        count(case when complaint_type = 'HEAT/HOT WATER' then 1 end) as heat_requests,
        count(case when complaint_type in ('Illegal Parking', 'Blocked Driveway') then 1 end) as parking_requests,
        count(case when agency = 'NYPD' then 1 end) as nypd_requests,
        count(case when agency = 'HPD' then 1 end) as hpd_requests,
        count(case when agency = 'DSNY' then 1 end) as dsny_requests
    FROM read_parquet('{pattern}')
    GROUP BY year_month
    ORDER BY year_month
''').fetchdf()

borough_df = con.execute(f'''
    SELECT 
        borough,
        count(*) as total_requests,
        count(case when complaint_type like '%Noise%' then 1 end) as noise_requests,
        count(case when complaint_type = 'HEAT/HOT WATER' then 1 end) as heat_requests,
        count(case when complaint_type in ('Illegal Parking', 'Blocked Driveway') then 1 end) as parking_requests,
        count(case when agency = 'HPD' then 1 end) as hpd_requests,
        count(case when agency = 'NYPD' then 1 end) as nypd_requests
    FROM read_parquet('{pattern}')
    GROUP BY borough
    ORDER BY total_requests DESC
''').fetchdf()

agency_df = con.execute(f'''
    SELECT 
        agency,
        count(*) as total_requests,
        sum(case when status = 'Closed' then 1 else 0 end) as closed_requests,
        round(sum(case when status = 'Closed' then 1 else 0 end) * 100.0 / count(*), 2) as close_rate_pct,
        round(avg(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 3600.0 end), 1) as avg_resolution_hours,
        round(median(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 3600.0 end), 1) as median_resolution_hours,
        round(quantile_cont(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 3600.0 end, 0.75), 1) as p75_resolution_hours,
        round(quantile_cont(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 3600.0 end, 0.90), 1) as p90_resolution_hours
    FROM read_parquet('{pattern}')
    GROUP BY agency
    HAVING count(*) >= 1000
    ORDER BY total_requests DESC
''').fetchdf()

complaint_types_df = con.execute(f'''
    SELECT 
        complaint_type,
        agency,
        count(*) as total_requests,
        round(count(*) * 100.0 / 7525498, 2) as pct_of_total,
        round(median(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 3600.0 end), 1) as median_resolution_hours,
        round(avg(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 86400.0 end), 1) as avg_resolution_days
    FROM read_parquet('{pattern}')
    GROUP BY complaint_type, agency
    ORDER BY total_requests DESC
    LIMIT 25
''').fetchdf()

hourly_df = con.execute(f'''
    SELECT 
        extract('hour' from try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f')) as hour_of_day,
        count(*) as total_requests,
        count(case when complaint_type like '%Noise%' then 1 end) as noise_requests,
        count(case when complaint_type in ('Illegal Parking', 'Blocked Driveway') then 1 end) as parking_requests,
        count(case when complaint_type = 'HEAT/HOT WATER' then 1 end) as heat_requests
    FROM read_parquet('{pattern}')
    GROUP BY hour_of_day
    ORDER BY hour_of_day
''').fetchdf()

dow_df = con.execute(f'''
    SELECT 
        dayname(try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f')) as day_of_week,
        dayofweek(try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f')) as dow_num,
        count(*) as total_requests,
        count(case when complaint_type like '%Noise%' then 1 end) as noise_requests,
        count(case when complaint_type in ('Illegal Parking', 'Blocked Driveway') then 1 end) as parking_requests,
        count(case when complaint_type = 'HEAT/HOT WATER' then 1 end) as heat_requests
    FROM read_parquet('{pattern}')
    GROUP BY day_of_week, dow_num
    ORDER BY dow_num
''').fetchdf()

top_addrs_df = con.execute(f'''
    SELECT 
        coalesce(incident_address, 'N/A') as incident_address,
        coalesce(borough, 'N/A') as borough,
        coalesce(incident_zip, 'N/A') as incident_zip,
        count(*) as total_requests,
        count(case when complaint_type like '%Noise%' then 1 end) as noise_requests,
        count(case when complaint_type in ('Illegal Parking', 'Blocked Driveway') then 1 end) as parking_requests,
        count(case when complaint_type = 'HEAT/HOT WATER' then 1 end) as heat_requests,
        round(count(case when open_data_channel_type = 'MOBILE' then 1 end) * 100.0 / count(*), 1) as mobile_pct,
        round(count(case when open_data_channel_type = 'ONLINE' then 1 end) * 100.0 / count(*), 1) as online_pct,
        round(count(case when open_data_channel_type = 'PHONE' then 1 end) * 100.0 / count(*), 1) as phone_pct
    FROM read_parquet('{pattern}')
    WHERE incident_address IS NOT NULL AND trim(incident_address) != ''
    GROUP BY incident_address, borough, incident_zip
    ORDER BY total_requests DESC
    LIMIT 25
''').fetchdf()

heat_zips_df = con.execute(f'''
    SELECT 
        incident_zip,
        borough,
        count(*) as heat_requests,
        count(distinct incident_address) as unique_addresses_with_heat_issues,
        round(count(*) * 1.0 / nullif(count(distinct incident_address), 0), 1) as avg_heat_requests_per_address
    FROM read_parquet('{pattern}')
    WHERE complaint_type = 'HEAT/HOT WATER' AND incident_zip IS NOT NULL AND trim(incident_zip) != ''
    GROUP BY incident_zip, borough
    ORDER BY heat_requests DESC
    LIMIT 20
''').fetchdf()

channel_df = con.execute(f'''
    SELECT 
        open_data_channel_type as channel,
        count(*) as total_requests,
        round(count(*) * 100.0 / 7525498, 2) as pct_of_total,
        round(median(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 3600.0 end), 1) as median_resolution_hours,
        round(avg(case when closed_date is not null and closed_date >= created_date then date_diff('second', try_strptime(created_date, '%Y-%m-%dT%H:%M:%S.%f'), try_strptime(closed_date, '%Y-%m-%dT%H:%M:%S.%f')) / 86400.0 end), 1) as avg_resolution_days
    FROM read_parquet('{pattern}')
    GROUP BY channel
    ORDER BY total_requests DESC
''').fetchdf()

resolution_actions_df = con.execute(f'''
    SELECT 
        complaint_type,
        case 
            when resolution_description like '%no evidence of the violation%' or resolution_description like '%observed no evidence%' or resolution_description like '%observed no criminal violation%' then '未发现违规证据 (No Evidence Observed)'
            when resolution_description like '%responsible for the condition were gone%' or resolution_description like '%condition was gone%' then '到达现场时当事人已离开 (Condition/Party Gone)'
            when resolution_description like '%issued a summons%' then '开具传票/罚单 (Summons Issued)'
            when resolution_description like '%took action to fix%' or resolution_description like '%action to fix the condition%' then '已采取处置措施修复 (Action Taken to Fix)'
            when resolution_description like '%police action was not necessary%' then '判定无需警察执法 (Action Not Necessary)'
            else '其他处理结果 (Other Resolutions)'
        end as action_category,
        count(*) as count_cases,
        round(count(*) * 100.0 / sum(count(*)) over (partition by complaint_type), 2) as pct_within_type
    FROM read_parquet('{pattern}')
    WHERE complaint_type in ('Noise - Residential', 'Illegal Parking') AND status = 'Closed'
    GROUP BY complaint_type, action_category
    ORDER BY complaint_type, count_cases DESC
''').fetchdf()

os.makedirs("D:/项目2/Gemini/data_cache", exist_ok=True)
monthly_df.to_parquet("D:/项目2/Gemini/data_cache/monthly.parquet")
borough_df.to_parquet("D:/项目2/Gemini/data_cache/borough.parquet")
agency_df.to_parquet("D:/项目2/Gemini/data_cache/agency.parquet")
complaint_types_df.to_parquet("D:/项目2/Gemini/data_cache/complaint_types.parquet")
hourly_df.to_parquet("D:/项目2/Gemini/data_cache/hourly.parquet")
dow_df.to_parquet("D:/项目2/Gemini/data_cache/dow.parquet")
top_addrs_df.to_parquet("D:/项目2/Gemini/data_cache/top_addrs.parquet")
heat_zips_df.to_parquet("D:/项目2/Gemini/data_cache/heat_zips.parquet")
channel_df.to_parquet("D:/项目2/Gemini/data_cache/channel.parquet")
resolution_actions_df.to_parquet("D:/项目2/Gemini/data_cache/resolution_actions.parquet")

print("All cache files successfully exported!")
