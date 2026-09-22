from pathlib import Path
import duckdb,pandas as pd,json
E=Path(__file__).parent/'evidence'
c=duckdb.connect();c.execute('SET threads=4')
c.execute("CREATE VIEW d AS SELECT *,try_cast(created_date AS TIMESTAMP) AS created,try_cast(closed_date AS TIMESTAMP) AS closed FROM read_parquet('D:/项目1/原始数据/*.parquet',hive_partitioning=false)")
c.execute("CREATE VIEW s AS SELECT *,year(created) AS yr,month(created) AS mo FROM d WHERE (created>=timestamp '2025-01-01' AND created<timestamp '2025-09-01') OR (created>=timestamp '2026-01-01' AND created<timestamp '2026-09-01')")
queries={
'control':"SELECT yr,count(*) n,count(DISTINCT created::DATE) AS day_count,count(DISTINCT unique_key) AS keys FROM s GROUP BY 1 ORDER BY 1",
'daily_focus':"SELECT created::DATE AS dt,complaint_type,count(*) n FROM s WHERE complaint_type IN ('Street Condition','Snow or Ice','HEAT/HOT WATER') GROUP BY ALL ORDER BY 1,2",
'descriptors':"SELECT yr,complaint_type,descriptor,agency,count(*) n FROM s WHERE complaint_type IN ('Street Condition','Snow or Ice','HEAT/HOT WATER') GROUP BY ALL ORDER BY n DESC",
'street_cross':"SELECT yr,mo,complaint_type,descriptor,borough,open_data_channel_type AS channel,count(*) n FROM s WHERE complaint_type IN ('Street Condition','Snow or Ice') GROUP BY ALL",
'taxonomy_dep':"SELECT date_trunc('month',created)::DATE AS mth,complaint_type,count(*) n FROM d WHERE agency='DEP' GROUP BY ALL ORDER BY 1,2",
'focus_records':"SELECT unique_key,created_date,complaint_type,descriptor,agency,borough,open_data_channel_type,closed_date,status FROM s WHERE complaint_type IN ('Street Condition','Snow or Ice') QUALIFY row_number() OVER(PARTITION BY yr,complaint_type ORDER BY created,unique_key)<=8",
'anomaly_counts':"SELECT agency,status,count(*) n,count(*) FILTER(WHERE closed<created) negative_n,count(*) FILTER(WHERE closed=created) zero_n,count(*) FILTER(WHERE status='Closed' AND closed IS NULL) missing_n,count(*) FILTER(WHERE status<>'Closed' AND closed IS NOT NULL) inconsistent_n FROM d GROUP BY ALL",
'anomaly_records':"SELECT unique_key,created_date,closed_date,status,agency,complaint_type FROM d WHERE closed<created OR closed>timestamp '2026-09-07' QUALIFY row_number() OVER(PARTITION BY (closed>timestamp '2026-09-07') ORDER BY closed-created,unique_key)<=8",
'partition_coverage':"SELECT filename,count(*) n,min(created) first_ts,max(created) last_ts FROM (SELECT filename,try_cast(created_date AS TIMESTAMP) created FROM read_parquet('D:/项目1/原始数据/*.parquet',hive_partitioning=false,filename=true)) GROUP BY 1 ORDER BY 1",
}
for name,sql in queries.items():
    data=c.sql(sql).df();data.to_csv(E/(name+'.csv'),index=False,encoding='utf-8-sig');print(name,data.head(12).to_string(index=False),flush=True)
(E/'deep_queries.json').write_text(json.dumps(queries,ensure_ascii=False,indent=2),encoding='utf-8')

d=pd.read_csv(E/'dimensions.csv');d['yr']=d.mth.str[:4].astype(int);d['mo']=d.mth.str[5:7].astype(int)
s=d[(d.yr.isin([2025,2026]))&(d.mo<=8)].copy()
results=[]
for title,mask in [('全量',s.n>=0),('排除雪冰与道路',~s.complaint_type.isin(['Snow or Ice','Street Condition'])),('排除雪冰、道路与供暖',~s.complaint_type.isin(['Snow or Ice','Street Condition','HEAT/HOT WATER'])),('只看4—8月',s.mo>=4)]:
    t=s[mask].groupby('yr').n.sum();results.append([title,int(t[2025]),int(t[2026]),int(t[2026]-t[2025]),t[2026]/t[2025]-1])
pd.DataFrame(results,columns=['check','base','current','delta','growth']).to_csv(E/'sensitivity.csv',index=False,encoding='utf-8-sig')
# Baseline standardization over common month x category x agency cells.
g=s.groupby(['yr','mo','complaint_type','agency']).n.sum().rename('total').reset_index()
o=s[s.channel=='ONLINE'].groupby(['yr','mo','complaint_type','agency']).n.sum().rename('online').reset_index()
g=g.merge(o,how='left').fillna({'online':0})
p=g.pivot(index=['mo','complaint_type','agency'],columns='yr',values=['total','online']).fillna(0)
common=p[(p[('total',2025)]>0)&(p[('total',2026)]>0)].copy()
common['weight']=common[('total',2025)]/common[('total',2025)].sum()
out={'common_base':int(common[('total',2025)].sum()),'common_current':int(common[('total',2026)].sum()),'base_total':int(s[s.yr==2025].n.sum()),'current_total':int(s[s.yr==2026].n.sum()),'base_online':int(s[(s.yr==2025)&(s.channel=='ONLINE')].n.sum()),'current_online':int(s[(s.yr==2026)&(s.channel=='ONLINE')].n.sum()),'base_common_rate':float((common.weight*common[('online',2025)]/common[('total',2025)]).sum()),'fixed_mix_current_rate':float((common.weight*common[('online',2026)]/common[('total',2026)]).sum())}
common.columns=['_'.join(str(v) for v in x if v!='') for x in common.columns]
common.reset_index().to_csv(E/'channel_standardization.csv',index=False,encoding='utf-8-sig')
(E/'standardization_summary.json').write_text(json.dumps(out,indent=2),encoding='utf-8');print('standardization',out)
print('sensitivity',results)
