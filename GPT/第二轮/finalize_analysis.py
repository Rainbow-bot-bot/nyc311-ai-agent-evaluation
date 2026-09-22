from pathlib import Path
import pandas as pd,json,duckdb
E=Path(__file__).parent/'evidence'
dd=pd.read_csv(E/'daily_focus.csv');dd['yr']=dd.dt.str[:4].astype(int);dd['mo']=dd.dt.str[5:7].astype(int)
stats=[]
for cat in ['Snow or Ice','Street Condition','HEAT/HOT WATER']:
    a=dd[(dd.yr==2025)&(dd.complaint_type==cat)];b=dd[(dd.yr==2026)&(dd.complaint_type==cat)]
    peak=b.nlargest(5,'n'); base=int(a.n.sum());current=int(b.n.sum());removed=int(peak.n.sum())
    stats.append(dict(category=cat,base=base,current=current,top5=removed,top5_share=removed/current,current_without_top5=current-removed,growth_without_top5=(current-removed)/base-1,months_up=int((pd.concat([a.groupby('mo').n.sum().rename('base'),b.groupby('mo').n.sum().rename('current')],axis=1).fillna(0).eval('current>base')).sum())))
pd.DataFrame(stats).to_csv(E/'peak_sensitivity.csv',index=False,encoding='utf-8-sig')
desc=pd.read_csv(E/'descriptors.csv');p=desc.pivot_table(index=['complaint_type','descriptor'],columns='yr',values='n',aggfunc='sum',fill_value=0);p.columns=['base','current'];p['delta']=p.current-p.base;p=p.sort_values('delta',ascending=False).reset_index();p.to_csv(E/'descriptor_comparison.csv',index=False,encoding='utf-8-sig')
c=duckdb.connect();c.execute('SET threads=4');c.execute("CREATE VIEW d AS SELECT *,try_cast(created_date AS TIMESTAMP) AS created FROM read_parquet('D:/项目1/原始数据/*.parquet',hive_partitioning=false)")
q="""SELECT unique_key,created_date,complaint_type,descriptor,agency,borough,open_data_channel_type FROM d WHERE created>=timestamp '2026-02-24' AND created<timestamp '2026-02-25' AND complaint_type='Snow or Ice' QUALIFY row_number() OVER(PARTITION BY borough,open_data_channel_type ORDER BY unique_key)<=2"""
c.sql(q).df().to_csv(E/'peak_records.csv',index=False,encoding='utf-8-sig')
q2="""SELECT borough,descriptor,open_data_channel_type AS channel,count(*) n,count(DISTINCT incident_address) addresses FROM d WHERE created>=timestamp '2026-02-24' AND created<timestamp '2026-02-25' AND complaint_type='Snow or Ice' GROUP BY ALL ORDER BY n DESC"""
c.sql(q2).df().to_csv(E/'peak_breakdown.csv',index=False,encoding='utf-8-sig')
q3="""SELECT count(*) n,count(DISTINCT incident_address) address_n,count(*) FILTER(WHERE nullif(trim(incident_address),'') IS NULL) missing_address FROM d WHERE created>=timestamp '2026-02-24' AND created<timestamp '2026-02-25' AND complaint_type='Snow or Ice'"""
c.sql(q3).df().to_csv(E/'peak_control.csv',index=False,encoding='utf-8-sig')
(E/'final_queries.json').write_text(json.dumps([q,q2,q3],indent=2),encoding='utf-8')
# Full evidence reconciliations, not self-reported status cells.
quality=pd.read_csv(E/'quality.csv').iloc[0];dim=pd.read_csv(E/'dimensions.csv');monthly=pd.read_csv(E/'monthly.csv');control=pd.read_csv(E/'control.csv')
sources=json.loads((E/'source_verification.json').read_text(encoding='utf-8'))
assert sum(x['rows'] for x in sources)==quality.n==monthly.n.sum()==dim.n.sum()
assert quality.n==quality.unique_keys and quality.missing_key==0 and quality.invalid_created==0
assert all(x['columns']==sources[0]['columns'] for x in sources)
assert (control.day_count==243).all()
checks={'source_files':len(sources),'rows':int(quality.n),'unique_keys':int(quality.unique_keys),'sha256_matches':sum(x['match'] for x in sources),'columns':len(sources[0]['columns'])}
for col in ['complaint_type','agency','borough','channel','month']:
    t=pd.read_csv(E/f'compare_{col}.csv'); assert t.base.sum()==control.iloc[0].n and t.current.sum()==control.iloc[1].n;assert t.delta.sum()==287342
checks['all_decompositions_reconcile']=True
pc=pd.read_csv(E/'partition_coverage.csv');assert pc.n.sum()==quality.n
for row in pc.itertuples():
    name=Path(row.filename).name; bounds=name.removeprefix('created_date=').removesuffix('.parquet').split('_to_')
    assert pd.Timestamp(row.first_ts)>=pd.Timestamp(bounds[0]) and pd.Timestamp(row.last_ts)<pd.Timestamp(bounds[1])
checks['all_partitions_in_bounds']=True
(E/'numerical_checks.json').write_text(json.dumps(checks,indent=2),encoding='utf-8')
print(pd.DataFrame(stats).to_string(index=False));print(p.head(15).to_string(index=False));print(checks)
