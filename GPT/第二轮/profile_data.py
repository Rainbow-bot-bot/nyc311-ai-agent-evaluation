from pathlib import Path
import json, hashlib, duckdb, pyarrow.parquet as pq
ROOT=Path(__file__).parent
OUT=ROOT/'evidence'; OUT.mkdir(exist_ok=True)
manifest=json.loads(Path(r'D:\项目2\input\data_manifest.json').read_text(encoding='utf-8-sig'))
files=[]
for item in manifest['files']:
    p=Path(manifest['source_path'])/item['name']
    digest=hashlib.file_digest(p.open('rb'),'sha256').hexdigest()
    meta=pq.ParquetFile(p)
    files.append(dict(file=p.name,bytes=p.stat().st_size,rows=meta.metadata.num_rows,sha256=digest,match=digest==item['sha256'],columns=meta.schema.names))
assert all(x['match'] for x in files)
(OUT/'source_verification.json').write_text(json.dumps(files,ensure_ascii=False,indent=2),encoding='utf-8')
c=duckdb.connect(); c.execute('SET threads=4'); c.execute("SET memory_limit='3GB'")
c.execute(f"SET temp_directory='{OUT.as_posix()}/tmp'")
c.execute("CREATE VIEW raw AS SELECT * FROM read_parquet('D:/项目1/原始数据/*.parquet',hive_partitioning=false,union_by_name=true)")
def save(name,sql):
    d=c.sql(sql).df(); d.to_csv(OUT/(name+'.csv'),index=False,encoding='utf-8-sig'); print(name, d.head(30).to_string(index=False),flush=True); return d
save('schema','DESCRIBE raw')
save('sample',"SELECT unique_key,created_date,closed_date,agency,complaint_type,status,borough,open_data_channel_type FROM raw LIMIT 5")
cols=files[0]['columns']
save('missing', 'SELECT '+','.join(f"count(*) FILTER(WHERE nullif(trim(\"{x}\"),'') IS NULL) AS \"{x}\"" for x in cols)+' FROM raw')
c.execute("CREATE VIEW d AS SELECT *,try_cast(created_date AS TIMESTAMP) AS created,try_cast(closed_date AS TIMESTAMP) AS closed FROM raw")
save('quality',"""SELECT count(*) n,count(DISTINCT unique_key) unique_keys,count(*) FILTER(WHERE unique_key IS NULL OR trim(unique_key)='') missing_key,min(created) start_date,max(created) finish,count(*) FILTER(WHERE created IS NULL) invalid_created,count(*) FILTER(WHERE closed_date IS NOT NULL AND trim(closed_date)<>'' AND closed IS NULL) invalid_closed,count(*) FILTER(WHERE closed<created) negative_duration,count(*) FILTER(WHERE closed=created) zero_duration,count(*) FILTER(WHERE closed>timestamp '2026-09-07') closed_after_end,count(*) FILTER(WHERE status='Closed' AND closed IS NULL) closed_missing,count(*) FILTER(WHERE status<>'Closed' AND closed IS NOT NULL) nonclosed_with_date FROM d""")
save('monthly',"SELECT date_trunc('month',created)::DATE AS mth,count(*) n,count(DISTINCT created::DATE) AS day_count,min(created) AS first_ts,max(created) AS last_ts FROM d GROUP BY 1 ORDER BY 1")
save('dimensions',"SELECT date_trunc('month',created)::DATE AS mth,complaint_type,agency,borough,open_data_channel_type channel,count(*) n FROM d GROUP BY ALL")
for col in ['complaint_type','agency','borough','status','open_data_channel_type']:
    save(col,f'SELECT "{col}",count(*) n FROM d GROUP BY 1 ORDER BY 2 DESC')
save('closure_month',"SELECT date_trunc('month',created)::DATE AS mth,count(*) n,count(*) FILTER(WHERE status='Closed') closed_status,count(*) FILTER(WHERE closed>=created) valid_duration,median(epoch(closed-created)/3600) FILTER(WHERE closed>=created) median_hours,quantile_cont(epoch(closed-created)/3600,0.9) FILTER(WHERE closed>=created) p90_hours FROM d GROUP BY 1 ORDER BY 1")


