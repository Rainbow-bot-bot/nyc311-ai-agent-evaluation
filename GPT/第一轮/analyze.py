"""Read-only full-shard analysis. All derived files stay next to this script."""
from pathlib import Path
import json, hashlib, collections, re
import pandas as pd
import pyarrow.parquet as pq

ROOT=Path(__file__).resolve().parent
SRC=Path(r'D:\项目1\原始数据')
manifest=json.loads(Path(r'D:\项目2\input\data_manifest.json').read_text(encoding='utf-8-sig'))
out=ROOT/'汇总'; out.mkdir(exist_ok=True)
stats=[]; missing=collections.Counter(); groups={k:[] for k in ['monthly','category','borough','status','agency','daily','category_borough']}; keys=set(); duplicates=0; nullkeys=0; schema=None; total=0; bad=collections.Counter()
for spec in manifest['files']:
    f=SRC/spec['name']; h=hashlib.sha256()
    with f.open('rb') as stream:
        for chunk in iter(lambda:stream.read(1024*1024),b''):h.update(chunk)
    assert h.hexdigest()==spec['sha256'] and f.stat().st_size==spec['size_bytes'],f
    t=pq.read_table(f); df=t.to_pandas(); n=len(df); total+=n
    if schema is None: schema={c:str(t.schema.field(c).type) for c in df.columns}
    assert {c:str(t.schema.field(c).type) for c in df.columns}==schema
    for c in df.columns: missing[c]+=int((df[c].isna()|df[c].fillna('').str.strip().eq('')).sum())
    kk=df.unique_key.dropna(); nullkeys+=n-len(kk); duplicates+=int(kk.isin(keys).sum()+kk.duplicated().sum()); keys.update(kk)
    created=pd.to_datetime(df.created_date,errors='coerce'); closed=pd.to_datetime(df.closed_date,errors='coerce')
    bad['created_parse']+=int((created.isna()&df.created_date.notna()).sum()); bad['closed_parse']+=int((closed.isna()&df.closed_date.notna()).sum())
    start,end=re.findall(r'\d{4}-\d{2}-\d{2}',f.name)
    bad['outside_partition']+=int((~(created.ge(start)&created.lt(end))).sum())
    hours=(closed-created).dt.total_seconds()/3600
    bad['negative_duration']+=int(hours.lt(0).sum()); bad['closed_status_no_date']+=int((df.status.eq('Closed')&closed.isna()).sum()); bad['nonclosed_with_date']+=int((~df.status.eq('Closed')&closed.notna()).sum()); bad['closed_after_2026_09_07']+=int(closed.ge('2026-09-07').sum())
    lat=pd.to_numeric(df.latitude,errors='coerce'); lon=pd.to_numeric(df.longitude,errors='coerce')
    bad['coordinate_outside_broad_nyc_box']+=int((lat.notna()&lon.notna()&~(lat.between(40.4,41)&lon.between(-74.3,-73.6))).sum())
    df['month']=created.dt.strftime('%Y-%m'); df['day']=created.dt.strftime('%Y-%m-%d')
    for c in ['complaint_type','borough','status','agency']: df[c]=df[c].fillna('(missing)').str.strip().replace('', '(missing)')
    df['valid_closed']=df.status.eq('Closed')&hours.ge(0)&closed.lt('2026-09-07')
    df['within24']=df.valid_closed&hours.le(24)
    for name,cols in [('monthly',['month']),('category',['month','complaint_type']),('borough',['month','borough']),('status',['status']),('agency',['month','agency']),('daily',['day']),('category_borough',['month','complaint_type','borough'])]:
        z=df.groupby(cols,dropna=False).agg(records=('unique_key','size'),valid_closed=('valid_closed','sum'),within24=('within24','sum')).reset_index();groups[name].append(z)
    stats.append(dict(file=f.name,rows=n,min_created=str(created.min()),max_created=str(created.max()),bytes=f.stat().st_size,sha256=h.hexdigest()))
    print(f.name,n,flush=True)
assert duplicates==0 and nullkeys==0, 'Review duplicate/missing IDs before interpreting aggregates'
for name,parts in groups.items():
    z=pd.concat(parts); cols=[c for c in z if c not in ['records','valid_closed','within24']]; z=z.groupby(cols,dropna=False,as_index=False)[['records','valid_closed','within24']].sum(); z.to_csv(out/f'{name}.csv',index=False,encoding='utf-8-sig')
pd.DataFrame(stats).to_csv(out/'files.csv',index=False,encoding='utf-8-sig')
pd.DataFrame([dict(field=k,type=schema[k],missing=v,missing_rate=v/total) for k,v in missing.items()]).to_csv(out/'fields.csv',index=False,encoding='utf-8-sig')
summary=dict(total=total,unique_keys=len(keys),duplicates=duplicates,nullkeys=nullkeys,quality=dict(bad),min_created=min(s['min_created'] for s in stats),max_created=max(s['max_created'] for s in stats),file_count=len(stats),bytes=sum(s['bytes'] for s in stats))
(out/'quality.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(summary,ensure_ascii=False,indent=2))
