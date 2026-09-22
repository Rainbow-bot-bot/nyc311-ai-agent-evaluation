from pathlib import Path
import pandas as pd
import json
P=Path(__file__).resolve().parent/'汇总'
quality=json.loads((P/'quality.json').read_text(encoding='utf-8'))
monthly=pd.read_csv(P/'monthly.csv'); cat=pd.read_csv(P/'category.csv'); bor=pd.read_csv(P/'borough.csv')
def compare(df,col):
    parts=[]
    for year in [2025,2026]:
        z=df[df.month.between(f'{year}-01',f'{year}-08')].groupby(col).records.sum().rename(str(year));parts.append(z)
    z=pd.concat(parts,axis=1).fillna(0).astype(int)
    z['change']=z['2026']-z['2025'];z['growth']=z['change']/z['2025'].replace(0,float('nan'));z['share26']=z['2026']/z['2026'].sum();z['contribution']=z['change']/z['change'].sum()
    return z.sort_values('change',ascending=False).reset_index()
c=compare(cat,'complaint_type');b=compare(bor,'borough');a=compare(pd.read_csv(P/'agency.csv'),'agency')
c.to_csv(P/'category_compare.csv',index=False,encoding='utf-8-sig');b.to_csv(P/'borough_compare.csv',index=False,encoding='utf-8-sig');a.to_csv(P/'agency_compare.csv',index=False,encoding='utf-8-sig')
daily=pd.read_csv(P/'daily.csv'); dates=pd.to_datetime(daily.day)
checks={name:int(pd.read_csv(P/f'{name}.csv').records.sum()) for name in ['monthly','category','borough','status','agency','daily','category_borough']}
assert all(v==quality['total'] for v in checks.values())
assert len(pd.date_range(dates.min(),dates.max()).difference(dates))==0
checks['missing_calendar_days']=0
checks['manifest_hashes_matched']=quality['file_count']
(P/'checks.json').write_text(json.dumps(checks,ensure_ascii=False,indent=2),encoding='utf-8')
print('QUALITY',quality)
print('MONTHLY\n',monthly.to_string(index=False))
print('CATEGORY\n',c.head(20).to_string(index=False))
print('DECLINES\n',c.tail(10).to_string(index=False))
print('BOROUGH\n',b.to_string(index=False))
print('AGENCY\n',a.to_string(index=False))
print('MISSING\n',pd.read_csv(P/'fields.csv').to_string(index=False))
print('STATUS\n',pd.read_csv(P/'status.csv').to_string(index=False))
payload={'quality':quality,'monthly':monthly.to_dict('records'),'categories':c.fillna('不适用').to_dict('records'),'boroughs':b.fillna('不适用').to_dict('records'),'fields':pd.read_csv(P/'fields.csv').to_dict('records'),'files':pd.read_csv(P/'files.csv').to_dict('records'),'checks':checks}
(P/'workbook_data.json').write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
