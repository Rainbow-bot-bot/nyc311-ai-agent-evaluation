from pathlib import Path
import pandas as pd, json, duckdb
ROOT=Path(__file__).parent; E=ROOT/'evidence'
d=pd.read_csv(E/'dimensions.csv'); d['year']=d.mth.str[:4].astype(int); d['month']=d.mth.str[5:7].astype(int)
s=d[(d.year.isin([2025,2026]))&(d.month<=8)].copy()
def compare(frame,cols,name):
    p=frame.pivot_table(index=cols,columns='year',values='n',aggfunc='sum',fill_value=0).reindex(columns=[2025,2026],fill_value=0)
    p.columns=['base','current']; p['delta']=p.current-p.base; p['growth']=p.current/p.base.replace(0,float('nan'))-1
    p['contribution']=p.delta/(s[s.year==2026].n.sum()-s[s.year==2025].n.sum())
    p=p.sort_values('delta',ascending=False).reset_index(); p.to_csv(E/(name+'.csv'),index=False,encoding='utf-8-sig'); print(name,p.head(20).to_string(index=False)); return p
for col in ['complaint_type','agency','borough','channel','month']:
    compare(s,[col],'compare_'+col)
compare(s[s.complaint_type=='HEAT/HOT WATER'],['month'],'heat_month')
compare(s[s.complaint_type=='Illegal Parking'],['month'],'parking_month')
compare(s,['complaint_type','channel'],'category_channel')
compare(s,['complaint_type','borough'],'category_borough')
