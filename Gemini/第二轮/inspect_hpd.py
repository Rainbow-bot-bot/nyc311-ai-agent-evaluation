import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_csv('data/hpd_heat_resolutions.csv')
for idx, row in df.head(10).iterrows():
    desc = str(row['resolution_description'])
    print(f"Rank {idx+1} ({row['pct']}%, cnt={row['cnt']}): {desc[:140]}...")
