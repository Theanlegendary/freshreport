import pandas as pd
import json

df = pd.read_excel('cache/latest_detail.xlsx')
oids = ['3305113555', '3305117888', '3305072300', '3305035144']
rows = df[df['ORDER ID'].astype(str).isin(oids)]

for _, r in rows.iterrows():
    print(f"\n=================== ORDER {r['ORDER ID']} ===================")
    for k, v in r.items():
        if pd.notna(v) and any(w in k for w in ['POST', 'STATUS', 'ACTION', 'CURRENT']):
            safe_v = str(v).encode('ascii', errors='replace').decode('ascii')
            print(f"  {k}: {safe_v}")
