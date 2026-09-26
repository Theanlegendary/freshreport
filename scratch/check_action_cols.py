import pandas as pd
import json

df = pd.read_excel('cache/latest_detail.xlsx')
b306 = df[df['CURRENT STATUS'].astype(str).str.startswith('306')].copy()

# Look at columns related to 306 actions
action_cols = [
    'ACTION POST OFFICE',
    'ACTION POST OFFICE.1',
    'ACTION POST OFFICE.2',
    'ACTION POST OFFICE.3',
    'ACTION POST OFFICE.4',
    'STATUS 306  AT ORIGIN HUB (FIRST TIME)',
    'STATUS 306 AT STORE / AGENT (LAST TIME)',
    'STATUS 306 AT STORE / AGENT FROM HUB (FIRST TIME)',
    'CURRENT POST OFFICE',
    'RECEIVE POST OFFICE',
    'DELIVERY POST OFFICE'
]

print(f"Total 306 bills: {len(b306)}")
# Check a few records of b306 where CURRENT POST OFFICE is DVCMEGA1 or KRAP001
for idx, r in b306[b306['CURRENT POST OFFICE'].astype(str).str.contains('DVCMEGA|DVMEGA', na=False)].head(10).iterrows():
    vals = {c: str(r[c]) for c in action_cols if pd.notna(r[c])}
    print(f"Order {r['ORDER ID']}: {vals}")
