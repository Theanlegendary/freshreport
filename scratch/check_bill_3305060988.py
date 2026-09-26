import json
import requests
import pandas as pd

with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

headers = {
    'Authorization': f'Bearer {cfg["api"]["bearer_token"]}',
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0'
}

oid = '3305060988'

# 1. Check API tracking
r = requests.get('https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking', params={'order_id': oid}, headers=headers, timeout=10)
print(f"Tracking API HTTP: {r.status_code}")
if r.status_code == 200:
    data = r.json()
    with open('scratch/bill_3305060988_api.json', 'w', encoding='utf-8') as f_out:
        json.dump(data, f_out, indent=2, ensure_ascii=False)
    print("API Overall Status:", data.get('status'), data.get('statusName'))
    trips = data.get('trackingTrips', [])
    print(f"Total trips: {len(trips)}")
    for i, t in enumerate(trips):
        ho = t.get('handoverInfo') or {}
        hp = (ho.get('handoverPoint') or {}).get('code')
        hpc = (ho.get('handoverPointCreation') or {}).get('code')
        desc = (t.get('desc') or '').encode('ascii', errors='replace').decode('ascii')
        print(f"Trip {i}: status={t.get('status')} postcode={t.get('postcode')} hp={hp} hpc={hpc} time={t.get('updatedAt')}")
        print(f"   desc={desc}")

# 2. Check Excel cache
try:
    df = pd.read_excel('cache/latest_detail.xlsx')
    row = df[df['ORDER ID'].astype(str) == oid]
    print(f"\nFound in cache/latest_detail.xlsx: {len(row)} rows")
    if len(row) > 0:
        for k, v in row.iloc[0].items():
            if pd.notna(v):
                print(f"  {k}: {str(v).encode('ascii', errors='replace').decode('ascii')}")
except Exception as e:
    print(f"Excel error: {e}")
