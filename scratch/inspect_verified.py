import json
import requests

with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

headers = {'Authorization': f'Bearer {cfg["api"]["bearer_token"]}'}

for oid in ['3305008111', '3304625177']:
    r = requests.get('https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking', params={'order_id': oid}, headers=headers)
    trips = r.json().get('trackingTrips', [])
    print(f"\n--- ORDER {oid} ---")
    for i, t in enumerate(trips[:4]):
        ho = t.get('handoverInfo') or {}
        hpc = (ho.get('handoverPointCreation') or {}).get('code')
        desc = (t.get('desc') or '').encode('ascii', errors='replace').decode('ascii')
        print(f"Trip {i}: status={t.get('status')} postcode={t.get('postcode')} hpc={hpc} time={t.get('updatedAt')}")
        print(f"  desc={desc}")
