import json
import requests

with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

headers = {
    'Authorization': f'Bearer {cfg["api"]["bearer_token"]}',
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0'
}

order_ids = ['3305035144', '3304889444', '3304890944', '3304956788']

for oid in order_ids:
    r = requests.get('https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking', params={'order_id': oid}, headers=headers, timeout=10)
    print(f"\n=================== ORDER {oid} (HTTP {r.status_code}) ===================")
    if r.status_code == 200:
        d = r.json()
        print(f"Overall Status: {d.get('status')} - {d.get('statusName')}")
        trips = d.get('trackingTrips', [])
        for i, t in enumerate(trips):
            ho = t.get('handoverInfo') or {}
            hp = ho.get('handoverPoint') or {}
            hpc = ho.get('handoverPointCreation') or {}
            print(f"Trip {i}: status={t.get('status')} postcode={t.get('postcode')} hp={hp.get('code')} hpc={hpc.get('code')} time={t.get('updatedAt')}")
            print(f"  desc={t.get('desc')}")
