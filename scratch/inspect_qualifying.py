import json
import requests

with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

headers = {
    'Authorization': f'Bearer {cfg["api"]["bearer_token"]}',
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0'
}

for oid in ['3305113555', '3305117888']:
    r = requests.get('https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking', params={'order_id': oid}, headers=headers)
    print(f"\n=================== ORDER {oid} ===================")
    for i, t in enumerate(r.json().get('trackingTrips', [])):
        ho = t.get('handoverInfo') or {}
        hp = (ho.get('handoverPoint') or {}).get('code')
        hpc = (ho.get('handoverPointCreation') or {}).get('code')
        desc = (t.get('desc') or '').encode('ascii', errors='replace').decode('ascii')
        print(f"Trip {i}: status={t.get('status')} postcode={t.get('postcode')} hp={hp} hpc={hpc} time={t.get('updatedAt')}")
        print(f"   desc={desc}")
