import json
import requests

with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

headers = {
    'Authorization': f'Bearer {cfg["api"]["bearer_token"]}',
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0'
}

order_ids = ['3305043133', '3305045344', '3305057355', '3305068299', '3305070899']

for oid in order_ids:
    r = requests.get('https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking', params={'order_id': oid}, headers=headers, timeout=10)
    print(f"\n=================== ORDER {oid} ===================")
    if r.status_code == 200:
        d = r.json()
        trips = d.get('trackingTrips', [])
        for i, t in enumerate(trips):
            ho = t.get('handoverInfo') or {}
            hp = ho.get('handoverPoint') or {}
            hpc = ho.get('handoverPointCreation') or {}
            st = t.get('status')
            postcode = t.get('postcode')
            time_str = t.get('updatedAt')
            hp_code = hp.get('code')
            hpc_code = hpc.get('code')
            desc = t.get('desc', '').encode('ascii', errors='replace').decode('ascii')
            print(f"Trip {i}: status={st} postcode={postcode} hp={hp_code} hpc={hpc_code} time={time_str}")
            print(f"  desc={desc}")
