import json
import concurrent.futures
import requests
import pandas as pd

with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)

headers = {
    'Authorization': f'Bearer {cfg["api"]["bearer_token"]}',
    'Accept': 'application/json, text/plain, */*',
    'User-Agent': 'Mozilla/5.0'
}

df = pd.read_excel('cache/latest_detail.xlsx')
df.columns = [str(c).strip().upper() for c in df.columns]
col_status = next((c for c in df.columns if 'CURRENT STATUS' in c or 'STATUS' in c), 'CURRENT STATUS')
df['sc'] = df[col_status].astype(str).str.extract(r'^(\d{3})')[0]
col_current_po = next((c for c in df.columns if 'CURRENT POST OFFICE' in c), 'CURRENT POST OFFICE')
df['current_po_clean'] = df[col_current_po].astype(str).str.strip().str.upper()

dvc = df[
    (df['current_po_clean'].isin(['DVCMEGA1', 'DVCMEGA', 'DVMEGA']) | df['current_po_clean'].str.contains('DVCMEGA|DVMEGA', na=False)) &
    (~df['current_po_clean'].isin(['MEGA1'])) &
    (df['sc'] == '306')
].copy()

order_ids = dvc['ORDER ID'].dropna().astype(str).str.strip().tolist()
print(f"Total candidate orders: {len(order_ids)}")

session = requests.Session()
session.headers.update(headers)

def check_order(oid):
    try:
        url = 'https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking'
        r = session.get(url, params={'order_id': oid}, timeout=10)
        if r.status_code != 200:
            return {'oid': oid, 'error': f"HTTP {r.status_code}"}
        data = r.json()
        trips = data.get('trackingTrips', [])
        if not trips:
            return {'oid': oid, 'status': 'no_trips'}
        
        # Latest trip (trip 0)
        latest_trip = trips[0]
        latest_status = str(latest_trip.get('status', ''))
        latest_postcode = str(latest_trip.get('postcode', '') or '').upper()
        latest_ho = latest_trip.get('handoverInfo') or {}
        latest_hpc = str(latest_ho.get('handoverPointCreation', {}).get('code', '') or '').upper()
        latest_hp = str(latest_ho.get('handoverPoint', {}).get('code', '') or '').upper()
        
        # Check if passed MEGA anywhere in the trips
        # Passing MEGA means a trip had postcode == 'MEGA1' or hp == 'MEGA1' or hpc == 'MEGA1'
        # Specifically, handover to or from MEGA1
        passed_mega = False
        mega_trip_idx = None
        for idx, t in enumerate(trips):
            t_pc = str(t.get('postcode', '') or '').upper()
            t_ho = t.get('handoverInfo') or {}
            t_hp = str(t_ho.get('handoverPoint', {}).get('code', '') or '').upper()
            t_hpc = str(t_ho.get('handoverPointCreation', {}).get('code', '') or '').upper()
            if 'MEGA' in t_pc or 'MEGA' in t_hp or 'MEGA' in t_hpc:
                # specifically MEGA1 hub
                if 'MEGA1' in (t_pc, t_hp, t_hpc):
                    passed_mega = True
                    mega_trip_idx = idx
                    break
        
        # Is latest action dvdriver?
        # DV driver codes typically start with 'DVC' or 'DV' or are driver codes
        # In our case: latest_postcode in ('DVCMEGA1', 'DVCMEGA', 'DVMEGA') or latest_postcode.startswith('DVC') or latest_postcode.startswith('DV')
        # AND latest handover was from MEGA1: latest_hpc == 'MEGA1' or 'MEGA1' in latest_hpc
        is_latest_dvdriver = (
            latest_postcode.startswith('DVC') or latest_postcode.startswith('DV')
        )
        
        # Is latest action on a post office?
        # A post office is typically 7 chars like KRAP001, PNPP005, KANP001, etc. (ending with P001, A001, S001, etc.)
        # Or postcode does NOT start with DVC/DV/MEGA
        is_latest_po = not is_latest_dvdriver and 'MEGA' not in latest_postcode
        
        # Does latest action receive handover from MEGA1?
        from_mega = 'MEGA1' in latest_hpc
        
        return {
            'oid': oid,
            'latest_status': latest_status,
            'latest_postcode': latest_postcode,
            'latest_hpc': latest_hpc,
            'latest_hp': latest_hp,
            'passed_mega': passed_mega,
            'mega_trip_idx': mega_trip_idx,
            'is_latest_dvdriver': is_latest_dvdriver,
            'is_latest_po': is_latest_po,
            'from_mega': from_mega,
            'desc': (latest_trip.get('desc') or '').encode('ascii', errors='replace').decode('ascii')
        }
    except Exception as e:
        return {'oid': oid, 'error': str(e)}

results = []
with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
    futures = [executor.submit(check_order, oid) for oid in order_ids]
    for f in concurrent.futures.as_completed(futures):
        results.append(f.result())

df_res = pd.DataFrame(results)
print("\n--- RESULTS SUMMARY ---")
print("Latest status counts:")
print(df_res['latest_status'].value_counts())
print("\nLatest postcode counts:")
print(df_res['latest_postcode'].value_counts())
print("\nPassed MEGA:")
print(df_res['passed_mega'].value_counts())
print("\nLatest is dvdriver:")
print(df_res['is_latest_dvdriver'].value_counts())
print("\nLatest from_mega (handover from MEGA1):")
print(df_res['from_mega'].value_counts())

# Filter according to user criteria:
# 1. Status 306
# 2. Passed MEGA in the middle (mega_trip_idx > 0, so MEGA was before the latest trip!)
# 3. Latest action is dvdriver (DVCMEGA1, DVC..., not post office)
# 4. Latest action is NOT on post office
print("\n--- DVDRIVER BILLS ---")
print(df_res[df_res['is_latest_dvdriver'] == True][['oid', 'latest_status', 'latest_postcode', 'latest_hpc', 'from_mega', 'passed_mega', 'mega_trip_idx', 'desc']])

print("\n--- S306 BILLS ---")
print(df_res[df_res['latest_status'] == 'S306'][['oid', 'latest_status', 'latest_postcode', 'latest_hpc', 'from_mega', 'passed_mega', 'mega_trip_idx', 'desc']])

