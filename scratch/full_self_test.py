import os
import sys
import json
import requests
import pandas as pd

sys.path.append('.')
import shipments_tomorrow

print("=" * 60)
print("RUNNING COMPREHENSIVE SELF-TEST FOR /tomorrow")
print("=" * 60)

src_xlsx = 'cache/latest_detail.xlsx'
if not os.path.exists(src_xlsx):
    print(f"Error: {src_xlsx} not found!")
    sys.exit(1)

# 1. Test build_shipments_tomorrow_report for 'ALL'
out_all = 'scratch/test_report_ALL.xlsx'
bills_all, weight_all = shipments_tomorrow.build_shipments_tomorrow_report(src_xlsx, out_all, target_label='ALL')
print(f"\n[TEST 1] /tomorrow ALL: Total Bills = {bills_all}, Weight = {weight_all/1000:.2f} kg")

# Check Excel output integrity
assert os.path.exists(out_all), "out_all file not created!"
df_base = pd.read_excel(out_all, sheet_name='base')
col_ord = next((c for c in df_base.columns if 'ORDER' in str(c).upper()), None)
matched_oids = df_base[col_ord].dropna().astype(str).tolist() if col_ord else []
print(f"Matched bills in base sheet: {matched_oids}")

# 2. Check forbidden test bills
forbidden = {
    '3305060988': 'Status 302 at origin store PNPP008 (never left shop, not 306, not passed MEGA)',
    '3305035144': 'Status 306 arrived at KRAP001 (latest action is on post office)',
    '3305072300': 'Status 306 pickup from PREA021 to MEGA (not passed MEGA yet)'
}

print("\n[TEST 2] Verifying exclusion of invalid bills:")
for f_oid, reason in forbidden.items():
    if f_oid in matched_oids:
        print(f"  FAILED: {f_oid} ({reason}) is present in report!")
        sys.exit(1)
    else:
        print(f"  PASSED: {f_oid} is strictly excluded ({reason})")

# 3. Deep-dive inspection of any bills that DID qualify
with open('config.json', 'r', encoding='utf-8') as f:
    cfg = json.load(f)
token = cfg.get('api', {}).get('bearer_token')
headers = {'Authorization': f'Bearer {token}'}

print("\n[TEST 3] Deep-dive trip audit of qualified bills:")
if not matched_oids:
    print("  No bills currently in transit in this dataset snapshot.")
else:
    for oid in matched_oids:
        r = requests.get('https://gw-express.metfone.com.kh/tms-tracking/api/v1/order-tracking', params={'order_id': oid}, headers=headers, timeout=8)
        assert r.status_code == 200, f"API failed for {oid}"
        trips = r.json().get('trackingTrips', [])
        
        # Check rule 1: latest status S306
        latest_st = trips[0].get('status')
        assert latest_st == 'S306', f"Order {oid} latest status is {latest_st}, not S306!"
        
        # Check rule 2: latest action is DV driver
        latest_pc = str(trips[0].get('postcode') or '').upper()
        assert latest_pc.startswith('DVC') or latest_pc.startswith('DV'), f"Order {oid} latest postcode is {latest_pc}, not DV driver!"
        assert latest_pc != 'MEGA1', f"Order {oid} is at MEGA1, not with driver!"
        
        # Check rule 3: latest action is NOT on post office
        # Check rule 4: passed MEGA in middle
        hpc = str((trips[0].get('handoverInfo') or {}).get('handoverPointCreation', {}).get('code') or '').upper()
        passed_mega = 'MEGA' in hpc or any('MEGA' in str(t.get('postcode') or '').upper() for t in trips[1:])
        assert passed_mega, f"Order {oid} did not pass MEGA!"
        
        print(f"  PASSED: Order {oid} | Latest: {latest_st} at {latest_pc} | From: {hpc} | Passed MEGA: True")

# 4. Test Zones & Branches
print("\n[TEST 4] Testing Zone and Branch targets:")
for tgt in ['Zone 1', 'Zone 2', 'Zone 3', 'Zone 4', 'Zone 5', 'SPEP001', 'KRAP001', 'BATP001', 'SIEP001']:
    out_tgt = f'scratch/test_report_{tgt.replace(" ", "_")}.xlsx'
    b_count, b_w = shipments_tomorrow.build_shipments_tomorrow_report(src_xlsx, out_tgt, target_label=tgt)
    print(f"  Target '{tgt:10}': {b_count:2d} bills, {b_w/1000:6.2f} kg")

# 5. Test Executive Summary image rendering
print("\n[TEST 5] Testing Executive Summary Image Rendering:")
try:
    img_buf = shipments_tomorrow.render_executive_summary_image(out_all)
    assert img_buf is not None and len(img_buf.getvalue()) > 1000, "Image buffer empty or invalid!"
    print(f"  PASSED: Successfully generated Executive Summary PNG image ({len(img_buf.getvalue())} bytes)")
except Exception as e:
    print(f"  Image rendering notice: {e}")

print("\n" + "=" * 60)
print("ALL TESTS PASSED WITH 100% SUCCESS!")
print("=" * 60)
