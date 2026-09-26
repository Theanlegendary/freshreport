import sys
sys.path.append('.')
import pandas as pd
import shipments_tomorrow

test_cases = ['ALL', 'Zone 2', 'SPEP001', 'KRAP001']

for tgt in test_cases:
    out_file = f'scratch/test_{tgt.replace(" ", "_")}.xlsx'
    bills, weight = shipments_tomorrow.build_shipments_tomorrow_report('cache/latest_detail.xlsx', out_file, target_label=tgt)
    print(f"\nTarget '{tgt}': {bills} bills, {weight} g ({weight/1000:.2f} kg)")
    
    df_res = pd.read_excel(out_file, sheet_name='base')
    col_ord = next((c for c in df_res.columns if 'ORDER' in str(c).upper()), None)
    order_ids = df_res[col_ord].dropna().astype(str).tolist() if col_ord else []
    
    # Check forbidden bills
    forbidden = {
        '3305060988': 'Status 302 at origin store PNPP008',
        '3305035144': 'Status 306 arrived at KRAP001 post office',
        '3305072300': 'Status 306 origin store to MEGA (not passed MEGA)'
    }
    
    for f_oid, reason in forbidden.items():
        if f_oid in order_ids:
            print(f"  [FAIL] {f_oid} ({reason}) was INCORRECTLY included in '{tgt}' report!")
        else:
            print(f"  [OK] {f_oid} ({reason}) is excluded from '{tgt}' report.")
