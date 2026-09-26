import sys, os
sys.path.insert(0, r'c:\Users\DELL\Desktop\daily_push')
os.chdir(r'c:\Users\DELL\Desktop\daily_push')

from total_pending_report import process_pending_data, DAY_COLS, DELAY_COLS, DELAY_COL_NAME

src = r'cache/latest_detail.xlsx'
if not os.path.exists(src):
    print('ERROR: No cache file found at', src)
    sys.exit(1)

print("Loading pending data...")
summary_df, grand_total, df_branch = process_pending_data(src)

print('\n=== GRAND TOTAL ===')
print(f'Total:        {grand_total["Total"]:,}')
print(f'Servicepoint: {grand_total["Servicepoint"]:,}')
print(f'Showroom:     {grand_total["Showroom"]:,}')
print(f'Agent:        {grand_total["Agent"]:,}')
fac_sum = grand_total["Servicepoint"] + grand_total["Showroom"] + grand_total["Agent"]
day_sum = sum(grand_total.get(d, 0) for d in DAY_COLS)
print(f'SP+SR+AG sum: {fac_sum:,}  {"OK" if fac_sum == grand_total["Total"] else "MISMATCH!"}')
print(f'Day buckets sum: {day_sum:,}  {"OK" if day_sum == grand_total["Total"] else "MISMATCH!"}')
print(f'Delay >=3D:   {grand_total.get(DELAY_COL_NAME, 0):,}')
print(f'Delay calc:   {sum(grand_total.get(d,0) for d in DELAY_COLS):,}')

print('\n=== SUMMARY TABLE ===')
print(summary_df.to_string(index=False))

print('\n=== PER-BRANCH CROSS-CHECK ===')
errors = 0
for _, r in summary_df.iterrows():
    facility_sum = r['Servicepoint'] + r['Showroom'] + r['Agent']
    day_sum = sum(r.get(d, 0) for d in DAY_COLS)
    ok_f = facility_sum == r['Total']
    ok_d = day_sum == r['Total']
    status = 'OK' if (ok_f and ok_d) else 'ERROR'
    if status == 'ERROR':
        errors += 1
        print(f"  {r['Branch']}: Total={r['Total']} | FacSum={facility_sum} {'OK' if ok_f else 'MISMATCH'} | DaySum={day_sum} {'OK' if ok_d else 'MISMATCH'}")

if errors == 0:
    print("  All branches OK - totals match!")
else:
    print(f"\n  {errors} branches have mismatches!")

print('\n=== RAW DETAIL STATS ===')
print(f'Total rows in df_branch: {len(df_branch):,}')
print(f'Facility types: {df_branch["Facility_Type"].value_counts().to_dict()}')
print(f'Age buckets:    {df_branch["Age_Bucket"].value_counts().sort_index().to_dict()}')
