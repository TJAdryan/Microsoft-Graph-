import pandas as pd
import os
from datetime import datetime
import time
import requests

# ================= CONFIGURATION =================
# 1. REVIEW MODE: Set to False first. It will only generate an Excel list.
# 2. DELETE MODE: Check the Excel list. If it looks good, set to True and run again.
DELETE_MODE = False 

# Protection: Never delete folders at the root of the library (Depth 0)
SKIP_TOP_LEVEL = False
# =================================================

stats_xlsx = os.path.join(reports_dir, "Library_Statistics_Latest.xlsx")
review_xlsx = os.path.join(reports_dir, "Empty_Folders_For_Review.xlsx")

# --- Helper Functions ---
def safe_graph_get(url):
    resp = graph_get(url) # Uses your existing graph_get function
    api_call_counter[0] += 1
    if api_call_counter[0] % 100 == 0: # Optimized: Pausing less frequently
        print(f"  ⏸️  Pausing after {api_call_counter[0]} API calls...")
        time.sleep(1)
    if getattr(resp, 'status_code', None) == 429:
        retry_after = int(resp.headers.get('Retry-After', 10))
        print(f"  ⚠️ Throttled! Waiting {retry_after}s...")
        time.sleep(retry_after)
        return safe_graph_get(url)
    return resp

def safe_delete(url, headers):
    resp = requests.delete(url, headers=headers)
    api_call_counter[0] += 1
    if getattr(resp, 'status_code', None) == 429:
        retry_after = int(resp.headers.get('Retry-After', 10))
        time.sleep(retry_after)
        return safe_delete(url, headers)
    return resp

def list_empty_folders_fast(drive_id, item_id="root", depth=0, empty_folders=None):
    """
    Optimized recursive function. 
    Checks 'childCount' immediately to avoid extra API calls.
    """
    if empty_folders is None:
        empty_folders = []
    
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$select=id,name,folder,parentReference&$top=5000"
    resp = safe_graph_get(url)
    
    if getattr(resp, 'status_code', None) != 200:
        return empty_folders
        
    items = resp.json().get('value', [])
    
    for item in items:
        # We only care about folders
        if 'folder' in item:
            child_count = item['folder'].get('childCount', 0)
            
            # OPTIMIZATION: Check childCount directly
            if child_count == 0:
                # Found an empty folder!
                if SKIP_TOP_LEVEL and depth == 0:
                    continue # Skip if it is a root folder
                
                empty_folders.append({
                    'id': item['id'],
                    'name': item['name'],
                    'depth': depth,
                    'drive_id': drive_id, # Store for deletion later
                    'parent_id': item.get('parentReference', {}).get('id', '')
                })
            else:
                # If not empty, dive deeper
                list_empty_folders_fast(drive_id, item['id'], depth+1, empty_folders)
                
    return empty_folders

# ================= MAIN LOGIC =================
if not os.path.exists(stats_xlsx):
    print(f"Stats file not found: {stats_xlsx}")
else:
    api_call_counter = [0]
    
    if not DELETE_MODE:
        # --- SCANNING PHASE (FAST) ---
        print("\n🚀 FAST SCAN MODE: Identifying empty folders...")
        stats_df = pd.read_excel(stats_xlsx, sheet_name="All_Libraries")
        all_empty_folders = []

        for idx, row in stats_df.iterrows():
            site_name = row['site_name']
            library_name = row['library_name']
            
            print(f"Scanning {idx+1}/{len(stats_df)}: {library_name}")

            # Get Drive ID
            drive_id = None
            if row.get('site_id') and row.get('library_id'):
                drive_resp = safe_graph_get(f"https://graph.microsoft.com/v1.0/sites/{row['site_id']}/lists/{row['library_id']}/drive")
                if getattr(drive_resp, 'status_code', None) == 200:
                    drive_id = drive_resp.json().get('id')

            if drive_id:
                # Use the FAST function
                found = list_empty_folders_fast(drive_id)
                if found:
                    for f in found:
                        f['site_name'] = site_name
                        f['library_name'] = library_name
                    all_empty_folders.extend(found)
                    print(f"  Found {len(found)} empty folders.")

        # Save to Excel for Review
        if all_empty_folders:
            results_df = pd.DataFrame(all_empty_folders)
            # Sort by depth descending (deepest first) just in case
            results_df = results_df.sort_values('depth', ascending=False)
            results_df.to_excel(review_xlsx, index=False)
            print(f"\n✅ Scan Complete. Review file created: {review_xlsx}")
            display(results_df.head())
        else:
            print("\n✅ Scan Complete. No empty folders found.")

    else:
        # --- DELETION PHASE ---
        print("\n🗑️ DELETION MODE: Processing review file...")
        
        if not os.path.exists(review_xlsx):
            print(f"❌ Review file not found: {review_xlsx}")
            print("Run with DELETE_MODE = False first.")
        else:
            df_to_delete = pd.read_excel(review_xlsx)
            print(f"Loaded {len(df_to_delete)} folders to delete.")
            
            deleted_count = 0
            for idx, row in df_to_delete.iterrows():
                # Double check safety
                if SKIP_TOP_LEVEL and row['depth'] == 0:
                    print(f"Skipping top-level folder: {row['name']}")
                    continue

                print(f"Deleting: {row['site_name']} -> {row['name']}")
                del_url = f"https://graph.microsoft.com/v1.0/drives/{row['drive_id']}/items/{row['id']}"
                resp = safe_delete(del_url, get_auth_headers())
                
                if getattr(resp, 'status_code', None) == 204:
                    deleted_count += 1
                else:
                    print(f"  ❌ Failed: {getattr(resp, 'status_code', None)}")

            print(f"\n✅ Operation Complete. Deleted {deleted_count} folders.")
