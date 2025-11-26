# Cell 22 - Optimized
import pandas as pd
import os
from datetime import datetime
import time
import requests

# Ensure reports_dir exists from previous context, or fallback
if 'reports_dir' not in globals():
    reports_dir = os.path.join(os.getcwd(), 'reports')

stats_xlsx = os.path.join(reports_dir, "Library_Statistics_Latest.xlsx")
results_xlsx = os.path.join(reports_dir, f"Empty_Folder_Deletion_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

if not os.path.exists(stats_xlsx):
    print(f"Library statistics file not found: {stats_xlsx}")
else:
    stats_df = pd.read_excel(stats_xlsx, sheet_name=0) # Reads first sheet by default
    print(f"Loaded {len(stats_df)} libraries from statistics report.")

    api_call_counter = [0]
    deletion_results = []

    def safe_graph_get(url):
        resp = graph_get(url)
        api_call_counter[0] += 1
        # Increased pause threshold to 200 to reduce console noise
        if api_call_counter[0] % 200 == 0:
            print(f"  ⏸️  Pausing after {api_call_counter[0]} API calls...")
            time.sleep(2)
        
        if getattr(resp, 'status_code', None) == 429:
            retry_after = int(resp.headers.get('Retry-After', 10))
            print(f"  ⚠️ Throttled! Waiting {retry_after}s...")
            time.sleep(retry_after)
            return safe_graph_get(url)
        return resp

    def safe_delete(url, headers):
        # Assuming get_auth_headers() is available from previous cells
        resp = requests.delete(url, headers=headers)
        api_call_counter[0] += 1
        if api_call_counter[0] % 200 == 0:
            print(f"  ⏸️  Pausing after {api_call_counter[0]} API calls...")
            time.sleep(2)
        
        if getattr(resp, 'status_code', None) == 429:
            retry_after = int(resp.headers.get('Retry-After', 10))
            print(f"  ⚠️ Throttled! Waiting {retry_after}s...")
            time.sleep(retry_after)
            return safe_delete(url, headers)
        return resp

    # Recursive function that now grabs childCount in the first pass
    def list_folders_recursive(drive_id, item_id="root", depth=0, folders=None):
        if folders is None:
            folders = []
        
        # Select specific fields including 'folder' which contains childCount
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$top=5000&$select=id,name,folder,parentReference,size"
        
        resp = safe_graph_get(url)
        if getattr(resp, 'status_code', None) != 200:
            return folders
            
        items = resp.json().get('value', [])
        for item in items:
            if 'folder' in item:
                # Store the folder with its childCount
                child_count = item['folder'].get('childCount', 0)
                folders.append({
                    'id': item['id'],
                    'name': item['name'],
                    'parent_id': item.get('parentReference', {}).get('id', ''),
                    'depth': depth,
                    'size': item.get('size', 0),
                    'child_count': child_count
                })
                # Recurse
                list_folders_recursive(drive_id, item['id'], depth+1, folders)
        return folders

    for idx, row in stats_df.iterrows():
        site_name = row.get('site_name', 'Unknown')
        library_name = row.get('library_name', 'Unknown')
        library_url = row.get('library_url', '')
        site_id = row.get('site_id')
        library_id = row.get('library_id')

        print(f"\nProcessing library {idx+1}/{len(stats_df)}: {library_name} ({site_name})")

        drive_id = None
        if site_id and library_id and not pd.isna(site_id) and not pd.isna(library_id):
            drive_resp = safe_graph_get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{library_id}/drive")
            if getattr(drive_resp, 'status_code', None) == 200:
                drive_id = drive_resp.json().get('id')

        if not drive_id:
            print("  Skipping: drive_id not available.")
            continue

        # 1. Get all folders (Optimized: gets childCount in this single pass)
        all_folders = list_folders_recursive(drive_id)
        print(f"  Found {len(all_folders)} total folders.")

        # 2. Filter for empty folders using the metadata we just got
        # No extra API calls needed here
        empty_folders = [f for f in all_folders if f['child_count'] == 0]
        print(f"  Found {len(empty_folders)} empty folders (childCount == 0).")

        # 3. Sort by depth (deepest first) to delete cleanly
        empty_folders_sorted = sorted(empty_folders, key=lambda x: -x['depth'])

        deleted_count = 0
        auth_headers = get_auth_headers() # Ensure this function is defined in previous cells

        for folder in empty_folders_sorted:
            del_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder['id']}"
            del_resp = safe_delete(del_url, auth_headers)
            
            if getattr(del_resp, 'status_code', None) == 204:
                print(f"    Deleted: {folder['name']}")
                deleted_count += 1
            else:
                print(f"    Failed: {folder['name']} (Status {getattr(del_resp, 'status_code', None)})")

        # 4. Calculate stats without re-scanning
        folders_after = len(all_folders) - deleted_count
        print(f"  New folder count (calculated): {folders_after}")

        deletion_results.append({
            'site_name': site_name,
            'library_name': library_name,
            'library_url': library_url,
            'folders_before': len(all_folders),
            'empty_folders_found': len(empty_folders),
            'folders_deleted': deleted_count,
            'folders_after': folders_after
        })

    # Save results
    if deletion_results:
        results_df = pd.DataFrame(deletion_results)
        results_df.to_excel(results_xlsx, index=False)
        print(f"\n✅ Deletion results saved: {results_xlsx}")
        display(results_df)
    else:
        print("\n✅ No deletion results to save.")
