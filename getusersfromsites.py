
import os
#Requires api auth to be set up
#This is a process to gather information on several sites and then the users and their permissions on those sites. 
# 1. Clean Dataframe
# 2. We assume 'df' is already loaded: df = pd.read_csv('...')  Have list of sites that are needed 
# 3. First gets required fields for groups then gets all group, members and owners.  User output xlsx marks all non member/owners as visitors
# 4. 
print(f"Loaded {len(df)} groups from CSV.")


df_sites = df[['Group ID', 'Group name']].copy()
df_sites.columns = ['group_id', 'name']

# Initialize webUrl column
df_sites['webUrl'] = ''

print("Fetching SharePoint URLs from Graph API...")

# 2. Loop through groups to get WebURL
for index, row in df_sites.iterrows():
    # Extract ID inside the loop
    group_id = str(row['group_id']).strip()
    group_name = row['name']
    
    # Endpoint: GET /groups/{id}/sites/root
    endpoint = f"https://graph.microsoft.com/v1.0/groups/{group_id}/sites/root"
    
    try:
        # Uses your existing graph_get function from Cell 1
        resp = graph_get(endpoint)
        
        if getattr(resp, 'status_code', None) == 200:
            data = resp.json()
            web_url = data.get('webUrl')
            df_sites.at[index, 'webUrl'] = web_url
            print(f"   Success: {group_name}: {web_url}")
        else:
            print(f"   Warning: {group_name}: Site not found (Status: {getattr(resp, 'status_code', 'N/A')})")
    except Exception as e:
        print(f"   Error: {group_name}: {str(e)}")

# 3. Save to Static Excel File
output_dir = r"C:\Python_Scripts\api_calls\reports"
os.makedirs(output_dir, exist_ok=True)
# Naming it 'Sites_For_Permissioning.xlsx' for use in PowerShell script
output_path = os.path.join(output_dir, "Sites_For_Permissioning.xlsx")

df_sites.to_excel(output_path, index=False)

print("\n" + "="*60)
print(f"Static file created: {output_path}")
print(f"Columns: {list(df_sites.columns)}")
print("="*60)
print(df_sites.shape)
display(df_sites.head())

# Cell 3: Generate Site Access Matrix from Static List
# Reads: Sites_For_Permissioning.xlsx
# Outputs: Site_Access_Matrix.xlsx

print(f"Generating Matrix at: {datetime.now().strftime('%I:%M %p')}")
print("=" * 60)

# 1. Load the Static Source of Truth
static_file_path = r"C:\Python_Scripts\api_calls\reports\Sites_For_Permissioning.xlsx"

if not os.path.exists(static_file_path):
    print(f"File not found: {static_file_path}")
    print("Please run the previous cell to generate the static site list.")
else:
    df_target = pd.read_excel(static_file_path)
    print(f"Loaded {len(df_target)} sites from static file.")

    all_site_users = []

    # 2. Iterate through the specific sites in the file
    for idx, row in df_target.iterrows():
        # UPDATED: Use 'group_id' instead of 'id'
        site_id = str(row['group_id']).strip()
        site_name = row['name']
        site_url = row.get('webUrl', '')
        
        print(f"\nProcessing {idx+1}/{len(df_target)}: {site_name}")
        
        group_id = site_id 
        
        # Get Owners
        owners_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/owners"
        owners_resp = graph_get(owners_url)
        
        if getattr(owners_resp, 'status_code', None) == 200:
            data = owners_resp.json()
            users = data.get('value', [])
            print(f"  Found {len(users)} owners")
            for u in users:
                all_site_users.append({
                    'user_email': u.get('mail', u.get('userPrincipalName')),
                    'user_displayName': u.get('displayName'),
                    'user_id': u.get('id'),
                    'site_name': site_name,
                    'role': 'Owner'
                })
        else:
            print(f"  Failed to get owners (Status: {getattr(owners_resp, 'status_code', 'N/A')})")

        # Get Members
        members_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members"
        members_resp = graph_get(members_url)
        
        if getattr(members_resp, 'status_code', None) == 200:
            data = members_resp.json()
            users = data.get('value', [])
            print(f"  Found {len(users)} members")
            for u in users:
                all_site_users.append({
                    'user_email': u.get('mail', u.get('userPrincipalName')),
                    'user_displayName': u.get('displayName'),
                    'user_id': u.get('id'),
                    'site_name': site_name,
                    'role': 'Member'
                })
        else:
            print(f"  Failed to get members (Status: {getattr(members_resp, 'status_code', 'N/A')})")

    # 3. Create Pivot Table
    if all_site_users:
        long_df = pd.DataFrame(all_site_users)
        long_df['role_rank'] = long_df['role'].map({'Owner': 1, 'Member': 2})
        long_df = long_df.sort_values('role_rank').drop_duplicates(subset=['user_email', 'site_name'], keep='first')
        
        pivot_df = long_df.pivot_table(
            index=['user_displayName', 'user_email', 'user_id'],
            columns='site_name',
            values='role',
            aggfunc='first',
            fill_value='Visitor'
        ).reset_index()

        out_path = r"C:\Python_Scripts\api_calls\reports\Site_Access_Matrix.xlsx"
        
        pivot_df.to_excel(out_path, index=False)
        print(f"\nMatrix saved: {out_path}")
        display(pivot_df.head())
    else:
        print("\nNo user data collected.")
