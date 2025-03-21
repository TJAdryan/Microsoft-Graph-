
#Get Token


import requests
from msal import ConfidentialClientApplication
import os
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from dotenv import load_dotenv,find_dotenv

#To point to specific secret directory
load_dotenv(find_dotenv("C:/Python_Scripts/N-Able/.env"))
web_token=os.getenv('webtoken')
sql_user =os.getenv('sqluser')
sql_pass=os.getenv('sqlpass')
DSN=os.getenv('dsn')

#update with your MSGraph Tokens
client_id=os.getenv('client_id')
tenant_id =os.getenv('tenant_id')
client_secret =os.getenv('client_secret')






msal_authority = f"https://login.microsoftonline.com/{tenant_id}"

msal_scope = ["https://graph.microsoft.com/.default"]




msal_app = ConfidentialClientApplication(
    client_id=client_id,
    client_credential=client_secret,
    authority=msal_authority,
)

result = msal_app.acquire_token_silent(
    scopes=msal_scope,
    account=None,
)

if not result:
    result = msal_app.acquire_token_for_client(scopes=msal_scope)

if "access_token" in result:
    access_token = result["access_token"]
    print('token acquired')
else:
    raise Exception("No Access Token found")

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}


# This url can be used to query recent sign ins to compare to device list
# response = requests.get(url="https://graph.microsoft.com//beta/users?$select=userprincipalname,signInActivity",
#     headers=headers)




##get licenses used by company
endpoint = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices'

#To get only Windows devices if needed
#endpoint = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$filter=operatingSystem eq \'Windows\''
http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                'Accept': 'application/json',
                'Content-Type': 'application/json'}

frames = []
data = requests.get(endpoint, headers=http_headers, stream=False).json()
eps = pd.json_normalize(data['value'])
print(eps.shape)
eps_cleaned= eps.dropna(axis=1, how='all')


# Use to remove any test users from the list of managed devices
#filtered = eps_cleaned[~eps_cleaned['userPrincipalName'].isin(['test@testcompanies.com'])]

filtered.replace('', pd.NA, inplace=True)
filtered = filtered.dropna(subset=['userPrincipalName'])
# Drop columns with all NaN values
filtered= filtered.dropna(axis=1, how='all')
filtered['last_sync_date'] = pd.to_datetime(filtered['lastSyncDateTime'],errors='coerce').dt.date
filtered = filtered[['id', 'userId','last_sync_date', 'deviceName', 'managedDeviceOwnerType',
       'enrolledDateTime', 'lastSyncDateTime', 'operatingSystem','complianceState',  'osVersion',
       'easActivated', 'easDeviceId', 'easActivationDateTime',
       'azureADRegistered', 'deviceEnrollmentType', 'emailAddress','azureADDeviceId', 
       'deviceCategoryDisplayName', 'isSupervised','isEncrypted', 'userPrincipalName','model', 'manufacturer', 
       'complianceGracePeriodExpirationDateTime', 'serialNumber','userDisplayName','wiFiMacAddress', 
       'totalStorageSpaceInBytes', 'freeStorageSpaceInBytes',
       'managedDeviceName', 'partnerReportedThreatState',
        'managementCertificateExpirationDate',
       'physicalMemoryInBytes', 'deviceActionResults']]

print(filtered.shape)
filtered['osVersion'].value_counts()

#Windows 10.0.22 or higher is Windows 11, below 10.
#Mac so many versions...
def identify_os_version(os_version):
    if os_version.startswith("10.0.22") or os_version.startswith("10.0.23") or os_version.startswith("10.0.24")or os_version.startswith("10.0.25")or os_version.startswith("10.0.26")or os_version.startswith("10.0.27"):
        return "Windows 11"
    elif os_version.startswith("10.0"):
        return "Windows 10"
    elif os_version.startswith("13"):
        return f"macOS Ventura {os_version}"
    elif os_version.startswith("14"):
        return f"macOS Big Sur {os_version}"
    elif os_version.startswith("15"):
        return f"macOS Monterey {os_version}"
    else:
        return f"Unknown {os_version}"


# Apply the identification function to the osVersion column
filtered['osVersionName'] = filtered['osVersion'].apply(identify_os_version)
os_versions = filtered[['id', 'userId','last_sync_date', 'deviceName', 'managedDeviceOwnerType',
  'complianceState',  'osVersionName','osVersion','operatingSystem',
   'azureADRegistered', 'deviceEnrollmentType', 'emailAddress','azureADDeviceId', 
   'deviceCategoryDisplayName', 'isSupervised','isEncrypted', 'userPrincipalName','model', 'manufacturer', 
 'serialNumber','userDisplayName','wiFiMacAddress']]
#os_versions = os_versions[~os_versions['model'].str.contains('Microsoft Dev Box')]


connection_url = URL.create(
    "mssql+pyodbc",
    username="SQLUser",
    password=sql_pass,
    host="127.0.0.1",
    port=1450,
    database="calls",
    query={
        "driver": "ODBC Driver 17 for SQL Server",
        "Encrypt": "yes",
        "TrustServerCertificate": "yes",
    },
)
engine = create_engine(connection_url)


#Update SQL Table
os_versions.to_sql('intune_devices', engine, if_exists='replace', index=False)
myQuery3 = '''
            SELECT *
            FROM intune_devices

'''



cid= pd.read_sql(myQuery3,engine)
cid=cid[[ 'last_sync_date', 'deviceName',
       'managedDeviceOwnerType', 'complianceState', 'osVersionName',
       'osVersion','deviceEnrollmentType', 'emailAddress', 
       'deviceCategoryDisplayName', 'isEncrypted',
       'userPrincipalName', 'model', 'manufacturer', 'serialNumber',
       'userDisplayName', 'wiFiMacAddress']]
print(cid.shape)
cintune = cid.rename(columns={'deviceName':'discoveredname'})
#Get Device List from RMM
#Replace CompanyName% RMM_table
myQuery = '''SELECT deviceid,devicename,deviceclass,license,supportedos,deviceclasslabel, discoveredname,lastloggedinuser,stillloggedin,sitename
          FROM RMM_devices
          WHERE customername LIKE 'CompanyName%'  
          '''
device =pd.read_sql(myQuery,engine)
print(device.shape)


merged_df = pd.merge(device, cintune, on='discoveredname', how='inner')

#change specific company_name

merged_df.to_sql('company_name_intune_join', engine, if_exists='replace')

print('done')
