#Get the total available licenses for the tenant and how many are being utilized- check each user for what licenses are assigned to them

import requests
from msal import ConfidentialClientApplication
import os
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from dotenv import load_dotenv,find_dotenv


load_dotenv(find_dotenv("C:/Python_Scripts/N-Able/.env"))
web_token=os.getenv('webtoken')
sql_user =os.getenv('sqluser')
sql_pass=os.getenv('sqlpass')
DSN=os.getenv('dsn')
client_id=os.getenv('comp_client_id')
tenant_id =os.getenv('comp_tenant_id')
client_secret =os.getenv('comp_client_secret')






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

# response = requests.get(url="https://graph.microsoft.com//beta/users?$select=userprincipalname,signInActivity",
#     headers=headers)




##get licenses used by company

endpoint = "https://graph.microsoft.com/v1.0/subscribedSkus"
http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                'Accept': 'application/json',
                'Content-Type': 'application/json'}
frames = []
data = requests.get(endpoint, headers=http_headers, stream=False).json()
df = pd.json_normalize(data['value'])


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

df2= df[['accountName', 'appliesTo', 'capabilityStatus', 'skuId', 'skuPartNumber',  'consumedUnits',  'prepaidUnits.enabled', 'prepaidUnits.suspended',
       'prepaidUnits.warning', 'prepaidUnits.lockedOut']]
df2.rename(columns={'skuId':'id'},inplace=True)
print(df2.shape)
dfnames= pd.read_csv('c:/users/public/Product names and service plan identifiers for licensing(2).csv')

dfnames.drop_duplicates(subset=['GUID'],inplace=True)

dfnames= dfnames[['GUID','Product_Display_Name', 'String_Id']]
dfnames.rename(columns={'GUID':'id'},inplace=True)
dfmerge =pd.merge(df2,dfnames,on='id',how='left')
dfmerge.drop_duplicates(subset=['Product_Display_Name'],inplace=True)
dfmerge.columns

comp_365 = dfmerge[['accountName', 'appliesTo','Product_Display_Name','id',  'prepaidUnits.enabled', 'consumedUnits' ]]


comp_365.columns = comp_365.columns.str.replace('.', '_')

comp_365.columns = comp_365.columns.str.lower().str.replace(' ', '')



comp_365['uid'] = range(1, len(comp_365) + 1)

comp_365.to_sql('comp_365_licenses', engine, if_exists='replace', index=False)


#Get users with licenses assigned



frames=[]
#ep = "https://graph.microsoft.com/v1.0/users?$filter=accountEnabled eq true&$select=id,userPrincipalName,assignedLicenses"
ep = "https://graph.microsoft.com/v1.0/users?$filter=accountEnabled eq true"
http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                'Accept': 'application/json',
                'Content-Type': 'application/json'}

data = requests.get(ep, headers=http_headers, stream=False).json()
all_users = pd.json_normalize(data['value'])
frames.append(all_users)
        
nextlink = data['@odata.nextLink']

while True:
    try:
        data2 = requests.get(nextlink, headers=http_headers, stream=False).json()
        nextlink = data2.get('@odata.nextLink')
        print(len(nextlink))
        au = pd.json_normalize(data2['value'])
        frames.append(au)
        
        if not nextlink:
            print('all users found')
            break
   
    except TypeError:
        print('all users found')
        break
    except TypeError:
        print('Type')
        break
        
df= pd.concat(frames)
df.reset_index(inplace=True,drop=True)
# all_lic = df[df['assignedLicenses'].apply(lambda x: len(x)>0)]
# user_id = all_lic['id'].to_list()
# principal = all_lic['userPrincipalName']
# print(len(user_id))
        
# all_lic = all_users[all_users['assignedLicenses'].apply(lambda x: len(x)>0)]

df = df[['userPrincipalName','id']]
userids = df['id'].tolist()
upns = df['userPrincipalName'].tolist()
maxid = len(userids)
number =0
lic_data=[]
while number < maxid:
    try:
        ep = f"https://graph.microsoft.com/v1.0/users/{userids[number]}/licenseDetails"
        http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                'Accept': 'application/json',
                'Content-Type': 'application/json'}

        data = requests.get(ep, headers=http_headers, stream=False).json()
        lic = pd.json_normalize(data['value'])
        lic['upn']=upns[number]
        lic['id']=userids[number]
        lic = lic[['id','skuPartNumber','upn']]
        lic_data.append(lic)
        number+=1
        if number % 100 == 0:
            print(number)
    except Exception:
        number +=1
        if number % 100 == 0:
            print(number)
        
lf = pd.concat(lic_data)
lf['skuPartNumber'].value_counts()
lf = pd.concat(lic_data)        
dummies = pd.get_dummies(lf['skuPartNumber'])
lf = lf.rename(columns={'skuPartNumber':'license'})
# Concatenate the original DataFrame with the dummy columns
lf = pd.concat([lf, dummies], axis=1)
lf = lf.groupby('upn').max().reset_index()
del lf['license']
