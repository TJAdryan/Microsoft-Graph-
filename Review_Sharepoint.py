import json
import requests
from msal import ConfidentialClientApplication
import pydantic
import pandas as pd
import io
from time import sleep

##updated 6/11/2024

client_id = 'update_this'
tenant_id = 'update_this'
client_secret ='update_this_too'

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

import time
time = time.strftime('%X %x %Z')
#Tokens expire after 1 hour or 3600 seconds

print(time)



#Get a list of Sharepoint sites the users have accessed in the past 90 days.  

url ="https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D90')"
http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                    'Accept': 'application/json',
                    'Content-Type': 'application/json'}

data = requests.get(url, headers=http_headers, stream=True)
df = pd.read_csv(io.StringIO(data.text))
df['Last Activity Date'] = pd.to_datetime(df['Last Activity Date'])
rec = df[df['Last Activity Date']> '2025-01-20']
rec['Storage Used GB'] =rec['Storage Used (Byte)']/1024/1024/1024
rec = rec[['Report Refresh Date', 'Site Id', 'Owner Display Name',
       'Is Deleted', 'Last Activity Date', 'File Count', 'Active File Count',
       'Page View Count', 'Visited Page Count','Storage Used GB', 
       'Storage Allocated (Byte)', 'Root Web Template', 'Owner Principal Name']]
	   
	   
	   
	   
sites = rec['Site Id'].tolist()
len(sites)

frames = []
number=0
for site in sites:
    url = f"https://graph.microsoft.com/v1.0/sites/{site}/drives"
    http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                'Accept': 'application/json',
                'Content-Type': 'application/json'}
    data = requests.get(url, headers=http_headers, stream=False).json()
    norm = data['value']
    df = pd.json_normalize(data['value'])
    frames.append(df)
    number+=1
    if number % 25 == 0:
        print(f"{number} completed")
    else:
        pass

    
    
print('done')

drive_list = pd.concat(frames)
dl = drive_list['id'].tolist()
len(dl)
#get a list of all the drives in each SharePoint site

frames = []
number=0
for site in sites:
    url = f"https://graph.microsoft.com/v1.0/sites/{site}/drives"
    http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                'Accept': 'application/json',
                'Content-Type': 'application/json'}
    data = requests.get(url, headers=http_headers, stream=False).json()
    norm = data['value']
    df = pd.json_normalize(data['value'])
    frames.append(df)
    number+=1
    if number % 25 == 0:
        print(f"{number} completed")
    else:
        pass

    
    
print('done')

drive_list = pd.concat(frames)
dl = drive_list['id'].tolist()
len(dl)

# Using list of Drives get a list of child items in each drive and when they were created/last modified
# I found many drives with 0 items which is to be expected and found many drives with 1000's of items


frame=[]
number = 0
while number < max_url:
    try:
        url = f"https://graph.microsoft.com/v1.0/drives/{dl[number]}/root/children"
        
        http_headers = {'Authorization': 'Bearer ' + result['access_token'],
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'}

        data = requests.get(url, headers=http_headers, stream=False).json()
        number+=1
        if number % 10 ==0:
            print(number)
        while url:
            response = requests.get(url, headers=http_headers, stream=False)
            sleep(.1)
            data = response.json()
           
            

            # Normalize and append the data
            df = pd.json_normalize(data['value'])
            frame.append(df)
            sleep(.1)

            # Check for the next page
       
           
            url = data.get('@odata.nextLink')
            if url:
                print(f"Next page URL: {url}")
        
          
    except IndexError as e:
        print(f"An error occurred: {e}")
        break

    except Exception as e:
        print(f"An error occurred: {e}")
    

# Concatenate all dataframes
changes = pd.concat(frame, ignore_index=True)
print('done')

#Now you have a dataframe with the root children for each folder in each drive along with thelast modified date

#updating that column to be a date type in pandas and performing a sort by last modified

timestamp_columns =['lastModifiedDateTime']
for col in timestamp_columns:
    changes[col] = pd.to_datetime(changes[col]).dt.date

date_to_compare = pd.to_datetime("2015-01-01").date()
filtered_df = changes[changes['lastModifiedDateTime'] < date_to_compare]
print(filtered_df.shape)
filtered_df['lastModifiedBy.user.email'].value_counts()




HilariousHippo@*****.com   50
WhimsicalWhale@*****.com   45
OddOtter@*****.com         30
LoopyLlama@*****.com       28
EccentricElephant@*****.com 17
SillyGoose@*****.com        9

# you can identify sites that have not been modified in over 10 years, which can be a good indicator that the files/sites are good candidates for archiving


