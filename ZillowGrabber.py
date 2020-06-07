import numpy as np
import pandas as pd
import zillow
import configparser
from oauth2client.service_account import ServiceAccountCredentials
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import datetime

# Open our config file
config = configparser.ConfigParser()
config.read('zillowkey.conf')

# Grab our key
key = config['DEFAULT']['zillowkey']

# Setup zillow api
api = zillow.ValuationApi()

# Load in our data and cut out irrelevant information
addresses = pd.read_excel('HousingData.xls')
address_df = addresses.loc[addresses['County'].isin(['Orange', 'Los Angeles', 'San Bernardino', 'Riverside'])]
address_df = address_df.loc[address_df['Status'].isin(['Postponed', 'Auction'])]
address_df = address_df.drop(['Original Principal Balance'], axis=1).dropna(subset=['Property Address', 'Zip'])

df = address_df.copy()

# Setup lists to be added to our dataframe from zillow api
zestimate = []
zillow_year_build = []
zillow_lot_size_sqft = []
zillow_bathrooms = []
zillow_bedrooms = []

# Go through each address, grab info, add it to our lists
for index, row in address_df.iterrows():
    address = address_df['Property Address'][index].rstrip('\n')
    postal_code = str(int(address_df['Zip'][index]))
    try:
        data = api.GetDeepSearchResults(key, address, postal_code)
        datadict = data.get_dict()
        zestimate.append(datadict['zestimate']['amount'])
        zillow_year_build.append(datadict['extended_data']['year_built'])
        zillow_lot_size_sqft.append(datadict['extended_data']['lot_size_sqft'])
        zillow_bathrooms.append(datadict['extended_data']['bathrooms'])
        zillow_bedrooms.append(datadict['extended_data']['bedrooms'])
    except:
        zestimate.append(None)
        zillow_year_build.append(None)
        zillow_lot_size_sqft.append(None)
        zillow_bathrooms.append(None)
        zillow_bedrooms.append(None)

# Create new columns in our dataframe from our lists
df['zestimate'] = zestimate
df['zillow_year_build'] = zillow_year_build
df['zillow_lot_size_sqft'] = zillow_lot_size_sqft
df['zillow_bathrooms'] = zillow_bathrooms
df['zillow_bedrooms'] = zillow_bedrooms


df.drop(columns=['State'], inplace=True)
df = df.fillna('None')

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
sheet_name = "TEST"
range_ = "Sheet1!A:A"

# Name for sheet
d = datetime.datetime.today()
sheet_name = 'Housing ' + d.strftime('%d-%m-%Y')

creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

# Create new Sheet
spreadsheet = {
    'properties': {
        'title': sheet_name
    }
}
spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                    fields='spreadsheetId').execute()
spreadsheet_id = spreadsheet.get('spreadsheetId')
print('Spreadsheet ID: {0}'.format(spreadsheet.get('spreadsheetId')))

# Fill in header
list = [df.columns.tolist()]

resource = {
  "majorDimension": "ROWS",
  "values": list
}

request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_, body=resource,
  valueInputOption="USER_ENTERED")
response = request.execute()

# Fill in body
list = df.values.tolist()

resource = {
  "majorDimension": "ROWS",
  "values": list
}

request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_, body=resource,
  valueInputOption="USER_ENTERED")
response = request.execute()
