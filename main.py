from __future__ import print_function

import os.path
import xlsxwriter 
import requests

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SAMPLE_SPREADSHEET_ID = '1aBy0Oe8f0pfMYh29f_k8SoNJWNulECsQ490OixuDdd0'
SAMPLE_RANGE_NAME = 'Página1'

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])
        
        
        workbook = xlsxwriter.Workbook('result.xlsx') 
        worksheet = workbook.add_worksheet() 
        
        data = requests.get('https://s7ozoqcftf.execute-api.ap-south-1.amazonaws.com/dev/patients').json()
        for value in data.items():
            print('')
        
        # Criar tabela
        
        worksheet.write('A1', 'ID') 
        worksheet.write('A2', '%s' % (value[1][0]['patientId'])) 
        worksheet.write('B1', '%s' % (values[0][0])) 
        worksheet.write('C1', '%s' % (values[0][1])) 
        worksheet.write('D1', 'Status') 
        worksheet.write('E1', 'Ativo') 
        worksheet.write('A3', '%s' % (value[1][1]['patientId'])) 
        worksheet.write('B2', '%s' % (values[1][0])) 
        worksheet.write('C2', '%s' % (values[1][1])) 
        worksheet.write('D2', '%s' % (value[1][0]['room-status'])) 
        worksheet.write('E2', '%s' % (value[1][0]['rolling'])) 
        worksheet.write('A4', '%s' % (value[1][2]['patientId'])) 
        worksheet.write('B3', '%s' % (values[2][0])) 
        worksheet.write('C3', '%s' % (values[2][1])) 
        worksheet.write('D3', '%s' % (value[1][1]['room-status'])) 
        worksheet.write('E3', '%s' % (value[1][1]['rolling'])) 
        # worksheet.write('A5', '%s' % (value[1][3]['patientId'])) 
        worksheet.write('B4', '%s' % (values[3][0])) 
        worksheet.write('C4', '%s' % (values[3][1]))
        worksheet.write('D4', '%s' % (value[1][2]['room-status']))
        worksheet.write('E4', '%s' % (value[1][2]['rolling']))
        workbook.close()

        if not values:
            print('No data found.')
            return

        for row in values:
            print('%s, %s' % (row[0], row[1]))

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()