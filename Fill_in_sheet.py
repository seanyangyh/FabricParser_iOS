from __future__ import print_function
import httplib2
import os
import datetime
import User_Input
from time import sleep

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
# SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def PATH(p):
    return os.path.abspath(os.path.join(os.path.dirname(__file__), p))


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def allsheetappenddate(date, sheetid, range, service):
    value_range_body = {
        'values': [
            [date, "", "", "", ""],
        ]
    }
    result = service.spreadsheets().values().append(spreadsheetId=sheetid, range=range,
                                                     valueInputOption='RAW', body=value_range_body).execute()
    beginsplit = result['updates']['updatedRange'].index(':')
    split = result['updates']['updatedRange'][beginsplit+2:]
    print(split)
    split1 = int(split)
    result2 = allsheetfillcolor(split1, sheetid, service)
    return result2


def allsheethandler(id, ver, title, url, snippet, sheetid, range, service):
    value_range_body = {
        'values': [
            [id, ver, title, url, User_Input.Default_status, User_Input.Default_owner, "", "", snippet],
        ]
    }
    result = service.spreadsheets().values().append(spreadsheetId=sheetid, range=range,
                                                     valueInputOption='RAW', body=value_range_body).execute()
    return result


def allsheetfillcolor(row, sheetid, service):
    batch_update_spreadsheet_request_color = {
        "requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": User_Input.sheet_id_all,
                        "startRowIndex": row-1,
                        "endRowIndex": row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 9
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 1,
                                "green": 1,
                                "blue": 0
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor)"
                }
            },
            {
                "mergeCells": {
                    "range": {
                        "sheetId": User_Input.sheet_id_all,
                        "startRowIndex": row-1,
                        "endRowIndex": row,
                        "startColumnIndex": 0,
                        "endColumnIndex": 9
                    },
                    "mergeType": "MERGE_ROWS"
                }
            }
        ]
    }
    result = service.spreadsheets().batchUpdate(spreadsheetId=sheetid,
                                                 body=batch_update_spreadsheet_request_color).execute()
    return result


def main():
    today = datetime.datetime.now()
    print(today.strftime("%Y/%m/%d"))
    # today = today.strftime('%Y%m%d')

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheet_id = User_Input.spreadsheet_id
    sheet_id_template = User_Input.sheet_id_template

    # Copy a sheet from Template
    copy_sheet_to_another_spreadsheet_request_body = {
        "destination_spreadsheet_id": spreadsheet_id
    }
    result1 = service.spreadsheets().sheets().copyTo(spreadsheetId=spreadsheet_id, sheetId=sheet_id_template,
                                                     body=copy_sheet_to_another_spreadsheet_request_body).execute()
    print(result1)

    # Rename the copied sheet to today
    batch_update_spreadsheet_request_body = {
      "requests": {
        "updateSheetProperties": {
          "properties": {
            "sheetId": result1['sheetId'],
            "title": today.strftime('%Y%m%d'),
          },
          "fields": "title",
        }
      }
    }
    result2 = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                 body=batch_update_spreadsheet_request_body).execute()
    print(result2)

    # Define range and read the local result file
    rangename = today.strftime('%Y%m%d') + '!A2:I'
    rangenameall = 'All!A2:F'

    file = open(str(PATH('./result/' + str(today.strftime('%Y%m%d') + '.txt'))), "r")
    readfile = file.readlines()
    file.close()
    print(readfile)

    # Append today's date into 'All' sheet
    appenddate = allsheetappenddate(today.strftime("%Y/%m/%d"), spreadsheet_id, rangenameall, service)
    print(appenddate)

    # Append result to 'today' and 'All' sheets
    for i in range(0, len(readfile), 1):
        # sleep(1)
        resultsplit = readfile[i].strip().split(" , ")
        fid = resultsplit[0]
        fver = resultsplit[1]
        ftitle = resultsplit[2]
        furl = resultsplit[3]
        fsnippet = resultsplit[4]
        if fver == User_Input.Version:
            appendall = allsheethandler(fid, fver, ftitle, furl, '', spreadsheet_id, rangenameall, service)
            print(appendall)
        else:
            appendtoday = allsheethandler(fid, fver, ftitle, furl, fsnippet, spreadsheet_id, rangename, service)
            print(appendtoday)


if __name__ == '__main__':
    main()
