from __future__ import print_function
import httplib2
import os
import base64
import json
import time
import datetime
import re
import User_Input

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
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'


def PATH(p):
    return os.path.abspath(os.path.join(os.path.dirname(__file__), p))


def checkFolder():
    if not os.path.exists(os.getcwd() + "/result"):
        os.makedirs(os.getcwd() + "/result")


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
                                   'gmail-python-quickstart.json')

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


def MailContentParser(rawdata):
    # Parsing raw data, replace format to json
    raw = str(rawdata).replace('\"', '\\\'')
    raw_to_json = str(raw).replace('\'', '\"')
    # json -> dict of python, and decode base64 to get url
    json_data = json.loads(raw_to_json)
    # Query subject
    subject = json_data['payload']['headers'][14]['value']

    try:
        # Query subject to get version and issue type
        begin_issuetype = subject.index("[")
        last_issuetype = subject.index("-")
        subjecttitle = subject[begin_issuetype:last_issuetype]
        subjectver = subject.strip().split()[-1]
    except:
        # Stability Digest or non-recognize subject
        subjecttitle = subject
        subjectver = ''

    try:
        # Get fabric num
        begin_fabric_num = json_data['snippet'].index("#")
        fabric_num = json_data['snippet'][begin_fabric_num+1:begin_fabric_num+6]
    except:
        # Cannot find # in snippet because string is too long
        fabric_num = 'unknown'

    try:
        # try to decode base64 to get url
        decode_result = base64.b64decode(json_data['payload']['parts'][0]['body']['data'])
        decode_result = str(decode_result, 'utf-8')
        urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',
                          decode_result)
        # try to get sub-title
        decode_result_split = decode_result.split('\r\n')
        print(decode_result_split)
        if decode_result_split[1] == '':
            subjecttitle = subjecttitle + decode_result_split[3]
        else:
            subjecttitle = subjecttitle + decode_result_split[2]
            begin_fabric_num = decode_result_split[0].index("#")
            fabric_num = decode_result_split[0][begin_fabric_num + 1:begin_fabric_num + 6]
        # url add query last 7 days
        urllast7days = urls[0] + "?time=last-seven-days"
        print(fabric_num + ' , ' + subjectver + ' , ' + subjecttitle + ' , ' + urllast7days + ' , ' + json_data['snippet'])
        result = dict(num=fabric_num, ver=subjectver, title=subjecttitle, url=urllast7days, detail=json_data['snippet'])
    except:
        # decode base64 fail case
        urlsError = 'url decode failed!'
        print(fabric_num + ' , ' + subjectver + ' , ' + subjecttitle + ' , ' + urlsError + ' , ' + json_data['snippet'])
        result = dict(num=fabric_num, ver=subjectver, title=subjecttitle, url=urlsError, detail=json_data['snippet'])

    return result


def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """
    today = datetime.datetime.now()
    # yesterday = today - datetime.timedelta(1)
    # yesterday = yesterday.strftime("%Y/%m/%d")
    # print('Yesterday: ' + yesterday)
    print('Today: ' + today.strftime("%Y/%m/%d"))
    today = today.strftime("%Y%m%d")
    file = open(str(PATH('./result/' + str(today + '.txt'))), "wb")
    file.close()

    # read timestamp last time
    readtime = open(str(PATH('./timestamp_lasttime.txt')), "r")
    last_timestamp = readtime.read()
    readtime.close()

    # Write concurrent timestamp to timestamp_lasttime.txt
    writefile = open(str(PATH('./timestamp_lasttime.txt')), "w")
    concurrent_timestamp = int(time.time())
    writefile.write(str(concurrent_timestamp))
    writefile.close()

    # Credential
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    # query and get mail id after last timestamp (ex. after:1492048800)
    getMessageList = service.users().messages().list(userId='me', labelIds=User_Input.Gmail_fabric_label, q='after:' + last_timestamp, maxResults=1000).execute()
    MailCount = getMessageList['resultSizeEstimate']
    getMessage = getMessageList['messages']
    getMessageJson = str(getMessage).replace('\'', '\"')
    getMessageJson_data = json.loads(getMessageJson)
    print('Mail Count: ' + str(MailCount))
    print(getMessageJson_data)

    for i in range(0, MailCount, 1):
        # get mail content from message id, and check/modify to correct json format
        getRawContent = service.users().messages().get(userId='me', id=getMessageJson_data[i]['id']).execute()
        fabric_data = MailContentParser(getRawContent)
        # write result separated by , to local .txt
        file = open(str(PATH('./result/' + str(today + '.txt'))), "a")
        file.write(fabric_data['num'] + " , " + fabric_data['ver'] + " , " + fabric_data['title'] + " , " + fabric_data['url'] + " , " + fabric_data['detail'] + "\n")
        file.close()


if __name__ == '__main__':
    checkFolder()
    main()
