import re
import os.path
import base64
from bs4 import BeautifulSoup

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from email.mime.text import MIMEText

def get_gmail_credentials():
    scopes = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.compose', 'https://www.googleapis.com/auth/gmail.modify']

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'gmail-credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def check_new_mail(service):
    results = service.users().messages().list(userId='me', labelIds=['UNREAD']).execute()
    return results

def get_mail(service, id):
    message = service.users().messages().get(userId='me', id=id, format="raw").execute()
    return message

def get_sender(raw):
    s = base64.urlsafe_b64decode(raw['raw']).decode('utf-8')
    sender = re.search(r'\nfrom:\s+[\s\S]*?\n', s, re.I).group(0)
    sender = re.search(r'\<(.*?)\>', sender)
    if (sender == None):
        return
    sender = sender.group(0)
    return sender[1:len(sender)-1].lower()

def construct_email(raw, email, name):
    s = base64.urlsafe_b64decode(raw['raw']).decode('utf-8')
    subject = re.search(r'\nsubject:\s+[\s\S]*?\n', s, re.I).group(0)
    s = s[s.find('\nContent-Type'):].strip()
    h = f'''
MIME-Version: 1.0
to: {email}
from: CISxIdeas Hackathon <hackathon@cis.edu.hk>{subject}
    '''
    s = s[:s.find('\n')] + h + s[s.find('\n'):]
    s = s.replace('{{name}}', name)
    body = {'raw': base64.b64encode(s.encode('utf-8')).decode('ascii')}
    return body

def get_drive_credentials():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

    creds = ServiceAccountCredentials.from_json_keyfile_name('sheets-credientials.json', scopes)
    return creds

def get_authorized_emails(client):
    sheet = client.open_by_key("1jf1Mh1X_297nkeFnkXFyhwxB_n1RHSkNkk05_XKeKFk").sheet1
    raw = sheet.col_values(2)
    return [e.strip() for e in raw if '@' in e]

def get_mass_emails(client):
    sheet = client.open_by_key("1ukTvbwpmfQcrXvwmQqEZxTLvECNUpo05KnIq8cdI7WQ").sheet1
    rawEmails = sheet.col_values(2)
    rawNames = sheet.col_values(3)
    return [(e.strip(), n.strip()) for (e, n) in zip(rawEmails, rawNames) if '@' in e and len(n.strip()) > 0]

def main():
    # get gmail and sheets credientials
    gmailCreds = get_gmail_credentials()
    driveCreds = get_drive_credentials()

    # get data from google sheets
    client = gspread.authorize(driveCreds)
    authorizedSenders = get_authorized_emails(client)
    massRecievers = get_mass_emails(client)

    # get unread mail
    service = build('gmail', 'v1', credentials=gmailCreds)
    unreadMail = check_new_mail(service)
    # parse unread mail individually
    for i in range(unreadMail['resultSizeEstimate']):
        id = unreadMail['messages'][i]['id']
        raw = get_mail(service, id)
        if not get_sender(raw) in authorizedSenders:
            continue
        service.users().messages().modify(userId='me', id=id, body={'removeLabelIds': ['UNREAD']}).execute()
        for reciever, name in massRecievers:
            email = construct_email(raw, reciever, name)
            service.users().messages().send(userId='me', body=email).execute()

if __name__ == '__main__':
    main()
