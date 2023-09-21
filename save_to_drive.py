import subprocess
import time
import io
import base64
import re
from googleapiclient.http import MediaIoBaseUpload
from email_credentials import email_address, email_password
from google_api import create_service

def construct_service(api_service):
    CLIENT_SERVICE_FILE = 'client_secret2.json'
    try:
        if api_service == 'drive':
            API_NAME = 'drive'
            API_VERSION = 'v3'
            SCOPES = ['https://www.googleapis.com/auth/drive']
            return create_service(CLIENT_SERVICE_FILE, API_NAME, API_VERSION, SCOPES)
        elif api_service == 'gmail':
            API_NAME = 'gmail'
            API_VERSION = 'v1'
            SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify']
            return create_service(CLIENT_SERVICE_FILE, API_NAME, API_VERSION, SCOPES)

    except Exception as e:
        print(e)
        return None
    

def search_email(service, query_string, label_ids=[]):
    try:
        message_list_response = service.users().messages().list(
            userId='me',
            labelIds=label_ids,
            q=query_string
        ).execute()
        message_items = message_list_response.get('messages')
        nextPageToken = message_list_response.get('nextPageToken')

        while nextPageToken:
            message_list_response = service.users().messages().list(
                userId='me',
                labelIds=label_ids,
                q=query_string,
                pageToken=nextPageToken
            ).execute()

            message_items.extend(message_list_response.get('messages'))
            nextPageToken = message_list_response.get('nextPageToken')
        return message_items

    except Exception as e:
        return None

def get_message_detail(service, message_id, format='metadata', metadata_headers=[]):
    try:
        message_detail = service.users().messages().get(
            userId='me',
            id=message_id,
            format=format,
            metadataHeaders=metadata_headers
        ).execute()
        return message_detail

    except Exception as e:
        print(e)
        return None
    
def mark_email_as_read(gmail_service, message_id):
    # Use the modify method to mark the email as read
    gmail_service.users().messages().modify(userId='me', id=message_id, body={'removeLabelIds': ['UNREAD']}).execute()   

def extract_name_from_email_body(email_body):
    # Use regular expressions to find the name between the first space and the first comma
    match = re.search(r'\s(.*?)\,', email_body)
    if match:
        return match.group(1)
    return None

# Step 1: create a google service instance
gmail_service = construct_service('gmail')
time.sleep(2)
drive_service = construct_service('drive')

# Step 2: Search email from safaricom that are unread
query_string = 'is:unread has:attachment from:m-pesastatements@safaricom.co.ke'
email_messages = search_email(gmail_service, query_string, ['INBOX'])

# Step 3: Download email and save them to google drive
if email_messages:
    for email_message in email_messages:
        message_id = email_message['threadId']
        messageDetail = get_message_detail(
            gmail_service,
            email_message['id'],
            format='full',
            metadata_headers=['parts'])
        messageDetailPayload = messageDetail.get('payload')

        for item in messageDetailPayload['headers']:
            if item['name'] == 'Subject':
                if item['value']:
                    message_subject = '{0} ({1})'.format(item['value'], message_id)
                else:
                    message_subject = '(No Subject) ({0})'.format(message_id)

        # Step 4: Select a drive folder
        folder_id = '1vACT2I0Xwr3kJoY55fz_3zHSPjsHjyvm'

        # Step 5: Inside the loop that processes email messages
        if 'parts' in messageDetailPayload:
            for msgPayload in messageDetailPayload['parts']:
                mime_type = msgPayload['mimeType']
                file_name = msgPayload['filename']
                body = msgPayload['body']

                if 'attachmentId' in body and mime_type == 'application/pdf':
                    attachment_id = body['attachmentId']

                    response = gmail_service.users().messages().attachments().get(
                        userId='me',
                        messageId=email_message['id'],
                        id=attachment_id
                    ).execute()

                    file_data = base64.urlsafe_b64decode(
                        response.get('data').encode('UTF-8'))

                    fh = io.BytesIO(file_data)

                    file_metadata = {
                        'name': file_name,
                        'parents': [folder_id],
                        'mimeType': mime_type
                    }

                    print(file_name)

                    media_body = MediaIoBaseUpload(fh, mimetype=mime_type, chunksize=1024*1024, resumable=True)

                    file = drive_service.files().create(
                        body=file_metadata,
                        media_body=media_body,
                        fields='id'
                    ).execute()

                    subprocess.run(["python", "open_pdf.py"])
                    # Step 6: Mark the email as read and break the loop.
                    mark_email_as_read(gmail_service, email_message['id'])
                        
else:
    subprocess.run(["python", "open_pdf.py"])
    print("No new emails found. Refreshing after 5 minutes.")