# This is a python script to automate email sending,
# especially when needing to send same email to multiple recipients separately,
# it saves a lot of time.

import os.path
import json
import base64

from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from html.parser import HTMLParser

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
# If modifying these scopes, delete the file token.json.


# Email content if using SMTP instead of Gmail API
# subject = "Email Subject"
# body =  # "Hei! \n Tämä on testisähköposti pythonilla lähetettynä.\n Terveisin, Lähettäjä"
# sender = "name.surname@gmail.com"
# recipients = ["reciever1@gmail.com", "reciever2@gmail.com", "reciever3@gmail.com"]  # List of recipients
# password = "password"  # or input("Type your password and press enter: ")


# Lataa sähköpostiosoitteet JSON-tiedostosta
def load_recipients(filename):
    with open(filename, 'r') as file:
        data = json.load(file)
    return data['recipients']


# HTML Parseri otsikon hakemiseen HTML-tiedostosta Title-tagista
class TitleParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_title = False
        self.title = None

    def handle_starttag(self, tag, attrs):
        if tag.lower() == "title":
            self.in_title = True

    def handle_endtag(self, tag):
        if tag.lower() == "title":
            self.in_title = False

    def handle_data(self, data):
        if self.in_title:
            self.title = data.strip()


# Lataa sähköpostin otsikon ja sisällön HTML-tiedostosta
def load_subject_and_body(filename):
    with open(filename, 'r', encoding='utf-8') as file:
        html = file.read()
    parser = TitleParser()
    parser.feed(html)
    otsikko = parser.title if parser.title else "Ei otsikkoa"
    return otsikko, html


# Add attachment to the email
def load_attachments(filename):
    with open(filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename= {os.path.basename(filename)}',
    )
    return part


recipients = load_recipients('recipients/recipients.json')
subject, body = load_subject_and_body('messages/message.html')


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    for recipient in recipients:

        # message = MIMEText(body, 'html')  # Create a MIMEText object with the email body

        message = MIMEMultipart()  # Create a MIMEMultipart object to handle attachments

        # message['From'] = "me"  # Set the sender email address
        message['Subject'] = subject  # Set the email subject
        message['To'] = recipient  # Set the current recipient email address
        message.attach(MIMEText(body, 'html'))  # Attach the email body as HTML
        # message.attach(load_attachments('attachments/attachment.pdf'))  # Attach the file

        create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

        try:
            sent = (service.users().messages().send(userId="me", body=create_message).execute())
            print(F'sent message to {recipient} {sent} Message Id: {message["id"]}')

        except HttpError as error:
            print(F'An error occurred: {error}')
            # sent = None

if __name__ == '__main__':
    main()