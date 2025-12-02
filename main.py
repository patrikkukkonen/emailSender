# This is a python script to automate email sending,
# when needing to send same email to multiple recipients separately,
# it saves a lot of time.
# Supports HTML email body with inline images and attachments.

import os.path
import json
import base64
from jinja2 import Environment, FileSystemLoader

from datetime import datetime

from email import encoders
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from html.parser import HTMLParser

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
# If modifying these scopes, delete the file token.json.

# Jinja2 template loader setup
env = Environment(loader=FileSystemLoader('.')) # current directory
template = env.get_template('messages/message.html') # Load the HTML template

# Email content if using SMTP instead of Gmail API
# subject = "Email Subject"
# body =  "Hei! \n Tämä on testisähköposti pythonilla lähetettynä.\n Terveisin, Lähettäjä"
# sender = "name.surname@gmail.com"
# recipients = ["reciever1@gmail.com", "reciever2@gmail.com", "reciever3@gmail.com"]  # List of recipients
# password = "password"  # or input("Type your password and press enter: ")

# Load sensitive configuration from a JSON file
with open('config.json', 'r') as config_file:
    config = json.load(config_file)

sender_name = config.get('sender_name', '') # Default to empty string if not found
sender_email = config.get('sender_email', '')
recipients_file = config.get('recipients_file', []) # List of recipient email addresses
attachment_paths = config.get('attachment_paths', '') # Paths to attachment files, if any.
inline_icons = config.get('inline_icons', []) # Paths and CIDs for inline images.


# Unique message ID based on timestamp
timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
final_html = template.render(timestamp=timestamp) # Render the template with timestamp variable


# Build email message with HTML body, attachments, and inline images
def build_message_with_inline_images(recipient, subject, html_body, attachment_paths, inline_icons):
    # Outer container for attachments + related content
    outer = MIMEMultipart('mixed')
    outer['To'] = recipient
    outer['Subject'] = subject
    outer['From'] = f'{sender_name} <{sender_email}>'  # Gmail API will send from authorized account when using userId="me"

    # Related part will contain the HTML body and inline images
    related = MIMEMultipart('related')

    # Alternative part for plain text + html
    alternative = MIMEMultipart('alternative')
    plain_fallback = ("Tämä viesti sisältää HTML-muotoisen sisällön. Mikäli näet tämän tekstin, sähköpostiohjelmasi ei tue HTML-renderöintiä.")
    alternative.attach(MIMEText(plain_fallback, 'plain'))
    alternative.attach(MIMEText(html_body, 'html'))

    # build structure: outer(mixed) -> related -> alternative(+ images)
    related.attach(alternative)
    outer.attach(related)

    # attach inline images to the related part
    for icon_path, cid in inline_icons:
        attach_image(related, Path(icon_path), cid)

    # attach regular attachments to the outer mixed part
    for ap in attachment_paths:
        outer.attach(load_attachments(ap))

    return outer

# Load recipients's email addresses from JSON file
def load_recipients(filename):
    with open(filename, 'r') as file:
        data = json.load(file)
    return data['recipients'] # expecting {"recipients": ["email1", "email2", ...]}

# HTML parser to extract title from HTML file's <title> tag
class TitleParser(HTMLParser):
    # Initialize the parser
    def __init__(self):
        super().__init__()
        self.in_title = False
        self.title = None

    # Handle start tag
    def handle_starttag(self, tag, attrs):
        if tag.lower() == "title":
            self.in_title = True
    # Handle end tag
    def handle_endtag(self, tag):
        if tag.lower() == "title":
            self.in_title = False
    # Handle data inside tags
    def handle_data(self, data):
        if self.in_title:
            self.title = data.strip()


# Get email subject (and body from HTML file)
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


# Attach images as inline MIMEImage with Content-ID
def attach_image(msg, image_path: Path, cid: str):
    data = image_path.read_bytes()
    suffix = image_path.suffix.lower()

    # Subtype mapping for common image formats
    subtype_map = {
        '.png': 'png',
        '.jpg': 'jpeg',
        '.jpeg': 'jpeg',
        '.gif': 'gif',
        '.bmp': 'bmp',
        '.webp': 'webp',
        '.tif': 'tiff',
        '.tiff': 'tiff',
    }

    # Use MIMEImage for formats imghdr recognizes
    if suffix in subtype_map:
        img = MIMEImage(data, _subtype=subtype_map[suffix])
        encoders.encode_base64(img)
        img.add_header("Content-ID", f"<{cid}>")
        img.add_header("Content-Disposition", "inline") # , filename=image_path.name)
        msg.attach(img)
        return

    # Handle SVG explicitly since MIMEImage does not support it (if needed)
    if suffix in {'.svg', '.svgz'}:
        part = MIMEBase('image', 'svg+xml')
        part.set_payload(data)
        encoders.encode_base64(part)
        part.add_header("Content-ID", f"<{cid}>")
        part.add_header("Content-Disposition", "inline") # , filename=image_path.name)
        msg.attach(part)
        return

    # Fallback to MIMEImage for other formats
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(data)
    encoders.encode_base64(part)
    part.add_header("Content-ID", f"<{cid}>")
    part.add_header("Content-Disposition", "inline") # filename=image_path.name)
    msg.attach(part)

# Load recipients and email content from files
# In the future could be:
#   Command line arguments (?)
recipients = load_recipients(recipients_file) # list of email addresses
subject, _ = load_subject_and_body('messages/message.html') # we only need the subject here

# Render the final HTML body using Jinja2 template
body = final_html # rendered HTML content

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
        # Save the credentials for the next run in token.json
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    # Build the Gmail API service
    service = build('gmail', 'v1', credentials=creds)

    # Loop through each recipient and send email
    # Future improvement ideas:
    #   * Use threading or async to send emails concurrently
    #   * Add delay between emails to avoid rate limiting
    #   * Log sent emails to a file or database
    #   * Handle bounces and delivery failures
    #   * Add CC and BCC support
    for recipient in recipients:
        # message = MIMEText(body, 'html')  # Create a MIMEText object with the email body
        # message = MIMEMultipart()  # Create a MIMEMultipart object to handle attachments

        # Build the email message with inline images and attachments
        message = build_message_with_inline_images(
            recipient, # to
            subject, # subject
            body, # html body
            attachment_paths=[attachment_paths], # attachments
            inline_icons=inline_icons # inline images with CIDs
        )

        # Encode the message as base64url
        create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

        try: # Send the email via Gmail API and print the message ID of the sent message
            sent = (service.users().messages().send(userId="me", body=create_message).execute())
            print(F'sent message to {recipient} Message Id: {sent.get("id", "unknown")}')

        # Future improvement: handle specific errors like invlaid email address etc.
        except HttpError as error: # Handle errors from Gmail API
            print(F'An error occurred: {error}')
            # sent = None

if __name__ == '__main__':
    main()