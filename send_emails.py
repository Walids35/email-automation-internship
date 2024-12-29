import os
import base64
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Define the scope for Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def authenticate_gmail_api():
    """Authenticate with Gmail API and return the service object."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def create_email_with_attachment(sender, to, subject, message_text, file_path):
    """Create an email with an attachment and custom sender name."""
    # Create the email
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = f"Walid Siala <{sender}>"
    message['subject'] = subject

    # Attach the message text
    message.attach(MIMEText(message_text, 'html'))

    # Attach the file
    if file_path:
        with open(file_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename={os.path.basename(file_path)}'
        )
        message.attach(part)

    # Encode the message in base64
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    return {'raw': raw}

def send_email(service, sender, recipient, subject, message_text, file_path=None):
    """Send an email with an optional attachment using Gmail API."""
    message = create_email_with_attachment(sender, recipient, subject, message_text, file_path)
    service.users().messages().send(userId="me", body=message).execute()

def main():
    service = authenticate_gmail_api()

    sender_email = "wsiala4@gmail.com"

    contacts_df = pd.read_csv('contacts.csv')
    with open('template.txt', 'r') as template_file:
        email_template = template_file.read()

    attachment_path = 'Walid_Resume.pdf'

    for index, row in contacts_df.iterrows():
        recipient_email = row['email']
        recipient_name = row['name']
        recipient_company = row['company_name']
        subject = f"Internship Opportunity - {recipient_name}"

        personalized_message = email_template.replace("{name}", recipient_name)
        personalized_message = personalized_message.replace("{company_name}", recipient_company)

        send_email(service, sender_email, recipient_email, subject, personalized_message, attachment_path)
        print(f"Email sent to {recipient_name} ({recipient_email}) with attachment {attachment_path}")

if __name__ == "__main__":
    main()
