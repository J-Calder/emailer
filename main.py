import os
import base64
import pickle
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import spacy
from email.mime.text import MIMEText
from difflib import SequenceMatcher

# Load spaCy's English model
nlp = spacy.load("en_core_web_sm")

# Define template responses
templates = [
    {
        'keywords': ['order', 'not', 'updated'],
        'response': "We apologize for the delay in updating your order. Our team will look into it and update you as soon as possible. Please provide your order number.",
    },
    {
        'keywords': ['where', 'tracking', 'number'],
        'response': "We apologize for the delay in providing your tracking number. We'll send it to you as soon as it's available. Please provide your order number.",
    },
    {
        'keywords': ['etransfer', 'accepted'],
        'response': "We're currently processing your e-transfer. You'll receive a confirmation email once it's been accepted. If you don't receive the confirmation within 24 hours, please contact us with your order number.",
    },
    {
        'keywords': ['order', 'shipped'],
        'response': "We typically ship orders within 1-2 business days after payment has been accepted. You'll receive a tracking number once your order has been shipped. If there's any delay, please provide your order number so we can check the status.",
    },
    {
        'keywords': ['accepted', 'e-transfer', 'order', 'on-hold'],
        'response': "We apologize for the confusion. We'll update your order status as soon as possible. Please provide your order number.",
    },
    {
        'keywords': ['e-transfer', 'not', 'accepted'],
        'response': "We're sorry for the delay in accepting your e-transfer. Please provide your order number and we'll investigate the issue.",
    },
    {
        'keywords': ['tracking', 'not', 'updated'],
        'response': "We apologize for the delay in updating your tracking information. Please provide your order number, and we'll look into the issue and update you as soon as possible.",
    },
    {
        'keywords': ['tracking'],
        'response': "We apologize for the delay in providing your tracking number. We'll send it to you as soon as it's available. Please provide your order number.",
    },
    {
        'keywords': ['order', 'completed'],
        'response': "We typically process orders within 1-2 business days after payment has been accepted. If you haven't received any updates on your order, please provide your order number so we can check the status.",
    },
    {
        'keywords': ['no', 'confirmation', 'order'],
        'response': "We apologize for the lack of confirmation on your order. We'll look into the issue and update you as soon as possible. Please provide your order number.",
    },
    {
        'keywords': ['missing', 'item', 'order'],
        'response': "We're sorry that an item is missing from your order. We'll send the missing item right away. Please share your order number and the missing item's details.",
    },
    {
        'keywords': ['order', 'still', 'on', 'hold'],
        'response': "We apologize for the delay in processing your order. We'll investigate the issue and update the status as soon as possible. Please provide your order number.",
    },
    {
        'keywords': ['status', 'order'],
        'response': "We apologize for any confusion regarding your order status. Please provide your order number, and we'll update you with the current status.",
    },
    {
        'keywords': ['pack', 'delivered', 'not', 'received'],
        'response': "We're sorry that you haven't received your package even though it's marked as delivered. We'll investigate the issue and get back to you as soon as possible. Please provide your order number and shipping address.",
    },
    {
        'keywords': ['pack', 'delayed'],
        'response': "We apologize for the delay in your package's delivery. Please provide your order number and tracking number, and we'll look into the issue.",
    },
    {
        'keywords': ['pack', 'wrong', 'address'],
        'response': "We're sorry that your package was delivered to the wrong address. We'll arrange to resend your order to the correct address. Please provide your order number and the correct shipping address.",
    },
    {
        'keywords': ['wrong', 'order', 'received'],
        'response': "We apologize for the mix-up in your order. We'll arrange to send the correct items immediately. Please provide your order number and a photo of the items you received.",
    },
    {
        'keywords': ['e-transfer', 'pending'],
        'response': "We're sorry for the delay in processing your e-transfer. We'll look into the issue and update you as soon as possible. Please provide your order number.",
    },
    {
        'keywords': ['lootly', 'points', 'update', 'purchase'],
        'response': "We apologize for the issue with your Lootly points not updating. We'll resolve the issue and update your points accordingly. Please provide your order number and the email address associated with your Lootly account.",
    },
    {
        'keywords': ['rewards', 'widget', 'website'],
        'response': "We're sorry for the inconvenience. The rewards widget should be visible on our website. Please try clearing your browser cache or using a different browser. If the issue persists, contact us with your device and browser details.",
    },
    {
        'keywords': ['received', 'smalls', 'regulars'],
        'response': "We apologize for the mix-up in your order. We'll send the correct regular-sized product as soon as possible. Please provide your order number and a photo of the product you received.",
    }
]

template_responses = {tuple(template['keywords']): template['response'] for template in templates}

def extract_keywords(text):
    keywords = []
    doc = nlp(text)
    for token in doc:
        if token.pos_ in ('NOUN', 'VERB', 'PROPN', 'ADJ'):
            keywords.append(token.text)
    return keywords

def find_matching_template(keywords):
    max_score = 0
    best_response = None
    for template_keyword, response in template_responses.items():
        template_keyword_set = set(template_keyword)
        common_keywords = template_keyword_set.intersection(keywords)
        score = len(common_keywords) / len(template_keyword_set)
        if score > max_score:
            max_score = score
            best_response = response
    return best_response if max_score > 0.5 else None

def send_email(service, user_id, email_data, response_text):
    # Find the recipient's email address
    to_email = None
    subject = None
    for header in email_data['payload']['headers']:
        if header['name'].lower() == 'to':
            to_email = header['value']
        if header['name'].lower() == 'subject':
            subject = header['value']
        if to_email is not None and subject is not None:
            break

    if to_email is None:
        print("Recipient's email address not found.")
        return None

    message = MIMEText(response_text)
    message['to'] = to_email
    message['subject'] = 'Re: ' + subject
    create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}
    
    try:
        send_message = (service.users().messages().send(userId=user_id, body=create_message).execute())
        print(F'sent message to {message["to"]} Message Id: {send_message["id"]}')
    except HttpError as error:
        print(F'An error occurred: {error}')
        send_message = None
    
    return send_message

def main():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', ['https://www.googleapis.com/auth/gmail.modify'])
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    user_id = 'me'

    # Get unread emails
    query = "is:unread"
    unread_emails = service.users().messages().list(userId=user_id, q=query).execute()

    print(f"Unread emails: {len(unread_emails['messages'])}")

    message_ids = [email['id'] for email in unread_emails['messages']]
    subject_header_index = None

    for msg_id in message_ids:
     email_data = service.users().messages().get(userId=user_id, id=msg_id).execute()

    print(f"Processing email with ID: {msg_id}")

    if subject_header_index is None:
        for i, header in enumerate(email_data['payload']['headers']):
            if header['name'] == 'Subject':
                subject_header_index = i
                break

    if subject_header_index is not None:
        email_subject = email_data['payload']['headers'][subject_header_index]['value']
    else:
        print("Subject header not found.")
        email_subject = ''

    if 'parts' in email_data['payload']:
        try:
            email_body = base64.urlsafe_b64decode(email_data['payload']['parts'][0]['body']['data']).decode('utf-8')
        except KeyError:
            email_body = ''

    # Process the email
    keywords = extract_keywords(email_subject + ' ' + email_body)

    print(f"Extracted keywords: {keywords}")

    response_text = find_matching_template(keywords)

    if response_text:
        send_email(service, user_id, email_data, response_text)

        # Mark email as read
        service.users().messages().modify(
            userId=user_id, id=msg_id,
            body={'removeLabelIds': ['UNREAD']}
        ).execute()
    else:
        print(f"No matching template found for email ID: {msg_id}")

if __name__ == '__main__':
    main()
