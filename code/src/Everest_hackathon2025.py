import imaplib
import email
import os
from email.header import decode_header
from bs4 import BeautifulSoup
import openai
import pdfminer.high_level
import docx
import pandas as pd
import io
import json
from datetime import datetime

# ----------- CONFIG -----------
IMAP_SERVER = "imap.gmail.com"
EMAIL_ACCOUNT = "xxx"
EMAIL_PASSWORD = "xxx"
OPENAI_API_KEY = "xxx"

openai.api_key = OPENAI_API_KEY

# ----------- MAPPING -----------
REQUEST_MAPPING = [
    {"Request type": "Adjustment", "Sub Request type": ""},
    {"Request type": "Accounting unit Transfer", "Sub Request type": ""},
    {"Request type": "Closing Notice", "Sub Request type": "Reallocation Fees"},
    {"Request type": "Closing Notice", "Sub Request type": "Amendment Fees"},
    {"Request type": "Closing Notice", "Sub Request type": "Reallocation Principal"},
    {"Request type": "Commitment Change", "Sub Request type": "Cashless Roll"},
    {"Request type": "Commitment Change", "Sub Request type": "Decrease"},
    {"Request type": "Commitment Change", "Sub Request type": "Increasee"},
    {"Request type": "Fee payment", "Sub Request type": "Ongoing Fee"},
    {"Request type": "Fee payment", "Sub Request type": "Letter of credit Fee"},
    {"Request type": "Money Movement-inbound", "Sub Request type": "Principal"},
    {"Request type": "Money Movement-inbound", "Sub Request type": "Interest"},
    {"Request type": "Money Movement-inbound", "Sub Request type": "Principal+Interest"},
    {"Request type": "Money Movement-inbound", "Sub Request type": "Principal+Interest+Fee"},
    {"Request type": "Money Movement-Outbound", "Sub Request type": "Timebound"},
    {"Request type": "Money Movement-Outbound", "Sub Request type": "Foreign Currency"}
]

# ----------- FUNCTIONS -----------
def extract_text_from_pdf(file_bytes):
    with io.BytesIO(file_bytes) as f:
        return pdfminer.high_level.extract_text(f)

def extract_text_from_docx(file_bytes):
    with io.BytesIO(file_bytes) as f:
        doc = docx.Document(f)
        return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_excel(file_bytes):
    with io.BytesIO(file_bytes) as f:
        df = pd.read_excel(f)
        return df.to_string()

# ----------- EMAIL FETCH -----------
def fetch_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    mail.select("inbox")

    status, messages = mail.search(None, '(UNSEEN)')
    email_ids = messages[0].split()

    emails = []
    for e_id in email_ids:
        status, msg_data = mail.fetch(e_id, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                sender = msg.get("From")
                date = msg.get("Date")
                body = ""
                attachments = []

                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        filename = part.get_filename()

                        if content_type == "text/html" and not filename:
                            body = part.get_payload(decode=True)
                        elif filename:
                            attachments.append({
                                "filename": filename,
                                "data": part.get_payload(decode=True)
                            })

                body = BeautifulSoup(body, "html.parser").get_text()
                emails.append({"subject": subject, "sender": sender, "body": body, "attachments": attachments, "date": date})

    mail.logout()
    return emails

# ----------- ATTACHMENT PROCESSING -----------
def process_attachments(attachments):
    full_text = ""
    for attachment in attachments:
        filename = attachment['filename']
        data = attachment['data']
        if filename.endswith(".pdf"):
            full_text += extract_text_from_pdf(data) + "\n"
        elif filename.endswith(".docx"):
            full_text += extract_text_from_docx(data) + "\n"
        elif filename.endswith(".xlsx"):
            full_text += extract_text_from_excel(data) + "\n"
    return full_text

# ----------- LLM CLASSIFICATION -----------
def classify_with_llm(email_obj, attachment_text):
    combined_text = email_obj['body'] + "\n" + attachment_text

    prompt = f"""
    You are an AI assistant for a bank service team.
    Based on the provided email and attachment content, classify strictly into one of the following request types and sub-request types:

    {json.dumps(REQUEST_MAPPING, indent=2)}

    Also extract the following fields if available:
    - CustomerID
    - LoanID
    - AccountNumber
    - RequestorName
    - CustomerName
    - Amount
    - DueDate
    - RequestTimestamp
    - RequestSummary

    Only return fields that are relevant for the specific Request type & Sub Request type.

    Email subject: {email_obj['subject']}
    Email date: {email_obj['date']}

    Email & Attachment content:
    """
    prompt += combined_text

    response = openai.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}]
    )

    print(response.choices[0].message.content)

emails = fetch_emails()
for email_obj in emails:
        attachment_text = process_attachments(email_obj["attachments"])
        classify_with_llm(email_obj, attachment_text)