import imaplib
import email
from email.header import decode_header
import time
import pywhatkit
import os

# Email credentials
EMAIL = "shaikjakeer9346@gmail.com"  
PASSWORD = "fhdr dixq fgnt glmx"  
IMAP_SERVER = "imap.gmail.com"

# WhatsApp details
WHATSAPP_NUMBER = "+919346758699"  
DOWNLOAD_FOLDER = "attachments"  

# Ensure the download folder exists
if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)

def decode_subject(subject):
    """Decode email subject if encoded."""
    decoded, encoding = decode_header(subject)[0]
    if isinstance(decoded, str):
        return decoded
    # If it'S bytes,decode it
    return decoded.decode(encoding or "utf-8")
    

def save_attachment(part, filename):
    """Save attachment (PDF or Image) locally."""
    filepath = os.path.join(DOWNLOAD_FOLDER, filename)
    with open(filepath, "wb") as f:
        f.write(part.get_payload(decode=True))
    return filepath

def check_latest_unread_email():
    """Check for the latest unread email and send a WhatsApp message."""
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select("inbox")

    # Search for the most recent unread email
    status, messages = mail.search(None, "UNSEEN")
    if status != "OK":
        print("No new emails found.")
        return

    email_ids = messages[0].split()
    if not email_ids:
        print("No unread emails.")
        return

    latest_email_id = email_ids[-1]  
    status, msg_data = mail.fetch(latest_email_id, "(RFC822)")
    if status != "OK":
        print("Error fetching email.")
        return

    raw_email = msg_data[0][1]
    email_message = email.message_from_bytes(raw_email)

    # Extract details
    subject = decode_subject(email_message["Subject"])
    sender = email_message["From"]
    body = ""
    attachment_paths = []  

    # Process email content and attachments
    for part in email_message.walk():
        content_type = part.get_content_type()

        # Extract email text content
        if content_type == "text/plain":
            try:
                body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
            except UnicodeDecodeError:
                body = "Unable to decode email content"

        # Check for PDF or Image attachments
        if part.get_filename():
            filename = decode_subject(part.get_filename())
            if filename.endswith((".pdf", ".jpg", ".jpeg", ".png")):
                filepath = save_attachment(part, filename)
                attachment_paths.append(filepath)

    # Prepare WhatsApp message
    whatsapp_message = f" New Email Alert \nFrom: {sender}\nSubject: {subject}\nBody: {body[:100]}..."

    try:
        # Send WhatsApp message
        pywhatkit.sendwhatmsg_instantly(WHATSAPP_NUMBER, whatsapp_message)
        print("WhatsApp message sent successfully!")

        # Send attachments via WhatsApp
        for attachment in attachment_paths:
            pywhatkit.sendwhats_document(WHATSAPP_NUMBER, attachment)
            print(f"Attachment sent successfully: {attachment}")

    except Exception as e:
        print(f"Failed to send WhatsApp message: {e}")

    # Mark email as read after processing
    mail.store(latest_email_id, "+FLAGS", "\\Seen")

    mail.logout()

def main():
    print("Starting Mail to WhatsApp connector... Press Ctrl + C to stop.")
    try:
        while True:
            check_latest_unread_email()
            time.sleep(60)  
    except KeyboardInterrupt:
        print("\nScript stopped by user. Exiting...")

if __name__ == "__main__":
    main()
