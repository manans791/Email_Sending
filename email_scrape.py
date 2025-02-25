import imaplib
import email
import re
import openpyxl
from bs4 import BeautifulSoup

# Email Account Credentials
EMAIL_USER = "manans791@gmail.com"
EMAIL_PASS = "ypbo pooa mfas kooz"
IMAP_SERVER = "imap.gmail.com"

print("[INFO] Connecting to the mail server...")

# Connect to the Mailbox
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL_USER, EMAIL_PASS)
print("[SUCCESS] Logged in successfully!")

mail.select("inbox")
print("[INFO] Inbox selected.")

# Search for all emails
status, email_ids = mail.search(None, "ALL")
email_ids = email_ids[0].split()

# Get the latest 100 emails (or all if less than 100)
email_ids = email_ids[-100:] if len(email_ids) > 100 else email_ids
print(f"[INFO] Processing {len(email_ids)} emails.")

emails_data = []

# Function to extract email addresses
def extract_emails(text):
    return re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)

# Function to clean HTML and extract only text from relevant tags
def extract_clean_text(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    allowed_tags = ["p", "div", "span", "b", "i", "strong", "em"]  # Tags that contain text
    extracted_text = " ".join(tag.get_text(strip=True) for tag in soup.find_all(allowed_tags))
    return extracted_text

# Fetch Emails and Extract Data
print("[INFO] Extracting email data...")
for index, e_id in enumerate(email_ids, start=1):
    status, msg_data = mail.fetch(e_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            email_from = msg.get("From", "Unknown")
            email_to = msg.get("To", "Unknown")
            subject = msg.get("Subject", "No Subject")

            # Extract emails from the 'From' and 'To' fields
            extracted_from = extract_emails(email_from)
            extracted_to = extract_emails(email_to)

            # Convert lists to strings
            email_from_cleaned = ", ".join(extracted_from) if extracted_from else "Unknown"
            email_to_cleaned = ", ".join(extracted_to) if extracted_to else "Unknown"

            # Extract email body
            email_body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    if content_type == "text/plain":
                        email_body = part.get_payload(decode=True).decode(errors="ignore")
                        break
                    elif content_type == "text/html":
                        html_body = part.get_payload(decode=True).decode(errors="ignore")
                        email_body = extract_clean_text(html_body)
                        break
            else:
                raw_body = msg.get_payload(decode=True).decode(errors="ignore")
                email_body = extract_clean_text(raw_body) if "<html" in raw_body else raw_body

            # Store email data
            emails_data.append([email_from_cleaned, email_to_cleaned, subject, email_body])
            print(f"[INFO] ({index}/{len(email_ids)}) Processed: {email_from_cleaned} -> {email_to_cleaned}")

# Save data to Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Scraped Emails"
ws.append(["From", "To", "Subject", "Body"])

for row in emails_data:
    ws.append(row)

wb.save("scraped_emails.xlsx")
print("[SUCCESS] Emails saved to scraped_emails.xlsx.")

# Close connection
mail.logout()
print("[INFO] Logged out and disconnected from mail server.")

print("Scraping complete! 🚀")
