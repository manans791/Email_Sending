import pandas as pd
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load the Excel file
print("Loading the Excel file...")
df = pd.read_excel("Emails data.xlsx")
print(f"Excel file loaded successfully! {len(df)} rows found.\n")

# Email configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_SENDER = "manans791@gmail.com"
EMAIL_PASSWORD = "ypbo pooa mfas kooz"  # Use App Password if needed

# Set up SMTP server
print("Connecting to the SMTP server...")
server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(EMAIL_SENDER, EMAIL_PASSWORD)
print("SMTP server connected successfully!\n")

# Iterate through the rows and send emails
for index, row in df.iterrows():
    recipient_email = row['Email']
    recipient_name = row['Name']
    recipient_number = row['Number']
    recipient_address = row['Address']

    print(f"Preparing email {index+1}/{len(df)} for: {recipient_email}")

    subject = "Please Confirm Your Information"
    body = f"""
Dear {recipient_name},

We are reaching out to you on behalf of Lodha Group. Kindly confirm the following details:

Email: {recipient_email}
Name: {recipient_name}
Number: {recipient_number}
Address: {recipient_address}
Looking forward to your confirmation.

Thanks and Regards,
Manan Sharma,
9987384152
Dun & Bradstreet
    """
    
    # Construct the email
    msg = MIMEMultipart()
    msg['From'] = EMAIL_SENDER
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        # Send the email
        server.sendmail(EMAIL_SENDER, recipient_email, msg.as_string())
        # Log the timestamp
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df.at[index, 'Email Sent Timestamp'] = timestamp
        print(f"✅ Email sent successfully to {recipient_email} at {timestamp}\n")
    except Exception as e:
        print(f"❌ Failed to send email to {recipient_email}. Error: {e}\n")

# Save the updated Excel file
print("Saving the updated Excel file with timestamps...")
df.to_excel("data.xlsx", index=False)
print("Excel file saved successfully as 'data.xlsx'!\n")

# Close the SMTP server
server.quit()
print("SMTP server disconnected. All emails processed successfully!")
