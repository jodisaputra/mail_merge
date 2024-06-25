import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Load the recipient list
df = pd.read_csv('recipients.csv')

# Load the HTML email template
with open('email_set outlook.html', 'r') as file:
    email_template = file.read()

# Email server configuration for Outlook (Office 365 or Outlook.com)
smtp_server = 'smtp.office365.com'
smtp_port = 587  # Port 587 is typically used for SMTP over TLS (STARTTLS)

sender_email = os.getenv('SENDER_EMAIL')
sender_password = os.getenv('SENDER_PASSWORD')
subject_email = os.getenv('SENDER_SUBJECT_EMAIL')

# Initialize the SMTP server
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()  # Secure the connection
server.login(sender_email, sender_password)

# Send emails
for index, row in df.iterrows():
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = sender_email
        msg['To'] = row['Email']
        msg['Subject'] = subject_email

        # Personalize the email template
        personalized_email = email_template.replace('{{FirstName}}', row['FirstName'])
        personalized_email = personalized_email.replace('{{LastName}}', row['LastName'])

        # Attach the HTML content to the email
        part = MIMEText(personalized_email, 'html')
        msg.attach(part)

        # Send the email
        server.sendmail(sender_email, row['Email'], msg.as_string())
        print(f"Email sent successfully to {row['Email']}")
    except Exception as e:
        print(f"Failed to send email to {row['Email']}: {e}")

# Close the SMTP server
server.quit()

print("Emails have been sent successfully.")
