import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import logging

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(filename='email_errors.log', level=logging.ERROR)

# Load the recipient list
df = pd.read_csv('recipients.csv')

# Validate CSV columns
required_columns = {'FirstName', 'Email', 'Cc'}
if not required_columns.issubset(df.columns):
    raise ValueError(f"CSV file must contain the following columns: {required_columns}")

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

# Send emails and track status
status_report = []

for index, row in df.iterrows():
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = sender_email
        msg['To'] = row['Email']
        msg['Subject'] = subject_email
        
        # Add CC addresses
        if pd.notna(row['Cc']):
            msg['Cc'] = row['Cc']
            cc_addresses = row['Cc'].split(';')
        else:
            cc_addresses = []

        # Personalize the email template
        personalized_email = email_template.replace('{{FirstName}}', row['FirstName'])
        
        # Check if LastName column exists and replace accordingly
        if 'LastName' in row and pd.notna(row['LastName']):
            personalized_email = personalized_email.replace('{{LastName}}', row['LastName'])
            personalized_message = f"Dear {row['FirstName']} {row['LastName']},<br><br>"
        else:
            personalized_email = personalized_email.replace('{{LastName}}', '')
            personalized_message = f"Dear {row['FirstName']},<br><br>"

        personalized_email = personalized_message + personalized_email

        # Attach the HTML content to the email
        part = MIMEText(personalized_email, 'html')
        msg.attach(part)

        # Send the email
        to_addresses = [row['Email']] + cc_addresses
        server.sendmail(sender_email, to_addresses, msg.as_string())
        status_report.append((row['Email'], 'Success'))
        print(f"Email sent successfully to {row['Email']} with CC to {row['Cc']}")
    except Exception as e:
        logging.error(f"Failed to send email to {row['Email']}: {e}")
        status_report.append((row['Email'], 'Failed'))
        print(f"Failed to send email to {row['Email']}: {e}")

# Close the SMTP server
server.quit()

# Save the status report to a CSV file
status_df = pd.DataFrame(status_report, columns=['Email', 'Status'])
status_df.to_csv('email_status_report.csv', index=False)

print("Emails have been sent. Status report saved to 'email_status_report.csv'.")
