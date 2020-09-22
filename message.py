import smtplib

from email.message import EmailMessage
from email.mime.text import MIMEText

from datetime import date

def error_message(SMTP_SERVER, sender_email, receiver_email, message):
    msg = EmailMessage()
    
    msg['Subject'] = 'ERROR: Documentation upload'
    msg['From'] = sender_email
    msg['To'] = receiver_email

    msg.set_content(message)

    s = smtplib.SMTP(SMTP_SERVER)
    s.send_message(msg)
    s.quit()

