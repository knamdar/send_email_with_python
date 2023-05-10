# -*- coding: utf-8 -*-
"""
Created in May  2023

@author: Ernest Namdar
"""

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Load data from Excel file
df = pd.read_excel('Samle_Sheet.xlsx', header=None)

# Set email template
subject = 'Welcome to AI in Medicine!' #Change the subject and the message template based on your need
message_template = '''
Dear {first_name},

Thank you for joining our class AI in Medicine. I’m sending you your new login info to the class web space. There you will find the assignments, announcements, and more.

User: {username}
Pass: {password}

Note that this new login info may take 24h to be activated. Please let me know if you have any questions and see you Wednesday evening!

Best,

Ernest Namdar
Ph.D. Candidate
Institute of Medical Science
University of Toronto
'''

# Set up SMTP server
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'ernest.namdar@gmail.com'
# if you have 2factor authentication, you ned to create an app password through your gmail account security settings
smtp_password = ''
smtp_from = 'ernest.namdar@gmail.com'

# Iterate over recipients and send email
for i in range(df.shape[0]):
    first_name, last_name, email, username, password, _ = df.iloc[i,0].split(',')
    message = message_template.format(first_name=first_name, last_name=last_name, username=username, password=password)
    msg = MIMEMultipart()
    msg['From'] = smtp_from
    msg['To'] = email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    smtp = smtplib.SMTP(smtp_server, smtp_port)
    smtp.starttls()
    smtp.login(smtp_username, smtp_password)
    smtp.sendmail(smtp_from, email, msg.as_string())
    smtp.quit()
    print(i+1, " Email sent to ", email)
