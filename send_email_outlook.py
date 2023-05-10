# -*- coding: utf-8 -*-
"""
Created on Tue May  9 10:25:43 2023

@author: Ernest Namdar
"""

import pandas as pd
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')

# Load data from Excel file
df = pd.read_excel('Sample_Sheet.xlsx', header=None)

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


# Iterate over recipients and send email
for i in range(df.shape[0]):
    first_name, last_name, email, username, password, _ = df.iloc[i,0].split(',')
    message = message_template.format(first_name=first_name, last_name=last_name, username=username, password=password)
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.CC = 'ernest.namdar@utoronto.ca' # Replace with the name of your account
    mail.Subject = subject
    mail.Body = message
    accounts = outlook.Session.Accounts
    for account in accounts:
        if account.DisplayName == 'ernest.namdar@utoronto.ca': # Replace with the name of your account
            mail.SendUsingAccount = account
    mail.Send()
    print(i+1, " Email sent to ", email)
