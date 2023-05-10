# send_email_with_python
Sending email using Python (Outlook and Gmail)

Sending emails with Python can save a significant amount of time and energy. In my case, I had to send a welcome message to each enrolled in our course and include their username and password in the email. I had an Excel sheet similar to Sample_Sheet.xlsx, with rows in the format of "firstname,lastname,email,username,password,date". If your scenario is different, you should revise the code. However, the main structure of the code remains the same.


send_email_gmail.py uses a Gmail account to send an email for each row in Sample_Sheet.xlsx.

send_email_outlook.py does the same, but it relies on your Outlook software. It saves you from the overhead of dealing with the authentication process.
