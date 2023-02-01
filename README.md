# Email--Validator

The Email Validator is a tool that validates the email addresses of prospects by sending an automated email and detecting hard bounces. The libraries used in this tool include Pandas, os, Openpyxl, smtplib, imaplib, time, and email.

The process involves the following steps:

The sender account must be set up before sending any emails. This account can be from any email provider (such as Gmail, Outlook, Hotmail, etc.).
An environmental variable should be created for the sender account for security purposes.
A list of email addresses to be validated should be added to an Excel sheet.
Upon execution of the code, the tool will automatically send emails to the addresses listed in the Excel sheet and then wait for a specified amount of time. Afterward, it will open the email account and check for any hard bounces and return the validated email addresses.
Finally, the tool will clean up the email account in preparation for the next execution.
