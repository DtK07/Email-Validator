import pandas as pd
import openpyxl
import os
import smtplib
import email
from email.message import EmailMessage
import imaplib
import time

#mail_id and password
mail_id = os.environ.get("EmailAddress")
pwd = os.environ.get("EmailPassword")

#Read excel to retrieve list of Mail Ids to be validated
def read_excel(File_Name = None, Sheet_Name = None):
    df = pd.read_excel(File_Name, Sheet_Name)
    contacts = []
    for i,row in df.iterrows():
        contacts.append(row[0])
    return contacts

#Send email to all the mail Ids retrieved from an excel
def send_mail(contact=None):
    msg = EmailMessage()
    msg['Subject'] = 'How about a short discussion?'
    msg['From'] = mail_id
    msg['To'] = contact
    msg.set_content('**message**')
    smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    smtp.login(mail_id, pwd)
    smtp.send_message(msg)
    smtp.close()
#Read Inbox to detect the bounced mails and store the returned message as a string
def read_inbox():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(mail_id, pwd)
    status, messages = mail.select("Inbox")
    messages = int(messages[0])
    N = messages
    message = ""
    for i in range(messages, messages-N, -1):
        res, msg = mail.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                raw_msg = email.message_from_bytes(response[1])
                message = str(message) + str(raw_msg)
    mail.close()
    mail.logout()
    return message
#Read spam box to detect the bounced mails and store the returned message as a string
def read_spambox():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(mail_id, pwd)
    status, messages = mail.select("[Gmail]/Spam")
    messages = int(messages[0])
    N = messages
    message = ""
    for i in range(messages, messages-N, -1):
        res, msg = mail.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                raw_msg = email.message_from_bytes(response[1])
                message = str(message) + str(raw_msg)
    mail.close()
    mail.logout()
    return message
#Delete the mails to clear the inbox
def delete_spambox():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(mail_id, pwd)
    mail.select("[Gmail]/Spam")
    status, message_id_list = mail.search(None, "ALL")
    messages = message_id_list[0].split(b' ')
    count = 1
    for m in messages:
        mail.store(m.decode(), "+FLAGS", "\\Deleted")
        count += 1
    mail.expunge()
    mail.close()
    mail.logout()
    return "All the mails in inbox are deleted"
#delete spam box
def delete_inbox():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(mail_id, pwd)
    mail.select("Inbox")
    status, message_id_list = mail.search(None, "ALL")
    messages = message_id_list[0].split(b' ')
    count = 1
    for m in messages:
        mail.store(m.decode(), "+FLAGS", "\\Deleted")
        count += 1
    mail.expunge()
    mail.close()
    mail.logout()
    return "All the mails in spam box are deleted"
#delete sent box
def delete_sentbox():
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(mail_id, pwd)
    mail.select("[Gmail]/Sent")
    status, message_id_list = mail.search(None, "ALL")
    messages = message_id_list[0].split(b' ')
    count = 1
    for m in messages:
        mail.store(m.decode(), "+FLAGS", "\\Deleted")
        count += 1
    mail.expunge()
    mail.close()
    mail.logout()
    return "All the mails in sent box are deleted"
#Validate the retrieved mails
def email_validation(contacts=None):
    returned_message1 = read_inbox()
    returned_message2 = read_spambox()
    returned_message = returned_message1 + returned_message2
    checked_mails = []
    for contact in contacts:
        checklist1 = f"Final-Recipient: RFC822; {contact}\nAction: failed"
        checklist2 = f"Final-Recipient: rfc822; {contact}\nAction: failed"
        if checklist1 in returned_message:
            checked_mails.append(f"{contact} is invalid")
        elif checklist2 in returned_message:
            checked_mails.append(f"{contact} is invalid")
        else:
            checked_mails.append(f"{contact} is valid")
    return checked_mails

if __name__ == '__main__':
    contacts = read_excel(File_Name = "Filename.xlsx", Sheet_Name = "Email List")
    print(contacts)
    for contact in contacts:
        print(f"Sending mail to {contact}")
        send_mail(contact)
        print(f"Mail sent to {contact}")
        time.sleep(5)
    print("Validating...Please wait")
    time.sleep(300)
    returned_mails = email_validation(contacts)
    if returned_mails == "":
        print("No mails to return")
    else:
        print(returned_mails)
    try:
        delete_status = delete_inbox()
        print(delete_status)
    except:
        pass
    try:
        delete_status = delete_spambox()
        print(delete_status)
    except:
        pass









