Info:
Email validator validates the Mail IDs of the prospects by sending them an automated mail and detecting the hard bounces.

Libraries used:
1.Pandas
2.os
3.Openpyxl
4.smtplib
5.imaplib
6.time
7.email

Process:
1. Initially for sending the mails sender account needs to be set-that may be of any account(gmail,outlook,hotmail,etc.,).
2. Once sender account is ready create an environmental variable for that account. (For security purpose)
2. Create an excel with the list of mails which needs to validated.
3. Once the code is executed, it will automatiaaly send the mail to the lists and then cools for particular time then open the mail box and
  check all the existing mailboxes for the bounces and then returns the validated mail IDs.
4. Finally cleans up the mail box and set for next execution.
