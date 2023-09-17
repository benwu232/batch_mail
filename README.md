# Batch Mail
## -- a tool for batch sending mail based on pyton-o365

- It can send a large number of emails with customization.
- It it based on pyton-o365 to call Office365 api to achieve aucentication on microsoft office.
- The main configuration is in config.yml
- Azure configuration
  - You need to do app registation in Azure/active directory, then you got client id (application id) and tenant id (director id)
  - Choose the third one: Accounts in any organizational directory and personal Microsoft accounts; And select Web in redict URI, and paste https://login.microsoftonline.com/common/oauth2/nativeclient in the right text box. (Please refer to https://github.com/O365/python-o365)
  - Then you need to go to API permissions of the registered app to ask the permission (Mail.send).
  - Then go to Certificates & Secrets / New client secret. After that, fill the seret value as client_secret in config.yml.

- Fill subject in mail_subject.txt
- Fill mail body with html format in mail_body.txt
- Fill recipents' information in recipients.csv
- Add attachments' paths to config.yml
- You can change the mail format in batch_mail.py
- Set work_start and work_end, the mail will be sent in the time span
- Set min_delay and max_delay for the delay period between each mail