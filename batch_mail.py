import time
import random
import yaml
import csv
import datetime as dt

from O365 import Account, FileSystemTokenBackend, Message

# 配置 Azure 应用程序信息
# client_id = 'e4da1008-d68e-45e7-a7ed-6a732565da8c'
# client_secret = '-wP8Q~xJgV7bQXtJgkraRrVNVTWzgxVB~EIoIbpc'
# tenant_id = '2d6b2a6d-cbed-45dd-82af-be8b70dd094e'
# redict_url = 'https://login.microsoftonline.com/common/oauth2/nativeclient'

# token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
# token_backend = FileSystemTokenBackend()

# src_mail = 'it@delight.co.nz'
# dst_mail_list = ['benwu232@gmail.com', 'benwu232@gmail.com']


def load_config(filename):
    with open(filename, 'r') as f:
        config = yaml.load(f, Loader=yaml.CLoader)
    return config

# 连接到 Office 365
def get_o365_account(client_id, client_secret, tenant_id, redict_url):
    account = Account(credentials=(client_id, client_secret), auth_flow_type='credentials', tenant_id=tenant_id)
    if account.authenticate():
        print('Authenticated!')
    return account

# def mass_send(account, src_mail, dst_mail_list, subject='test', body='', attachments=[]):
def mass_send(account, cfg):
    if not account.is_authenticated:
        print('Authentication failed.')
        return

    recipients = []
    with open(cfg['recipients'], 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            title, name, email = row
            if title == 'Title':
                continue  # 跟进标题行，不处理'
            recipients.append((title, name, email))
    
    with open(cfg['mail_subject'], 'r') as fp:
        mail_subject = fp.read()

    with open(cfg['mail_body'], 'r') as fp:
        mail_body_raw = fp.read()

    work_start = dt.datetime.strptime(cfg['work_start'], '%H:%M').time()
    work_end = dt.datetime.strptime(cfg['work_end'], '%H:%M').time()

    mailbox = account.mailbox(resource=cfg['src_mail'])
    # for k, (title, name, ds_mail_addr) in enumerate(recipients):
    iter_recipients = iter(recipients)
    k = 1
    while True:
        curtime = dt.datetime.now().time()
        if curtime < work_start or curtime > work_end:
            print('not in work time')
        else:
            title, name, ds_mail_addr = next(iter_recipients)
            print(f"No. {k} Sending to {ds_mail_addr}...")
            m = mailbox.new_message()
            m.to.add(ds_mail_addr)
            m.subject = mail_subject
            m.body = f"{title} {name}, <br><br> {mail_body_raw}"
            m.attachments.add(cfg['mail_attachments'])
            m.send()
            k += 1
            if k > len(recipients):
                break
        delay = random.randint(cfg['min_delay'], cfg['max_delay'])
        print(f'Sleep {delay} seconds...')
        time.sleep(delay)
        # print('ok')        
    print(f'{k-1} emails sent successfully.')


if __name__ == '__main__':
    cfg = load_config('config.yml')
    account = get_o365_account(cfg['client_id'], cfg['client_secret'], cfg['tenant_id'], cfg['redict_url'])
    mass_send(account, cfg)
    pass