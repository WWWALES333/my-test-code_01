#!/usr/bin/env python3
import email
import json
import imap_tools
from imap_tools import MailBox

with open('config.json', 'r') as f:
    config = json.load(f)

mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])

typ, msg_ids = mailbox.client.search(None, 'ALL')
msg_id_list = msg_ids[0].split()

print(f"总邮件数: {len(msg_id_list)}")

# 找最早的邮件
print("\n最早的10封邮件日期:")
for msg_id in msg_id_list[:10]:
    try:
        typ, msg_data = mailbox.client.fetch(msg_id, '(RFC822)')
        if typ != 'OK':
            continue
        msg_bytes = msg_data[0][1]
        msg = email.message_from_bytes(msg_bytes)
        date = msg.get('date', '')
        subject = msg.get('subject', '')
        print(f"  {date[:40]}")
        print(f"    Subject: {subject[:50]}")
    except:
        continue

mailbox.logout()
