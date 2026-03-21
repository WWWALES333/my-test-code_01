#!/usr/bin/env python3
# 分析邮箱中的发件人

import email
import json
import re

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

import imap_tools
from imap_tools import MailBox

mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])

typ, msg_ids = mailbox.client.search(None, 'ALL')
msg_id_list = msg_ids[0].split()

senders = {}
count = 0
for msg_id in reversed(msg_id_list[:300]):  # 分析最近300封
    try:
        typ, msg_data = mailbox.client.fetch(msg_id, '(RFC822)')
        if typ != 'OK':
            continue
        msg_bytes = msg_data[0][1]
        msg = email.message_from_bytes(msg_bytes)

        # 解码主题
        subject_raw = msg.get('subject', '')
        if subject_raw:
            subject = email.header.decode_header(subject_raw)
            if isinstance(subject, list):
                subject_str = ''.join([s[0].decode(s[1] or 'utf-8') if s[0] else '' for s in subject])
            else:
                subject_str = str(subject)
        else:
            subject_str = ''

        if '周报' not in subject_str and '月报' not in subject_str:
            continue

        # 提取发件人
        from_header = msg.get('from', '')
        # 简单提取邮箱
        match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+)', from_header)
        if match:
            email_addr = match.group(1)
        else:
            email_addr = from_header[:50]

        if email_addr not in senders:
            senders[email_addr] = []
        senders[email_addr].append(subject_str[:50])
        count += 1

    except Exception as e:
        continue

mailbox.logout()

print(f"找到 {len(senders)} 个发送周报/月的发件人:\n")
for sender, subjects in sorted(senders.items(), key=lambda x: -len(x[1])):
    print(f"{sender}: {len(subjects)} 封")
    for s in subjects[:5]:
        print(f"  - {s}")
    print()
