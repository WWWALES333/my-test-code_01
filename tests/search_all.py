#!/usr/bin/env python3
# 全面搜索所有周报/月报

import email
import json
import re

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

import imap_tools
from imap_tools import MailBox

mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])

# 搜索所有邮件
typ, msg_ids = mailbox.client.search(None, 'ALL')
msg_id_list = msg_ids[0].split()
total = len(msg_id_list)

print(f"邮箱中共有 {total} 封邮件\n")

# 统计
stats = {
    'weekly': 0,  # 周报
    'monthly': 0,  # 月报
    'other': 0    # 其他
}

subjects = {'weekly': [], 'monthly': [], 'other': []}

for msg_id in msg_id_list:
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

        if not subject_str:
            continue

        # 判断类型
        if '月报' in subject_str:
            stats['monthly'] += 1
            subjects['monthly'].append(subject_str[:60])
        elif '周报' in subject_str:
            stats['weekly'] += 1
            subjects['weekly'].append(subject_str[:60])
        else:
            stats['other'] += 1

    except:
        continue

mailbox.logout()

print(f"统计结果:")
print(f"  周报: {stats['weekly']} 封")
print(f"  月报: {stats['monthly']} 封")
print(f"  其他: {stats['other']} 封\n")

print("最近10封周报:")
for s in subjects['weekly'][-10:]:
    print(f"  - {s}")

print("\n最近10封月报:")
for s in subjects['monthly'][-10:]:
    print(f"  - {s}")

print("\n最早的5封周报:")
for s in subjects['weekly'][:5]:
    print(f"  - {s}")

print("\n最早的5封月报:")
for s in subjects['monthly'][:5]:
    print(f"  - {s}")
