#!/usr/bin/env python3
# 尝试使用UID获取更多邮件

import email
import json
import imap_tools
from imap_tools import MailBox

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])

# 使用UID获取所有邮件UID
print("=== 使用 UID FETCH 获取所有邮件 ===")
typ, data = mailbox.client.uid('FETCH', '1:*', '(UID)')
if typ == 'OK' and data:
    uids = []
    for response in data:
        if isinstance(response, tuple) and len(response) >= 2:
            uid_data = response[0]
            if isinstance(uid_data, bytes):
                uid_str = uid_data.decode()
                # 解析UID
                import re
                match = re.search(r'UID (\d+)', uid_str)
                if match:
                    uids.append(int(match.group(1)))

    print(f"找到 UID 数量: {len(uids)}")
    if uids:
        print(f"UID范围: {min(uids)} - {max(uids)}")

# 尝试使用 RANGE 获取
print("\n=== 使用 UID FETCH 1:8015 ===")
typ, data = mailbox.client.uid('FETCH', '1:8015', '(UID FLAGS)')
if typ == 'OK' and data:
    count = len([x for x in data if x])
    print(f"获取到: {count} 条")

mailbox.logout()
