#!/usr/bin/env python3
import email
import json
import imap_tools
from imap_tools import MailBox

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# 方式1: 标准搜索
print("=== 方式1: 标准搜索 ALL ===")
mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])
typ, msg_ids = mailbox.client.search(None, 'ALL')
print(f"ALL: {len(msg_ids[0].split()) if msg_ids[0] else 0} 封")
mailbox.logout()

# 方式2: 直接获取邮箱状态
print("\n=== 方式2: 直接获取邮箱状态 ===")
mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])
status = mailbox.client.status('INBOX', '(MESSAGES UNSEEN RECENT)')
print(f"INBOX状态: {status}")
mailbox.logout()

# 方式3: 使用不同的搜索方式
print("\n=== 方式3: 使用不同的搜索方式 ===")
mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])
for charset in [None, 'UTF-8', 'GBK']:
    try:
        if charset:
            typ, msg_ids = mailbox.client.search(charset, 'ALL')
        else:
            typ, msg_ids = mailbox.client.search(None, 'ALL')
        print(f"charset={charset}: {typ}, {len(msg_ids[0].split()) if msg_ids[0] else 0} 封")
    except Exception as e:
        print(f"charset={charset}: error - {e}")
mailbox.logout()

# 方式4: 尝试获取文件夹列表
print("\n=== 方式4: 检查文件夹 ===")
mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])

for folder in mailbox.folder.list():
    try:
        mailbox.folder.set(folder.name)
        typ, msg_ids = mailbox.client.search(None, 'ALL')
        count = len(msg_ids[0].split()) if msg_ids[0] else 0
        print(f"文件夹 '{folder.name}': {count} 封")
    except Exception as e:
        print(f"文件夹 '{folder.name}': 错误 - {e}")

mailbox.logout()
