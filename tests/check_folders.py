#!/usr/bin/env python3
# 检查邮箱文件夹结构

import json
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

import imap_tools
from imap_tools import MailBox

mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])

# 列出所有文件夹
folders = mailbox.folder.list()
print("邮箱文件夹列表:")
for folder in folders:
    print(f"  - {folder}")

mailbox.logout()
