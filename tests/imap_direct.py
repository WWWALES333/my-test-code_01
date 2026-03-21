#!/usr/bin/env python3
import imaplib
import email
import json

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

print("=== 使用 imaplib 直接连接 ===")
mail = imaplib.IMAP4_SSL(config['imap_server'])
mail.login(config['email'], config['password'])

# 先选择INBOX
print("=== 选择 INBOX ===")
typ, data = mail.select('INBOX')
print(f"选择结果: {typ}, {data}")

# 获取邮件数量
print("\n=== 获取邮件数量 ===")
typ, data = mail.select('INBOX')
if typ == 'OK':
    # 解析邮件数量
    import re
    match = re.search(r'(\d+)', str(data))
    if match:
        print(f"INBOX 中有 {match.group(1)} 封邮件")

# 搜索所有邮件
print("\n=== 搜索所有邮件 ===")
typ, msg_ids = mail.search(None, 'ALL')
if typ == 'OK':
    ids = msg_ids[0].split()
    print(f"搜索到: {len(ids)} 封")
    if ids:
        print(f"ID范围: {ids[0]} - {ids[-1]}")

# 检查文件夹列表
print("\n=== 检查所有文件夹 ===")
typ, folders = mail.list()
if typ == 'OK':
    for folder in folders:
        print(f"  {folder}")

mail.logout()
