#!/usr/bin/env python3
import imaplib
import json

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

print("=== 检查所有邮箱文件夹 ===")
mail = imaplib.IMAP4_SSL(config['imap_server'])
mail.login(config['email'], config['password'])

# 获取所有文件夹
typ, folders = mail.list()
if typ == 'OK':
    for folder in folders:
        print(f"  {folder}")
        # 尝试选择并获取数量
        folder_name = folder.decode().split('"')[-2] if '"' in folder.decode() else folder.decode()
        try:
            typ, data = mail.select(f'"{folder_name}"', readonly=True)
            if typ == 'OK':
                import re
                match = re.search(r'(\d+)', str(data))
                if match:
                    print(f"    -> {match.group(1)} 封")
        except:
            pass

mail.logout()
