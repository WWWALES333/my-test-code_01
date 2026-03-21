#!/usr/bin/env python3
# 调试脚本：分析邮件结构，找出附件提取失败的原因

import email
import imap_tools
from imap_tools import MailBox

# 加载配置
import json
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# 连接邮箱
mailbox = MailBox(config['imap_server'], config['imap_port'])
mailbox.login(config['email'], config['password'])

# 搜索所有包含"周报"或"月报"的邮件
typ, msg_ids = mailbox.client.search(None, 'ALL')
msg_id_list = msg_ids[0].split()

# 只分析前10封
count = 0
for msg_id in reversed(msg_id_list[:20]):
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
                subject = ''.join([s[0].decode(s[1] or 'utf-8') if s[0] else '' for s in subject])
            else:
                subject = str(subject)
        else:
            subject = ''

        if '月报' not in subject and '周报' not in subject:
            continue

        count += 1
        print(f"\n{'='*60}")
        print(f"邮件 #{count}: {subject[:50]}")
        print(f"Content-Type: {msg.get_content_type()}")
        print(f"Is Multipart: {msg.is_multipart()}")

        # 分析附件
        attachments_found = []

        if msg.is_multipart():
            print("\n--- 各部分分析 ---")
            for i, part in enumerate(msg.walk()):
                content_type = part.get_content_type()
                filename = part.get_filename()
                content_disp = str(part.get('Content-Disposition', ''))

                # 解码文件名
                if filename:
                    try:
                        decoded = email.header.decode_header(filename)
                        filename = decoded[0][0]
                        if isinstance(filename, bytes):
                            filename = filename.decode('utf-8')
                    except:
                        pass

                print(f"\nPart {i}:")
                print(f"  Content-Type: {content_type}")
                print(f"  Filename: {filename}")
                print(f"  Content-Disposition: {content_disp[:50] if content_disp else 'None'}")

                # 检查是否是附件
                if content_disp and 'attachment' in content_disp:
                    attachments_found.append({
                        'filename': filename,
                        'content_type': content_type,
                        'part': part
                    })
                elif filename and (filename.endswith('.docx') or filename.endswith('.doc') or filename.endswith('.pdf')):
                    attachments_found.append({
                        'filename': filename,
                        'content_type': content_type,
                        'part': part
                    })

        print(f"\n找到附件数量: {len(attachments_found)}")
        for att in attachments_found:
            print(f"  - {att['filename']} ({att['content_type']})")

        if count >= 10:
            break

    except Exception as e:
        print(f"Error: {e}")
        continue

mailbox.logout()
