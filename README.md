# 销售周报自动下载与分类系统

## 项目结构

```
sales-weekly-report/
├── config.json          # 邮箱配置
├── main.py             # 主程序
├── requirements.txt    # 依赖
└── README.md           # 说明文档
```

## 配置说明

### config.json 字段说明

| 字段 | 说明 | 示例 |
|------|------|------|
| imap_server | IMAP服务器地址 | imap.qiye.aliyun.com |
| imap_port | IMAP端口 | 993 |
| email | 邮箱账号 | your_email@company.com |
| password | 邮箱密码或授权码 | xxxxxx |
| output_root | 周报存档根目录 | /path/to/06 销售周报 |
| region_keywords | 区域关键词列表 | ["将军汤", "广西区域", ...] |

## 使用方法

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 修改 config.json 配置您的邮箱信息

3. 运行程序：
```bash
python main.py
```

## 区域关键词配置

默认支持以下区域关键词（可在 config.json 中修改）：
- 将军汤
- 广西区域
- 粤琼区域
- 华中区域
- 华东区域
- 华北区域
- 西南区域

## 日志说明

- 运行日志保存在当前目录：run_log.json
- 下载记录保存在当前目录：downloaded_history.json
