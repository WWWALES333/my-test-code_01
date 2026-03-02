# 技术架构文档

## 1. 项目概述

销售周报自动下载与分类系统 - 从阿里邮箱自动下载销售部门的周报/月报邮件附件，并根据时间自动分类存档。

## 2. 目录结构

```
my-test-code/
├── config.json               # 配置文件（邮箱IMAP、输出路径、过滤规则）
├── main.py                  # 主程序入口
├── requirements.txt         # Python依赖
├── downloaded_history.json  # 下载历史（MD5去重）
├── run_log.json             # 运行日志（最近50条）
├── system.log               # 系统日志
├── prd.md                   # 产品需求文档
├── README.md                # 说明文档
├── 06 销售周报/             # 报告存档目录
│   └── {年份}/
│       ├── {月份}月第{N}周/  # 周报目录
│       │   └── *.docx
│       └── {月份}月报/       # 月报目录
│           └── *.docx
└── __pycache__/             # Python缓存（不提交）
```

## 3. 技术栈

| 技术 | 用途 |
|------|------|
| Python 3.x | 运行环境 |
| imap_tools | 邮箱IMAP连接 |
| email (标准库) | 邮件解析 |
| hashlib | MD5去重 |
| logging | 日志记录 |
| json | 配置和历史数据 |

## 4. 核心模块

### 4.1 WeeklyReportDownloader 类 (main.py)

| 方法 | 职责 |
|------|------|
| `__init__` | 初始化：加载配置、加载下载历史 |
| `_load_config` | 读取 config.json |
| `_load_history` | 读取 downloaded_history.json |
| `_save_history` | 保存下载历史 |
| `_save_run_log` | 保存运行日志（保留最近50条） |
| `connect_mailbox` | 连接阿里邮箱IMAP |
| `search_weekly_report_emails` | 遍历邮箱，搜索周报/月报邮件，解析附件 |
| `download_attachment` | 从邮件中提取附件内容 |
| `_extract_region` | 从文件名提取区域关键词（已实现，未使用） |
| `_extract_time_info` | 解析周报时间（年、月、周次） |
| `_extract_month_info` | 解析月报时间（年、月） |
| `_get_week_number` | 根据日期计算周次（1-7日=第1周，8-14=第2周...） |
| `_get_output_path` | 计算输出文件路径 |
| `download_and_classify` | 主流程：连接→搜索→下载→分类→记录 |
| `_print_summary` | 打印运行总结 |

## 5. 数据流

```
1. 读取 config.json
2. 连接阿里邮箱 IMAP
3. 遍历所有邮件（倒序，从新到旧）
4. 过滤条件：
   - 主题包含"周报"或"月报"
   - 有 .docx/.doc 附件
   - 符合日期过滤（可选）
   - 符合类型过滤（周报/月报/全部）
5. 对每个符合条件的邮件：
   - 下载附件内容
   - 计算MD5，检查是否已下载
   - 解析文件名提取时间信息
   - 生成输出路径
   - 保存文件
   - 记录到历史和日志
6. 断开邮箱连接
7. 保存历史记录和运行日志
8. 打印总结
```

## 6. 关键配置 (config.json)

| 字段 | 说明 |
|------|------|
| imap_server | IMAP服务器地址（默认：imap.qiye.aliyun.com） |
| imap_port | IMAP端口（默认：993） |
| email | 邮箱账号 |
| password | 邮箱授权码 |
| output_root | 报告存档根目录（默认：./06 销售周报） |
| region_keywords | 区域关键词列表 |
| date_filter.enabled | 是否启用日期过滤 |
| date_filter.year | 过滤年份 |
| date_filter.month | 过滤月份 |
| report_type_filter | 报告类型（weekly/monthly/all） |
| download_history_file | 下载历史文件名 |
| log_file | 运行日志文件名 |

## 7. 注意事项

- 区域分类功能 `_extract_region` 已实现但未在输出路径中使用
- 当前输出目录结构为扁平结构：`年度/月份第N周/文件名`
- 区域信息未被用于创建子目录
- 运行日志保留最近50条记录
- 使用MD5哈希进行去重，支持断点续传
