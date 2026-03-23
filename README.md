# 销售周报自动下载与分类系统

当前封板版本：`V1.2`（2026-03-23）

## 项目结构

```text
my-test-code_01/
├── src/                     # 核心业务代码
│   └── main.py
├── docs/                    # 文档、截图、参考资料
├── data/
│   ├── input/
│   │   └── config.json      # 邮箱配置
│   └── output/              # 程序运行产物（归档、日志、审计）
├── tests/                   # 排查与测试脚本
├── requirements.txt
├── rules.md
└── README.md
```

## 使用方法

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 修改 `data/input/config.json` 配置邮箱信息。

3. 运行程序：
```bash
python src/main.py -c data/input/config.json --once
```

## 输出位置

- 周报/月报归档：`data/output/sales_reports/06 销售周报/`
- 运行日志：`data/output/runtime/run_log.json`
- 下载历史：`data/output/runtime/downloaded_history.json`
- 审计报告：`data/output/audit/reports/`

## 数据基线

- 历史归档仅保留 `2023` 年及以后数据。
- `2023` 年之前历史目录已清理，可按需重新拉取。
