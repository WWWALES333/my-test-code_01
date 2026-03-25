# 销售周报自动下载与分类系统

当前封板版本：`V1.3`（2026-03-25）

`v1.3` 已完成离线 AI 专题分析 MVP 封板，版本标签：`V1.3`。

## 项目结构

```text
my-test-code_01/
├── src/                     # 核心业务代码
│   ├── main.py
│   └── analysis_v13/        # v1.3 离线分析链路
├── docs/                    # 长期文档、版本文档、截图、参考资料
│   ├── 01_business_context.md
│   ├── 02_domain_glossary.md
│   ├── 10_product_roadmap.md
│   ├── 11_information_architecture.md
│   ├── assets/
│   ├── reference/
│   └── releases/
│       ├── v1.2/
│       └── v1.3/
├── data/
│   ├── input/
│   │   ├── config.json      # 邮箱配置
│   │   └── v1.3/            # v1.3 标注基线与样本清单
│   └── output/              # 程序运行产物（归档、日志、审计）
│       └── insights/v1.3/   # v1.3 专题分析产物
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

3. 运行下载归档主链路（v1.2稳定链路）：
```bash
python src/main.py -c data/input/config.json --once
```

4. 运行 v1.3 离线分析（冻结样本）：
```bash
python3 -m src.analysis_v13.run \
  --samples data/input/v1.3/samples \
  --annotations data/input/v1.3/annotations \
  --out data/output/insights/v1.3 \
  --model-mode mock
```

## 文档入口

- 基线规则：`rules.md`
- 业务背景：`docs/01_business_context.md`
- 术语表：`docs/02_domain_glossary.md`
- 产品路书：`docs/10_product_roadmap.md`
- 信息架构：`docs/11_information_architecture.md`
- `v1.2` 归档文档：`docs/releases/v1.2/`
- `v1.3` 封板文档：`docs/releases/v1.3/`

## 输出位置

- 周报/月报归档：`data/output/sales_reports/06 销售周报/`
- 运行日志：`data/output/runtime/run_log.json`
- 下载历史：`data/output/runtime/downloaded_history.json`
- 审计报告：`data/output/audit/reports/`
- `v1.3` 专题产物：
  - `data/output/insights/v1.3/extracted/`
  - `data/output/insights/v1.3/reports/`

## 数据基线

- 历史归档仅保留 `2023` 年及以后数据。
- `2023` 年之前历史目录已清理，可按需重新拉取。
