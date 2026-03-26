# 销售周报自动下载与分类系统

当前封板版本：`V1.2`（2026-03-23）
当前开发版本：`v1.4`（AI专题试运行版，离线开发中）

## 项目结构

```text
my-test-code_01/
├── src/                     # 核心业务代码
│   ├── main.py
│   ├── analysis_v13/        # v1.3 离线分析模块（MVP）
│   └── analysis_v14/        # v1.4 离线分析模块（试运行）
├── docs/                    # 长期文档、版本文档、截图、参考资料
│   ├── 01_business_context.md
│   ├── 02_domain_glossary.md
│   ├── 10_product_roadmap.md
│   ├── 11_information_architecture.md
│   ├── assets/
│   ├── reference/
│   └── releases/
│       ├── v1.2/
│       ├── v1.3/
│       └── v1.4/
├── data/
│   ├── input/
│   │   ├── config.example.json  # 配置模板（可提交）
│   │   ├── config.json          # 本地私有配置（已忽略，不提交）
│   │   ├── v1.3/                # v1.3 冻结样本与标注基线
│   │   └── v1.4/                # v1.4 试运行样本与复核记录
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

2. 初始化本地配置（仅本地使用，不提交）：
```bash
cp data/input/config.example.json data/input/config.json
```

3. 修改 `data/input/config.json` 配置邮箱信息。

4. 运行程序：
```bash
python src/main.py -c data/input/config.json --once
```

5. 运行 `v1.4` 离线分析（mock 模式）：
```bash
python -m src.analysis_v14.run \
  --samples data/input/v1.4/samples \
  --annotations data/input/v1.4/annotations \
  --out data/output/insights/v1.4 \
  --model-mode mock
```

> 安全要求：`data/input/config.json`、`.claude/settings.local.json` 含本地敏感信息，禁止提交到 GitHub。

## 安全发布检查（必做）

1. 一次性启用仓库 pre-commit 钩子：
```bash
git config core.hooksPath .githooks
```

2. 发布前手动执行敏感信息检查：
```bash
python3 tests/check_no_secrets.py
```

检查结果必须为 `[PASS]` 才允许打包或发布到 GitHub。

## 文档入口

- 基线规则：`rules.md`
- 业务背景：`docs/01_business_context.md`
- 术语表：`docs/02_domain_glossary.md`
- 产品路书：`docs/10_product_roadmap.md`
- 信息架构：`docs/11_information_architecture.md`
- `v1.2` 归档文档：`docs/releases/v1.2/`
- `v1.3` 基线文档：`docs/releases/v1.3/`
- `v1.4` 试运行文档：`docs/releases/v1.4/`

## 输出位置

- 周报/月报归档：`data/output/sales_reports/06 销售周报/`
- 运行日志：`data/output/runtime/run_log.json`
- 下载历史：`data/output/runtime/downloaded_history.json`
- 审计报告：`data/output/audit/reports/`

## 数据基线

- 历史归档仅保留 `2023` 年及以后数据。
- `2023` 年之前历史目录已清理，可按需重新拉取。
