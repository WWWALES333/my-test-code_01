# v1.6 技术方案

## 1. 系统形态
`v1.6` 继续采用离线分析 + 本地 Web 工作台。基础抽取、归一和 v1.5 中间层复用 `src/analysis_v15/`，新增能力放在 `src/analysis_v16/`。

## 2. 代码边界
- `src/main.py`：下载、归档、审计、通知。
- `src/analysis_v15/`：可复用的抽取、归一、销售画像、基础工作台数据。
- `src/analysis_v16/`：业务问题识别、Minimax 适配、20 条复核批次、学习候选、v1.6 页面和报告。

## 3. Minimax 适配
使用 OpenAI 兼容接口：
- `OPENAI_BASE_URL=https://api.minimaxi.com/v1`
- `OPENAI_MODEL=MiniMax-M2.7`
- `OPENAI_API_KEY` 从环境变量或 macOS Keychain 读取。

适配层必须处理：
- `<think>...</think>` 清洗。
- 混合文本中的 JSON 提取。
- JSON 字段校验。
- 失败重试和失败兜底。

## 4. 数据流
1. 运行 `v1.5` 基础链路生成 report/evidence/trend/profile。
2. `v1.6` 对 evidence_facts 做业务问题识别。
3. 生成 business_question_facts、business_insights 和 business_question_summary。
4. 生成每轮 20 条 review_batch。
5. 人工复核写入 review_decisions。
6. 系统生成学习摘要、规则候选、Prompt 候选、标签扩展候选和黄金样本集。

## 5. 复核设计
复核页面只让用户判断业务结果，不要求用户判断该改规则还是 Prompt。系统根据复核前后差异自动归因。

## 6. 代码治理
- 不复制整套 v1.5。
- Prompt、标签枚举、业务问题定义保持单一事实源。
- 封板前删除临时代码，跑测试和敏感信息扫描。
