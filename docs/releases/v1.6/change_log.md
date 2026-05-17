# v1.6 变更记录

**版本**：v1.6
**状态**：开发中
**版本名称**：AI 一线情报工作台质量提升版

## 本版目标
- 复核学习闭环。
- 业务问题驱动分析。
- Minimax 边界语义判定。
- 工作台 UI/UX 重构。
- 代码审核与精简治理。

## 计划新增
- `src/analysis_v16/`
- `business_question_facts`
- `review_batch`
- `learning_summary`
- `rule_candidates`
- `prompt_candidates`
- `label_gap_candidates`
- `golden_set`
- v1.6 工作台和周/月摘要

## 当前已完成
- 新增 `src/analysis_v16/`，复用 v1.5 normalized 层，不改 `src/main.py`。
- 新增 Minimax/OpenAI 兼容模型适配层，默认 `https://api.minimaxi.com/v1` + `MiniMax-M2.7`。
- 新增 `<think>` 清洗、混合文本 JSON 提取和模型失败兜底。
- 新增业务问题层：医生接纳度、医生直接诉求、医生间接机会、销售 AI 使用、竞品动作、区域/销售差异。
- 新增 20 条一轮复核批次、复核写回、学习摘要、规则/Prompt/标签候选和黄金样本集。
- 新增 v1.6 本地工作台页面和 `/api/v16-review-decisions` 复核提交接口。
- 新增 `tests/test_analysis_v16.py`，并通过 v1.3-v1.6 相关回归测试。

## 本版不做
- 不做数据库。
- 不做多人权限。
- 不做多专题平台。

## 已知风险
- Minimax 可能输出 `<think>`，必须在适配层清洗。
- 业务问题字段不能一次性过度细分。
- UI/UX 必须早期落地，不能最后再补。
