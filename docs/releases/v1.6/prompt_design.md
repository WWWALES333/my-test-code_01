# v1.6 Prompt 设计说明

## 设计目标
`v1.6` 的 Minimax 调用不再只接收“分类/总结任务”，而是先接收稳定的业务背景包，再执行边界样本判定或证据簇洞察归纳。

核心目标是避免模型因为缺少上下文，把销售自述、医生反馈、市场观察、竞品动作和公司内部机会混在一起。

## Prompt 背景来源
- `docs/01_business_context.md`
- `docs/02_domain_glossary.md`
- `data/input/v1.6/business_question_taxonomy.md`
- `docs/releases/v1.6/PRD.md`

运行时使用的压缩背景版本为：`v1.6-business-context-20260518`。

## 规则与模型分工
- 规则负责：文件读取、时间窗口、同比/环比、基础聚合、明显非 AI 噪声、明显格式化字段。
- Minimax 负责：上下文语义理解、角色区分、医生反馈/销售自述/市场观察边界判断、证据簇洞察归纳。
- 人工复核负责：纠偏结果写回、沉淀黄金样本、形成规则/Prompt/标签候选。

## 边界样本判定 Prompt
模型需要先判断：
- 谁在说
- 谁在行动
- 说给谁
- 发生在哪个业务动作里
- 是否和 AI 有真实业务关系

新增结构化字段：
- `speaker_role`
- `business_actor`
- `evidence_type`
- `reasoning_summary`

## 洞察归纳 Prompt
模型必须输出：
- `insight_title`
- `conclusion`
- `evidence_basis`
- `trend_judgement`
- `driving_factors`
- `counter_evidence_or_uncertainty`
- `why_it_matters`
- `product_implication`
- `sales_management_implication`
- `action_recommendation`
- `caveats`

## 运行产物
每次 `v1.6` 运行会输出：
- `data/output/insights/v1.6/normalized/prompt_context.json`
- `data/output/insights/v1.6/reports/当前使用Prompt说明.md`
- `data/output/insights/v1.6/run_manifest.json` 中的 `prompt_context_version`

## 重要边界
- Prompt 不替代事实数据和证据追溯。
- Prompt 不自动修改规则。
- Prompt 候选来自复核结果，但必须经过人工确认和回归测试后才能落地。
