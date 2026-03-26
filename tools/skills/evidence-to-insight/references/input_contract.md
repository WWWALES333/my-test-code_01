# 输入契约

本 skill 只读取 `v1.5` 的证据与结论候选对象。

## 首选输入
- `data/output/insights/v1.5/normalized/evidence_facts.jsonl`
- `data/output/insights/v1.5/normalized/insight_candidates.jsonl`（如实现）

## 可选补充输入
- `data/output/insights/v1.5/normalized/review_decisions.jsonl`
- `data/output/insights/v1.5/normalized/report_facts.jsonl`

## 允许主题
- `cloud_clinic`
- `cloud_steward`
- `sales_behavior`
- `product_opportunity`

## 证据筛选原则
- 优先保留高质量证据
- 降权模板、资料整理、泛 AI 表达
- 同一证据重复出现时去重

## 禁止依赖
- 禁止把单条片段直接包装成结论卡
- 禁止忽略复核状态
- 禁止脱离原始文件路径做结论
