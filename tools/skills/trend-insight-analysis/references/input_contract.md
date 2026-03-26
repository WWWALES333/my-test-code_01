# 输入契约

本 skill 只读取 `v1.5` 规范化后的趋势对象。

## 首选输入
- `data/output/insights/v1.5/normalized/dashboard_snapshot.json`
- `data/output/insights/v1.5/normalized/sales_monthly_rollup.jsonl`

## 可选补充输入
- `data/output/insights/v1.5/normalized/report_facts.jsonl`
- `data/output/insights/v1.5/normalized/evidence_facts.jsonl`

## 关键字段
- 时间字段：`year`、`month`
- 销售字段：`salesperson_id`、`salesperson_name`
- 趋势字段：`ai_mentions`、`ai_report_rate`、`active_sales_count`
- 分布字段：`actor_primary`、`business_line`
- 质量字段：`confirmed_count`、`pending_review_count`

## 禁止依赖
- 禁止直接基于原始 HTML 页面做分析
- 禁止只看单一月度命中量就得出趋势结论
- 禁止忽略销售个人层数据
