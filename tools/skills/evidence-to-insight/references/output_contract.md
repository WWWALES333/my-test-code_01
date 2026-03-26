# 输出契约

每张结论卡固定输出以下字段：

- `insight_id`
- `topic`
- `title`
- `summary`
- `business_line`
- `signal_type`
- `evidence_refs`
- `owner_refs`
- `needs_review`
- `confidence`

## 字段说明
- `insight_id`：结论卡主键
- `topic`：结论主题
- `title`：业务可读标题
- `summary`：结论说明，回答“发生了什么、意味着什么”
- `business_line`：`云诊室 | 云管家 | 混合 | 待判断`
- `signal_type`：`opportunity | concern | behavior | feedback | trend`
- `evidence_refs`：代表证据引用
- `owner_refs`：涉及的销售或区域对象
- `needs_review`：`true | false`
- `confidence`：`high | medium | low`

## 输出风格
- 标题必须是业务语言，不得写成“潜在 AI 机会-1”
- `summary` 应简洁、结论化、可执行
- 证据不足时，宁可降低置信度，也不要强行总结
