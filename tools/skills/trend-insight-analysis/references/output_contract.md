# 输出契约

每条趋势判断固定输出以下字段：

- `period`
- `metric_name`
- `change_type`
- `change_summary`
- `business_interpretation`
- `drilldown_targets`
- `evidence_refs`
- `confidence`

## 字段说明
- `period`：例如 `2026Q1 vs 2025Q1` 或 `2026-03 vs 2026-02`
- `metric_name`：例如 `AI提及量`、`活跃销售数`、`销售渗透率`
- `change_type`：`up | down | flat | volatile`
- `change_summary`：客观变化描述
- `business_interpretation`：业务解释，回答“意味着什么”
- `drilldown_targets`：建议下钻的销售、区域或证据对象
- `evidence_refs`：用于支撑该趋势判断的证据引用
- `confidence`：`high | medium | low`

## 输出风格
- 先给判断，再给解释
- 禁止复读统计表
- 禁止把不完整输入包装成高置信结论
