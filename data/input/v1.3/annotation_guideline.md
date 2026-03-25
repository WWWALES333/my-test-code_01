# v1.3 标注规范（优化版）

## 1. 文档目的
本文档用于统一 `v1.3` 样本的标注与验收口径，重点解决以下问题：
- 单标签无法表达复合语义
- 关键词直判导致语境误判
- 宽泛表达缺少稳定归类口径

## 2. 标注对象
- 冻结样本集中的周报和月报
- 文档中的可分析片段（segment）
- AI 专题相关的范围、业务线、主体、结果和不确定项

## 3. 字段设计（主标签 + 扩展标签）

### 3.1 主兼容字段
- `ai_actor`：兼容旧口径，值等于 `actor_primary`
- `business_line`：`云诊室 | 云管家 | 混合 | 待判断`
- `decision_status`：`confirmed | uncertain | pending_human_review`

### 3.2 新增字段
- `actor_primary`：`销售自用 | 销售对外介绍 | 医生反馈 | 潜在 AI 机会`
- `actor_subtype`：细分标签，允许 1-2 个值，使用 `;` 分隔
- `ai_scope`：`product_ai | market_trend | competitor_ai | general_ai`
- `interaction_outcome`：`no_feedback | positive_feedback | negative_or_observing | converted | not_applicable`
- `certainty_level`：`high | medium | low`

## 4. AI 范围（ai_scope）判定
- `product_ai`：与我方产品能力、销售动作、医生使用或产品机会直接相关
- `market_trend`：行业政策、趋势、外部环境变化、AI 搜索线索
- `competitor_ai`：竞品相关 AI 动态、竞品借 AI 形成竞争压力
- `general_ai`：泛 AI 认知或个人思考，不构成明确业务动作

说明：
- `market_trend` 与 `competitor_ai` 需要单独统计，不直接混入主业务 KPI。

## 5. 业务线（business_line）判定
- 优先使用“片段内容信号 + 文档上下文先验（文件路径、客户类型）”联合判断
- 仅有弱信号时允许 `混合` 或 `待判断`
- 不允许为了凑标签强行落单边业务线

## 6. 主体（actor_primary / actor_subtype）判定
- `actor_primary` 必须是单值，不允许写 `A / B`
- 复合语义放到 `actor_subtype`，例如：
  - `销售介绍后收到反馈`
  - `销售介绍_无明确反馈`
  - `客户AI搜索线索`
  - `行业趋势观察`
  - `竞品AI动态`
  - `销售个人思考`
  - `效率提效机会`

## 7. 互动结果（interaction_outcome）判定
- `no_feedback`：有介绍动作但未出现明确反馈
- `positive_feedback`：出现兴趣、认可、正向态度
- `negative_or_observing`：出现否定、观望、暂不接受
- `converted`：出现明确转化或行为变化
- `not_applicable`：不适用（如纯趋势观察）

## 8. 不确定与复核
- 允许输出 `pending_human_review`，不强行分类
- `review_reason_code` 固定使用：
  - `ACTOR_OVERLAP`
  - `SCOPE_AMBIGUOUS`
  - `BROAD_STATEMENT`
  - `BUSINESSLINE_LOW_SIGNAL`
  - `OUTCOME_UNCLEAR`

## 9. 标注记录格式
每条标注至少包含：
- `sample_id`
- `segment_id`
- `source_text`
- `business_line`
- `ai_actor` / `actor_primary` / `actor_subtype`
- `ai_scope`
- `interaction_outcome`
- `certainty_level`
- `decision_status`
- `review_reason_code`
- `review_comment`

## 10. 文件落点
- 主审批文件：`data/input/v1.3/annotations/labels_full_hits_review.csv`
- 样本级模板：`data/input/v1.3/annotations/labels_template.csv`
- 人工正式版（验收后）：`data/input/v1.3/annotations/labels_v1.csv`
