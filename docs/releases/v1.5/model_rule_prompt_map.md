# V1.5 模型 / 规则 / 提示词参与点说明

## 1. 文档目的
这份材料回答三个问题：

- 当前系统里哪些步骤是纯规则 / 纯代码
- 哪些步骤是真正由模型参与
- 哪些问题优先查规则、查数据，哪些才应该查提示词

## 2. 当前系统总体判断
当前 `v1.5` 的总体实现状态是：

- 规则主导
- 模型点状参与
- 提示词参与范围有限

更具体地说：
- 文件读取、正文抽取、时间解析、owner 归一、趋势聚合、页面渲染、复核写回，都是规则/代码逻辑
- 真正的运行时模型参与，当前只出现在 [`src/analysis_v14/tagger.py`](/Users/wales/.codex/worktrees/5fd6/my-test-code_01/src/analysis_v14/tagger.py) 的 `real` 分类分支
- 趋势解释文本和结论卡文本，目前主要仍是规则聚合 + 模板化生成，不是运行时大模型总结
- 项目内两个 skill 只是方法说明，不是当前运行时系统模块

## 3. 模块总表
| 模块名称 | 模块作用 | 输入字段 | 输出字段 | 使用方式 | 当前实现位置 | 当前已知问题 | 最可能误差来源 |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 文件扫描与报告识别 | 扫描归档目录并识别周报/月报与时间 | 文件路径 | `report_id/report_type/year/month/week_of_month` | 规则 | `src/analysis_v14/loader.py` | 依赖目录命名质量 | 归档目录命名错误 |
| 正文抽取 | 从 `docx/doc/pdf/txt/md` 提取正文 | 文件路径 | 正文文本 / 解析失败原因 | 规则+工具链 | `src/analysis_v14/parser.py` | `doc/pdf` 稳定性依赖本地工具 | 工具缺失、正文抽取不完整 |
| 片段切分 | 把正文切成可分析片段 | 正文文本 | `segment_items` | 规则 | `src/analysis_v15/parser.py` | 句段不一定等于业务完整语义单元 | 切分粒度过碎或过粗 |
| 销售/区域/战区归一 | 把文件名、段落负责人、花名册挂接 | 文件路径、段落 owner、花名册 | `owner_registry/sales_roster` | 规则 | `src/analysis_v15/owner.py`、`src/analysis_v15/roster.py` | owner 噪声仍有残留 | 名册映射不足、文件命名噪声 |
| 花名册接入 | 建当前在岗销售主名单 | `xlsx` | `sales_roster.jsonl` | 规则 | `src/analysis_v15/roster.py` | 当前只支持花名册快照 | 花名册字段缺失或不完整 |
| AI 命中判断 | 判断片段是否属于 AI 信号 | `text/context` | `is_ai_hit` | 规则为主，模型可选 | `src/analysis_v14/tagger.py` | 宽泛表达和复杂隐式语义仍有误判 | 规则覆盖不足或 prompt 理解偏差 |
| 业务线识别 | 判断云诊室/云管家/混合/待判断 | `text/file_path` | `business_line` | 规则为主，模型可选 | `src/analysis_v14/tagger.py` | 低信号文本容易落入待判断 | 规则词命中不足 |
| 主体识别 | 判断销售自用/对外介绍/医生反馈/机会 | `text` | `actor_primary/actor_subtype` | 规则为主，模型可选 | `src/analysis_v14/tagger.py` | `ACTOR_OVERLAP` 仍多 | 语义复合、规则冲突 |
| AI 范围识别 | 判断 `product_ai/market_trend/competitor_ai/general_ai` | `text` | `ai_scope` | 规则为主，模型可选 | `src/analysis_v14/tagger.py` | 泛趋势与产品信号边界仍不稳 | 范围词规则不足 |
| 置信度与状态 | 生成 `certainty_level/confidence/decision_status` | 标签结果 + reason codes | 状态字段 | 规则 | `src/analysis_v14/tagger.py` | 当前分层仍偏硬编码 | 状态规则不符合业务容忍度 |
| `review_queue` 生成 | 把不稳定标签转成待复核项 | `tag_result` | `review_queue` | 规则 | `src/analysis_v14/review.py` | 当前 reason code 粒度仍不够丰富 | 标签层不稳定 |
| 复核写回消费 | 用人工结果覆盖原判断 | `review_decisions` + 原标签 | `evidence_facts/review_tasks` 更新 | 规则 | `src/analysis_v15/review_state.py`、`src/analysis_v15/normalize.py` | 现在是文件回写 MVP | 回写字段不全或状态设计偏轻 |
| 趋势聚合 | 生成月/周指标和快照 | `evidence_facts/sales_roster` | `trend_cube/dashboard_snapshot` | 规则 | `src/analysis_v15/metrics.py` | 当前指标解释仍偏浅 | 指标定义不够丰富 |
| 结论卡生成 | 把证据簇转成洞察树与结论卡 | `evidence_facts` | `insight_tree/insight_cards` | 规则+模板 | `src/analysis_v15/insights.py` | 结论“食之无味” | 证据簇质量、模板表达、动作标准不清 |
| 页面展示与下钻 | 把对象渲染为工作台和交互 | `normalized/*` | 页面与接口 | 规则/前端实现 | `src/analysis_v15/reporter.py`、`src/analysis_v15/webapp.py` | 受上游问题放大 | 数据对象不足、交互设计不足 |

## 4. 纯规则 / 纯代码步骤
当前完全不依赖大模型的步骤包括：

- 文件扫描与扩展名判断
- 周报/月报与时间识别
- 正文抽取
- 文本切分
- 花名册读取
- 销售/区域/战区归一
- `review_queue` 生成
- 复核结果写回文件
- 复核结果二次消费
- 趋势统计与时间窗口聚合
- 销售画像、区域构成聚合
- 结论树的证据分组与模板输出
- 页面渲染、本地服务、重建机制

这意味着：当前很多你看到的“浅”或“不准”，其实先要查规则和数据对象，不是先查模型。

## 5. 模型参与步骤
当前真正的运行时模型参与只有一个主入口：

- [`src/analysis_v14/tagger.py`](/Users/wales/.codex/worktrees/5fd6/my-test-code_01/src/analysis_v14/tagger.py) 的 `Tagger._classify_real()`

这个模型调用只负责片段级分类精修，目标字段是：
- `is_ai_hit`
- `business_line`
- `actor_primary`
- `actor_subtype`
- `ai_scope`
- `interaction_outcome`
- `certainty_level`
- `review_reason_code`
- `decision_status`
- `confidence`
- `reason`

当前没有运行时大模型参与的模块：
- 趋势解释
- 结论卡生成
- 销售画像总结
- 页面话术生成

所以：
- 如果标签不对，才需要考虑模型和 prompt
- 如果趋势和结论“食之无味”，优先不是查模型，而是查聚合、分组和模板

## 6. 提示词参与步骤
### 6.1 当前真实运行时提示词
当前在线上逻辑中实际使用的提示词只有一处：[`src/analysis_v14/tagger.py`](/Users/wales/.codex/worktrees/5fd6/my-test-code_01/src/analysis_v14/tagger.py)

#### system prompt
```text
你是销售周报AI专题分类器。请仅返回JSON对象，不要包含其他文本。字段: is_ai_hit,business_line,actor_primary,actor_subtype,ai_scope,interaction_outcome,certainty_level,review_reason_code,decision_status,confidence,reason。
```

#### user prompt 结构
```json
{
  "text": "...当前片段原文...",
  "context": {...报告上下文...},
  "rule_baseline": {...规则分类基线...},
  "constraints": {
    "decision_status_values": [...],
    "review_reason_code_values": [...],
    "do_not_force_classify": true
  }
}
```

### 6.2 这个提示词的目标
- 不是从零做分析
- 而是基于规则基线 `rule_baseline` 做片段级精修
- 重点是纠正规则处理不了的复杂语义和边界样本

### 6.3 输入与输出
输入：
- 当前文本片段
- 报告上下文
- 规则基线结果
- 合法状态枚举

输出：
- 一组固定 JSON 字段
- 用于覆盖或修正规则分类结果

### 6.4 失败时会发生什么
- 如果模型失败、接口未配置、返回非 JSON，系统不会静默成功
- 会回退到规则基线，并把：
  - `decision_status` 改为 `pending_human_review`
  - `certainty_level` 改为 `low`
  - `review_reason_code` 合并 `MODEL_CALL_FAILED`

### 6.5 当前哪些问题更可能是提示词问题
更可能是 prompt / 模型问题的场景：
- 明显复杂的复合语义
- 同时包含客户反馈、销售表达和机会判断的片段
- 规则已经给出相对合理基线，但模型反而改坏

当前不太像 prompt 问题的场景：
- 时间口径不清楚
- 趋势解释空
- 结论卡缺乏判断力
- 销售画像维度不够
- 复核工作流不完整

这些更偏 `BI 数据层 / 聚合层 / 结论模板层 / 复核系统层`。

## 7. skill 不是运行时 prompt
项目内有两个 skill：
- [`tools/skills/trend-insight-analysis/SKILL.md`](/Users/wales/.codex/worktrees/5fd6/my-test-code_01/tools/skills/trend-insight-analysis/SKILL.md)
- [`tools/skills/evidence-to-insight/SKILL.md`](/Users/wales/.codex/worktrees/5fd6/my-test-code_01/tools/skills/evidence-to-insight/SKILL.md)

它们当前是：
- 方法说明
- 面向 AI 代理的分析手册

它们不是：
- 当前运行时 prompt
- 当前系统正式调用的大模型模块

这点必须分清，否则很容易误判系统到底哪里在用模型。

## 8. 失败排查顺序
### 8.1 如果结论不对，先查什么
优先顺序：
1. 查证据质量
2. 查 owner 归一是否错人
3. 查证据簇分组逻辑
4. 查结论模板
5. 最后才考虑模型

原因：
- 当前结论卡不是模型总结产物，主要是 `evidence_facts -> insight_tree` 的规则链路

### 8.2 如果标签不对，先查什么
优先顺序：
1. 查正文抽取是否有缺字/错字
2. 查切分是否把语义拆坏
3. 查 `tagger` 规则
4. 如果是 `real` 模式，再查 prompt / 模型输出

### 8.3 如果趋势话术不对，先查什么
优先顺序：
1. 查时间口径
2. 查聚合指标
3. 查趋势解释模板
4. 最后才查页面文案

### 8.4 如果洞察空泛，先查什么
优先顺序：
1. 查证据簇质量
2. 查主题分组逻辑
3. 查结论模板
4. 不要先怪大模型

## 9. 排错决策树
```text
坏结果出现
├─ 是标签级问题？
│  ├─ 先查正文抽取和切分
│  ├─ 再查 tagger 规则
│  └─ 若为 real 模式，再查 prompt / 模型返回
├─ 是趋势级问题？
│  ├─ 先查时间口径和聚合指标
│  ├─ 再查 dashboard_snapshot / trend_cube
│  └─ 最后查展示文案
├─ 是结论卡问题？
│  ├─ 先查 evidence_facts 是否够稳
│  ├─ 再查证据簇分组
│  ├─ 再查结论模板
│  └─ 不要先怪模型
└─ 是复核闭环问题？
   ├─ 先查 review_tasks 是否正确生成
   ├─ 再查 review_decisions 是否正确写回
   └─ 再查 run_pipeline 是否正确消费复核结果
```

## 10. 当前最值得产品参与的提示词 / 规则点
当前最值得你参与的不是“把 prompt 写得更花”，而是这几类判断标准：

- 主体标签边界：哪些算销售对外介绍，哪些算医生反馈
- `pending_human_review` 的进入标准：什么值得进复核，什么可直接确认
- 结论卡的动作标准：什么叫“值得行动”
- 产品机会与背景观察的分界：哪些进入产品池，哪些只保留为观察

其中只有第一类会直接影响 prompt；后面三类更多会先影响规则和模板。

## 11. 本文结论
当前系统里：
- 模型是真实存在的，但只参与片段级分类精修
- 提示词是真实存在的，但只有分类 prompt 在运行时使用
- 你现在最不满意的“趋势浅、结论空、看板不好用”，主要不是 prompt 问题

当前最可能的根因顺序是：
- `BI 数据层`
- `分析方法层`
- `结论抽取层`
- `复核闭环层`

展示层更多是这些问题的外显。
