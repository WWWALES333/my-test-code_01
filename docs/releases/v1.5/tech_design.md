# AI 一线情报工作台 MVP 技术方案

**版本**：v1.5  
**状态**：正式基线  
**阶段**：开发前确认版

## 1. 版本说明
`v1.5` 技术方案只承接 `PRD.md` 已锁定的业务目标：把 `v1.4` 的离线分析报告升级为一个面向业务判断的统一工作台，并建立文件回写型复核闭环。

## 2. 当前现状与 `v1.4` 产物问题复盘
当前项目已具备：
- `src/main.py`：邮件拉取、下载、归档、审计、通知
- `src/analysis_v13/`：AI 专题离线 MVP
- `src/analysis_v14/`：真实归档试运行、CSV/Markdown/HTML 产物

基于真实 `v1.4` 产物复盘，现存问题明确如下：
- `dashboard_monthly.csv` 与 `dashboard_weekly.csv` 主要输出数量，不足以支撑业务解释。
- `evidence_trace.csv` 的 `owner_hint` 混有战区、区域、模板名、报告名，无法稳定映射到销售个人。
- `opportunity_backlog.csv` 混入模板文本、泛趋势表达和低价值证据，机会池噪声高。
- `review_worklist.csv` 已形成真实复核压力，但目前只有清单，没有正式任务流转。
- `AI专题看板.html` 是静态页面，不具备持续查看和回写能力。

## 3. 为什么 `v1.5` 不能只继续堆 HTML / CSV
继续增强 HTML/CSV 只能得到“更完整的展示物”，无法解决以下结构性问题：
- 缺少稳定销售主键
- 缺少中间层数据对象
- 缺少结论卡对象
- 缺少复核回写对象
- 缺少工作台状态与交互逻辑

因此，`v1.5` 必须先补数据分层和中间层，再做 Web 工作台。

## 4. 首期系统形态
- 运行形态：离线构建数据 + 本地 Web 工作台
- 存储形态：文件系统为主，不引入正式数据库
- 分析形态：规则 + 现有 AI 分析链路 + 首批 2 个 skill
- 复核形态：文件回写 MVP
- 页面形态：统一工作台，不再以单一离线 HTML 为主产物

## 5. 数据分层设计
`v1.5` 数据层固定分为四层：

### 5.1 原始层
沿用 `v1.4` 输出，不改动原始产物契约：
- `report_index.jsonl`
- `tag_result.jsonl`
- `evidence_span.jsonl`
- `review_queue.jsonl`

### 5.2 规范化层
把原始层清洗为稳定可消费的数据对象。

### 5.3 洞察层
趋势结论、结论卡、销售画像等业务对象。

### 5.4 展示与复核层
供 Web 工作台直接消费的数据快照和复核结果。

## 6. 规范化中间层设计
正式中间层目录固定为：
- `data/output/insights/v1.5/normalized/`

核心对象职责如下：

### 6.1 `owner_registry.jsonl`
- 统一销售主键
- 记录姓名、区域、战区、别名、归一来源
- 解决 `owner_hint` 无法稳定下钻的问题

### 6.2 `report_facts.jsonl`
- 报告级事实
- 字段包含：报告时间、报告类型、销售主键、区域、解析状态、原文路径

### 6.3 `evidence_facts.jsonl`
- 片段级事实
- 字段包含：业务线、主体、范围、状态、证据原文、所属销售、所属报告、复核状态

### 6.4 `sales_monthly_rollup.jsonl`
- 销售月度聚合
- 输出销售个人层面的趋势、渗透、反馈、机会、复核积压

### 6.5 `insight_cards.jsonl`
- 结论中心正式消费对象
- 每条记录对应一张结论卡

### 6.6 `review_tasks.jsonl`
- 正式待复核任务对象
- 由原始 `review_queue` 经过标准化得到

### 6.7 `review_decisions.jsonl`
- 人工复核结果
- 首版由文件回写方式生成

### 6.8 `dashboard_snapshot.json`
- 首页与趋势页的快照数据
- 避免页面直接读散落 CSV

## 7. Web 工作台模块设计
正式输出目录固定为：
- `data/output/insights/v1.5/web/`

统一工作台模块为：
- `overview`
- `trends`
- `sales`
- `insights`
- `review`
- `evidence`

模块职责：
- `overview`：汇总本期判断、机会、风险和关键指标
- `trends`：展示趋势和结构变化，支持下钻
- `sales`：销售个人画像与活跃分层
- `insights`：结论卡与机会卡
- `review`：复核任务查看与提交
- `evidence`：原文证据详情与追溯

## 8. 复核回写设计
首版复核闭环不引入数据库，采用文件回写 MVP。

正式目录固定为：
- `data/output/insights/v1.5/review/`

核心文件：
- `review_tasks.jsonl`
- `review_decisions.jsonl`
- `review_result.csv`

处理流程：
1. 分析链路输出待复核任务
2. Web 工作台读取 `review_tasks`
3. 人工修改关键字段并提交
4. 系统写入 `review_decisions`
5. 下一轮规范化和聚合优先消费人工结论

## 9. 证据追溯设计
所有页面层对象必须能追溯到：
- 结论卡 -> 证据片段
- 证据片段 -> 报告事实
- 报告事实 -> 原始文件路径

禁止出现：
- 只有结论没有证据
- 只有证据没有原文路径
- 只有销售聚合没有原始片段

## 10. skill 试点设计
首批 2 个 skill 固定为：
- `trend-insight-analysis`
- `evidence-to-insight`

放置位置固定为仓库内：
- `tools/skills/trend-insight-analysis/`
- `tools/skills/evidence-to-insight/`

边界固定：
- skill 只服务分析代理，不作为业务入口
- skill 不替代 `src/analysis_v15/` 的正式处理逻辑
- skill 不负责正式状态写回
- skill 只辅助：
  - 趋势解释
  - 结论卡生成

## 11. 代码模块落点
`v1.5` 继续独立于 `src/main.py`，不把新分析逻辑堆回下载主链路。

建议代码落点：
- `src/analysis_v15/loader.py`
- `src/analysis_v15/normalizer.py`
- `src/analysis_v15/aggregator.py`
- `src/analysis_v15/insight_builder.py`
- `src/analysis_v15/review_writer.py`
- `src/analysis_v15/run.py`
- `src/web_v15/` 或等价目录承接工作台页面

职责边界：
- `src/main.py`：仅承担 `v1.2` 主链路
- `src/analysis_v15/`：数据构建与结论构建
- `src/web_v15/`：展示与复核交互
- `tools/skills/`：分析代理说明文档

## 12. 风险与折中
- 销售归一规则短期内不可能完美，首版允许用映射表和规则归一先落地。
- 文件回写方案简单，但不适合多人并发，首版接受这个折中。
- 结论卡自动归纳存在噪声风险，必须用证据回链和人工复核兜底。
- skill 能提升分析稳定性，但不能替代正式产品逻辑。

## 13. 本期刻意不做
- 不改 `src/main.py`
- 不做正式数据库
- 不做多专题统一平台
- 不做经营结果数据关联
- 不做复杂权限
- 不把 skill 提炼到全局 `~/.codex/skills`
