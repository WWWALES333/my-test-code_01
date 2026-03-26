# v1.4 开发实现清单（执行版）

## 1. 开发原则
- 仅实现 `v1.4 PRD` 已锁定范围。
- 不改 `src/main.py` 主职责。
- 先保证可追溯与准确率，再优化效率。

## 2. 阶段任务拆解
### Phase A：基础框架与目录
- [x] 新建 `src/analysis_v14/` 独立模块与运行入口。
- [x] 建立 `data/input/v1.4/` 与 `data/output/insights/v1.4/` 目录约束。
- [x] 确保 `v1.3` 产物不被覆盖。

### Phase B：解析与切分升级
- [x] 落地 `docx/pdf/doc` 正文解析链路。
- [x] 对解析失败输出标准原因码并进入 `review_queue`。
- [x] 稳定片段切分策略，保留回归可比性。

### Phase C：判定与模型双模
- [x] 复用既有规则初筛与标签口径。
- [x] 接入 `real` 模式（OpenAI 兼容接口）并保留 `mock` 对照。
- [x] 统一输出状态：`confirmed | uncertain | pending_human_review`。

### Phase D：产物与复核闭环
- [x] 生成标准产物：`report_index/tag_result/evidence_span/review_queue`。
- [x] 输出 `review_queue.csv` 与 `AI专题摘要.md`。
- [x] 补齐复核记录模板（`data/input/v1.4/review/review_record_template.csv`）。

### Phase E：验收与封板准备
- [ ] 执行冻结样本回归 + 2026 试运行。
- [ ] 统计严格门槛：主标签一致率 `>=90%`、追溯率 `=100%`。
- [x] 回填 `change_log.md`（首轮实现进展）。

## 3. 验收检查项（完成定义）
- 标准产物完整输出且路径规范正确。
- 结论可追溯到证据，证据可追溯到原文。
- 不确定片段全部进入正式复核链路。
- `v1.2` 主链路无行为回归。
- 安全检查脚本通过。

## 4. 风险关注点
- `doc/pdf` 解析质量波动导致复核量上升。
- `real` 模式输出波动导致一致率下降。
- 试运行样本覆盖不足导致验收结果失真。

## 5. 当前执行顺序
1. 先完成 `Phase A + Phase B`。  
2. 再完成 `Phase C + Phase D`。  
3. 最后执行 `Phase E` 并决定是否封板。  
