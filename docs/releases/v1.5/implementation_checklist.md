# v1.5 开发实现清单（执行版）

## 1. 开发原则
- 仅实现 `v1.5 PRD` 已锁定范围。
- 不改 `src/main.py` 主职责。
- 先保证业务可读性、追溯性和复核闭环，再优化样式与效率。

## 2. 数据层与中间层
- [x] 新建 `src/analysis_v15/` 独立模块与运行入口。
- [x] 建立 `data/output/insights/v1.5/normalized/`、`reports/`、`review/`、`web/` 目录约束。
- [x] 实现销售归一对象 `owner_registry`。
- [x] 实现 `report_facts`、`evidence_facts`、`sales_monthly_rollup`。
- [x] 实现 `insight_cards`、`review_tasks`、`review_decisions`、`dashboard_snapshot`。
- [x] 确保 `v1.3`、`v1.4` 产物不被覆盖。

## 3. Web 工作台
- [x] 建立统一工作台入口。
- [x] 完成 `overview` 模块。
- [x] 完成 `trends` 模块。
- [x] 完成 `sales` 模块。
- [x] 完成 `insights` 模块。
- [x] 完成 `review` 模块。
- [x] 完成 `evidence` 模块。

## 4. 复核闭环
- [x] 生成标准化 `review_tasks`。
- [x] 在工作台中展示待复核项、上下文和可编辑字段。
- [x] 提交复核结果并写入 `review_decisions`。
- [x] 下一轮分析优先消费人工复核结果。
- [x] 学习字段沉淀为候选池，不自动修改规则或 Prompt。

## 5. skill 试点
- [x] 落地项目内 `trend-insight-analysis` skill 文档。
- [x] 落地项目内 `evidence-to-insight` skill 文档。
- [x] 明确 skill 只作为分析方法辅助，不作为当前运行时系统模块。
- [x] 验证 skill 不阻塞主链路。

## 6. 验收与封板准备
- [x] 使用既有 `v1.5_real_round2` 归一化产物重建工作台快照。
- [x] 用单元测试覆盖 `v1.3/v1.4/v1.5` 关键链路。
- [x] 执行安全检查脚本并通过。
- [x] 回填 `change_log.md` 并准备封板。
- [ ] 全量 2023-当前数据重建耗时较长，作为后续后台任务，不阻塞本次封板。

## 7. 当前不进入实现的事项
- [ ] 数据库方案（留待后续版本评估）
- [ ] 多专题平台化
- [ ] 经营结果数据关联
- [ ] 复杂权限系统
