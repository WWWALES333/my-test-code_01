# v1.5 开发实现清单（执行版）

## 1. 开发原则
- 仅实现 `v1.5 PRD` 已锁定范围。
- 不改 `src/main.py` 主职责。
- 先保证业务可读性、追溯性和复核闭环，再优化样式与效率。

## 2. 数据层与中间层
- [ ] 新建 `src/analysis_v15/` 独立模块与运行入口。
- [ ] 建立 `data/output/insights/v1.5/normalized/`、`reports/`、`review/`、`web/` 目录约束。
- [ ] 实现销售归一对象 `owner_registry`。
- [ ] 实现 `report_facts`、`evidence_facts`、`sales_monthly_rollup`。
- [ ] 实现 `insight_cards`、`review_tasks`、`review_decisions`、`dashboard_snapshot`。
- [ ] 确保 `v1.3`、`v1.4` 产物不被覆盖。

## 3. Web 工作台
- [ ] 建立统一工作台入口。
- [ ] 完成 `overview` 模块。
- [ ] 完成 `trends` 模块。
- [ ] 完成 `sales` 模块。
- [ ] 完成 `insights` 模块。
- [ ] 完成 `evidence` 模块。

## 4. 复核闭环
- [ ] 生成标准化 `review_tasks`。
- [ ] 在工作台中展示待复核项、上下文和可编辑字段。
- [ ] 提交复核结果并写入 `review_decisions`。
- [ ] 下一轮分析优先消费人工复核结果。

## 5. skill 试点
- [ ] 落地项目内 `trend-insight-analysis` skill。
- [ ] 落地项目内 `evidence-to-insight` skill。
- [ ] 接通趋势中心对 trend skill 的消费链路。
- [ ] 接通结论中心对 insight skill 的消费链路。
- [ ] 验证 skill 失败不阻塞主链路。

## 6. 验收与封板准备
- [ ] 以 `2025-01` 至当前数据跑通 `v1.5` 工作台。
- [ ] 用 `v1.3` 冻结样本做回归验证。
- [ ] 完成趋势、销售、结论、复核四类人工验收。
- [ ] 执行安全检查脚本并通过。
- [ ] 回填 `change_log.md` 并准备封板。

## 7. 当前不进入实现的事项
- [ ] 数据库方案（留待后续版本评估）
- [ ] 多专题平台化
- [ ] 经营结果数据关联
- [ ] 复杂权限系统
