# v1.3 变更记录

## 1. 版本目标
围绕销售周报和月报，完成 AI 专题分析第一版 MVP。

## 2. 本版新增
- 新增离线分析模块 `src/analysis_v13/`，支持样本读取、文本解析、分段、打标、复核队列与摘要生成。
- 新增统一离线入口：`python3 -m src.analysis_v13.run`。
- 新增 `v1.3` 标注体系升级：
  - 双层主体标签：`ai_actor/actor_primary + actor_subtype`
  - 范围维度：`ai_scope`（`product_ai | market_trend | competitor_ai | general_ai`）
  - 互动结果维度：`interaction_outcome`
  - 置信度维度：`certainty_level`
  - 复核原因码：`review_reason_code`
- 新增双视图摘要口径：
  - 主业务 AI 结论（`product_ai` 为主）
  - 市场雷达（`market_trend + competitor_ai`）
- 完成 24 份冻结样本跑通，形成 `labels_v1.csv` 封版标注基线。

## 3. 本版未做
- 未建设正式 Web 页面。
- 未引入数据库正式实现。
- 未改动 `v1.2` 下载/归档主链路（`src/main.py` 继续保持现有职责）。
- 未做全历史样本回填，仅完成冻结样本集闭环。

## 4. 已知限制
- 当前规则仍以启发式为主，跨语境长文本在部分场景下仍需人工复核。
- `ai_scope` 与 `business_line` 的边界在极宽泛表达下仍可能依赖业务判读。
- PDF/DOC 在 `v1.3` 首版中不作为主解析对象。

## 5. 后续迭代方向
- 基于 `labels_v1.csv` 增量优化规则与模型策略，降低误判并提升可解释性。
- 增加针对历史样本的批量回放与稳定性对比报告。
- 评估将离线产物接入轻量可视化页面，提升业务查看效率。
