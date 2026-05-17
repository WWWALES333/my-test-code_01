# v1.6 代码审核记录

## 当前结论
`v1.5` 已形成可复用基础层，`v1.6` 不应复制整套实现，而应复用 v1.5 的抽取、归一、销售画像、趋势基础数据，再叠加业务问题、复核学习和 UI/UX。

## 保留边界
- `src/main.py`：保留下载、归档、审计、通知职责。
- `src/analysis_v13/`、`src/analysis_v14/`、`src/analysis_v15/`：作为历史版本和复用基础保留。
- `src/analysis_v16/`：只放 v1.6 新增能力。

## 已做收敛
- Minimax 适配独立为 `model_adapter.py`。
- 业务问题枚举独立为 `schema.py`。
- 业务问题识别独立为 `business_questions.py`。
- 复核学习闭环独立为 `review_learning.py`。
- 页面和报告独立为 `reporter.py`。

## 后续封板前检查
- 检查是否存在重复 Prompt、重复枚举。
- 检查是否有临时调试输出。
- 检查是否误提交私有数据和运行产物。
- 跑 `tests/check_no_secrets.py`。
