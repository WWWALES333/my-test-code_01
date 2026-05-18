# v1.6 实施检查清单

## 代码治理
- [x] 从 `release/v1.5` 创建 `release/v1.6`。
- [x] 明确 `src/main.py` 不承载 AI 分析逻辑。
- [x] 新增 `src/analysis_v16/`，避免复制整套 v1.5。
- [x] 封板前删除临时代码和重复定义。

## 文档基线
- [x] PRD
- [x] tech_design
- [x] test_plan
- [x] change_log
- [x] business_question_taxonomy

## Minimax 适配
- [x] `.com` 域名和 `MiniMax-M2.7` 作为默认配置。
- [x] Keychain / 环境变量读取。
- [x] `<think>` 清洗。
- [x] JSON 提取。
- [x] real 模式全链路跑通。
- [x] 边界样本批量 LLM 判定跑通。
- [x] 证据簇 LLM 洞察归纳跑通。

## 复核学习
- [x] 20 条一轮主动复核批次。
- [x] 自动错因归因。
- [x] 规则 / Prompt / 标签扩展候选。
- [x] 黄金样本集。

## 工作台
- [x] 总览
- [x] 趋势
- [x] 销售
- [x] 洞察
- [x] 复核
- [x] 证据
- [x] 首页补齐分析窗口、5 个业务问题回答、代表原文和待复核提示。
- [x] 页面主结论不暴露 `external_pitch`、`unknown` 等技术枚举。

## 验收
- [x] 单元测试通过。
- [x] 核心脚本 smoke test 通过。
- [x] 敏感信息扫描通过。
- [x] README 更新。
- [ ] 封板提交。
