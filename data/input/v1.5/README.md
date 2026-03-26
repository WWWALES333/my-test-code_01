# v1.5 输入目录说明

本目录用于 `v1.5` AI 一线情报工作台的开发前输入基线，不与 `v1.3`、`v1.4` 输入目录混用。

## 1. 目录用途
- `annotation_guideline.md`：`v1.5` 标注与证据规范
- `sample_inventory.md`：`v1.5` 样本分层、冻结范围与补样原则
- `samples/`：需要人工挑选或冻结的样本链接 / 副本
- `annotations/`：人工标注或回归基线文件
- `review/`：人工复核模板、批次补充说明

## 2. 哪些是模板
以下文件默认是模板或基线说明：
- `annotation_guideline.md`
- `sample_inventory.md`
- `review/review_record_template.csv`

## 3. 哪些是人工填写结果
以下内容应由人工或后续流程补齐：
- `samples/` 下的冻结样本
- `annotations/` 下的人工标注文件
- `review/` 下按批次填写的复核记录

## 4. 维护原则
- 本目录只放 `v1.5` 输入基线，不放运行产物
- 真实运行产物统一输出到 `data/output/insights/v1.5/`
- 如样本路径或复核模板发生变化，必须同步更新 `docs/releases/v1.5/` 文档
