# v1.4 输入目录说明

本目录用于 `v1.4` AI 专题试运行输入，不与 `v1.3` 冻结样本混用。

## 子目录约定
- `samples/`：试运行样本（支持 `docx/doc/pdf/txt/md`）
- `annotations/`：人工标注或复核基线
- `review/`：人工复核记录与批次补充说明

## 说明
- 建议按批次维护输入，如 `samples/batch_2026w13/`
- 真实运行入口见 `python -m src.analysis_v14.run`
