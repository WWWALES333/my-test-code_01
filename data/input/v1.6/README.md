# v1.6 输入目录说明

`data/input/v1.6/` 存放 v1.6 的业务定义、人工复核输入和回归样本，不存放私有密钥。

## 目录约定
- `business_question_taxonomy.md`：业务问题分类口径。
- `annotations/`：人工复核或黄金样本输入。
- `review/`：人工复核模板和轮次记录。
- `samples/`：冻结样本或小范围测试样本。

## 密钥要求
Minimax / OpenAI 兼容接口密钥不得放在本目录。正式运行只允许从环境变量或 macOS Keychain 读取。
