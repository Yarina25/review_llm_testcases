# AI 测试用例评审脚本

读取 Excel 中的测试用例（`.xlsx/.xlsm/.xls`），调用**火山方舟（豆包大模型）**逐条评审，
输出带**多维评分、评审结论、问题清单、修改建议**的评审报告 Excel（三张表）。

```
用例Excel ──► 解析/列映射 ──► 逐条调用方舟LLM评审 ──► 评审报告.xlsx
                                    ▲                       ├─ Sheet1 评审结果（逐条得分+结论）
                                    │                       ├─ Sheet2 问题明细（级别+维度+建议）
                        并发/重试/JSON解析                  └─ Sheet3 汇总统计（通过率+维度均值）
```

## 目录结构

| 文件 | 说明 |
|---|---|
| `review_llm_testcases.py` | 主脚本（核心交付物） |
| `mock_ark_server.py` | 本地 mock 方舟接口，用于不花钱演练全流程 |
| `requirements.txt` | 运行依赖 |
| `samples/样例用例.xlsx` | 含典型缺陷的样例用例（8 条） |
| `samples/样例用例_评审报告.xlsx` | 样例评审报告（mock 生成，供查看效果） |
| `samples/评审结果.json` | 样例结构化评审结果（便于平台/CI 对接） |
| `samples/gen_sample.py` | 样例用例生成器，可自行改数据重新生成 |

## 环境准备

```powershell
# 要求 Python 3.8+
pip install -r requirements.txt
```

依赖仅两项：`openpyxl`（读写 xlsx）、`xlrd`（读旧版 xls）。API 调用使用 Python 标准库 `urllib`，无额外依赖。

## 快速开始

### 1. 先做不花钱的演练（推荐）

```powershell
# 终端1：启动本地 mock 接口
python mock_ark_server.py --port 8000

# 终端2：用 mock 跑样例
python review_llm_testcases.py --input samples/样例用例.xlsx `
    --base-url http://127.0.0.1:8000 --api-key mock
```

### 2. 真实运行

```powershell
# 设置 API Key（获取：https://console.volcengine.com/ark/region:ark+cn-beijing/apikey）
$env:ARK_API_KEY = "你的方舟APIKey"

# 先试跑 2 条，确认列映射和效果
python review_llm_testcases.py --input 用例.xlsx --max-cases 2

# 全量评审
python review_llm_testcases.py --input 用例.xlsx
```

生成的报告默认在输入文件同目录：`<输入名>_评审报告.xlsx`。

## 参数说明

| 参数 | 说明 | 默认 |
|---|---|---|
| `--input` | 用例 Excel 路径（必填） | - |
| `--output` | 报告输出路径 | `<输入名>_评审报告.xlsx` |
| `--sheet` | 读取指定 sheet 名 | 第一个非空 sheet |
| `--model` | 方舟模型 ID | `doubao-seed-2-1-turbo-260628` |
| `--base-url` | 兼容接口地址 | `https://ark.cn-beijing.volces.com/api/v3` |
| `--api-key` | API Key（优先环境变量 `ARK_API_KEY`） | 空 |
| `--max-workers` | 并发数 | 4 |
| `--max-cases` | 只评审前 N 条（试跑用） | 0=全部 |
| `--timeout` | 单次接口超时（秒） | 120 |
| `--max-retries` | 失败重试次数（429/5xx/网络） | 3 |
| `--dry-run` | 只解析并打印列映射，不调接口 | 关 |
| `--json-out` | 额外输出结构化结果 JSON | 空 |
| `--col-steps` 等 | 表头列名覆盖（见下） | 自动识别 |

环境变量：`ARK_API_KEY`、`ARK_MODEL`、`ARK_BASE_URL`。

## 评审维度与权重

| 维度 | 权重 | 关注点 |
|---|---|---|
| 需求可追溯性 | 0.10 | 是否关联需求编号/描述 |
| 完整性 | 0.15 | 前置、步骤、数据、预期结果是否齐全 |
| 准确性 | 0.15 | 是否符合需求语义、无自相矛盾 |
| 可执行性 | 0.10 | 步骤无歧义、可操作、顺序合理 |
| 预期结果可验证性 | 0.15 | 预期结果明确、可判定通过/失败 |
| 边界与异常覆盖 | 0.10 | 边界值、非法输入、异常容错 |
| 独立性与去重 | 0.10 | 独立可执行、无重复/强依赖 |
| 优先级合理性 | 0.05 | 优先级与重要程度匹配 |
| 命名与可读性 | 0.05 | 用例名称准确表达意图 |
| 数据与安全考虑 | 0.05 | 数据合规、敏感数据、越权 |

- 综合得分 = 各维度加权平均（保留 1 位小数）
- 结论分档：`>= 8.5` 通过；`>= 6.5` 需修改；`< 6.5` 不通过
- 维度、权重、分档阈值均在脚本顶部 `DIMENSIONS / PASS_THRESHOLD / REVISE_THRESHOLD` 常量中，按团队规范直接改即可

## 输出说明

**Sheet1 评审结果**：原始用例列 + 综合得分、评审结论（颜色标注：绿=通过/黄=需修改/红=不通过）、问题数、总评、10 个维度得分。

**Sheet2 问题明细**：每条问题一行：用例编号、名称、严重级别（严重/一般/轻微，颜色标注）、涉及维度、问题描述、修改建议。

**Sheet3 汇总统计**：用例总数、通过/需修改/不通过数量、平均综合得分、问题总数、评审模型/时间、各维度平均分与最低分。

## 表头列名识别

脚本按关键字自动识别常见表头（支持中英文、大小写、空格/下划线差异），例如：

```
用例编号/id/编号 → 编号字段    测试步骤/操作步骤/step → 步骤字段
用例名称/标题    → 名称字段    预期结果/expected    → 预期字段
```

表头与内置别名不一致时，用参数显式指定（需 `--dry-run` 先看映射结果）：

```powershell
python review_llm_testcases.py --input 用例.xlsx --dry-run
python review_llm_testcases.py --input 用例.xlsx --col-steps 操作步骤 --col-expected 期望输出
```

支持覆盖字段：`--col-id / --col-title / --col-module / --col-priority / --col-pre /
--col-steps / --col-expected / --col-data / --col-req`。

## 常见问题

| 现象 | 处理 |
|---|---|
| `未提供 API Key` | 设置 `$env:ARK_API_KEY` 后重跑 |
| `API 鉴权失败(HTTP 401/403)` | Key 错误或模型未开通；到方舟控制台核实 Key 和模型 |
| `接口返回 HTTP 429` | 触发限流，脚本会自动退避重试；可降低 `--max-workers` |
| `未能定位「测试步骤」或「预期结果」列` | 用 `--dry-run` 看实际识别结果，再用 `--col-*` 指定 |
| `.xls` 旧格式 | 已支持（依赖 xlrd），无需转换 |
| 模型名怎么选 | 效果优先 `doubao-seed-2-1-pro-260915`；成本敏感 `doubao-seed-2-1-lite-260915`；默认 `doubao-seed-2-1-turbo-260628` 效果/成本均衡。具体以方舟控制台模型广场为准 |
| 大批量用例成本 | 按条计费，几百条建议分文件跑或用 `--max-cases` 分批；单条用例字段过长会被截断（默认 1500 字/字段，脚本内 `MAX_FIELD_LEN` 可调） |
