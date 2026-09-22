# 项目2｜AI 数据分析 Agent 自主交付评测

基于同一份 NYC311 服务请求数据，比较八个 AI 系统在两轮任务中的交付质量、执行可靠性、耗时与 API 目录价估算。系统自行选题、分析并提交 Excel 和分析说明；失败与人工接管记录单独列出。

## 项目索引

| 项目 | 内容与入口 |
| --- | --- |
| [项目1：NYC311 服务请求分析](https://github.com/Rainbow-bot-bot/nyc311-service-request-analysis) | 数据采集与服务请求分析，包含清洗、SQL、专题核验、Power BI 和项目报告。 |
| [项目2：AI 自主交付评测](https://github.com/Rainbow-bot-bot/nyc311-ai-agent-evaluation) | 当前仓库：比较 AI 完成数据分析任务的实际结果。 |

项目2使用项目1采集并保存的 NYC311 数据。原始来源为 [NYC Open Data 311 Service Requests](https://data.cityofnewyork.us/Social-Services/311-Service-Requests-from-2020-to-Present/erm2-nwe9)，快照日期为2026-09-07，共25份Parquet、7,525,498条记录、44个字段。文件清单和哈希见 [input/data_manifest.json](input/data_manifest.json)。

## 阅读入口

先看 [7页概览PDF](交付成果/项目2_研究报告_概览.pdf)，再查看 [完整报告PDF](交付成果/项目2_研究报告.pdf)、[Excel评测看板](交付成果/项目2_AI评测分析.xlsx) 或 [Markdown报告](交付成果/项目2_自主交付研究报告.md)。

[Word报告](交付成果/项目2_AI数据分析Agent自主交付评测报告.docx) 可编辑。[Notebook](交付成果/01_评测数据整理与验收.ipynb) 包含整理逻辑、计算和执行输出。

下载仓库后，可用浏览器打开 [开始阅读.html](交付成果/开始阅读.html) 和 [查证.html](交付成果/查证.html)。GitHub 文件页仅展示这两个页面的源码。

## 本次结果

正式配对为 GPT、Grok、DeepSeek、GLM、Qwen 五个系统。第二轮最终质量分依次为100、98、95、94、59；Qwen分项合计为90，按冻结规则封顶59。五系统最终质量均分由88.6变为89.2，增加0.6分。

Gemini第二轮为观察评分；Claude第一轮由Claude主导、Grok收尾；Kimi未形成完整成品。评分依据和成本口径见完整报告及 [最终评分裁定](./_实验系统/报告素材/最终评分裁定.json)。

结果限于本次NYC311任务，每种条件运行一次；模型、工具环境和提示条件共同影响结果。

## 仓库内容与运行条件

`交付成果/` 保存正式报告、工作簿、Notebook和查证页面；各AI目录保存原始提交及分析脚本；`资料/` 保存任务、评分和维护说明；根目录及 `_实验系统/` 保存相关处理与生成代码。

报告与Excel可直接阅读。原始Parquet、内部SQLite、原生日志及部分中间数据未上传，重跑分析需要这些输入和原环境依赖。具体文件范围与运行条件见 [公开仓库说明](资料/公开仓库说明.md)。
