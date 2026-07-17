# AGENTS.md

本仓库 `bucea-etech-design` 用于北京建筑大学智能科学与技术学院电气工程及其自动化专业第四学期《电子技术课程设计》。

## 核心约定

- **最终只交付一份完整报告**，输出在 `report/`，不要按题目拆多套报告目录。
- 任务书含两道题目（`docs/电子技术课程设计任务书.doc`），设计图分别对应：
  - `design/Design1.pdf` → 频率可调的多谐振荡器设计
  - `design/Design2.pdf` → 简易十字路口红绿交通灯电路
- **实际电路与任务书并不完全一致。** 写报告时以 `design/*.pdf` 为准，说明差异；不得把任务书示例写成已实现事实。
- 报告格式依据 `docs/课设报告编写模板.doc`，正文维护源位于 `report-src/电子技术课程设计报告.md`。
- 最终报告保留目录：目录页使用大写罗马数字，正文使用阿拉伯数字并从 `1` 开始。
- 不要保留“参考文献”部分。
- 缺实验/仿真素材时只写“待补充”或占位，不虚构结果。
- 两题及额外完成内容统一写入**同一份**报告结构中，不拆分单题报告。

## 目录规范

| 路径 | 说明 |
| --- | --- |
| `docs/` | 任务书、大纲、报告模板 |
| `design/` | 设计 PDF（`Design1` / `Design2`） |
| `report/` | 唯一最终报告 |
| `report-src/` | 报告 Markdown 源文与图片 |
| `scripts/` | 报告生成脚本 |
| `tmp/`、`output/` | 本地临时产物，不提交 |

- 不保留 `.ms14`，不维护 `multisim/` 子目录。
- 不要把 `Design1` / `Design2` 的题目归属写反。

## Python

使用 `uv`，不要用系统 Python 直接跑脚本：

```bash
uv run --locked python scripts/build_report.py
```

依赖由 `pyproject.toml` 和 `uv.lock` 管理。脚本输出到仓库根下的 `output/`，不得直接覆盖 `report/` 中已审定的最终版。

公式转换依赖 macOS Microsoft Word 的 `mathml2omml.xsl`。生成稿首次在 Word 中打开时需要更新目录字段并进行逐页检查。

## DOCX 写作

- 标题：`1` / `1.1` / `1.1.1`
- 正文小四宋体风格、首行缩进
- 先文后图表；表题在上、图题在下
- 表格简洁黑白
- 无参考文献章节、无引用角标

## 交付前检查

- `unzip -t report/*.docx`
- `textutil -convert txt -stdout report/*.docx`（查目录、参考文献、错误占位）
- 目录页码为大写罗马数字，正文页码为阿拉伯数字并从 `1` 开始
- 正文与 `design/*.pdf`、任务书差异说明是否一致
- 无 LibreOffice 时说明无法做逐页渲染检查
