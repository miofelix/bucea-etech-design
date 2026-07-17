# bucea-etech-design

北京建筑大学智能科学与技术学院电气工程及其自动化专业第四学期《电子技术课程设计》项目仓库。

## 最终交付

仓库只维护一份最终报告：`report/电子技术课程设计报告.docx`。目录页使用大写罗马数字，正文从阿拉伯数字 `1` 开始。

任务书包含两道题目，实际电路与任务书示例并不完全一致。报告以设计图和实际完成情况为准：

| 设计图 | 对应题目 |
| --- | --- |
| `design/Design1.pdf` | 频率可调的多谐振荡器设计 |
| `design/Design2.pdf` | 简易十字路口红绿交通灯电路 |

数字时钟是额外完成的实物电路，没有对应的 `design/*.pdf`。

## 仓库结构

```text
.
├── docs/                         # 任务书、教学大纲、原始报告模板
├── design/                       # 两道题目的设计图 PDF
├── report-src/                   # 报告 Markdown 源文与原始图片
│   ├── images/
│   └── 电子技术课程设计报告.md
├── report/                       # 唯一最终 DOCX
├── scripts/
│   └── build_report.py           # 生成可审阅 DOCX
├── pyproject.toml / uv.lock      # Python 依赖与锁文件
├── README.md
└── AGENTS.md
```

`tmp/` 和 `output/` 仅用于本地中间产物，均被 Git 忽略。

## 生成审阅稿

生成脚本不会覆盖最终报告，而是写入 `output/电子技术课程设计报告.docx`：

```bash
uv run --locked python scripts/build_report.py
```

脚本需要 macOS 版 Microsoft Word 提供公式转换文件：

```text
/Applications/Microsoft Word.app/Contents/Resources/mathml2omml.xsl
```

生成稿包含自动目录字段。首次在 Word 中打开后需要更新字段，再进行逐页检查；只有人工复核通过后才可替换 `report/` 中的最终版。

## 交付检查

```bash
unzip -t report/*.docx
textutil -convert txt -stdout report/*.docx
```

同时检查报告与 `design/*.pdf`、任务书差异说明是否一致，并确认没有参考文献、占位内容、Office 锁文件或第二份报告。
