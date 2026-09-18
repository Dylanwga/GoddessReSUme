# GoddessReSUme

用于中文技术与科研简历的可复用 Codex skill，同时优化专业表达和 Word 格式排版。

## 能做什么

- **经历专业化**：组织业务背景、技术难点、个人职责和成果；根据已确认事实明确项目 Owner 的责任范围。
- **统一排版**：采用楷体与 Times New Roman、可编辑文本框、机构条带、分级标题、图标联系方式和右对齐日期。
- **从经历图标取色**：选择一段经历的图标作为全篇主题来源，派生标题深色与条带浅色。
- **局部精修与补漏**：尊重最新手工修改，支持恢复缺失内容、补充图标、调整联系方式，并核对文字与版面。

只改一句话时直接提供文案；完整优化或生成 Word 简历时执行内容与版式双重检查。专业化表达以已有事实为依据，不编造职责、指标或成果。

## 安装

将仓库克隆到 Codex 的个人 skills 目录。默认位置如下；如果已设置 `CODEX_HOME`，请使用对应目录。

```sh
git clone https://github.com/Dylanwga/GoddessReSUme.git ~/.codex/skills/goddessresume
```

已有同名目录时先检查内容，避免覆盖本地修改。展示名称为 **GoddessReSUme**；内部标识遵循 skill 的小写命名规范，安装后在新会话中使用 `$goddessresume`。

## 使用示例

```text
使用 $goddessresume 优化这份简历：专业化改写实习和科研经历，
保留事实与指标，按内置版式统一排版，并从最新实习经历的图标选择主题色。
```

```text
使用 $goddessresume，只修改指定经历，明确我作为项目 Owner 的职责，
保留其他部分的内容与位置。
```

```text
使用 $goddessresume，在 GitHub 链接前补一个图标，保持原链接和其他排版。
```

Word 文件处理需要可用的 DOCX 编辑、字体和渲染能力；若环境提供 documents 技能，可配合使用。附带的清点脚本仅依赖 Python 标准库，可检查 OOXML 文本、对象与链接，不能替代视觉检查。

## 版式预览

以下样例从空白文档重建，身份、机构、经历、论文及日期均为虚构，照片为原生占位框。蓝色仅演示图标取色规则；实际简历按选定经历的图标配色。

| 第 1 页 | 第 2 页 |
| --- | --- |
| ![脱敏版式示例，第 1 页](assets/layout-page-1.png) | ![脱敏版式示例，第 2 页](assets/layout-page-2.png) |

[下载可编辑的脱敏 Word 样例](assets/layout-reference.docx)

## 文件说明

| 文件 | 用途 |
| --- | --- |
| [SKILL.md](SKILL.md) | 工作范围、事实边界与内容/版式验收要求 |
| [references/writing.md](references/writing.md) | 专业化表达与职责、成果的写法 |
| [references/layout.md](references/layout.md) | 字体、配色、尺寸、间距和对齐规范 |
| [references/word-editing.md](references/word-editing.md) | Word 编辑、补漏与渲染检查 |
| [scripts/docx_inventory.py](scripts/docx_inventory.py) | 只读 DOCX 内容与对象清点工具 |
| [agents/openai.yaml](agents/openai.yaml) | Codex 展示名称和默认调用提示 |

```sh
python scripts/docx_inventory.py /path/to/resume.docx
python scripts/docx_inventory.py /path/to/before.docx --compare /path/to/after.docx
```
