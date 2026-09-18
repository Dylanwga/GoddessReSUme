# 用原生组件构建完整简历

适用于从经历材料新建整份 DOCX。局部修改用户当前文件仍按 [Word 精修](word-editing.md) 操作，不通过重新生成覆盖手工调整。

`scripts/docx_components.py` 从随包脱敏样例复用对象结构，清空示例正文、图片关系与示例外链，再插入调用者提供的内容。每个内容块是一个原生文本框；机构条带把底色、行内图标和字段放在同一框内。模块依赖 `python-docx` 与 `lxml`，优先使用环境已有运行时。

## 先准备内容与资产

1. 把经历整理成姓名、联系行、栏目、机构抬头、项目标题、正文段落、分项和论文题录。内容单元决定框数，样例框数和两页长度不决定新简历的篇幅。
2. 按 [图标获取](icons.md) 准备任务本地图片，确认各位置与来源，再选择经历图标的主题色。模块不联网、不会识别机构或自动下载图标，这部分由执行 skill 的模型使用可用搜索和下载工具完成。
3. 确认 KaiTi 与 Times New Roman 可用，查看两页样例。给每个块分配页号、坐标和高度，最后按渲染结果迭代；不把字体缩小当作容量控制。

## 基本调用

以下文字与路径都是演示，正式任务只传入已确认的经历、链接和本地图标。导入路径为当前 skill 的 `scripts` 目录，不依赖某台机器的固定位置。

```python
from docx_components import ResumeBuilder

# 色值须由当前经历图标推导；此处仅演示参数。
b = ResumeBuilder(theme="2F6FBA")
b.add_block("name", "虚构候选人", y=28, width=300, height=36)
b.add_block("contact", [
    {"icon_path": "/task/assets/phone.png", "icon_height": 11},
    {"text": "电话演示", "url": "tel:+10000000000"},
    {"text": " ｜ "},
    {"icon_path": "/task/assets/mail.png", "icon_height": 11},
    {"text": "person@example.com", "url": "mailto:person@example.com"},
], y=66, height=24)
b.add_block("section", "教育背景", y=108, height=28)
b.add_institution(
    "虚构大学甲", "示例专业", "示例学位", "20XX—20XX",
    logo_path="/task/assets/fictional-school.png",
    y=146, height=28,
)
b.add_block("project", "示例项目", y=186, height=24)
b.add_block("bullet", [
    {"text": "职责：", "bold": True},
    {"text": "这里传入已确认的责任范围。"},
], level=0, y=216, height=40)
b.add_block("bullet", [
    {"text": "技术实现：", "bold": True},
    {"text": "这里传入有依据的技术说明。"},
], level=1, y=260, height=44)
b.save("/task/outputs/resume.docx")
# 接下来渲染全部页面并检查，save 成功不代表版式通过。
```

## 接口与默认值

所有坐标、宽高及字号以 pt 表示，页号从 1 开始；默认采用样例纸型及页边距。

| 接口 | 用法 |
| --- | --- |
| `ResumeBuilder(theme=...)` | 主题须在准备图标后确定。省略时为中性灰，仅适用于图标确实不可取得且没有已有主题的回退。可通过 `east_asia_font`、`western_font` 指定用户授权的替代字体。 |
| `add_block(role, runs_or_text, ...)` | 每次增加一个独立框和一个段落；返回 `Block`，可访问其 `paragraph` 继续精修。传入字符串、一个 run 字典或 run 列表。 |
| `add_institution(name, detail, degree, dates, ...)` | 一条完整机构抬头，字段统一 11 pt 加粗。`logo_path` 为本地图片；`tab_stops` 为三处真实制表位，最后一处右对齐。没有中间字段可传空字符串。 |
| `add_contact(text, url=..., icon_path=..., ...)` | 一行单个联系项目。电话、邮箱等多个项目同排时使用 `add_block("contact", [...])`，每项分别插入图标和真实链接。 |
| `validate()` / `save(path)` | 检查对象、关系与页面边界；`save` 拒绝覆盖组件样例。不会测量文字高度、判断图标身份或检测视觉遮挡。 |

`role` 可选：`name`（20 pt）、`section`（14 pt、2 pt 全宽下边框）、`institution`（11 pt）、`contact`（11 pt）、`project`（10.5 pt 加粗）、`body`（10.5 pt 普通正文）、`bullet`（10.5 pt 原生编号）、`citation`（10.5 pt 题录）。`bullet` 显式指定 `level=0` 为一级实心、`level=1` 为二级空心；默认是二级。题录的粗体范围按论文层级调整，不能全部机械加粗。

位置参数：`page`、`x`、`y`、`width`、`height`。普通文字 run 支持 `text`、`bold`、`italic`、`underline`、`color`、`url`、`font_size`；图标片段支持 `icon_path` 和 `icon_height`。例如 GitHub 标签保持 11 pt、URL 可用 `font_size=10.5`。传入有效 PNG/JPEG；SVG 先按图标流程转成兼容图片。图片和链接写入当前文档的关系，图标不使用远程外链。

`font_size` 等覆盖参数服务于用户指定的样式和局部例外，不能用于绕开 10.5 pt 正文要求。同级元素保持一致；姓名与联系方式外的正文使用样例段落对齐，编号保留原生定义与悬挂缩进。

## 容量与分页

模块提供可靠组件，不自动完成内容布局。默认高度只是起点：按实际行数、1.2 倍行距、内边距及段间距预估高度，渲染后修正。不用 `\n` 在一个框里合并多条分项。机构字段较长时重新设制表位或增加条带高度，并检查右侧日期没有挤出或折行错位。

先安排姓名和实际联系行；无照片时收回头像位置，无 Scholar/GitHub 地址时省略对应行。栏目与至少一段正文同页。下一块超出页内可用区域时移到下一页；调用时用 `page=2` 等显式安排，不能将跨页段落塞进一个固定框。用户明确指定 A4 时可调整 `b.document.sections[0]` 的纸型及边距，再重新计算所有坐标和分页。

## 交付验证

1. 运行 `docx_inventory.py`，对照实际内容单元确认一段一框、无示例文字或遗漏；单独检查图片、链接、字号及原生编号，清点脚本不替代这些检查。
2. 核对学校/任职机构/联系字段资产清单：每个应有图标的位置都嵌入正确图片，链接目标来自用户材料。
3. 渲染所有页面，检查字体实际匹配、续行缩进、条带字段、全宽粗线、图标基线，以及页尾和框内文字是否完整。默认页面下方保留 28.35 pt 安全边界，不能仅凭未越过纸边就通过。
4. 调整高度、间距或分页后重新渲染受影响页面。确认所有内容可见且结构可编辑后交付；没有渲染能力时只能标为待验证草稿。

测试与预览使用从空白构建的虚构材料。用户简历、测试输出及其渲染件留在当前任务范围，不能加入可复用 skill。
