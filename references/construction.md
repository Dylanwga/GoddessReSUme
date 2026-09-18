# 用原生组件构建完整简历

这是 [排版流程](workflow-layout.md) 的组件接口参考，适用于从经历材料新建整份 DOCX。局部修改用户当前文件仍按 [Word 精修](word-editing.md) 操作，不通过重新生成覆盖手工调整。样式参数以 [版式规范](style-spec.md) 为准。

`scripts/docx_components.py` 从随包脱敏样例复用对象结构，清空示例正文、图片关系与示例外链，再插入调用者提供的内容。每个内容块是一个原生文本框；机构条带把底色、行内图标和字段放在同一框内。模块依赖 `python-docx` 与 `lxml`，优先使用环境已有运行时。

## 调用前的输入

由内容与排版流程提供内容单元、本地图标、真实链接、选定主题色与初始几何位置。模块不联网、不识别机构或自动下载图标，也不测量文字或自动分页；字体可用性、图标获取和页面容量由流程处理。

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
| `add_institution(name, detail, degree, dates, ...)` | 一条完整机构抬头，字段使用规范中的机构样式。`logo_path` 为本地图片；`tab_stops` 为三处真实制表位，最后一处右对齐。没有中间字段可传空字符串。 |
| `add_contact(text, url=..., icon_path=..., ...)` | 一行单个联系项目。电话、邮箱等多个项目同排时使用 `add_block("contact", [...])`，每项分别插入图标和真实链接。 |
| `validate()` / `save(path)` | 检查对象、关系与页面边界；`save` 拒绝覆盖组件样例。不会测量文字高度、判断图标身份或检测视觉遮挡。 |

`role` 可选：`name`（姓名）、`section`（栏目及全宽下边框）、`institution`（机构）、`contact`（联系方式）、`project`（项目标题）、`body`（普通正文）、`bullet`（原生编号）、`citation`（题录）。对应字体与字号见 [参数表](style-spec.md#页面与文字参数)。`bullet` 显式指定 `level=0` 为一级实心、`level=1` 为二级空心；默认是二级。题录的粗体范围按论文层级调整，不能全部机械加粗。

位置参数：`page`、`x`、`y`、`width`、`height`。普通文字 run 支持 `text`、`bold`、`italic`、`underline`、`color`、`url`、`font_size`；图标片段支持 `icon_path` 和 `icon_height`。GitHub 标签和 URL 等字号差异按参数表分别传入 `font_size`。传入有效 PNG/JPEG；SVG 先按图标流程转成兼容图片。图片和链接写入当前文档的关系，图标不使用远程外链。

`font_size` 等覆盖参数服务于用户指定的样式和局部例外，不能用于绕开正文规范。同级元素保持一致；正文使用样例段落对齐，编号保留原生定义与悬挂缩进。

## 容量与分页

模块提供可靠组件，不自动完成内容布局。默认高度只是起点：按实际行数和规范行距、内边距及段间距预估高度，渲染后修正。不用 `\n` 在一个框里合并多条分项。机构字段较长时重新设制表位或增加条带高度，并检查右侧日期没有挤出或折行错位。

先安排姓名和实际联系行；无照片时收回头像位置，无 Scholar/GitHub 地址时省略对应行。栏目与至少一段正文同页。下一块超出页内可用区域时移到下一页；调用时用 `page=2` 等显式安排，不能将跨页段落塞进一个固定框。用户明确指定 A4 时可调整 `b.document.sections[0]` 的纸型及边距，再重新计算所有坐标和分页。

## 保存后的交接

`save()` 只保证已实现的结构与边界检查通过，不能证明字体实际匹配、内容齐全或文字可见。保存后按 [排版流程](workflow-layout.md) 渲染，再执行 [独立验收](acceptance.md)。任务文件不加入可复用 skill，参考资产维护遵循 [脱敏要求](style-spec.md#参考资产)。
