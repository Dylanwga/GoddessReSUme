# 从空白文档构建原生简历组件

本页是 [排版流程](workflow-layout.md) 的组件接口参考。局部修改用户当前文件按 [Word 精修](word-editing.md) 操作，不能重建后覆盖手工调整。字体、颜色关系和层级由 [版式规范](style-spec.md) 统一规定。

新建简历默认目标为**内容充实、紧凑且恰好一页**。首轮用安全高度防裁字，渲染后必须按实际文字底边收紧和重排；不能把保守估高直接当成最终段间距。纸型、字号和事实不能为填满页面而任意修改。资料确实不足时说明缺哪些内容，不虚构经历或用拉大空白制造“满一页”。

`scripts/docx_components.py` 使用 `python-docx` 与 `lxml` **从空白 Document 创建文档**。不再读取、克隆或要求一个已有 DOCX 模板。脱敏样例是组件生成的可视参考，不是代码运行必需的输入，也不提供候选人事实。

## 组件与对象关系

每个逻辑段落只有一个原生文本框。装饰、图标和文字是否同框，按用途区分：

| 组件 | 原生对象结构 |
| --- | --- |
| 栏目标题 | 一个透明标题框 + 一个正文列等宽的实色粗横线对象，组成局部原生组；用于“教育背景”“项目经历”等栏目。 |
| 机构抬头 | 一个浅色底框 + 一个独立图片对象 + 一个透明文字框，组成局部原生组。图片不存在时保留两层；缺图原因须在图标流程中说明。 |
| 项目小标题及仓库 | 一个文字框；左侧项目小标题，右侧行内 GitHub 图标和真实链接，以右对齐制表位放在同一行。此层级不加栏目粗线。 |
| 联系方式 | 每个实际设计行一个框，图标、联系文字、真实链接和分隔符同行。 |
| 技术栈 | 联系方式下方的独立文字框，可选代码图标；只用有依据的项目相关技术，不自动附加“精通”等熟练度声明。 |
| 背景、职责、技术说明、成果、教育介绍 | 每个逻辑段落一个文字框。每条原生项目符号独立成框，段落自然换行不拆框。 |
| 照片占位 | 原生灰色矩形文字框，仅用于脱敏样例或用户明确要求的可填写模板。普通简历无照片时收回槽位。 |

机构的文字框不带填色，`bodyPr anchor="ctr"`，上下内边距对称；独立图片按条带高度居中。真实图标的宽高比决定文字避让距离，图片不参与文字行高。条带作为整体可移动，组内的底色、图标与文字仍能分别选择编辑。

栏目粗线必须在渲染中真实可见，不能用 `w:u`、下划线字符或大量带下划线空格模拟。组件直接创建原生矩形粗线，不依赖浮动文本框中的段落下边框支持情况。

## 调用前准备

流程应传入：已确认的内容单元、本地图标、真实目标链接、选定主题色、页面尺寸，以及初始排版位置。组件不联网、不搜索图标、不推断项目职责、不确认个人贡献。图标按 [图标流程](icons.md) 获取；用户提供的真实标识可直接使用。SVG 先转为兼容 PNG，图片嵌入文档，不用远程图片链接。

每段教育经历默认包含机构条带和**两条独立介绍框**：学术信息（硕博的导师/研究方向、本科的毕业设计/指导教师）及在校活动/荣誉。有资料就填写；缺资料保留字段标签与可编辑空白，不写虚构事实，也不重复增加一套空白框。仅用户明确要求不留空、删除字段，或局部精修范围不包含教育时省略。新建使用 `add_education()`，避免只建机构条而漏掉介绍位置。

项目正文使用 `add_project_point()`：背景与职责为一级实心分点，技术说明为二级空心分点；职责正文加粗。已确认的主要成果可用一级实心分点，没有成果资料就不调用成果组件。每个背景说明场景、问题或目标，不以重复职责充数。

## 完整示例

以下文字和图标均为明确虚构示例。示例路径由当前任务提供，不指向个人原件。`scripts` 加入当前 Python 导入路径后：

```python
from docx_components import ResumeBuilder

b = ResumeBuilder(theme="2F6FBA")  # 演示色；正式任务按图标流程选色。
b.add_block("name", "虚构候选人", y=28, width=320)
contact = b.add_block("contact", [
    {"icon_path": "/task/assets/phone.png", "icon_height": 11},
    {"text": " 000-0000-0000", "url": "tel:+10000000000", "color": "262626", "underline": False},
    {"text": " ｜ "},
    {"icon_path": "/task/assets/mail.png", "icon_height": 11},
    {"text": " person@example.com", "url": "mailto:person@example.com"},
], y=64)
stack = b.add_tech_stack(
    ["示例语言", "示例框架", "示例数据库"],
    icon_path="/task/assets/code.png", y=contact.next_y(2),
)
section = b.add_section("教育背景", y=stack.next_y(5))
education = b.add_education(
    "虚构大学甲", "示例专业", "硕士", "20XX—20XX",
    logo_path="/task/assets/fictional-school.png", y=section.next_y(4),
    academic_info="导师：示例教授；研究方向：虚构技术专题。",
    activities="",  # 默认生成“在校活动 / 荣誉：”及可编辑空白。
)

# 默认一页目标；先安排初稿，再根据真实渲染收紧和审查内容密度。
section = b.add_section("项目经历", y=education.next_y(10))
band = b.add_institution(
    "虚构实验室", "示例方向", "", "项目 Owner",
    logo_path="/task/assets/fictional-lab.png", y=section.next_y(4),
)
project = b.add_project_header(
    "示例项目：调查流程自动化",
    repo_url="https://example.com/repository",
    repo_label="example.com/repository",
    icon_path="/task/assets/github.png", y=band.next_y(3),
)
background = b.add_project_point(
    "background", "描述有依据的使用场景、现有问题和项目目标。",
    y=project.next_y(.5),
)
responsibility = b.add_project_point(
    "responsibility", "明确已确认的 Owner 角色及责任范围。",
    y=background.next_y(.5),
)
technical = b.add_project_point(
    "technical", "传入有依据的技术方案及实现说明。",
    y=responsibility.next_y(.5),
)

report = b.validate()  # 报告结构、页边界和可能过长的标题链接。
b.save("/task/outputs/resume.docx")
# 接下来渲染、读取实际文本 bbox，再执行下文的紧凑定高与单页检查。
```

示例只演示组件顺序，不是足量的完整候选人素材。同项目框外间距默认 0–1 pt，项目之间 6–8 pt，栏目之间 8–12 pt；教育单行介绍框通常约 20 pt。按实际渲染保持同级一致。溢出时先收紧多余框高和重复文字，再调整有依据的内容组织；不得自动把一页目标变成多页。

## 接口契约

所有坐标、字号、间距和宽高以 **pt** 表示。页码从 1 开始。未传 `x`/`y` 时使用当前节的左/上边距；未传 `width` 时使用 `页宽 − 右边距 − x`。可修改 `b.document.sections[0]`，但修改纸型后必须重新计算已有对象位置。

### 初始化与页面

```python
ResumeBuilder(
    theme=None, east_asia_font="KaiTi", western_font="Times New Roman",
    page_width=595.3, page_height=864.55,
    margin_left=34, margin_right=34, margin_top=28.35, margin_bottom=28.35,
)
```

默认纸型来自版式规范，不是 A4。用户要 A4 时显式传相应页面尺寸。`theme=None` 使用中性灰，仅适用于图标确实不可取得、用户也无既有主题的回退。更换字体需有用户授权并验证可用性。

- `b.content_width`：当前节的正文列宽。
- `b.safe_bottom`：当前页高减下边距。
- `b.next_position(block, gap=.5, required_height=0)`：返回 `(page, y)` 的位置建议，不移动现有对象，也不拆分段落。它支持用户明确要求的多页布局；一页任务收到下一页建议时应先修订布局与内容，不能直接套用。栏目与首个内容块须将组合高度传给 `required_height`。
- `GAPS` 提供 `.5 / 7 / 10 pt` 的同项目、跨项目、跨栏目起点；是紧凑关系的默认值，不是忽略实际显示效果的硬阈值。

### 内容与几何

`add_block(role, runs_or_text, *, page=1, x=None, y=None, width=None, height=None, font_size=None, color=None, level=1, icon_path=None, icon_height=11, line_spacing=1.2, line_height=None, padding_y=2)`

- 常用 `role`：`name`、`body`、`bullet`、`citation`、`contact`、`project`、`tech_stack`。`section` 转交栏目组件；`institution` 简写入口**只接收无制表符的机构名称字符串**，不支持 run 列表或把多个字段拼在一起。完整机构字段、制表位和图标参数使用 `add_institution()`。
- 文本参数可为字符串、run 字典或 run 列表。文本 run 支持 `text`、`bold`、`italic`、`underline`、`color`、`url`、`font_size`。图标 run 支持 `icon_path`、`icon_height`。
- 每个调用只创建一个逻辑段落；不允许传入 `\n`/`\r` 合并条目。文字自然换行由 Word 排版。
- `line_spacing` 是相对倍数；`line_height` 传 pt 时改为固定行高。联系方式、项目小标题和技术栈默认固定 18 pt 行高。正文默认 1.2 倍，**不是固定 14.4 pt**。
- `height=None` 按文本宽度、字号、行高和内边距估算。中文自然行高预留额外字体度量余量，避免把名义字号乘以倍数误当成实际行框高度；实际段落仍保持指定的 1.2 倍行距。估算不读取 Word 的实际排版结果，不保证刚好装下；显式 `height` 不会被覆盖。
- `bullet` 使用原生编号；`level=0` 为一级实心，`level=1` 为二级空心。符号字体不改变正文中西文字体。
- `citation` 默认斜体，具体加粗与作者强调由 run 传入，不能把所有论文一律加粗。

除教育组合外，创建方法返回 `Block`，提供 `.paragraph`、`.anchor`、`.page`、`.x`、`.y`、`.width`、`.height`、`.bottom`、`.next_y(gap=.5)`。`add_education()` 返回 `EducationBlock`，可读 `.band`、`.details`、`.field_names`、`.page`、`.bottom`、`.next_y(gap=.5)`；它只是逻辑关联，两个介绍框仍是页面上的独立对象，不放进机构组。`metadata` 保存估算与组件信息，不属于候选人事实。

### 专用组件

| 接口 | 参数与行为 |
| --- | --- |
| `add_section(text, ..., height=26, font_size=14, rule_thickness=1.5, rule_y=18.5, padding_y=2)` | 原生标题/粗线组。`rule_y` 相对组顶部，`rule_thickness` 为粗线高度。与正文列等宽；仅用于栏目层级。 |
| `add_institution(name, detail="", degree="", dates="", ..., logo_path=None, logo_height=18, logo_gap=4, tab_stops=None, height=26, font_size=11, text_baseline=0, line_height=16, tint_alpha=10000)` | 独立背景/图标/文字组。三处真实制表位偏移相对于整个条带，最后一处右对齐；默认按当前宽度计算。无中间字段传空字符串。图片等比缩放、垂直居中，文字框透明居中。`text_baseline` 仅用于渲染后的光学补偿，可正负微调，默认 0。 |
| `add_education(name, detail="", degree="", dates="", academic_info="", activities="", academic_label=None, activities_label="在校活动 / 荣誉：", reserve_missing=True, band_gap=1, detail_gap=.5, detail_line_height=16, **institution_placement)` | 机构条后默认增加学术信息、在校活动两个独立框。`academic_label` 默认按本科/学士选择“毕业设计 / 指导教师：”，其他选择“导师 / 研究方向：”，允许覆盖。缺资料显示标签和空白；`reserve_missing=False` 只省略缺资料字段，已有资料仍保留。机构图片、坐标、宽高等由 `institution_placement` 传入。 |
| `add_project_header(title, ..., repo_url=None, repo_label=None, icon_path=None, icon_height=11, height=24, font_size=10.5, link_font_size=10.5, min_gap=10)` | 小标题与仓库同行。`repo_label` 只控制显示；默认从 URL 省略协议，真实链接目标不变。链接不继承标题加粗，图标在链接前。宽度估算超限会报告警告，需缩短显示文字并渲染确认。 |
| `add_project_point(kind, text, label=None, **placement)` | `kind` 必须为 `background`、`responsibility`、`technical` 或 `result`。背景/职责/主要成果为原生一级实心，技术为二级空心；职责全文加粗。标签可改，层级由语义决定。只写传入内容，不自动生成结果。 |
| `add_contact(text, url=None, icon_path=None, **placement)` | 单项联系行。多个联系项目同行用 `add_block("contact", [...])`，不要用重复调用制造重叠。 |
| `add_tech_stack(items, icon_path=None, label="技术栈：", separator=" · ", **placement)` | `items` 为字符串或字符串列表。技术项只来自提供的事实/项目证据；可手动使用 `｜` 分组。 |
| `add_photo_placeholder(label="照片占位", ..., width=76.17, height=106.64)` | 原生灰色占位框，只用于明确请求的模板或完全虚构样例。 |
| `estimate_height(role, runs_or_text, width=None, font_size=None, line_spacing=1.2, line_height=None, padding_y=2, level=1)` | 返回初步框高，便于安排下一块。调用者仍需检查渲染后的实际文字。 |

上述 1.5 pt、18.5 pt、26 pt 等是**可调整的组件起点**，不是所有字体、纸型或用户参考的硬性验收阈值。不得把某份样例的框数、页数、0.5 pt 基线补偿或所有段落框高固定下来。

## 内容驱动的高度与间距

布局按“当前文字需要的框高 + 同级间距”依次计算，不能所有正文无条件分配同样的两三行大框。一行内容与两行内容的框高应不同；栏目间距大于同一项目内的段间距，同级保持一致。首轮保守框高用来防裁字，**不是交付时可直接保留的最终高度**。

高度估算有字体度量误差，尤其中文、西文、粗体混排和长 URL。渲染后同时检查**实际文字间距**和框位置：若框的空白过大，应收紧框高；如果文字已经贴底，则应增高而非压缩。文字、图标、底色分别可编辑并不代表视觉已对齐。

使用批准的渲染器输出 PDF，测得每段末行墨迹底边相对文本框顶部的位置后：

```python
# observation 来自与该次渲染配套的 manifest 和 layout_metrics 输出。
# 必须减“观测帧当时的 y”；不能减已被 move_block 改过的 background.y。
measured_background_ink_bottom = observation["ink_bounds"]["yMax"] - observation["frame"]["y"]
b.set_body_layout(background, rendered_ink_bottom=measured_background_ink_bottom,
                  bottom_padding=1, safety_margin=.5)
b.move_block(responsibility, y=background.next_y(.5))
# 依次收紧并重排后续块，再渲染全部页面复核；教育详情也可逐框处理。
```

`set_body_layout(block, rendered_ink_bottom=None, rendered_lines=None, line_pitch=17.75, bottom_padding=1, safety_margin=.5)` 只接受正文、分点或题录框。优先用实际墨迹底边，设置 `高度 = ink_bottom + bottom_padding + safety_margin`，保留顶部内边距和实际段落行距。没有 bbox 而只有真实行数时，可用 `rendered_lines` 和测得的 `line_pitch` 得到保守回退；此回退不能代替对最终空白的测量。调用后必须重新渲染，结构计算本身不证明未裁字。

`move_block(block, y=..., page=None, x=None)` 只移动指定组件；栏目/机构小组整体移动，教育介绍框彼此不锁定。不要通过移动一条教育介绍就假设其余关联内容自动随动。

### 向渲染测量脚本导出清单

`layout_manifest()` 返回可直接写为 JSON 的 `{"blocks": [...]}`。每个对象包含 `id`（外层 `wp:docPr` 名称）、`page`、`x`、`y`、`width`、`height`、`role`、`text`；尺寸单位均为 pt。文字从实际 OOXML 聚合，包括超链接内文字及真正的 run 制表符，不依赖可能漏掉超链接的 `paragraph.text`，也不把制表位定义误计为文字。

教育介绍导出为 `education_detail`；有意预留的空白字段导出为 `education_placeholder` 且 `intentional_blank=true`，帮助验收区分预留内容与意外漏字。已有资料的教育介绍为 `intentional_blank=false`。正文 `metadata` 中有 `group_id`、`semantic_kind` 时一并导出，供测量脚本识别所属项目和分点层级。其余组件保留原 `role`。

这些是导出角色别名，**不会改变教育详情 Block 的内部 `body` 类型**；可将 `education.details[i]` 直接传给 `set_body_layout()`。收紧或移动后重新调用 `layout_manifest()`，清单即反映最新几何位置。清单用于匹配实际渲染，不能据此推断真实文字底边或宣布视觉通过。

### 两轮渲染的具体步骤

以下路径为任务目录示例。`TASK_PYTHON`、`DOCX_RENDERER`、`PDFTOTEXT` 应先解析为批准的运行时、`render_docx.py` 和 Poppler 工具路径；字体配置按环境准备。不要默认调用用户桌面上的 LibreOffice。整个构建过程中保留同一个 `b` 和它的 Block 对象，或用确定性的构建脚本恢复它们；每次 manifest 必须对应同一次保存的 DOCX。

**第一轮：保存安全初稿和清单。**

```python
import json
from pathlib import Path

task = Path("/task")
b.save(task / "resume-pass1.docx")
(task / "manifest-pass1.json").write_text(
    json.dumps(b.layout_manifest(), ensure_ascii=False, indent=2), encoding="utf-8",
)
```

执行第一轮实际渲染及测量：

```sh
"$TASK_PYTHON" "$DOCX_RENDERER" /task/resume-pass1.docx --output_dir /task/render-pass1 --emit_pdf
"$PDFTOTEXT" -bbox-layout /task/render-pass1/resume-pass1.pdf /task/render-pass1/bbox.xhtml
"$TASK_PYTHON" /path/to/GoddessReSUme/scripts/layout_metrics.py /task/render-pass1/bbox.xhtml --manifest /task/manifest-pass1.json --output /task/metrics-pass1.json
```

**第二轮：读取第一轮观测，收紧并重排，再生成新的清单。**

```python
from docx.oxml.ns import qn

metrics = json.loads((task / "metrics-pass1.json").read_text(encoding="utf-8"))
observations = {row["id"]: row for row in metrics["blocks"]}
for block in b.blocks:
    if block.role not in {"body", "bullet", "citation"}:
        continue
    observation = observations[block.anchor.find(qn("wp:docPr")).get("name")]
    if observation["match_status"] != "matched" or not observation["ink_bounds"]:
        continue  # 未匹配项必须另行检查，不能把 unknown 当通过。
    relative_bottom = observation["ink_bounds"]["yMax"] - observation["frame"]["y"]
    b.set_body_layout(block, rendered_ink_bottom=relative_bottom)

# 示例：按照第一轮相同顺序显式重排；实际任务须覆盖所有后续组件。
b.move_block(responsibility, y=background.next_y(.5))
b.move_block(technical, y=responsibility.next_y(.5))
b.save(task / "resume-pass2.docx")
(task / "manifest-pass2.json").write_text(
    json.dumps(b.layout_manifest(), ensure_ascii=False, indent=2), encoding="utf-8",
)
```

重新渲染修改后的文件，不能复用第一轮 PDF 或清单：

```sh
"$TASK_PYTHON" "$DOCX_RENDERER" /task/resume-pass2.docx --output_dir /task/render-pass2 --emit_pdf
"$PDFTOTEXT" -bbox-layout /task/render-pass2/resume-pass2.pdf /task/render-pass2/bbox.xhtml
"$TASK_PYTHON" /path/to/GoddessReSUme/scripts/layout_metrics.py /task/render-pass2/bbox.xhtml --manifest /task/manifest-pass2.json --output /task/metrics-pass2.json
```

最后读取**第二轮**实测页数和正文底边，并逐页看图核对裁切、分点、间距、图标和字体：

```python
metrics = json.loads((task / "metrics-pass2.json").read_text(encoding="utf-8"))
last_page = metrics["pages"][-1]
measured_bottom = last_page["substantive_last_line_bottom"] if metrics["body_coverage"] == "matched" else None
density = b.one_page_report(
    rendered_page_count=metrics["page_count"],
    rendered_content_bottom=measured_bottom,
)
# unknown、overflow、page_count_mismatch 或其他不达标状态都不能作为交付通过。
```

这两轮是最低闭环，不是固定“两次即可”。文字或框宽再变化、匹配仍不确定、出现裁切或间距不一致时，需要修正并重新测量。

## 单页密度报告

`one_page_report(target_pages=1, bottom_whitespace_target=(12, 28), rendered_content_bottom=None, rendered_page_count=None)` 报告目标页数、当前结构页数、末尾框底、距安全页底的估计空白及溢出。`rendered_page_count` 使用 `layout_metrics` 的 `metrics["page_count"]`，不能由 builder 的页宿主数量代替。`rendered_content_bottom` 必须是实际页面上正文底边的 pt 坐标，不是框底。

- 实测页数或正文底边缺少任一项时，`density_status="unknown"`。`page_count` 和 `estimated_page_count` 都只是结构值，`page_count_source="builder_page_structure"`；它们不能证明实际 PDF 为一页。
- 两项实测值齐备后分别报告 `within_density_target`、`underfilled`、`too_close_to_bottom`、`overflow` 或 `page_count_mismatch`。实际多页、页数不符或与结构不一致不能通过；`visual_acceptance_proven` 始终为 false。
- 一页任务默认末端到页底安全线保留约 12–28 pt。过空时先检查被漏掉的真实信息及可充实的技术/职责内容；资料不足要说明缺口。过满时优化组织与冗余，不改纸长、不压字号、不裁事实。
- 报告不会拉伸对象、自动增大间距或编造内容。“页数为 1”和“没有越界”都不足以证明简历内容充实、分点正确或排版通过。

机构名称、专业、日期较长时调整真实制表位、字段措辞或条带尺寸；不能用连续空格撑列，也不能让字段覆盖独立图标。项目标题和仓库链接发生挤压时，先用准确的短显示标签和紧凑标题；不能丢弃真实链接、强制裁切或缩小到低于字体规范。

## 验证、渲染与交接

`validate()` 按组件角色检查：一个段落一个文字框、栏目/机构原生组结构、透明文字框、独立填色对象、机构图片与文本分离、垂直居中设置、组内对象边界、栏目线宽、项目右对齐制表位及超链接、关系完整性、对象 ID 和正文下边距安全线。教育组合会检查关联介绍框没有被移除；语义项目分点会检查实际原生编号层级。机构字段分别按对应制表位的可用宽度估计；可能跳列或折行时返回明确警告，保留全部文字，要求调整字段措辞、制表位或列宽并重新渲染。`save(path)` 调用验证后保存；不会自动填入示例身份。

验证报告明确 `requires_rendering=True`、`measures_text_fit=False`。结构通过不代表文字可见、字体替换正确或没有重叠，交付前仍须执行 [独立验收](acceptance.md)。

针对性结构回归测试：

```sh
python -m unittest discover -s scripts -p 'test_docx_components.py' -v
```

测试使用临时的全虚构文档，检查文本/链接保留、原生可编辑组、图标独立、不同页面尺寸、分页及退化检测；不替代视觉检查。维护样例时从空白文档重新生成并同步预览，不把任务简历或其部件纳入 skill。
