# 从空白文档构建原生简历组件

本页是 [排版流程](workflow-layout.md) 的组件接口参考。局部修改用户当前文件按 [Word 精修](word-editing.md) 操作，不能重建后覆盖手工调整。字体、颜色关系和层级由 [版式规范](style-spec.md) 统一规定。

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

教育信息按事实排入导师、研究方向、活动、荣誉等独立段落。**仅用户要求保留填写位置时**可传入空白字段标签；不要为填满页面虚构事实或默认留下“待补充”。每个项目背景独立传入，说明场景、问题或目标，不把职责重复一遍充作背景。

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
school = b.add_institution(
    "虚构大学甲", "示例专业", "硕士", "20XX—20XX",
    logo_path="/task/assets/fictional-school.png", y=section.next_y(4),
)
education = b.add_block("body", [
    {"text": "教育介绍：", "bold": True},
    {"text": "传入真实的研究方向、导师或在校活动；这里仅为虚构样例。"},
], y=school.next_y(3))

# 预留栏目和第一条机构带的组合空间；已有对象不会被隐式搬动。
page, y = b.next_position(education, gap=8, required_height=26 + 4 + 26)
section = b.add_section("项目经历", page=page, y=y)
band = b.add_institution(
    "虚构实验室", "示例方向", "", "项目 Owner",
    logo_path="/task/assets/fictional-lab.png", page=page, y=section.next_y(4),
)
project = b.add_project_header(
    "示例项目：调查流程自动化",
    repo_url="https://example.com/repository",
    repo_label="example.com/repository",
    icon_path="/task/assets/github.png", page=page, y=band.next_y(3),
)
background = b.add_block("bullet", [
    {"text": "项目背景：", "bold": True},
    {"text": "描述有依据的使用场景、现有问题和项目目标。"},
], level=0, page=page, y=project.next_y(2))
responsibility = b.add_block("bullet", [
    {"text": "项目职责：", "bold": True},
    {"text": "明确已确认的 Owner 角色及责任范围。"},
], level=0, page=page, y=background.next_y(1))
b.add_block("bullet", [
    {"text": "技术实现：", "bold": True},
    {"text": "传入有依据的技术方案及实现说明。"},
], level=1, page=page, y=responsibility.next_y(1))

report = b.validate()  # 报告结构、页边界和可能过长的标题链接。
b.save("/task/outputs/resume.docx")
# 接下来用批准的文档渲染器渲染全部页面，核对实际字体、行数和视觉效果。
```

示例中的 `gap` 是布局起点，不能作为所有简历的固定间距。页面不足时继续按内容分页，不压字号或裁文字。

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
- `b.next_position(block, gap=2, required_height=0)`：返回 `(page, y)`。空间不足则返回下一页和上边距，不移动现有对象，也不自动拆分段落。栏目与首个内容块须将组合高度传给 `required_height`。

### 内容与几何

`add_block(role, runs_or_text, *, page=1, x=None, y=None, width=None, height=None, font_size=None, color=None, level=1, icon_path=None, icon_height=11, line_spacing=1.2, line_height=None, padding_y=2)`

- 常用 `role`：`name`、`body`、`bullet`、`citation`、`contact`、`project`、`tech_stack`。`section` 转交栏目组件；`institution` 简写入口**只接收无制表符的机构名称字符串**，不支持 run 列表或把多个字段拼在一起。完整机构字段、制表位和图标参数使用 `add_institution()`。
- 文本参数可为字符串、run 字典或 run 列表。文本 run 支持 `text`、`bold`、`italic`、`underline`、`color`、`url`、`font_size`。图标 run 支持 `icon_path`、`icon_height`。
- 每个调用只创建一个逻辑段落；不允许传入 `\n`/`\r` 合并条目。文字自然换行由 Word 排版。
- `line_spacing` 是相对倍数；`line_height` 传 pt 时改为固定行高。联系方式、项目小标题和技术栈默认固定 18 pt 行高。正文默认 1.2 倍，**不是固定 14.4 pt**。
- `height=None` 按文本宽度、字号、行高和内边距估算。中文自然行高预留额外字体度量余量，避免把名义字号乘以倍数误当成实际行框高度；实际段落仍保持指定的 1.2 倍行距。估算不读取 Word 的实际排版结果，不保证刚好装下；显式 `height` 不会被覆盖。
- `bullet` 使用原生编号；`level=0` 为一级实心，`level=1` 为二级空心。符号字体不改变正文中西文字体。
- `citation` 默认斜体，具体加粗与作者强调由 run 传入，不能把所有论文一律加粗。

所有创建方法返回 `Block`，提供 `.paragraph`、`.anchor`、`.page`、`.x`、`.y`、`.width`、`.height`、`.bottom`、`.next_y(gap=2)`，便于连续排版与局部精修。`metadata` 保存估算与组件信息，不属于候选人事实。

### 专用组件

| 接口 | 参数与行为 |
| --- | --- |
| `add_section(text, ..., height=26, font_size=14, rule_thickness=1.5, rule_y=18.5, padding_y=2)` | 原生标题/粗线组。`rule_y` 相对组顶部，`rule_thickness` 为粗线高度。与正文列等宽；仅用于栏目层级。 |
| `add_institution(name, detail="", degree="", dates="", ..., logo_path=None, logo_height=18, logo_gap=4, tab_stops=None, height=26, font_size=11, text_baseline=0, line_height=16, tint_alpha=10000)` | 独立背景/图标/文字组。三处真实制表位偏移相对于整个条带，最后一处右对齐；默认按当前宽度计算。无中间字段传空字符串。图片等比缩放、垂直居中，文字框透明居中。`text_baseline` 仅用于渲染后的光学补偿，可正负微调，默认 0。 |
| `add_project_header(title, ..., repo_url=None, repo_label=None, icon_path=None, icon_height=11, height=24, font_size=10.5, link_font_size=10.5, min_gap=10)` | 小标题与仓库同行。`repo_label` 只控制显示；默认从 URL 省略协议，真实链接目标不变。链接不继承标题加粗，图标在链接前。宽度估算超限会报告警告，需缩短显示文字并渲染确认。 |
| `add_contact(text, url=None, icon_path=None, **placement)` | 单项联系行。多个联系项目同行用 `add_block("contact", [...])`，不要用重复调用制造重叠。 |
| `add_tech_stack(items, icon_path=None, label="技术栈：", separator=" · ", **placement)` | `items` 为字符串或字符串列表。技术项只来自提供的事实/项目证据；可手动使用 `｜` 分组。 |
| `add_photo_placeholder(label="照片占位", ..., width=76.17, height=106.64)` | 原生灰色占位框，只用于明确请求的模板或完全虚构样例。 |
| `estimate_height(role, runs_or_text, width=None, font_size=None, line_spacing=1.2, line_height=None, padding_y=2, level=1)` | 返回初步框高，便于安排下一块。调用者仍需检查渲染后的实际文字。 |

上述 1.5 pt、18.5 pt、26 pt 等是**可调整的组件起点**，不是所有字体、纸型或用户参考的硬性验收阈值。不得把某份样例的框数、页数、0.5 pt 基线补偿或所有段落框高固定下来。

## 内容驱动的高度与间距

布局按“当前文字需要的框高 + 同级间距”依次计算，不能所有正文无条件分配同样的两三行大框。一行内容与两行内容的框高应不同；栏目间距通常大于同一项目内的段间距，同级保持一致。

高度估算有字体度量误差，尤其中文、西文、粗体混排和长 URL。渲染后同时检查**实际文字间距**和框位置：若框的空白过大，应收紧框高；如果文字已经贴底，则应增高而非压缩。文字、图标、底色分别可编辑并不代表视觉已对齐。

机构名称、专业、日期较长时调整真实制表位、字段措辞或条带尺寸；不能用连续空格撑列，也不能让字段覆盖独立图标。项目标题和仓库链接发生挤压时，先用准确的短显示标签和紧凑标题；不能丢弃真实链接、强制裁切或缩小到低于字体规范。

## 验证、渲染与交接

`validate()` 按组件角色检查：一个段落一个文字框、栏目/机构原生组结构、透明文字框、独立填色对象、机构图片与文本分离、垂直居中设置、组内对象边界、栏目线宽、项目右对齐制表位及超链接、关系完整性、对象 ID 和正文下边距安全线。机构的学校/公司名称、专业/岗位以及学位与日期，分别按对应制表位的可用宽度估计；可能跳列或折行时返回明确警告，保留全部文字，要求调整字段措辞、制表位或列宽并重新渲染。`save(path)` 调用验证后保存；不会自动覆盖样例内容或填入示例身份。

验证报告明确 `requires_rendering=True`、`measures_text_fit=False`。结构通过不代表文字可见、字体替换正确或没有重叠，交付前仍须执行 [独立验收](acceptance.md)。

针对性结构回归测试：

```sh
python -m unittest discover -s scripts -p 'test_docx_components.py' -v
```

测试使用临时的全虚构文档，检查文本/链接保留、原生可编辑组、图标独立、不同页面尺寸、分页及退化检测；不替代视觉检查。维护样例时从空白文档重新生成并同步预览，不把任务简历或其部件纳入 skill。
