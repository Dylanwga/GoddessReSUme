#!/usr/bin/env python3
"""Rebuild the fictional layout reference from a blank document.

No input resume, document part, screenshot, or downloaded image is read.
Pillow draws original illustrative icons; they are not real institution logos.
Render and inspect every page after building. See references/acceptance.md.
"""
import argparse
import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
import posixpath
import tempfile
import zipfile
from lxml import etree
from PIL import Image, ImageDraw
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_components import ResumeBuilder


def icons(directory):
    result = {}
    blue = '#2F6FBA'
    for name in ('school', 'lab', 'company', 'phone', 'mail', 'repo', 'code', 'scholar'):
        im = Image.new('RGBA', (128, 128), (255, 255, 255, 0))
        d = ImageDraw.Draw(im)
        ink = blue if name in ('school', 'lab', 'company') else '#303840'
        if name == 'school':
            d.ellipse((6, 6, 122, 122), outline=ink, width=9)
            d.polygon([(29, 52), (64, 31), (99, 52)], fill=ink)
            for x in (37, 59, 81):
                d.rectangle((x, 57, x + 10, 88), fill=ink)
            d.rectangle((29, 94, 99, 102), fill=ink)
        elif name == 'company':
            d.rounded_rectangle((6, 6, 122, 122), radius=18, fill=ink)
            d.polygon([(27, 83), (49, 43), (66, 72), (85, 27), (105, 83)], fill='white')
        elif name == 'lab':
            d.polygon([(64, 5), (122, 64), (64, 123), (6, 64)], fill=ink)
            d.ellipse((39, 39, 89, 89), fill='white')
            d.ellipse((52, 52, 76, 76), fill=ink)
        elif name == 'mail':
            d.rounded_rectangle((8, 24, 120, 105), radius=8, outline=ink, width=9)
            d.line([(12, 31), (64, 70), (116, 31)], fill=ink, width=9)
        elif name == 'phone':
            d.rounded_rectangle((33, 5, 95, 123), radius=12, outline=ink, width=10)
            d.line((47, 23, 81, 23), fill=ink, width=6)
            d.ellipse((59, 102, 69, 112), fill=ink)
        elif name == 'repo':
            d.line([(35, 28), (35, 102), (95, 64), (95, 28)], fill=ink, width=10)
            for x, y in ((35, 24), (35, 104), (95, 24)):
                d.ellipse((x - 15, y - 15, x + 15, y + 15), fill=ink)
        elif name == 'code':
            d.line([(42, 26), (12, 64), (42, 102)], fill=ink, width=10)
            d.line([(86, 26), (116, 64), (86, 102)], fill=ink, width=10)
            d.line((75, 16, 54, 112), fill=ink, width=9)
        else:
            d.polygon([(5, 43), (64, 14), (123, 43), (64, 72)], fill=ink)
            d.rectangle((33, 66, 94, 99), fill=ink)
        path = directory / (name + '.png')
        im.save(path)
        result[name] = path
    return result


def clean_package(path):
    """Remove the blank library template thumbnail and identify the sample."""
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    removed = {n for n in parts if n.startswith(('docProps/thumbnail', 'customXml/'))}
    for name in removed:
        del parts[name]
    for name in list(parts):
        if not name.endswith('.rels'):
            continue
        rels = etree.fromstring(parts[name])
        owner_dir = posixpath.dirname(posixpath.dirname(name))
        for rel in list(rels):
            target = posixpath.normpath(posixpath.join(owner_dir, rel.get('Target', ''))).lstrip('/')
            if rel.get('TargetMode') != 'External' and target in removed:
                rels.remove(rel)
        parts[name] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
    types = etree.fromstring(parts['[Content_Types].xml'])
    for node in list(types):
        if node.get('PartName', '').lstrip('/') in removed:
            types.remove(node)
    parts['[Content_Types].xml'] = etree.tostring(types, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def build(output, *, manifest=None, metrics=None, body_align='left'):
    """Build a single fictional page; optional metrics tighten body frames.

    Measurements are reused only when the complete rendered text still matches.
    Re-render after tightening and inspect coverage, spacing and density.
    """
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    measured = {}
    if metrics:
        report = json.loads(Path(metrics).read_text(encoding='utf-8'))
        measured = {item['id']: item for item in report['blocks']}
    with tempfile.TemporaryDirectory(prefix='goddessresume-fictional-') as tmp:
        ico = icons(Path(tmp))
        b = ResumeBuilder(theme='2F6FBA')
        props = b.document.core_properties
        props.title = 'GoddessReSUme 虚构版式样例'
        props.subject = '全部身份、机构、经历、日期及图标均为虚构，仅演示可编辑版式'
        props.author = props.last_modified_by = 'GoddessReSUme'
        props.created = props.modified = datetime(2020, 1, 1, tzinfo=timezone.utc)
        b.add_block('name', '林知遥（虚构样例）', y=28, width=420, height=30)
        b.add_photo_placeholder(x=485, y=29, width=76.3, height=78)
        b.add_block('contact', [
            {'icon_path': ico['phone'], 'icon_height': 11}, {'text': ' 000-0000-0000'},
            {'text': ' ｜ '}, {'icon_path': ico['mail'], 'icon_height': 11},
            {'text': ' lin@example.com', 'url': 'mailto:lin@example.com'},
            {'text': ' ｜ '}, {'icon_path': ico['repo'], 'icon_height': 11},
            {'text': ' GitHub', 'url': 'https://example.com/profile'},
        ], y=63, width=440, height=22)
        b.add_tech_stack('Python / Go｜FastAPI · React · Kafka', icon_path=ico['code'],
                         y=87, width=440, height=22, font_size=10.5)

        def tighten(block):
            from docx.oxml.ns import qn
            key = block.anchor.find(qn('wp:docPr')).get('name')
            item = measured.get(key)
            if item:
                if item['match_status'] != 'matched':
                    raise ValueError(f'Measured text is incomplete: {key}')
                current = ''.join(n.text or '' for n in block.paragraph._p.iter(qn('w:t')))
                normalize = lambda value: re.sub(r'\s+', '', unicodedata.normalize('NFKC', value))
                if normalize(current) != normalize(item['observed_text']):
                    return block  # Revised content needs a fresh, conservative render.
                ink = item['ink_bounds']['yMax'] - item['frame']['y']
                b.set_body_layout(block, rendered_ink_bottom=ink)
            return block

        def education(y, degree, dates, detail):
            entry = b.add_education('星澜大学（虚构）', detail, degree, dates,
                logo_path=ico['school'], y=y, tab_stops=(184, 355, 524.3))
            previous = entry.band
            for i, block in enumerate(entry.details):
                b.move_block(block, y=previous.next_y(1 if i == 0 else .5))
                tighten(block)
                block.metadata['group_id'] = 'education-' + degree
                previous = block
            return entry

        sec = b.add_section('教育背景', y=119)
        masters = education(sec.next_y(4), '硕士', '2023—2026', '计算机科学与技术')
        bachelors = education(masters.next_y(7), '本科', '2019—2023', '软件工程')
        sec = b.add_section('项目经历', y=bachelors.next_y(10))
        band = b.add_institution('云舟科技（虚构）', '平台研发', '实习生', '2025.03—2025.09',
            logo_path=ico['company'], y=sec.next_y(4), tab_stops=(184, 355, 524.3))

        projects = [
            ('atlasflow', 'AtlasFlow：研发任务协作平台（虚构）', [
                ('background', '项目背景：', '面向跨角色研发任务中信息分散、状态难以同步和交接成本高的问题，构建连接任务拆解、执行跟踪与结果归档的协作工作流。', 34.2),
                ('responsibility', '项目职责：', '作为项目 Owner，负责流程建模与服务接口设计，拆解客户端和执行服务的接口边界，组织联调与异常场景验收。', 34.2),
                ('technical', '任务状态与失败恢复：', '针对超时重试、重复回调和阶段失败造成的状态不一致，以任务标识关联执行记录，先核对事件是否处理，再按允许的迁移推进任务，避免重复通知再次触发已完成阶段。失败时保留阶段产物与检查点，从最近成功阶段之后继续执行。通过重复回调、回调迟到和中断恢复样例，逐项核对状态迁移与产物引用，检查是否重复执行或丢失历史结果。', 69.8),
                ('technical', '接口契约与结果追溯：', '围绕客户端、任务服务与执行器之间的交接，使用 FastAPI 定义任务创建、阶段回调和结果查询接口，以任务标识、阶段编号及产物位置建立统一的数据契约。将失败状态与执行日志关联返回，使调用方可区分尚未完成与执行失败，并沿同一标识定位对应阶段和复现输入。', 52),
                ('result', '项目产出：', '交付可运行原型、接口说明与复现脚本；将联调中的失败样例沉淀为回归用例，为接入新工具和迭代流程提供可重复的验收基线。', 34.2),
            ]),
            ('tracenest', 'TraceNest：安全事件关联工具（虚构）', [
                ('background', '项目背景：', '针对告警上下文分散、重复人工检索和证据交接困难的问题，聚合主机与网络事件，为分析人员提供可追溯的事件调查视图。', 34.2),
                ('responsibility', '项目职责：', '作为项目 Owner，负责事件模型、关联接口与分析页面的设计，明确证据保留要求，组织端到端联调及边界条件验证。', 34.2),
                ('technical', '事件归一化与重复输入：', '针对不同来源的时间、主体及事件标识不一致，使用 Go 统一字段并保留原始记录位置，经 Kafka 向检索侧传递标准事件。以来源标识关联归一化结果和调查对象，结合重复、迟到与缺失字段样例核对记录及检索展示，避免清洗后丢失回溯依据。', 52),
                ('technical', '调查证据与来源关联：', '通过 React 时间线组织调查记录，将结论与对应原始事件、查询条件一起呈现。复核时可沿来源标识返回证据记录，区分当前分析结论与原始事实，支持调查过程交接。', 34.2),
                ('result', '项目产出：', '交付事件检索与证据回溯原型，并整理字段映射、异常处理约定及验证样例，供新增来源按统一接口接入。', 34.2),
            ]),
        ]
        previous = band
        for index, (slug, title_text, points) in enumerate(projects):
            title = b.add_project_header(title_text, repo_url='https://example.com/projects/' + slug,
                repo_label='GitHub', icon_path=ico['repo'], height=22,
                y=previous.next_y(3 if index == 0 else 7))
            previous = title
            for kind, label, text, initial_height in points:
                block = b.add_project_point(kind, text, label=label,
                    y=previous.next_y(.5), height=initial_height)
                block.paragraph.alignment = (WD_ALIGN_PARAGRAPH.LEFT if body_align == 'left'
                                             else WD_ALIGN_PARAGRAPH.JUSTIFY)
                block.metadata['group_id'] = slug
                previous = tighten(block)
        report = b.validate()
        b.save(output)
        if manifest:
            Path(manifest).write_text(json.dumps(b.layout_manifest(), ensure_ascii=False, indent=2), encoding='utf-8')
    clean_package(output)
    print(output)
    print(report)
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', nargs='?', type=Path, default=Path(__file__).resolve().parents[1] / 'assets/layout-reference.docx')
    parser.add_argument('--manifest', type=Path, help='Export frame geometry and complete text for rendered measurement')
    parser.add_argument('--metrics', type=Path, help='Use matching previous-render measurements to tighten body frames')
    parser.add_argument('--body-align', choices=('left', 'both'), default='left',
                        help='Sample body alignment; left avoids verified LibreOffice justified-punctuation clipping')
    args = parser.parse_args()
    build(args.output, manifest=args.manifest, metrics=args.metrics, body_align=args.body_align)
