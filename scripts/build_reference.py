#!/usr/bin/env python3
"""Rebuild the fictional layout reference from a blank document.

No input resume, document part, screenshot, or downloaded image is read.
Pillow draws original illustrative icons; they are not real institution logos.
Render and inspect every page after building. See references/acceptance.md.
"""
import argparse
from datetime import datetime, timezone
from pathlib import Path
import posixpath
import tempfile
import zipfile
from lxml import etree
from PIL import Image, ImageDraw
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


def build(output):
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='goddessresume-fictional-') as tmp:
        ico = icons(Path(tmp))
        b = ResumeBuilder(theme='2F6FBA')
        props = b.document.core_properties
        props.title = 'GoddessReSUme 虚构版式样例'
        props.subject = '身份、机构、经历、论文、日期及图标均为虚构，仅演示可编辑版式'
        props.author = props.last_modified_by = 'GoddessReSUme'
        props.created = props.modified = datetime(2020, 1, 1, tzinfo=timezone.utc)
        b.add_block('name', '林知遥（虚构样例）', y=28, width=420, height=30)
        b.add_photo_placeholder(x=485, y=29, width=76.3, height=98)
        b.add_block('contact', [
            {'icon_path': ico['phone'], 'icon_height': 11}, {'text': ' 000-0000-0000'},
            {'text': ' ｜ '}, {'icon_path': ico['mail'], 'icon_height': 11},
            {'text': ' lin@example.com', 'url': 'mailto:lin@example.com'},
        ], y=63, width=440, height=22)
        b.add_block('contact', [
            {'icon_path': ico['repo'], 'icon_height': 11},
            {'text': ' GitHub', 'url': 'https://example.com/profile'}, {'text': ' ｜ '},
            {'icon_path': ico['scholar'], 'icon_height': 11},
            {'text': ' Google Scholar', 'url': 'https://example.com/scholar'},
        ], y=86, width=440, height=22)
        b.add_tech_stack('Python / Go｜FastAPI · React · Kafka', icon_path=ico['code'],
                         y=109, width=440, height=22, font_size=10.5)

        def paragraph(label, text, y, *, page=1, height=None, bullet=False, bold=False):
            runs = [{'text': label, 'bold': True}, {'text': text, 'bold': bold}]
            return b.add_block('bullet' if bullet else 'body', runs, page=page, y=y,
                               height=height, level=0)

        sec = b.add_section('教育背景', y=145)
        band = b.add_institution('星澜大学（虚构）', '计算机科学与技术', '硕士', '2023—2026',
                                 logo_path=ico['school'], y=sec.bottom + 4, tab_stops=(184, 355, 524.3))
        line = paragraph('研究方向：', '面向开发者工具的智能体协作；导师：周知衡教授（虚构）。', band.bottom + 3)
        band = b.add_institution('星澜大学（虚构）', '软件工程', '本科', '2019—2023',
                                 logo_path=ico['school'], y=line.bottom + 7, tab_stops=(184, 355, 524.3))
        line = paragraph('在校经历：', '参与校内开源社团，组织代码阅读与工程实践活动。', band.bottom + 3)

        sec = b.add_section('项目经历', y=line.bottom + 12)
        band = b.add_institution('云舟科技（虚构）', '平台研发', '实习生', '2025.03—2025.09',
                                 logo_path=ico['company'], y=sec.bottom + 4, tab_stops=(184, 355, 524.3))
        title = b.add_project_header('AtlasFlow：研发任务协作平台（虚构）', repo_url='https://example.com/projects/atlasflow',
                                     repo_label='GitHub', icon_path=ico['repo'], y=band.bottom + 3)
        line = paragraph('项目背景：', '面向多角色研发任务中信息分散与交接成本高的问题，构建任务拆解、执行跟踪和结果归档的一体化工作流。', title.bottom, height=36)
        line = paragraph('项目职责：', '作为项目 Owner，负责已约定范围内的流程建模与服务接口设计，协调客户端和执行服务的联调验证。', line.next_y(1), height=36, bold=True)
        line = paragraph('执行编排：', '将任务状态、工具调用和异常恢复组织为可追踪流程，通过消息事件连接执行服务与结果存储。', line.next_y(1), height=20, bullet=True)
        line = paragraph('质量验证：', '建立固定样例与失败分类，检查超时重试、重复事件和中断恢复，并以执行日志定位交接环节的问题。', line.next_y(1), height=36, bullet=True)
        line = paragraph('项目产出：', '交付接口说明、复现脚本和验收样例，形成可继续扩展的任务执行原型。', line.next_y(1), height=20)

        title = b.add_project_header('TraceNest：安全事件关联工具（虚构）', repo_url='https://example.com/projects/tracenest',
                                     repo_label='GitHub', icon_path=ico['repo'], y=line.bottom + 9)
        line = paragraph('项目背景：', '针对告警上下文分散、人工排查重复的问题，聚合主机与网络事件，为分析人员提供可追溯的事件视图。', title.bottom, height=36)
        line = paragraph('项目职责：', '作为项目 Owner，负责事件模型、关联接口和验证样例；将分析结论与对应证据一并保留。', line.next_y(1), height=20, bold=True)
        line = paragraph('工程实现：', '使用 Go 接入事件流，结合 Kafka 处理异步消息；通过统一字段和关联标识连接检索结果与分析页面。', line.next_y(1), height=36, bullet=True)
        paragraph('项目产出：', '形成事件检索与证据回溯原型，并整理边界条件与复现步骤。', line.next_y(1), height=20)

        sec = b.add_section('科研经历', page=2, y=30)
        band = b.add_institution('澄星研究院（虚构）', '智能系统研究组', '研究助理', '2024.10—2025.06',
                                 logo_path=ico['lab'], page=2, y=sec.bottom + 4, tab_stops=(184, 355, 524.3))
        title = b.add_project_header('FlexRoute：约束条件下的任务路由（虚构）', repo_url='https://example.com/research/flexroute',
                                     repo_label='GitHub', icon_path=ico['repo'], page=2, y=band.bottom + 3)
        line = paragraph('研究背景：', '面向工具成本和响应时限不同的任务，探索在资源预算内选择执行路径，避免统一路由策略造成不必要的调用。', title.bottom, page=2, height=36)
        line = paragraph('个人职责：', '负责路由策略原型和评测流程，将任务约束、候选路径与失败恢复条件表示为可重复执行的配置。', line.next_y(1), page=2, height=20, bold=True)
        line = paragraph('方法设计：', '按任务难度与可用工具建立分层决策过程，对比静态规则与动态选择，并保留路径选择的解释信息。', line.next_y(1), page=2, height=20, bullet=True)
        line = paragraph('实验分析：', '固定数据划分、调用预算和评估脚本，分别分析完成率、耗时与调用成本，定位收益来自哪些任务类型。', line.next_y(1), page=2, height=36, bullet=True)
        line = paragraph('研究产出：', '整理实验配置、失败案例和消融结论，形成内部研究报告与可复现实验包。', line.next_y(1), page=2, height=20)

        title = b.add_project_header('CacheWeave：多阶段检索缓存（虚构）', repo_url='https://example.com/research/cacheweave',
                                     repo_label='GitHub', icon_path=ico['repo'], page=2, y=line.bottom + 10)
        line = paragraph('研究背景：', '针对连续查询中重复检索和上下文重建的开销，探索可复用中间结果及其失效条件，兼顾响应效率与结果一致性。', title.bottom, page=2, height=36)
        line = paragraph('个人职责：', '负责缓存键设计和一致性验证，分析数据更新、查询变化与中间结果复用之间的约束关系。', line.next_y(1), page=2, height=20, bold=True)
        line = paragraph('系统实现：', '按检索阶段组织缓存记录，保留来源版本与依赖关系；在数据变化时失效关联条目，并记录命中与重算原因。', line.next_y(1), page=2, height=36, bullet=True)
        line = paragraph('验证方式：', '构造重复查询、数据更新及并发访问样例，对照无缓存路径核对结果一致性，输出可定位的差异记录。', line.next_y(1), page=2, height=36, bullet=True)

        sec = b.add_section('论文与研究记录', page=2, y=line.bottom + 12)
        entries = [
            ('Lin Zhiyao', ', Zhou Zhiheng. FlexRoute: Budget-Aware Task Routing. 虚构研究报告，2025。', 'https://example.com/papers/flexroute'),
            ('Lin Zhiyao', ', Chen Shuyi. CacheWeave: Reusable Retrieval States. 虚构研究报告，2025。', 'https://example.com/papers/cacheweave'),
        ]
        y = sec.bottom + 3
        for i, (author, rest, url) in enumerate(entries, 1):
            line = b.add_block('citation', [{'text': f'[{i}] '}, {'text': author, 'bold': True, 'color': b.theme},
                  {'text': rest, 'italic': True}, {'text': ' 链接', 'url': url}], page=2, y=y, height=22)
            y = line.next_y(2)
        b.add_block('body', '说明：本样例的全部资料与标识均为虚构，仅供版式参考；正式简历应使用真实资料和机构图标。',
                    page=2, y=y + 8, font_size=10.5, height=36, color='666666')
        report = b.validate()
        b.save(output)
    clean_package(output)
    print(output)
    print(report)
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', nargs='?', type=Path, default=Path(__file__).resolve().parents[1] / 'assets/layout-reference.docx')
    build(parser.parse_args().output)
