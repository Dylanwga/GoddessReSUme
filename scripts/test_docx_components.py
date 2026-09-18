"""Component regression checks; temporary, fully fictional DOCX fixtures only.

Run with the bundled Python: python -m unittest discover -s scripts -p 'test_docx_components.py'
These checks cover editing structure and content preservation, not visual QA.
"""
from pathlib import Path
import json
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree
from PIL import Image

from docx_components import ResumeBuilder

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}


class ResumeComponentsTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.icon = self.root / 'fictional-icon.png'
        Image.new('RGB', (80, 40), '#2f6fba').save(self.icon)

    def tearDown(self):
        self.temp.cleanup()

    def read_output(self, builder, name='fictional.docx'):
        output = builder.save(self.root / name)
        with ZipFile(output) as archive:
            document = etree.fromstring(archive.read('word/document.xml'))
            relationships = etree.fromstring(archive.read('word/_rels/document.xml.rels'))
            core = etree.fromstring(archive.read('docProps/core.xml'))
        return document, relationships, core

    def test_editable_components_preserve_text_links_and_local_images(self):
        b = ResumeBuilder(theme='2F6FBA')
        b.add_section('示例教育背景', y=100)
        b.add_institution('虚构大学甲', '示例专业', '硕士', '20XX—20XX', logo_path=self.icon, y=134)
        b.add_project_header('虚构项目：调查流程', repo_url='https://example.com/repository?branch=demo&view=code',
                             repo_label='example.com/repository', icon_path=self.icon, y=172)
        b.add_block('bullet', [{'text': '项目背景：', 'bold': True}, {'text': '这是明确虚构的流程演示内容。'}], y=205)
        b.add_tech_stack(['示例语言', '示例框架'], icon_path=self.icon, y=240)
        root, rels, core = self.read_output(b)
        text = ''.join(root.xpath('//w:t/text()', namespaces=NS))
        for supplied in ('虚构大学甲', '示例专业', '硕士', '20XX—20XX', '虚构项目：调查流程',
                         '项目背景：', '这是明确虚构的流程演示内容。', '技术栈：', '示例语言 · 示例框架'):
            self.assertIn(supplied, text)
        self.assertIn('https://example.com/repository?branch=demo&view=code', [r.get('Target') for r in rels])
        self.assertEqual(len(root.xpath('//w:txbxContent', namespaces=NS)), 5)
        self.assertTrue(all(len(tx.xpath('./w:p', namespaces=NS)) == 1 for tx in root.xpath('//w:txbxContent', namespaces=NS)))
        self.assertFalse(root.xpath('//w:br | //w:pBdr', namespaces=NS))
        self.assertTrue(root.xpath('//w:hyperlink', namespaces=NS))
        self.assertFalse(root.xpath('//a:blip/@r:link', namespaces={**NS, 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}))
        self.assertFalse(core.xpath('//*[local-name()="creator" or local-name()="lastModifiedBy"]/text()'))
        result = b.validate()
        self.assertTrue(result['requires_rendering'])
        self.assertFalse(result['measures_text_fit'])

    def test_section_line_and_band_layers_are_independently_editable(self):
        b = ResumeBuilder(theme='2F6FBA')
        b.add_section('栏目示例', y=80)
        b.add_institution('虚构机构', '示例方向', '岗位', '20XX', logo_path=self.icon, y=115, height=30)
        root, _, _ = self.read_output(b)
        groups = root.xpath('//wpg:wgp', namespaces=NS)
        self.assertEqual(len(groups), 2)
        section, band = groups
        # A line that is a separate filled object survives a renderer's lack of paragraph borders.
        self.assertEqual(len(section.xpath('./wps:wsp', namespaces=NS)), 2)
        self.assertEqual(len(section.xpath('./wps:wsp[not(wps:txbx)]/wps:spPr/a:solidFill', namespaces=NS)), 1)
        rule_width = section.xpath('./wps:wsp[not(wps:txbx)]/wps:spPr/a:xfrm/a:ext/@cx', namespaces=NS)[0]
        group_width = section.xpath('./wpg:grpSpPr/a:xfrm/a:ext/@cx', namespaces=NS)[0]
        self.assertEqual(rule_width, group_width)
        self.assertEqual(len(band.xpath('./wps:wsp[not(wps:txbx)]', namespaces=NS)), 1)
        self.assertEqual(len(band.xpath('./pic:pic', namespaces=NS)), 1)
        self.assertEqual(len(band.xpath('./wps:wsp[wps:txbx]', namespaces=NS)), 1)
        self.assertFalse(band.xpath('./wps:wsp[wps:txbx]//pic:pic', namespaces=NS))
        self.assertTrue(band.xpath('./wps:wsp[wps:txbx]/wps:spPr/a:noFill', namespaces=NS))
        self.assertEqual(band.xpath('./wps:wsp[wps:txbx]/wps:bodyPr/@anchor', namespaces=NS), ['ctr'])
        body = band.xpath('./wps:wsp[wps:txbx]/wps:bodyPr', namespaces=NS)[0]
        self.assertEqual(body.get('tIns'), body.get('bIns'))
        logo = band.xpath('./pic:pic/pic:spPr/a:xfrm', namespaces=NS)[0]
        ly = int(logo.find('a:off', NS).get('y'))
        lh = int(logo.find('a:ext', NS).get('cy'))
        self.assertEqual(ly * 2 + lh, 30 * 12700)

    def test_other_page_sizes_and_explicit_pagination_keep_safe_boundaries(self):
        b = ResumeBuilder(page_width=420, page_height=595, margin_left=30, margin_right=25,
                          margin_top=24, margin_bottom=28)
        block = b.add_block('body', '虚构短句。', y=540, height=20)
        self.assertAlmostEqual(block.width, 365)
        page, y = b.next_position(block, gap=2, required_height=45)
        self.assertEqual(page, 2)
        self.assertAlmostEqual(y, 24)
        heading = b.add_section('第二页栏目', page=page, y=y)
        institution = b.add_institution('虚构大学', dates='20XX', page=page, y=heading.next_y(3))
        root, _, _ = self.read_output(b)
        self.assertEqual(b.validate()['pages'], 2)
        self.assertEqual(len(root.xpath('//w:pageBreakBefore', namespaces=NS)), 1)
        self.assertAlmostEqual(institution.width, 365)
        self.assertTrue(root.xpath('//w:tab[@w:val="right"]', namespaces=NS))
        bad = ResumeBuilder(page_width=420, page_height=595, margin_bottom=28)
        bad.add_block('body', '虚构越界内容。', y=560, height=20)
        with self.assertRaisesRegex(ValueError, 'safe-bottom'):
            bad.validate()

    def test_content_height_changes_with_content_and_does_not_delete_long_titles(self):
        b = ResumeBuilder()
        short = b.estimate_height('body', '简短的示例内容。', width=240)
        long = b.estimate_height('body', '较长的虚构背景说明，用于检查不同自然换行数量。' * 5, width=240)
        self.assertGreater(long, short)
        title = '虚构项目标题' * 15
        header = b.add_project_header(title, repo_url='https://example.com/long/repository', y=100, width=240)
        root, _, _ = self.read_output(b)
        self.assertIn(title, ''.join(root.xpath('//w:t/text()', namespaces=NS)))
        self.assertTrue(b.validate()['warnings'])
        self.assertEqual(header.page, 1)
        self.assertAlmostEqual(header.y, 100)
        with self.assertRaisesRegex(ValueError, 'line breaks'):
            b.add_block('body', '第一条\n第二条')

    def test_validation_detects_regressed_structure(self):
        b = ResumeBuilder()
        band = b.add_institution('虚构机构', y=100)
        shape = band.anchor.xpath('.//wps:wsp[wps:txbx]')[0]
        shape.find('wps:bodyPr', NS).set('anchor', 't')
        with self.assertRaisesRegex(ValueError, 'vertically centered'):
            b.validate()

    def test_multiline_cjk_estimate_covers_measured_render_regression(self):
        b = ResumeBuilder()
        text = '虚构系统用于演示长段落自然换行，所有内容均为组件测试。' * 6 + '末尾验证标记甲乙丙丁。'
        block = b.add_block('body', text, y=185, width=300)
        # An independent bundled-renderer probe at KaiTi 10.5 pt / 1.2 spacing
        # placed the final ink bottom 121.42 pt below this box's top. The old
        # nominal-em estimate (114.5 pt) clipped it. This measured lower bound
        # guards the actual failure rather than restating the estimator formula.
        self.assertGreater(block.height, 121.42)
        following = b.add_block('body', '下一段起始标记', y=block.next_y(2), width=300)
        self.assertGreater(following.y, 185 + 121.42)
        root, _, _ = self.read_output(b)
        self.assertIn('末尾验证标记甲乙丙丁。', ''.join(root.xpath('//w:t/text()', namespaces=NS)))
        self.assertEqual(block.paragraph._p.xpath('./w:pPr/w:spacing/@w:line'), ['288'])
        self.assertEqual(block.paragraph._p.xpath('./w:pPr/w:spacing/@w:lineRule'), ['auto'])
        self.assertGreater(b.estimate_height('body', text, width=300, font_size=12), block.height)
        self.assertGreater(b.estimate_height('body', text, width=240), block.height)

    def test_long_institution_fields_warn_without_deleting_content(self):
        b = ResumeBuilder()
        short = b.add_institution('虚构大学', '示例专业', '硕士', '20XX—20XX', logo_path=self.icon, y=60)
        self.assertEqual(b.validate()['warnings'], [])
        name = '虚构大学人工智能与网络空间安全联合研究学院'
        detail = '智能系统安全与隐私保护示例研究方向'
        band = b.add_institution(name, detail, '硕士', '20XX—20XX', logo_path=self.icon, y=100)
        warnings = b.validate()['warnings']
        self.assertTrue(any('name' in warning and 'tab field' in warning for warning in warnings))
        self.assertTrue(any('detail' in warning and 'tab field' in warning for warning in warnings))
        root, _, _ = self.read_output(b)
        self.assertIn(name, ''.join(root.xpath('//w:t/text()', namespaces=NS)))
        self.assertIn(detail, ''.join(root.xpath('//w:t/text()', namespaces=NS)))
        self.assertAlmostEqual(band.height, short.height)  # no silent resize or field truncation
        for unsupported in ([{'text': '虚构机构'}], '虚构机构\t示例方向'):
            with self.assertRaisesRegex(ValueError, 'use add_institution'):
                b.add_block('institution', unsupported)

    def test_education_defaults_keep_two_independent_editable_detail_boxes(self):
        b = ResumeBuilder()
        entry = b.add_education('虚构大学', '示例专业', '本科', '20XX—20XX', logo_path=self.icon, y=80)
        self.assertEqual(len(entry.details), 2)
        self.assertEqual(entry.field_names, ('academic', 'activities'))
        self.assertTrue(all(block.metadata['education_missing'] for block in entry.details))
        self.assertTrue(all(block.height == 20 for block in entry.details))
        self.assertIn('毕业设计 / 指导教师：', entry.details[0].paragraph.text)
        root, _, _ = self.read_output(b)
        self.assertEqual(len(root.xpath('//wp:anchor', namespaces=NS)), 3)
        self.assertEqual(len(root.xpath('//wpg:wgp', namespaces=NS)), 1)
        self.assertEqual(len(root.xpath('//wpg:wgp//w:txbxContent', namespaces=NS)), 1)
        self.assertEqual(len(root.xpath('//w:txbxContent', namespaces=NS)), 3)
        self.assertEqual(b.validate()['education_entries'], 1)
        self.assertAlmostEqual(entry.bottom, entry.details[-1].bottom)
        previous_second_y = entry.details[1].y
        b.move_block(entry.details[0], y=115)
        self.assertAlmostEqual(entry.details[1].y, previous_second_y)
        self.assertAlmostEqual(entry.band.y, 80)

    def test_education_partial_complete_and_explicit_no_reservation(self):
        cases = [
            ({'academic_info': '示例研究方向'}, 2, [False, True]),
            ({'academic_info': '示例研究方向', 'activities': '示例社团活动'}, 2, [False, False]),
            ({'reserve_missing': False}, 0, []),
            ({'reserve_missing': False, 'activities': '示例社团活动'}, 1, [False]),
        ]
        for options, count, missing in cases:
            with self.subTest(options=options):
                b = ResumeBuilder()
                entry = b.add_education('虚构大学', degree='硕士', y=80, **options)
                self.assertEqual(len(entry.details), count)
                self.assertEqual([d.metadata['education_missing'] for d in entry.details], missing)
                root, _, _ = self.read_output(b)
                text = ''.join(root.xpath('//w:t/text()', namespaces=NS))
                for supplied in (options.get('academic_info'), options.get('activities')):
                    if supplied:
                        self.assertEqual(text.count(supplied), 1)
                self.assertEqual(len(root.xpath('//w:txbxContent', namespaces=NS)), 1 + count)

    def test_education_cannot_silently_lose_associated_detail(self):
        b = ResumeBuilder()
        entry = b.add_education('虚构大学', degree='硕士', y=80)
        removed = entry.details[0].anchor
        removed.getparent().remove(removed)
        with self.assertRaisesRegex(ValueError, 'Education academic detail textbox was removed'):
            b.save(self.root / 'incomplete-education.docx')

    def test_project_points_apply_and_validate_semantic_bullet_hierarchy(self):
        b = ResumeBuilder()
        blocks = []
        for index, kind in enumerate(('background', 'responsibility', 'technical', 'result')):
            block = b.add_project_point(kind, '调用者提供的虚构示例内容。', y=80 + index * 40)
            blocks.append(block)
            self.assertEqual(block.paragraph._p.xpath('./w:pPr/w:numPr/w:ilvl/@w:val'),
                             ['1' if kind == 'technical' else '0'])
            if kind == 'responsibility':
                self.assertTrue(all(run.bold for run in block.paragraph.runs))
        root, _, _ = self.read_output(b)
        self.assertEqual(len(root.xpath('//w:numPr', namespaces=NS)), 4)
        self.assertEqual(len(root.xpath('//w:txbxContent', namespaces=NS)), 4)
        blocks[2].paragraph._p.xpath('./w:pPr/w:numPr/w:ilvl')[0].set('{%s}val' % NS['w'], '0')
        with self.assertRaisesRegex(ValueError, 'technical requires native bullet level 1'):
            b.validate()

    def test_measured_body_layout_reduces_geometry_without_changing_text_or_spacing(self):
        b = ResumeBuilder()
        text = '虚构系统用于演示长段落自然换行，所有内容均为组件测试。' * 6 + '末尾验证标记甲乙丙丁。'
        block = b.add_block('body', text, width=300, y=100)
        old_height = block.height
        old_paragraph = etree.tostring(block.paragraph._p)
        b.set_body_layout(block, rendered_ink_bottom=121.42, rendered_lines=7)
        self.assertLess(block.height, old_height)
        self.assertGreater(block.height, 121.42)
        self.assertEqual(etree.tostring(block.paragraph._p), old_paragraph)
        self.assertEqual(block.metadata['body_layout']['source'], 'rendered_ink_bottom')
        self.assertTrue(block.metadata['body_layout']['requires_rerendering'])
        following = b.add_block('body', '后续虚构段落', y=300)
        b.move_block(following, y=block.next_y())
        self.assertAlmostEqual(following.y - block.bottom, .5)
        self.read_output(b)

    def test_one_page_report_distinguishes_estimates_from_render_measurements(self):
        b = ResumeBuilder()
        block = b.add_block('body', '虚构内容', y=700, height=20)
        estimate = b.one_page_report()
        self.assertEqual(estimate['target_pages'], 1)
        self.assertEqual(estimate['density_status'], 'unknown')
        self.assertEqual(estimate['estimated_page_count'], estimate['page_count'])
        self.assertEqual(estimate['page_count_source'], 'builder_page_structure')
        self.assertFalse(estimate['visual_acceptance_proven'])
        self.assertEqual(b.one_page_report(rendered_content_bottom=b.safe_bottom - 20)['density_status'], 'unknown')
        self.assertEqual(b.one_page_report(rendered_page_count=1)['density_status'], 'unknown')
        for remaining, status in ((20, 'within_density_target'), (100, 'underfilled'),
                                  (5, 'too_close_to_bottom'), (-2, 'overflow')):
            report = b.one_page_report(rendered_content_bottom=b.safe_bottom - remaining, rendered_page_count=1)
            self.assertEqual(report['density_status'], status)
            self.assertFalse(report['visual_acceptance_proven'])
        # A renderer can add a page even though the builder has only one host.
        self.assertEqual(b.one_page_report(rendered_content_bottom=b.safe_bottom - 20,
                                          rendered_page_count=2)['density_status'], 'overflow')
        self.assertEqual(b.one_page_report(target_pages=2, rendered_content_bottom=b.safe_bottom - 20,
                                          rendered_page_count=1)['density_status'], 'page_count_mismatch')
        for invalid_count in (0, 1.5, True):
            with self.assertRaisesRegex(ValueError, 'rendered_page_count'):
                b.one_page_report(rendered_page_count=invalid_count)
        extra = b.add_block('body', '虚构第二页内容', page=2, y=50)
        self.assertTrue(b.one_page_report()['estimated_overflow'])
        self.assertEqual(b.one_page_report(rendered_content_bottom=b.safe_bottom - 20,
                                          rendered_page_count=1)['density_status'], 'page_count_mismatch')
        b.move_block(extra, page=1, y=block.next_y())
        self.assertEqual(b.one_page_report()['page_count'], 1)
        self.read_output(b)

    def test_manifest_preserves_hyperlinked_text_run_tabs_and_semantic_associations(self):
        b = ResumeBuilder()
        header = b.add_project_header('虚构项目', repo_url='https://example.com/repo',
                                      repo_label='example.com/repo', icon_path=self.icon, y=80)
        point = b.add_project_point('technical', '示例实现内容。', y=110)
        point.metadata['group_id'] = 'fictional-project-a'
        education = b.add_education('虚构大学', '示例专业', '硕士', '20XX',
                                     academic_info='示例研究方向', y=150)
        manifest = b.layout_manifest()
        json.dumps(manifest, ensure_ascii=False)  # suitable for an external renderer/measurement script
        records = {row['id']: row for row in manifest['blocks']}
        header_record = records[header.anchor.find('wp:docPr', NS).get('name')]
        self.assertEqual(header_record['text'], '虚构项目\t example.com/repo')
        point_record = records[point.anchor.find('wp:docPr', NS).get('name')]
        self.assertEqual(point_record['semantic_kind'], 'technical')
        self.assertEqual(point_record['group_id'], 'fictional-project-a')
        band_record = records[education.band.anchor.find('wp:docPr', NS).get('name')]
        self.assertEqual(band_record['text'], '虚构大学\t示例专业\t硕士\t20XX')
        academic = records[education.details[0].anchor.find('wp:docPr', NS).get('name')]
        placeholder = records[education.details[1].anchor.find('wp:docPr', NS).get('name')]
        self.assertEqual(academic['role'], 'education_detail')
        self.assertFalse(academic['intentional_blank'])
        self.assertEqual(placeholder['role'], 'education_placeholder')
        self.assertTrue(placeholder['intentional_blank'])
        self.assertEqual(len(records), len(b.blocks))

    def test_education_details_support_measured_body_layout_and_manifest_refresh(self):
        b = ResumeBuilder()
        education = b.add_education('虚构大学', degree='本科', y=80)
        detail = education.details[0]
        self.assertEqual(detail.role, 'body')
        b.set_body_layout(detail, rendered_ink_bottom=15, bottom_padding=1, safety_margin=.5)
        b.move_block(detail, y=112)
        record = next(row for row in b.layout_manifest()['blocks']
                      if row['id'] == detail.anchor.find('wp:docPr', NS).get('name'))
        self.assertEqual(record['role'], 'education_placeholder')
        self.assertAlmostEqual(record['height'], 16.5)
        self.assertAlmostEqual(record['y'], 112)
        self.assertTrue(record['intentional_blank'])
        b.validate()


if __name__ == '__main__':
    unittest.main()
