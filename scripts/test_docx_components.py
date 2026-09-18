"""Component regression checks; temporary, fully fictional DOCX fixtures only.

Run with the bundled Python: python -m unittest discover -s scripts -p 'test_docx_components.py'
These checks cover editing structure and content preservation, not visual QA.
"""
from pathlib import Path
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


if __name__ == '__main__':
    unittest.main()
