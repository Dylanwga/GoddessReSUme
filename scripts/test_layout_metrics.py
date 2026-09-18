"""Rendered-bounds regression checks with entirely fictional XHTML fixtures."""

from pathlib import Path
import tempfile
import unittest
from xml.sax.saxutils import escape

from layout_metrics import measure, read_bbox, read_manifest


def fixture(pages):
    body = []
    for words in pages:
        lines = []
        for text, left, top, right, bottom in words:
            lines.append(f'<line xMin="{left}" yMin="{top}" xMax="{right}" yMax="{bottom}">'
                         f'<word xMin="{left}" yMin="{top}" xMax="{right}" yMax="{bottom}">'
                         f'{escape(text)}</word></line>')
        body.append('<page width="300" height="200"><flow><block>' + ''.join(lines) + '</block></flow></page>')
    return ('<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><body><doc>'
            + ''.join(body) + '</doc></body></html>')


def frame(identifier, text, y, *, height=14, role="body", page=1, **extra):
    return dict(id=identifier, page=page, x=20, y=y, width=260, height=height,
                role=role, text=text, **extra)


class LayoutMetricsTests(unittest.TestCase):
    def parse(self, pages):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'fictional.xhtml'
            path.write_text(fixture(pages), encoding='utf-8')
            before = path.read_bytes()
            result = read_bbox(path)
            self.assertEqual(before, path.read_bytes())
            return result

    def test_actual_text_not_frame_footer_or_decoration_defines_bottom(self):
        pages = self.parse([[('虚构姓名', 25, 20, 80, 30), ('虚构项目正文', 25, 150, 130, 160),
                             ('1', 140, 188, 145, 196)]])
        frames = [frame('name', '虚构姓名', 18, role='name'),
                  frame('body', '虚构项目正文', 148, height=32),
                  frame('page-number', '1', 186, role='footer')]
        result = measure(pages, frames, safe_bottom_margin=20)
        page = result['pages'][0]
        self.assertEqual(page['all_text_last_line_bottom'], 196)
        self.assertEqual(page['body_last_line_bottom'], 160)
        self.assertEqual(page['substantive_bottom_gap'], 20)
        self.assertEqual(page['substantive_fill_status'], 'within_target')
        self.assertEqual(result['blocks'][1]['unused_height'], 22)
        self.assertEqual(result['blocks'][1]['actual_bottom_slack'], 20)

    def test_no_manifest_reports_unknown_not_a_fill_pass(self):
        result = measure(self.parse([[('虚构正文', 25, 150, 90, 160)]]), safe_bottom_margin=20)
        self.assertEqual(result['page_count'], 1)
        self.assertEqual(result['all_text_line_count'], 1)
        self.assertIsNone(result['pages'][0]['body_last_line_bottom'])
        self.assertEqual(result['pages'][0]['substantive_fill_status'], 'unknown')
        self.assertEqual(result['line_break_status'], 'unknown')
        self.assertEqual(result['line_break_issues'], [])

    def test_mixed_font_fragments_share_one_visual_line_in_reading_order(self):
        pages = self.parse([[('示例学院', 25, 40.4, 100, 50.4), ('2020', 210, 40, 240, 51)]])
        result = measure(pages, [frame('a', '示例学院\t2020', 38, role='education_detail')])
        self.assertEqual(result['blocks'][0]['match_status'], 'matched')
        self.assertEqual(result['blocks'][0]['actual_line_count'], 1)
        self.assertEqual(result['all_text_line_count'], 1)
        self.assertEqual(result['pages'][0]['pdf_line_fragment_count'], 2)

    def test_real_lines_gaps_and_tall_frame_are_separate_measurements(self):
        pages = self.parse([[('第一段第一行', 25, 40, 140, 50), ('第一段第二行', 25, 55, 140, 65),
                             ('第二段正文', 25, 85, 140, 95)]])
        frames = [frame('one', '第一段第一行第一段第二行', 38, height=45, group_id='fictional-a'),
                  frame('two', '第二段正文', 83, group_id='fictional-a')]
        result = measure(pages, frames, safe_bottom_margin=20)
        self.assertEqual(result['blocks'][0]['actual_line_count'], 2)
        self.assertEqual(result['blocks'][0]['actual_bottom_slack'], 18)
        gap = result['adjacent_body_gaps'][0]
        self.assertEqual(gap['visible_gap'], 20)
        self.assertEqual(gap['status'], 'wide')
        self.assertEqual(result['pages'][0]['substantive_fill_status'], 'underfilled')

    def test_complete_text_disambiguates_overbroad_geometry(self):
        pages = self.parse([[('甲项目正文', 25, 40, 130, 50), ('乙项目正文', 25, 70, 130, 80)]])
        frames = [frame('a', '甲项目正文', 35, height=50), frame('b', '乙项目正文', 68)]
        result = measure(pages, frames)
        self.assertEqual([b['match_status'] for b in result['blocks']], ['matched', 'matched'])
        self.assertEqual([b['actual_line_count'] for b in result['blocks']], [1, 1])
        self.assertEqual(result['blocks'][0]['ink_bounds']['yMax'], 50)

    def test_duplicate_text_inside_broad_frame_remains_ambiguous(self):
        pages = self.parse([[('重复示例', 25, 40, 130, 50), ('重复示例', 25, 70, 130, 80)]])
        result = measure(pages, [frame('a', '重复示例', 35, height=50)])
        self.assertEqual(result['blocks'][0]['match_status'], 'ambiguous')
        self.assertEqual(result['blocks'][0]['complete_text_match_count'], 2)
        self.assertIsNone(result['blocks'][0]['actual_line_count'])
        self.assertEqual(result['pages'][0]['substantive_fill_status'], 'unknown')

    def test_overlapping_claims_and_missing_text_do_not_pass(self):
        pages = self.parse([[('虚构正文', 25, 150, 130, 160)]])
        duplicate = measure(pages, [frame('a', '虚构正文', 148), frame('b', '虚构正文', 148)])
        self.assertTrue(all(b['match_status'] == 'ambiguous' for b in duplicate['blocks']))
        absent = measure(pages, [frame('a', '另一段并未渲染的正文', 148)])
        self.assertEqual(absent['blocks'][0]['match_status'], 'unknown')
        self.assertEqual(absent['pages'][0]['substantive_fill_status'], 'unknown')

    def test_education_filling_slot_cannot_fake_substantive_density(self):
        pages = self.parse([[('项目正文', 25, 80, 100, 90), ('导师／活动：____', 25, 150, 180, 160)]])
        frames = [frame('body', '项目正文', 78),
                  frame('slot', '导师/活动：____', 148, role='education_placeholder', intentional_blank=True)]
        result = measure(pages, frames, safe_bottom_margin=20)
        page = result['pages'][0]
        self.assertEqual(page['body_bottom_gap'], 20)
        self.assertEqual(page['substantive_bottom_gap'], 90)
        self.assertEqual(page['substantive_fill_status'], 'underfilled')
        self.assertEqual(page['eligible_body_line_count'], 2)

    def test_bullets_and_whitespace_are_not_missing_content(self):
        pages = self.parse([[('\uf09f', 20, 40, 24, 50), ('示例 A', 27, 40, 90, 50),
                             ('后续内容', 25, 55, 110, 65)]])
        frames = [frame('a', '示例 A 后续内容', 38, height=30, role='bullet')]
        result = measure(pages, frames)
        self.assertEqual(result['blocks'][0]['match_status'], 'matched')
        self.assertEqual(result['body_coverage'], 'matched')

    def test_page_overflow_and_configurable_thresholds(self):
        pages = self.parse([[('第一页内容', 25, 140, 150, 150)], [('第二页内容', 25, 175, 150, 190)]])
        result = measure(pages, [frame('a', '第一页内容', 138), frame('b', '第二页内容', 173, height=20, page=2)],
                         safe_bottom_margin=20, target_bottom_gap=(10, 35))
        self.assertFalse(result['page_count_matches_target'])
        self.assertEqual(result['pages'][0]['substantive_fill_status'], 'within_target')
        self.assertEqual(result['pages'][1]['substantive_fill_status'], 'past_safe_bottom')

    def test_native_second_level_o_bullet_is_ignored_only_at_bullet_position(self):
        pages = self.parse([[('o', 20, 40, 24, 50), ('技术正文', 28, 40, 100, 50)]])
        result = measure(pages, [frame('a', '技术正文', 38, role='bullet')])
        self.assertEqual(result['body_coverage'], 'matched')
        extra = self.parse([[('o', 20, 110, 24, 120), ('技术正文', 28, 40, 100, 50)]])
        result = measure(extra, [frame('a', '技术正文', 38, role='bullet')])
        self.assertEqual(result['body_coverage'], 'unknown')

    def test_line_break_diagnostics_preserve_normal_punctuation_and_latin_text(self):
        pages = self.parse([[('“虚构验证（完成）”，', 25, 40, 190, 50),
                             ('下一阶段。', 25, 55, 110, 65),
                             ('FastAPI', 25, 90, 100, 100), ('2026', 25, 105, 70, 115)]])
        result = measure(pages, [frame('chinese', '“虚构验证（完成）”，下一阶段。', 38, height=30),
                                 frame('latin', 'FastAPI 2026', 88, height=30)])
        self.assertEqual(result['line_break_status'], 'no_issues_detected')
        self.assertEqual(result['line_break_issues'], [])
        self.assertTrue(all(b['line_break_status'] == 'no_issues_detected' for b in result['blocks']))

    def test_leading_comma_is_located_after_nested_native_bullet(self):
        pages = self.parse([[('•', 20, 20, 24, 30), ('上层正文。', 28, 20, 120, 30),
                             ('o', 20, 50, 24, 60), ('嵌套技术说明', 28, 50, 140, 60),
                             ('，继续核对。', 28, 65, 140, 75)]])
        result = measure(pages, [frame('parent', '上层正文。', 18, role='bullet'),
                                 frame('nested', '嵌套技术说明，继续核对。', 48, height=30, role='bullet')])
        self.assertEqual(result['body_coverage'], 'matched')
        self.assertEqual(result['blocks'][0]['line_break_status'], 'no_issues_detected')
        self.assertEqual(result['line_break_status'], 'issues_detected')
        issues = result['line_break_issues']
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0], dict(block_id='nested', page=1, line_number=2, character='，',
                                        issue_type='forbidden_line_start', line_text='，继续核对。',
                                        line_bounds=dict(xMin=28, yMin=65, xMax=140, yMax=75)))
        self.assertEqual(result['blocks'][1]['line_break_issues'], issues)
        self.assertEqual(result['pages'][0]['line_break_issues'], issues)

    def test_opening_bracket_only_flags_a_nonfinal_visual_line(self):
        pages = self.parse([[('虚构条件（', 25, 40, 120, 50), ('完整说明）。', 25, 55, 130, 65),
                             ('保留原文（', 25, 90, 120, 100)]])
        result = measure(pages, [frame('wrapped', '虚构条件（完整说明）。', 38, height=30),
                                 frame('final', '保留原文（', 88)])
        self.assertEqual([(i['block_id'], i['line_number'], i['character'], i['issue_type'])
                          for i in result['line_break_issues']],
                         [('wrapped', 1, '（', 'forbidden_line_end')])
        self.assertEqual(result['blocks'][1]['line_break_status'], 'no_issues_detected')

    def test_attached_bullet_does_not_hide_forbidden_first_body_character(self):
        pages = self.parse([[('◦，虚构正文', 25, 40, 150, 50)]])
        result = measure(pages, [frame('a', '，虚构正文', 38, role='bullet')])
        self.assertEqual(result['blocks'][0]['match_status'], 'matched')
        self.assertEqual(result['line_break_issues'][0]['character'], '，')
        self.assertEqual(result['line_break_issues'][0]['line_text'], '，虚构正文')

    def test_unverified_text_and_unassigned_words_cannot_clear_line_breaks(self):
        pages = self.parse([[('，未匹配正文', 25, 40, 150, 50)]])
        unknown = measure(pages, [frame('missing', '另一段正文', 38)])
        self.assertEqual(unknown['blocks'][0]['line_break_status'], 'unknown')
        self.assertEqual(unknown['pages'][0]['line_break_status'], 'unknown')
        self.assertEqual(unknown['line_break_status'], 'unknown')
        self.assertEqual(unknown['line_break_issues'], [])
        ambiguous = measure(pages, [frame('a', '，未匹配正文', 38), frame('b', '，未匹配正文', 38)])
        self.assertEqual(ambiguous['line_break_status'], 'unknown')
        self.assertEqual(ambiguous['line_break_issues'], [])
        partial = self.parse([[('已匹配正文。', 25, 40, 150, 50), ('，遗漏正文', 25, 90, 150, 100)]])
        result = measure(partial, [frame('known', '已匹配正文。', 38)])
        self.assertEqual(result['blocks'][0]['line_break_status'], 'no_issues_detected')
        self.assertEqual(result['line_break_status'], 'unknown')

    def test_line_break_checks_use_merged_lines_and_ignore_nonbody(self):
        pages = self.parse([[('，标题', 25, 20, 100, 30), ('接口返回', 25, 50, 100, 60),
                             ('“', 102, 50.3, 112, 60.3), ('值”', 115, 50, 140, 60),
                             ('”后续正文', 25, 65, 140, 75)]])
        result = measure(pages, [frame('heading', '，标题', 18, role='section'),
                                 frame('body', '接口返回“值””后续正文', 48, height=30)])
        self.assertEqual(result['blocks'][0]['line_break_status'], 'not_applicable')
        self.assertEqual(result['blocks'][1]['actual_line_count'], 2)
        self.assertEqual([(i['line_number'], i['character'], i['issue_type'])
                          for i in result['line_break_issues']], [(2, '”', 'forbidden_line_start')])

    def test_inter_project_spacing_and_intervening_heading_are_not_body_gaps(self):
        pages = self.parse([[('甲项目正文', 25, 40, 130, 50), ('乙项目正文', 25, 80, 130, 90),
                             ('新栏目', 25, 110, 80, 120), ('新栏目正文', 25, 140, 130, 150)]])
        frames = [frame('a', '甲项目正文', 38, group_id='a'), frame('b', '乙项目正文', 78, group_id='b'),
                  frame('heading', '新栏目', 108, role='section'), frame('c', '新栏目正文', 138, group_id='c')]
        result = measure(pages, frames)
        self.assertEqual(len(result['adjacent_body_gaps']), 1)
        self.assertEqual(result['adjacent_body_gaps'][0]['status'], 'between_groups')

    def test_unassigned_content_and_nonexistent_manifest_page_are_unknown(self):
        pages = self.parse([[('项目正文', 25, 80, 100, 90), ('未归属文字', 25, 150, 180, 160)]])
        result = measure(pages, [frame('body', '项目正文', 78)])
        self.assertEqual(result['pages'][0]['unassigned_word_count'], 1)
        self.assertEqual(result['body_coverage'], 'unknown')
        missing_page = measure(pages, [frame('body', '项目正文', 78),
                                      frame('missing', '第三页正文', 78, page=3)])
        self.assertEqual(missing_page['body_coverage'], 'unknown')

    def test_manifest_validation_and_file_shapes(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'manifest.json'
            path.write_text('{"blocks": []}', encoding='utf-8')
            self.assertEqual(read_manifest(path), [])
            path.write_text('[]', encoding='utf-8')
            self.assertEqual(read_manifest(path), [])
        pages = self.parse([[('示例', 25, 40, 70, 50)]])
        with self.assertRaisesRegex(ValueError, 'Duplicate'):
            measure(pages, [frame('same', '示例', 38), frame('same', '示例', 38)])
        with self.assertRaisesRegex(ValueError, 'starting at 1'):
            measure(pages, [frame('a', '示例', 38, page=0)])
        with self.assertRaisesRegex(ValueError, 'must be finite'):
            measure(pages, [frame('a', '示例', float('nan'))])


if __name__ == '__main__':
    unittest.main()
