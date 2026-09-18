#!/usr/bin/env python3
"""Standard-library tests using fictional XML fixtures built from scratch."""

import hashlib
import importlib.util
from pathlib import Path
import tempfile
import unittest
import zipfile


SPEC = importlib.util.spec_from_file_location("docx_inventory", Path(__file__).with_name("docx_inventory.py"))
inventory_module = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(inventory_module)
NAMESPACES = " ".join(f'xmlns:{prefix}="{uri}"' for prefix, uri in inventory_module.NS.items())


def paragraph(text="虚构示例大学", extra=""):
    return f"<w:p><w:r><w:t>{text}</w:t>{extra}</w:r></w:p>"


def transform(width=6000000, height=330200, x=0, group=False):
    children = f'<a:chOff x="0" y="0"/><a:chExt cx="{width}" cy="{height}"/>' if group else ""
    return f'<a:xfrm><a:off x="{x}" y="0"/><a:ext cx="{width}" cy="{height}"/>{children}</a:xfrm>'


def shape(name, content=None, fill="noFill", anchor="ctr", width=6000000, x=0):
    text = f"<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>" if content is not None else ""
    fill_xml = f"<a:{fill}/>" if fill else ""
    return (f'<wps:wsp><wps:cNvPr id="2" name="{name}"/><wps:cNvSpPr/>'
            f'<wps:spPr>{transform(width=width, x=x)}<a:prstGeom prst="rect"/>{fill_xml}</wps:spPr>'
            f'{text}<wps:bodyPr anchor="{anchor}"/></wps:wsp>')


def picture(name="institution-1-icon"):
    return (f'<pic:pic><pic:nvPicPr><pic:cNvPr id="3" name="{name}"/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdIcon"/></pic:blipFill>'
            f'<pic:spPr>{transform(width=228600, height=228600)}</pic:spPr></pic:pic>')


def group(name, children, nested=False):
    tag = "grpSp" if nested else "wgp"
    return (f'<wpg:{tag}><wpg:cNvPr id="1" name="{name}"/><wpg:cNvGrpSpPr/>'
            f'<wpg:grpSpPr>{transform(group=True)}</wpg:grpSpPr>{children}</wpg:{tag}>')


def drawing(content, name="example-1", inline=False):
    kind = "inline" if inline else "anchor"
    return (f'<w:drawing><wp:{kind}><wp:docPr id="1" name="{name}"/>'
            f'<a:graphic><a:graphicData>{content}</a:graphicData></a:graphic></wp:{kind}></w:drawing>')


def institution(include_icon=True, text_fill="noFill", anchor="ctr", inline_icon=False):
    content = paragraph(extra=drawing(picture(), inline=True) if inline_icon else "")
    return group("institution-1", shape("institution-1-background", fill="solidFill")
                 + (picture() if include_icon else "")
                 + shape("institution-1-text", content, fill=text_fill, anchor=anchor))


class DrawingStructureTests(unittest.TestCase):
    def inspect(self, content, extra_parts=None):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "fictional.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", f'<w:document {NAMESPACES}><w:body>{content}</w:body></w:document>')
                archive.writestr("word/_rels/document.xml.rels",
                                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                 '<Relationship Id="rIdIcon" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fictional.svg"/>'
                                 '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/fictional" TargetMode="External"/>'
                                 '</Relationships>')
                archive.writestr("word/media/fictional.svg", '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20"/></svg>')
                for name, value in (extra_parts or {}).items():
                    archive.writestr(name, value)
            before = path.read_bytes()
            result = inventory_module.inventory(path)
            self.assertEqual(before, path.read_bytes())
            self.assertEqual(hashlib.sha256(before).hexdigest(), result["sha256"])
            return result

    def test_three_independent_institution_objects(self):
        result = self.inspect(drawing(institution()))
        item = result["native_groups"][0]
        self.assertEqual("institution", item["role"])
        self.assertEqual(3, len(item["child_object_ids"]))
        self.assertTrue(all(item["structure_observations"].values()))
        children = [child for child in result["drawing_objects"] if child["parent_group_id"] == item["object_id"]]
        self.assertEqual(["background", "icon", "text"], [child["role"] for child in children])
        self.assertEqual("group", children[0]["geometry"]["coordinate_space"])
        self.assertEqual(6000000, children[0]["geometry"]["extent_emu"]["cx"])
        self.assertEqual("ctr", children[2]["vertical_anchor"])
        self.assertEqual("noFill", children[2]["fill"]["kind"])
        self.assertEqual("rIdIcon", result["images"][0]["relationship"]["id"])
        self.assertNotIn("passed", item)

    def test_missing_icon_is_reported_without_acceptance_verdict(self):
        result = self.inspect(drawing(institution(include_icon=False)))
        observations = result["native_groups"][0]["structure_observations"]
        self.assertTrue(observations["has_separate_background_and_text"])
        self.assertFalse(observations["has_independent_picture"])
        self.assertNotIn("passed", result["drawing_structure_summary"])

    def test_names_do_not_override_actual_fill_anchor_or_inline_picture(self):
        result = self.inspect(drawing(institution(include_icon=False, text_fill="solidFill", anchor="t", inline_icon=True)))
        item = result["native_groups"][0]
        observations = item["structure_observations"]
        self.assertFalse(observations["text_has_explicit_no_fill"])
        self.assertFalse(observations["text_has_center_anchor"])
        self.assertFalse(observations["text_has_no_inline_pictures"])
        self.assertFalse(observations["has_independent_picture"])
        self.assertEqual(2, len(item["child_object_ids"]), "Nested inline pictures are not independent group children")

    def test_unspecified_fill_does_not_mean_transparency(self):
        result = self.inspect(drawing(institution(text_fill=None)))
        self.assertFalse(result["native_groups"][0]["structure_observations"]["text_has_explicit_no_fill"])

    def test_misnamed_picture_is_not_a_native_textbox(self):
        children = shape("institution-1-background", fill="solidFill") + picture("institution-1-text")
        result = self.inspect(drawing(group("institution-1", children)))
        self.assertFalse(result["native_groups"][0]["structure_observations"]["has_separate_background_and_text"])
        self.assertEqual("blipFill", result["drawing_objects"][-1]["fill"]["kind"])

    def test_section_rule_must_span_full_local_width(self):
        title = shape("section-1-title", paragraph("虚构栏目"))
        result = self.inspect(drawing(group("section-1", title + shape("section-1-rule", fill="solidFill"))))
        self.assertTrue(all(result["native_groups"][0]["structure_observations"].values()))
        shifted = self.inspect(drawing(group("section-1", title + shape("section-1-rule", fill="solidFill", x=100))))
        self.assertFalse(shifted["native_groups"][0]["structure_observations"]["rule_spans_group_width"])
        narrow = self.inspect(drawing(group("section-1", title + shape("section-1-rule", fill="solidFill", width=5000000))))
        self.assertFalse(narrow["native_groups"][0]["structure_observations"]["rule_spans_group_width"])

    def test_nested_groups_keep_direct_members_and_local_transforms(self):
        nested = group("institution-2", shape("institution-2-text", paragraph()), nested=True)
        result = self.inspect(drawing(group("section-1", nested)))
        outer, inner = result["native_groups"]
        self.assertEqual([inner["object_id"]], outer["child_object_ids"])
        self.assertEqual(1, len(inner["child_object_ids"]))
        self.assertEqual(outer["object_id"], inner["parent_group_id"])

    def test_selected_mc_and_accepted_revisions_apply_to_objects(self):
        content = ('<mc:AlternateContent><mc:Choice Requires="wps">' + institution()
                   + '</mc:Choice><mc:Fallback>' + institution() + '</mc:Fallback></mc:AlternateContent>')
        content += '<w:del>' + institution() + '</w:del>'
        result = self.inspect(drawing(content))
        self.assertEqual(1, len(result["native_groups"]))
        self.assertEqual(["虚构示例大学"], [p["text"] for p in result["paragraphs"]])

    def test_legacy_paragraph_relationship_and_compare_semantics_remain(self):
        link = '<w:hyperlink r:id="rIdLink"><w:r><w:t>虚构链接</w:t></w:r></w:hyperlink>'
        content = drawing(shape("project-1", '<w:p>' + link + '</w:p>' + paragraph("虚构第二段", '<w:br/><w:t>续行</w:t>')))
        result = self.inspect(content)
        self.assertEqual(["虚构链接", "虚构第二段\n续行"], [p["text"] for p in result["paragraphs"]])
        self.assertEqual(1, result["structure_summary"]["multi_paragraph_textbox_count"])
        self.assertEqual(1, result["structure_summary"]["line_break_paragraph_count"])
        self.assertEqual("https://example.invalid/fictional", result["hyperlinks"][0]["relationship"]["target"])
        comparison = inventory_module.compare(result, result)
        self.assertEqual([], comparison["paragraphs_removed_or_regrouped"])
        self.assertEqual({}, comparison["characters_with_lower_global_count"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
