"""Native, independently editable Word components for GoddessReSUme.

Documents start blank. No personal document or bundled sample part is cloned.
Coordinates and sizes are points; page numbers start at one. Each logical text
unit has one paragraph in one textbox. Sections and institution bands are small
native groups whose text, decoration and pictures remain separate objects.

Height estimation is conservative typography guidance, not a rendering engine.
Explicit page/y coordinates are never silently moved. Render every saved page.
"""

from dataclasses import dataclass, field
import math
from pathlib import Path
import re
import unicodedata

from docx import Document
from docx.image.image import Image
from docx.oxml import OxmlElement
from docx.oxml.ns import nsmap, qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph


nsmap.update({
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
})
ROLES = {"name", "section", "project", "body", "bullet", "citation", "contact",
         "institution", "project_header", "tech_stack", "photo_placeholder"}
SIZES = {"name": 20, "section": 14, "institution": 11, "contact": 11}
MIN_HEIGHTS = {"name": 36, "section": 26, "institution": 26, "project": 24,
               "project_header": 24, "contact": 22, "tech_stack": 22}


def _element(tag, attributes=None, parent=None):
    node = OxmlElement(tag)
    for key, value in (attributes or {}).items():
        node.set(qn(key) if ":" in key else key, str(value))
    if parent is not None:
        parent.append(node)
    return node


def _child(parent, tag, **attributes):
    node = parent.find(qn(tag))
    if node is None:
        node = _element(tag, parent=parent)
    for key, value in attributes.items():
        node.set(qn("w:" + key), str(value))
    return node


def _hex(value):
    value = str(value).lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", value):
        raise ValueError(f"Expected six-digit RGB color, got {value!r}")
    return value


def _number(value, name, *, positive=False):
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise ValueError(f"{name} must be a finite number of points")
    if value < 0 or (positive and value == 0):
        raise ValueError(f"{name} must be {'positive' if positive else 'nonnegative'}")
    return value


def _emu(value):
    return round(value * 12700)


def _pieces(value):
    items = [value] if isinstance(value, (str, dict)) else list(value)
    allowed = {"text", "bold", "italic", "underline", "color", "url", "font_size",
               "icon_path", "icon_height"}
    result = []
    for item in items:
        item = {"text": item} if isinstance(item, str) else dict(item)
        if set(item) - allowed or not isinstance(item.get("text", ""), str):
            raise ValueError(f"Invalid run specification: {item!r}")
        if any(char in item.get("text", "") for char in "\r\n"):
            raise ValueError("One textbox represents one logical paragraph; do not supply line breaks")
        if item.get("url") and not isinstance(item["url"], str):
            raise ValueError("url must be a string")
        if "color" in item:
            _hex(item["color"])
        if "font_size" in item:
            _number(item["font_size"], "font_size", positive=True)
        if "icon_path" in item:
            if item.get("text") or item.get("url"):
                raise ValueError("An icon run cannot also contain text or a URL")
            _number(item.get("icon_height", 11), "icon_height", positive=True)
        result.append(item)
    return result


def _text_width(text, size):
    """Approximate glyph advances only; no claim to a font-engine measurement."""
    units = 0
    for char in text:
        if unicodedata.combining(char):
            continue
        if char == "\t":
            units += 2
        elif char.isspace():
            units += .3
        elif unicodedata.east_asian_width(char) in {"W", "F"}:
            units += 1
        elif char in "ilI.,:;'!|":
            units += .28
        elif char in "MW@%":
            units += .85
        elif char.isupper():
            units += .64
        else:
            units += .51
    return units * size


@dataclass
class Block:
    role: str
    page: int
    anchor: object
    paragraph: Paragraph
    metadata: dict = field(default_factory=dict)

    @property
    def x(self):
        return int(self.anchor.find(qn("wp:positionH")).find(qn("wp:posOffset")).text) / 12700

    @property
    def y(self):
        return int(self.anchor.find(qn("wp:positionV")).find(qn("wp:posOffset")).text) / 12700

    @property
    def width(self):
        return int(self.anchor.find(qn("wp:extent")).get("cx")) / 12700

    @property
    def height(self):
        return int(self.anchor.find(qn("wp:extent")).get("cy")) / 12700

    @property
    def bottom(self):
        return self.y + self.height

    def next_y(self, gap=2):
        return self.bottom + _number(gap, "gap")


class ResumeBuilder:
    """Build from a blank DOCX using explicit typography and native shapes.

    Callers own candidate facts, icon acquisition and theme selection. The
    defaults reproduce the skill's typography, not a fixed candidate's layout.
    ``document`` remains available for metadata and intentional page overrides.
    """

    def __init__(self, *, theme=None, east_asia_font="KaiTi", western_font="Times New Roman",
                 page_width=595.3, page_height=864.55, margin_left=34, margin_right=34,
                 margin_top=28.35, margin_bottom=28.35):
        for value, name in ((page_width, "page_width"), (page_height, "page_height"),
                            (margin_left, "margin_left"), (margin_right, "margin_right"),
                            (margin_top, "margin_top"), (margin_bottom, "margin_bottom")):
            _number(value, name, positive=name.startswith("page"))
        if margin_left + margin_right >= page_width or margin_top + margin_bottom >= page_height:
            raise ValueError("Margins must leave a positive content area")
        self.document = Document()
        section = self.document.sections[0]
        section.page_width, section.page_height = Pt(page_width), Pt(page_height)
        section.left_margin, section.right_margin = Pt(margin_left), Pt(margin_right)
        section.top_margin, section.bottom_margin = Pt(margin_top), Pt(margin_bottom)
        section.header_distance = section.footer_distance = Pt(0)
        self.theme = _hex(theme if theme is not None else "4A5568")
        self.dark = "".join(f"{round(int(self.theme[i:i + 2], 16) * .8):02X}" for i in (0, 2, 4))
        self.east_asia_font, self.western_font = east_asia_font, western_font
        self.blocks, self._pages, self._next_id = [], {}, 1
        self._numbering_id = self._numbering()
        normal = self.document.styles["Normal"]
        normal.font.size = Pt(10.5)
        normal.paragraph_format.space_before = normal.paragraph_format.space_after = Pt(0)
        normal.paragraph_format.line_spacing = 1.2
        fonts = _child(normal.element.get_or_add_rPr(), "w:rFonts")
        fonts.attrib.clear()
        for key in ("ascii", "hAnsi", "cs"):
            fonts.set(qn("w:" + key), self.western_font)
        fonts.set(qn("w:eastAsia"), self.east_asia_font)
        for key in ("author", "last_modified_by", "title", "subject", "keywords", "comments", "category"):
            setattr(self.document.core_properties, key, "")

    @property
    def content_width(self):
        section = self.document.sections[0]
        return section.page_width.pt - section.left_margin.pt - section.right_margin.pt

    @property
    def safe_bottom(self):
        section = self.document.sections[0]
        return section.page_height.pt - section.bottom_margin.pt

    def _id(self):
        value = self._next_id
        self._next_id += 1
        return value

    def _numbering(self):
        root = self.document.part.numbering_part.element
        abstracts = [int(x.get(qn("w:abstractNumId"))) for x in root.findall(qn("w:abstractNum"))]
        numbers = [int(x.get(qn("w:numId"))) for x in root.findall(qn("w:num"))]
        abstract_id, number_id = max(abstracts, default=-1) + 1, max(numbers, default=0) + 1
        abstract = _element("w:abstractNum", {"w:abstractNumId": abstract_id})
        _element("w:multiLevelType", {"w:val": "multilevel"}, abstract)
        for level, char, font, size, indent in ((0, "\uf09f", "Wingdings", 10.5, 7.1),
                                                (1, "o", "Courier New", 7.5, 12.75)):
            lvl = _element("w:lvl", {"w:ilvl": level}, abstract)
            for tag, val in (("w:start", 1), ("w:numFmt", "bullet"), ("w:lvlText", char),
                             ("w:lvlJc", "left")):
                _element(tag, {"w:val": val}, lvl)
            ppr = _element("w:pPr", parent=lvl)
            _element("w:ind", {"w:left": round(indent * 20), "w:hanging": 142}, ppr)
            rpr = _element("w:rPr", parent=lvl)
            _element("w:rFonts", {"w:ascii": font, "w:hAnsi": font, "w:eastAsia": font}, rpr)
            _element("w:sz", {"w:val": round(size * 2)}, rpr)
            _element("w:szCs", {"w:val": round(size * 2)}, rpr)
        # abstractNum entries precede concrete num entries in the part schema.
        first_num = root.find(qn("w:num"))
        root.insert(root.index(first_num) if first_num is not None else len(root), abstract)
        num = _element("w:num", {"w:numId": number_id}, root)
        _element("w:abstractNumId", {"w:val": abstract_id}, num)
        return number_id

    def _page(self, page):
        if isinstance(page, bool) or not isinstance(page, int) or page < 1:
            raise ValueError("page must be a positive, one-based integer")
        for number in range(len(self._pages) + 1, page + 1):
            host = self.document.add_paragraph()
            props = host._p.get_or_add_pPr()
            if number > 1:
                _child(props, "w:pageBreakBefore")
            _child(props, "w:spacing", before=0, after=0, line=20, lineRule="exact")
            self._pages[number] = host
        return self._pages[page]

    def _font(self, run, size, color, bold=False, italic=False, underline=False, baseline=0):
        run.font.size, run.font.color.rgb = Pt(size), RGBColor.from_string(_hex(color))
        run.bold, run.italic, run.underline = bold, italic, underline
        rpr = run._r.get_or_add_rPr()
        fonts = _child(rpr, "w:rFonts")
        fonts.attrib.clear()
        for key in ("ascii", "hAnsi", "cs"):
            fonts.set(qn("w:" + key), self.western_font)
        fonts.set(qn("w:eastAsia"), self.east_asia_font)
        _child(rpr, "w:szCs", val=round(size * 2))
        if baseline:
            _child(rpr, "w:position", val=round(baseline * 2))

    def _image_size(self, path, height):
        _number(height, "icon_height", positive=True)
        image = Image.from_file(str(path))
        return height * image.px_width / image.px_height, height

    def _picture_shape(self, path, height, name, *, x=0, y=0):
        width, height = self._image_size(path, height)
        rid, _ = self.document.part.get_or_add_image(str(path))
        pic = _element("pic:pic")
        props = _element("pic:nvPicPr", parent=pic)
        _element("pic:cNvPr", {"id": self._id(), "name": name}, props)
        _element("pic:cNvPicPr", parent=props)
        fill = _element("pic:blipFill", parent=pic)
        _element("a:blip", {"r:embed": rid}, fill)
        _element("a:fillRect", parent=_element("a:stretch", parent=fill))
        sppr = _element("pic:spPr", parent=pic)
        self._transform(sppr, x, y, width, height)
        _element("a:prstGeom", {"prst": "rect"}, sppr)
        return pic, width, height

    def _picture(self, paragraph, path, height, size, *, baseline=-1):
        pic, width, height = self._picture_shape(path, height, f"inline-icon-{self._next_id}")
        run = paragraph.add_run()
        self._font(run, size, "262626", baseline=baseline)
        drawing = _element("w:drawing", parent=run._r)
        inline = _element("wp:inline", parent=drawing)
        _element("wp:extent", {"cx": _emu(width), "cy": _emu(height)}, inline)
        _element("wp:docPr", {"id": self._id(), "name": "inline-icon"}, inline)
        frame = _element("wp:cNvGraphicFramePr", parent=inline)
        _element("a:graphicFrameLocks", {"noChangeAspect": 1}, frame)
        graphic = _element("a:graphic", parent=inline)
        data = _element("a:graphicData", {"uri": nsmap["pic"]}, graphic)
        data.append(pic)
        return run

    @staticmethod
    def _transform(parent, x, y, width, height, *, group=False):
        transform = _element("a:xfrm", parent=parent)
        _element("a:off", {"x": _emu(x), "y": _emu(y)}, transform)
        _element("a:ext", {"cx": _emu(width), "cy": _emu(height)}, transform)
        if group:
            _element("a:chOff", {"x": 0, "y": 0}, transform)
            _element("a:chExt", {"cx": _emu(width), "cy": _emu(height)}, transform)
        return transform

    def _shape(self, name, width, height, *, paragraph=None, fill=None, alpha=None,
               x=0, y=0, padding_y=2, vertical_anchor="t"):
        shape = _element("wps:wsp")
        _element("wps:cNvPr", {"id": self._id(), "name": name}, shape)
        _element("wps:cNvSpPr", {"txBox": 1} if paragraph is not None else {}, shape)
        props = _element("wps:spPr", parent=shape)
        self._transform(props, x, y, width, height)
        _element("a:avLst", parent=_element("a:prstGeom", {"prst": "rect"}, props))
        if fill:
            color = _element("a:srgbClr", {"val": _hex(fill)}, _element("a:solidFill", parent=props))
            if alpha is not None:
                _element("a:alpha", {"val": alpha}, color)
        else:
            _element("a:noFill", parent=props)
        _element("a:noFill", parent=_element("a:ln", parent=props))
        if paragraph is not None:
            _element("w:txbxContent", parent=_element("wps:txbx", parent=shape)).append(paragraph._p)
            body = _element("wps:bodyPr", {"rot": 0, "vert": "horz", "wrap": "square",
                            "lIns": 0, "rIns": 0, "tIns": _emu(padding_y),
                            "bIns": _emu(padding_y), "anchor": vertical_anchor}, shape)
            _element("a:noAutofit", parent=body)
        else:
            _element("wps:bodyPr", parent=shape)
        return shape

    def _group(self, width, height, children):
        group = _element("wpg:wgp")
        _element("wpg:cNvGrpSpPr", parent=group)
        self._transform(_element("wpg:grpSpPr", parent=group), 0, 0, width, height, group=True)
        for child in children:
            group.append(child)
        return group

    def _paragraph(self, pieces, role, size, color, *, line_spacing=1.2, line_height=None,
                   level=1, indent=0, tabs=None, align=None, baseline=0):
        paragraph = Paragraph(_element("w:p"), self.document._body)
        props = paragraph._p.get_or_add_pPr()
        if role == "bullet":
            numbering = _child(props, "w:numPr")
            _child(numbering, "w:ilvl", val=level)
            _child(numbering, "w:numId", val=self._numbering_id)
        if tabs:
            tab_set = _child(props, "w:tabs")
            for position, mode in tabs:
                _element("w:tab", {"w:val": mode, "w:pos": round(position * 20)}, tab_set)
        _child(props, "w:snapToGrid", val=0)
        _child(props, "w:spacing", before=0, after=0,
               line=round(line_height * 20) if line_height is not None else round(line_spacing * 240),
               lineRule="exact" if line_height is not None else "auto")
        if role == "bullet":
            _child(props, "w:ind", left=142 if level == 0 else 255, hanging=142)
        elif indent:
            _child(props, "w:ind", left=round(indent * 20), hanging=0)
        _child(props, "w:jc", val=align or ("both" if role in {"body", "bullet", "citation"} else "left"))
        for item in pieces:
            run_size = item.get("font_size", size)
            if "icon_path" in item:
                self._picture(paragraph, item["icon_path"], item.get("icon_height", 11), run_size)
                continue
            run = paragraph.add_run(item.get("text", ""))
            self._font(run, run_size, item.get("color", "0000FF" if item.get("url") else color),
                       item.get("bold", role in {"name", "section", "project", "project_header", "institution"}),
                       item.get("italic", role == "citation"), item.get("underline", bool(item.get("url"))),
                       baseline=baseline)
            if item.get("url"):
                link = _element("w:hyperlink", {"r:id": self.document.part.relate_to(item["url"], RT.HYPERLINK, is_external=True)})
                link.append(run._r)
                paragraph._p.append(link)
        return paragraph

    def estimate_height(self, role, runs_or_text, *, width=None, font_size=None,
                        line_spacing=1.2, line_height=None, padding_y=2, level=1):
        """Estimate content height; verify actual line count and clipping in render."""
        width = self.content_width if width is None else _number(width, "width", positive=True)
        size = SIZES.get(role, 10.5) if font_size is None else _number(font_size, "font_size", positive=True)
        _number(line_spacing, "line_spacing", positive=True)
        _number(padding_y, "padding_y")
        if line_height is not None:
            _number(line_height, "line_height", positive=True)
        items = _pieces(runs_or_text)
        total = 0
        max_size = size
        for item in items:
            run_size = item.get("font_size", size)
            max_size = max(max_size, run_size)
            total += (self._image_size(item["icon_path"], item.get("icon_height", 11))[0]
                      if "icon_path" in item else _text_width(item.get("text", ""), run_size))
        usable = width - (7.1 if level == 0 else 12.75) if role == "bullet" else width
        if usable <= 0:
            raise ValueError("Width leaves no room for text")
        lines = max(1, math.ceil(total * 1.04 / usable))
        # KaiTi's rendered line box is appreciably taller than its nominal em.
        # A multilingual seven-line probe measured 17.75 pt baseline advances at
        # 10.5 pt / 1.2 spacing. Reserve font-metric headroom without changing
        # the paragraph's actual 1.2 line-spacing setting or forcing a line count.
        has_cjk = any(unicodedata.east_asian_width(c) in {"W", "F"}
                      for item in items for c in item.get("text", ""))
        metric_factor = 1.45 if has_cjk else 1.25
        advance = line_height if line_height is not None else max_size * line_spacing * metric_factor
        return max(MIN_HEIGHTS.get(role, 0), math.ceil((lines * advance + 2 * padding_y) * 2) / 2)

    def _placement(self, x, y, width, height):
        section = self.document.sections[0]
        x = section.left_margin.pt if x is None else x
        y = section.top_margin.pt if y is None else y
        width = section.page_width.pt - section.right_margin.pt - x if width is None else width
        for value, name in ((x, "x"), (y, "y"), (width, "width"), (height, "height")):
            _number(value, name, positive=name in {"width", "height"})
        return x, y, width, height

    def _attach(self, role, page, x, y, width, height, content, paragraph, identifier, metadata=None):
        anchor = _element("wp:anchor", {"distT": 0, "distB": 0, "distL": 0, "distR": 0,
                          "simplePos": 0, "relativeHeight": 251660000 + identifier,
                          "behindDoc": 0, "locked": 0, "layoutInCell": 1, "allowOverlap": 1})
        _element("wp:simplePos", {"x": 0, "y": 0}, anchor)
        for axis, value in (("H", x), ("V", y)):
            _element("wp:posOffset", parent=_element("wp:position" + axis, {"relativeFrom": "page"}, anchor)).text = str(_emu(value))
        _element("wp:extent", {"cx": _emu(width), "cy": _emu(height)}, anchor)
        _element("wp:effectExtent", {"l": 0, "t": 0, "r": 0, "b": 0}, anchor)
        _element("wp:wrapNone", parent=anchor)
        _element("wp:docPr", {"id": identifier, "name": f"{role}-{identifier}"}, anchor)
        _element("wp:cNvGraphicFramePr", parent=anchor)
        data = _element("a:graphicData", {"uri": nsmap["wpg"] if content.tag == qn("wpg:wgp") else nsmap["wps"]},
                        _element("a:graphic", parent=anchor))
        data.append(content)
        _element("w:drawing", parent=self._page(page).add_run()._r).append(anchor)
        block = Block(role, page, anchor, paragraph, metadata or {})
        self.blocks.append(block)
        return block

    def add_block(self, role, runs_or_text, *, page=1, x=None, y=None, width=None,
                  height=None, font_size=None, color=None, level=1, icon_path=None,
                  icon_height=11, line_spacing=1.2, line_height=None, padding_y=2):
        """One logical paragraph; section/institution dispatch to native groups.

        Text run keys: text, bold, italic, underline, color, url, font_size.
        Image-only run keys: icon_path, icon_height. No manual line breaks.
        The institution shorthand accepts only a plain institution name;
        use add_institution() to supply separate fields, tabs and logo options.
        """
        if role not in ROLES - {"project_header", "photo_placeholder"}:
            raise ValueError(f"Use a supported text role or the dedicated component method, got {role!r}")
        if role == "section":
            return self.add_section(runs_or_text, page=page, x=x, y=y, width=width,
                                    height=26 if height is None else height, font_size=font_size or 14)
        if role == "institution":
            if not isinstance(runs_or_text, str) or any(c in runs_or_text for c in "\t\r\n"):
                raise ValueError("add_block('institution') accepts only a plain name string; use add_institution() for fields")
            return self.add_institution(runs_or_text, logo_path=icon_path, logo_height=icon_height if icon_path else 18,
                                        page=page, x=x, y=y, width=width, height=26 if height is None else height,
                                        font_size=font_size or 11)
        if level not in (0, 1):
            raise ValueError("bullet level must be 0 or 1")
        size = SIZES.get(role, 10.5) if font_size is None else _number(font_size, "font_size", positive=True)
        items = _pieces(runs_or_text)
        if icon_path is not None:
            items = [{"icon_path": icon_path, "icon_height": icon_height}, {"text": " "}] + items
        effective_width = self.document.sections[0].page_width.pt - self.document.sections[0].right_margin.pt - (self.document.sections[0].left_margin.pt if x is None else x) if width is None else width
        if line_height is None and role in {"contact", "tech_stack", "project"}:
            line_height = 18
        estimated_height = self.estimate_height(role, items, width=effective_width, font_size=size,
                                                line_spacing=line_spacing, line_height=line_height,
                                                padding_y=padding_y, level=level)
        x, y, width, height = self._placement(x, y, width, estimated_height if height is None else height)
        color = _hex(color or (self.dark if role == "project" else "262626"))
        paragraph = self._paragraph(items, role, size, color, line_spacing=line_spacing,
                                    line_height=line_height, level=level)
        identifier = self._id()
        shape = self._shape(f"{role}-{identifier}-text", width, height, paragraph=paragraph, padding_y=padding_y)
        return self._attach(role, page, x, y, width, height, shape, paragraph, identifier,
                            {"estimated_height": estimated_height, "height_is_estimate": height == estimated_height})

    def add_section(self, text, *, page=1, x=None, y=None, width=None, height=26,
                    font_size=14, rule_thickness=1.5, rule_y=18.5, padding_y=2):
        """Column-wide native rule grouped with a separate transparent title box."""
        _number(font_size, "font_size", positive=True)
        _number(rule_thickness, "rule_thickness", positive=True)
        _number(rule_y, "rule_y")
        x, y, width, height = self._placement(x, y, width, height)
        if rule_y + rule_thickness > height:
            raise ValueError("Section rule extends below the section group")
        paragraph = self._paragraph(_pieces(text), "section", font_size, self.theme, line_height=18)
        identifier = self._id()
        name = f"section-{identifier}"
        children = [self._shape(name + "-title", width, height, paragraph=paragraph, padding_y=padding_y),
                    self._shape(name + "-rule", width, rule_thickness, y=rule_y, fill=self.theme)]
        return self._attach("section", page, x, y, width, height, self._group(width, height, children),
                            paragraph, identifier)

    def add_institution(self, name, detail="", degree="", dates="", *, logo_path=None,
                        logo_height=18, logo_gap=4, tab_stops=None, page=1, x=None,
                        y=None, width=None, height=26, font_size=11, text_baseline=0,
                        line_height=16, tint_alpha=10000):
        """Group an independent background, optional logo and transparent text.

        Tab offsets are relative to the full band, not the indented first text.
        Logo height and its actual aspect ratio determine the text indentation.
        text_baseline is optional optical compensation in points, default zero.
        """
        fields = (name, detail, degree, dates)
        if any(not isinstance(v, str) or any(c in v for c in "\t\r\n") for v in fields):
            raise ValueError("Institution fields must be single-line strings without tabs")
        _number(logo_gap, "logo_gap")
        _number(font_size, "font_size", positive=True)
        _number(line_height, "line_height", positive=True)
        if not isinstance(text_baseline, (int, float)) or not math.isfinite(text_baseline):
            raise ValueError("text_baseline must be finite")
        if isinstance(tint_alpha, bool) or not isinstance(tint_alpha, int) or not 0 <= tint_alpha <= 100000:
            raise ValueError("tint_alpha is an integer from 0 to 100000")
        x, y, width, height = self._placement(x, y, width, height)
        identifier = self._id()
        component_name = f"institution-{identifier}"
        children = [self._shape(component_name + "-background", width, height, fill=self.theme, alpha=tint_alpha)]
        indent = 0
        if logo_path:
            if logo_height > height:
                raise ValueError("Logo must fit inside the institution band")
            pic, logo_width, _ = self._picture_shape(logo_path, logo_height, component_name + "-icon",
                                                    y=(height - logo_height) / 2)
            children.append(pic)
            indent = logo_width + logo_gap
            if indent >= width:
                raise ValueError("Logo and gap leave no room for institution text")
        stops = tuple(tab_stops) if tab_stops is not None else (width * .37, width * .67, width - 3)
        if len(stops) != 3:
            raise ValueError("tab_stops must have three offsets")
        for stop in stops:
            _number(stop, "tab stop", positive=True)
        if list(stops) != sorted(set(stops)) or stops[-1] > width or stops[0] <= indent:
            raise ValueError("Tab stops must increase within the band and begin after the logo")
        paragraph = self._paragraph(_pieces("\t".join(fields)), "institution", font_size, self.dark,
                                    line_height=line_height, indent=indent,
                                    tabs=[(p, "right" if i == 2 else "left") for i, p in enumerate(stops)],
                                    baseline=text_baseline)
        children.append(self._shape(component_name + "-text", width, height, paragraph=paragraph,
                                    padding_y=0, vertical_anchor="ctr"))
        return self._attach("institution", page, x, y, width, height, self._group(width, height, children),
                            paragraph, identifier, {"logo_expected": bool(logo_path), "font_size": font_size})

    def add_project_header(self, title, *, repo_url=None, repo_label=None, icon_path=None,
                           icon_height=11, page=1, x=None, y=None, width=None, height=24,
                           font_size=10.5, link_font_size=10.5, min_gap=10):
        """Project subtitle and its repository share one line with a right tab.

        Long text is preserved. A width warning in validate() requires render
        review and, if needed, a shorter display label; the URL stays unchanged.
        """
        if not isinstance(title, str) or any(c in title for c in "\t\r\n"):
            raise ValueError("Project title must be a single-line string")
        x, y, width, height = self._placement(x, y, width, height)
        _number(font_size, "font_size", positive=True)
        _number(link_font_size, "link_font_size", positive=True)
        _number(min_gap, "min_gap")
        if repo_url is None and (repo_label is not None or icon_path is not None):
            raise ValueError("A project link label/icon requires repo_url")
        label = repo_label if repo_label is not None else re.sub(r"^https?://", "", repo_url or "")
        if any(c in label for c in "\t\r\n"):
            raise ValueError("Repository display label must be single-line")
        items = [{"text": title}]
        required = _text_width(title, font_size)
        if repo_url:
            items.append({"text": "\t"})
            if icon_path:
                items.append({"icon_path": icon_path, "icon_height": icon_height})
                required += self._image_size(icon_path, icon_height)[0]
            items.append({"text": (" " if icon_path else "") + label, "url": repo_url,
                          "font_size": link_font_size, "bold": False})
            required += _text_width(label, link_font_size) + min_gap
        paragraph = self._paragraph(_pieces(items), "project_header", font_size, self.dark, line_height=18,
                                    tabs=[(width, "right")] if repo_url else None)
        identifier = self._id()
        shape = self._shape(f"project_header-{identifier}-text", width, height, paragraph=paragraph)
        return self._attach("project_header", page, x, y, width, height, shape, paragraph, identifier,
                            {"repository_url": repo_url, "estimated_line_width": required})

    def add_contact(self, text, *, url=None, icon_path=None, **placement):
        return self.add_block("contact", {"text": text, "url": url}, icon_path=icon_path, **placement)

    def add_tech_stack(self, items, *, icon_path=None, label="技术栈：", separator=" · ", **placement):
        """Use supplied project-relevant technologies without proficiency claims."""
        text = items if isinstance(items, str) else separator.join(items)
        runs = [{"text": label, "bold": True, "color": self.dark, "font_size": 11}, {"text": text}]
        return self.add_block("tech_stack", runs, icon_path=icon_path, **placement)

    def add_photo_placeholder(self, label="照片占位", *, page=1, x=None, y=None,
                              width=76.17, height=106.64):
        """Native gray slot, only for fictional examples or requested templates."""
        x, y, width, height = self._placement(x, y, width, height)
        paragraph = self._paragraph(_pieces(label), "photo_placeholder", 11, "777777", align="center")
        identifier = self._id()
        shape = self._shape(f"photo_placeholder-{identifier}-text", width, height, paragraph=paragraph,
                            fill="EEEEEE", padding_y=0, vertical_anchor="ctr")
        return self._attach("photo_placeholder", page, x, y, width, height, shape, paragraph, identifier)

    def next_position(self, block, *, gap=2, required_height=0):
        """Return explicit (page, y); keep a following block within safe margins.

        For a heading + first body unit, pass their combined required height.
        This helper never moves an already-created block or splits paragraphs.
        """
        _number(required_height, "required_height")
        top = self.document.sections[0].top_margin.pt
        if required_height > self.safe_bottom - top:
            raise ValueError("Required block/cluster is taller than the content area")
        y = block.next_y(gap)
        return (block.page + 1, top) if y + required_height > self.safe_bottom else (block.page, y)

    @staticmethod
    def _institution_width_warnings(block):
        """Flag likely tab jumping/field collisions, while preserving all text."""
        fields = block.paragraph.text.split("\t")
        props = block.paragraph._p.get_or_add_pPr()
        tab_set = props.find(qn("w:tabs"))
        tabs = list(tab_set) if tab_set is not None else []
        if len(fields) != 4 or len(tabs) != 3:
            return ["institution fields/tab stops changed; review four-field layout and render"]
        stops = [int(t.get(qn("w:pos"))) / 20 for t in tabs]
        indent_node = props.find(qn("w:ind"))
        indent = int(indent_node.get(qn("w:left"), "0")) / 20 if indent_node is not None else 0
        size = block.metadata.get("font_size", 11)
        widths = [_text_width(field, size) * 1.04 for field in fields]
        slots = (stops[0] - indent, stops[1] - stops[0], stops[2] - stops[1])
        problems = []
        for label, required, available in (("name", widths[0], slots[0]),
                                            ("detail", widths[1], slots[1]),
                                            ("degree/date", widths[2] + widths[3] + (4 if fields[2] and fields[3] else 0), slots[2])):
            if required > available:
                problems.append(f"institution {label} may exceed its tab field ({required:.1f} > {available:.1f} pt); "
                                "adjust field wording or tab widths and render; text is preserved")
        return problems

    def validate(self):
        """Check structural invariants and safe bounds; visual review is separate."""
        section = self.document.sections[0]
        warnings = []
        for block in self.blocks:
            if (block.x < 0 or block.y < 0 or block.x + block.width > section.page_width.pt + .01
                    or block.bottom > self.safe_bottom + .01):
                raise ValueError(f"{block.role} on page {block.page} exceeds page/safe-bottom bounds")
            boxes = list(block.anchor.iter(qn("w:txbxContent")))
            if len(boxes) != 1 or len(boxes[0].findall(qn("w:p"))) != 1:
                raise ValueError(f"{block.role} requires one textbox containing one paragraph")
            if list(boxes[0].iter(qn("w:br"))):
                raise ValueError("Manual line breaks cannot combine logical content units")
            group = block.anchor.find(".//" + qn("wpg:wgp"))
            if block.role in {"section", "institution"}:
                if group is None:
                    raise ValueError(f"{block.role} must be a native group")
                shapes = group.findall(qn("wps:wsp"))
                pictures = group.findall(qn("pic:pic"))
                if len(shapes) != 2 or len(pictures) != (int(block.metadata.get("logo_expected", False)) if block.role == "institution" else 0):
                    raise ValueError(f"{block.role} has missing or unexpected group components")
                text_shape = next((s for s in shapes if s.find(".//" + qn("w:txbxContent")) is not None), None)
                decoration = next((s for s in shapes if s is not text_shape), None)
                if text_shape is None or text_shape.find(qn("wps:spPr")).find(qn("a:noFill")) is None:
                    raise ValueError("Grouped title/institution text must be a transparent separate textbox")
                if decoration.find(".//" + qn("a:solidFill")) is None:
                    raise ValueError("Section rule/institution background requires a visible fill")
                if list(text_shape.iter(qn("pic:pic"))):
                    raise ValueError("Institution logo must be independent of text line metrics")
                if block.role == "institution":
                    body = text_shape.find(qn("wps:bodyPr"))
                    if body.get("anchor") != "ctr" or body.get("tIns") != body.get("bIns"):
                        raise ValueError("Institution text must be vertically centered with symmetric padding")
                    component_name = block.anchor.find(qn("wp:docPr")).get("name")
                    warnings.extend(f"{component_name}: {warning}"
                                    for warning in self._institution_width_warnings(block))
                for child in shapes + pictures:
                    transform = child.find(".//" + qn("a:xfrm"))
                    off, extent = transform.find(qn("a:off")), transform.find(qn("a:ext"))
                    cx, cy, cw, ch = [int(v) / 12700 for v in (off.get("x"), off.get("y"), extent.get("cx"), extent.get("cy"))]
                    if cx < 0 or cy < 0 or cx + cw > block.width + .01 or cy + ch > block.height + .01:
                        raise ValueError("Group child extends outside its component bounds")
                if block.role == "section":
                    ext = decoration.find(".//" + qn("a:ext"))
                    if abs(int(ext.get("cx")) / 12700 - block.width) > .01:
                        raise ValueError("Section rule must span the component's content column")
            if block.role == "project_header" and block.metadata.get("repository_url"):
                tabs = block.paragraph._p.findall(".//" + qn("w:tab"))
                if not any(t.get(qn("w:val")) == "right" for t in tabs):
                    raise ValueError("Project repository requires a right-aligned tab stop")
                if not list(block.paragraph._p.iter(qn("w:hyperlink"))):
                    raise ValueError("Project repository must retain a live hyperlink")
            if block.role == "project_header" and block.metadata.get("estimated_line_width", 0) > block.width:
                warnings.append(f"{block.anchor.find(qn('wp:docPr')).get('name')}: title/link may not fit one line; shorten display text and render")
        seen = set()
        for node in self.document.element.body.iter():
            if node.tag in {qn("wp:docPr"), qn("wps:cNvPr"), qn("pic:cNvPr")}:
                identifier = node.get("id")
                if identifier in seen:
                    raise ValueError(f"Duplicate drawing id {identifier}")
                seen.add(identifier)
            for key in ("r:id", "r:embed", "r:link"):
                rid = node.get(qn(key))
                if rid is not None and rid not in self.document.part.rels:
                    raise ValueError(f"Unresolved document relationship {rid}")
        return {"pages": len(self._pages), "textboxes": len(self.blocks),
                "native_groups": sum(b.role in {"section", "institution"} for b in self.blocks),
                "requires_rendering": True, "measures_text_fit": False, "warnings": warnings}

    def save(self, path):
        """Save after structural validation; render all pages before delivery."""
        path = Path(path).resolve()
        self.validate()
        self.document.save(path)
        return path
