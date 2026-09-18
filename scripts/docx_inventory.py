#!/usr/bin/env python3
"""Read-only DOCX content inventory; comparison is not semantic or visual QA."""

import argparse
from collections import Counter
import hashlib
import io
import json
from pathlib import Path
import posixpath
from urllib.parse import unquote
import xml.etree.ElementTree as ET
import zipfile


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
}
SUPPORTED = set(NS.values()) | {
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/word/2010/wordml",
    "http://schemas.microsoft.com/office/word/2012/wordml",
}
STORIES = {"document": "body", "hdr": "header", "ftr": "footer",
           "footnotes": "footnote", "endnotes": "endnote"}
NOTE_SEPARATORS = {"separator", "continuationSeparator", "continuationNotice"}


class InventoryError(Exception):
    pass


def q(prefix, name):
    return "{" + NS[prefix] + "}" + name


def local(tag):
    return tag.rsplit("}", 1)[-1]


def parse_xml(data, part):
    """Retain in-scope prefix mappings so Choice/@Requires can be resolved."""
    scopes, stack, pending = {}, [{}], []
    try:
        parser = ET.iterparse(io.BytesIO(data), events=("start", "end", "start-ns"))
        for event, item in parser:
            if event == "start-ns":
                pending.append(item)
            elif event == "start":
                scope = dict(stack[-1])
                scope.update(pending)
                pending.clear()
                scopes[id(item)] = scope
                stack.append(scope)
            else:
                stack.pop()
        return parser.root, scopes
    except ET.ParseError as exc:
        raise InventoryError(f"Invalid XML in {part}: {exc}") from exc


def accepted_view(root, scopes, part, decisions):
    """Select one MC representation and the accepted revision view in memory."""
    def visit(element):
        if element.tag in {q("w", "del"), q("w", "moveFrom"), q("w", "delText")}:
            return []
        if element.tag in {q("w", "footnote"), q("w", "endnote")}:
            if element.get(q("w", "type")) in NOTE_SEPARATORS:
                return []
        if element.tag == q("mc", "AlternateContent"):
            selected = None
            for choice in element.findall(q("mc", "Choice")):
                requirements = choice.get("Requires", "").split()
                mapping = scopes[id(choice)]
                if requirements and all(mapping.get(p) in SUPPORTED for p in requirements):
                    selected = choice
                    break
            if selected is None:
                selected = element.find(q("mc", "Fallback"))
            if selected is None:
                raise InventoryError(f"{part}: AlternateContent has no supported Choice or Fallback")
            decisions.append({"part": part, "selected": local(selected.tag),
                              "requires": selected.get("Requires")})
            return [out for child in selected for out in visit(child)]
        copy = ET.Element(element.tag, element.attrib)
        copy.text = element.text
        for child in element:
            copy.extend(visit(child))
        return [copy]
    return visit(root)[0]


def text_of(element):
    """Inline logical text; nested paragraphs/textboxes belong to their own records."""
    output = []

    def visit(node, first=False):
        if node.tag in {q("w", x) for x in ("pPr", "rPr", "sdtPr", "sdtEndPr")}:
            return
        if not first and node.tag in {q("w", "p"), q("w", "txbxContent"), q("v", "textbox")}:
            return
        if node.tag in {q("w", "t"), q("m", "t")}:
            output.append(node.text or "")
        elif node.tag in {q("w", "tab"), q("w", "ptab")}:
            output.append("\t")
        elif node.tag in {q("w", "br"), q("w", "cr")}:
            output.append("\n")
        elif node.tag == q("w", "sym"):
            # A symbol's meaning depends on its font; never silently treat it as ordinary Unicode.
            output.append("⟦sym:" + node.get(q("w", "font"), "") + ":"
                          + node.get(q("w", "char"), "").upper() + "⟧")
        elif node.tag == q("w", "noBreakHyphen"):
            output.append("\u2011")
        elif node.tag == q("w", "softHyphen"):
            output.append("\u00ad")
        elif node.tag == q("m", "chr"):
            # Explicit OMML operator characters can live in properties, outside m:t.
            output.append(node.get(q("m", "val"), ""))
        for child in node:
            visit(child)

    visit(element, True)
    return "".join(output)


def relationships(archive, part):
    rel_path = (posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
                if part else "_rels/.rels")
    if rel_path not in archive.namelist():
        return {}
    root, _ = parse_xml(archive.read(rel_path), rel_path)
    result = {}
    for rel in root:
        rid, target = rel.get("Id"), rel.get("Target", "")
        if not rid or rid in result:
            raise InventoryError(f"{rel_path}: missing or duplicate relationship Id {rid!r}")
        external = rel.get("TargetMode") == "External"
        resolved = target if external else posixpath.normpath(
            posixpath.join(posixpath.dirname(part), unquote(target))).lstrip("/")
        result[rid] = {"id": rid, "target": target, "resolved_target": resolved,
                       "external": external, "type": rel.get("Type", "")}
    return result


def summarize_structure(result):
    """Describe logical paragraph ownership, not rendered lines or layout quality.

    Indices refer to the existing, zero-based ``paragraphs`` list. Keep that
    list unchanged for compatibility; whitespace-only entries are excluded
    from these nonempty-paragraph statistics.
    """
    textboxes = {item["textbox_id"]: item for item in result["textboxes"]}
    line_break_indices, body_indices, ordinary_body_indices = [], [], []
    for index, paragraph in enumerate(result["paragraphs"]):
        if not paragraph["text"].strip():
            continue
        context = paragraph["context"]
        has_line_break = "\n" in paragraph["text"]
        if has_line_break:
            line_break_indices.append(index)
        textbox_id = context["textbox_id"]
        if textbox_id is not None:
            textbox = textboxes[textbox_id]
            textbox["paragraph_indices"].append(index)
            if has_line_break:
                textbox["line_break_paragraph_indices"].append(index)
        elif context["story"] == "body":
            body_indices.append(index)
            if not context["in_table"]:
                ordinary_body_indices.append(index)

    for textbox in result["textboxes"]:
        textbox["nonempty_paragraph_count"] = len(textbox["paragraph_indices"])
        textbox["line_break_paragraph_count"] = len(textbox["line_break_paragraph_indices"])
    multiple = [item["textbox_id"] for item in result["textboxes"]
                if item["nonempty_paragraph_count"] > 1]
    return {
        "paragraph_index_base": 0,
        "paragraph_index_source": "paragraphs",
        "textbox_count": len(textboxes),
        "nonempty_textbox_count": sum(bool(item["paragraph_indices"])
                                      for item in result["textboxes"]),
        "textbox_nonempty_paragraph_count": sum(item["nonempty_paragraph_count"]
                                                for item in result["textboxes"]),
        "multi_paragraph_textbox_count": len(multiple),
        "multi_paragraph_textbox_ids": multiple,
        "line_break_paragraph_count": len(line_break_indices),
        "line_break_paragraph_indices": line_break_indices,
        "body_nonempty_paragraph_count": len(body_indices),
        "body_paragraph_indices": body_indices,
        "ordinary_body_nonempty_paragraph_count": len(ordinary_body_indices),
        "ordinary_body_paragraph_indices": ordinary_body_indices,
        "limitations": [
            "Counts use non-whitespace inline text in logical w:p paragraphs, after MC selection and accepted-view filtering.",
            "Explicit w:br/w:cr stays inside its paragraph; automatic line wrapping is not a new paragraph.",
            "Body counts exclude textboxes; ordinary body also excludes tables. Headers, footers, and notes are excluded from both.",
            "Each textbox ID identifies one w:txbxContent in this inventory, not a persistent OOXML object ID.",
            "These counts locate shared textboxes and explicit breaks; they do not establish independent movability, visibility, or layout quality.",
        ],
    }


def inventory(filename):
    path = Path(filename).expanduser().resolve()
    raw = path.read_bytes()
    result = {"path": str(path), "sha256": hashlib.sha256(raw).hexdigest(),
              "paragraphs": [], "hyperlinks": [], "images": [], "anchors": [],
              "vml_shapes": [], "math": [], "alternate_content": [], "warnings": [],
              "textboxes": []}
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        names = set(archive.namelist())
        package_rels = relationships(archive, "")
        mains = [v["resolved_target"] for v in package_rels.values()
                 if v["type"].endswith("/officeDocument") and not v["external"]]
        main = mains[0] if len(mains) == 1 else "word/document.xml"
        if main not in names:
            raise InventoryError("DOCX has no readable main document part")
        # Follow story relationships, rather than every orphaned XML file in the ZIP.
        pending, seen = [main], set()
        while pending:
            part = pending.pop(0)
            if part in seen:
                continue
            seen.add(part)
            if part not in names:
                raise InventoryError(f"Missing referenced story part: {part}")
            root, scopes = parse_xml(archive.read(part), part)
            story = STORIES.get(local(root.tag))
            if story is None or not root.tag.startswith("{" + NS["w"] + "}"):
                raise InventoryError(f"Unsupported story root in {part}: {root.tag}")
            rels = relationships(archive, part)
            root = accepted_view(root, scopes, part, result["alternate_content"])
            for rel in rels.values():
                if not rel["external"] and rel["type"].rsplit("/", 1)[-1] in {
                    "header", "footer", "footnotes", "endnotes"
                }:
                    pending.append(rel["resolved_target"])

            def resolve(rid):
                if rid not in rels:
                    result["warnings"].append(f"{part}: unresolved relationship {rid!r}")
                    return {"id": rid, "target": None, "resolved_target": None}
                rel = rels[rid]
                if not rel["external"] and rel["resolved_target"] not in names:
                    result["warnings"].append(f"{part}: missing target {rel['resolved_target']}")
                return dict(rel)

            textbox_number = 0

            def walk(node, context):
                nonlocal textbox_number
                context = dict(context)
                if node.tag in {q("w", "footnote"), q("w", "endnote")}:
                    context["note_id"] = node.get(q("w", "id"))
                if node.tag == q("w", "tbl"):
                    context["in_table"] = True
                if node.tag == q("wp", "anchor") or node.tag == q("wp", "inline"):
                    dp = node.find(q("wp", "docPr"))
                    context["shape_id"] = dp.get("id") if dp is not None else None
                if node.tag == q("w", "txbxContent"):
                    # Grouped drawings can share shape_id. Identify the actual
                    # content node after branch/revision selection instead.
                    textbox_number += 1
                    parent_textbox_id = context["textbox_id"]
                    context["textbox_id"] = f"{part}#textbox-{textbox_number}"
                    context["container"] = "textbox"
                    result["textboxes"].append({
                        "part": part, "textbox_id": context["textbox_id"],
                        "parent_textbox_id": parent_textbox_id,
                        "context": dict(context), "paragraph_indices": [],
                        "line_break_paragraph_indices": [],
                    })
                base = {"part": part, "context": dict(context)}
                if node.tag == q("w", "p"):
                    value = text_of(node)
                    if value:
                        result["paragraphs"].append({**base, "text": value})
                elif node.tag == q("w", "hyperlink"):
                    rid = node.get(q("r", "id"))
                    result["hyperlinks"].append({**base, "text": text_of(node),
                        "bookmark": node.get(q("w", "anchor")),
                        "relationship": resolve(rid) if rid else None})
                elif node.tag in {q("a", "blip"), q("v", "imagedata")} or local(node.tag) == "svgBlip":
                    for attribute in (q("r", "embed"), q("r", "link"), q("r", "id"), q("o", "relid")):
                        if node.get(attribute):
                            result["images"].append({**base, "reference_kind": local(attribute),
                                "relationship": resolve(node.get(attribute))})
                    if node.tag == q("v", "imagedata") and node.get("src"):
                        result["images"].append({**base, "src": node.get("src")})
                elif node.tag == q("wp", "anchor"):
                    dp, extent = node.find(q("wp", "docPr")), node.find(q("wp", "extent"))
                    entry = {**base, "attributes": dict(node.attrib),
                             "doc_properties": dict(dp.attrib) if dp is not None else {},
                             "extent_emu": dict(extent.attrib) if extent is not None else {}}
                    for axis in ("H", "V"):
                        position = node.find(q("wp", "position" + axis))
                        if position is not None:
                            entry["position_" + axis.lower()] = {
                                "relative_from": position.get("relativeFrom"),
                                "values": {local(x.tag): x.text for x in position}}
                    result["anchors"].append(entry)
                elif node.tag in {q("v", x) for x in ("shape", "rect", "roundrect", "oval", "line", "group")}:
                    context["shape_id"] = node.get("id")
                    result["vml_shapes"].append({**base, "id": node.get("id"),
                                                  "style": node.get("style", "")})
                elif node.tag == q("m", "oMath"):
                    canonical = ET.canonicalize(ET.tostring(node, encoding="unicode"), rewrite_prefixes=True)
                    result["math"].append({**base, "text": text_of(node),
                        "structure_sha256": hashlib.sha256(canonical.encode()).hexdigest()})
                elif node.tag == q("w", "altChunk"):
                    result["warnings"].append(f"{part}: altChunk text is not expanded; inspect it separately")
                for child in node:
                    walk(child, context)

            walk(root, {"story": story, "container": "body", "in_table": False,
                        "textbox_id": None})
    result["structure_summary"] = summarize_structure(result)
    result["counts"] = {key: len(result[key]) for key in (
        "paragraphs", "hyperlinks", "images", "anchors", "vml_shapes", "math", "textboxes")}
    result["warnings"] = sorted(set(result["warnings"]))
    return result


def compare(before, after):
    normalize = lambda value: " ".join(value.split())
    old = Counter(normalize(p["text"]) for p in before["paragraphs"] if normalize(p["text"]))
    new = Counter(normalize(p["text"]) for p in after["paragraphs"] if normalize(p["text"]))
    chars = lambda inv: Counter(c for p in inv["paragraphs"] for c in p["text"] if not c.isspace())
    old_chars, new_chars = chars(before), chars(after)
    math_signatures = lambda inv: Counter((item["structure_sha256"], item["text"])
                                          for item in inv["math"])
    old_math, new_math = math_signatures(before), math_signatures(after)
    math_records = lambda counts: [{"structure_sha256": key[0], "text": key[1], "count": count}
                                  for key, count in sorted(counts.items())]
    records = lambda counts: [{"text": text, "count": count} for text, count in sorted(counts.items())]
    return {
        "paragraphs_removed_or_regrouped": records(old - new),
        "paragraphs_added_or_regrouped": records(new - old),
        "characters_with_lower_global_count": dict(sorted((old_chars - new_chars).items())),
        "characters_with_higher_global_count": dict(sorted((new_chars - old_chars).items())),
        "math_count_before": len(before["math"]), "math_count_after": len(after["math"]),
        "math_structures_removed_or_changed": math_records(old_math - new_math),
        "math_structures_added_or_changed": math_records(new_math - old_math),
        "limitations": [
            "Paragraph differences can be regrouping or editing, not missing content; order is ignored.",
            "Character counts are clues only: replacements, symbols, and punctuation can change counts.",
            "Math structure hashes also change when formula formatting changes; inspect differences.",
            "Equal inventories do not prove meaning, reading order, glyph display, visibility, or layout.",
            "This is an accepted-revision, supported-MC-branch view; it is not a renderer or DOCX validator.",
        ],
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="DOCX to inventory (never modified)")
    parser.add_argument("--compare", metavar="OTHER_DOCX", help="compare input to a second DOCX")
    args = parser.parse_args()
    try:
        before = inventory(args.input)
        result = before
        if args.compare:
            after = inventory(args.compare)
            result = {"before": before, "after": after, "comparison": compare(before, after)}
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except (InventoryError, OSError, zipfile.BadZipFile, KeyError) as exc:
        parser.exit(2, f"docx_inventory: {exc}\n")


if __name__ == "__main__":
    main()
