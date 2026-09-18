#!/usr/bin/env python3
"""Measure rendered text; never treat DOCX frame estimates as rendered evidence.

Input: Poppler ``pdftotext -bbox-layout input.pdf output.xhtml`` and an optional
JSON array (or {"blocks": [...]}) of frames. Frame coordinates are PDF points,
pages start at 1, and each frame has id, page, x, y, width, height, role, text.
Optional group_id identifies one project; intentional_blank excludes a filling
slot from substantive-content density, while still measuring its spacing.

This read-only tool reports measurements and uncertainty, not visual approval.
No manifest means body classification is unknown. Whitespace-only text, images,
rules and frame bottoms never count as the last line of substantive content.
Chinese line-break diagnostics inspect matched body lines only, preserving the
extracted punctuation. They locate likely violations, not visual approval;
unknown text coverage cannot establish that line breaks are clear.
"""

import argparse
import json
import math
from pathlib import Path
import re
import unicodedata
import xml.etree.ElementTree as ET


BODY_ROLES = {"body", "bullet", "background", "owner", "education_detail",
              "education_placeholder", "citation"}
BULLETS = {"•", "●", "▪", "◦", "‧", "\uf09f", "\uf0b7"}
# Deliberately exclude ASCII punctuation and ambiguous straight quotes: this is
# a conservative Chinese punctuation check, not English word-wrap validation.
FORBIDDEN_LINE_START = set("，。、；：？！）］｝】〕〉》」』”’〗〙〛〞〟")
FORBIDDEN_LINE_END = set("（［｛【〔〈《「『“‘〖〘〚〝")


def _name(node):
    return node.tag.rsplit("}", 1)[-1]


def _number(value, label):
    if isinstance(value, bool):
        raise ValueError(f"{label} must be a finite number")
    value = float(value)
    if not math.isfinite(value):
        raise ValueError(f"{label} must be finite")
    return value


def _text_key(text):
    text = unicodedata.normalize("NFKC", text).strip()
    while text and text[0] in BULLETS:
        text = text[1:].lstrip()
    return re.sub(r"\s+", "", text)


def read_bbox(path):
    """Read actual text bounds from Poppler XHTML using only the standard library."""
    root = ET.parse(path).getroot()
    pages = []
    for page_number, node in enumerate((n for n in root.iter() if _name(n) == "page"), 1):
        page = {"page": page_number, "width": _number(node.get("width"), "page width"),
                "height": _number(node.get("height"), "page height"), "lines": []}
        if page["width"] <= 0 or page["height"] <= 0:
            raise ValueError("Page dimensions must be positive")
        for line_number, line in enumerate(n for n in node.iter() if _name(n) == "line"):
            words = []
            for word_number, word in enumerate(n for n in line.iter() if _name(n) == "word"):
                text = "".join(word.itertext())
                if not text.strip():
                    continue
                bounds = [_number(word.get(key), key) for key in ("xMin", "yMin", "xMax", "yMax")]
                if bounds[2] < bounds[0] or bounds[3] < bounds[1]:
                    raise ValueError("Word bounds are inverted")
                words.append({"key": (page_number, line_number, word_number), "text": text,
                              "xMin": bounds[0], "yMin": bounds[1], "xMax": bounds[2], "yMax": bounds[3]})
            if words:
                page["lines"].append({"source_line": line_number, "words": words})
        pages.append(page)
    if not pages:
        raise ValueError("No pages found: use pdftotext -bbox-layout XHTML")
    return pages


def read_manifest(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    blocks = data.get("blocks") if isinstance(data, dict) else data
    if not isinstance(blocks, list):
        raise ValueError("Manifest must be an array or an object containing a blocks array")
    return blocks


def _validated_blocks(blocks):
    result, ids = [], set()
    for index, supplied in enumerate(blocks):
        if not isinstance(supplied, dict):
            raise ValueError("Each manifest block must be an object")
        block = dict(supplied)
        block["id"] = str(block.get("id", f"block-{index + 1}"))
        if block["id"] in ids:
            raise ValueError(f"Duplicate manifest id: {block['id']}")
        ids.add(block["id"])
        page = _number(block.get("page", 1), "block page")
        if page < 1 or page != int(page):
            raise ValueError("Manifest pages must be positive integers starting at 1")
        block["page"] = int(page)
        for key in ("x", "y", "width", "height"):
            block[key] = _number(block.get(key), f"{block['id']}.{key}")
        if block["width"] <= 0 or block["height"] <= 0:
            raise ValueError("Manifest frame dimensions must be positive")
        block["role"] = str(block.get("role", "unknown"))
        if "text" in block and not isinstance(block["text"], str):
            raise ValueError("Manifest text must be a string, including links and excluding image descriptions")
        result.append(block)
    return result


def _bounds(words):
    if not words:
        return None
    return {"xMin": min(w["xMin"] for w in words), "yMin": min(w["yMin"] for w in words),
            "xMax": max(w["xMax"] for w in words), "yMax": max(w["yMax"] for w in words)}


def _within(word, block, tolerance):
    x, y = (word["xMin"] + word["xMax"]) / 2, (word["yMin"] + word["yMax"]) / 2
    return (block["x"] - tolerance <= x <= block["x"] + block["width"] + tolerance
            and block["y"] - tolerance <= y <= block["y"] + block["height"] + tolerance)


def _line_observations(lines, block, tolerance):
    observed = []
    for source in lines:
        words = [w for w in source["words"] if _within(w, block, tolerance)]
        if words:
            words.sort(key=lambda w: w["xMin"])
            observed.append({**_bounds(words), "text": " ".join(w["text"] for w in words),
                             "source_line": source["source_line"], "words": words})
    return _merge_line_fragments(observed)


def _merge_line_fragments(observed):
    observed = sorted(observed, key=lambda line: (line["yMin"], line["xMin"]))
    # Poppler can emit separate <line>s for Chinese/Latin runs and tab fields on
    # one visual baseline. Coalesce strongly overlapping vertical spans before
    # reading left to right, so tiny font-bound differences do not reorder them.
    merged = []
    for line in observed:
        previous = merged[-1] if merged else None
        overlap = min(previous["yMax"], line["yMax"]) - max(previous["yMin"], line["yMin"]) if previous else 0
        shorter = min(previous["yMax"] - previous["yMin"], line["yMax"] - line["yMin"]) if previous else 0
        if previous and shorter > 0 and overlap / shorter >= .65:
            words = sorted(previous["words"] + line["words"], key=lambda w: w["xMin"])
            merged[-1] = {**_bounds(words), "text": " ".join(w["text"] for w in words),
                          "source_line": previous["source_line"], "words": words}
        else:
            merged.append(line)
    return merged


def _select_text(lines, expected):
    """Narrow broad frame geometry to one complete, word-aligned text sequence."""
    if expected is None or not _text_key(expected):
        return lines, None
    words = [word for line in lines for word in line["words"]]
    keys = [_text_key(word["text"]) for word in words]
    target = _text_key(expected)
    matches = []
    for start in range(len(words)):
        if not keys[start]:
            continue
        joined = ""
        for end in range(start, len(words)):
            joined += keys[end]
            if joined == target:
                matches.append((start, end + 1))
                break
            if not target.startswith(joined):
                break
    if len(matches) != 1:
        return lines, len(matches)
    start, end = matches[0]
    selected = {word["key"] for word in words[start:end]}
    narrowed = []
    for line in lines:
        kept = [word for word in line["words"] if word["key"] in selected]
        if kept:
            narrowed.append({**_bounds(kept), "text": " ".join(w["text"] for w in kept),
                             "source_line": line["source_line"], "words": kept})
    return narrowed, 1


def _fit(gap, target):
    if gap is None:
        return "unknown"
    if gap < 0:
        return "past_safe_bottom"
    if gap < target[0]:
        return "too_close_to_safe_bottom"
    if gap > target[1]:
        return "underfilled"
    return "within_target"


def _generated_bullet(word, blocks):
    if word["text"].strip() in BULLETS:
        return True
    # Word's second native bullet level uses a literal Courier "o". Only ignore
    # it before the first matched text line of a bullet, not an arbitrary word o.
    if word["text"].strip() != "o":
        return False
    for block in blocks:
        if block["role"] != "bullet" or block["match_status"] != "matched" or not block["lines"]:
            continue
        line = block["lines"][0]
        if (block["frame"]["x"] - 6 <= word["xMin"] < line["xMin"]
                and word["xMax"] <= line["xMin"] + .5
                and min(word["yMax"], line["yMax"]) > max(word["yMin"], line["yMin"])):
            return True
    return False


def _line_break_issues(lines, block):
    """Locate forbidden edge characters in one fully matched body block.

    line_number is one-based within the block's merged visual lines. Bounds
    locate the whole extracted line, not the punctuation's exact raster pixels.
    Do not normalize width: NFKC would erase fullwidth punctuation distinctions.
    """
    issues = []
    for number, line in enumerate(lines, 1):
        text = line["text"].strip()
        # Separate native markers are already removed by complete-text matching.
        # A marker may share one extracted word with the first body text instead.
        if number == 1:
            while text and text[0] in BULLETS:
                text = text[1:].lstrip()
        if not text:
            continue
        edges = []
        if text[0] in FORBIDDEN_LINE_START:
            edges.append(("forbidden_line_start", text[0]))
        if number < len(lines) and text[-1] in FORBIDDEN_LINE_END:
            edges.append(("forbidden_line_end", text[-1]))
        for issue_type, character in edges:
            issues.append({"block_id": block["id"], "page": block["page"],
                           "line_number": number, "character": character,
                           "issue_type": issue_type, "line_text": text,
                           "line_bounds": {k: line[k] for k in ("xMin", "yMin", "xMax", "yMax")}})
    return issues


def _line_break_status(issues, known, applicable=True):
    # A discovered issue remains useful even if other body text is unverified.
    # An empty list alone never implies that unknown text has passed inspection.
    if issues:
        return "issues_detected"
    if not known:
        return "unknown"
    return "no_issues_detected" if applicable else "not_applicable"


def measure(pages, manifest=None, *, safe_bottom_margin=28.35, target_bottom_gap=(12, 28),
            max_body_gap=11, body_roles=None, geometry_tolerance=1, target_pages=1):
    """Return actual metrics; unknown/ambiguous matches cannot satisfy fill checks."""
    body_roles = set(BODY_ROLES if body_roles is None else body_roles)
    for value, label in ((safe_bottom_margin, "safe bottom margin"), (max_body_gap, "maximum body gap"),
                         (geometry_tolerance, "geometry tolerance"), *[(v, "target gap") for v in target_bottom_gap]):
        if _number(value, label) < 0:
            raise ValueError(f"{label} cannot be negative")
    if len(target_bottom_gap) != 2 or target_bottom_gap[0] > target_bottom_gap[1]:
        raise ValueError("Target bottom gap must be ordered: minimum maximum")
    if not isinstance(target_pages, int) or isinstance(target_pages, bool) or target_pages < 1:
        raise ValueError("Target pages must be a positive integer")
    if any(safe_bottom_margin >= page["height"] for page in pages):
        raise ValueError("Safe bottom margin must leave usable page height")
    frames = _validated_blocks(manifest) if manifest is not None else []
    all_words = {w["key"]: w for p in pages for line in p["lines"] for w in line["words"]}
    owners, observations, match_counts, results = {}, {}, {}, []
    for frame in frames:
        page = next((p for p in pages if p["page"] == frame["page"]), None)
        lines = _line_observations(page["lines"], frame, geometry_tolerance) if page else []
        lines, match_counts[frame["id"]] = _select_text(lines, frame.get("text"))
        observations[frame["id"]] = lines
        for line in lines:
            for word in line["words"]:
                owners.setdefault(word["key"], []).append(frame["id"])
    for frame in frames:
        lines = observations[frame["id"]]
        words = [w for line in lines for w in line["words"]]
        bounds = _bounds(words)
        expected = frame.get("text")
        observed = "\n".join(line["text"] for line in lines)
        reasons = []
        if frame["page"] > len(pages):
            reasons.append("manifest_page_absent_from_pdf")
        if match_counts[frame["id"]] and match_counts[frame["id"]] > 1:
            status = "ambiguous"
            reasons.append("multiple_complete_text_matches_inside_frame")
        elif any(len(owners[w["key"]]) > 1 for w in words):
            status = "ambiguous"
            reasons.append("rendered_words_claimed_by_multiple_frames")
        elif expected is None:
            status = "unknown"
            reasons.append("manifest_text_missing")
        elif not _text_key(expected):
            status = "intentional_blank" if frame.get("intentional_blank") and not words else "unknown"
            if status == "unknown":
                reasons.append("empty_expected_text_or_unexpected_rendered_text")
        elif _text_key(expected) != _text_key(observed):
            status = "unknown"
            reasons.append("rendered_text_differs_or_is_missing")
        else:
            status = "matched"
        eligible = frame["role"] in body_roles
        substantive = eligible and not (frame.get("intentional_blank") or frame["role"] == "education_placeholder")
        line_break_issues = _line_break_issues(lines, frame) if eligible and status == "matched" else []
        line_break_status = (_line_break_status(line_break_issues, status in {"matched", "intentional_blank"},
                                                status == "matched") if eligible else "not_applicable")
        results.append({"id": frame["id"], "page": frame["page"], "role": frame["role"],
                        "group_id": frame.get("group_id"), "eligible_body": eligible,
                        "substantive_body": substantive, "match_status": status, "reasons": reasons,
                        "complete_text_match_count": match_counts[frame["id"]],
                        "frame": {k: frame[k] for k in ("x", "y", "width", "height")},
                        "actual_line_count": len(lines) if status == "matched" else None,
                        "observed_line_count": len(lines), "ink_bounds": bounds,
                        "actual_top_slack": bounds["yMin"] - frame["y"] if bounds else None,
                        "actual_bottom_slack": frame["y"] + frame["height"] - bounds["yMax"] if bounds else None,
                        "unused_height": frame["height"] - (bounds["yMax"] - bounds["yMin"]) if bounds else None,
                        "observed_text": observed,
                        "line_break_status": line_break_status, "line_break_issues": line_break_issues,
                        "lines": [{k: v for k, v in line.items() if k != "words"} for line in lines]})
    report = {"schema_version": 1, "units": "pt", "measurement_source": "rendered_pdf_text_bounds",
              "page_count": len(pages), "target_page_count": target_pages,
              "page_count_matches_target": len(pages) == target_pages,
              "parameters": {"safe_bottom_margin": safe_bottom_margin, "target_bottom_gap": list(target_bottom_gap),
                             "max_body_gap": max_body_gap, "body_roles": sorted(body_roles),
                             "geometry_tolerance": geometry_tolerance},
              "pages": [], "blocks": results, "adjacent_body_gaps": [],
              "limitations": ["Text extraction does not establish font fidelity, visual centering, icon quality or editability.",
                              "Bounds describe extracted glyph boxes, not exact raster ink pixels.",
                              "Chinese line-break issues are location hints from matched body text; inspect rendered glyphs to verify them.",
                              "Unknown or ambiguous matches require inspection; they are not acceptance passes."]}
    for page in pages:
        page_blocks = [b for b in results if b["page"] == page["page"]]
        eligible = [b for b in page_blocks if b["eligible_body"]]
        substantive = [b for b in eligible if b["substantive_body"]]
        unassigned = [w for k, w in all_words.items() if k[0] == page["page"] and k not in owners
                      and not _generated_bullet(w, page_blocks)]
        known = (manifest is not None and bool(eligible) and not unassigned
                 and all(b["match_status"] in {"matched", "intentional_blank"} for b in eligible))
        line_break_issues = [issue for block in eligible for issue in block["line_break_issues"]]
        safe_bottom = page["height"] - safe_bottom_margin
        def last_bottom(blocks):
            return max((b["ink_bounds"]["yMax"] for b in blocks
                        if b["match_status"] == "matched" and b["ink_bounds"]), default=None)
        body_bottom, content_bottom = last_bottom(eligible), last_bottom(substantive)
        body_gap = safe_bottom - body_bottom if body_bottom is not None else None
        content_gap = safe_bottom - content_bottom if content_bottom is not None else None
        page_words = [w for line in page["lines"] for w in line["words"]]
        visual_lines = _merge_line_fragments([{**_bounds(line["words"]), **line} for line in page["lines"]])
        report["pages"].append({"page": page["page"], "width": page["width"], "height": page["height"],
                                "safe_bottom": safe_bottom, "all_text_line_count": len(visual_lines),
                                "pdf_line_fragment_count": len(page["lines"]),
                                "all_text_last_line_bottom": _bounds(page_words)["yMax"] if page_words else None,
                                "body_coverage": "matched" if known else "unknown",
                                "eligible_body_line_count": sum(b["actual_line_count"] or 0 for b in eligible) if known else None,
                                "observed_matched_body_last_line_bottom": body_bottom,
                                "body_last_line_bottom": body_bottom if known else None,
                                "body_bottom_gap": body_gap if known else None,
                                "substantive_last_line_bottom": content_bottom if known else None,
                                "substantive_bottom_gap": content_gap if known else None,
                                "substantive_fill_status": _fit(content_gap, target_bottom_gap) if known else "unknown",
                                "line_break_status": _line_break_status(line_break_issues, known,
                                                                         any(b["match_status"] == "matched" for b in eligible)),
                                "line_break_issues": line_break_issues,
                                "unassigned_word_count": len(unassigned),
                                "unassigned_text": " ".join(w["text"] for w in unassigned)})
        ordered = sorted(page_blocks, key=lambda b: (b["frame"]["y"], b["frame"]["x"]))
        for previous, current in zip(ordered, ordered[1:]):
            if not (previous["eligible_body"] and current["eligible_body"]):
                continue
            same_group = previous["group_id"] == current["group_id"] if previous["group_id"] is not None and current["group_id"] is not None else None
            same_column = min(previous["frame"]["x"] + previous["frame"]["width"], current["frame"]["x"] + current["frame"]["width"]) > max(previous["frame"]["x"], current["frame"]["x"])
            measured = same_column and all(b["match_status"] == "matched" for b in (previous, current))
            gap = current["ink_bounds"]["yMin"] - previous["ink_bounds"]["yMax"] if measured else None
            status = "unknown" if gap is None else "overlap" if gap < 0 else "between_groups" if same_group is False else "wide" if gap > max_body_gap else "within_target"
            report["adjacent_body_gaps"].append({"page": page["page"], "previous_id": previous["id"],
                                                  "next_id": current["id"], "same_group": same_group,
                                                  "visible_gap": gap, "status": status})
    report["all_text_line_count"] = sum(p["all_text_line_count"] for p in report["pages"])
    report["body_coverage"] = "matched" if all(p["body_coverage"] == "matched" for p in report["pages"]) and not any(b["page"] > len(pages) for b in results) else "unknown"
    report["line_break_issues"] = [issue for block in results for issue in block["line_break_issues"]]
    report["line_break_status"] = _line_break_status(report["line_break_issues"], report["body_coverage"] == "matched",
                                                     any(b["eligible_body"] and b["match_status"] == "matched" for b in results))
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("bbox", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--safe-bottom-margin", type=float, default=28.35)
    parser.add_argument("--target-bottom-gap", type=float, nargs=2, default=(12, 28), metavar=("MIN", "MAX"))
    parser.add_argument("--max-body-gap", type=float, default=11)
    parser.add_argument("--geometry-tolerance", type=float, default=1)
    parser.add_argument("--target-pages", type=int, default=1)
    parser.add_argument("--body-roles", help="Comma-separated roles counted as body text")
    args = parser.parse_args()
    try:
        report = measure(read_bbox(args.bbox), read_manifest(args.manifest) if args.manifest else None,
                         safe_bottom_margin=args.safe_bottom_margin, target_bottom_gap=args.target_bottom_gap,
                         max_body_gap=args.max_body_gap, geometry_tolerance=args.geometry_tolerance,
                         target_pages=args.target_pages,
                         body_roles=set(args.body_roles.split(",")) if args.body_roles else None)
    except (OSError, ValueError, ET.ParseError) as exc:
        parser.error(str(exc))
    serialized = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.write_text(serialized, encoding="utf-8")
    else:
        print(serialized, end="")


if __name__ == "__main__":
    main()
