"""Shared helpers for the format-monograph command-line tools."""

from __future__ import annotations

import hashlib
import json
import re
import zipfile
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O_NS = "urn:schemas-microsoft-com:office:office"
V_NS = "urn:schemas-microsoft-com:vml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "m": M_NS, "o": O_NS, "v": V_NS, "r": R_NS}
CONTENT_PART = re.compile(
    r"^word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$"
)

ROLE_STYLE_MAP = {
    "body": "Normal",
    "title": "Title",
    "subtitle": "Subtitle",
    "caption": "Caption",
    "bibliography": "Bibliography",
    "body_text": "Normal",
    "chapter_title": "Heading 1",
    "level_2_section": "Heading 2",
    "level_3_section": "Heading 3",
    "level_4_section": "Heading 4",
    "long_quote": "Quote",
    **{f"heading{i}": f"Heading {i}" for i in range(1, 10)},
}

STYLE_PROPERTIES = {
    "font_name",
    "font_name_ascii",
    "font_name_east_asia",
    "font_name_complex_script",
    "font_size_pt",
    "bold",
    "italic",
    "color_hex",
    "alignment",
    "space_before_pt",
    "space_after_pt",
    "line_spacing",
    "line_spacing_rule",
    "line_spacing_pt",
    "first_line_indent_pt",
    "first_line_indent_chars",
    "left_indent_pt",
    "right_indent_pt",
    "keep_with_next",
    "keep_together",
    "page_break_before",
    "widow_control",
}

SECTION_PROPERTIES = {
    "page_width_mm",
    "page_height_mm",
    "orientation",
    "margin_top_mm",
    "margin_bottom_mm",
    "margin_left_mm",
    "margin_right_mm",
    "gutter_mm",
    "different_first_page_header_footer",
    "odd_and_even_pages_header_footer",
    "page_size_policy",
    "mirror_margins",
    "margin_inner_ratio",
    "margin_outer_ratio",
    "margin_top_ratio",
    "margin_bottom_ratio",
    "header_distance_ratio",
    "footer_distance_ratio",
}

TABLE_PROPERTIES = {
    "table_style",
    "alignment",
    "repeat_header_row",
    "prevent_row_split",
}

FIELD_PROPERTIES = {
    "update_on_open",
    "mark_fields_dirty",
    "convert_explicit_markers",
    "rebuild_heading_numbering",
    "heading_levels",
    "strip_manual_heading_prefixes",
}

EQUATION_PROPERTIES = {
    "require_editable_equations",
    "preserve_editable_objects",
    "block_formula_images",
}

ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "distributed": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
}

TABLE_ALIGNMENTS = {
    "left": WD_TABLE_ALIGNMENT.LEFT,
    "center": WD_TABLE_ALIGNMENT.CENTER,
    "right": WD_TABLE_ALIGNMENT.RIGHT,
}


class FormatMonographError(RuntimeError):
    """Expected user-facing validation or processing error."""


def ensure_docx(path: Path) -> None:
    if path.suffix.lower() != ".docx":
        raise FormatMonographError(f"Expected a .docx file: {path}")
    if not path.is_file():
        raise FormatMonographError(f"DOCX file does not exist: {path}")
    if not zipfile.is_zipfile(path):
        raise FormatMonographError(f"File is not a valid DOCX package: {path}")
    with zipfile.ZipFile(path) as package:
        if "word/document.xml" not in package.namelist():
            raise FormatMonographError(f"DOCX package has no word/document.xml: {path}")


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise FormatMonographError(f"JSON file does not exist: {path}") from exc
    except json.JSONDecodeError as exc:
        raise FormatMonographError(
            f"Invalid JSON at line {exc.lineno}, column {exc.colno}: {exc.msg}"
        ) from exc


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )


def profile_schema_path() -> Path:
    return Path(__file__).resolve().parent.parent / "references" / "format-profile.schema.json"


def _paragraph_text_without_field_results(paragraph: etree._Element) -> str:
    pieces: list[str] = []
    field_stack: list[str] = []

    for element in paragraph.iter():
        if element.tag == f"{{{W_NS}}}fldChar":
            kind = element.get(f"{{{W_NS}}}fldCharType")
            if kind == "begin":
                field_stack.append("instruction")
            elif kind == "separate" and field_stack:
                field_stack[-1] = "result"
            elif kind == "end" and field_stack:
                field_stack.pop()
            continue

        if element.tag in {f"{{{W_NS}}}t", f"{{{W_NS}}}delText"}:
            if any(state == "result" for state in field_stack):
                continue
            if any(parent.tag == f"{{{W_NS}}}fldSimple" for parent in element.iterancestors()):
                continue
            pieces.append(element.text or "")
        elif element.tag == f"{{{W_NS}}}tab":
            pieces.append("\t")
        elif element.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
            pieces.append("\n")

    return "".join(pieces)


def _normalize_derived_paragraph_text(paragraph: etree._Element, value: str) -> str:
    if _field_instruction_for_marker(value.strip()) is not None:
        return ""
    styles = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    if styles:
        match = re.fullmatch(r"Heading([1-4])", styles[0])
        if match:
            prefix = _heading_prefix_pattern(int(match.group(1))).match(value)
            if prefix:
                return value[prefix.end():]
    return value


def content_inventory(path: Path, normalize_derived: bool = False) -> dict[str, list[str]]:
    """Return authored text by OOXML part, excluding generated field results."""
    ensure_docx(path)
    result: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in sorted(package.namelist()):
            if not CONTENT_PART.match(name):
                continue
            root = etree.fromstring(package.read(name))
            values = []
            for paragraph in root.xpath(".//w:p", namespaces=NS):
                value = _paragraph_text_without_field_results(paragraph)
                if normalize_derived:
                    value = _normalize_derived_paragraph_text(paragraph, value)
                values.append(value)
            result[name] = values
    return result


def content_fingerprint(path: Path, normalize_derived: bool = False) -> str:
    encoded = json.dumps(
        content_inventory(path, normalize_derived=normalize_derived),
        ensure_ascii=False,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def word_xml_counts(path: Path) -> dict[str, int]:
    ensure_docx(path)
    counts = {
        "fields": 0,
        "equations": 0,
        "drawings": 0,
        "text_boxes": 0,
        "footnotes": 0,
        "endnotes": 0,
        "comments": 0,
    }
    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        counts["footnotes"] = int("word/footnotes.xml" in names)
        counts["endnotes"] = int("word/endnotes.xml" in names)
        counts["comments"] = int("word/comments.xml" in names)
        for name in names:
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = etree.fromstring(package.read(name))
            counts["fields"] += len(root.xpath(".//w:fldChar | .//w:fldSimple", namespaces=NS))
            counts["equations"] += len(
                root.xpath(
                    ".//*[local-name()='oMath' or local-name()='oMathPara']"
                )
            )
            counts["drawings"] += len(root.xpath(".//w:drawing", namespaces=NS))
            counts["text_boxes"] += len(root.xpath(".//w:txbxContent", namespaces=NS))
    return counts


def iter_document_paragraphs(document: Any) -> Iterable[Any]:
    for paragraph in document.paragraphs:
        yield paragraph
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph


def style_name_for_selector(selector: dict[str, str]) -> str | None:
    kind, value = selector["kind"], selector["value"]
    if kind == "style_name":
        return value
    if kind == "paragraph_role":
        return ROLE_STYLE_MAP.get(value)
    if kind == "caption_role":
        return "Caption"
    if kind == "bibliography_role":
        return "Bibliography"
    return None


def supported_properties(rule: dict[str, Any]) -> set[str]:
    kind = rule["selector"]["kind"]
    if kind in {"document", "section_role"}:
        return SECTION_PROPERTIES
    if kind == "table_role":
        return TABLE_PROPERTIES
    if kind == "field_role":
        return FIELD_PROPERTIES
    if kind == "equation_role":
        return EQUATION_PROPERTIES
    if style_name_for_selector(rule["selector"]):
        return STYLE_PROPERTIES
    return set()


def unsupported_properties(rule: dict[str, Any]) -> list[str]:
    return sorted(set(rule["properties"]) - supported_properties(rule))


def _font_attributes(font_element: Any) -> Any:
    r_pr = font_element.get_or_add_rPr()
    r_fonts = r_pr.rFonts
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    return r_fonts


def _set_font_attributes(font_element: Any, name: str, attributes: tuple[str, ...]) -> None:
    r_fonts = _font_attributes(font_element)
    for attr in attributes:
        r_fonts.set(qn(f"w:{attr}"), name)


def _set_east_asian_font(font_element: Any, name: str) -> None:
    _set_font_attributes(font_element, name, ("ascii", "hAnsi", "eastAsia", "cs"))


def apply_style_properties(style: Any, properties: dict[str, Any]) -> None:
    font = style.font
    paragraph_format = style.paragraph_format

    if "font_name" in properties:
        font.name = properties["font_name"]
        _set_east_asian_font(style.element, properties["font_name"])
    if "font_name_ascii" in properties:
        font.name = properties["font_name_ascii"]
        _set_font_attributes(
            style.element, properties["font_name_ascii"], ("ascii", "hAnsi")
        )
    if "font_name_east_asia" in properties:
        _set_font_attributes(
            style.element, properties["font_name_east_asia"], ("eastAsia",)
        )
    if "font_name_complex_script" in properties:
        _set_font_attributes(
            style.element, properties["font_name_complex_script"], ("cs",)
        )
    if "font_size_pt" in properties:
        font.size = Pt(float(properties["font_size_pt"]))
    if "bold" in properties:
        font.bold = bool(properties["bold"])
    if "italic" in properties:
        font.italic = bool(properties["italic"])
    if "color_hex" in properties:
        color = str(properties["color_hex"]).lstrip("#")
        if not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
            raise FormatMonographError(f"Invalid color_hex: {properties['color_hex']}")
        font.color.rgb = RGBColor.from_string(color.upper())
    if "alignment" in properties:
        value = properties["alignment"]
        if value not in ALIGNMENTS:
            raise FormatMonographError(f"Unsupported paragraph alignment: {value}")
        paragraph_format.alignment = ALIGNMENTS[value]
    for key, attr in (
        ("space_before_pt", "space_before"),
        ("space_after_pt", "space_after"),
        ("first_line_indent_pt", "first_line_indent"),
        ("left_indent_pt", "left_indent"),
        ("right_indent_pt", "right_indent"),
    ):
        if key in properties:
            setattr(paragraph_format, attr, Pt(float(properties[key])))
    if "line_spacing" in properties:
        paragraph_format.line_spacing = float(properties["line_spacing"])
    if "line_spacing_rule" in properties:
        rules = {
            "single": WD_LINE_SPACING.SINGLE,
            "one_point_five": WD_LINE_SPACING.ONE_POINT_FIVE,
            "double": WD_LINE_SPACING.DOUBLE,
            "at_least": WD_LINE_SPACING.AT_LEAST,
            "exact": WD_LINE_SPACING.EXACTLY,
            "multiple": WD_LINE_SPACING.MULTIPLE,
        }
        value = properties["line_spacing_rule"]
        if value not in rules:
            raise FormatMonographError(f"Unsupported line_spacing_rule: {value}")
        paragraph_format.line_spacing_rule = rules[value]
    if "line_spacing_pt" in properties:
        paragraph_format.line_spacing = Pt(float(properties["line_spacing_pt"]))
        if "line_spacing_rule" not in properties:
            paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    if "first_line_indent_chars" in properties:
        p_pr = style.element.get_or_add_pPr()
        ind = p_pr.get_or_add_ind()
        ind.set(
            qn("w:firstLineChars"),
            str(int(round(float(properties["first_line_indent_chars"]) * 100))),
        )
        ind.attrib.pop(qn("w:firstLine"), None)
    for key, attr in (
        ("keep_with_next", "keep_with_next"),
        ("keep_together", "keep_together"),
        ("page_break_before", "page_break_before"),
        ("widow_control", "widow_control"),
    ):
        if key in properties:
            setattr(paragraph_format, attr, bool(properties[key]))


def _set_document_toggle(document: Any, name: str, enabled: bool) -> None:
    settings = document.settings.element
    element = settings.find(qn(f"w:{name}"))
    if enabled:
        if element is None:
            element = OxmlElement(f"w:{name}")
            settings.append(element)
        element.set(qn("w:val"), "true")
    elif element is not None:
        settings.remove(element)


def apply_section_properties(document: Any, properties: dict[str, Any]) -> int:
    sections = list(document.sections)
    page_size_policy = properties.get("page_size_policy")
    if page_size_policy not in {None, "preserve"}:
        raise FormatMonographError(f"Unsupported page_size_policy: {page_size_policy}")

    explicit_size = "page_width_mm" in properties or "page_height_mm" in properties
    for section in sections:
        if "page_width_mm" in properties:
            section.page_width = Mm(float(properties["page_width_mm"]))
        if "page_height_mm" in properties:
            section.page_height = Mm(float(properties["page_height_mm"]))
        if "orientation" in properties:
            orientation = str(properties["orientation"]).lower()
            if orientation not in {"portrait", "landscape"}:
                raise FormatMonographError(f"Unsupported orientation: {orientation}")
            target = WD_ORIENT.PORTRAIT if orientation == "portrait" else WD_ORIENT.LANDSCAPE
            if section.orientation != target and not explicit_size:
                section.page_width, section.page_height = section.page_height, section.page_width
            section.orientation = target

        for key, attr in (
            ("margin_top_mm", "top_margin"),
            ("margin_bottom_mm", "bottom_margin"),
            ("margin_left_mm", "left_margin"),
            ("margin_right_mm", "right_margin"),
            ("gutter_mm", "gutter"),
        ):
            if key in properties:
                setattr(section, attr, Mm(float(properties[key])))

        width_mm = float(section.page_width.mm)
        height_mm = float(section.page_height.mm)
        ratio_values = (
            ("margin_inner_ratio", "left_margin", width_mm),
            ("margin_outer_ratio", "right_margin", width_mm),
            ("margin_top_ratio", "top_margin", height_mm),
            ("margin_bottom_ratio", "bottom_margin", height_mm),
            ("header_distance_ratio", "header_distance", height_mm),
            ("footer_distance_ratio", "footer_distance", height_mm),
        )
        for key, attr, basis in ratio_values:
            if key in properties:
                ratio = float(properties[key])
                if not 0 <= ratio < 0.5:
                    raise FormatMonographError(f"{key} must be between 0 and 0.5.")
                setattr(section, attr, Mm(basis * ratio))

        if "different_first_page_header_footer" in properties:
            section.different_first_page_header_footer = bool(
                properties["different_first_page_header_footer"]
            )

    if "odd_and_even_pages_header_footer" in properties:
        document.settings.odd_and_even_pages_header_footer = bool(
            properties["odd_and_even_pages_header_footer"]
        )
    if "mirror_margins" in properties:
        _set_document_toggle(document, "mirrorMargins", bool(properties["mirror_margins"]))

    return len(sections)

def _set_repeat_table_header(row: Any, enabled: bool) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    existing = tr_pr.find(qn("w:tblHeader"))
    if enabled:
        if existing is None:
            existing = OxmlElement("w:tblHeader")
            tr_pr.append(existing)
        existing.set(qn("w:val"), "true")
    elif existing is not None:
        tr_pr.remove(existing)


def _set_prevent_row_split(row: Any, enabled: bool) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    existing = tr_pr.find(qn("w:cantSplit"))
    if enabled:
        if existing is None:
            existing = OxmlElement("w:cantSplit")
            tr_pr.append(existing)
        existing.set(qn("w:val"), "true")
    elif existing is not None:
        tr_pr.remove(existing)


def apply_table_properties(document: Any, properties: dict[str, Any]) -> int:
    for table in document.tables:
        if "table_style" in properties:
            table.style = properties["table_style"]
        if "alignment" in properties:
            value = properties["alignment"]
            if value not in TABLE_ALIGNMENTS:
                raise FormatMonographError(f"Unsupported table alignment: {value}")
            table.alignment = TABLE_ALIGNMENTS[value]
        if "repeat_header_row" in properties and table.rows:
            _set_repeat_table_header(table.rows[0], bool(properties["repeat_header_row"]))
        if "prevent_row_split" in properties:
            for row in table.rows:
                _set_prevent_row_split(row, bool(properties["prevent_row_split"]))
        for row in table.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return len(document.tables)



def _record_derived_change(document: Any, change: dict[str, Any]) -> None:
    changes = getattr(document, "_format_monograph_derived_changes", None)
    if changes is None:
        changes = []
        setattr(document, "_format_monograph_derived_changes", changes)
    changes.append(change)


def _set_update_fields_on_open(document: Any, enabled: bool) -> None:
    settings = document.settings.element
    element = settings.find(qn("w:updateFields"))
    if enabled:
        if element is None:
            element = OxmlElement("w:updateFields")
            settings.append(element)
        element.set(qn("w:val"), "true")
    elif element is not None:
        settings.remove(element)


def _mark_fields_dirty(document: Any) -> int:
    root = document.element
    fields = root.xpath(".//w:fldSimple | .//w:fldChar[@w:fldCharType='begin']")
    for field in fields:
        field.set(qn("w:dirty"), "true")
    return len(fields)


def _clear_paragraph_content(paragraph: Any) -> None:
    p = paragraph._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def _append_simple_field(paragraph: Any, instruction: str, placeholder: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    field.set(qn("w:dirty"), "true")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = placeholder
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def _field_instruction_for_marker(marker: str) -> tuple[str, str] | None:
    fixed = {
        "[[TOC]]": ('TOC \\o "1-3" \\h \\z \\u', "Update table of contents"),
        "[[PAGE]]": ("PAGE", "1"),
    }
    if marker in fixed:
        return fixed[marker]
    match = re.fullmatch(r"\[\[(REF|PAGEREF|SEQ):([A-Za-z0-9_.-]+)\]\]", marker)
    if not match:
        return None
    kind, value = match.groups()
    if kind == "REF":
        return f"REF {value} \\h", "0"
    if kind == "PAGEREF":
        return f"PAGEREF {value} \\h", "0"
    return f"SEQ {value} \\* ARABIC \\s 1", "1"


def _convert_explicit_field_markers(document: Any) -> int:
    converted = 0
    for index, paragraph in enumerate(iter_document_paragraphs(document)):
        marker = paragraph.text.strip()
        field = _field_instruction_for_marker(marker)
        if field is None or marker != paragraph.text.strip():
            continue
        _clear_paragraph_content(paragraph)
        _append_simple_field(paragraph, field[0], field[1])
        _record_derived_change(
            document,
            {"kind": "field_marker", "paragraph": index, "source": marker, "field": field[0]},
        )
        converted += 1
    return converted


def _remove_prefix_from_runs(paragraph: Any, length: int) -> None:
    remaining = length
    for run in paragraph.runs:
        if remaining <= 0:
            break
        text = run.text
        take = min(len(text), remaining)
        run.text = text[take:]
        remaining -= take
    if remaining:
        raise FormatMonographError("Heading prefix spans an unsupported non-text object.")


def _heading_prefix_pattern(level: int) -> re.Pattern[str]:
    if level == 1:
        return re.compile(r"^\s*第\s*[0-9一二三四五六七八九十百]+\s*章\s*")
    separators = level - 1
    return re.compile(rf"^\s*\d+(?:[.-]\d+){{{separators}}}\s*")


def _strip_manual_heading_prefixes(document: Any, levels: int) -> int:
    changed = 0
    for index, paragraph in enumerate(document.paragraphs):
        if not paragraph.style or not paragraph.style.name.startswith("Heading "):
            continue
        try:
            level = int(paragraph.style.name.split()[-1])
        except ValueError:
            continue
        if level > levels:
            continue
        text = paragraph.text
        match = _heading_prefix_pattern(level).match(text)
        if not match:
            if re.match(r"^\s*(?:第\s*\d+\s*章|\d+[.-]\d+)", text):
                raise FormatMonographError(
                    f"Ambiguous manual heading number at paragraph {index}: {text[:80]}"
                )
            continue
        prefix = match.group(0)
        _remove_prefix_from_runs(paragraph, len(prefix))
        _record_derived_change(
            document,
            {"kind": "heading_prefix", "paragraph": index, "removed": prefix},
        )
        changed += 1
    return changed


def _next_numbering_id(root: Any, tag: str, attribute: str) -> int:
    values = []
    for element in root.findall(qn(f"w:{tag}")):
        value = element.get(qn(f"w:{attribute}"))
        if value is not None and value.isdigit():
            values.append(int(value))
    return max(values, default=0) + 1


def _ensure_heading_numbering(document: Any, levels: int) -> int:
    if not 1 <= levels <= 4:
        raise FormatMonographError("heading_levels must be between 1 and 4.")
    root = document.part.numbering_part.element
    abstract_id = _next_numbering_id(root, "abstractNum", "abstractNumId")
    num_id = _next_numbering_id(root, "num", "numId")

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)

    for level in range(levels):
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(level))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), "decimal")
        p_style = OxmlElement("w:pStyle")
        p_style.set(qn("w:val"), f"Heading{level + 1}")
        lvl_text = OxmlElement("w:lvlText")
        if level == 0:
            lvl_text.set(qn("w:val"), "第%1章")
        else:
            lvl_text.set(
                qn("w:val"),
                ".".join(f"%{number}" for number in range(1, level + 2)),
            )
        suff = OxmlElement("w:suff")
        suff.set(qn("w:val"), "space")
        lvl.extend([start, num_fmt, p_style, lvl_text, suff])
        abstract.append(lvl)
    root.insert(0, abstract)

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    root.append(num)

    for level in range(levels):
        style = document.styles[f"Heading {level + 1}"]
        p_pr = style.element.get_or_add_pPr()
        num_pr = p_pr.find(qn("w:numPr"))
        if num_pr is None:
            num_pr = OxmlElement("w:numPr")
            p_pr.append(num_pr)
        for child_name in ("ilvl", "numId"):
            existing = num_pr.find(qn(f"w:{child_name}"))
            if existing is not None:
                num_pr.remove(existing)
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), str(level))
        number = OxmlElement("w:numId")
        number.set(qn("w:val"), str(num_id))
        num_pr.extend([ilvl, number])
    return levels


def apply_field_properties(document: Any, properties: dict[str, Any]) -> int:
    changed = 0
    if "update_on_open" in properties:
        _set_update_fields_on_open(document, bool(properties["update_on_open"]))
        changed += 1
    if properties.get("mark_fields_dirty"):
        changed += _mark_fields_dirty(document)
    if properties.get("convert_explicit_markers"):
        changed += _convert_explicit_field_markers(document)

    levels = int(properties.get("heading_levels", 4))
    if properties.get("strip_manual_heading_prefixes"):
        changed += _strip_manual_heading_prefixes(document, levels)
    if properties.get("rebuild_heading_numbering"):
        changed += _ensure_heading_numbering(document, levels)
    return changed


def apply_equation_properties(document: Any, properties: dict[str, Any]) -> int:
    unsupported_values = {
        key: value
        for key, value in properties.items()
        if key in EQUATION_PROPERTIES and value not in {True, False}
    }
    if unsupported_values:
        raise FormatMonographError(
            "Equation policy properties must be boolean: "
            + ", ".join(sorted(unsupported_values))
        )
    return len(document.element.xpath(".//*[local-name()='oMath']"))


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def field_inventory(path: Path) -> dict[str, Any]:
    ensure_docx(path)
    counts: dict[str, int] = {}
    total = 0
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = etree.fromstring(package.read(name))
            instructions = root.xpath(".//w:fldSimple/@w:instr", namespaces=NS)
            instructions.extend(
                text
                for text in root.xpath(".//w:instrText/text()", namespaces=NS)
                if text.strip()
            )
            for instruction in instructions:
                match = re.match(r"\s*([A-Za-z]+)", instruction)
                kind = match.group(1).upper() if match else "UNKNOWN"
                counts[kind] = counts.get(kind, 0) + 1
                total += 1
    return {"total": total, "types": dict(sorted(counts.items()))}


def equation_inventory(path: Path) -> dict[str, Any]:
    ensure_docx(path)
    result = {
        "omml": 0,
        "mathtype_ole": 0,
        "legacy_equation_ole": 0,
        "other_ole": 0,
        "formula_image_candidates": 0,
    }
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = etree.fromstring(package.read(name))
            result["omml"] += len(root.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS))
            for ole in root.xpath(".//*[local-name()='OLEObject']"):
                prog_id = (ole.get("ProgID") or ole.get(f"{{{O_NS}}}ProgID") or "").lower()
                if "dsmt" in prog_id or "mathtype" in prog_id:
                    result["mathtype_ole"] += 1
                elif prog_id in {"equation.3", "equation.2", "equation"}:
                    result["legacy_equation_ole"] += 1
                else:
                    result["other_ole"] += 1

            for paragraph in root.xpath(".//w:p", namespaces=NS):
                if not paragraph.xpath(".//w:drawing | .//v:imagedata", namespaces=NS):
                    continue
                styles = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
                metadata = paragraph.xpath(
                    ".//*[local-name()='docPr']/@name | "
                    ".//*[local-name()='docPr']/@title | "
                    ".//*[local-name()='docPr']/@descr"
                )
                hint = " ".join(styles + metadata).lower()
                if any(token in hint for token in ("equation", "formula", "公式")):
                    result["formula_image_candidates"] += 1
    result["editable_total"] = (
        result["omml"] + result["mathtype_ole"] + result["legacy_equation_ole"]
    )
    return result


def protected_object_manifest(path: Path) -> dict[str, Any]:
    ensure_docx(path)
    embeddings: dict[str, str] = {}
    media: dict[str, str] = {}
    omml: list[str] = []
    with zipfile.ZipFile(path) as package:
        for name in sorted(package.namelist()):
            data = package.read(name)
            if name.startswith("word/embeddings/"):
                embeddings[name] = _sha256_bytes(data)
            elif name.startswith("word/media/"):
                media[name] = _sha256_bytes(data)
            if name.startswith("word/") and name.endswith(".xml"):
                root = etree.fromstring(data)
                for equation in root.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS):
                    canonical = etree.tostring(equation, method="c14n", exclusive=True)
                    omml.append(_sha256_bytes(canonical))
    return {
        "embeddings": embeddings,
        "media": media,
        "omml_sha256": sorted(omml),
    }

def apply_rule(document: Any, rule: dict[str, Any]) -> int:
    unsupported = unsupported_properties(rule)
    if unsupported:
        raise FormatMonographError(
            f"Rule {rule['id']} has unsupported automatic properties: {', '.join(unsupported)}"
        )

    selector = rule["selector"]
    if selector["kind"] in {"document", "section_role"}:
        return apply_section_properties(document, rule["properties"])
    if selector["kind"] == "table_role":
        return apply_table_properties(document, rule["properties"])
    if selector["kind"] == "field_role":
        return apply_field_properties(document, rule["properties"])
    if selector["kind"] == "equation_role":
        return apply_equation_properties(document, rule["properties"])

    style_name = style_name_for_selector(selector)
    if style_name is None:
        raise FormatMonographError(
            f"Rule {rule['id']} uses an unsupported automatic selector: {selector}"
        )
    try:
        style = document.styles[style_name]
    except KeyError as exc:
        raise FormatMonographError(
            f"Rule {rule['id']} targets missing Word style: {style_name}"
        ) from exc
    apply_style_properties(style, rule["properties"])
    return sum(
        1
        for paragraph in iter_document_paragraphs(document)
        if paragraph.style and paragraph.style.name == style_name
    )


def first_anchor_paragraph(document: Any, rule: dict[str, Any]) -> Any | None:
    style_name = style_name_for_selector(rule["selector"])
    paragraphs = list(document.paragraphs)
    if style_name:
        for paragraph in paragraphs:
            if (
                paragraph.text.strip()
                and paragraph.style
                and paragraph.style.name == style_name
                and paragraph.runs
            ):
                return paragraph
    for paragraph in paragraphs:
        if paragraph.text.strip() and paragraph.runs:
            return paragraph
    return None


def summarize_rule(rule: dict[str, Any]) -> str:
    pairs = ", ".join(f"{key}={value}" for key, value in sorted(rule["properties"].items()))
    sources = ", ".join(rule["source_ids"])
    return f"{rule['id']}: {pairs}. Sources: {sources}."


def load_document(path: Path) -> Any:
    ensure_docx(path)
    try:
        return Document(str(path))
    except Exception as exc:
        raise FormatMonographError(f"Unable to open DOCX {path}: {exc}") from exc
