"""Shared helpers for the format-monograph command-line tools."""

from __future__ import annotations

import copy
import hashlib
import json
import os
import platform
import re
import shutil
import subprocess
import zipfile
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
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
FIELD_MARKER_PATTERN = re.compile(
    r"\[\[(?:TOC|PAGE|(?:REF|PAGEREF|SEQ):[A-Za-z0-9_.-]+)\]\]"
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
    "heading_1": "Heading 1",
    "heading_2": "Heading 2",
    "heading_3": "Heading 3",
    "heading_4": "Heading 4",
    "figure_caption": "Caption",
    "table_caption": "Caption",
    "equation_caption": "Caption",
    "reference_entry": "Bibliography",
    "answer": "Normal",
    "teaching_callout": "Normal",
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

CAPTION_POLICY_PROPERTIES = {
    "numbering_mode",
    "preserve_identifier",
    "domain_context",
    "allow_automatic_renumbering",
    "preserve_table_cell_caption_position",
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
    "chapter_start",
}

FONT_ALIAS_GROUPS = (
    {"å®‹ä½“", "simsun", "nsimsun", "æ–°å®‹ä½“"},
    {"é»‘ä½“", "simhei"},
    {"æ¥·ä½“", "kaiti", "simkai", "æ¥·ä½“_gb2312"},
    {"ä»¿å®‹", "fangsong", "simfang", "ä»¿å®‹_gb2312"},
)

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


def _font_directories() -> list[Path]:
    candidates: list[Path] = []
    system = platform.system()
    if system == "Windows":
        if os.environ.get("WINDIR"):
            candidates.append(Path(os.environ["WINDIR"]) / "Fonts")
        if os.environ.get("LOCALAPPDATA"):
            candidates.append(
                Path(os.environ["LOCALAPPDATA"]) / "Microsoft" / "Windows" / "Fonts"
            )
    elif system == "Darwin":
        candidates.extend(
            [
                Path("/System/Library/Fonts"),
                Path("/Library/Fonts"),
                Path.home() / "Library" / "Fonts",
            ]
        )
    else:
        candidates.extend(
            [
                Path("/usr/share/fonts"),
                Path("/usr/local/share/fonts"),
                Path.home() / ".fonts",
                Path.home() / ".local" / "share" / "fonts",
            ]
        )
    return [path for path in candidates if path.is_dir()]


def available_font_names() -> set[str]:
    names: set[str] = set()
    fc_list = shutil.which("fc-list")
    if fc_list:
        try:
            result = subprocess.run(
                [fc_list, "--format=%{family}\\n"],
                capture_output=True,
                text=True,
                check=False,
                timeout=20,
            )
            for line in result.stdout.splitlines():
                names.update(part.strip() for part in line.split(",") if part.strip())
        except (OSError, subprocess.TimeoutExpired):
            pass

    if platform.system() == "Windows":
        try:
            import winreg

            key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                index = 0
                while True:
                    try:
                        label, _, _ = winreg.EnumValue(key, index)
                    except OSError:
                        break
                    label = re.sub(r"\s*\([^)]*\)\s*$", "", label)
                    names.update(
                        part.strip()
                        for part in re.split(r"\s*&\s*", label)
                        if part.strip()
                    )
                    index += 1
        except OSError:
            pass

    for directory in _font_directories():
        for pattern in ("*.ttf", "*.ttc", "*.otf", "*.dfont"):
            for path in directory.rglob(pattern):
                names.add(path.stem)
    return names


def _font_key(value: str) -> str:
    return re.sub(r"[\s_-]+", "", value).casefold()


def font_alias_keys(value: str) -> set[str]:
    key = _font_key(value)
    for group in FONT_ALIAS_GROUPS:
        normalized = {_font_key(item) for item in group}
        if key in normalized:
            return normalized
    return {key}


def resolve_font_name(value: str, available_names: set[str] | None = None) -> dict[str, Any]:
    available_names = available_font_names() if available_names is None else available_names
    available_by_key = {_font_key(name): name for name in available_names}
    for key in font_alias_keys(value):
        if key in available_by_key:
            return {
                "requested": value,
                "available": True,
                "matched_name": available_by_key[key],
                "match": "exact" if key == _font_key(value) else "verified_alias",
            }
    return {
        "requested": value,
        "available": False,
        "matched_name": None,
        "match": "missing",
    }


def required_profile_fonts(profile: dict[str, Any]) -> list[str]:
    keys = {
        "font_name",
        "font_name_ascii",
        "font_name_east_asia",
        "font_name_complex_script",
    }
    return sorted(
        {
            str(value)
            for rule in profile.get("rules", [])
            if rule.get("status") == "approved"
            and rule.get("application") == "automatic"
            for key, value in rule.get("properties", {}).items()
            if key in keys and str(value).strip()
        }
    )


def missing_profile_fonts(profile: dict[str, Any]) -> list[str]:
    return [
        item["requested"]
        for item in profile_font_resolutions(profile)
        if not item["available"]
    ]


def profile_font_resolutions(profile: dict[str, Any]) -> list[dict[str, Any]]:
    available = available_font_names()
    return [
        resolve_font_name(name, available)
        for name in required_profile_fonts(profile)
    ]


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
    value = FIELD_MARKER_PATTERN.sub("", value)
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
            counts["equations"] += len(root.xpath(".//*[local-name()='oMath']"))
            counts["drawings"] += len(root.xpath(".//w:drawing", namespaces=NS))
            counts["text_boxes"] += len(root.xpath(".//w:txbxContent", namespaces=NS))
    return counts


def _iter_container_paragraphs(container: Any) -> Iterable[Any]:
    for paragraph in container.paragraphs:
        yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_container_paragraphs(cell)


def iter_document_paragraphs(document: Any) -> Iterable[Any]:
    # Accessing an undefined header/footer through python-docx materializes a new
    # part. Restrict automatic paragraph traversal to existing body containers
    # so a read-only rule count cannot alter the package.
    yield from _iter_container_paragraphs(document)


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
    kind = rule["selecï­¶¶‰žËkºwµçQ¥¹}ÁÉ•™¥àˆ°€‰Á…É…É…Á ˆè¥¹‘•à°€‰É•µ½Ù•ˆèÁÉ•™¥áô°(€€€€€€€€¤(€€€€€€€¡…¹•€¬ô€Ä(€€€É•ÑÕÉ¸¡…¹•(()‘•˜}¹•áÑ}¹Õµ‰•É¥¹}¥¡É½½Ðè¹ä°Ñ…œèÍÑÈ°…ÑÑÉ¥‰ÕÑ”èÍÑÈ¤€´ø¥¹Ðè(€€€Ù…±Õ•Ì€ômt(€€€™½È•±•µ•¹Ð¥¸É½½Ð¹™¥¹‘…±°¡Å¸¡˜‰ÜéíÑ…ôˆ¤¤è(€€€€€€€Ù…±Õ”€ô•±•µ•¹Ð¹•Ð¡Å¸¡˜‰Üéí…ÑÑÉ¥‰ÕÑ•ôˆ¤¤(€€€€€€€¥˜Ù…±Õ”¥Ì¹½Ð9½¹”…¹Ù…±Õ”¹¥Í‘¥¥Ð ¤è(€€€€€€€€€€€Ù…±Õ•Ì¹…ÁÁ•¹¡¥¹Ð¡Ù…±Õ”¤¤(€€€É•ÑÕÉ¸µ…à¡Ù…±Õ•Ì°‘•™…Õ±ÐôÀ¤€¬€Ä(()‘•˜}•¹ÍÕÉ•}¡•…‘¥¹}¹Õµ‰•É¥¹œ¡‘½Õµ•¹Ðè¹ä°±•Ù•±Ìè¥¹Ð°¡…ÁÑ•É}ÍÑ…ÉÐè¥¹Ð€ô€Ä¤€´ø¥¹Ðè(€€€¥˜¹½Ð€Ä€ðô±•Ù•±Ì€ðô€Ðè(€€€€€€€É…¥Í”½Éµ…Ñ5½¹½É…Á¡ÉÉ½È ‰¡•…‘¥¹}±•Ù•±ÌµÕÍÐ‰”‰•ÑÝ••¸€Ä…¹€Ð¸ˆ¤(€€€¥˜¡…ÁÑ•É}ÍÑ…ÉÐ€ð€Äè(€€€€€€€É…¥Í”½Éµ…Ñ5½¹½É…Á¡ÉÉ½È ‰¡…ÁÑ•É}ÍÑ…ÉÐµÕÍÐ‰”„Á½Í¥Ñ¥Ù”¥¹Ñ••È¸ˆ¤(€€€É½½Ð€ô‘½Õµ•¹Ð¹Á…ÉÐ¹¹Õµ‰•É¥¹}Á…ÉÐ¹•±•µ•¹Ð(€€€…‰ÍÑÉ…Ñ}¥€ô}¹•áÑ}¹Õµ‰•É¥¹}¥¡É½½Ð°€‰…‰ÍÑÉ…Ñ9Õ´ˆ°€‰…‰ÍÑÉ…Ñ9Õµ%ˆ¤(€€€¹Õµ}¥€ô}¹•áÑ}¹Õµ‰•É¥¹}¥¡É½½Ð°€‰¹Õ´ˆ°€‰¹Õµ%ˆ¤((€€€…‰ÍÑÉ…Ð€ô=áµ±±•µ•¹Ð ‰Üé…‰ÍÑÉ…Ñ9Õ´ˆ¤(€€€…‰ÍÑÉ…Ð¹Í•Ð¡Å¸ ‰Üé…‰ÍÑÉ…Ñ9Õµ%ˆ¤°ÍÑÈ¡…‰ÍÑÉ…Ñ}¥¤¤(€€€µÕ±Ñ¤€ô=áµ±±•µ•¹Ð ‰ÜéµÕ±Ñ¥1•Ù•±QåÁ”ˆ¤(€€€µÕ±Ñ¤¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°€‰µÕ±Ñ¥±•Ù•°ˆ¤(€€€…‰ÍÑÉ…Ð¹…ÁÁ•¹¡µÕ±Ñ¤¤((€€€™½È±•Ù•°¥¸É…¹”¡±•Ù•±Ì¤è(€€€€€€€±Ù°€ô=áµ±±•µ•¹Ð ‰Üé±Ù°ˆ¤(€€€€€€€±Ù°¹Í•Ð¡Å¸ ‰Üé¥±Ù°ˆ¤°ÍÑÈ¡±•Ù•°¤¤(€€€€€€€ÍÑ…ÉÐ€ô=áµ±±•µ•¹Ð ‰ÜéÍÑ…ÉÐˆ¤(€€€€€€€ÍÑ…ÉÐ¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°ÍÑÈ¡¡…ÁÑ•É}ÍÑ…ÉÐ¥˜±•Ù•°€ôô€À•±Í”€Ä¤¤(€€€€€€€¹Õµ}™µÐ€ô=áµ±±•µ•¹Ð ‰Üé¹ÕµµÐˆ¤(€€€€€€€¹Õµ}™µÐ¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°€‰‘•¥µ…°ˆ¤(€€€€€€€Á}ÍÑå±”€ô=áµ±±•µ•¹Ð ‰ÜéÁMÑå±”ˆ¤(€€€€€€€Á}ÍÑå±”¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°˜‰!•…‘¥¹í±•Ù•°€¬€Åôˆ¤(€€€€€€€±Ù±}Ñ•áÐ€ô=áµ±±•µ•¹Ð ‰Üé±Ù±Q•áÐˆ¤(€€€€€€€¥˜±•Ù•°€ôô€Àè(€€€€€€€€€€€±Ù±}Ñ•áÐ¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°€‹ž²°”Çž®€ˆ¤(€€€€€€€•±Í”è(€€€€€€€€€€€±Ù±}Ñ•áÐ¹Í•Ð (€€€€€€€€€€€€€€€Å¸ ‰ÜéÙ…°ˆ¤°(€€€€€€€€€€€€€€€€ˆ¸ˆ¹©½¥¸¡˜ˆ•í¹Õµ‰•Éôˆ™½È¹Õµ‰•È¥¸É…¹” Ä°±•Ù•°€¬€È¤¤°(€€€€€€€€€€€€¤(€€€€€€€ÍÕ™˜€ô=áµ±±•µ•¹Ð ‰ÜéÍÕ™˜ˆ¤(€€€€€€€ÍÕ™˜¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°€‰ÍÁ…”ˆ¤(€€€€€€€±Ù°¹•áÑ•¹¡mÍÑ…ÉÐ°¹Õµ}™µÐ°Á}ÍÑå±”°±Ù±}Ñ•áÐ°ÍÕ™™t¤(€€€€€€€…‰ÍÑÉ…Ð¹…ÁÁ•¹¡±Ù°¤(€€€É½½Ð¹¥¹Í•ÉÐ À°…‰ÍÑÉ…Ð¤((€€€¹Õ´€ô=áµ±±•µ•¹Ð ‰Üé¹Õ´ˆ¤(€€€¹Õ´¹Í•Ð¡Å¸ ‰Üé¹Õµ%ˆ¤°ÍÑÈ¡¹Õµ}¥¤¤(€€€…‰ÍÑÉ…Ñ}É•˜€ô=áµ±±•µ•¹Ð ‰Üé…‰ÍÑÉ…Ñ9Õµ%ˆ¤(€€€…‰ÍÑÉ…Ñ}É•˜¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°ÍÑÈ¡…‰ÍÑÉ…Ñ}¥¤¤(€€€¹Õ´¹…ÁÁ•¹¡…‰ÍÑÉ…Ñ}É•˜¤(€€€¥˜¡…ÁÑ•É}ÍÑ…ÉÐ€„ô€Äè(€€€€€€€½Ù•ÉÉ¥‘”€ô=áµ±±•µ•¹Ð ‰Üé±Ù±=Ù•ÉÉ¥‘”ˆ¤(€€€€€€€½Ù•ÉÉ¥‘”¹Í•Ð¡Å¸ ‰Üé¥±Ù°ˆ¤°€ˆÀˆ¤(€€€€€€€ÍÑ…ÉÑ}½Ù•ÉÉ¥‘”€ô=áµ±±•µ•¹Ð ‰ÜéÍÑ…ÉÑ=Ù•ÉÉ¥‘”ˆ¤(€€€€€€€ÍÑ…ÉÑ}½Ù•ÉÉ¥‘”¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°ÍÑÈ¡¡…ÁÑ•É}ÍÑ…ÉÐ¤¤(€€€€€€€½Ù•ÉÉ¥‘”¹…ÁÁ•¹¡ÍÑ…ÉÑ}½Ù•ÉÉ¥‘”¤(€€€€€€€¹Õ´¹…ÁÁ•¹¡½Ù•ÉÉ¥‘”¤(€€€É½½Ð¹…ÁÁ•¹¡¹Õ´¤((€€€™½È±•Ù•°¥¸É…¹”¡±•Ù•±Ì¤è(€€€€€€€ÍÑå±”€ô•¹ÍÕÉ•}Á…É…É…Á¡}ÍÑå±”¡‘½Õµ•¹Ð°˜‰!•…‘¥¹œí±•Ù•°€¬€Åôˆ¤(€€€€€€€Á}ÁÈ€ôÍÑå±”¹•±•µ•¹Ð¹•Ñ}½É}…‘‘}ÁAÈ ¤(€€€€€€€¹Õµ}ÁÈ€ôÁ}ÁÈ¹™¥¹¡Å¸ ‰Üé¹ÕµAÈˆ¤¤(€€€€€€€¥˜¹Õµ}ÁÈ¥Ì9½¹”è(€€€€€€€€€€€¹Õµ}ÁÈ€ô=áµ±±•µ•¹Ð ‰Üé¹ÕµAÈˆ¤(€€€€€€€€€€€Á}ÁÈ¹…ÁÁ•¹¡¹Õµ}ÁÈ¤(€€€€€€€™½È¡¥±‘}¹…µ”¥¸€ ‰¥±Ù°ˆ°€‰¹Õµ%ˆ¤è(€€€€€€€€€€€•á¥ÍÑ¥¹œ€ô¹Õµ}ÁÈ¹™¥¹¡Å¸¡˜‰Üéí¡¥±‘}¹…µ•ôˆ¤¤(€€€€€€€€€€€¥˜•á¥ÍÑ¥¹œ¥Ì¹½Ð9½¹”è(€€€€€€€€€€€€€€€¹Õµ}ÁÈ¹É•µ½Ù”¡•á¥ÍÑ¥¹œ¤(€€€€€€€¥±Ù°€ô=áµ±±•µ•¹Ð ‰Üé¥±Ù°ˆ¤(€€€€€€€¥±Ù°¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°ÍÑÈ¡±•Ù•°¤¤(€€€€€€€¹Õµ‰•È€ô=áµ±±•µ•¹Ð ‰Üé¹Õµ%ˆ¤(€€€€€€€¹Õµ‰•È¹Í•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°ÍÑÈ¡¹Õµ}¥¤¤(€€€€€€€¹Õµ}ÁÈ¹•áÑ•¹¡m¥±Ù°°¹Õµ‰•Ét¤(€€€É•ÑÕÉ¸±•Ù•±Ì(()‘•˜…ÁÁ±å}™¥•±‘}ÁÉ½Á•ÉÑ¥•Ì¡‘½Õµ•¹Ðè¹ä°ÁÉ½Á•ÉÑ¥•Ìè‘¥ÑmÍÑÈ°¹åt¤€´ø¥¹Ðè(€€€¡…¹•€ô€À(€€€¥˜€‰ÕÁ‘…Ñ•}½¹}½Á•¸ˆ¥¸ÁÉ½Á•ÉÑ¥•Ìè(€€€€€€€}Í•Ñ}ÕÁ‘…Ñ•}™¥•±‘Í}½¹}½Á•¸¡‘½Õµ•¹Ð°‰½½°¡ÁÉ½Á•ÉÑ¥•Íl‰ÕÁ‘…Ñ•}½¹}½Á•¸‰t¤¤(€€€€€€€¡…¹•€¬ô€Ä(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰µ…É­}™¥•±‘Í}‘¥ÉÑäˆ¤è(€€€€€€€¡…¹•€¬ô}µ…É­}™¥•±‘Í}‘¥ÉÑä¡‘½Õµ•¹Ð¤(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰½¹Ù•ÉÑ}•áÁ±¥¥Ñ}µ…É­•ÉÌˆ¤è(€€€€€€€¡…¹•€¬ô}½¹Ù•ÉÑ}•áÁ±¥¥Ñ}™¥•±‘}µ…É­•ÉÌ¡‘½Õµ•¹Ð¤((€€€±•Ù•±Ì€ô¥¹Ð¡ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰¡•…‘¥¹}±•Ù•±Ìˆ°€Ð¤¤(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰ÍÑÉ¥Á}µ…¹Õ…±}¡•…‘¥¹}ÁÉ•™¥á•Ìˆ¤è(€€€€€€€¡…¹•€¬ô}ÍÑÉ¥Á}µ…¹Õ…±}¡•…‘¥¹}ÁÉ•™¥á•Ì¡‘½Õµ•¹Ð°±•Ù•±Ì¤(€€€¥˜ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰É•‰Õ¥±‘}¡•…‘¥¹}¹Õµ‰•É¥¹œˆ¤è(€€€€€€€¡…¹•€¬ô}•¹ÍÕÉ•}¡•…‘¥¹}¹Õµ‰•É¥¹œ (€€€€€€€€€€€‘½Õµ•¹Ð°±•Ù•±Ì°¥¹Ð¡ÁÉ½Á•ÉÑ¥•Ì¹•Ð ‰¡…ÁÑ•É}ÍÑ…ÉÐˆ°€Ä¤¤(€€€€€€€€¤(€€€É•ÑÕÉ¸¡…¹•(()‘•˜…ÁÁ±å}•ÅÕ…Ñ¥½¹}ÁÉ½Á•ÉÑ¥•Ì¡‘½Õµ•¹Ðè¹ä°ÁÉ½Á•ÉÑ¥•Ìè‘¥ÑmÍÑÈ°¹åt¤€´ø¥¹Ðè(€€€Õ¹ÍÕÁÁ½ÉÑ•‘}Ù…±Õ•Ì€ôì(€€€€€€€­•äèÙ…±Õ”(€€€€€€€™½È­•ä°Ù…±Õ”¥¸ÁÉ½Á•ÉÑ¥•Ì¹¥Ñ•µÌ ¤(€€€€€€€¥˜­•ä¥¸EUQ%=9}AI=AIQ%L…¹Ù…±Õ”¹½Ð¥¸íQÉÕ”°…±Í•ô(€€€ô(€€€¥˜Õ¹ÍÕÁÁ½ÉÑ•‘}Ù…±Õ•Ìè(€€€€€€€É…¥Í”½Éµ…Ñ5½¹½É…Á¡ÉÉ½È (€€€€€€€€€€€€‰ÅÕ…Ñ¥½¸Á½±¥äÁÉ½Á•ÉÑ¥•ÌµÕÍÐ‰”‰½½±•…¸è€ˆ(€€€€€€€€€€€€¬€ˆ°€ˆ¹©½¥¸¡Í½ÉÑ•¡Õ¹ÍÕÁÁ½ÉÑ•‘}Ù…±Õ•Ì¤¤(€€€€€€€€¤(€€€É•ÑÕÉ¸±•¸¡‘½Õµ•¹Ð¹•±•µ•¹Ð¹áÁ…Ñ  ˆ¸¼¼©m±½…°µ¹…µ” ¤ô½5…Ñ tˆ¤¤(()‘•˜}Í¡„ÈÔÙ}‰åÑ•Ì¡Ù…±Õ”è‰åÑ•Ì¤€´øÍÑÈè(€€€É•ÑÕÉ¸¡…Í¡±¥ˆ¹Í¡„ÈÔØ¡Ù…±Õ”¤¹¡•á‘¥•ÍÐ ¤(()‘•˜}½µÁ±•á}™¥•±‘}¥¹ÍÑÉÕÑ¥½¹Ì¡É½½Ðè•ÑÉ•”¹}±•µ•¹Ð¤€´ø±¥ÍÑmÍÑÉtè(€€€¥¹ÍÑÉÕÑ¥½¹Ìè±¥ÍÑmÍÑÉt€ômt(€€€™½ÈÁ…É…É…Á ¥¸É½½Ð¹áÁ…Ñ  ˆ¸¼½ÜéÀˆ°¹…µ•ÍÁ…•Ìõ9L¤è(€€€€€€€ÍÑ…¬è±¥ÍÑm‘¥ÑmÍÑÈ°¹åut€ômt(€€€€€€€™½È•±•µ•¹Ð¥¸Á…É…É…Á ¹¥Ñ•È ¤è(€€€€€€€€€€€¥˜•±•µ•¹Ð¹Ñ…œ€ôôÅ¸ ‰Üé™±‘¡…Èˆ¤è(€€€€€€€€€€€€€€€­¥¹€ô•±•µ•¹Ð¹•Ð¡Å¸ ‰Üé™±‘¡…ÉQåÁ”ˆ¤¤(€€€€€€€€€€€€€€€¥˜­¥¹€ôô€‰‰•¥¸ˆè(€€€€€€€€€€€€€€€€€€€ÍÑ…¬¹…ÁÁ•¹¡ì‰Á…ÉÑÌˆèmt°€‰…ÁÑÕÉ•ˆè…±Í•ô¤(€€€€€€€€€€€€€€€•±¥˜­¥¹€ôô€‰Í•Á…É…Ñ”ˆ…¹ÍÑ…¬…¹¹½ÐÍÑ…­l´Åul‰…ÁÑÕÉ•‰tè(€€€€€€€€€€€€€€€€€€€Ù…±Õ”€ô€ˆˆ¹©½¥¸¡ÍÑ…­l´Åul‰Á…ÉÑÌ‰t¤¹ÍÑÉ¥À ¤(€€€€€€€€€€€€€€€€€€€¥˜Ù…±Õ”è(€€€€€€€€€€€€€€€€€€€€€€€¥¹ÍÑÉÕÑ¥½¹Ì¹…ÁÁ•¹¡Ù…±Õ”¤(€€€€€€€€€€€€€€€€€€€ÍÑ…­l´Åul‰…ÁÑÕÉ•‰t€ôQÉÕ”(€€€€€€€€€€€€€€€•±¥˜­¥¹€ôô€‰•¹ˆ…¹ÍÑ…¬è(€€€€€€€€€€€€€€€€€€€™¥•±€ôÍÑ…¬¹Á½À ¤(€€€€€€€€€€€€€€€€€€€¥˜¹½Ð™¥•±‘l‰…ÁÑÕÉ•‰tè(€€€€€€€€€€€€€€€€€€€€€€€Ù…±Õ”€ô€ˆˆ¹©½¥¸¡™¥•±‘l‰Á…ÉÑÌ‰t¤¹ÍÑÉ¥À ¤(€€€€€€€€€€€€€€€€€€€€€€€¥˜Ù…±Õ”è(€€€€€€€€€€€€€€€€€€€€€€€€€€€¥¹ÍÑÉÕÑ¥½¹Ì¹…ÁÁ•¹¡Ù…±Õ”¤(€€€€€€€€€€€•±¥˜€ (€€€€€€€€€€€€€€€•±•µ•¹Ð¹Ñ…œ€ôôÅ¸ ‰Üé¥¹ÍÑÉQ•áÐˆ¤(€€€€€€€€€€€€€€€…¹ÍÑ…¬(€€€€€€€€€€€€€€€…¹¹½ÐÍÑ…­l´Åul‰…ÁÑÕÉ•‰t(€€€€€€€€€€€€¤è(€€€€€€€€€€€€€€€ÍÑ…­l´Åul‰Á…ÉÑÌ‰t¹…ÁÁ•¹¡•±•µ•¹Ð¹Ñ•áÐ½È€ˆˆ¤(€€€É•ÑÕÉ¸¥¹ÍÑÉÕÑ¥½¹Ì(()‘•˜™¥•±‘}¥¹Ù•¹Ñ½Éä¡Á…Ñ èA…Ñ ¤€´ø‘¥ÑmÍÑÈ°¹åtè(€€€•¹ÍÕÉ•}‘½à¡Á…Ñ ¤(€€€½Õ¹ÑÌè‘¥ÑmÍÑÈ°¥¹Ñt€ôíô(€€€¥¹ÍÑÉÕÑ¥½¹Ìè±¥ÍÑmÍÑÉt€ômt(€€€‰½½­µ…É­ÌèÍ•ÑmÍÑÉt€ôÍ•Ð ¤(€€€Ý¥Ñ é¥Á™¥±”¹i¥Á¥±”¡Á…Ñ ¤…ÌÁ…­…”è(€€€€€€€™½È¹…µ”¥¸Á…­…”¹¹…µ•±¥ÍÐ ¤è(€€€€€€€€€€€¥˜¹½Ð¹…µ”¹ÍÑ…ÉÑÍÝ¥Ñ  ‰Ý½É¼ˆ¤½È¹½Ð¹…µ”¹•¹‘ÍÝ¥Ñ  ˆ¹áµ°ˆ¤è(€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”(€€€€€€€€€€€É½½Ð€ô•ÑÉ•”¹™É½µÍÑÉ¥¹œ¡Á…­…”¹É•…¡¹…µ”¤¤(€€€€€€€€€€€‰½½­µ…É­Ì¹ÕÁ‘…Ñ” (€€€€€€€€€€€€€€€É½½Ð¹áÁ…Ñ  ˆ¸¼½Üé‰½½­µ…É­MÑ…ÉÐ½Üé¹…µ”ˆ°¹…µ•ÍÁ…•Ìõ9L¤(€€€€€€€€€€€€¤(€€€€€€€€€€€¥¹ÍÑÉÕÑ¥½¹Ì¹•áÑ•¹ (€€€€€€€€€€€€€€€Ù…±Õ”¹ÍÑÉ¥À ¤(€€€€€€€€€€€€€€€™½ÈÙ…±Õ”¥¸É½½Ð¹áÁ…Ñ  ˆ¸¼½Üé™±‘M¥µÁ±”½Üé¥¹ÍÑÈˆ°¹…µ•ÍÁ…•Ìõ9L¤(€€€€€€€€€€€€€€€¥˜Ù…±Õ”¹ÍÑÉ¥À ¤(€€€€€€€€€€€€¤(€€€€€€€€€€€¥¹ÍÑÉÕÑ¥½¹Ì¹•áÑ•¹¡}½µÁ±•á}™¥•±‘}¥¹ÍÑÉÕÑ¥½¹Ì¡É½½Ð¤¤((€€€É•™•É•¹•Ì€ômt(€€€Í•ÅÕ•¹•Ì€ômt(€€€™½È¥¹ÍÑÉÕÑ¥½¸¥¸¥¹ÍÑÉÕÑ¥½¹Ìè(€€€€€€€µ…Ñ €ôÉ”¹µ…Ñ ¡È‰qÌ¨¡mµi„µét¬¤ üéqÌ¬¡myqqqÍt¬¤¤üˆ°¥¹ÍÑÉÕÑ¥½¸¤(€€€€€€€­¥¹€ôµ…Ñ ¹É½ÕÀ Ä¤¹ÕÁÁ•È ¤¥˜µ…Ñ •±Í”€‰U9-9=]8ˆ(€€€€€€€…ÉÕµ•¹Ð€ôµ…Ñ ¹É½ÕÀ È¤¥˜µ…Ñ •±Í”9½¹”(€€€€€€€½Õ¹ÑÍm­¥¹‘t€ô½Õ¹ÑÌ¹•Ð¡­¥¹°€À¤€¬€Ä(€€€€€€€¥˜­¥¹¥¸ì‰Iˆ°€‰AI‰ô…¹…ÉÕµ•¹Ðè(€€€€€€€€€€€É•™•É•¹•Ì¹…ÁÁ•¹¡ì‰ÑåÁ”ˆè­¥¹°€‰Ñ…É•Ðˆè…ÉÕµ•¹Ñô¤(€€€€€€€¥˜­¥¹€ôô€‰MDˆ…¹…ÉÕµ•¹Ðè(€€€€€€€€€€€Í•ÅÕ•¹•Ì¹…ÁÁ•¹¡…ÉÕµ•¹Ð¤((€€€Õ¹É•Í½±Ù•€ôÍ½ÉÑ• (€€€€€€€ì(€€€€€€€€€€€¥Ñ•µl‰Ñ…É•Ð‰t(€€€€€€€€€€€™½È¥Ñ•´¥¸É•™•É•¹•Ì(€€€€€€€€€€€¥˜¥Ñ•µl‰Ñ…É•Ð‰t¹½Ð¥¸‰½½­µ…É­Ì(€€€€€€€ô(€€€€¤(€€€É•ÑÕÉ¸ì(€€€€€€€€‰Ñ½Ñ…°ˆè±•¸¡¥¹ÍÑÉÕÑ¥½¹Ì¤°(€€€€€€€€‰ÑåÁ•Ìˆè‘¥Ð¡Í½ÉÑ•¡½Õ¹ÑÌ¹¥Ñ•µÌ ¤¤¤°(€€€€€€€€‰‰½½­µ…É­ÌˆèÍ½ÉÑ•¡‰½½­µ…É­Ì¤°(€€€€€€€€‰É•™•É•¹•ÌˆèÉ•™•É•¹•Ì°(€€€€€€€€‰Õ¹É•Í½±Ù•‘}É•™•É•¹•ÌˆèÕ¹É•Í½±Ù•°(€€€€€€€€‰Í•ÅÕ•¹•ÌˆèÍ½ÉÑ•¡Í•Ð¡Í•ÅÕ•¹•Ì¤¤°(€€€ô(()‘•˜™¥•±‘}…¡•}¥¹Ù•¹Ñ½Éä¡Á…Ñ èA…Ñ ¤€´ø‘¥ÑmÍÑÈ°¹åtè(€€€¥¹Ù•¹Ñ½Éä€ô™¥•±‘}¥¹Ù•¹Ñ½Éä¡Á…Ñ ¤(€€€Ñ½}™¥•±‘Ì€ô¥¹Ð¡¥¹Ù•¹Ñ½Éål‰ÑåÁ•Ì‰t¹•Ð ‰Q=ˆ°€À¤¤(€€€Ñ½}•¹ÑÉ¥•Ì€ô€À(€€€‘¥ÉÑå}™¥•±‘Ì€ô€À(€€€ÕÁ‘…Ñ•}½¹}½Á•¸€ô…±Í”(€€€Ý¥Ñ é¥Á™¥±”¹i¥Á¥±”¡Á…Ñ ¤…ÌÁ…­…”è(€€€€€€€‘½Õµ•¹Ð€ô•ÑÉ•”¹™É½µÍÑÉ¥¹œ¡Á…­…”¹É•… ‰Ý½É½‘½Õµ•¹Ð¹áµ°ˆ¤¤(€€€€€€€Ñ½}•¹ÑÉ¥•Ì€ô±•¸ (€€€€€€€€€€€‘½Õµ•¹Ð¹áÁ…Ñ  (€€€€€€€€€€€€€€€€ˆ¸¼½ÜéÁmÍÑ…ÉÑÌµÝ¥Ñ ¡ÑÉ…¹Í±…Ñ” ¸½ÜéÁAÈ½ÜéÁMÑå±”½ÜéÙ…°°€Ñ½Œœ°€Q=œ¤°€Q=œ¥tˆ°(€€€€€€€€€€€€€€€¹…µ•ÍÁ…•Ìõ9L°(€€€€€€€€€€€€¤(€€€€€€€€¤(€€€€€€€‘¥ÉÑå}™¥•±‘Ì€ô±•¸ (€€€€€€€€€€€‘½Õµ•¹Ð¹áÁ…Ñ  (€€€€€€€€€€€€€€€€ˆ¸¼½Üé™±‘M¥µÁ±•mÜé‘¥ÉÑäôÑÉÕ”œ½ÈÜé‘¥ÉÑäôœÄtð€ˆ(€€€€€€€€€€€€€€€€ˆ¸¼½Üé™±‘¡…ÉmÜé™±‘¡…ÉQåÁ”ô‰•¥¸umÜé‘¥ÉÑäôÑÉÕ”œ½ÈÜé‘¥ÉÑäôœÄtˆ°(€€€€€€€€€€€€€€€¹…µ•ÍÁ…•Ìõ9L°(€€€€€€€€€€€€¤(€€€€€€€€¤(€€€€€€€¥˜€‰Ý½É½Í•ÑÑ¥¹Ì¹áµ°ˆ¥¸Á…­…”¹¹…µ•±¥ÍÐ ¤è(€€€€€€€€€€€Í•ÑÑ¥¹Ì€ô•ÑÉ•”¹™É½µÍÑÉ¥¹œ¡Á…­…”¹É•… ‰Ý½É½Í•ÑÑ¥¹Ì¹áµ°ˆ¤¤(€€€€€€€€€€€ÕÁ‘…Ñ•}½¹}½Á•¸€ô‰½½°¡Í•ÑÑ¥¹Ì¹áÁ…Ñ  ˆ¸¼½ÜéÕÁ‘…Ñ•¥•±‘Ìˆ°¹…µ•ÍÁ…•Ìõ9L¤¤((€€€¥˜Ñ½}™¥•±‘Ì€ôô€Àè(€€€€€€€ÍÑ…ÑÕÌ€ô€‰…‰Í•¹Ðˆ(€€€•±¥˜Ñ½}•¹ÑÉ¥•Ì€ôô€Àè(€€€€€€€ÍÑ…ÑÕÌ€ô€‰½‘•}½¹±äˆ(€€€•±¥˜‘¥ÉÑå}™¥•±‘Ìè(€€€€€€€ÍÑ…ÑÕÌ€ô€‰ÍÑ…±”ˆ(€€€•±Í”è(€€€€€€€ÍÑ…ÑÕÌ€ô€‰É•™É•Í¡•ˆ(€€€É•ÑÕÉ¸ì(€€€€€€€€‰ÍÑ…ÑÕÌˆèÍÑ…ÑÕÌ°(€€€€€€€€‰µ…¥¹}Ñ½}™¥•±‘ÌˆèÑ½}™¥•±‘Ì°(€€€€€€€€‰Ñ½}•¹ÑÉ¥•ÌˆèÑ½}•¹ÑÉ¥•Ì°(€€€€€€€€‰‘¥ÉÑå}™¥•±‘Ìˆè‘¥ÉÑå}™¥•±‘Ì°(€€€€€€€€‰ÕÁ‘…Ñ•}½¹}½Á•¸ˆèÕÁ‘…Ñ•}½¹}½Á•¸°(€€€€€€€€‰™¥•±‘}ÑåÁ•Ìˆè¥¹Ù•¹Ñ½Éål‰ÑåÁ•Ì‰t°(€€€ô(()‘•˜•ÅÕ…Ñ¥½¹}¥¹Ù•¹Ñ½Éä¡Á…Ñ èA…Ñ ¤€´ø‘¥ÑmÍÑÈ°¹åtè(€€€•¹ÍÕÉ•}‘½à¡Á…Ñ ¤(€€€É•ÍÕ±Ð€ôì(€€€€€€€€‰½µµ°ˆè€À°(€€€€€€€€‰µ…Ñ¡ÑåÁ•}½±”ˆè€À°(€€€€€€€€‰±•…å}•ÅÕ…Ñ¥½¹}½±”ˆè€À°(€€€€€€€€‰½Ñ¡•É}½±”ˆè€À°(€€€€€€€€‰™½ÉµÕ±…}¥µ…•}…¹‘¥‘…Ñ•Ìˆè€À°(€€€ô(€€€Ý¥Ñ é¥Á™¥±”¹i¥Á¥±”¡Á…Ñ ¤…ÌÁ…­…”è(€€€€€€€™½È¹…µ”¥¸Á…­…”¹¹…µ•±¥ÍÐ ¤è(€€€€€€€€€€€¥˜¹½Ð¹…µ”¹ÍÑ…ÉÑÍÝ¥Ñ  ‰Ý½É¼ˆ¤½È¹½Ð¹…µ”¹•¹‘ÍÝ¥Ñ  ˆ¹áµ°ˆ¤è(€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”(€€€€€€€€€€€É½½Ð€ô•ÑÉ•”¹™É½µÍÑÉ¥¹œ¡Á…­…”¹É•…¡¹…µ”¤¤(€€€€€€€€€€€É•ÍÕ±Ñl‰½µµ°‰t€¬ô±•¸¡É½½Ð¹áÁ…Ñ  ˆ¸¼½´é½5…Ñ ˆ°¹…µ•ÍÁ…•Ìõ9L¤¤(€€€€€€€€€€€™½È½±”¥¸É½½Ð¹áÁ…Ñ  ˆ¸¼¼©m±½…°µ¹…µ” ¤ô=1=‰©•Ðtˆ¤è(€€€€€€€€€€€€€€€ÁÉ½}¥€ô€¡½±”¹•Ð ‰AÉ½%ˆ¤½È½±”¹•Ð¡˜‰ííí=}9MõõõAÉ½%ˆ¤½È€ˆˆ¤¹±½Ý•È ¤(€€€€€€€€€€€€€€€¥˜€‰‘ÍµÐˆ¥¸ÁÉ½}¥½È€‰µ…Ñ¡ÑåÁ”ˆ¥¸ÁÉ½}¥è(€€€€€€€€€€€€€€€€€€€É•ÍÕ±Ñl‰µ…Ñ¡ÑåÁ•}½±”‰t€¬ô€Ä(€€€€€€€€€€€€€€€•±¥˜ÁÉ½}¥¥¸ì‰•ÅÕ…Ñ¥½¸¸Ìˆ°€‰•ÅÕ…Ñ¥½¸¸Èˆ°€‰•ÅÕ…Ñ¥½¸‰ôè(€€€€€€€€€€€€€€€€€€€É•ÍÕ±Ñl‰±•…å}•ÅÕ…Ñ¥½¹}½±”‰t€¬ô€Ä(€€€€€€€€€€€€€€€•±Í”è(€€€€€€€€€€€€€€€€€€€É•ÍÕ±Ñl‰½Ñ¡•É}½±”‰t€¬ô€Ä((€€€€€€€€€€€™½ÈÁ…É…É…Á ¥¸É½½Ð¹áÁ…Ñ  ˆ¸¼½ÜéÀˆ°¹…µ•ÍÁ…•Ìõ9L¤è(€€€€€€€€€€€€€€€¥˜¹½ÐÁ…É…É…Á ¹áÁ…Ñ  ˆ¸¼½Üé‘É…Ý¥¹œð€¸¼½Øé¥µ…•‘…Ñ„ˆ°¹…µ•ÍÁ…•Ìõ9L¤è(€€€€€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”(€€€€€€€€€€€€€€€ÍÑå±•Ì€ôÁ…É…É…Á ¹áÁ…Ñ  ˆ¸½ÜéÁAÈ½ÜéÁMÑå±”½ÜéÙ…°ˆ°¹…µ•ÍÁ…•Ìõ9L¤(€€€€€€€€€€€€€€€µ•Ñ…‘…Ñ„€ôÁ…É…É…Á ¹áÁ…Ñ  (€€€€€€€€€€€€€€€€€€€€ˆ¸¼¼©m±½…°µ¹…µ” ¤ô‘½AÈt½¹…µ”ð€ˆ(€€€€€€€€€€€€€€€€€€€€ˆ¸¼¼©m±½…°µ¹…µ” ¤ô‘½AÈt½Ñ¥Ñ±”ð€ˆ(€€€€€€€€€€€€€€€€€€€€ˆ¸¼¼©m±½…°µ¹…µ” ¤ô‘½AÈt½‘•ÍÈˆ(€€€€€€€€€€€€€€€€¤(€€€€€€€€€€€€€€€¡¥¹Ð€ô€ˆ€ˆ¹©½¥¸¡ÍÑå±•Ì€¬µ•Ñ…‘…Ñ„¤¹±½Ý•È ¤(€€€€€€€€€€€€€€€¥˜…¹ä¡Ñ½­•¸¥¸¡¥¹Ð™½ÈÑ½­•¸¥¸€ ‰•ÅÕ…Ñ¥½¸ˆ°€‰™½ÉµÕ±„ˆ°€‹–³–ò<ˆ¤¤è(€€€€€€€€€€€€€€€€€€€É•ÍÕ±Ñl‰™½ÉµÕ±…}¥µ…•}…¹‘¥‘…Ñ•Ì‰t€¬ô€Ä(€€€É•ÍÕ±Ñl‰•‘¥Ñ…‰±•}Ñ½Ñ…°‰t€ô€ (€€€€€€€É•ÍÕ±Ñl‰½µµ°‰t€¬É•ÍÕ±Ñl‰µ…Ñ¡ÑåÁ•}½±”‰t€¬É•ÍÕ±Ñl‰±•…å}•ÅÕ…Ñ¥½¹}½±”‰t(€€€€¤(€€€É•ÑÕÉ¸É•ÍÕ±Ð(()‘•˜ÁÉ½Ñ•Ñ•‘}½‰©•Ñ}µ…¹¥™•ÍÐ¡Á…Ñ èA…Ñ ¤€´ø‘¥ÑmÍÑÈ°¹åtè(€€€•¹ÍÕÉ•}‘½à¡Á…Ñ ¤(€€€•µ‰•‘‘¥¹Ìè‘¥ÑmÍÑÈ°ÍÑÉt€ôíô(€€€µ•‘¥„è‘¥ÑmÍÑÈ°ÍÑÉt€ôíô(€€€½µµ°è±¥ÍÑmÍÑÉt€ômt(€€€Ý¥Ñ é¥Á™¥±”¹i¥Á¥±”¡Á…Ñ ¤…ÌÁ…­…”è(€€€€€€€™½È¹…µ”¥¸Í½ÉÑ•¡Á…­…”¹¹…µ•±¥ÍÐ ¤¤è(€€€€€€€€€€€‘…Ñ„€ôÁ…­…”¹É•…¡¹…µ”¤(€€€€€€€€€€€¥˜¹…µ”¹ÍÑ…ÉÑÍÝ¥Ñ  ‰Ý½É½•µ‰•‘‘¥¹Ì¼ˆ¤è(€€€€€€€€€€€€€€€•µ‰•‘‘¥¹Ím¹…µ•t€ô}Í¡„ÈÔÙ}‰åÑ•Ì¡‘…Ñ„¤(€€€€€€€€€€€•±¥˜¹…µ”¹ÍÑ…ÉÑÍÝ¥Ñ  ‰Ý½É½µ•‘¥„¼ˆ¤è(€€€€€€€€€€€€€€€µ•‘¥…m¹…µ•t€ô}Í¡„ÈÔÙ}‰åÑ•Ì¡‘…Ñ„¤(€€€€€€€€€€€¥˜¹…µ”¹ÍÑ…ÉÑÍÝ¥Ñ  ‰Ý½É¼ˆ¤…¹¹…µ”¹•¹‘ÍÝ¥Ñ  ˆ¹áµ°ˆ¤è(€€€€€€€€€€€€€€€É½½Ð€ô•ÑÉ•”¹™É½µÍÑÉ¥¹œ¡‘…Ñ„¤(€€€€€€€€€€€€€€€™½È•ÅÕ…Ñ¥½¸¥¸É½½Ð¹áÁ…Ñ  ˆ¸¼½´é½5…Ñ ˆ°¹…µ•ÍÁ…•Ìõ9L¤è(€€€€€€€€€€€€€€€€€€€…¹½¹¥…°€ô•ÑÉ•”¹Ñ½ÍÑÉ¥¹œ¡•ÅÕ…Ñ¥½¸°µ•Ñ¡½ô‰ŒÄÑ¸ˆ°•á±ÕÍ¥Ù”õQÉÕ”¤(€€€€€€€€€€€€€€€€€€€½µµ°¹…ÁÁ•¹¡}Í¡„ÈÔÙ}‰åÑ•Ì¡…¹½¹¥…°¤¤(€€€É•ÑÕÉ¸ì(€€€€€€€€‰•µ‰•‘‘¥¹Ìˆè•µ‰•‘‘¥¹Ì°(€€€€€€€€‰µ•‘¥„ˆèµ•‘¥„°(€€€€€€€€‰½µµ±}Í¡„ÈÔØˆèÍ½ÉÑ•¡½µµ°¤°(€€€ô(()‘•˜ÁÉ½Ñ•Ñ•‘}Á…å±½…‘}µ…¹¥™•ÍÐ¡Á…Ñ èA…Ñ ¤€´ø‘¥ÑmÍÑÈ°±¥ÍÑmÍÑÉutè(€€€€ˆˆ‰½µÁ…É”ÁÉ½Ñ•Ñ•Á…å±½…‘Ì¥¹‘•Á•¹‘•¹Ñ±ä½˜Á…­…”Á…ÉÐÉ•¹…µ¥¹œ¸ˆˆˆ(€€€µ…¹¥™•ÍÐ€ôÁÉ½Ñ•Ñ•‘}½‰©•Ñ}µ…¹¥™•ÍÐ¡Á…Ñ ¤(€€€É•ÑÕÉ¸ì(€€€€€€€€‰•µ‰•‘‘¥¹ÌˆèÍ½ÉÑ•¡µ…¹¥™•ÍÑl‰•µ‰•‘‘¥¹Ì‰t¹Ù…±Õ•Ì ¤¤°(€€€€€€€€‰µ•‘¥„ˆèÍ½ÉÑ•¡µ…¹¥™•ÍÑl‰µ•‘¥„‰t¹Ù…±Õ•Ì ¤¤°(€€€€€€€€‰½µµ±}Í¡„ÈÔØˆèÍ½ÉÑ•¡µ…¹¥™•ÍÑl‰½µµ±}Í¡„ÈÔØ‰t¤°(€€€ô()‘•˜…ÁÁ±å}ÉÕ±” (€€€‘½Õµ•¹Ðè¹ä°(€€€ÉÕ±”è‘¥ÑmÍÑÈ°¹åt°(€€€€¨°(€€€Á…É…É…Á¡}Ñ…É•ÑÌè±¥ÍÑm¹åtð9½¹”€ô9½¹”°(€€€Ñ…‰±•}Ñ…É•ÑÌè±¥ÍÑmÑÕÁ±•m¹ä°‘¥ÑmÍÑÈ°¹åuutð9½¹”€ô9½¹”°(€€€¡…ÁÑ•É}ÍÑ…ÉÐè¥¹Ðð9½¹”€ô9½¹”°(¤€´ø¥¹Ðè(€€€Õ¹ÍÕÁÁ½ÉÑ•€ôÕ¹ÍÕÁÁ½ÉÑ•‘}ÁÉ½Á•ÉÑ¥•Ì¡ÉÕ±”¤(€€€¥˜Õ¹ÍÕÁÁ½ÉÑ•è(€€€€€€€É…¥Í”½Éµ…Ñ5½¹½É…Á¡ÉÉ½È (€€€€€€€€€€€˜‰IÕ±”íÉÕ±•l¥uô¡…ÌÕ¹ÍÕÁÁ½ÉÑ•…ÕÑ½µ…Ñ¥ŒÁÉ½Á•ÉÑ¥•Ìèìœ°€œ¹©½¥¸¡Õ¹ÍÕÁÁ½ÉÑ•¥ôˆ(€€€€€€€€¤((€€€Í•±•Ñ½È€ôÉÕ±•l‰Í•±•Ñ½È‰t(€€€¥˜Í•±•Ñ½Él‰­¥¹‰t¥¸ì‰‘½Õµ•¹Ðˆ°€‰Í•Ñ¥½¹}É½±”‰ôè(€€€€€€€É•ÑÕÉ¸…ÁÁ±å}Í•Ñ¥½¹}ÁÉ½Á•ÉÑ¥•Ì¡‘½Õµ•¹Ð°ÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t¤(€€€¥˜Í•±•Ñ½Él‰­¥¹‰t€ôô€‰Ñ…‰±•}É½±”ˆè(€€€€€€€É•ÑÕÉ¸…ÁÁ±å}Ñ…‰±•}ÁÉ½Á•ÉÑ¥•Ì¡‘½Õµ•¹Ð°ÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t°Ñ…‰±•}Ñ…É•ÑÌ¤(€€€¥˜Í•±•Ñ½Él‰­¥¹‰t€ôô€‰™¥•±‘}É½±”ˆè(€€€€€€€ÁÉ½Á•ÉÑ¥•Ì€ô‘¥Ð¡ÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t¤(€€€€€€€¥˜¡…ÁÑ•É}ÍÑ…ÉÐ¥Ì¹½Ð9½¹”è(€€€€€€€€€€€ÁÉ½Á•ÉÑ¥•Íl‰¡…ÁÑ•É}ÍÑ…ÉÐ‰t€ô¡…ÁÑ•É}ÍÑ…ÉÐ(€€€€€€€É•ÑÕÉ¸…ÁÁ±å}™¥•±‘}ÁÉ½Á•ÉÑ¥•Ì¡‘½Õµ•¹Ð°ÁÉ½Á•ÉÑ¥•Ì¤(€€€¥˜Í•±•Ñ½Él‰­¥¹‰t€ôô€‰•ÅÕ…Ñ¥½¹}É½±”ˆè(€€€€€€€É•ÑÕÉ¸…ÁÁ±å}•ÅÕ…Ñ¥½¹}ÁÉ½Á•ÉÑ¥•Ì¡‘½Õµ•¹Ð°ÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t¤((€€€ÍÑå±•}¹…µ”€ôÍÑå±•}¹…µ•}™½É}Í•±•Ñ½È¡Í•±•Ñ½È¤(€€€¥˜ÍÑå±•}¹…µ”¥Ì9½¹”è(€€€€€€€É…¥Í”½Éµ…Ñ5½¹½É…Á¡ÉÉ½È (€€€€€€€€€€€˜‰IÕ±”íÉÕ±•l¥uôÕÍ•Ì…¸Õ¹ÍÕÁÁ½ÉÑ•…ÕÑ½µ…Ñ¥ŒÍ•±•Ñ½ÈèíÍ•±•Ñ½Éôˆ(€€€€€€€€¤(€€€¥˜Á…É…É…Á¡}Ñ…É•ÑÌ¥Ì¹½Ð9½¹”è(€€€€€€€É•ÑÕÉ¸…ÁÁ±å}ÍÑå±•}ÉÕ±•}Ñ½}Á…É…É…Á¡Ì¡‘½Õµ•¹Ð°ÉÕ±”°Á…É…É…Á¡}Ñ…É•ÑÌ¤(€€€ÍÑå±”€ô•¹ÍÕÉ•}Á…É…É…Á¡}ÍÑå±”¡‘½Õµ•¹Ð°ÍÑå±•}¹…µ”¤(€€€…ÁÁ±å}ÍÑå±•}ÁÉ½Á•ÉÑ¥•Ì¡ÍÑå±”°ÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t¤(€€€É•ÑÕÉ¸ÍÕ´ (€€€€€€€€Ä(€€€€€€€™½ÈÁ…É…É…Á ¥¸¥Ñ•É}‘½Õµ•¹Ñ}Á…É…É…Á¡Ì¡‘½Õµ•¹Ð¤(€€€€€€€¥˜Á…É…É…Á ¹ÍÑå±”…¹Á…É…É…Á ¹ÍÑå±”¹¹…µ”€ôôÍÑå±•}¹…µ”(€€€€¤(()‘•˜™¥ÉÍÑ}…¹¡½É}Á…É…É…Á ¡‘½Õµ•¹Ðè¹ä°ÉÕ±”è‘¥ÑmÍÑÈ°¹åt¤€´ø¹äð9½¹”è(€€€ÍÑå±•}¹…µ”€ôÍÑå±•}¹…µ•}™½É}Í•±•Ñ½È¡ÉÕ±•l‰Í•±•Ñ½È‰t¤(€€€Á…É…É…Á¡Ì€ô±¥ÍÐ¡‘½Õµ•¹Ð¹Á…É…É…Á¡Ì¤(€€€¥˜ÍÑå±•}¹…µ”è(€€€€€€€™½ÈÁ…É…É…Á ¥¸Á…É…É…Á¡Ìè(€€€€€€€€€€€¥˜€ (€€€€€€€€€€€€€€€Á…É…É…Á ¹Ñ•áÐ¹ÍÑÉ¥À ¤(€€€€€€€€€€€€€€€…¹Á…É…É…Á ¹ÍÑå±”(€€€€€€€€€€€€€€€…¹Á…É…É…Á ¹ÍÑå±”¹¹…µ”€ôôÍÑå±•}¹…µ”(€€€€€€€€€€€€€€€…¹Á…É…É…Á ¹ÉÕ¹Ì(€€€€€€€€€€€€¤è(€€€€€€€€€€€€€€€É•ÑÕÉ¸Á…É…É…Á (€€€™½ÈÁ…É…É…Á ¥¸Á…É…É…Á¡Ìè(€€€€€€€¥˜Á…É…É…Á ¹Ñ•áÐ¹ÍÑÉ¥À ¤…¹Á…É…É…Á ¹ÉÕ¹Ìè(€€€€€€€€€€€É•ÑÕÉ¸Á…É…É…Á (€€€É•ÑÕÉ¸9½¹”(()‘•˜ÍÕµµ…É¥é•}ÉÕ±”¡ÉÕ±”è‘¥ÑmÍÑÈ°¹åt¤€´øÍÑÈè(€€€Á…¥ÉÌ€ô€ˆ°€ˆ¹©½¥¸¡˜‰í­•åôõíÙ…±Õ•ôˆ™½È­•ä°Ù…±Õ”¥¸Í½ÉÑ•¡ÉÕ±•l‰ÁÉ½Á•ÉÑ¥•Ì‰t¹¥Ñ•µÌ ¤¤¤(€€€Í½ÕÉ•Ì€ô€ˆ°€ˆ¹©½¥¸¡ÉÕ±•l‰Í½ÕÉ•}¥‘Ì‰t¤(€€€É•ÑÕÉ¸˜‰íÉÕ±•l¥uôèíÁ…¥ÉÍô¸M½ÕÉ•ÌèíÍ½ÕÉ•Íô¸ˆ(()‘•˜±½…‘}‘½Õµ•¹Ð¡Á…Ñ èA…Ñ ¤€´ø¹äè(€€€•¹ÍÕÉ•}‘½à¡Á…Ñ ¤(€€€ÑÉäè(€€€€€€€É•ÑÕÉ¸½Õµ•¹Ð¡ÍÑÈ¡Á…Ñ ¤¤(€€€•á•ÁÐá•ÁÑ¥½¸…Ì•áŒè(€€€€€€€É…¥Í”½Éµ…Ñ5½¹½É…Á¡ÉÉ½È¡˜‰U¹…‰±”Ñ¼½Á•¸=`íÁ…Ñ¡ôèí•áôˆ¤™É½´•áŒ(