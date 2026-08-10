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
    "toc_level_1": "TOC 1",
    "toc_level_2": "TOC 2",
    "toc_level_3": "TOC 3",
    "table_of_contents_level_1": "TOC 1",
    "table_of_contents_level_2": "TOC 2",
    "table_of_contents_level_3": "TOC 3",
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
    "available_width_percent",
    "preferred_column_widths_percent",
    "allow_autofit",
    "cell_margins_mm",
    "vertical_alignment",
    "border_preset",
    "column_roles",
    "column_alignments",
    "header_bold",
    "header_shading_hex",
    "font_name_ascii",
    "font_name_east_asia",
    "font_name_complex_script",
    "font_size_pt",
    "line_spacing_pt",
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

FONT_THEME_ATTRIBUTES = {
    "ascii": "asciiTheme",
    "hAnsi": "hAnsiTheme",
    "eastAsia": "eastAsiaTheme",
    "cs": "cstheme",
}

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


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
            counts["drawings"] += len(root.xpath(".//w:dr×ouÞÚ$z{-®éÜj×&WGW&âÖ‚‡fÇVW2ÂFVfVÇCÓ’²Ð Ð Ð¦FVböVç7W&Uö†VF–æuöçVÖ&W&–ær†Fö7VÖVçC¢ç’ÂÆWfVÇ3¢–çBÂ6†FW%÷7F'C¢–çBÒ’Óâ–çC Ð¢–bæ÷BÃÒÆWfVÇ2ÃÒC Ð¢&—6Rf÷&ÖDÖöæöw&„W'&÷"‚&†VF–æuöÆWfVÇ2×W7B&R&WGvVVâæBBâ"Ð¢–b6†FW%÷7F'BÂ Ð¢&—6Rf÷&ÖDÖöæöw&„W'&÷"‚&6†FW%÷7F'B×W7B&R÷6—F—fR–çFVvW"â"Ð¢&ö÷BÒFö7VÖVçBç'BæçVÖ&W&–æu÷'BæVÆVÖVç@Ð¢'7G&7Eö–BÒöæW‡EöçVÖ&W&–æuö–B‡&ö÷BÂ&'7G&7DçVÒ"Â&'7G&7DçVÔ–B"Ð¢çVÕö–BÒöæW‡EöçVÖ&W&–æuö–B‡&ö÷BÂ&çVÒ"Â&çVÔ–B"Ð Ð¢'7G&7BÒ÷†ÖÄVÆVÖVçB‚'s¦'7G&7DçVÒ"Ð¢'7G&7Bç6WB‡â‚'s¦'7G&7DçVÔ–B"’Â7G"†'7G&7Eö–B’Ð¢×VÇF’Ò÷†ÖÄVÆVÖVçB‚'s¦×VÇF”ÆWfVÅG—R"Ð¢×VÇF’ç6WB‡â‚'s§fÂ"’Â&×VÇF–ÆWfVÂ"Ð¢'7G&7BæVæB†×VÇF’Ð Ð¢f÷"ÆWfVÂ–â&ævR†ÆWfVÇ2“ Ð¢ÇfÂÒ÷†ÖÄVÆVÖVçB‚'s¦ÇfÂ"Ð¢ÇfÂç6WB‡â‚'s¦–ÇfÂ"’Â7G"†ÆWfVÂ’Ð¢7F'BÒ÷†ÖÄVÆVÖVçB‚'s§7F'B"Ð¢7F'Bç6WB‡â‚'s§fÂ"’Â7G"†6†FW%÷7F'B–bÆWfVÂÓÒVÇ6R’Ð¢çVÕöf×BÒ÷†ÖÄVÆVÖVçB‚'s¦çVÔf×B"Ð¢çVÕöf×Bç6WB‡â‚'s§fÂ"’Â&FV6–ÖÂ"Ð¢÷7G–ÆRÒ÷†ÖÄVÆVÖVçB‚'s§7G–ÆR"Ð¢÷7G–ÆRç6WB‡â‚'s§fÂ"’Âb$†VF–æw¶ÆWfVÂ²Ò"Ð¢ÇfÅ÷FW‡BÒ÷†ÖÄVÆVÖVçB‚'s¦ÇfÅFW‡B"Ð¢–bÆWfVÂÓÒ Ð¢ÇfÅ÷FW‡Bç6WB‡â‚'s§fÂ"’Â.zÊÂSzº"Ð¢VÇ6S Ð¢ÇfÅ÷FW‡Bç6WB€Ð¢â‚'s§fÂ"’ÀÐ¢"â"æ¦ö–â†b"W¶çVÖ&W'Ò"f÷"çVÖ&W"–â&ævRƒÂÆWfVÂ²"’’ÀÐ¢Ð¢7VfbÒ÷†ÖÄVÆVÖVçB‚'s§7Vfb"Ð¢7Vfbç6WB‡â‚'s§fÂ"’Â'76R"Ð¢ÇfÂæW‡FVæB…·7F'BÂçVÕöf×BÂ÷7G–ÆRÂÇfÅ÷FW‡BÂ7VfeÒÐ¢'7G&7BæVæB†ÇfÂÐ¢&ö÷Bæ–ç6W'BƒÂ'7G&7BÐ Ð¢çVÒÒ÷†ÖÄVÆVÖVçB‚'s¦çVÒ"Ð¢çVÒç6WB‡â‚'s¦çVÔ–B"’Â7G"†çVÕö–B’Ð¢'7G&7E÷&VbÒ÷†ÖÄVÆVÖVçB‚'s¦'7G&7DçVÔ–B"Ð¢'7G&7E÷&Vbç6WB‡â‚'s§fÂ"’Â7G"†'7G&7Eö–B’Ð¢çVÒæVæB†'7G&7E÷&VbÐ¢–b6†FW%÷7F'BÒ Ð¢÷fW'&–FRÒ÷†ÖÄVÆVÖVçB‚'s¦ÇfÄ÷fW'&–FR"Ð¢÷fW'&–FRç6WB‡â‚'s¦–ÇfÂ"’Â#"Ð¢7F'Eö÷fW'&–FRÒ÷†ÖÄVÆVÖVçB‚'s§7F'D÷fW'&–FR"Ð¢7F'Eö÷fW'&–FRç6WB‡â‚'s§fÂ"’Â7G"†6†FW%÷7F'B’Ð¢÷fW'&–FRæVæB‡7F'Eö÷fW'&–FRÐ¢çVÒæVæB†÷fW'&–FRÐ¢&ö÷BæVæB†çVÒÐ Ð¢f÷"ÆWfVÂ–â&ævR†ÆWfVÇ2“ Ð¢7G–ÆRÒVç7W&U÷&w&…÷7G–ÆR†Fö7VÖVçBÂb$†VF–ær¶ÆWfVÂ²Ò"Ð¢÷"Ò7G–ÆRæVÆVÖVçBævWEö÷%öFE÷"‚Ð¢çVÕ÷"Ò÷"æf–æB‡â‚'s¦çVÕ""’Ð¢–bçVÕ÷"—2æöæS Ð¢çVÕ÷"Ò÷†ÖÄVÆVÖVçB‚'s¦çVÕ""Ð¢÷"æVæB†çVÕ÷"Ð¢f÷"6†–ÆEöæÖR–â‚&–ÇfÂ"Â&çVÔ–B"“ Ð¢W†—7F–ærÒçVÕ÷"æf–æB‡â†b's§¶6†–ÆEöæÖWÒ"’Ð¢–bW†—7F–ær—2æ÷BæöæS Ð¢çVÕ÷"ç&VÖ÷fR†W†—7F–ærÐ¢–ÇfÂÒ÷†ÖÄVÆVÖVçB‚'s¦–ÇfÂ"Ð¢–ÇfÂç6WB‡â‚'s§fÂ"’Â7G"†ÆWfVÂ’Ð¢çVÖ&W"Ò÷†ÖÄVÆVÖVçB‚'s¦çVÔ–B"Ð¢çVÖ&W"ç6WB‡â‚'s§fÂ"’Â7G"†çVÕö–B’Ð¢çVÕ÷"æW‡FVæB…¶–ÇfÂÂçVÖ&W%ÒÐ¢&WGW&âÆWfVÇ0Ð Ð Ð¦FVbÇ•öf–VÆE÷&÷W'F–W2†Fö7VÖVçC¢ç’Â&÷W'F–W3¢F–7E·7G"Âç•Ò’Óâ–çC Ð¢6†ævVBÒ Ð¢–b'WFFUööåö÷Vâ"–â&÷W'F–W3 Ð¢÷6WE÷WFFUöf–VÆG5ööåö÷Vâ†Fö7VÖVçBÂ&ööÂ‡&÷W'F–W5²'WFFUööåö÷Vâ%Ò’Ð¢6†ævVB³ÒÐ¢–b&÷W'F–W2ævWB‚&Ö&µöf–VÆG5öF—'G’"“ Ð¢6†ævVB³ÒöÖ&µöf–VÆG5öF—'G’†Fö7VÖVçBÐ¢–b&÷W'F–W2ævWB‚&6öçfW'EöW‡Æ–6—EöÖ&¶W'2"“ Ð¢6†ævVB³Òö6öçfW'EöW‡Æ–6—Eöf–VÆEöÖ&¶W'2†Fö7VÖVçBÐ Ð¢ÆWfVÇ2Ò–çB‡&÷W'F–W2ævWB‚&†VF–æuöÆWfVÇ2"ÂB’Ð¢–b&÷W'F–W2ævWB‚'7G&—öÖçVÅö†VF–æu÷&Vf—†W2"“ Ð¢6†ævVB³Ò÷7G&—öÖçVÅö†VF–æu÷&Vf—†W2†Fö7VÖVçBÂÆWfVÇ2Ð¢–b&÷W'F–W2ævWB‚'&V'V–ÆEö†VF–æuöçVÖ&W&–ær"“ Ð¢6†ævVB³ÒöVç7W&Uö†VF–æuöçVÖ&W&–ær€Ð¢Fö7VÖVçBÂÆWfVÇ2Â–çB‡&÷W'F–W2ævWB‚&6†FW%÷7F'B"Â’Ð¢Ð¢&WGW&â6†ævV@Ð Ð Ð¦FVbÇ•öWVF–öå÷&÷W'F–W2†Fö7VÖVçC¢ç’Â&÷W'F–W3¢F–7E·7G"Âç•Ò’Óâ–çC Ð¢Vç7W÷'FVE÷fÇVW2Ò°Ð¢¶W“¢fÇVPÐ¢f÷"¶W’ÂfÇVR–â&÷W'F–W2æ—FV×2‚Ð¢–b¶W’–âUTD”ôåõ$õU%D”U2æBfÇVRæ÷B–âµG'VRÂfÇ6WÐÐ¢ÐÐ¢–bVç7W÷'FVE÷fÇVW3 Ð¢&—6Rf÷&ÖDÖöæöw&„W'&÷"€Ð¢$WVF–öâöÆ–7’&÷W'F–W2×W7B&R&ööÆVã¢ Ð¢²"Â"æ¦ö–â‡6÷'FVB‡Vç7W÷'FVE÷fÇVW2’Ð¢Ð¢&WGW&âÆVâ†Fö7VÖVçBæVÆVÖVçBç‡F‚‚"âòò¥¶Æö6ÂÖæÖR‚“ÒvôÖF‚uÒ"’Ð Ð Ð¦FVb÷6†#Seö'—FW2‡fÇVS¢'—FW2’Óâ7G# Ð¢&WGW&â†6†Æ–"ç6†#Sb‡fÇVR’æ†W†F–vW7B‚Ð Ð Ð¦FVbö6ö×ÆW…öf–VÆEö–ç7G'V7F–öç2‡&ö÷C¢WG&VRåôVÆVÖVçB’ÓâÆ—7E·7G%Ó Ð¢–ç7G'V7F–öç3¢Æ—7E·7G%ÒÒµÐÐ¢f÷"&w&‚–â&ö÷Bç‡F‚‚"âò÷s§"ÂæÖW76W3Ôå2“ Ð¢7F6³¢Æ—7E¶F–7E·7G"Âç•ÕÒÒµÐÐ¢f÷"VÆVÖVçB–â&w&‚æ—FW"‚“ Ð¢–bVÆVÖVçBçFrÓÒâ‚'s¦fÆD6†""“ Ð¢¶–æBÒVÆVÖVçBævWB‡â‚'s¦fÆD6†%G—R"’Ð¢–b¶–æBÓÒ&&Vv–â# Ð¢7F6²æVæB‡²''G2#¢µÒÂ&6GW&VB#¢fÇ6WÒÐ¢VÆ–b¶–æBÓÒ'6W&FR"æB7F6²æBæ÷B7F6µ²ÓÕ²&6GW&VB%Ó Ð¢fÇVRÒ""æ¦ö–â‡7F6µ²ÓÕ²''G2%Ò’ç7G&—‚Ð¢–bfÇVS Ð¢–ç7G'V7F–öç2æVæB‡fÇVRÐ¢7F6µ²ÓÕ²&6GW&VB%ÒÒG'VPÐ¢VÆ–b¶–æBÓÒ&VæB"æB7F6³ Ð¢f–VÆBÒ7F6²ç÷‚Ð¢–bæ÷Bf–VÆE²&6GW&VB%Ó Ð¢fÇVRÒ""æ¦ö–â†f–VÆE²''G2%Ò’ç7G&—‚Ð¢–bfÇVS Ð¢–ç7G'V7F–öç2æVæB‡fÇVRÐ¢VÆ–b€Ð¢VÆVÖVçBçFrÓÒâ‚'s¦–ç7G%FW‡B"Ð¢æB7F6°Ð¢æBæ÷B7F6µ²ÓÕ²&6GW&VB%ÐÐ¢“ Ð¢7F6µ²ÓÕ²''G2%ÒæVæB†VÆVÖVçBçFW‡B÷"""Ð¢&WGW&â–ç7G'V7F–öç0Ð Ð Ð¦FVbf–VÆEö–çfVçF÷'’‡Fƒ¢F‚’ÓâF–7E·7G"Âç•Ó Ð¢Vç7W&UöFö7‚‡F‚Ð¢6÷VçG3¢F–7E·7G"Â–çEÒÒ·ÐÐ¢–ç7G'V7F–öç3¢Æ—7E·7G%ÒÒµÐÐ¢&öö¶Ö&·3¢6WE·7G%ÒÒ6WB‚Ð¢v—F‚¦—f–ÆRå¦—f–ÆR‡F‚’26¶vS Ð¢f÷"æÖR–â6¶vRææÖVÆ—7B‚“ Ð¢–bæ÷BæÖRç7F'G7v—F‚‚'v÷&Bò"’÷"æ÷BæÖRæVæG7v—F‚‚"ç†ÖÂ"“ Ð¢6öçF–çVPÐ¢&ö÷BÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB†æÖR’Ð¢&öö¶Ö&·2çWFFR€Ð¢&ö÷Bç‡F‚‚"âò÷s¦&öö¶Ö&µ7F'Bôs¦æÖR"ÂæÖW76W3Ôå2Ð¢Ð¢–ç7G'V7F–öç2æW‡FVæB€Ð¢fÇVRç7G&—‚Ð¢f÷"fÇVR–â&ö÷Bç‡F‚‚"âò÷s¦fÆE6–×ÆRôs¦–ç7G""ÂæÖW76W3Ôå2Ð¢–bfÇVRç7G&—‚Ð¢Ð¢–ç7G'V7F–öç2æW‡FVæB…ö6ö×ÆW…öf–VÆEö–ç7G'V7F–öç2‡&ö÷B’Ð Ð¢&VfW&Væ6W2ÒµÐÐ¢6WVVæ6W2ÒµÐÐ¢f÷"–ç7G'V7F–öâ–â–ç7G'V7F–öç3 Ð¢ÖF6‚Ò&RæÖF6‚‡"%Ç2¢…´Õ¦×¥Ò²’ƒó¥Ç2²…µåÅÅÇ5Ò²’“ò"Â–ç7G'V7F–öâÐ¢¶–æBÒÖF6‚æw&÷Wƒ’çWW"‚’–bÖF6‚VÇ6R%Tä´äõtâ Ð¢&wVÖVçBÒÖF6‚æw&÷Wƒ"’–bÖF6‚VÇ6RæöæPÐ¢6÷VçG5¶¶–æEÒÒ6÷VçG2ævWB†¶–æBÂ’²Ð¢–b¶–æB–â²%$Tb"Â%tU$Tb'ÒæB&wVÖVçC Ð¢&VfW&Væ6W2æVæB‡²'G—R#¢¶–æBÂ'F&vWB#¢&wVÖVçGÒÐ¢–b¶–æBÓÒ%4U"æB&wVÖVçC Ð¢6WVVæ6W2æVæB†&wVÖVçBÐ Ð¢Vç&W6öÇfVBÒ6÷'FVB€Ð¢°Ð¢—FVÕ²'F&vWB%ÐÐ¢f÷"—FVÒ–â&VfW&Væ6W0Ð¢–b—FVÕ²'F&vWB%Òæ÷B–â&öö¶Ö&·0Ð¢ÐÐ¢Ð¢&WGW&â°Ð¢'F÷FÂ#¢ÆVâ†–ç7G'V7F–öç2’ÀÐ¢'G—W2#¢F–7B‡6÷'FVB†6÷VçG2æ—FV×2‚’’’ÀÐ¢&&öö¶Ö&·2#¢6÷'FVB†&öö¶Ö&·2’ÀÐ¢'&VfW&Væ6W2#¢&VfW&Væ6W2ÀÐ¢'Vç&W6öÇfVE÷&VfW&Væ6W2#¢Vç&W6öÇfVBÀÐ¢'6WVVæ6W2#¢6÷'FVB‡6WB‡6WVVæ6W2’’ÀÐ¢ÐÐ Ð Ð¦FVbf–VÆEö66†Uö–çfVçF÷'’‡Fƒ¢F‚’ÓâF–7E·7G"Âç•Ó Ð¢–çfVçF÷'’Òf–VÆEö–çfVçF÷'’‡F‚Ð¢Fö5öf–VÆG2Ò–çB†–çfVçF÷'•²'G—W2%ÒævWB‚%Dô2"Â’Ð¢Fö5öVçG&–W2Ò Ð¢F—'G•öf–VÆG2Ò Ð¢WFFUööåö÷VâÒfÇ6PÐ¢v—F‚¦—f–ÆRå¦—f–ÆR‡F‚’26¶vS Ð¢Fö7VÖVçBÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‚'v÷&BöFö7VÖVçBç†ÖÂ"’Ð¢Fö5öVçG&–W2ÒÆVâ€Ð¢Fö7VÖVçBç‡F‚€Ð¢"âò÷s§·7F'G2×v—F‚‡G&ç6ÆFR‚â÷s§"÷s§7G–ÆRôs§fÂÂwFö2rÂuDô2r’ÂuDô2r•Ò"ÀÐ¢æÖW76W3Ôå2ÀÐ¢Ð¢Ð¢F—'G•öf–VÆG2ÒÆVâ€Ð¢Fö7VÖVçBç‡F‚€Ð¢"âò÷s¦fÆE6–×ÆU´s¦F—'G“ÒwG'VRr÷"s¦F—'G“ÒsuÒÂ Ð¢"âò÷s¦fÆD6†%´s¦fÆD6†%G—SÒv&Vv–âuÕ´s¦F—'G“ÒwG'VRr÷"s¦F—'G“ÒsuÒ"ÀÐ¢æÖW76W3Ôå2ÀÐ¢Ð¢Ð¢–b'v÷&B÷6WGF–æw2ç†ÖÂ"–â6¶vRææÖVÆ—7B‚“ Ð¢6WGF–æw2ÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB‚'v÷&B÷6WGF–æw2ç†ÖÂ"’Ð¢WFFUööåö÷VâÒ&ööÂ‡6WGF–æw2ç‡F‚‚"âò÷s§WFFTf–VÆG2"ÂæÖW76W3Ôå2’Ð Ð¢–bFö5öf–VÆG2ÓÒ Ð¢7FGW2Ò&'6VçB Ð¢VÆ–bFö5öVçG&–W2ÓÒ Ð¢7FGW2Ò&6öFUööæÇ’ Ð¢VÆ–bF—'G•öf–VÆG3 Ð¢7FGW2Ò'7FÆR Ð¢VÇ6S Ð¢7FGW2Ò'&Vg&W6†VB Ð¢&WGW&â°Ð¢'7FGW2#¢7FGW2ÀÐ¢&Ö–å÷Fö5öf–VÆG2#¢Fö5öf–VÆG2ÀÐ¢'Fö5öVçG&–W2#¢Fö5öVçG&–W2ÀÐ¢&F—'G•öf–VÆG2#¢F—'G•öf–VÆG2ÀÐ¢'WFFUööåö÷Vâ#¢WFFUööåö÷VâÀÐ¢&f–VÆE÷G—W2#¢–çfVçF÷'•²'G—W2%ÒÀÐ¢ÐÐ Ð Ð¦FVbWVF–öåö–çfVçF÷'’‡Fƒ¢F‚’ÓâF–7E·7G"Âç•Ó Ð¢Vç7W&UöFö7‚‡F‚Ð¢&W7VÇBÒ°Ð¢&öÖÖÂ#¢ÀÐ¢&ÖF‡G—UööÆR#¢ÀÐ¢&ÆVv7•öWVF–öåööÆR#¢ÀÐ¢&÷F†W%ööÆR#¢ÀÐ¢&f÷&×VÆö–ÖvUö6æF–FFW2#¢ÀÐ¢ÐÐ¢v—F‚¦—f–ÆRå¦—f–ÆR‡F‚’26¶vS Ð¢f÷"æÖR–â6¶vRææÖVÆ—7B‚“ Ð¢–bæ÷BæÖRç7F'G7v—F‚‚'v÷&Bò"’÷"æ÷BæÖRæVæG7v—F‚‚"ç†ÖÂ"“ Ð¢6öçF–çVPÐ¢&ö÷BÒWG&VRæg&ö×7G&–ær‡6¶vRç&VB†æÖR’Ð¢&W7VÇE²&öÖÖÂ%Ò³ÒÆVâ‡&ö÷Bç‡F‚‚"âòöÓ¦ôÖF‚"ÂæÖW76W3Ôå2’Ð¢f÷"öÆR–â&ö÷Bç‡F‚‚"âòò¥¶Æö6ÂÖæÖR‚“ÒtôÄTö&¦V7BuÒ"“ Ð¢&öuö–BÒ†öÆRævWB‚%&öt”B"’÷"öÆRævWB†b'··´õôå7××Õ&öt”B"’÷"""’æÆ÷vW"‚Ð¢–b&G6×B"–â&öuö–B÷"&ÖF‡G—R"–â&öuö–C Ð¢&W7VÇE²&ÖF‡G—UööÆR%Ò³ÒÐ¢VÆ–b&öuö–B–â²&WVF–öâã2"Â&WVF–öâã""Â&WVF–öâ'Ó Ð¢&W7VÇE²&ÆVv7•öWVF–öåööÆR%Ò³ÒÐ¢VÇ6S Ð¢&W7VÇE²&÷F†W%ööÆR%Ò³ÒÐ Ð¢f÷"&w&‚–â&ö÷Bç‡F‚‚"âò÷s§"ÂæÖW76W3Ôå2“ Ð¢–bæ÷B&w&‚ç‡F‚‚"âò÷s¦G&v–ærÂâò÷c¦–ÖvVFF"ÂæÖW76W3Ôå2“ Ð¢6öçF–çVPÐ¢7G–ÆW2Ò&w&‚ç‡F‚‚"â÷s§"÷s§7G–ÆRôs§fÂ"ÂæÖW76W3Ôå2Ð¢ÖWFFFÒ&w&‚ç‡F‚€Ð¢"âòò¥¶Æö6ÂÖæÖR‚“ÒvFö5"uÒôæÖRÂ Ð¢"âòò¥¶Æö6ÂÖæÖR‚“ÒvFö5"uÒôF—FÆRÂ Ð¢"âòò¥¶Æö6ÂÖæÖR‚“ÒvFö5"uÒôFW67" Ð¢Ð¢†–çBÒ""æ¦ö–â‡7G–ÆW2²ÖWFFF’æÆ÷vW"‚Ð¢–bç’‡Fö¶Vâ–â†–çBf÷"Fö¶Vâ–â‚&WVF–öâ"Â&f÷&×VÆ"Â.XZÎ[Èò"’“ Ð¢&W7VÇE²&f÷&×VÆö–ÖvUö6æF–FFW2%Ò³ÒÐ¢&W7VÇE²&VF—F&ÆU÷F÷FÂ%ÒÒ€Ð¢&W7VÇE²&öÖÖÂ%Ò²&W7VÇE²&ÖF‡G—UööÆR%Ò²&W7VÇE²&ÆVv7•öWVF–öåööÆR%ÐÐ¢Ð¢&WGW&â&W7VÇ@Ð Ð Ð¦FVb&÷FV7FVEöö&¦V7EöÖæ–fW7B‡Fƒ¢F‚’ÓâF–7E·7G"Âç•Ó Ð¢Vç7W&UöFö7‚‡F‚Ð¢VÖ&VFF–æw3¢F–7E·7G"Â7G%ÒÒ·ÐÐ¢ÖVF–¢F–7E·7G"Â7G%ÒÒ·ÐÐ¢öÖÖÃ¢Æ—7E·7G%ÒÒµÐÐ¢v—F‚¦—f–ÆRå¦—f–ÆR‡F‚’26¶vS Ð¢f÷"æÖR–â6÷'FVB‡6¶vRææÖVÆ—7B‚’“ Ð¢FFÒ6¶vRç&VB†æÖRÐ¢–bæÖRç7F'G7v—F‚‚'v÷&BöVÖ&VFF–æw2ò"“ Ð¢VÖ&VFF–æw5¶æÖUÒÒ÷6†#Seö'—FW2†FFÐ¢VÆ–bæÖRç7F'G7v—F‚‚'v÷&BöÖVF–ò"“ Ð¢ÖVF–¶æÖUÒÒ÷6†#Seö'—FW2†FFÐ¢–bæÖRç7F'G7v—F‚‚'v÷&Bò"’æBæÖRæVæG7v—F‚‚"ç†ÖÂ"“ Ð¢&ö÷BÒWG&VRæg&ö×7G&–ær†FFÐ¢f÷"WVF–öâ–â&ö÷Bç‡F‚‚"âòöÓ¦ôÖF‚"ÂæÖW76W3Ôå2“ Ð¢6æöæ–6ÂÒWG&VRçF÷7G&–ær†WVF–öâÂÖWF†öCÒ&3Fâ"ÂW†6ÇW6—fSÕG'VRÐ¢öÖÖÂæVæB…÷6†#Seö'—FW2†6æöæ–6Â’Ð¢&WGW&â°Ð¢&VÖ&VFF–æw2#¢VÖ&VFF–æw2ÀÐ¢&ÖVF–#¢ÖVF–ÀÐ¢&öÖÖÅ÷6†#Sb#¢6÷'FVB†öÖÖÂ’ÀÐ¢ÐÐ Ð Ð¦FVb&÷FV7FVE÷–ÆöEöÖæ–fW7B‡Fƒ¢F‚’ÓâF–7E·7G"ÂÆ—7E·7G%ÕÓ Ð¢""$6ö×&R&÷FV7FVB–ÆöG2–æFWVæFVçFÇ’öb6¶vR'B&VæÖ–ærâ"" Ð¢Öæ–fW7BÒ&÷FV7FVEöö&¦V7EöÖæ–fW7B‡F‚Ð¢&WGW&â°Ð¢&VÖ&VFF–æw2#¢6÷'FVB†Öæ–fW7E²&VÖ&VFF–æw2%ÒçfÇVW2‚’’ÀÐ¢&ÖVF–#¢6÷'FVB†Öæ–fW7E²&ÖVF–%ÒçfÇVW2‚’’ÀÐ¢&öÖÖÅ÷6†#Sb#¢6÷'FVB†Öæ–fW7E²&öÖÖÅ÷6†#Sb%Ò’ÀÐ¢ÐÐ Ð¦FVbÇ•÷'VÆR€Ð¢Fö7VÖVçC¢ç’ÀÐ¢'VÆS¢F–7E·7G"Âç•ÒÀÐ¢¢ÀÐ¢&w&…÷F&vWG3¢Æ—7E´ç•ÒÂæöæRÒæöæRÀÐ¢F&ÆU÷F&vWG3¢Æ—7E·GWÆU´ç’ÂF–7E·7G"Âç•ÕÕÒÂæöæRÒæöæRÀÐ¢6†FW%÷7F'C¢–çBÂæöæRÒæöæRÀÐ¢’Óâ–çC Ð¢Vç7W÷'FVBÒVç7W÷'FVE÷&÷W'F–W2‡'VÆRÐ¢–bVç7W÷'FVC Ð¢&—6Rf÷&ÖDÖöæöw&„W'&÷"€Ð¢b%'VÆR·'VÆU²v–Bu×Ò†2Vç7W÷'FVBWFöÖF–2&÷W'F–W3¢²rÂræ¦ö–â‡Vç7W÷'FVB—Ò Ð¢Ð Ð¢6VÆV7F÷"Ò'VÆU²'6VÆV7F÷"%ÐÐ¢–b6VÆV7F÷%²&¶–æB%Ò–â²&Fö7VÖVçB"Â'6V7F–öå÷&öÆR'Ó Ð¢&WGW&âÇ•÷6V7F–öå÷&÷W'F–W2†Fö7VÖVçBÂ'VÆU²'&÷W'F–W2%ÒÐ¢–b6VÆV7F÷%²&¶–æB%ÒÓÒ'F&ÆU÷&öÆR# Ð¢&WGW&âÇ•÷F&ÆU÷&÷W'F–W2†Fö7VÖVçBÂ'VÆU²'&÷W'F–W2%ÒÂF&ÆU÷F&vWG2Ð¢–b6VÆV7F÷%²&¶–æB%ÒÓÒ&f–VÆE÷&öÆR# Ð¢&÷W'F–W2ÒF–7B‡'VÆU²'&÷W'F–W2%ÒÐ¢–b6†FW%÷7F'B—2æ÷BæöæS Ð¢&÷W'F–W5²&6†FW%÷7F'B%ÒÒ6†FW%÷7F'@Ð¢&WGW&âÇ•öf–VÆE÷&÷W'F–W2†Fö7VÖVçBÂ&÷W'F–W2Ð¢–b6VÆV7F÷%²&¶–æB%ÒÓÒ&WVF–öå÷&öÆR# Ð¢&WGW&âÇ•öWVF–öå÷&÷W'F–W2†Fö7VÖVçBÂ'VÆU²'&÷W'F–W2%ÒÐ Ð¢7G–ÆUöæÖRÒ7G–ÆUöæÖUöf÷%÷6VÆV7F÷"‡6VÆV7F÷"Ð¢–b7G–ÆUöæÖR—2æöæS Ð¢&—6Rf÷&ÖDÖöæöw&„W'&÷"€Ð¢b%'VÆR·'VÆU²v–Bu×ÒW6W2âVç7W÷'FVBWFöÖF–26VÆV7F÷#¢·6VÆV7F÷'Ò Ð¢Ð¢–b&w&…÷F&vWG2—2æ÷BæöæS Ð¢&WGW&âÇ•÷7G–ÆU÷'VÆU÷Fõ÷&w&‡2†Fö7VÖVçBÂ'VÆRÂ&w&…÷F&vWG2Ð¢7G–ÆRÒVç7W&U÷&w&…÷7G–ÆR†Fö7VÖVçBÂ7G–ÆUöæÖRÐ¢Ç•÷7G–ÆU÷&÷W'F–W2‡7G–ÆRÂ'VÆU²'&÷W'F–W2%ÒÐ¢&WGW&â7VÒ€Ð¢Ð¢f÷"&w&‚–â—FW%öFö7VÖVçE÷&w&‡2†Fö7VÖVçBÐ¢–b&w&‚ç7G–ÆRæB&w&‚ç7G–ÆRææÖRÓÒ7G–ÆUöæÖPÐ¢Ð Ð Ð¦FVbf—'7Eöæ6†÷%÷&w&‚†Fö7VÖVçC¢ç’Â'VÆS¢F–7E·7G"Âç•Ò’Óâç’ÂæöæS Ð¢7G–ÆUöæÖRÒ7G–ÆUöæÖUöf÷%÷6VÆV7F÷"‡'VÆU²'6VÆV7F÷"%ÒÐ¢&w&‡2ÒÆ—7B†Fö7VÖVçBç&w&‡2Ð¢–b7G–ÆUöæÖS Ð¢f÷"&w&‚–â&w&‡3 Ð¢–b€Ð¢&w&‚çFW‡Bç7G&—‚Ð¢æB&w&‚ç7G–ÆPÐ¢æB&w&‚ç7G–ÆRææÖRÓÒ7G–ÆUöæÖPÐ¢æB&w&‚ç'Vç0Ð¢“ Ð¢&WGW&â&w&€Ð¢f÷"&w&‚–â&w&‡3 Ð¢–b&w&‚çFW‡Bç7G&—‚’æB&w&‚ç'Vç3 Ð¢&WGW&â&w&€Ð¢&WGW&âæöæPÐ Ð Ð¦FVb7VÖÖ&—¦U÷'VÆR‡'VÆS¢F–7E·7G"Âç•Ò’Óâ7G# Ð¢—'2Ò"Â"æ¦ö–â†b'¶¶W—Ó×·fÇVWÒ"f÷"¶W’ÂfÇVR–â6÷'FVB‡'VÆU²'&÷W'F–W2%Òæ—FV×2‚’’Ð¢6÷W&6W2Ò"Â"æ¦ö–â‡'VÆU²'6÷W&6Uö–G2%ÒÐ¢&WGW&âb'·'VÆU²v–Bu×Ó¢·—'7Òâ6÷W&6W3¢·6÷W&6W7Òâ Ð Ð Ð¦FVbÆöEöFö7VÖVçB‡Fƒ¢F‚’Óâç“ Ð¢Vç7W&UöFö7‚‡F‚Ð¢G'“ Ð¢&WGW&âFö7VÖVçB‡7G"‡F‚’Ð¢W†6WBW†6WF–öâ2W†3 Ð¢&—6Rf÷&ÖDÖöæöw&„W'&÷"†b%Væ&ÆRFò÷VâDô5‚·F‡Ó¢¶W†7Ò"’g&öÒW†0Ð