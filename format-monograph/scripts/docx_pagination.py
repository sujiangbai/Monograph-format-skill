"""Apply and inspect approved DOCX pagination sections."""

from __future__ import annotations

import copy
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Callable

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from _common import (
    NS,
    STRUCTURAL_INDENT_ATTRIBUTES,
    FormatMonographError,
    style_effective_font,
)


FOOTER_PART = re.compile(r"word/footer\d+\.xml$")
HEADER_FOOTER_PART = re.compile(r"word/(?:header|footer)\d+\.xml$")
def _section_properties(document: Any) -> list[Any]:
    return document.element.body.xpath("./w:p/w:pPr/w:sectPr | ./w:sectPr")


def _body_position(document: Any, paragraph: Any) -> int:
    children = list(document.element.body)
    try:
        return children.index(paragraph._p)
    except ValueError as exc:
        raise FormatMonographError(
            "Pagination locators must resolve to top-level body paragraphs."
        ) from exc


def section_index_for_paragraph(document: Any, paragraph: Any) -> int:
    position = _body_position(document, paragraph)
    section_index = 0
    for child in list(document.element.body)[:position]:
        p_pr = child.find(qn("w:pPr")) if child.tag == qn("w:p") else None
        if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
            section_index += 1
    return section_index


def _section_for_position(document: Any, position: int) -> Any:
    children = list(document.element.body)
    for child in children[position:]:
        if child.tag == qn("w:p"):
            p_pr = child.find(qn("w:pPr"))
            sect_pr = None if p_pr is None else p_pr.find(qn("w:sectPr"))
            if sect_pr is not None:
                return sect_pr
        elif child.tag == qn("w:sectPr"):
            return child
    raise FormatMonographError("The DOCX body has no section properties.")


def _boundary_before(document: Any, paragraph: Any) -> tuple[Any, bool]:
    body = document.element.body
    children = list(body)
    position = _body_position(document, paragraph)
    if position:
        previous = children[position - 1]
        p_pr = previous.find(qn("w:pPr")) if previous.tag == qn("w:p") else None
        existing = None if p_pr is None else p_pr.find(qn("w:sectPr"))
        if existing is not None:
            return existing, False

    source = _section_for_position(document, position)
    boundary = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    sect_pr = copy.deepcopy(source)
    section_type = sect_pr.find(qn("w:type"))
    if section_type is None:
        section_type = OxmlElement("w:type")
        sect_pr.insert(0, section_type)
    section_type.set(qn("w:val"), "nextPage")
    p_pr.append(sect_pr)
    boundary.append(p_pr)
    body.insert(position, boundary)
    return sect_pr, True


def _paragraph_has_layout_payload(element: Any) -> bool:
    if element.tag != qn("w:p"):
        return True
    if "".join(element.xpath(".//w:t/text()")).strip():
        return True
    return bool(
        element.xpath(
            ".//w:drawing | .//w:object | .//w:pict | .//w:br | .//w:fldChar | "
            ".//w:instrText | .//w:footnoteReference | .//w:endnoteReference"
        )
    )


def _is_generated_empty_section_boundary(element: Any) -> bool:
    if element.tag != qn("w:p") or list(element) == []:
        return False
    children = list(element)
    if len(children) != 1 or children[0].tag != qn("w:pPr"):
        return False
    properties = list(children[0])
    return len(properties) == 1 and properties[0].tag == qn("w:sectPr")


def _nonzero_indent_attributes(ind: Any | None) -> dict[str, str]:
    if ind is None:
        return {}
    result = {}
    for attribute in STRUCTURAL_INDENT_ATTRIBUTES:
        value = ind.get(qn(f"w:{attribute}"))
        if value is not None and value != "0":
            result[attribute] = value
    return result


def _boundary_after_title(
    document: Any, title_paragraph: Any, toc_paragraph: Any
) -> tuple[Any, bool]:
    body = document.element.body
    children = list(body)
    title_position = _body_position(document, title_paragraph)
    toc_position = _body_position(document, toc_paragraph)
    if title_position >= toc_position:
        raise FormatMonographError("The book title must precede the TOC heading.")

    between = children[title_position + 1 : toc_position]
    if any(_paragraph_has_layout_payload(element) for element in between):
        raise FormatMonographError(
            "Non-empty authored content between the book title and TOC requires QA."
        )

    title_p_pr = title_paragraph._p.get_or_add_pPr()
    existing = title_p_pr.find(qn("w:sectPr"))
    generated_boundaries = [
        element for element in between if _is_generated_empty_section_boundary(element)
    ]
    source = existing
    if source is None and generated_boundaries:
        generated_p_pr = generated_boundaries[-1].find(qn("w:pPr"))
        source = generated_p_pr.find(qn("w:sectPr"))
    if source is None:
        source = _section_for_position(document, title_position)

    inserted = existing is None
    if existing is None:
        existing = copy.deepcopy(source)
        title_p_pr.append(existing)
    for boundary in generated_boundaries:
        body.remove(boundary)
    return existing, inserted


def _set_page_numbering(sect_pr: Any, start: int | None, number_format: str) -> None:
    pg_num = sect_pr.find(qn("w:pgNumType"))
    if pg_num is None and start is None:
        return
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        sect_pr.append(pg_num)
    pg_num.set(qn("w:fmt"), number_format)
    if start is None:
        pg_num.attrib.pop(qn("w:start"), None)
    else:
        pg_num.set(qn("w:start"), str(start))


def _suppress_redundant_body_page_break(paragraph: Any, sect_pr: Any) -> bool:
    section_type = sect_pr.find(qn("w:type"))
    section_value = (
        section_type.get(qn("w:val")) if section_type is not None else "nextPage"
    )
    if section_value == "continuous":
        return False
    direct = paragraph.paragraph_format.page_break_before
    if direct is False:
        return False
    inherited = (
        paragraph.style.paragraph_format.page_break_before
        if direct is None and paragraph.style is not None
        else None
    )
    if direct is None and inherited is not True:
        return False
    paragraph.paragraph_format.page_break_before = False
    return True


def _hide_single_title_page_running_elements(sect_pr: Any) -> None:
    title_page = sect_pr.find(qn("w:titlePg"))
    if title_page is None:
        title_page = OxmlElement("w:titlePg")
        sect_pr.append(title_page)
    title_page.set(qn("w:val"), "true")
    page_number = sect_pr.find(qn("w:pgNumType"))
    if page_number is not None:
        sect_pr.remove(page_number)


def _set_section_type(sect_pr: Any, value: str) -> None:
    section_type = sect_pr.find(qn("w:type"))
    if section_type is None:
        section_type = OxmlElement("w:type")
        sect_pr.insert(0, section_type)
    section_type.set(qn("w:val"), value)


def _set_vertical_alignment(sect_pr: Any, value: str) -> None:
    if value not in {"top", "center", "both", "bottom"}:
        raise FormatMonographError(f"Unsupported section vertical alignment: {value}")
    alignment = sect_pr.find(qn("w:vAlign"))
    if alignment is None:
        alignment = OxmlElement("w:vAlign")
        sect_pr.append(alignment)
    alignment.set(qn("w:val"), value)


def _field_instructions(root: Any) -> list[str]:
    def xpath(expression: str) -> list[Any]:
        try:
            return root.xpath(expression, namespaces=NS)
        except TypeError:
            return root.xpath(expression)

    result = [
        value.strip()
        for value in xpath(".//w:fldSimple/@w:instr")
        if value.strip()
    ]
    stack: list[dict[str, Any]] = []
    for element in root.iter():
        if element.tag == qn("w:fldChar"):
            kind = element.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append({"collecting": True, "parts": []})
            elif kind == "separate" and stack:
                field = stack[-1]
                instruction = "".join(field["parts"]).strip()
                if instruction:
                    result.append(instruction)
                field["collecting"] = False
            elif kind == "end" and stack:
                field = stack.pop()
                if field["collecting"]:
                    instruction = "".join(field["parts"]).strip()
                    if instruction:
                        result.append(instruction)
        elif (
            element.tag == qn("w:instrText")
            and stack
            and stack[-1]["collecting"]
        ):
            stack[-1]["parts"].append(element.text or "")
    return result


def _field_types(root: Any) -> set[str]:
    return {
        instruction.split(maxsplit=1)[0].upper()
        for instruction in _field_instructions(root)
        if instruction.strip()
    }


def _page_field_count(root: Any) -> int:
    def xpath(expression: str) -> list[Any]:
        try:
            return root.xpath(expression, namespaces=NS)
        except TypeError:
            return root.xpath(expression)

    simple = xpath(
        ".//w:fldSimple[starts-with(translate(normalize-space(@w:instr), 'page', 'PAGE'), 'PAGE')]"
    )
    complex_instructions = [
        value
        for value in xpath(".//w:instrText/text()")
        if value.strip().upper().startswith("PAGE")
    ]
    return len(simple) + len(complex_instructions)


def _page_only_footer(footer: Any) -> bool:
    root = footer._element
    if root.xpath(
        ".//w:tbl | .//w:object | "
        ".//*[local-name()='blip']"
    ):
        return False
    instructions = _field_instructions(root)
    field_types = _field_types(root)
    page_count = _page_field_count(root)
    if page_count == 0:
        return not field_types and not root.xpath(
            ".//w:drawing | .//w:pict | .//w:t[normalize-space(.) != '']"
        )
    if field_types - {"PAGE", "="}:
        return False
    formulas = [
        re.sub(r"\s+", "", instruction).upper()
        for instruction in instructions
        if instruction.lstrip().startswith("=")
    ]
    if formulas and (formulas != ["=-1"] or page_count != 1):
        return False
    for visual in root.xpath(".//w:drawing | .//w:pict"):
        if (
            not visual.xpath(".//*[local-name()='txbxContent']")
            or not _page_only_payload(visual)
        ):
            return False
        if visual.xpath(".//*[local-name()='blip']"):
            return False
        for image_data in visual.xpath(".//*[local-name()='imagedata']"):
            if any(str(value).strip() for value in image_data.attrib.values()):
                return False
    visible = [
        value.strip()
        for value in root.xpath(".//w:t/text()")
        if value.strip()
    ]
    return all(
        value.isdigit() or re.fullmatch(r"[IVXLCDMivxlcdm]+", value)
        for value in visible
    )


def _page_only_payload(root: Any) -> bool:
    field_types = _field_types(root)
    page_count = _page_field_count(root)
    if field_types - {"PAGE", "="} or page_count < 1:
        return False
    formulas = [
        re.sub(r"\s+", "", instruction).upper()
        for instruction in _field_instructions(root)
        if instruction.lstrip().startswith("=")
    ]
    return not formulas or (formulas == ["=-1"] and page_count == 1)


def _replace_with_page_field(footer: Any, alignment: Any) -> None:
    paragraphs = list(footer.paragraphs)
    paragraph = paragraphs[0] if paragraphs else footer.add_paragraph()
    for extra in paragraphs[1:]:
        footer._element.remove(extra._p)
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)
    paragraph.alignment = alignment
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def _ensure_page_field(
    footer: Any, alignment: Any, *, replace_static_page_text: bool = False
) -> bool:
    page_fields = _page_field_count(footer._element)
    page_only = _page_only_footer(footer)
    if page_fields > 1:
        if not page_only:
            raise FormatMonographError(
                "Footer contains multiple PAGE fields mixed with non-page content."
            )
        _replace_with_page_field(footer, alignment)
        return True
    if page_fields == 1:
        if page_only and len(footer.paragraphs) != 1:
            _replace_with_page_field(footer, alignment)
            return True
        field = footer._element.xpath(
            ".//w:fldSimple[starts-with(translate(normalize-space(@w:instr), 'page', 'PAGE'), 'PAGE')]"
        )
        paragraph = field[0].getparent() if field else footer.paragraphs[0]._p
        while paragraph is not None and paragraph.tag != qn("w:p"):
            paragraph = paragraph.getparent()
        if paragraph is not None:
            from docx.text.paragraph import Paragraph

            Paragraph(paragraph, footer._element).alignment = alignment
        return False
    visible = "".join(footer._element.xpath(".//w:t/text()")).strip()
    has_non_text_payload = bool(
        footer._element.xpath(".//w:tbl | .//w:drawing | .//w:object | .//w:pict")
    )
    if visible or has_non_text_payload:
        if replace_static_page_text and visible.isdigit() and not has_non_text_payload:
            _replace_with_page_field(footer, alignment)
            return True
        raise FormatMonographError(
            "Footer contains non-page content; page fields require explicit QA."
        )
    paragraph = next((item for item in footer.paragraphs if not item.text.strip()), None)
    if paragraph is None:
        paragraph = footer.add_paragraph()
    paragraph.alignment = alignment
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)
    return True


def _ensure_visible_footers(
    document: Any,
    start_index: int,
    *,
    restart_indexes: set[int] | None = None,
    replace_static_page_text: bool = False,
) -> int:
    changed = 0
    restart_indexes = restart_indexes or set()
    document.settings.odd_and_even_pages_header_footer = True
    sections = list(document.sections)
    for index, section in enumerate(sections[start_index:], start=start_index):
        section.different_first_page_header_footer = False
        if index in restart_indexes and index > 0:
            for footer in (section.footer, section.even_page_footer):
                if not footer.is_linked_to_previous:
                    continue
                if not _page_only_footer(footer):
                    raise FormatMonographError(
                        "A restarted section inherits non-page footer content."
                    )
                footer.is_linked_to_previous = False
                changed += 1
        changed += int(
            _ensure_page_field(
                section.footer,
                WD_ALIGN_PARAGRAPH.RIGHT,
                replace_static_page_text=replace_static_page_text,
            )
        )
        changed += int(
            _ensure_page_field(
                section.even_page_footer,
                WD_ALIGN_PARAGRAPH.LEFT,
                replace_static_page_text=replace_static_page_text,
            )
        )
        if index > start_index:
            # A missing relationship inherits the prior footer. Existing independent
            # footer content is retained and receives its own PAGE field above.
            pass
    return changed


def apply_pagination_sections(
    document: Any,
    settings: dict[str, Any],
    resolver: Callable[[Any, dict[str, Any]], Any],
    *,
    replace_static_page_text: bool = False,
    front_matter: dict[str, Any] | None = None,
) -> dict[str, Any]:
    if not settings.get("approved"):
        return {}
    toc_locator = settings.get("toc_start")
    body_locator = settings.get("body_start")
    if not isinstance(toc_locator, dict) or not isinstance(body_locator, dict):
        raise FormatMonographError(
            "Approved pagination_sections requires toc_start and body_start locators."
        )
    toc_paragraph = resolver(document, toc_locator)
    body_paragraph = resolver(document, body_locator)
    if _body_position(document, toc_paragraph) >= _body_position(document, body_paragraph):
        raise FormatMonographError("TOC pagination start must precede body pagination start.")

    front_matter = front_matter or {}
    title_boundary_inserted = False
    title_index = None
    if front_matter.get("approved") and front_matter.get("separate_title_page"):
        toc_section_start = getattr(
            document, "_format_monograph_toc_heading", toc_paragraph
        )
        title_paragraph = getattr(document, "_format_monograph_book_title", None)
        if title_paragraph is None:
            title_paragraph = resolver(document, front_matter["book_title"])
        title_boundary, title_boundary_inserted = _boundary_after_title(
            document, title_paragraph, toc_section_start
        )
        _set_section_type(title_boundary, "nextPage")
        _set_vertical_alignment(
            title_boundary,
            str(front_matter.get("title_page_vertical_alignment", "top")),
        )
        toc_section_start.paragraph_format.page_break_before = False
        _hide_single_title_page_running_elements(title_boundary)
        title_index = section_index_for_paragraph(document, title_paragraph)

    body_boundary, inserted = _boundary_before(document, body_paragraph)
    if front_matter.get("approved"):
        _set_section_type(body_boundary, "nextPage")
        body_paragraph.paragraph_format.page_break_before = False
    page_break_suppressed = _suppress_redundant_body_page_break(
        body_paragraph, body_boundary
    )
    toc_index = section_index_for_paragraph(document, toc_paragraph)
    body_index = section_index_for_paragraph(document, body_paragraph)
    if body_index != toc_index + 1:
        raise FormatMonographError(
            "The approved TOC and body starts must form adjacent pagination sections."
        )

    sections = _section_properties(document)
    starts = settings.get("start_at", {})
    number_format = str(settings.get("number_format", "decimal"))
    _set_page_numbering(sections[toc_index], int(starts.get("toc", 1)), number_format)
    _set_page_numbering(sections[body_index], int(starts.get("body", 1)), number_format)
    if settings.get("continue_after_body_start", True):
        for sect_pr in sections[body_index + 1 :]:
            _set_page_numbering(sect_pr, None, number_format)

    fields_added = _ensure_visible_footers(
        document,
        toc_index,
        restart_indexes={toc_index, body_index},
        replace_static_page_text=replace_static_page_text,
    )
    return {
        "title_section": title_index,
        "toc_section": toc_index,
        "body_section": body_index,
        "inserted_title_section_break": title_boundary_inserted,
        "inserted_body_section_break": inserted,
        "suppressed_redundant_body_page_break": page_break_suppressed,
        "page_fields_added": fields_added,
    }


def finalize_pagination_sections(
    document: Any,
    settings: dict[str, Any],
    resolver: Callable[[Any, dict[str, Any]], Any],
    *,
    front_matter: dict[str, Any] | None = None,
) -> bool:
    if not settings.get("approved"):
        return False
    body_locator = settings.get("body_start")
    if not isinstance(body_locator, dict):
        raise FormatMonographError(
            "Approved pagination_sections requires a body_start locator."
        )
    body_paragraph = resolver(document, body_locator)
    body_index = section_index_for_paragraph(document, body_paragraph)
    sections = _section_properties(document)
    if not 0 <= body_index < len(sections):
        raise FormatMonographError("Body pagination section is out of range.")
    front_matter = front_matter or {}
    if front_matter.get("approved"):
        title_paragraph = getattr(document, "_format_monograph_book_title", None)
        if title_paragraph is None:
            title_paragraph = resolver(document, front_matter["book_title"])
        title_index = section_index_for_paragraph(document, title_paragraph)
        _hide_single_title_page_running_elements(sections[title_index])
        _set_section_type(sections[title_index], "nextPage")
        toc_heading = getattr(document, "_format_monograph_toc_heading", None)
        if toc_heading is not None:
            toc_heading.paragraph_format.page_break_before = False
        _set_section_type(sections[body_index], "nextPage")
        body_paragraph.paragraph_format.page_break_before = False
        return True
    return _suppress_redundant_body_page_break(body_paragraph, sections[body_index])


def _relationship_targets(package: zipfile.ZipFile) -> dict[str, str]:
    name = "word/_rels/document.xml.rels"
    if name not in package.namelist():
        return {}
    root = etree.fromstring(package.read(name))
    result = {}
    for relationship in root:
        identifier = relationship.get("Id")
        target = relationship.get("Target", "")
        if identifier:
            result[identifier] = posixpath.normpath(posixpath.join("word", target))
    return result


def pagination_inventory(path: Path) -> dict[str, Any]:
    with zipfile.ZipFile(path) as package:
        document = etree.fromstring(package.read("word/document.xml"))
        settings = (
            etree.fromstring(package.read("word/settings.xml"))
            if "word/settings.xml" in package.namelist()
            else None
        )
        relationships = _relationship_targets(package)
        referenced_parts: set[str] = set()
        inherited: dict[str, str] = {}
        sections = []
        for index, sect_pr in enumerate(
            document.xpath("./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr", namespaces=NS)
        ):
            explicit = {}
            for reference in sect_pr.xpath(
                "./w:footerReference", namespaces=NS
            ):
                kind = reference.get(qn("w:type"), "default")
                rel_id = reference.get(qn("r:id"))
                if rel_id:
                    explicit[kind] = rel_id
                    inherited[kind] = rel_id
                    target = relationships.get(rel_id)
                    if target:
                        referenced_parts.add(target)
            effective = dict(inherited)
            footer_fields = {}
            footer_page_field_counts = {}
            footer_non_page_payload = {}
            for kind, rel_id in effective.items():
                target = relationships.get(rel_id)
                if target and target in package.namelist():
                    footer_root = etree.fromstring(package.read(target))
                    field_types = _field_types(footer_root)
                    footer_fields[kind] = sorted(field_types)
                    footer_page_field_counts[kind] = _page_field_count(footer_root)
                    visible = [
                        value.strip()
                        for value in footer_root.xpath(".//w:t/text()", namespaces=NS)
                        if value.strip()
                    ]
                    footer_non_page_payload[kind] = bool(
                        not _page_only_payload(footer_root)
                        or footer_root.xpath(
                            ".//w:tbl | .//w:drawing | .//w:object | .//w:pict",
                            namespaces=NS,
                        )
                        or any(not value.isdigit() for value in visible)
                    )
                else:
                    footer_fields[kind] = []
                    footer_page_field_counts[kind] = 0
                    footer_non_page_payload[kind] = False
            pg_num = sect_pr.find(qn("w:pgNumType"))
            vertical_alignment = sect_pr.find(qn("w:vAlign"))
            sections.append(
                {
                    "index": index,
                    "page_number_start": (
                        None if pg_num is None else pg_num.get(qn("w:start"))
                    ),
                    "page_number_format": (
                        None if pg_num is None else pg_num.get(qn("w:fmt"))
                    ),
                    "different_first_page": sect_pr.find(qn("w:titlePg")) is not None,
                    "vertical_alignment": (
                        None
                        if vertical_alignment is None
                        else vertical_alignment.get(qn("w:val"), "top")
                    ),
                    "footer_references": explicit,
                    "effective_footer_references": effective,
                    "footer_fields": footer_fields,
                    "footer_page_field_counts": footer_page_field_counts,
                    "footer_non_page_payload": footer_non_page_payload,
                    "missing_page_footer_types": [
                        kind
                        for kind in ("default", "even")
                        if "PAGE" not in footer_fields.get(kind, [])
                    ],
                }
            )
        all_parts = {
            name for name in package.namelist() if HEADER_FOOTER_PART.fullmatch(name)
        }
        all_referenced = {
            relationships.get(rel_id)
            for sect_pr in document.xpath(".//w:sectPr", namespaces=NS)
            for rel_id in sect_pr.xpath(
                "./w:headerReference/@r:id | ./w:footerReference/@r:id", namespaces=NS
            )
        }
        return {
            "odd_and_even_pages_header_footer": bool(
                settings is not None
                and settings.xpath("./w:evenAndOddHeaders", namespaces=NS)
            ),
            "sections": sections,
            "page_number_restarts": [
                item["index"] for item in sections if item["page_number_start"] is not None
            ],
            "orphan_header_footer_parts": sorted(all_parts - set(all_referenced)),
        }


def _resolve_audit_boundary(
    document: Any,
    locator: dict[str, Any],
    resolver: Callable[[Any, dict[str, Any]], Any],
    *,
    allow_unique_toc_field: bool = False,
    alternate_locators: list[dict[str, Any]] | None = None,
) -> Any:
    try:
        return resolver(document, locator)
    except FormatMonographError as original_error:
        for alternate in alternate_locators or []:
            try:
                return resolver(document, alternate)
            except FormatMonographError:
                continue
        if not allow_unique_toc_field:
            raise original_error
        matches = [
            paragraph
            for paragraph in document.paragraphs
            if paragraph._p.xpath(
                ".//w:fldSimple[starts-with(translate(normalize-space(@w:instr), "
                "'toc', 'TOC'), 'TOC')] | "
                ".//w:instrText[starts-with(translate(normalize-space(.), "
                "'toc', 'TOC'), 'TOC')]"
            )
        ]
        if len(matches) == 1:
            return matches[0]
        raise original_error


def audit_pagination_sections(
    path: Path,
    document: Any,
    settings: dict[str, Any],
    resolver: Callable[[Any, dict[str, Any]], Any],
    structure_map: dict[str, Any] | None = None,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    inventory = pagination_inventory(path)
    if not settings.get("approved"):
        return [], inventory
    toc = section_index_for_paragraph(
        document,
        _resolve_audit_boundary(
            document,
            settings["toc_start"],
            resolver,
            allow_unique_toc_field=True,
        ),
    )
    body_locator = settings["body_start"]
    alternate_body_locators = []
    for heading in (structure_map or {}).get("headings", []):
        heading_locator = heading.get("locator", {})
        if (
            heading.get("approved")
            and heading_locator.get("text_sha256")
            == body_locator.get("text_sha256")
            and heading.get("normalized_text_sha256")
        ):
            alternate = copy.deepcopy(body_locator)
            alternate["text_sha256"] = heading["normalized_text_sha256"]
            alternate_body_locators.append(alternate)
    body_paragraph = _resolve_audit_boundary(
        document,
        body_locator,
        resolver,
        alternate_locators=alternate_body_locators,
    )
    body = section_index_for_paragraph(document, body_paragraph)
    failures = []
    front_matter = (structure_map or {}).get("front_matter", {})
    if front_matter.get("approved"):
        title_paragraph = _resolve_audit_boundary(
            document,
            front_matter["book_title"],
            resolver,
        )
        title_section = section_index_for_paragraph(document, title_paragraph)
        toc_headings = [
            paragraph
            for paragraph in document.paragraphs
            if paragraph.text.strip() == front_matter.get("toc_heading_text")
            and paragraph.style is not None
            and paragraph.style.name == "Monograph TOC Heading"
        ]
        if len(toc_headings) != 1:
            failures.append(
                {
                    "property": "toc_heading",
                    "expected": front_matter.get("toc_heading_text"),
                    "actual_count": len(toc_headings),
                }
            )
        elif section_index_for_paragraph(document, toc_headings[0]) != toc:
            failures.append(
                {"property": "toc_heading_section", "expected": toc}
            )
        if title_paragraph.style is None or title_paragraph.style.name != "Monograph Book Title":
            failures.append(
                {"property": "book_title_style", "expected": "Monograph Book Title"}
            )
        title_alignment = title_paragraph.alignment
        if title_alignment is None and title_paragraph.style is not None:
            title_alignment = title_paragraph.style.paragraph_format.alignment
        if title_alignment != WD_ALIGN_PARAGRAPH.CENTER:
            failures.append(
                {"property": "book_title_alignment", "expected": "center"}
            )
        title_p_pr = title_paragraph._p.pPr
        title_direct_num_pr = (
            None if title_p_pr is None else title_p_pr.find(qn("w:numPr"))
        )
        if title_direct_num_pr is not None:
            failures.append(
                {"property": "book_title_direct_numbering", "expected": "absent"}
            )
        direct_indent = (
            None if title_p_pr is None else title_p_pr.find(qn("w:ind"))
        )
        direct_indent_conflicts = _nonzero_indent_attributes(direct_indent)
        if direct_indent_conflicts:
            failures.append(
                {
                    "property": "book_title_direct_indent",
                    "expected": 0,
                    "actual": direct_indent_conflicts,
                }
            )
        expected_title_format = {
            "font_name_east_asia": "黑体",
            "font_name_ascii": "Times New Roman",
            "font_size_pt": 22,
            "bold": True,
        }
        expected_title_format.update(front_matter.get("book_title_format", {}))
        title_style = title_paragraph.style
        if title_style is not None:
            style_p_pr = title_style.element.pPr
            style_indent = (
                None if style_p_pr is None else style_p_pr.find(qn("w:ind"))
            )
            style_indent_conflicts = _nonzero_indent_attributes(style_indent)
            required_zero_attributes = {
                attribute: (
                    None
                    if style_indent is None
                    else style_indent.get(qn(f"w:{attribute}"))
                )
                for attribute in (
                    "left",
                    "leftChars",
                    "right",
                    "rightChars",
                    "firstLine",
                    "firstLineChars",
                )
            }
            if style_indent_conflicts or any(
                value != "0" for value in required_zero_attributes.values()
            ):
                failures.append(
                    {
                        "property": "book_title_style_indent",
                        "expected": 0,
                        "actual": {
                            **required_zero_attributes,
                            **style_indent_conflicts,
                        },
                    }
                )
            actual_ascii = style_effective_font(document, title_style, "ascii")[0]
            actual_east_asia = style_effective_font(document, title_style, "eastAsia")[0]
            if actual_ascii != expected_title_format["font_name_ascii"]:
                failures.append(
                    {
                        "property": "book_title_font_name_ascii",
                        "expected": expected_title_format["font_name_ascii"],
                        "actual": actual_ascii,
                    }
                )
            if actual_east_asia != expected_title_format["font_name_east_asia"]:
                failures.append(
                    {
                        "property": "book_title_font_name_east_asia",
                        "expected": expected_title_format["font_name_east_asia"],
                        "actual": actual_east_asia,
                    }
                )
            actual_size = None if title_style.font.size is None else title_style.font.size.pt
            if actual_size != float(expected_title_format["font_size_pt"]):
                failures.append(
                    {
                        "property": "book_title_font_size_pt",
                        "expected": float(expected_title_format["font_size_pt"]),
                        "actual": actual_size,
                    }
                )
        style_bold = (
            title_paragraph.style is not None
            and title_paragraph.style.font.bold is True
        )
        if any(
            run.text and run.bold is not True and not (run.bold is None and style_bold)
            for run in title_paragraph.runs
        ):
            failures.append({"property": "book_title_bold", "expected": True})
        title_inventory = inventory["sections"][title_section]
        title_section_properties = (
            None if title_p_pr is None else title_p_pr.find(qn("w:sectPr"))
        )
        if title_section_properties is None:
            failures.append(
                {
                    "property": "book_title_section_boundary",
                    "expected": "attached_to_book_title_paragraph",
                }
            )
        expected_vertical = front_matter.get("title_page_vertical_alignment")
        if expected_vertical and title_inventory.get("vertical_alignment") != expected_vertical:
            failures.append(
                {
                    "section": title_section,
                    "property": "title_page_vertical_alignment",
                    "expected": expected_vertical,
                    "actual": title_inventory.get("vertical_alignment"),
                }
            )
        expected_spacing = expected_title_format.get("line_spacing_pt")
        if expected_spacing is not None and title_style is not None:
            actual_spacing = title_style.paragraph_format.line_spacing
            actual_spacing_pt = getattr(actual_spacing, "pt", None)
            if actual_spacing_pt != float(expected_spacing):
                failures.append(
                    {
                        "property": "book_title_line_spacing_pt",
                        "expected": float(expected_spacing),
                        "actual": actual_spacing_pt,
                    }
                )
            expected_rule = expected_title_format.get("line_spacing_rule")
            rule_names = {
                WD_LINE_SPACING.AT_LEAST: "at_least",
                WD_LINE_SPACING.EXACTLY: "exact",
                WD_LINE_SPACING.SINGLE: "single",
                WD_LINE_SPACING.ONE_POINT_FIVE: "one_point_five",
                WD_LINE_SPACING.DOUBLE: "double",
            }
            actual_rule = rule_names.get(title_style.paragraph_format.line_spacing_rule)
            if expected_rule and actual_rule != expected_rule:
                failures.append(
                    {
                        "property": "book_title_line_spacing_rule",
                        "expected": expected_rule,
                        "actual": actual_rule,
                    }
                )
        if not title_inventory["different_first_page"]:
            failures.append(
                {"section": title_section, "property": "title_page_number_visible"}
            )
        if title_inventory["footer_page_field_counts"].get("first", 0):
            failures.append(
                {"section": title_section, "property": "title_page_first_footer_page_field"}
            )
    direct_page_break = body_paragraph.paragraph_format.page_break_before
    inherited_page_break = (
        body_paragraph.style.paragraph_format.page_break_before
        if body_paragraph.style is not None
        else None
    )
    if front_matter.get("approved"):
        body_section_type = document.sections[body]._sectPr.find(qn("w:type"))
        body_section_value = (
            "nextPage"
            if body_section_type is None
            else body_section_type.get(qn("w:val"), "nextPage")
        )
        if direct_page_break is not True and body_section_value == "continuous":
            failures.append(
                {
                    "section": body,
                    "property": "body_page_break_for_continuous_restart",
                    "expected": True,
                    "actual": direct_page_break,
                }
            )
    elif direct_page_break is True or (
        direct_page_break is None and inherited_page_break is True
    ):
        failures.append(
            {
                "section": body,
                "property": "redundant_page_break_before_at_body_start",
                "expected": False,
                "actual": True,
            }
        )
    starts = settings.get("start_at", {})
    expected = {toc: str(starts.get("toc", 1)), body: str(starts.get("body", 1))}
    sections = inventory["sections"]
    for index, start in expected.items():
        if index >= len(sections) or sections[index]["page_number_start"] != start:
            failures.append(
                {"section": index, "property": "page_number_start", "expected": start}
            )
    if settings.get("continue_after_body_start", True):
        for item in sections[body + 1 :]:
            if item["page_number_start"] is not None:
                failures.append(
                    {
                        "section": item["index"],
                        "property": "unexpected_page_number_restart",
                        "actual": item["page_number_start"],
                    }
                )
    if not inventory["odd_and_even_pages_header_footer"]:
        failures.append({"property": "odd_and_even_pages_header_footer", "expected": True})
    for item in sections[toc:]:
        if item["different_first_page"]:
            failures.append(
                {"section": item["index"], "property": "show_on_first_page", "expected": True}
            )
        if item["missing_page_footer_types"]:
            failures.append(
                {
                    "section": item["index"],
                    "property": "visible_page_footer_types",
                    "missing": item["missing_page_footer_types"],
                }
            )
        for kind in ("default", "even"):
            count = item["footer_page_field_counts"].get(kind, 0)
            if count != 1:
                failures.append(
                    {
                        "section": item["index"],
                        "property": "page_footer_field_count",
                        "footer_type": kind,
                        "expected": 1,
                        "actual": count,
                    }
                )
            if item["footer_non_page_payload"].get(kind, False):
                failures.append(
                    {
                        "section": item["index"],
                        "property": "page_footer_non_page_payload",
                        "footer_type": kind,
                        "expected": False,
                    }
                )
    if inventory["orphan_header_footer_parts"]:
        failures.append(
            {
                "property": "orphan_header_footer_parts",
                "actual_count": len(inventory["orphan_header_footer_parts"]),
            }
        )
    return failures, inventory
