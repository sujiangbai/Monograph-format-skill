"""Apply and inspect approved DOCX pagination sections."""

from __future__ import annotations

import copy
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any, Callable

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from _common import NS, FormatMonographError


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


def _field_types(root: Any) -> set[str]:
    def xpath(expression: str) -> list[Any]:
        try:
            return root.xpath(expression, namespaces=NS)
        except TypeError:
            return root.xpath(expression)

    result = {
        value.strip().split(maxsplit=1)[0].upper()
        for value in xpath(".//w:fldSimple/@w:instr")
        if value.strip()
    }
    parts: list[str] = []
    collecting = False
    for element in root.iter():
        if element.tag == qn("w:fldChar"):
            kind = element.get(qn("w:fldCharType"))
            if kind == "begin":
                collecting = True
                parts = []
            elif kind == "separate" and collecting:
                instruction = "".join(parts).strip()
                if instruction:
                    result.add(instruction.split(maxsplit=1)[0].upper())
                collecting = False
            elif kind == "end":
                collecting = False
        elif element.tag == qn("w:instrText") and collecting:
            parts.append(element.text or "")
    return result


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
    if _field_types(root) - {"PAGE"}:
        return False
    for visual in root.xpath(".//w:drawing | .//w:pict"):
        if (
            not visual.xpath(".//*[local-name()='txbxContent']")
            or _field_types(visual) != {"PAGE"}
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
    document: Any, start_index: int, *, replace_static_page_text: bool = False
) -> int:
    changed = 0
    document.settings.odd_and_even_pages_header_footer = True
    sections = list(document.sections)
    for index, section in enumerate(sections[start_index:], start=start_index):
        section.different_first_page_header_footer = False
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

    body_boundary, inserted = _boundary_before(document, body_paragraph)
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
        replace_static_page_text=replace_static_page_text,
    )
    return {
        "toc_section": toc_index,
        "body_section": body_index,
        "inserted_body_section_break": inserted,
        "suppressed_redundant_body_page_break": page_break_suppressed,
        "page_fields_added": fields_added,
    }


def finalize_pagination_sections(
    document: Any,
    settings: dict[str, Any],
    resolver: Callable[[Any, dict[str, Any]], Any],
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
    return _suppress_redundant_body_page_break(
        body_paragraph, sections[body_index]
    )


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
                        field_types - {"PAGE"}
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
    direct_page_break = body_paragraph.paragraph_format.page_break_before
    inherited_page_break = (
        body_paragraph.style.paragraph_format.page_break_before
        if body_paragraph.style is not None
        else None
    )
    if direct_page_break is True or (
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

