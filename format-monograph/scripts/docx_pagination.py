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


def _ensure_page_field(footer: Any, alignment: Any) -> bool:
    if "PAGE" in _field_types(footer._element):
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


def _ensure_visible_footers(document: Any, start_index: int) -> int:
    changed = 0
    document.settings.odd_and_even_pages_header_footer = True
    sections = list(document.sections)
    for index, section in enumerate(sections[start_index:], start=start_index):
        section.different_first_page_header_footer = False
        changed += int(_ensure_page_field(section.footer, WD_ALIGN_PARAGRAPH.RIGHT))
        changed += int(
            _ensure_page_field(section.even_page_footer, WD_ALIGN_PARAGRAPH.LEFT)
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

    _, inserted = _boundary_before(document, body_paragraph)
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

    fields_added = _ensure_visible_footers(document, toc_index)
    return {
        "toc_section": toc_index,
        "body_section": body_index,
        "inserted_body_section_break": inserted,
        "page_fields_added": fields_added,
    }


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
            for kind, rel_id in effective.items():
                target = relationships.get(rel_id)
                if target and target in package.namelist():
                    footer_fields[kind] = sorted(
                        _field_types(etree.fromstring(package.read(target)))
                    )
                else:
                    footer_fields[kind] = []
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


def audit_pagination_sections(
    path: Path,
    document: Any,
    settings: dict[str, Any],
    resolver: Callable[[Any, dict[str, Any]], Any],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    inventory = pagination_inventory(path)
    if not settings.get("approved"):
        return [], inventory
    toc = section_index_for_paragraph(document, resolver(document, settings["toc_start"]))
    body = section_index_for_paragraph(document, resolver(document, settings["body_start"]))
    failures = []
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
    if inventory["orphan_header_footer_parts"]:
        failures.append(
            {
                "property": "orphan_header_footer_parts",
                "actual_count": len(inventory["orphan_header_footer_parts"]),
            }
        )
    return failures, inventory
