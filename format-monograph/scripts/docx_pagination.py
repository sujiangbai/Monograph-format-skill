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
    document.seã^:¶‰žËkºwµçUÉ¹…Ñ•}±½…Ñ½ÉÌè±¥ÍÑm‘¥ÑmÍÑÈ°¹åutð9½¹”€ô9½¹”°4(¤€´ø¹äè4(€€€ÑÉäè4(€€€€€€€É•ÑÕÉ¸É•Í½±Ù•È¡‘½Õµ•¹Ð°±½…Ñ½È¤4(€€€•á•ÁÐ½Éµ…Ñ5½¹½É…Á¡ÉÉ½È…Ì½É¥¥¹…±}•ÉÉ½Èè4(€€€€€€€™½È…±Ñ•É¹…Ñ”¥¸…±Ñ•É¹…Ñ•}±½…Ñ½ÉÌ½Èmtè4(€€€€€€€€€€€ÑÉäè4(€€€€€€€€€€€€€€€É•ÑÕÉ¸É•Í½±Ù•È¡‘½Õµ•¹Ð°…±Ñ•É¹…Ñ”¤4(€€€€€€€€€€€•á•ÁÐ½Éµ…Ñ5½¹½É…Á¡ÉÉ½Èè4(€€€€€€€€€€€€€€€½¹Ñ¥¹Õ”4(€€€€€€€¥˜¹½Ð…±±½Ý}Õ¹¥ÅÕ•}Ñ½}™¥•±è4(€€€€€€€€€€€É…¥Í”½É¥¥¹…±}•ÉÉ½È4(€€€€€€€µ…Ñ¡•Ì€ôl4(€€€€€€€€€€€Á…É…É…Á 4(€€€€€€€€€€€™½ÈÁ…É…É…Á ¥¸‘½Õµ•¹Ð¹Á…É…É…Á¡Ì4(€€€€€€€€€€€¥˜Á…É…É…Á ¹}À¹áÁ…Ñ  4(€€€€€€€€€€€€€€€€ˆ¸¼½Üé™±‘M¥µÁ±•mÍÑ…ÉÑÌµÝ¥Ñ ¡ÑÉ…¹Í±…Ñ”¡¹½Éµ…±¥é”µÍÁ…”¡Üé¥¹ÍÑÈ¤°€ˆ4(€€€€€€€€€€€€€€€€ˆÑ½Œœ°€Q=œ¤°€Q=œ¥tð€ˆ4(€€€€€€€€€€€€€€€€ˆ¸¼½Üé¥¹ÍÑÉQ•áÑmÍÑ…ÉÑÌµÝ¥Ñ ¡ÑÉ…¹Í±…Ñ”¡¹½Éµ…±¥é”µÍÁ…” ¸¤°€ˆ4(€€€€€€€€€€€€€€€€ˆÑ½Œœ°€Q=œ¤°€Q=œ¥tˆ4(€€€€€€€€€€€€¤4(€€€€€€€t4(€€€€€€€¥˜±•¸¡µ…Ñ¡•Ì¤€ôô€Äè4(€€€€€€€€€€€É•ÑÕÉ¸µ…Ñ¡•ÍlÁt4(€€€€€€€É…¥Í”½É¥¥¹…±}•ÉÉ½È4(4(4)‘•˜…Õ‘¥Ñ}Á…¥¹…Ñ¥½¹}Í•Ñ¥½¹Ì 4(€€€Á…Ñ èA…Ñ °4(€€€‘½Õµ•¹Ðè¹ä°4(€€€Í•ÑÑ¥¹Ìè‘¥ÑmÍÑÈ°¹åt°4(€€€É•Í½±Ù•Èè…±±…‰±•mm¹ä°‘¥ÑmÍÑÈ°¹åut°¹åt°4(€€€ÍÑÉÕÑÕÉ•}µ…Àè‘¥ÑmÍÑÈ°¹åtð9½¹”€ô9½¹”°4(¤€´øÑÕÁ±•m±¥ÍÑm‘¥ÑmÍÑÈ°¹åut°‘¥ÑmÍÑÈ°¹åutè4(€€€¥¹Ù•¹Ñ½Éä€ôÁ…¥¹…Ñ¥½¹}¥¹Ù•¹Ñ½Éä¡Á…Ñ ¤4(€€€¥˜¹½ÐÍ•ÑÑ¥¹Ì¹•Ð ‰…ÁÁÉ½Ù•ˆ¤è4(€€€€€€€É•ÑÕÉ¸mt°¥¹Ù•¹Ñ½Éä4(€€€Ñ½Œ€ôÍ•Ñ¥½¹}¥¹‘•á}™½É}Á…É…É…Á  4(€€€€€€€‘½Õµ•¹Ð°4(€€€€€€€}É•Í½±Ù•}…Õ‘¥Ñ}‰½Õ¹‘…Éä 4(€€€€€€€€€€€‘½Õµ•¹Ð°4(€€€€€€€€€€€Í•ÑÑ¥¹Íl‰Ñ½}ÍÑ…ÉÐ‰t°4(€€€€€€€€€€€É•Í½±Ù•È°4(€€€€€€€€€€€…±±½Ý}Õ¹¥ÅÕ•}Ñ½}™¥•±õQÉÕ”°4(€€€€€€€€¤°4(€€€€¤4(€€€‰½‘å}±½…Ñ½È€ôÍ•ÑÑ¥¹Íl‰‰½‘å}ÍÑ…ÉÐ‰t4(€€€…±Ñ•É¹…Ñ•}‰½‘å}±½…Ñ½ÉÌ€ômt4(€€€™½È¡•…‘¥¹œ¥¸€¡ÍÑÉÕÑÕÉ•}µ…À½Èíô¤¹•Ð ‰¡•…‘¥¹Ìˆ°mt¤è4(€€€€€€€¡•…‘¥¹}±½…Ñ½È€ô¡•…‘¥¹œ¹•Ð ‰±½…Ñ½Èˆ°íô¤4(€€€€€€€¥˜€ 4(€€€€€€€€€€€¡•…‘¥¹œ¹•Ð ‰…ÁÁÉ½Ù•ˆ¤4(€€€€€€€€€€€…¹¡•…‘¥¹}±½…Ñ½È¹•Ð ‰Ñ•áÑ}Í¡„ÈÔØˆ¤4(€€€€€€€€€€€€ôô‰½‘å}±½…Ñ½È¹•Ð ‰Ñ•áÑ}Í¡„ÈÔØˆ¤4(€€€€€€€€€€€…¹¡•…‘¥¹œ¹•Ð ‰¹½Éµ…±¥é•‘}Ñ•áÑ}Í¡„ÈÔØˆ¤4(€€€€€€€€¤è4(€€€€€€€€€€€…±Ñ•É¹…Ñ”€ô½Áä¹‘••Á½Áä¡‰½‘å}±½…Ñ½È¤4(€€€€€€€€€€€…±Ñ•É¹…Ñ•l‰Ñ•áÑ}Í¡„ÈÔØ‰t€ô¡•…‘¥¹l‰¹½Éµ…±¥é•‘}Ñ•áÑ}Í¡„ÈÔØ‰t4(€€€€€€€€€€€…±Ñ•É¹…Ñ•}‰½‘å}±½…Ñ½ÉÌ¹…ÁÁ•¹¡…±Ñ•É¹…Ñ”¤4(€€€‰½‘å}Á…É…É…Á €ô}É•Í½±Ù•}…Õ‘¥Ñ}‰½Õ¹‘…Éä 4(€€€€€€€‘½Õµ•¹Ð°4(€€€€€€€‰½‘å}±½…Ñ½È°4(€€€€€€€É•Í½±Ù•È°4(€€€€€€€…±Ñ•É¹…Ñ•}±½…Ñ½ÉÌõ…±Ñ•É¹…Ñ•}‰½‘å}±½…Ñ½ÉÌ°4(€€€€¤4(€€€‰½‘ä€ôÍ•Ñ¥½¹}¥¹‘•á}™½É}Á…É…É…Á ¡‘½Õµ•¹Ð°‰½‘å}Á…É…É…Á ¤4(€€€™…¥±ÕÉ•Ì€ômt4(€€€™É½¹Ñ}µ…ÑÑ•È€ô€¡ÍÑÉÕÑÕÉ•}µ…À½Èíô¤¹•Ð ‰™É½¹Ñ}µ…ÑÑ•Èˆ°íô¤4(€€€¥˜™É½¹Ñ}µ…ÑÑ•È¹•Ð ‰…ÁÁÉ½Ù•ˆ¤è4(€€€€€€€Ñ¥Ñ±•}Á…É…É…Á €ô}É•Í½±Ù•}…Õ‘¥Ñ}‰½Õ¹‘…Éä 4(€€€€€€€€€€€‘½Õµ•¹Ð°4(€€€€€€€€€€€™É½¹Ñ}µ…ÑÑ•Él‰‰½½­}Ñ¥Ñ±”‰t°4(€€€€€€€€€€€É•Í½±Ù•È°4(€€€€€€€€¤4(€€€€€€€Ñ¥Ñ±•}Í•Ñ¥½¸€ôÍ•Ñ¥½¹}¥¹‘•á}™½É}Á…É…É…Á ¡‘½Õµ•¹Ð°Ñ¥Ñ±•}Á…É…É…Á ¤4(€€€€€€€Ñ½}¡•…‘¥¹Ì€ôl4(€€€€€€€€€€€Á…É…É…Á 4(€€€€€€€€€€€™½ÈÁ…É…É…Á ¥¸‘½Õµ•¹Ð¹Á…É…É…Á¡Ì4(€€€€€€€€€€€¥˜Á…É…É…Á ¹Ñ•áÐ¹ÍÑÉ¥À ¤€ôô™É½¹Ñ}µ…ÑÑ•È¹•Ð ‰Ñ½}¡•…‘¥¹}Ñ•áÐˆ¤4(€€€€€€€€€€€…¹Á…É…É…Á ¹ÍÑå±”¥Ì¹½Ð9½¹”4(€€€€€€€€€€€…¹Á…É…É…Á ¹ÍÑå±”¹¹…µ”€ôô€‰5½¹½É…Á Q=!•…‘¥¹œˆ4(€€€€€€€t4(€€€€€€€¥˜±•¸¡Ñ½}¡•…‘¥¹Ì¤€„ô€Äè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰Ñ½}¡•…‘¥¹œˆ°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè™É½¹Ñ}µ…ÑÑ•È¹•Ð ‰Ñ½}¡•…‘¥¹}Ñ•áÐˆ¤°4(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…±}½Õ¹Ðˆè±•¸¡Ñ½}¡•…‘¥¹Ì¤°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€€€€€•±¥˜Í•Ñ¥½¹}¥¹‘•á}™½É}Á…É…É…Á ¡‘½Õµ•¹Ð°Ñ½}¡•…‘¥¹ÍlÁt¤€„ôÑ½Œè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì‰ÁÉ½Á•ÉÑäˆè€‰Ñ½}¡•…‘¥¹}Í•Ñ¥½¸ˆ°€‰•áÁ•Ñ•ˆèÑ½ô4(€€€€€€€€€€€€¤4(€€€€€€€¥˜Ñ¥Ñ±•}Á…É…É…Á ¹ÍÑå±”¥Ì9½¹”½ÈÑ¥Ñ±•}Á…É…É…Á ¹ÍÑå±”¹¹…µ”€„ô€‰5½¹½É…Á 	½½¬Q¥Ñ±”ˆè(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}ÍÑå±”ˆ°€‰•áÁ•Ñ•ˆè€‰5½¹½É…Á 	½½¬Q¥Ñ±”‰ô4(€€€€€€€€€€€€¤4(€€€€€€€Ñ¥Ñ±•}…±¥¹µ•¹Ð€ôÑ¥Ñ±•}Á…É…É…Á ¹…±¥¹µ•¹Ð(€€€€€€€¥˜Ñ¥Ñ±•}…±¥¹µ•¹Ð¥Ì9½¹”…¹Ñ¥Ñ±•}Á…É…É…Á ¹ÍÑå±”¥Ì¹½Ð9½¹”è4(€€€€€€€€€€€Ñ¥Ñ±•}…±¥¹µ•¹Ð€ôÑ¥Ñ±•}Á…É…É…Á ¹ÍÑå±”¹Á…É…É…Á¡}™½Éµ…Ð¹…±¥¹µ•¹Ð4(€€€€€€€¥˜Ñ¥Ñ±•}…±¥¹µ•¹Ð€„ô]}1%9}AIIA ¹9QHè(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}…±¥¹µ•¹Ðˆ°€‰•áÁ•Ñ•ˆè€‰•¹Ñ•È‰ô(€€€€€€€€€€€€¤(€€€€€€€Ñ¥Ñ±•}Á}ÁÈ€ôÑ¥Ñ±•}Á…É…É…Á ¹}À¹ÁAÈ(€€€€€€€Ñ¥Ñ±•}‘¥É•Ñ}¹Õµ}ÁÈ€ô€ (€€€€€€€€€€€9½¹”¥˜Ñ¥Ñ±•}Á}ÁÈ¥Ì9½¹”•±Í”Ñ¥Ñ±•}Á}ÁÈ¹™¥¹¡Å¸ ‰Üé¹ÕµAÈˆ¤¤(€€€€€€€€¤(€€€€€€€¥˜Ñ¥Ñ±•}‘¥É•Ñ}¹Õµ}ÁÈ¥Ì¹½Ð9½¹”è(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ (€€€€€€€€€€€€€€€ì‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}‘¥É•Ñ}¹Õµ‰•É¥¹œˆ°€‰•áÁ•Ñ•ˆè€‰…‰Í•¹Ð‰ô(€€€€€€€€€€€€¤(€€€€€€€‘¥É•Ñ}¥¹‘•¹Ð€ô€ (€€€€€€€€€€€9½¹”¥˜Ñ¥Ñ±•}Á}ÁÈ¥Ì9½¹”•±Í”Ñ¥Ñ±•}Á}ÁÈ¹™¥¹¡Å¸ ‰Üé¥¹ˆ¤¤(€€€€€€€€¤(€€€€€€€‘¥É•Ñ}¥¹‘•¹Ñ}½¹™±¥ÑÌ€ô}¹½¹é•É½}¥¹‘•¹Ñ}…ÑÑÉ¥‰ÕÑ•Ì¡‘¥É•Ñ}¥¹‘•¹Ð¤(€€€€€€€¥˜‘¥É•Ñ}¥¹‘•¹Ñ}½¹™±¥ÑÌè(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ (€€€€€€€€€€€€€€€ì(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}‘¥É•Ñ}¥¹‘•¹Ðˆ°(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€À°(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè‘¥É•Ñ}¥¹‘•¹Ñ}½¹™±¥ÑÌ°(€€€€€€€€€€€€€€€ô(€€€€€€€€€€€€¤(€€€€€€€•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ð€ôì4(€€€€€€€€€€€€‰™½¹Ñ}¹…µ•}•…ÍÑ}…Í¥„ˆè€‹¦îG’öLˆ°4(€€€€€€€€€€€€‰™½¹Ñ}¹…µ•}…Í¥¤ˆè€‰Q¥µ•Ì9•ÜI½µ…¸ˆ°4(€€€€€€€€€€€€‰™½¹Ñ}Í¥é•}ÁÐˆè€ÈÈ°4(€€€€€€€€€€€€‰‰½±ˆèQÉÕ”°4(€€€€€€€ô4(€€€€€€€•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ð¹ÕÁ‘…Ñ”¡™É½¹Ñ}µ…ÑÑ•È¹•Ð ‰‰½½­}Ñ¥Ñ±•}™½Éµ…Ðˆ°íô¤¤4(€€€€€€€Ñ¥Ñ±•}ÍÑå±”€ôÑ¥Ñ±•}Á…É…É…Á ¹ÍÑå±”(€€€€€€€¥˜Ñ¥Ñ±•}ÍÑå±”¥Ì¹½Ð9½¹”è(€€€€€€€€€€€ÍÑå±•}Á}ÁÈ€ôÑ¥Ñ±•}ÍÑå±”¹•±•µ•¹Ð¹ÁAÈ(€€€€€€€€€€€ÍÑå±•}¥¹‘•¹Ð€ô€ (€€€€€€€€€€€€€€€9½¹”¥˜ÍÑå±•}Á}ÁÈ¥Ì9½¹”•±Í”ÍÑå±•}Á}ÁÈ¹™¥¹¡Å¸ ‰Üé¥¹ˆ¤¤(€€€€€€€€€€€€¤(€€€€€€€€€€€ÍÑå±•}¥¹‘•¹Ñ}½¹™±¥ÑÌ€ô}¹½¹é•É½}¥¹‘•¹Ñ}…ÑÑÉ¥‰ÕÑ•Ì¡ÍÑå±•}¥¹‘•¹Ð¤(€€€€€€€€€€€É•ÅÕ¥É•‘}é•É½}…ÑÑÉ¥‰ÕÑ•Ì€ôì(€€€€€€€€€€€€€€€…ÑÑÉ¥‰ÕÑ”è€ (€€€€€€€€€€€€€€€€€€€9½¹”(€€€€€€€€€€€€€€€€€€€¥˜ÍÑå±•}¥¹‘•¹Ð¥Ì9½¹”(€€€€€€€€€€€€€€€€€€€•±Í”ÍÑå±•}¥¹‘•¹Ð¹•Ð¡Å¸¡˜‰Üéí…ÑÑÉ¥‰ÕÑ•ôˆ¤¤(€€€€€€€€€€€€€€€€¤(€€€€€€€€€€€€€€€™½È…ÑÑÉ¥‰ÕÑ”¥¸€ (€€€€€€€€€€€€€€€€€€€€‰±•™Ðˆ°(€€€€€€€€€€€€€€€€€€€€‰±•™Ñ¡…ÉÌˆ°(€€€€€€€€€€€€€€€€€€€€‰É¥¡Ðˆ°(€€€€€€€€€€€€€€€€€€€€‰É¥¡Ñ¡…ÉÌˆ°(€€€€€€€€€€€€€€€€€€€€‰™¥ÉÍÑ1¥¹”ˆ°(€€€€€€€€€€€€€€€€€€€€‰™¥ÉÍÑ1¥¹•¡…ÉÌˆ°(€€€€€€€€€€€€€€€€¤(€€€€€€€€€€€ô(€€€€€€€€€€€¥˜ÍÑå±•}¥¹‘•¹Ñ}½¹™±¥ÑÌ½È…¹ä (€€€€€€€€€€€€€€€Ù…±Õ”€„ô€ˆÀˆ™½ÈÙ…±Õ”¥¸É•ÅÕ¥É•‘}é•É½}…ÑÑÉ¥‰ÕÑ•Ì¹Ù…±Õ•Ì ¤(€€€€€€€€€€€€¤è(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ (€€€€€€€€€€€€€€€€€€€ì(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}ÍÑå±•}¥¹‘•¹Ðˆ°(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€À°(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆèì(€€€€€€€€€€€€€€€€€€€€€€€€€€€€¨©É•ÅÕ¥É•‘}é•É½}…ÑÑÉ¥‰ÕÑ•Ì°(€€€€€€€€€€€€€€€€€€€€€€€€€€€€¨©ÍÑå±•}¥¹‘•¹Ñ}½¹™±¥ÑÌ°(€€€€€€€€€€€€€€€€€€€€€€€ô°(€€€€€€€€€€€€€€€€€€€ô(€€€€€€€€€€€€€€€€¤(€€€€€€€€€€€…ÑÕ…±}…Í¥¤€ôÍÑå±•}•™™•Ñ¥Ù•}™½¹Ð¡‘½Õµ•¹Ð°Ñ¥Ñ±•}ÍÑå±”°€‰…Í¥¤ˆ¥lÁt4(€€€€€€€€€€€…ÑÕ…±}•…ÍÑ}…Í¥„€ôÍÑå±•}•™™•Ñ¥Ù•}™½¹Ð¡‘½Õµ•¹Ð°Ñ¥Ñ±•}ÍÑå±”°€‰•…ÍÑÍ¥„ˆ¥lÁt4(€€€€€€€€€€€¥˜…ÑÕ…±}…Í¥¤€„ô•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ñl‰™½¹Ñ}¹…µ•}…Í¥¤‰tè4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}™½¹Ñ}¹…µ•}…Í¥¤ˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ñl‰™½¹Ñ}¹…µ•}…Í¥¤‰t°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}…Í¥¤°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€¥˜…ÑÕ…±}•…ÍÑ}…Í¥„€„ô•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ñl‰™½¹Ñ}¹…µ•}•…ÍÑ}…Í¥„‰tè4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}™½¹Ñ}¹…µ•}•…ÍÑ}…Í¥„ˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ñl‰™½¹Ñ}¹…µ•}•…ÍÑ}…Í¥„‰t°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}•…ÍÑ}…Í¥„°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€…ÑÕ…±}Í¥é”€ô9½¹”¥˜Ñ¥Ñ±•}ÍÑå±”¹™½¹Ð¹Í¥é”¥Ì9½¹”•±Í”Ñ¥Ñ±•}ÍÑå±”¹™½¹Ð¹Í¥é”¹ÁÐ4(€€€€€€€€€€€¥˜…ÑÕ…±}Í¥é”€„ô™±½…Ð¡•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ñl‰™½¹Ñ}Í¥é•}ÁÐ‰t¤è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}™½¹Ñ}Í¥é•}ÁÐˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè™±½…Ð¡•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ñl‰™½¹Ñ}Í¥é•}ÁÐ‰t¤°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}Í¥é”°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€ÍÑå±•}‰½±€ô€ 4(€€€€€€€€€€€Ñ¥Ñ±•}Á…É…É…Á ¹ÍÑå±”¥Ì¹½Ð9½¹”4(€€€€€€€€€€€…¹Ñ¥Ñ±•}Á…É…É…Á ¹ÍÑå±”¹™½¹Ð¹‰½±¥ÌQÉÕ”4(€€€€€€€€¤4(€€€€€€€¥˜…¹ä 4(€€€€€€€€€€€ÉÕ¸¹Ñ•áÐ…¹ÉÕ¸¹‰½±¥Ì¹½ÐQÉÕ”…¹¹½Ð€¡ÉÕ¸¹‰½±¥Ì9½¹”…¹ÍÑå±•}‰½±¤4(€€€€€€€€€€€™½ÈÉÕ¸¥¸Ñ¥Ñ±•}Á…É…É…Á ¹ÉÕ¹Ì4(€€€€€€€€¤è4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹¡ì‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}‰½±ˆ°€‰•áÁ•Ñ•ˆèQÉÕ•ô¤4(€€€€€€€Ñ¥Ñ±•}¥¹Ù•¹Ñ½Éä€ô¥¹Ù•¹Ñ½Éål‰Í•Ñ¥½¹Ì‰umÑ¥Ñ±•}Í•Ñ¥½¹t(€€€€€€€Ñ¥Ñ±•}Í•Ñ¥½¹}ÁÉ½Á•ÉÑ¥•Ì€ô€ (€€€€€€€€€€€9½¹”¥˜Ñ¥Ñ±•}Á}ÁÈ¥Ì9½¹”•±Í”Ñ¥Ñ±•}Á}ÁÈ¹™¥¹¡Å¸ ‰ÜéÍ•ÑAÈˆ¤¤(€€€€€€€€¤(€€€€€€€¥˜Ñ¥Ñ±•}Í•Ñ¥½¹}ÁÉ½Á•ÉÑ¥•Ì¥Ì9½¹”è(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ (€€€€€€€€€€€€€€€ì(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}Í•Ñ¥½¹}‰½Õ¹‘…Éäˆ°(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€‰…ÑÑ…¡•‘}Ñ½}‰½½­}Ñ¥Ñ±•}Á…É…É…Á ˆ°(€€€€€€€€€€€€€€€ô(€€€€€€€€€€€€¤(€€€€€€€•áÁ•Ñ•‘}Ù•ÉÑ¥…°€ô™É½¹Ñ}µ…ÑÑ•È¹•Ð ‰Ñ¥Ñ±•}Á…•}Ù•ÉÑ¥…±}…±¥¹µ•¹Ðˆ¤4(€€€€€€€¥˜•áÁ•Ñ•‘}Ù•ÉÑ¥…°…¹Ñ¥Ñ±•}¥¹Ù•¹Ñ½Éä¹•Ð ‰Ù•ÉÑ¥…±}…±¥¹µ•¹Ðˆ¤€„ô•áÁ•Ñ•‘}Ù•ÉÑ¥…°è4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰Í•Ñ¥½¸ˆèÑ¥Ñ±•}Í•Ñ¥½¸°4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰Ñ¥Ñ±•}Á…•}Ù•ÉÑ¥…±}…±¥¹µ•¹Ðˆ°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè•áÁ•Ñ•‘}Ù•ÉÑ¥…°°4(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆèÑ¥Ñ±•}¥¹Ù•¹Ñ½Éä¹•Ð ‰Ù•ÉÑ¥…±}…±¥¹µ•¹Ðˆ¤°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€€€€€•áÁ•Ñ•‘}ÍÁ…¥¹œ€ô•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ð¹•Ð ‰±¥¹•}ÍÁ…¥¹}ÁÐˆ¤4(€€€€€€€¥˜•áÁ•Ñ•‘}ÍÁ…¥¹œ¥Ì¹½Ð9½¹”…¹Ñ¥Ñ±•}ÍÑå±”¥Ì¹½Ð9½¹”è4(€€€€€€€€€€€…ÑÕ…±}ÍÁ…¥¹œ€ôÑ¥Ñ±•}ÍÑå±”¹Á…É…É…Á¡}™½Éµ…Ð¹±¥¹•}ÍÁ…¥¹œ4(€€€€€€€€€€€…ÑÕ…±}ÍÁ…¥¹}ÁÐ€ô•Ñ…ÑÑÈ¡…ÑÕ…±}ÍÁ…¥¹œ°€‰ÁÐˆ°9½¹”¤4(€€€€€€€€€€€¥˜…ÑÕ…±}ÍÁ…¥¹}ÁÐ€„ô™±½…Ð¡•áÁ•Ñ•‘}ÍÁ…¥¹œ¤è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}±¥¹•}ÍÁ…¥¹}ÁÐˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè™±½…Ð¡•áÁ•Ñ•‘}ÍÁ…¥¹œ¤°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}ÍÁ…¥¹}ÁÐ°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€•áÁ•Ñ•‘}ÉÕ±”€ô•áÁ•Ñ•‘}Ñ¥Ñ±•}™½Éµ…Ð¹•Ð ‰±¥¹•}ÍÁ…¥¹}ÉÕ±”ˆ¤4(€€€€€€€€€€€ÉÕ±•}¹…µ•Ì€ôì4(€€€€€€€€€€€€€€€]}1%9}MA%9¹Q}1MPè€‰…Ñ}±•…ÍÐˆ°4(€€€€€€€€€€€€€€€]}1%9}MA%9¹aQ1dè€‰•á…Ðˆ°4(€€€€€€€€€€€€€€€]}1%9}MA%9¹M%91è€‰Í¥¹±”ˆ°4(€€€€€€€€€€€€€€€]}1%9}MA%9¹=9}A=%9Q}%Yè€‰½¹•}Á½¥¹Ñ}™¥Ù”ˆ°4(€€€€€€€€€€€€€€€]}1%9}MA%9¹=U	1è€‰‘½Õ‰±”ˆ°4(€€€€€€€€€€€ô4(€€€€€€€€€€€…ÑÕ…±}ÉÕ±”€ôÉÕ±•}¹…µ•Ì¹•Ð¡Ñ¥Ñ±•}ÍÑå±”¹Á…É…É…Á¡}™½Éµ…Ð¹±¥¹•}ÍÁ…¥¹}ÉÕ±”¤4(€€€€€€€€€€€¥˜•áÁ•Ñ•‘}ÉÕ±”…¹…ÑÕ…±}ÉÕ±”€„ô•áÁ•Ñ•‘}ÉÕ±”è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½½­}Ñ¥Ñ±•}±¥¹•}ÍÁ…¥¹}ÉÕ±”ˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè•áÁ•Ñ•‘}ÉÕ±”°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè…ÑÕ…±}ÉÕ±”°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€¥˜¹½ÐÑ¥Ñ±•}¥¹Ù•¹Ñ½Éål‰‘¥™™•É•¹Ñ}™¥ÉÍÑ}Á…”‰tè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì‰Í•Ñ¥½¸ˆèÑ¥Ñ±•}Í•Ñ¥½¸°€‰ÁÉ½Á•ÉÑäˆè€‰Ñ¥Ñ±•}Á…•}¹Õµ‰•É}Ù¥Í¥‰±”‰ô4(€€€€€€€€€€€€¤4(€€€€€€€¥˜Ñ¥Ñ±•}¥¹Ù•¹Ñ½Éål‰™½½Ñ•É}Á…•}™¥•±‘}½Õ¹ÑÌ‰t¹•Ð ‰™¥ÉÍÐˆ°€À¤è4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì‰Í•Ñ¥½¸ˆèÑ¥Ñ±•}Í•Ñ¥½¸°€‰ÁÉ½Á•ÉÑäˆè€‰Ñ¥Ñ±•}Á…•}™¥ÉÍÑ}™½½Ñ•É}Á…•}™¥•±‰ô4(€€€€€€€€€€€€¤4(€€€‘¥É•Ñ}Á…•}‰É•…¬€ô‰½‘å}Á…É…É…Á ¹Á…É…É…Á¡}™½Éµ…Ð¹Á…•}‰É•…­}‰•™½É”4(€€€¥¹¡•É¥Ñ•‘}Á…•}‰É•…¬€ô€ 4(€€€€€€€‰½‘å}Á…É…É…Á ¹ÍÑå±”¹Á…É…É…Á¡}™½Éµ…Ð¹Á…•}‰É•…­}‰•™½É”4(€€€€€€€¥˜‰½‘å}Á…É…É…Á ¹ÍÑå±”¥Ì¹½Ð9½¹”4(€€€€€€€•±Í”9½¹”4(€€€€¤4(€€€¥˜™É½¹Ñ}µ…ÑÑ•È¹•Ð ‰…ÁÁÉ½Ù•ˆ¤è4(€€€€€€€‰½‘å}Í•Ñ¥½¹}ÑåÁ”€ô‘½Õµ•¹Ð¹Í•Ñ¥½¹Ím‰½‘åt¹}Í•ÑAÈ¹™¥¹¡Å¸ ‰ÜéÑåÁ”ˆ¤¤4(€€€€€€€‰½‘å}Í•Ñ¥½¹}Ù…±Õ”€ô€ 4(€€€€€€€€€€€€‰¹•áÑA…”ˆ4(€€€€€€€€€€€¥˜‰½‘å}Í•Ñ¥½¹}ÑåÁ”¥Ì9½¹”4(€€€€€€€€€€€•±Í”‰½‘å}Í•Ñ¥½¹}ÑåÁ”¹•Ð¡Å¸ ‰ÜéÙ…°ˆ¤°€‰¹•áÑA…”ˆ¤4(€€€€€€€€¤4(€€€€€€€¥˜‘¥É•Ñ}Á…•}‰É•…¬¥Ì¹½ÐQÉÕ”…¹‰½‘å}Í•Ñ¥½¹}Ù…±Õ”€ôô€‰½¹Ñ¥¹Õ½ÕÌˆè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰Í•Ñ¥½¸ˆè‰½‘ä°4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰‰½‘å}Á…•}‰É•…­}™½É}½¹Ñ¥¹Õ½ÕÍ}É•ÍÑ…ÉÐˆ°4(€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆèQÉÕ”°4(€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè‘¥É•Ñ}Á…•}‰É•…¬°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€•±¥˜‘¥É•Ñ}Á…•}‰É•…¬¥ÌQÉÕ”½È€ 4(€€€€€€€‘¥É•Ñ}Á…•}‰É•…¬¥Ì9½¹”…¹¥¹¡•É¥Ñ•‘}Á…•}‰É•…¬¥ÌQÉÕ”4(€€€€¤è4(€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€‰Í•Ñ¥½¸ˆè‰½‘ä°4(€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰É•‘Õ¹‘…¹Ñ}Á…•}‰É•…­}‰•™½É•}…Ñ}‰½‘å}ÍÑ…ÉÐˆ°4(€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè…±Í”°4(€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆèQÉÕ”°4(€€€€€€€€€€€ô4(€€€€€€€€¤4(€€€ÍÑ…ÉÑÌ€ôÍ•ÑÑ¥¹Ì¹•Ð ‰ÍÑ…ÉÑ}…Ðˆ°íô¤4(€€€•áÁ•Ñ•€ôíÑ½ŒèÍÑÈ¡ÍÑ…ÉÑÌ¹•Ð ‰Ñ½Œˆ°€Ä¤¤°‰½‘äèÍÑÈ¡ÍÑ…ÉÑÌ¹•Ð ‰‰½‘äˆ°€Ä¤¥ô4(€€€Í•Ñ¥½¹Ì€ô¥¹Ù•¹Ñ½Éål‰Í•Ñ¥½¹Ì‰t4(€€€™½È¥¹‘•à°ÍÑ…ÉÐ¥¸•áÁ•Ñ•¹¥Ñ•µÌ ¤è4(€€€€€€€¥˜¥¹‘•à€øô±•¸¡Í•Ñ¥½¹Ì¤½ÈÍ•Ñ¥½¹Ím¥¹‘•ául‰Á…•}¹Õµ‰•É}ÍÑ…ÉÐ‰t€„ôÍÑ…ÉÐè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì‰Í•Ñ¥½¸ˆè¥¹‘•à°€‰ÁÉ½Á•ÉÑäˆè€‰Á…•}¹Õµ‰•É}ÍÑ…ÉÐˆ°€‰•áÁ•Ñ•ˆèÍÑ…ÉÑô4(€€€€€€€€€€€€¤4(€€€¥˜Í•ÑÑ¥¹Ì¹•Ð ‰½¹Ñ¥¹Õ•}…™Ñ•É}‰½‘å}ÍÑ…ÉÐˆ°QÉÕ”¤è4(€€€€€€€™½È¥Ñ•´¥¸Í•Ñ¥½¹Ím‰½‘ä€¬€Ä€étè4(€€€€€€€€€€€¥˜¥Ñ•µl‰Á…•}¹Õµ‰•É}ÍÑ…ÉÐ‰t¥Ì¹½Ð9½¹”è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰Í•Ñ¥½¸ˆè¥Ñ•µl‰¥¹‘•à‰t°4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰Õ¹•áÁ•Ñ•‘}Á…•}¹Õµ‰•É}É•ÍÑ…ÉÐˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè¥Ñ•µl‰Á…•}¹Õµ‰•É}ÍÑ…ÉÐ‰t°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€¥˜¹½Ð¥¹Ù•¹Ñ½Éål‰½‘‘}…¹‘}•Ù•¹}Á…•Í}¡•…‘•É}™½½Ñ•È‰tè4(€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹¡ì‰ÁÉ½Á•ÉÑäˆè€‰½‘‘}…¹‘}•Ù•¹}Á…•Í}¡•…‘•É}™½½Ñ•Èˆ°€‰•áÁ•Ñ•ˆèQÉÕ•ô¤4(€€€™½È¥Ñ•´¥¸Í•Ñ¥½¹ÍmÑ½Œétè4(€€€€€€€¥˜¥Ñ•µl‰‘¥™™•É•¹Ñ}™¥ÉÍÑ}Á…”‰tè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì‰Í•Ñ¥½¸ˆè¥Ñ•µl‰¥¹‘•à‰t°€‰ÁÉ½Á•ÉÑäˆè€‰Í¡½Ý}½¹}™¥ÉÍÑ}Á…”ˆ°€‰•áÁ•Ñ•ˆèQÉÕ•ô4(€€€€€€€€€€€€¤4(€€€€€€€¥˜¥Ñ•µl‰µ¥ÍÍ¥¹}Á…•}™½½Ñ•É}ÑåÁ•Ì‰tè4(€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰Í•Ñ¥½¸ˆè¥Ñ•µl‰¥¹‘•à‰t°4(€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰Ù¥Í¥‰±•}Á…•}™½½Ñ•É}ÑåÁ•Ìˆ°4(€€€€€€€€€€€€€€€€€€€€‰µ¥ÍÍ¥¹œˆè¥Ñ•µl‰µ¥ÍÍ¥¹}Á…•}™½½Ñ•É}ÑåÁ•Ì‰t°4(€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€¤4(€€€€€€€™½È­¥¹¥¸€ ‰‘•™…Õ±Ðˆ°€‰•Ù•¸ˆ¤è4(€€€€€€€€€€€½Õ¹Ð€ô¥Ñ•µl‰™½½Ñ•É}Á…•}™¥•±‘}½Õ¹ÑÌ‰t¹•Ð¡­¥¹°€À¤4(€€€€€€€€€€€¥˜½Õ¹Ð€„ô€Äè4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰Í•Ñ¥½¸ˆè¥Ñ•µl‰¥¹‘•à‰t°4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰Á…•}™½½Ñ•É}™¥•±‘}½Õ¹Ðˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰™½½Ñ•É}ÑåÁ”ˆè­¥¹°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè€Ä°4(€€€€€€€€€€€€€€€€€€€€€€€€‰…ÑÕ…°ˆè½Õ¹Ð°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€€€€€€€€€¥˜¥Ñ•µl‰™½½Ñ•É}¹½¹}Á…•}Á…å±½…‰t¹•Ð¡­¥¹°…±Í”¤è4(€€€€€€€€€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€€€€€‰Í•Ñ¥½¸ˆè¥Ñ•µl‰¥¹‘•à‰t°4(€€€€€€€€€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰Á…•}™½½Ñ•É}¹½¹}Á…•}Á…å±½…ˆ°4(€€€€€€€€€€€€€€€€€€€€€€€€‰™½½Ñ•É}ÑåÁ”ˆè­¥¹°4(€€€€€€€€€€€€€€€€€€€€€€€€‰•áÁ•Ñ•ˆè…±Í”°4(€€€€€€€€€€€€€€€€€€€ô4(€€€€€€€€€€€€€€€€¤4(€€€¥˜¥¹Ù•¹Ñ½Éål‰½ÉÁ¡…¹}¡•…‘•É}™½½Ñ•É}Á…ÉÑÌ‰tè4(€€€€€€€™…¥±ÕÉ•Ì¹…ÁÁ•¹ 4(€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€‰ÁÉ½Á•ÉÑäˆè€‰½ÉÁ¡…¹}¡•…‘•É}™½½Ñ•É}Á…ÉÑÌˆ°4(€€€€€€€€€€€€€€€€‰…ÑÕ…±}½Õ¹Ðˆè±•¸¡¥¹Ù•¹Ñ½Éål‰½ÉÁ¡…¹}¡•…‘•É}™½½Ñ•É}Á…ÉÑÌ‰t¤°4(€€€€€€€€€€€ô4(€€€€€€€€¤4(€€€É•ÑÕÉ¸™…¥±ÕÉ•Ì°¥¹Ù•¹Ñ½Éä4(