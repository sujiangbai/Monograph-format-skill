#!/usr/bin/env python3
"""Selectively import verified field results into an audited DOCX baseline."""

from __future__ import annotations

import copy
import hashlib
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from lxml import etree
from docx.oxml.ns import qn

from _common import NS, FormatMonographError, protected_payload_manifest


FIELD_RESULT_PART = re.compile(
    r"^word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$"
)
DEFAULT_ALLOWED_FIELD_TYPES = frozenset(
    {"TOC", "PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "REF"}
)
SCALAR_FIELD_TYPES = frozenset(
    {"=", "PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "REF", "SEQ"}
)
FIELD_TEXT_TAGS = {qn("w:t"), qn("w:delText")}
VISIBLE_PAYLOAD_TAGS = {
    qn("w:t"),
    qn("w:delText"),
    qn("w:drawing"),
    qn("w:object"),
    qn("w:pict"),
}


def _normalized_instruction(value: str) -> str:
    return " ".join(value.split())


def _field_type(instruction: str) -> str:
    match = re.match(r"\s*([A-Za-z]+|=)", instruction)
    return match.group(1).upper() if match else "UNKNOWN"


@dataclass
class FieldRecord:
    form: str
    order: int
    parent_order: int | None = None
    instruction_parts: list[str] = field(default_factory=list)
    instruction: str = ""
    field_type: str = "UNKNOWN"
    simple: etree._Element | None = None
    begin: etree._Element | None = None
    separate: etree._Element | None = None
    end: etree._Element | None = None

    @property
    def key(self) -> tuple[str, str, str]:
        return self.form, self.field_type, self.instruction.casefold()

    @property
    def semantic_key(self) -> tuple[str, str]:
        return self.field_type, self.instruction.casefold()


def parse_fields(root: etree._Element) -> list[FieldRecord]:
    records: list[FieldRecord] = []
    stack: list[FieldRecord] = []
    simple_elements = set(root.iter(qn("w:fldSimple")))
    for element in root.iter():
        if element in simple_elements:
            instruction = _normalized_instruction(element.get(qn("w:instr"), ""))
            records.append(
                FieldRecord(
                    form="simple",
                    order=len(records),
                    parent_order=stack[-1].order if stack else None,
                    instruction=instruction,
                    field_type=_field_type(instruction),
                    simple=element,
                )
            )
            continue
        if element.tag == qn("w:fldChar"):
            kind = element.get(qn("w:fldCharType"))
            if kind == "begin":
                record = FieldRecord(
                    form="complex",
                    order=len(records),
                    parent_order=stack[-1].order if stack else None,
                    begin=element,
                )
                records.append(record)
                stack.append(record)
            elif kind == "separate" and stack:
                stack[-1].separate = element
            elif kind == "end" and stack:
                record = stack.pop()
                record.end = element
                record.instruction = _normalized_instruction(
                    "".join(record.instruction_parts)
                )
                record.field_type = _field_type(record.instruction)
        elif element.tag == qn("w:instrText") and stack:
            if stack[-1].separate is None:
                stack[-1].instruction_parts.append(element.text or "")
    if stack:
        raise FormatMonographError("DOCX contains an unterminated complex field.")
    for record in records:
        if record.form == "complex" and (
            record.begin is None
            or record.end is None
            or record.separate is None
            and record.field_type != "TC"
        ):
            raise FormatMonographError(
                "DOCX contains a complex field without begin, separate, and end boundaries."
            )
    return records


def _record_map(records: Iterable[FieldRecord]) -> dict[int, FieldRecord]:
    return {record.order: record for record in records}


def _has_toc_ancestor(record: FieldRecord, records: list[FieldRecord]) -> bool:
    by_order = _record_map(records)
    parent = record.parent_order
    while parent is not None:
        ancestor = by_order[parent]
        if ancestor.field_type == "TOC":
            return True
        parent = ancestor.parent_order
    return False


def _element_slice(root: etree._Element, start: etree._Element, end: etree._Element) -> list[etree._Element]:
    elements = list(root.iter())
    try:
        first = elements.index(start)
        last = elements.index(end)
    except ValueError as exc:
        raise FormatMonographError("Field boundaries are not present in their XML part.") from exc
    if first >= last:
        raise FormatMonographError("Field boundaries are out of order.")
    return elements[first + 1 : last]


def _result_text_nodes(root: etree._Element, record: FieldRecord) -> list[etree._Element]:
    if record.form == "simple":
        assert record.simple is not None
        return [item for item in record.simple.iter() if item.tag in FIELD_TEXT_TAGS]
    if record.field_type == "TC" and record.separate is None:
        return []
    assert record.separate is not None and record.end is not None
    return [
        item
        for item in _element_slice(root, record.separate, record.end)
        if item.tag in FIELD_TEXT_TAGS
    ]


def _result_payload(root: etree._Element, record: FieldRecord) -> list[etree._Element]:
    if record.form == "simple":
        assert record.simple is not None
        return list(record.simple.iter())
    if record.field_type == "TC" and record.separate is None:
        return []
    assert record.separate is not None and record.end is not None
    return _element_slice(root, record.separate, record.end)


def _approved_result_text_ids(
    root: etree._Element,
    records: list[FieldRecord],
    allowed: set[str],
) -> set[etree._Element]:
    result: set[etree._Element] = set()
    for record in records:
        if record.field_type not in allowed:
            continue
        for element in _result_text_nodes(root, record):
            result.add(element)
    return result


def _authored_text_manifest(
    root: etree._Element,
    records: list[FieldRecord],
    allowed: set[str],
) -> list[str]:
    ignored = _approved_result_text_ids(root, records, allowed)
    paragraphs = []
    for paragraph in root.xpath(".//w:p", namespaces=NS):
        value = "".join(
            element.text or ""
            for element in paragraph.iter()
            if element.tag in FIELD_TEXT_TAGS and element not in ignored
        )
        if value:
            paragraphs.append(value)
    return paragraphs


def _field_contract(records: list[FieldRecord]) -> list[tuple[str, str]]:
    return [
        record.semantic_key
        for record in records
        if not _has_toc_ancestor(record, records)
    ]


def _record_context_key(
    root: etree._Element,
    record: FieldRecord,
    records: list[FieldRecord],
    allowed: set[str],
) -> tuple[str, str] | None:
    marker = record.simple if record.form == "simple" else record.begin
    if marker is None:
        return None
    paragraph = marker
    while paragraph is not None and paragraph.tag != qn("w:p"):
        paragraph = paragraph.getparent()
    if paragraph is None:
        return None
    paragraph_id = paragraph.get(qn("w14:paraId"))
    if paragraph_id:
        return "paragraph_id", paragraph_id.casefold()
    ignored = _approved_result_text_ids(root, records, allowed)
    authored = "".join(
        element.text or ""
        for element in paragraph.iter()
        if element.tag in FIELD_TEXT_TAGS and element not in ignored
    )
    if not authored:
        return None
    return "authored_text", hashlib.sha256(authored.encode("utf-8")).hexdigest()


def _validate_page_offset_formula(
    record: FieldRecord,
    records: list[FieldRecord],
) -> None:
    children = [item for item in records if item.parent_order == record.order]
    if (
        record.form != "complex"
        or re.sub(r"\s+", "", record.instruction) != "=-1"
        or len(children) != 1
        or children[0].form != "complex"
        or children[0].field_type != "PAGE"
        or children[0].instruction.strip().upper() != "PAGE"
        or any(item.parent_order == children[0].order for item in records)
    ):
        raise FormatMonographError(
            "Only the exact core-generated PAGE-minus-one display formula can be refreshed."
        )


def _bookmark_ranges(root: etree._Element) -> dict[str, str]:
    starts: dict[str, tuple[int, str]] = {}
    result: dict[str, str] = {}
    elements = list(root.iter())
    positions = {element: index for index, element in enumerate(elements)}
    for element in root.xpath(".//w:bookmarkStart", namespaces=NS):
        name = element.get(qn("w:name"))
        identifier = element.get(qn("w:id"))
        if name and identifier is not None:
            starts[identifier] = (positions[element], name)
    for element in root.xpath(".//w:bookmarkEnd", namespaces=NS):
        identifier = element.get(qn("w:id"))
        if identifier not in starts:
            continue
        start, name = starts[identifier]
        end = positions[element]
        text = "".join(
            item.text or ""
            for item in elements[start + 1 : end]
            if item.tag in FIELD_TEXT_TAGS
        )
        result[name] = hashlib.sha256(text.encode("utf-8")).hexdigest()
    return result


def _canonical_attributes(element: etree._Element | None) -> tuple[tuple[str, str], ...]:
    if element is None:
        return ()
    return tuple(
        sorted((etree.QName(name).localname, value) for name, value in element.attrib.items())
    )


def _section_manifest(root: etree._Element) -> list[tuple[Any, ...]]:
    result: list[tuple[Any, ...]] = []
    for section in root.xpath(".//w:sectPr", namespaces=NS):
        section_type = section.find(qn("w:type"))
        pg_num = section.find(qn("w:pgNumType"))
        vertical = section.find(qn("w:vAlign"))
        result.append(
            (
                "nextPage"
                if section_type is None
                else section_type.get(qn("w:val"), "nextPage"),
                _canonical_attributes(section.find(qn("w:pgSz"))),
                _canonical_attributes(section.find(qn("w:pgMar"))),
                None
                if pg_num is None
                else (
                    pg_num.get(qn("w:start")),
                    pg_num.get(qn("w:fmt"), "decimal"),
                ),
                section.find(qn("w:titlePg")) is not None,
                "top" if vertical is None else vertical.get(qn("w:val"), "top"),
            )
        )
    return result


def _validate_backend_part(
    baseline_root: etree._Element,
    refreshed_root: etree._Element,
    baseline_records: list[FieldRecord],
    refreshed_records: list[FieldRecord],
    allowed: set[str],
) -> list[str]:
    if _field_contract(baseline_records) != _field_contract(refreshed_records):
        raise FormatMonographError(
            "Target application changed field instructions, ordering, or boundaries."
        )
    if _authored_text_manifest(
        baseline_root, baseline_records, allowed
    ) != _authored_text_manifest(refreshed_root, refreshed_records, allowed):
        raise FormatMonographError(
            "Target application changed visible content outside approved field results."
        )
    baseline_bookmarks = _bookmark_ranges(baseline_root)
    refreshed_bookmarks = _bookmark_ranges(refreshed_root)
    for name, digest in baseline_bookmarks.items():
        if refreshed_bookmarks.get(name) != digest:
            raise FormatMonographError(
                "Target application changed or removed an approved bookmark target."
            )
    extra_bookmarks = set(refreshed_bookmarks) - set(baseline_bookmarks)
    if any(not re.fullmatch(r"_(?:Toc\d+|GoBack)", name) for name in extra_bookmarks):
        raise FormatMonographError(
            "Target application added a non-TOC bookmark outside the approved contract."
        )
    if _section_manifest(baseline_root) != _section_manifest(refreshed_root):
        raise FormatMonographError(
            "Target application changed approved pagination or section structure."
        )
    return [
        "field_form_normalization",
        "run_serialization",
        "proofing_or_revision_metadata",
        "table_serialization",
    ]


def _document_relationships(package: zipfile.ZipFile) -> dict[str, str]:
    name = "word/_rels/document.xml.rels"
    if name not in package.namelist():
        return {}
    root = etree.fromstring(package.read(name))
    result = {}
    for relationship in root:
        identifier = relationship.get("Id")
        target = relationship.get("Target")
        if identifier and target:
            result[identifier] = posixpath.normpath(posixpath.join("word", target))
    return result


def _story_roles(package: zipfile.ZipFile) -> dict[tuple[str, int, str], str]:
    if "word/document.xml" not in package.namelist():
        return {}
    relationships = _document_relationships(package)
    document = etree.fromstring(package.read("word/document.xml"))
    inherited: dict[str, dict[str, str]] = {"header": {}, "footer": {}}
    roles: dict[tuple[str, int, str], str] = {}
    sections = document.xpath(".//w:sectPr", namespaces=NS)
    for section_index, section in enumerate(sections):
        for story in ("header", "footer"):
            for reference in section.xpath(f"./w:{story}Reference", namespaces=NS):
                kind = reference.get(qn("w:type"), "default")
                rel_id = reference.get(qn("r:id"))
                target = relationships.get(rel_id or "")
                if target:
                    inherited[story][kind] = target
            for kind, target in inherited[story].items():
                roles[(story, section_index, kind)] = target
    return roles


def _semantic_part_sources(
    baseline: zipfile.ZipFile, refreshed: zipfile.ZipFile
) -> tuple[dict[str, list[str]], list[str]]:
    baseline_roles = _story_roles(baseline)
    refreshed_roles = _story_roles(refreshed)
    if set(baseline_roles) != set(refreshed_roles):
        raise FormatMonographError(
            "Target application changed the effective header or footer role set."
        )
    sources: dict[str, list[str]] = {}
    for role, baseline_part in baseline_roles.items():
        refreshed_part = refreshed_roles[role]
        values = sources.setdefault(baseline_part, [])
        if refreshed_part not in values:
            values.append(refreshed_part)
    return sources, [
        "header_footer_relationship_serialization",
        "header_footer_part_renumbering",
    ]


def _set_scalar_result(
    baseline_root: etree._Element,
    baseline: FieldRecord,
    refreshed_root: etree._Element,
    refreshed: FieldRecord,
) -> None:
    if _is_dirty(refreshed):
        raise FormatMonographError(
            f"Target application left field {baseline.field_type} dirty after refresh."
        )
    payload = _result_payload(refreshed_root, refreshed)
    if any(
        element.tag in {qn("w:drawing"), qn("w:object"), qn("w:pict"), qn("w:tbl")}
        for element in payload
    ):
        raise FormatMonographError(
            f"Field {baseline.field_type} returned a non-scalar result payload."
        )
    value = "".join(
        element.text or "" for element in payload if element.tag in FIELD_TEXT_TAGS
    )
    if baseline.field_type in {
        "=",
        "PAGE",
        "NUMPAGES",
        "SECTIONPAGES",
        "PAGEREF",
    }:
        if value.strip() and not re.fullmatch(r"[\s0-9IVXLCDMivxlcdm.\-–—]+", value):
            raise FormatMonographError(
                f"Field {baseline.field_type} returned a non-page scalar result."
            )
    targets = _result_text_nodes(baseline_root, baseline)
    if not targets:
        raise FormatMonographError(
            f"Baseline field {baseline.field_type} has no result text container."
        )
    targets[0].text = value
    for target in targets[1:]:
        target.text = ""
    dirty = qn("w:dirty")
    marker = baseline.simple if baseline.form == "simple" else baseline.begin
    assert marker is not None
    marker.attrib.pop(dirty, None)


def _body_child(body: etree._Element, element: etree._Element) -> etree._Element | None:
    current = element
    while current.getparent() is not None and current.getparent() is not body:
        current = current.getparent()
    return current if current.getparent() is body else None


def _toc_span(root: etree._Element, record: FieldRecord) -> tuple[etree._Element, int, int]:
    bodies = root.xpath("/w:document/w:body", namespaces=NS)
    if len(bodies) != 1 or record.begin is None or record.end is None:
        raise FormatMonographError("TOC must be a complete complex field in the main body.")
    body = bodies[0]
    start_block = _body_child(body, record.begin)
    end_block = _body_child(body, record.end)
    if start_block is None or end_block is None:
        raise FormatMonographError("TOC field boundaries are outside the main body.")
    start = body.index(start_block)
    end = body.index(end_block)
    if start > end:
        raise FormatMonographError("TOC field boundaries are out of order.")
    inside = False
    for block in list(body)[start : end + 1]:
        for element in block.iter():
            if element is record.begin:
                inside = True
                continue
            if element is record.end:
                inside = False
                continue
            if not inside and element.tag in VISIBLE_PAYLOAD_TAGS:
                if element.tag not in FIELD_TEXT_TAGS or (element.text or ""):
                    raise FormatMonographError(
                        "TOC boundary blocks share authored content with the field."
                    )
    return body, start, end


def _unwrap_element(element: etree._Element) -> None:
    parent = element.getparent()
    if parent is None:
        return
    index = parent.index(element)
    for child in list(element):
        element.remove(child)
        parent.insert(index, child)
        index += 1
    parent.remove(element)


def _toc_entry_level(paragraph: etree._Element) -> int | None:
    styles = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    if not styles:
        return None
    match = re.fullmatch(r"TOC\s*([1-9])", str(styles[-1]), re.IGNORECASE)
    return None if match is None else int(match.group(1))


def _toc_entry_text(paragraph: etree._Element) -> tuple[str, str]:
    segments: list[list[str]] = [[]]
    tab_count = 0
    for element in paragraph.iter():
        if element.tag == qn("w:tab"):
            segments.append([])
            tab_count += 1
            continue
        if element.tag not in FIELD_TEXT_TAGS:
            continue
        segments[-1].append(element.text or "")
    values = [" ".join("".join(segment).split()) for segment in segments]
    values = [value for value in values if value]
    if len(values) < 2:
        if values and tab_count:
            return "", values[0]
        return (values[0] if values else ""), ""
    return " ".join(values[:-1]), values[-1]


def _toc_contract_text(value: str, level: int, kind: str) -> str:
    if kind == "heading":
        if level == 1:
            pattern = re.compile(
                r"^\s*第\s*[0-9一二三四五六七八九十百]+\s*章\s*"
            )
        else:
            pattern = re.compile(rf"^\s*\d+(?:[.-]\d+){{{level - 1}}}\s*")
        match = pattern.match(value)
        if match:
            value = value[match.end() :]
    return " ".join(value.split())


def _validate_toc_contract(
    container: etree._Element,
    contract: list[dict[str, Any]],
) -> None:
    entries: list[dict[str, Any]] = []
    for paragraph in container.xpath(".//w:p", namespaces=NS):
        level = _toc_entry_level(paragraph)
        title, page = _toc_entry_text(paragraph)
        has_visible_text = bool(title or page)
        if level is None:
            if has_visible_text and not paragraph.xpath(
                ".//w:instrText", namespaces=NS
            ):
                raise FormatMonographError(
                    "Refreshed TOC contains text outside a TOC entry."
                )
            continue
        if not title:
            raise FormatMonographError("Refreshed TOC contains an empty entry.")
        if not re.fullmatch(r"[0-9IVXLCDMivxlcdm.\-–—]+", page):
            raise FormatMonographError(
                "Refreshed TOC entry has no verifiable page value."
            )
        internal_target = bool(
            paragraph.xpath(".//w:hyperlink[@w:anchor]", namespaces=NS)
        ) or bool(
            re.search(
                r"\bPAGEREF\b",
                " ".join(
                    paragraph.xpath(".//w:instrText/text()", namespaces=NS)
                ),
                re.IGNORECASE,
            )
        )
        if not internal_target:
            raise FormatMonographError(
                "Refreshed TOC entry has no internal target."
            )
        entries.append({"level": level, "title": title})
    if len(entries) != len(contract):
        raise FormatMonographError(
            "Refreshed TOC entry count does not match approved sources."
        )
    for entry, expected in zip(entries, contract):
        level = int(expected.get("level", 0))
        kind = str(expected.get("kind", ""))
        if entry["level"] != level:
            raise FormatMonographError(
                "Refreshed TOC entry level does not match approved sources."
            )
        value = _toc_contract_text(entry["title"], level, kind)
        digest = hashlib.sha256(value.encode("utf-8")).hexdigest()
        if not value or digest != expected.get("text_sha256"):
            raise FormatMonographError(
                "Refreshed TOC entry text or order does not match approved sources."
            )


def _sanitize_toc_blocks(
    blocks: list[etree._Element],
    toc_contract: list[dict[str, Any]] | None = None,
) -> None:
    container = etree.Element("container")
    for block in blocks:
        container.append(block)
    prohibited = {
        "AlternateContent",
        "altChunk",
        "control",
        "drawing",
        "imagedata",
        "object",
        "oleObject",
        "pict",
        "sdt",
        "sectPr",
        "shape",
        "tbl",
        "txbxContent",
    }
    if any(etree.QName(element).localname in prohibited for element in container.iter()):
        raise FormatMonographError("Refreshed TOC contains a non-text result payload.")
    if container.xpath(".//*[@r:id or @r:embed or @r:link]", namespaces=NS):
        raise FormatMonographError("Refreshed TOC contains an external relationship.")
    if toc_contract is not None:
        _validate_toc_contract(container, toc_contract)
    for hyperlink in list(container.xpath(".//w:hyperlink", namespaces=NS)):
        _unwrap_element(hyperlink)
    for bookmark in list(
        container.xpath(".//w:bookmarkStart | .//w:bookmarkEnd", namespaces=NS)
    ):
        parent = bookmark.getparent()
        if parent is not None:
            parent.remove(bookmark)
    nested = parse_fields(container)
    for record in reversed(nested):
        if record.field_type == "TOC":
            continue
        if record.form == "simple":
            assert record.simple is not None
            _unwrap_element(record.simple)
            continue
        assert record.begin is not None and record.separate is not None and record.end is not None
        elements = _element_slice(container, record.begin, record.end)
        before_result = True
        for element in elements:
            if element is record.separate:
                before_result = False
                continue
            if before_result and element.tag == qn("w:instrText"):
                element.text = ""
        for marker in (record.begin, record.separate, record.end):
            parent = marker.getparent()
            if parent is not None:
                parent.remove(marker)
    for block in list(container):
        container.remove(block)


def _replace_toc_result(
    baseline_root: etree._Element,
    baseline: FieldRecord,
    refreshed_root: etree._Element,
    refreshed: FieldRecord,
    toc_contract: list[dict[str, Any]] | None = None,
) -> None:
    baseline_body, baseline_start, baseline_end = _toc_span(baseline_root, baseline)
    refreshed_body, refreshed_start, refreshed_end = _toc_span(refreshed_root, refreshed)
    blocks = [
        copy.deepcopy(block)
        for block in list(refreshed_body)[refreshed_start : refreshed_end + 1]
    ]
    _sanitize_toc_blocks(blocks, toc_contract)
    container = etree.Element("container")
    for block in blocks:
        container.append(block)
    copied_toc = next(
        (record for record in parse_fields(container) if record.field_type == "TOC"),
        None,
    )
    if copied_toc is None or copied_toc.begin is None:
        raise FormatMonographError("Refreshed TOC lost its outer field boundary.")
    instruction_nodes = [
        element
        for element in _element_slice(container, copied_toc.begin, copied_toc.separate)
        if element.tag == qn("w:instrText")
    ]
    if not instruction_nodes:
        raise FormatMonographError("Refreshed TOC lost its field instruction.")
    instruction_nodes[0].text = " " + baseline.instruction + " "
    for element in instruction_nodes[1:]:
        element.text = ""
    copied_toc.begin.attrib.pop(qn("w:dirty"), None)
    for block in list(container):
        container.remove(block)
    for index in range(baseline_end, baseline_start - 1, -1):
        baseline_body.remove(baseline_body[index])
    for offset, block in enumerate(blocks):
        baseline_body.insert(baseline_start + offset, block)


def _matched_records(
    baseline_root: etree._Element,
    baseline: list[FieldRecord],
    refreshed_root: etree._Element,
    refreshed: list[FieldRecord],
    allowed: set[str],
) -> list[tuple[FieldRecord, FieldRecord]]:
    baseline_groups: dict[tuple[str, str], list[FieldRecord]] = {}
    refreshed_groups: dict[tuple[str, str], list[FieldRecord]] = {}
    for record in baseline:
        if _has_toc_ancestor(record, baseline):
            continue
        baseline_groups.setdefault(record.semantic_key, []).append(record)
    for record in refreshed:
        if _has_toc_ancestor(record, refreshed):
            continue
        refreshed_groups.setdefault(record.semantic_key, []).append(record)
    if set(baseline_groups) != set(refreshed_groups):
        raise FormatMonographError("Target application changed the field set.")
    matches: list[tuple[FieldRecord, FieldRecord]] = []
    for key, baseline_values in baseline_groups.items():
        refreshed_values = refreshed_groups[key]
        if len(baseline_values) != len(refreshed_values):
            raise FormatMonographError("Target application changed a field occurrence count.")
        if len(baseline_values) == 1:
            matches.append((baseline_values[0], refreshed_values[0]))
            continue
        if key[0] == "TC" and "TC" in allowed:
            matches.extend(zip(baseline_values, refreshed_values))
            continue
        baseline_contexts = {
            _record_context_key(baseline_root, record, baseline, allowed): record
            for record in baseline_values
        }
        refreshed_contexts = {
            _record_context_key(refreshed_root, record, refreshed, allowed): record
            for record in refreshed_values
        }
        if (
            None in baseline_contexts
            or None in refreshed_contexts
            or len(baseline_contexts) != len(baseline_values)
            or len(refreshed_contexts) != len(refreshed_values)
            or set(baseline_contexts) != set(refreshed_contexts)
        ):
            raise FormatMonographError(
                "Duplicate fields cannot be matched by a unique authored paragraph context."
            )
        matches.extend(
            (baseline_contexts[context], refreshed_contexts[context])
            for context in baseline_contexts
        )
    return sorted(matches, key=lambda pair: pair[0].order)


def _is_dirty(record: FieldRecord) -> bool:
    marker = record.simple if record.form == "simple" else record.begin
    if marker is None:
        return False
    return marker.get(qn("w:dirty")) in {"1", "true", "on"}


def selective_field_result_writeback(
    baseline_path: Path,
    refreshed_path: Path,
    output_path: Path,
    *,
    allowed_field_types: Iterable[str] = DEFAULT_ALLOWED_FIELD_TYPES,
    toc_contract: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Write only approved field results into a package copied from ``baseline_path``."""
    allowed = {str(value).upper() for value in allowed_field_types}
    if not allowed:
        raise FormatMonographError("Selective field writeback requires approved field types.")
    if toc_contract is not None and (
        not isinstance(toc_contract, list)
        or not toc_contract
        or any(
            not isinstance(item, dict)
            or int(item.get("level", 0)) not in {1, 2, 3, 4}
            or item.get("kind") not in {"heading", "appendix"}
            or not re.fullmatch(r"[0-9a-f]{64}", str(item.get("text_sha256", "")))
            for item in toc_contract
        )
    ):
        raise FormatMonographError("TOC result contract is invalid or empty.")
    unsupported = allowed - (DEFAULT_ALLOWED_FIELD_TYPES | {"=", "SEQ", "TC"})
    if unsupported:
        raise FormatMonographError(
            "Selective field writeback received unsupported field types: "
            + ", ".join(sorted(unsupported))
        )
    if protected_payload_manifest(baseline_path) != protected_payload_manifest(refreshed_path):
        raise FormatMonographError(
            "Target application changed a protected media, formula, or embedded payload."
        )

    patched_parts: dict[str, bytes] = {}
    updated_fields = 0
    matched_fields = 0
    discarded_categories: set[str] = set()
    all_field_types: set[str] = set()
    unapproved_dirty_fields = 0
    approved_source_fields = 0
    verified_toc_fields = 0
    with zipfile.ZipFile(baseline_path) as baseline, zipfile.ZipFile(refreshed_path) as refreshed:
        refreshed_names = set(refreshed.namelist())
        story_sources, story_discarded = _semantic_part_sources(baseline, refreshed)
        discarded_categories.update(story_discarded)
        for name in baseline.namelist():
            if not FIELD_RESULT_PART.fullmatch(name):
                continue
            baseline_data = baseline.read(name)
            baseline_root = etree.fromstring(baseline_data)
            baseline_records = parse_fields(baseline_root)
            if "=" in allowed:
                for record in baseline_records:
                    if record.field_type == "=":
                        _validate_page_offset_formula(record, baseline_records)
            all_field_types.update(record.field_type for record in baseline_records)
            unapproved_dirty_fields += sum(
                1
                for record in baseline_records
                if record.field_type not in allowed and _is_dirty(record)
            )
            source_names = story_sources.get(name)
            if source_names is None:
                source_names = [name] if name in refreshed_names else []
            if not source_names:
                if baseline_records:
                    raise FormatMonographError(
                        "Target application removed a field-bearing XML part."
                    )
                continue
            refreshed_options = []
            for source_name in source_names:
                if source_name not in refreshed_names:
                    raise FormatMonographError(
                        "Target application returned a missing header or footer part."
                    )
                refreshed_root = etree.fromstring(refreshed.read(source_name))
                refreshed_records = parse_fields(refreshed_root)
                if "=" in allowed:
                    for record in refreshed_records:
                        if record.field_type == "=":
                            _validate_page_offset_formula(record, refreshed_records)
                discarded_categories.update(
                    _validate_backend_part(
                        baseline_root,
                        refreshed_root,
                        baseline_records,
                        refreshed_records,
                        allowed,
                    )
                )
                refreshed_options.append((refreshed_root, refreshed_records))
            refreshed_root, refreshed_records = refreshed_options[0]
            if not baseline_records and not refreshed_records:
                continue
            matches = _matched_records(
                baseline_root,
                baseline_records,
                refreshed_root,
                refreshed_records,
                allowed,
            )
            for _option_root, option_records in refreshed_options[1:]:
                _matched_records(
                    baseline_root,
                    baseline_records,
                    _option_root,
                    option_records,
                    allowed,
                )
            matched_fields += len(matches)
            toc_matches = [pair for pair in matches if pair[0].field_type == "TOC"]
            if len(toc_matches) > 1:
                raise FormatMonographError("Only one approved TOC field is supported per part.")
            for baseline_field, refreshed_field in matches:
                if baseline_field.field_type not in allowed:
                    continue
                if baseline_field.field_type == "TOC":
                    continue
                if baseline_field.field_type == "TC":
                    approved_source_fields += 1
                    continue
                if baseline_field.field_type not in SCALAR_FIELD_TYPES:
                    raise FormatMonographError(
                        f"Field type {baseline_field.field_type} has no selective handler."
                    )
                _set_scalar_result(
                    baseline_root,
                    baseline_field,
                    refreshed_root,
                    refreshed_field,
                )
                updated_fields += 1
            if toc_matches:
                if _is_dirty(toc_matches[0][1]):
                    raise FormatMonographError(
                        "Target application left the TOC field dirty after refresh."
                    )
                _replace_toc_result(
                    baseline_root,
                    toc_matches[0][0],
                    refreshed_root,
                    toc_matches[0][1],
                    toc_contract,
                )
                updated_fields += 1
                verified_toc_fields += 1
            candidate = etree.tostring(
                baseline_root,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )
            if candidate != baseline_data:
                patched_parts[name] = candidate
        if "word/settings.xml" in baseline.namelist():
            settings = etree.fromstring(baseline.read("word/settings.xml"))
            changed = False
            for update in list(settings.xpath("./w:updateFields", namespaces=NS)):
                settings.remove(update)
                changed = True
            if changed:
                patched_parts["word/settings.xml"] = etree.tostring(
                    settings,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                )

    if toc_contract is not None and verified_toc_fields != 1:
        raise FormatMonographError(
            "Approved TOC result contract requires exactly one verified TOC field."
        )

    temp_output = output_path.with_name(f".{output_path.stem}.field-writeback.tmp.docx")
    temp_output.unlink(missing_ok=True)
    try:
        with zipfile.ZipFile(baseline_path) as baseline, zipfile.ZipFile(
            temp_output, "w", compression=zipfile.ZIP_DEFLATED
        ) as output:
            for info in baseline.infolist():
                output.writestr(info, patched_parts.get(info.filename, baseline.read(info.filename)))
        if protected_payload_manifest(temp_output) != protected_payload_manifest(baseline_path):
            raise FormatMonographError("Selective field writeback changed a protected payload.")
        temp_output.replace(output_path)
    finally:
        temp_output.unlink(missing_ok=True)

    return {
        "status": "selective_verified",
        "allowed_field_types": sorted(allowed),
        "matched_fields": matched_fields,
        "updated_fields": updated_fields,
        "approved_source_fields": approved_source_fields,
        "toc_result_status": (
            "verified_text_only" if toc_contract is not None else "not_requested"
        ),
        "toc_source_count": 0 if toc_contract is None else len(toc_contract),
        "unapproved_field_types": sorted(all_field_types - allowed),
        "unapproved_dirty_fields": unapproved_dirty_fields,
        "patched_parts": sorted(patched_parts),
        "discarded_backend_differences": sorted(discarded_categories),
    }
