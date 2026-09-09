from __future__ import annotations

import tempfile
import unittest
import zipfile
import json
import copy
import subprocess
from pathlib import Path

from docx import Document
from docx.enum.section import WD_HEADER_FOOTER, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
import sys

sys.path.insert(0, str(SCRIPTS))

from _common import NS, FormatMonographError  # noqa: E402
from field_writeback import (  # noqa: E402
    _semantic_part_sources,
    _story_roles,
    parse_fields,
    selective_field_result_writeback,
)
from finalize_docx import (  # noqa: E402
    apply_measured_layout_adjustments,
    apply_page_display_offsets,
    external_measure,
    external_verify,
    effective_font_failures,
    remove_measured_block_spacers,
)


def add_complex_field(paragraph, instruction: str, value: str, *, dirty: bool = True) -> None:
    begin_run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    if dirty:
        begin.set(qn("w:dirty"), "true")
    begin_run._r.append(begin)
    instruction_run = paragraph.add_run()
    instruction_node = OxmlElement("w:instrText")
    instruction_node.set(qn("xml:space"), "preserve")
    instruction_node.text = f" {instruction} "
    instruction_run._r.append(instruction_node)
    separate_run = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    separate_run._r.append(separate)
    paragraph.add_run(value)
    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)


def add_simple_field(paragraph, instruction: str, value: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = value
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def add_nested_formula(
    paragraph,
    nested_instruction: str,
    value: str,
) -> None:
    begin = paragraph.add_run()
    outer_begin = OxmlElement("w:fldChar")
    outer_begin.set(qn("w:fldCharType"), "begin")
    begin._r.append(outer_begin)
    instruction = paragraph.add_run()
    instruction_node = OxmlElement("w:instrText")
    instruction_node.text = " = "
    instruction._r.append(instruction_node)
    add_complex_field(paragraph, nested_instruction, "2", dirty=False)
    suffix = paragraph.add_run()
    suffix_node = OxmlElement("w:instrText")
    suffix_node.text = " - 1 "
    suffix._r.append(suffix_node)
    separator = paragraph.add_run()
    separate_node = OxmlElement("w:fldChar")
    separate_node.set(qn("w:fldCharType"), "separate")
    separator._r.append(separate_node)
    paragraph.add_run(value)
    end = paragraph.add_run()
    end_node = OxmlElement("w:fldChar")
    end_node.set(qn("w:fldCharType"), "end")
    end._r.append(end_node)


def rewrite_package(source: Path, output: Path, transform) -> None:
    with zipfile.ZipFile(source) as package, zipfile.ZipFile(
        output, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in package.infolist():
            target.writestr(info, transform(info.filename, package.read(info.filename)))


def field_values(path: Path, part: str = "word/document.xml") -> list[str]:
    with zipfile.ZipFile(path) as package:
        root = etree.fromstring(package.read(part))
    values = []
    elements = list(root.iter())
    for record in parse_fields(root):
        if record.form != "complex":
            continue
        start = elements.index(record.separate)
        end = elements.index(record.end)
        values.append(
            "".join(
                element.text or ""
                for element in elements[start + 1 : end]
                if element.tag == qn("w:t")
            )
        )
    return values


class V032SelectiveFieldWritebackTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_final_audit_body_scalar_refresh_and_toc_shift_remain_strict(self) -> None:
        from docx.shared import Pt
        from field_writeback import _result_text_nodes
        from structure_map import candidate_structure_map
        from test_v024_finalization import expand_toc_cache
        from test_v11_execution import approved_v11_profile

        source = self.root / "audit-source.docx"
        document = Document()
        indent = OxmlElement("w:ind")
        for name in ("left", "leftChars", "right", "rightChars", "firstLine", "firstLineChars"):
            indent.set(qn("w:" + name), "0")
        document.styles["Heading 1"].element.get_or_add_pPr().append(indent)
        add_complex_field(document.add_paragraph(), 'TOC \\o "1-3" \\h \\z \\u', "TOC placeholder")
        document.add_paragraph("Heading", style="Heading 1")
        body = document.add_paragraph("Current page ")
        add_complex_field(body, "PAGE", "0")
        body.add_run(" of ")
        add_complex_field(body, "NUMPAGES", "0")
        add_complex_field(body, "REF anchor", "0")
        add_complex_field(body, "SECTIONPAGES", "0")
        add_complex_field(body, "DATE", "Unapproved cache", dirty=False)
        reference = document.add_paragraph("Chapter location ")
        add_complex_field(reference, "PAGEREF anchor", "0")
        document.add_paragraph("After anchor")
        document.save(source)
        structure = candidate_structure_map(source)
        structure["status"] = "approved"
        structure["toc_source"].update(approved=True, mode="heading_styles", levels=3)
        for heading in structure["headings"]:
            heading["approved"] = True
        for entry in structure["paragraph_roles"]:
            if entry["locator"].get("paragraph") in (2, 3):
                entry.update(approved=True, role="body_text", canonical_role="body")
        map_path = self.root / "audit-map.json"
        map_path.write_text(json.dumps(structure))
        profile = approved_v11_profile()
        rule = copy.deepcopy(profile["rules"][0])
        rule.update(id="FMT-BODY-901", category="body",
                    selector={"kind": "paragraph_role", "value": "body_text"},
                    properties={"font_size_pt": 12})
        profile["rules"] = [rule]
        profile_path = self.root / "audit-profile.json"
        profile_path.write_text(json.dumps(profile))

        # Formatting differs from the source; identity must not suppress the
        # actual rule audit. Field instructions and all non-cache XML stay exact.
        for paragraph in (body, reference):
            for run in paragraph.runs:
                run.font.size = Pt(12)
            for record in parse_fields(paragraph._p):
                if record.field_type in {"PAGE", "NUMPAGES", "PAGEREF", "REF", "SECTIONPAGES"}:
                    for node in _result_text_nodes(paragraph._p, record):
                        node.text = "3"
                    record.begin.attrib.pop(qn("w:dirty"), None)
        formatted = self.root / "audit-formatted.docx"
        document.save(formatted)
        final = self.root / "audit-final.docx"
        expand_toc_cache(formatted, final)
        original_source = source.read_bytes()
        original_map = map_path.read_bytes()

        def audit(path):
            return subprocess.run(
                [sys.executable, str(SCRIPTS / "audit_docx.py"), str(source), str(path),
                 "--profile", str(profile_path), "--structure-map", str(map_path)],
                capture_output=True, text=True, cwd=SCRIPTS.parent,
            )

        passed = audit(final)
        self.assertEqual(0, passed.returncode, passed.stdout + passed.stderr)
        self.assertTrue(json.loads(passed.stdout)["passed"])
        for mode in ("authored", "instruction", "boundary", "ref_instruction", "sectionpages_instruction", "ref_boundary", "sectionpages_boundary", "unapproved_cache", "duplicate", "format", "object", "section_boundary"):
            with self.subTest(mode=mode):
                changed = Document(final)
                target = next(p for p in changed.paragraphs if p.text.startswith("Current page"))
                if mode == "authored":
                    target.runs[0].text = "Changed author text "
                elif mode == "instruction":
                    target._p.xpath(".//w:instrText")[0].text = " NUMPAGES "
                elif mode == "boundary":
                    end = target._p.xpath(".//w:fldChar[@w:fldCharType='end']")[-1]
                    end.getparent().remove(end)
                elif mode in {"ref_instruction", "sectionpages_instruction", "ref_boundary", "sectionpages_boundary"}:
                    kind = mode.split("_")[0].upper()
                    record = next(r for r in parse_fields(target._p) if r.field_type == kind)
                    if mode.endswith("boundary"):
                        record.end.getparent().remove(record.end)
                    else:
                        instruction = record.begin.getparent().getnext().find(qn("w:instrText"))
                        instruction.text = " PAGE "
                elif mode == "unapproved_cache":
                    record = next(r for r in parse_fields(target._p) if r.field_type == "DATE")
                    _result_text_nodes(target._p, record)[0].text = "Changed unapproved cache"
                elif mode == "duplicate":
                    target._p.addnext(copy.deepcopy(target._p))
                elif mode == "format":
                    target.runs[0].font.size = Pt(8)
                elif mode == "object":
                    target.add_run()._r.append(OxmlElement("w:object"))
                elif mode == "section_boundary":
                    target._p.get_or_add_pPr().append(OxmlElement("w:sectPr"))
                path = self.root / ("audit-" + mode + ".docx")
                changed.save(path)
                failed = audit(path)
                self.assertNotEqual(0, failed.returncode, failed.stdout)
        self.assertEqual(original_source, source.read_bytes())
        self.assertEqual(original_map, map_path.read_bytes())

    def test_font_audit_reconciles_only_approved_scalar_cache_identity(self) -> None:
        from field_writeback import _result_text_nodes
        from finalize_docx import _font_scalar_cache_key
        from structure_map import candidate_structure_map

        nested = Document().add_paragraph()
        add_nested_formula(nested, "PAGE", "2")
        self.assertIsNone(_font_scalar_cache_key(nested, {"PAGE"}))

        baseline = self.root / "font-baseline.docx"
        document = Document()
        fonts = OxmlElement("w:rFonts")
        for attr, name in (("ascii", "Times New Roman"), ("eastAsia", "Songti"), ("cs", "Times New Roman")):
            fonts.set(qn("w:" + attr), name)
        document.styles["Normal"].element.get_or_add_rPr().append(fonts)
        document.add_paragraph("Before anchor")
        paragraph = document.add_paragraph("Approved body ")
        for kind in ("PAGE", "NUMPAGES", "PAGEREF anchor"):
            add_complex_field(paragraph, kind, "0")
        document.add_paragraph("After anchor")
        document.save(baseline)
        structure = candidate_structure_map(baseline)
        for entry in structure["paragraph_roles"]:
            if entry["locator"]["paragraph"] == 1:
                entry.update(approved=True, role="body_text", canonical_role="body")
        profile = {"rules": [{"id": "FMT-BODY-501", "status": "approved", "application": "automatic",
                    "selector": {"kind": "paragraph_role", "value": "body_text"},
                    "properties": {"font_name_ascii": "Times New Roman", "font_name_east_asia": "Songti", "font_name_complex_script": "Times New Roman"}}]}
        allowed = {"PAGE", "NUMPAGES", "PAGEREF"}
        refreshed = self.root / "font-refreshed.docx"

        def refresh(name, data):
            if name != "word/document.xml":
                return data
            root = etree.fromstring(data)
            for field in parse_fields(root):
                for node in _result_text_nodes(root, field):
                    node.text = "3"
                field.begin.attrib.pop(qn("w:dirty"), None)
            return etree.tostring(root)

        rewrite_package(baseline, refreshed, refresh)
        output = self.root / "font-output.docx"
        writeback = selective_field_result_writeback(baseline, refreshed, output, allowed_field_types=allowed)
        self.assertEqual("selective_verified", writeback["status"])
        original_bytes = output.read_bytes()
        self.assertEqual([], effective_font_failures(baseline, profile, structure))
        self.assertTrue(effective_font_failures(output, profile, structure))
        self.assertTrue(effective_font_failures(output, profile, structure, baseline_path=baseline))
        self.assertTrue(effective_font_failures(output, profile, structure, baseline_path=output, scalar_refresh_fields=allowed))
        self.assertTrue(effective_font_failures(output, profile, structure, baseline_path=baseline, scalar_refresh_fields={"PAGE"}))
        self.assertEqual([], effective_font_failures(output, profile, structure, baseline_path=baseline, scalar_refresh_fields=allowed))
        self.assertEqual(original_bytes, output.read_bytes())

        for mode in ("authored", "style", "instruction", "boundary", "duplicate", "font", "style_font", "unknown_field"):
            with self.subTest(mode=mode):
                changed = self.root / (mode + ".docx")

                def mutate(name, data):
                    if mode == "style_font" and name == "word/styles.xml":
                        root = etree.fromstring(data)
                        root.xpath("./w:style[@w:styleId='Normal']/w:rPr/w:rFonts", namespaces=NS)[0].set(qn("w:eastAsia"), "Wrong Font")
                        return etree.tostring(root)
                    if name != "word/document.xml":
                        return data
                    root = etree.fromstring(data)
                    target = root.xpath("./w:body/w:p", namespaces=NS)[1]
                    if mode == "authored":
                        target.xpath(".//w:t", namespaces=NS)[0].text = "Changed author text"
                    elif mode == "style":
                        ppr = etree.Element(qn("w:pPr"))
                        etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "Heading1")
                        target.insert(0, ppr)
                    elif mode in {"instruction", "unknown_field"}:
                        target.xpath(".//w:instrText", namespaces=NS)[0].text = " DATE " if mode == "unknown_field" else " NUMPAGES "
                    elif mode == "boundary":
                        end = target.xpath(".//w:fldChar[@w:fldCharType='end']", namespaces=NS)[-1]
                        end.getparent().remove(end)
                    elif mode == "duplicate":
                        target.addnext(copy.deepcopy(target))
                    elif mode == "font":
                        run = target.xpath("./w:r[w:t]", namespaces=NS)[0]
                        rpr = etree.Element(qn("w:rPr"))
                        etree.SubElement(rpr, qn("w:rFonts")).set(qn("w:eastAsia"), "Wrong Font")
                        run.insert(0, rpr)
                    return etree.tostring(root)

                rewrite_package(output, changed, mutate)
                failures = effective_font_failures(changed, profile, structure, baseline_path=baseline, scalar_refresh_fields=allowed)
                self.assertTrue(failures)
                if mode == "style_font":
                    self.assertTrue(any(item.get("actual") == "Wrong Font" and item.get("expected") == "Songti" for item in failures))

    def test_backend_serialization_noise_is_discarded(self) -> None:
        baseline = self.root / "baseline.docx"
        document = Document()
        document.add_paragraph("Authored content")
        add_complex_field(document.add_paragraph(), "PAGE", "1", dirty=False)
        update = OxmlElement("w:updateFields")
        update.set(qn("w:val"), "true")
        document.settings._element.append(update)
        document.save(baseline)
        refreshed = self.root / "refreshed.docx"

        def transform(name: str, data: bytes) -> bytes:
            if name != "word/document.xml":
                return data
            root = etree.fromstring(data)
            root.xpath(".//w:p[1]", namespaces=NS)[0].set(qn("w:rsidR"), "00ABCDEF")
            authored_run = root.xpath(".//w:p[1]/w:r[1]", namespaces=NS)[0]
            properties = OxmlElement("w:rPr")
            properties.append(OxmlElement("w:b"))
            authored_run.insert(0, properties)
            record = next(item for item in parse_fields(root) if item.field_type == "PAGE")
            elements = list(root.iter())
            start = elements.index(record.separate)
            end = elements.index(record.end)
            next(
                item
                for item in elements[start + 1 : end]
                if item.tag == qn("w:t")
            ).text = "2"
            return etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        rewrite_package(baseline, refreshed, transform)
        output = self.root / "output.docx"
        report = selective_field_result_writeback(baseline, refreshed, output)
        self.assertEqual("selective_verified", report["status"])
        self.assertEqual(["2"], field_values(output))
        with zipfile.ZipFile(output) as package:
            root = etree.fromstring(package.read("word/document.xml"))
            self.assertFalse(root.xpath(".//w:p[1]/w:r[1]/w:rPr/w:b", namespaces=NS))
            self.assertIsNone(root.xpath(".//w:p[1]", namespaces=NS)[0].get(qn("w:rsidR")))
            settings = etree.fromstring(package.read("word/settings.xml"))
            self.assertFalse(settings.xpath("./w:updateFields", namespaces=NS))

    def test_changed_instruction_or_authored_text_is_rejected(self) -> None:
        baseline = self.root / "baseline.docx"
        document = Document()
        document.add_paragraph("Protected authored text")
        add_complex_field(document.add_paragraph(), "PAGE", "1")
        document.save(baseline)

        for mode in ("instruction", "authored"):
            refreshed = self.root / f"{mode}.docx"

            def transform(name: str, data: bytes, selected: str = mode) -> bytes:
                if name != "word/document.xml":
                    return data
                root = etree.fromstring(data)
                if selected == "instruction":
                    root.xpath(".//w:instrText", namespaces=NS)[0].text = " NUMPAGES "
                else:
                    root.xpath(".//w:p[1]//w:t", namespaces=NS)[0].text = "Changed"
                return etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )

            rewrite_package(baseline, refreshed, transform)
            with self.assertRaises(FormatMonographError):
                selective_field_result_writeback(
                    baseline, refreshed, self.root / f"{mode}-output.docx"
                )

    def test_duplicate_scalar_fields_match_by_unique_authored_context(self) -> None:
        baseline = self.root / "duplicate.docx"
        document = Document()
        add_complex_field(document.add_paragraph("First "), "PAGE", "1", dirty=False)
        add_complex_field(document.add_paragraph("Second "), "PAGE", "1", dirty=False)
        document.save(baseline)
        refreshed = self.root / "duplicate-refreshed.docx"

        def transform(name: str, data: bytes) -> bytes:
            if name != "word/document.xml":
                return data
            root = etree.fromstring(data)
            elements = list(root.iter())
            for value, record in zip(("2", "3"), parse_fields(root)):
                start = elements.index(record.separate)
                end = elements.index(record.end)
                next(
                    item
                    for item in elements[start + 1 : end]
                    if item.tag == qn("w:t")
                ).text = value
            return etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        rewrite_package(baseline, refreshed, transform)
        output = self.root / "duplicate-output.docx"
        report = selective_field_result_writeback(baseline, refreshed, output)
        self.assertEqual(2, report["matched_fields"])
        self.assertEqual(["2", "3"], field_values(output))

    def test_ambiguous_duplicate_scalar_fields_are_rejected(self) -> None:
        baseline = self.root / "ambiguous.docx"
        document = Document()
        add_complex_field(document.add_paragraph(), "PAGE", "1")
        add_complex_field(document.add_paragraph(), "PAGE", "1")
        document.save(baseline)
        refreshed = self.root / "ambiguous-refreshed.docx"
        rewrite_package(baseline, refreshed, lambda _name, data: data)
        with self.assertRaisesRegex(FormatMonographError, "Duplicate fields"):
            selective_field_result_writeback(
                baseline,
                refreshed,
                self.root / "ambiguous-output.docx",
            )

    def test_simple_field_can_match_word_complex_field_serialization(self) -> None:
        baseline = self.root / "simple.docx"
        document = Document()
        add_simple_field(document.add_paragraph(), "PAGE", "1")
        document.save(baseline)
        refreshed = self.root / "complex.docx"

        def transform(name: str, data: bytes) -> bytes:
            if name != "word/document.xml":
                return data
            root = etree.fromstring(data)
            paragraph = root.xpath(".//w:p[w:fldSimple]", namespaces=NS)[0]
            field = paragraph.find(qn("w:fldSimple"))
            paragraph.remove(field)
            temporary = Document()
            temporary_paragraph = temporary.add_paragraph()
            add_complex_field(temporary_paragraph, "PAGE", "8", dirty=False)
            for child in list(temporary_paragraph._p):
                if child.tag != qn("w:pPr"):
                    paragraph.append(child)
            return etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        rewrite_package(baseline, refreshed, transform)
        output = self.root / "simple-output.docx"
        report = selective_field_result_writeback(baseline, refreshed, output)
        self.assertEqual(1, report["matched_fields"])
        with zipfile.ZipFile(output) as package:
            root = etree.fromstring(package.read("word/document.xml"))
            self.assertEqual(
                "8",
                "".join(root.xpath(".//w:fldSimple//w:t/text()", namespaces=NS)),
            )

    def test_header_footer_results_are_patched_without_importing_other_xml(self) -> None:
        baseline = self.root / "footer.docx"
        document = Document()
        document.add_paragraph("Body")
        add_complex_field(
            document.sections[0].footer.paragraphs[0],
            "PAGE",
            "1",
            dirty=False,
        )
        document.save(baseline)
        with zipfile.ZipFile(baseline) as package:
            footer_part = next(
                name for name in package.namelist() if name.startswith("word/footer")
            )
        refreshed = self.root / "footer-refreshed.docx"

        refreshed_footer = "word/footer9.xml"
        with zipfile.ZipFile(baseline) as source, zipfile.ZipFile(
            refreshed, "w", zipfile.ZIP_DEFLATED
        ) as target:
            for info in source.infolist():
                name = info.filename
                data = source.read(name)
                output_name = name
                if name == footer_part:
                    root = etree.fromstring(data)
                    record = parse_fields(root)[0]
                    elements = list(root.iter())
                    start = elements.index(record.separate)
                    end = elements.index(record.end)
                    next(
                        item
                        for item in elements[start + 1 : end]
                        if item.tag == qn("w:t")
                    ).text = "9"
                    root.set("backend-noise", "discard-me")
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                    output_name = refreshed_footer
                elif name == "word/_rels/document.xml.rels":
                    root = etree.fromstring(data)
                    for relationship in root:
                        if relationship.get("Target") == footer_part.removeprefix("word/"):
                            relationship.set("Target", "footer9.xml")
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                elif name == "[Content_Types].xml":
                    root = etree.fromstring(data)
                    for override in root:
                        if override.get("PartName") == f"/{footer_part}":
                            override.set("PartName", f"/{refreshed_footer}")
                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                target.writestr(output_name, data)
        output = self.root / "footer-output.docx"
        report = selective_field_result_writeback(baseline, refreshed, output)
        self.assertIn(footer_part, report["patched_parts"])
        self.assertIn(
            "header_footer_part_renumbering",
            report["discarded_backend_differences"],
        )
        self.assertEqual(["9"], field_values(output, footer_part))
        with zipfile.ZipFile(output) as package:
            root = etree.fromstring(package.read(footer_part))
            self.assertIsNone(root.get("backend-noise"))

    def test_empty_extra_header_footer_roles_are_discarded(self) -> None:
        baseline = self.root / "empty-role-baseline.docx"
        document = Document()
        document.add_paragraph("Body")
        add_complex_field(
            document.sections[0].footer.paragraphs[0],
            "PAGE",
            "1",
            dirty=False,
        )
        document.save(baseline)

        for attribute, kind, active in (
            ("header", "default", True),
            ("first_page_header", "first", False),
            ("even_page_header", "even", False),
        ):
            with self.subTest(kind=kind):
                refreshed = self.root / f"empty-{kind}.docx"
                candidate = Document(baseline)
                story = getattr(candidate.sections[0], attribute)
                story.is_linked_to_previous = False
                self.assertEqual(1, len(story.paragraphs))
                candidate.save(refreshed)

                with zipfile.ZipFile(baseline) as baseline_package, zipfile.ZipFile(
                    refreshed
                ) as refreshed_package:
                    baseline_roles = _story_roles(baseline_package)
                    refreshed_roles = _story_roles(refreshed_package)
                    document_root = etree.fromstring(
                        refreshed_package.read("word/document.xml")
                    )
                    settings_root = etree.fromstring(
                        refreshed_package.read("word/settings.xml")
                    )
                extra_roles = set(refreshed_roles) - set(baseline_roles)
                self.assertEqual({("header", 0, kind)}, extra_roles)
                title_page = document_root.find(".//" + qn("w:titlePg")) is not None
                even_and_odd = (
                    settings_root.find(qn("w:evenAndOddHeaders")) is not None
                )
                observed_active = (
                    kind == "default"
                    or (kind == "first" and title_page)
                    or (kind == "even" and even_and_odd)
                )
                self.assertEqual(active, observed_active)

                output = self.root / f"empty-{kind}-output.docx"
                report = selective_field_result_writeback(baseline, refreshed, output)
                self.assertEqual(1, report['candidate_story_audit']['discarded_candidate_role_count'])
                with zipfile.ZipFile(output) as package, zipfile.ZipFile(baseline) as original:
                    self.assertEqual(baseline_roles, _story_roles(package))
                    self.assertEqual(set(original.namelist()), set(package.namelist()))
                    self.assertEqual(original.read('word/styles.xml'), package.read('word/styles.xml'))

    def test_header_footer_activation_value_matrix(self) -> None:
        for tag in ("titlePg", "evenAndOddHeaders"):
            for before, after, accepted in (
                ("false", "true", False), ("true", "false", False),
                ("0", "on", False), ("off", None, True),
                (None, "false", True), ("true", "1", True),
                ("on", "true", True), ("false", "off", True),
                (None, "invalid", False), (None, "duplicate", False),
            ):
                with self.subTest(tag=tag, before=before, after=after):
                    paths = []
                    for index, value in enumerate((before, after)):
                        doc = Document()
                        doc.add_paragraph("Body")
                        add_complex_field(doc.sections[0].footer.paragraphs[0], "PAGE", "1", dirty=False)
                        parent = doc.sections[0]._sectPr if tag == "titlePg" else doc.settings.element
                        if value is not None:
                            element = OxmlElement("w:" + tag)
                            element.set(qn("w:val"), "true" if value == "duplicate" else value)
                            parent.append(element)
                            if value == "duplicate":
                                parent.append(OxmlElement("w:" + tag))
                        path = self.root / f"activation-{index}.docx"
                        doc.save(path)
                        paths.append(path)
                    output = self.root / "activation-output.docx"
                    if accepted:
                        selective_field_result_writeback(*paths, output)
                    else:
                        with self.assertRaises(FormatMonographError):
                            selective_field_result_writeback(*paths, output)

    def test_extra_role_cannot_share_original_candidate_source(self) -> None:
        baseline = self.root / "shared-baseline.docx"
        candidate = self.root / "shared-candidate.docx"
        doc = Document()
        doc.add_paragraph("Body")
        add_complex_field(doc.sections[0].footer.paragraphs[0], "PAGE", "1", dirty=False)
        doc.save(baseline)
        reference = doc.sections[0]._sectPr.find(qn("w:footerReference"))
        extra = OxmlElement("w:footerReference")
        extra.set(qn("w:type"), "first")
        extra.set(qn("r:id"), reference.get(qn("r:id")))
        doc.sections[0]._sectPr.append(extra)
        doc.save(candidate)
        with self.assertRaisesRegex(FormatMonographError, "shares a baseline"):
            selective_field_result_writeback(baseline, candidate, self.root / "shared-output.docx")

    def test_shared_baseline_footer_split_into_conflicting_sources_is_rejected(self) -> None:
        from copy import deepcopy

        baseline = self.root / "split-source-baseline.docx"
        candidate = self.root / "split-source-candidate.docx"
        output = self.root / "split-source-output.docx"
        document = Document()
        document.add_paragraph("First section")
        add_complex_field(document.sections[0].footer.paragraphs[0], "PAGE", "1", dirty=False)
        document.add_section(WD_SECTION.NEW_PAGE)
        document.add_paragraph("Second section")
        self.assertTrue(document.sections[1].footer.is_linked_to_previous)
        document.save(baseline)

        refreshed = Document(baseline)
        original_footer = refreshed.sections[0].footer._element
        split_footer = refreshed.sections[1].footer
        split_footer.is_linked_to_previous = False
        for child in list(split_footer._element):
            split_footer._element.remove(child)
        for child in original_footer:
            split_footer._element.append(deepcopy(child))
        for text in split_footer._element.iter(qn("w:t")):
            text.text = "9"
        refreshed.save(candidate)
        with zipfile.ZipFile(baseline) as original, zipfile.ZipFile(candidate) as changed:
            self.assertEqual(1, len(set(_story_roles(original).values())))
            roles = _story_roles(changed)
            self.assertEqual(2, len(set(roles.values())))
            for section, value in ((0, "1"), (1, "9")):
                root = etree.fromstring(changed.read(roles[("footer", section, "default")]))
                self.assertEqual([value], root.xpath(".//w:t/text()", namespaces=NS))
        with self.assertRaisesRegex(FormatMonographError, "Ambiguous baseline header/footer field source"):
            selective_field_result_writeback(baseline, candidate, output)
        self.assertFalse(output.exists())

    def test_empty_extra_role_with_localized_paragraph_style_is_discarded(self) -> None:
        baseline = self.root / "localized-empty-style-baseline.docx"
        style_name = "本地空页眉样式"
        document = Document()
        document.add_paragraph("Body")
        document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        add_complex_field(
            document.sections[0].footer.paragraphs[0],
            "PAGE",
            "1",
            dirty=False,
        )
        document.save(baseline)

        refreshed = self.root / "localized-empty-style-refreshed.docx"
        candidate = Document(baseline)
        header = candidate.sections[0].header
        header.is_linked_to_previous = False
        localized_style = candidate.styles[style_name]
        self.assertTrue(any(ord(character) > 127 for character in localized_style.style_id))
        header.paragraphs[0].style = localized_style
        candidate.save(refreshed)

        output = self.root / "localized-empty-style-output.docx"
        selective_field_result_writeback(baseline, refreshed, output)
        with zipfile.ZipFile(output) as package, zipfile.ZipFile(baseline) as original:
            self.assertEqual(_story_roles(original), _story_roles(package))

    def test_discarded_extra_role_does_not_resolve_style_definition(self) -> None:
        style_name = "基准空页眉样式"

        def make_pair(label: str, *, style_only_in_refreshed: bool = False):
            baseline = self.root / f"style-{label}-baseline.docx"
            document = Document()
            document.add_paragraph("Body")
            if not style_only_in_refreshed:
                document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            add_complex_field(
                document.sections[0].footer.paragraphs[0],
                "PAGE",
                "1",
                dirty=False,
            )
            document.save(baseline)
            refreshed = self.root / f"style-{label}-refreshed.docx"
            candidate = Document(baseline)
            if style_only_in_refreshed:
                candidate.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
            header = candidate.sections[0].header
            header.is_linked_to_previous = False
            header.paragraphs[0].style = candidate.styles[style_name]
            candidate.save(refreshed)
            return baseline, refreshed, candidate.styles[style_name].style_id

        def rewrite_styles(path: Path, label: str, mode: str, style_id: str) -> Path:
            output = self.root / f"style-{label}-{mode}.docx"
            with zipfile.ZipFile(path) as source, zipfile.ZipFile(
                output, "w", zipfile.ZIP_DEFLATED
            ) as target:
                for info in source.infolist():
                    if mode == "missing_styles" and info.filename == "word/styles.xml":
                        continue
                    data = source.read(info.filename)
                    if info.filename == "word/styles.xml":
                        root = etree.fromstring(data)
                        matches = [
                            style
                            for style in root.xpath("./w:style", namespaces=NS)
                            if style.get(qn("w:styleId")) == style_id
                        ]
                        self.assertEqual(1, len(matches))
                        style = matches[0]
                        if mode == "changed_definition":
                            marker = OxmlElement("w:semiHidden")
                            style.append(marker)
                        elif mode == "wrong_type":
                            style.set(qn("w:type"), "character")
                        elif mode == "duplicate":
                            root.append(etree.fromstring(etree.tostring(style)))
                        data = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )
                    target.writestr(info, data)
            return output

        baseline, refreshed, _style_id = make_pair(
            "refreshed-only", style_only_in_refreshed=True
        )
        cases = [("style_only_in_refreshed", baseline, refreshed)]

        baseline, refreshed, style_id = make_pair("baseline-missing")
        cases.append(
            (
                "baseline_styles_missing",
                rewrite_styles(baseline, "baseline-missing", "missing_styles", style_id),
                refreshed,
            )
        )

        baseline, refreshed, style_id = make_pair("definition-changed")
        cases.append(
            (
                "definition_changed",
                baseline,
                rewrite_styles(
                    refreshed, "definition-changed", "changed_definition", style_id
                ),
            )
        )

        baseline, refreshed, style_id = make_pair("wrong-type")
        cases.append(
            (
                "type_wrong",
                rewrite_styles(baseline, "wrong-type-baseline", "wrong_type", style_id),
                rewrite_styles(refreshed, "wrong-type-refreshed", "wrong_type", style_id),
            )
        )

        baseline, refreshed, style_id = make_pair("duplicate-baseline")
        cases.append(
            (
                "baseline_duplicate",
                rewrite_styles(
                    baseline, "duplicate-baseline", "duplicate", style_id
                ),
                refreshed,
            )
        )

        baseline, refreshed, style_id = make_pair("duplicate-refreshed")
        cases.append(
            (
                "refreshed_duplicate",
                baseline,
                rewrite_styles(
                    refreshed, "duplicate-refreshed", "duplicate", style_id
                ),
            )
        )

        for label, baseline, refreshed in cases:
            with self.subTest(label=label), zipfile.ZipFile(
                baseline
            ) as baseline_package, zipfile.ZipFile(refreshed) as refreshed_package:
                sources, _ = _semantic_part_sources(baseline_package, refreshed_package)
                self.assertEqual(set(_story_roles(baseline_package).values()), set(sources))

    def test_nonempty_or_invalid_extra_header_footer_roles_are_rejected(self) -> None:
        baseline = self.root / "invalid-extra-role-baseline.docx"
        document = Document()
        document.add_paragraph("Body")
        add_complex_field(
            document.sections[0].footer.paragraphs[0],
            "PAGE",
            "1",
            dirty=False,
        )
        document.save(baseline)

        for mode in (
            "field",
            "text",
            "table",
            "drawing",
            "symbol",
            "break",
            "layout_attribute",
            "layout_subtree",
            "relationship",
            "relationship_part",
            "malformed",
            "missing",
        ):
            with self.subTest(mode=mode):
                materialized = self.root / f"extra-{mode}-materialized.docx"
                candidate = Document(baseline)
                header = candidate.sections[0].header
                header.is_linked_to_previous = False
                self.assertEqual(1, len(header.paragraphs))
                candidate.save(materialized)
                with zipfile.ZipFile(baseline) as baseline_package, zipfile.ZipFile(
                    materialized
                ) as refreshed_package:
                    extra_roles = set(_story_roles(refreshed_package)) - set(
                        _story_roles(baseline_package)
                    )
                    self.assertEqual(1, len(extra_roles))
                    extra_part = _story_roles(refreshed_package)[extra_roles.pop()]

                refreshed = self.root / f"extra-{mode}.docx"
                with zipfile.ZipFile(materialized) as source, zipfile.ZipFile(
                    refreshed, "w", zipfile.ZIP_DEFLATED
                ) as target:
                    for info in source.infolist():
                        if mode == "missing" and info.filename == extra_part:
                            continue
                        data = source.read(info.filename)
                        if info.filename == extra_part:
                            if mode == "malformed":
                                data = b"<malformed"
                            else:
                                root = etree.fromstring(data)
                                if mode == "field":
                                    field = OxmlElement("w:fldSimple")
                                    field.set(qn("w:instr"), " PAGE ")
                                    root.append(field)
                                elif mode == "text":
                                    text = OxmlElement("w:t")
                                    text.text = "unauthorized"
                                    root.append(text)
                                elif mode == "table":
                                    root.append(OxmlElement("w:tbl"))
                                elif mode == "drawing":
                                    root.append(OxmlElement("w:drawing"))
                                elif mode == "symbol":
                                    root.append(OxmlElement("w:sym"))
                                elif mode == "break":
                                    root.append(OxmlElement("w:br"))
                                elif mode == "layout_attribute":
                                    root.set(qn("w:rsidR"), "00000001")
                                elif mode == "layout_subtree":
                                    properties = root.find(".//" + qn("w:pPr"))
                                    assert properties is not None
                                    properties.append(OxmlElement("w:spacing"))
                                elif mode == "relationship":
                                    root.set(qn("r:id"), "rIdSynthetic")
                                data = etree.tostring(
                                    root,
                                    xml_declaration=True,
                                    encoding="UTF-8",
                                    standalone=True,
                                )
                        target.writestr(info, data)
                    if mode == "relationship_part":
                        part_path = Path(extra_part)
                        relationship_part = str(
                            part_path.parent
                            / "_rels"
                            / f"{part_path.name}.rels"
                        )
                        target.writestr(
                            relationship_part,
                            b'<?xml version="1.0" encoding="UTF-8"?>'
                            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
                        )

                output = self.root / f"extra-{mode}-output.docx"
                if mode == "malformed":
                    with zipfile.ZipFile(baseline) as baseline_package, zipfile.ZipFile(
                        refreshed
                    ) as refreshed_package, self.assertRaises(FormatMonographError):
                        _semantic_part_sources(baseline_package, refreshed_package)
                else:
                    with self.assertRaises(FormatMonographError):
                        selective_field_result_writeback(baseline, refreshed, output)
                self.assertFalse(output.exists())

    def test_missing_baseline_header_footer_role_is_rejected(self) -> None:
        baseline = self.root / "missing-role-baseline.docx"
        document = Document()
        document.add_paragraph("Body")
        add_complex_field(
            document.sections[0].footer.paragraphs[0],
            "PAGE",
            "1",
            dirty=False,
        )
        document.save(baseline)
        refreshed = self.root / "missing-role-refreshed.docx"

        def remove_footer_role(name: str, data: bytes) -> bytes:
            if name != "word/document.xml":
                return data
            root = etree.fromstring(data)
            reference = root.find(".//" + qn("w:footerReference"))
            assert reference is not None
            reference.getparent().remove(reference)
            return etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        rewrite_package(baseline, refreshed, remove_footer_role)
        output = self.root / "missing-role-output.docx"
        with self.assertRaisesRegex(
            FormatMonographError, "changed the effective header or footer role set"
        ):
            selective_field_result_writeback(baseline, refreshed, output)
        self.assertFalse(output.exists())

    def test_unapproved_dirty_seq_is_reported_and_not_updated(self) -> None:
        baseline = self.root / "seq.docx"
        document = Document()
        add_complex_field(document.add_paragraph(), "SEQ Figure", "1")
        document.save(baseline)
        refreshed = self.root / "seq-refreshed.docx"
        rewrite_package(baseline, refreshed, lambda _name, data: data)
        output = self.root / "seq-output.docx"
        report = selective_field_result_writeback(baseline, refreshed, output)
        self.assertEqual(["SEQ"], report["unapproved_field_types"])
        self.assertEqual(1, report["unapproved_dirty_fields"])
        self.assertEqual(["1"], field_values(output))

    def test_global_update_on_open_is_removed_with_clean_unapproved_fields(self) -> None:
        baseline = self.root / "mixed-fields.docx"
        document = Document()
        add_complex_field(document.add_paragraph(), "PAGE", "1", dirty=False)
        add_complex_field(document.add_paragraph("Figure "), "SEQ Figure", "1", dirty=False)
        update = OxmlElement("w:updateFields")
        update.set(qn("w:val"), "true")
        document.settings._element.append(update)
        document.save(baseline)
        refreshed = self.root / "mixed-fields-refreshed.docx"
        rewrite_package(baseline, refreshed, lambda _name, data: data)
        output = self.root / "mixed-fields-output.docx"
        report = selective_field_result_writeback(baseline, refreshed, output)
        self.assertEqual(0, report["unapproved_dirty_fields"])
        self.assertEqual(["SEQ"], report["unapproved_field_types"])
        with zipfile.ZipFile(output) as package:
            settings = etree.fromstring(package.read("word/settings.xml"))
        self.assertFalse(settings.xpath("./w:updateFields", namespaces=NS))

    def test_dirty_approved_result_is_rejected(self) -> None:
        baseline = self.root / "dirty-page.docx"
        document = Document()
        add_complex_field(document.add_paragraph(), "PAGE", "1")
        document.save(baseline)
        refreshed = self.root / "dirty-page-refreshed.docx"
        rewrite_package(baseline, refreshed, lambda _name, data: data)
        with self.assertRaisesRegex(FormatMonographError, "left field PAGE dirty"):
            selective_field_result_writeback(
                baseline,
                refreshed,
                self.root / "dirty-page-output.docx",
            )

    def test_arbitrary_formula_is_rejected_even_when_formula_type_is_allowed(self) -> None:
        baseline = self.root / "formula.docx"
        document = Document()
        add_complex_field(document.add_paragraph(), "= 2 + 2", "4")
        document.save(baseline)
        refreshed = self.root / "formula-refreshed.docx"
        rewrite_package(baseline, refreshed, lambda _name, data: data)

        with self.assertRaisesRegex(
            FormatMonographError,
            "Only the exact core-generated PAGE-minus-one display formula",
        ):
            selective_field_result_writeback(
                baseline,
                refreshed,
                self.root / "formula-output.docx",
                allowed_field_types={"="},
            )

    def test_non_page_nested_formula_and_non_numeric_result_are_rejected(self) -> None:
        for nested, value in (("NUMPAGES", "1"), ("PAGE", "FORGED")):
            baseline = self.root / f"nested-{nested}-{value}.docx"
            document = Document()
            add_nested_formula(document.add_paragraph(), nested, value)
            document.save(baseline)
            refreshed = self.root / f"nested-{nested}-{value}-refreshed.docx"
            rewrite_package(baseline, refreshed, lambda _name, data: data)
            with self.assertRaises(FormatMonographError):
                selective_field_result_writeback(
                    baseline,
                    refreshed,
                    self.root / f"nested-{nested}-{value}-output.docx",
                    allowed_field_types={"=", "PAGE", "NUMPAGES"},
                )

    def test_external_verify_requires_no_save_contract_and_matching_page_count(self) -> None:
        source = self.root / "verify.docx"
        Document().save(source)
        helper = self.root / "verify_backend.py"
        helper.write_text(
            "import json, pathlib, sys\n"
            "request=json.load(sys.stdin)\n"
            "assert request['operation']=='verify_only'\n"
            "assert request['target_software']=='microsoft_word'\n"
            "pathlib.Path(request['pdf_output_path']).write_bytes(b'%PDF-1.4\\n%%EOF')\n"
            "print(json.dumps({'protocol_version':'1.1','status':'success','operation':'verify_only',"
            "'backend':'test_word','software':'Microsoft Word','repaginated':True,"
            "'saved':False,'read_only_verified':True,'pdf_exported':True,"
            "'page_count':7}))\n",
            encoding="utf-8",
        )
        pdf = self.root / "verify.pdf"
        response = external_verify(
            source,
            json.dumps([sys.executable, str(helper)]),
            self.root / "profile.json",
            self.root / "map.json",
            pdf,
            "Microsoft Word 2021",
            expected_page_count=7,
        )
        self.assertTrue(response["read_only_verified"])
        with self.assertRaises(FormatMonographError):
            external_verify(
                source,
                json.dumps([sys.executable, str(helper)]),
                self.root / "profile.json",
                self.root / "map.json",
                self.root / "mismatch.pdf",
                "Microsoft Word 2021",
                expected_page_count=8,
            )
        helper.write_text(
            "import json, pathlib, sys\n"
            "request=json.load(sys.stdin)\n"
            "pathlib.Path(request['pdf_output_path']).write_bytes(b'%PDF-1.4\\n%%EOF')\n"
            "print(json.dumps({'protocol_version':'1.1','status':'success','operation':'verify_only',"
            "'backend':'test_word','software':'Microsoft Word','repaginated':True,"
            "'saved':False,'read_only_verified':True,'pdf_exported':True}))\n",
            encoding="utf-8",
        )
        with self.assertRaisesRegex(FormatMonographError, "valid page count"):
            external_verify(
                source,
                json.dumps([sys.executable, str(helper)]),
                self.root / "profile.json",
                self.root / "map.json",
                self.root / "missing-count.pdf",
                "Microsoft Word 2021",
                expected_page_count=7,
            )

    def test_layout_measurement_and_core_spacer_removal_are_separate(self) -> None:
        source = self.root / "spacers.docx"
        document = Document()
        style = document.styles.add_style(
            "Monograph Figure Table Spacer", WD_STYLE_TYPE.PARAGRAPH
        )
        document.add_paragraph("Body before")
        document.add_paragraph(style=style)
        document.add_paragraph("Body after")
        document.save(source)
        helper = self.root / "measure_backend.py"
        helper.write_text(
            "import json, sys\n"
            "request=json.load(sys.stdin)\n"
            "assert request['operation']=='measure_layout'\n"
            "assert request['target_software']=='microsoft_word'\n"
            "print(json.dumps({'protocol_version':'1.1','status':'success','operation':'measure_layout',"
            "'backend':'test_word','software':'Microsoft Word','repaginated':True,"
            "'saved':False,'read_only_verified':True,'structural_changes_applied':0,"
            "'page_count':3,'sections':[{'section_index':0}],"
            "'page_boundary_spacer_ordinals':[0]}))\n",
            encoding="utf-8",
        )
        measured = external_measure(
            source,
            json.dumps([sys.executable, str(helper)]),
            self.root / "profile.json",
            self.root / "map.json",
            "Microsoft Word 2021",
        )
        self.assertEqual([0], measured["page_boundary_spacer_ordinals"])
        output = self.root / "spacers-normalized.docx"
        self.assertEqual(
            2,
            apply_measured_layout_adjustments(
                source,
                output,
                [0],
                {0: "evenPage"},
            ),
        )
        self.assertEqual(
            ["Body before", "Body after"],
            [paragraph.text for paragraph in Document(output).paragraphs],
        )
        section_type = Document(output).sections[0]._sectPr.find(qn("w:type"))
        self.assertEqual("evenPage", section_type.get(qn("w:val")))

        spacer_only = self.root / "spacer-only.docx"
        self.assertEqual(1, remove_measured_block_spacers(source, spacer_only, [0]))

        offset_output = self.root / "page-offset.docx"
        self.assertEqual(2, apply_page_display_offsets(source, offset_output, {0: 1}))
        with zipfile.ZipFile(offset_output) as package:
            field_types = []
            for name in package.namelist():
                if not name.startswith("word/footer"):
                    continue
                field_types.extend(
                    record.field_type
                    for record in parse_fields(etree.fromstring(package.read(name)))
                )
        self.assertEqual(2, field_types.count("="))
        self.assertEqual(2, field_types.count("PAGE"))

    def test_page_offset_footer_is_isolated_from_adjacent_sections(self) -> None:
        source = self.root / "shared-footers.docx"
        document = Document()
        document.settings.odd_and_even_pages_header_footer = True
        document.add_paragraph("Title")
        for footer in (
            document.sections[0].footer,
            document.sections[0].even_page_footer,
        ):
            add_complex_field(footer.paragraphs[0], "PAGE", "1", dirty=False)
        document.add_section(WD_SECTION.NEW_PAGE)
        document.add_paragraph("Contents")
        document.add_section(WD_SECTION.NEW_PAGE)
        document.add_paragraph("Body")
        first = document.sections[0]
        for section in document.sections[1:]:
            for footer_type in (
                WD_HEADER_FOOTER.PRIMARY,
                WD_HEADER_FOOTER.EVEN_PAGE,
            ):
                if section._sectPr.get_footerReference(footer_type) is not None:
                    section._sectPr.remove_footerReference(footer_type)
                reference = first._sectPr.get_footerReference(footer_type)
                section._sectPr.add_footerReference(footer_type, reference.rId)
        document.save(source)

        output = self.root / "isolated-footers.docx"
        self.assertEqual(2, apply_page_display_offsets(source, output, {1: 1}))
        result = Document(output)
        for footer_name in ("footer", "even_page_footer"):
            field_types = []
            for section in result.sections:
                footer = getattr(section, footer_name)
                root = etree.fromstring(footer._element.xml.encode("utf-8"))
                field_types.append([record.field_type for record in parse_fields(root)])
            self.assertEqual([["PAGE"], ["=", "PAGE"], ["PAGE"]], field_types)


if __name__ == "__main__":
    unittest.main()
