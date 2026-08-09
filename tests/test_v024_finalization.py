from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "format-monograph"
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

from _common import NS, field_cache_inventory  # noqa: E402
from finalize_docx import field_contract_preserved  # noqa: E402
from structure_map import (  # noqa: E402
    candidate_structure_map,
    structure_content_fingerprint,
    text_sha256,
)
from test_v11_execution import approved_v11_profile  # noqa: E402


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def run_element(name: str, text: str | None = None) -> etree._Element:
    run = etree.Element(qn("w:r"))
    child = etree.SubElement(run, qn(name))
    if text is not None:
        child.text = text
    return run


def field_char(kind: str) -> etree._Element:
    run = etree.Element(qn("w:r"))
    child = etree.SubElement(run, qn("w:fldChar"))
    child.set(qn("w:fldCharType"), kind)
    return run


def set_paragraph_style(paragraph: etree._Element, style: str) -> None:
    p_pr = paragraph.find(qn("w:pPr"))
    if p_pr is None:
        p_pr = etree.Element(qn("w:pPr"))
        paragraph.insert(0, p_pr)
    p_style = p_pr.find(qn("w:pStyle"))
    if p_style is None:
        p_style = etree.SubElement(p_pr, qn("w:pStyle"))
    p_style.set(qn("w:val"), style)


def expand_toc_cache(source: Path, output: Path) -> None:
    temp = output.with_suffix(".tmp")
    with zipfile.ZipFile(source) as package, zipfile.ZipFile(
        temp, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in package.infolist():
            data = package.read(info.filename)
            if info.filename == "word/document.xml":
                root = etree.fromstring(data)
                body = root.xpath("/w:document/w:body", namespaces=NS)[0]
                toc = body.xpath(
                    "./w:p[.//w:instrText[contains(translate(., 'toc', 'TOC'), 'TOC')]]",
                    namespaces=NS,
                )[0]
                for child in list(toc):
                    if child.tag != qn("w:pPr"):
                        toc.remove(child)
                set_paragraph_style(toc, "TOC1")
                toc.extend(
                    [
                        field_char("begin"),
                        run_element("w:instrText", ' TOC \\o "1-3" \\h \\z \\u '),
                        field_char("separate"),
                        run_element("w:t", "Generated chapter entry"),
                    ]
                )
                second = etree.Element(qn("w:p"))
                set_paragraph_style(second, "TOC2")
                second.extend(
                    [
                        run_element("w:t", "Generated section entry"),
                        field_char("end"),
                    ]
                )
                body.insert(body.index(toc) + 1, second)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            elif info.filename == "word/settings.xml":
                root = etree.fromstring(data)
                for update in root.xpath("./w:updateFields", namespaces=NS):
                    root.remove(update)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            target.writestr(info, data)
    os.replace(temp, output)


class V024FinalizationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "synthetic-finalization.docx"
        image = self.root / "pixel.png"
        image.write_bytes(PNG_1X1)

        document = Document()
        document.add_paragraph("Synthetic book title")
        document.add_paragraph("4.1 Static contents entry", style="Heading 1")
        document.add_paragraph("4.1.1 Static contents child", style="Heading 2")
        document.add_paragraph("第4章 Synthetic chapter")
        document.add_paragraph("4.1 Synthetic section")
        polluted = document.add_paragraph("Synthetic body with inherited outline", style="Heading 2")
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), "1")
        polluted._p.get_or_add_pPr().append(outline)
        figure = document.add_paragraph()
        figure.add_run().add_picture(str(image))
        document.add_paragraph("图 4.1-1 Synthetic figure")
        document.add_paragraph("表 4.1 Synthetic standalone table")
        first = document.add_table(rows=2, cols=2)
        first.cell(0, 0).text = "Header A"
        first.cell(0, 1).text = "Header B"
        first.cell(1, 0).text = "Value A"
        first.cell(1, 1).text = "Value B"
        repeated = document.add_table(rows=3, cols=2)
        repeated.cell(0, 0).merge(repeated.cell(0, 1)).text = (
            "表 4.2 Synthetic repeated title"
        )
        repeated.cell(1, 0).text = "Header C"
        repeated.cell(1, 1).text = "Header D"
        repeated.cell(2, 0).text = "Value C"
        repeated.cell(2, 1).text = "Value D"
        layout = document.add_table(rows=1, cols=1)
        layout.cell(0, 0).text = "Synthetic image-layout container"
        document.save(self.source)

        structure = candidate_structure_map(self.source)
        structure["status"] = "approved"
        structure["toc_ranges"] = [
            {
                "start_paragraph": 1,
                "end_paragraph": 2,
                "paragraph_sha256": [
                    text_sha256(document.paragraphs[1].text),
                    text_sha256(document.paragraphs[2].text),
                ],
                "levels": 3,
                "approved": True,
            }
        ]
        for entry in structure["headings"]:
            entry["approved"] = entry["paragraph"] in {3, 4}
        for entry in structure["paragraph_roles"]:
            locator = entry["locator"]
            if locator["kind"] == "body_paragraph" and locator["paragraph"] == 5:
                entry.update({"role": "body", "approved": True})
            elif entry["role"] in {
                "heading_1",
                "heading_2",
                "figure_caption",
                "table_caption",
            } and not (
                locator["kind"] == "body_paragraph"
                and locator["paragraph"] in {1, 2}
            ):
                entry["approved"] = True
        for entry in structure["captions"]:
            entry.update({"approved": True, "action": "style_only"})
        structure["tables"][0].update(
            {
                "kind": "data",
                "approved": True,
                "repeat_header_rows": [0],
                "prevent_normal_row_split": True,
            }
        )
        structure["tables"][1].update(
            {
                "kind": "data",
                "approved": True,
                "caption_row": 0,
                "header_rows": [1],
                "repeat_header_rows": [1],
                "repeat_caption_with_header": True,
                "prevent_normal_row_split": True,
            }
        )
        structure["tables"][2].update(
            {
                "kind": "layout",
                "approved": True,
                "pagination_only": True,
                "keep_rows_together": True,
            }
        )
        for group in structure["pagination_groups"]:
            group["approved"] = True
        self.structure = self.root / "structure-map.json"
        self.structure.write_text(
            json.dumps(structure, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        profile = approved_v11_profile()
        profile["rules"] = [
            rule
            for rule in profile["rules"]
            if rule["selector"]["kind"] == "paragraph_role"
        ]
        profile["rules"].append(
            {
                "id": "FMT-CAP-924",
                "category": "caption",
                "selector": {"kind": "caption_role", "value": "all"},
                "properties": {"font_size_pt": 9, "alignment": "center"},
                "source_ids": ["SRC-901"],
                "evidence_summary": "Synthetic caption formatting.",
                "confidence": "high",
                "status": "approved",
                "application": "automatic",
            }
        )
        self.profile = self.root / "profile.json"
        self.profile.write_text(
            json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        self.output_dir = self.root / "out"

    def tearDown(self) -> None:
        self.temp.cleanup()

    def run_script(self, name: str, *args: object) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(SCRIPTS / name), *(str(value) for value in args)],
            cwd=SKILL,
            capture_output=True,
            text=True,
            check=False,
        )

    def apply(self) -> Path:
        result = self.run_script(
            "apply_profile.py",
            self.source,
            "--profile",
            self.profile,
            "--structure-map",
            self.structure,
            "--output-dir",
            self.output_dir,
            "--allow-missing-fonts",
        )
        self.assertEqual(0, result.returncode, result.stdout + result.stderr)
        return self.output_dir / "synthetic-finalization-formatted.docx"

    def test_v13_applies_toc_outline_and_pagination_patches(self) -> None:
        formatted = self.apply()
        document = Document(formatted)
        self.assertEqual("Synthetic book title", document.paragraphs[0].text)
        self.assertTrue(
            document.paragraphs[1]._p.xpath(
                ".//w:instrText[contains(translate(., 'toc', 'TOC'), 'TOC')]"
            )
        )
        self.assertEqual("第4章 Synthetic chapter", document.paragraphs[2].text)
        body = next(
            p for p in document.paragraphs if p.text == "Synthetic body with inherited outline"
        )
        self.assertEqual("Normal", body.style.name)
        self.assertIsNone(body._p.get_or_add_pPr().find(qn("w:outlineLvl")))
        figure = next(p for p in document.paragraphs if p._p.xpath(".//w:drawing"))
        self.assertIsNotNone(figure._p.get_or_add_pPr().find(qn("w:keepNext")))
        table_caption = next(
            p for p in document.paragraphs if p.text == "表 4.1 Synthetic standalone table"
        )
        self.assertIsNotNone(
            table_caption._p.get_or_add_pPr().find(qn("w:keepNext"))
        )
        for row_index in (0, 1):
            self.assertIsNotNone(
                document.tables[1].rows[row_index]
                ._tr.get_or_add_trPr()
                .find(qn("w:tblHeader"))
            )
        self.assertIsNotNone(
            document.tables[2].rows[0]._tr.get_or_add_trPr().find(qn("w:cantSplit"))
        )

    def test_expanded_toc_keeps_stable_audit_and_single_main_toc(self) -> None:
        formatted = self.apply()
        expanded = self.root / "expanded-toc.docx"
        expand_toc_cache(formatted, expanded)
        field_cache = field_cache_inventory(expanded)
        self.assertEqual("refreshed", field_cache["status"])
        self.assertEqual(1, field_cache["main_toc_fields"])
        self.assertEqual(
            structure_content_fingerprint(self.source, json.loads(self.structure.read_text(encoding="utf-8"))),
            structure_content_fingerprint(expanded, json.loads(self.structure.read_text(encoding="utf-8"))),
        )
        audited = self.run_script(
            "audit_docx.py",
            self.source,
            expanded,
            "--profile",
            self.profile,
            "--structure-map",
            self.structure,
        )
        self.assertEqual(0, audited.returncode, audited.stdout + audited.stderr)

        damaged = self.root / "damaged.docx"
        document = Document(expanded)
        body = next(
            p for p in document.paragraphs if p.text == "Synthetic body with inherited outline"
        )
        body.add_run(" changed")
        document.save(damaged)
        failed = self.run_script(
            "audit_docx.py",
            self.source,
            damaged,
            "--profile",
            self.profile,
            "--structure-map",
            self.structure,
        )
        self.assertEqual(1, failed.returncode)

    def test_deferred_finalization_requires_explicit_qa(self) -> None:
        formatted = self.apply()
        denied = self.run_script(
            "finalize_docx.py",
            formatted,
            "--source",
            self.source,
            "--profile",
            self.profile,
            "--structure-map",
            self.structure,
            "--output",
            self.root / "denied.docx",
            "--field-updater",
            "deferred",
        )
        self.assertEqual(1, denied.returncode)
        self.assertIn("--approve-deferred", denied.stderr)

        finalized = self.root / "finalized.docx"
        status = self.root / "finalization.json"
        allowed = self.run_script(
            "finalize_docx.py",
            formatted,
            "--source",
            self.source,
            "--profile",
            self.profile,
            "--structure-map",
            self.structure,
            "--output",
            finalized,
            "--status-output",
            status,
            "--field-updater",
            "deferred",
            "--approve-deferred",
        )
        self.assertEqual(0, allowed.returncode, allowed.stdout + allowed.stderr)
        payload = json.loads(status.read_text(encoding="utf-8"))
        self.assertEqual("deferred", payload["delivery_field_status"])
        self.assertEqual("pass", payload["content_integrity"])
        self.assertEqual("pass", payload["protected_object_integrity"])

    def test_field_contract_rejects_removed_editable_fields(self) -> None:
        before = {
            "main_toc_fields": 1,
            "field_types": {"TOC": 1, "PAGE": 2},
        }
        self.assertTrue(field_contract_preserved(before, before))
        self.assertFalse(
            field_contract_preserved(
                before,
                {"main_toc_fields": 0, "field_types": {"PAGE": 2}},
            )
        )
        self.assertFalse(
            field_contract_preserved(
                before,
                {"main_toc_fields": 1, "field_types": {"TOC": 1, "PAGE": 1}},
            )
        )


if __name__ == "__main__":
    unittest.main()
