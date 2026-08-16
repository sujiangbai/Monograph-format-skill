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
from docx.shared import Inches
from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "format-monograph"
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

from _common import NS  # noqa: E402
from finalize_docx import controlled_field_result_writeback  # noqa: E402
from run_monograph import has_target_layout_evidence, validate_visual_manifest  # noqa: E402
from structure_map import (  # noqa: E402
    candidate_structure_map,
    load_structure_map,
)


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


def add_page_field(paragraph, value: str = "1") -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    run._r.append(begin)
    instruction_run = paragraph.add_run()
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    instruction_run._r.append(instruction)
    separate_run = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    separate_run._r.append(separate)
    paragraph.add_run(value)
    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)


def add_toc_field(paragraph) -> None:
    add_page_field(paragraph, "Update table of contents")
    instruction = paragraph._p.xpath(".//w:instrText")[0]
    instruction.text = ' TOC \\o "1-4" \\h \\z '


def rewrite_package(source: Path, output: Path, transform) -> None:
    with zipfile.ZipFile(source) as package, zipfile.ZipFile(
        output, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in package.infolist():
            target.writestr(info, transform(info.filename, package.read(info.filename)))


class V030WholeBookRuntimeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def synthetic_book(self) -> Path:
        image = self.root / "pixel.png"
        image.write_bytes(PNG_1X1)
        path = self.root / "synthetic-book.docx"
        document = Document()
        document.add_paragraph("Synthetic title")
        document.add_paragraph("Chapter body")
        document.add_paragraph("APPENDIX A Load cases")
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Column A"
        table.cell(0, 1).text = "Column B"
        table.cell(1, 0).text = "1"
        table.cell(1, 1).text = "2"
        document.add_paragraph().add_run().add_picture(str(image), width=Inches(1))
        document.save(path)
        return path

    def test_candidate_map_is_private_bounded_and_preserves_appendices(self) -> None:
        source = self.synthetic_book()
        structure = candidate_structure_map(source)
        serialized = json.dumps(structure, ensure_ascii=False)
        self.assertEqual("1.5", structure["schema_version"])
        self.assertNotIn("APPENDIX A Load cases", serialized)
        self.assertEqual("preserve_existing", structure["appendices"][0]["numbering_mode"])
        self.assertIsNone(structure["appendices"][0]["include_in_toc"])
        self.assertFalse(structure["trial_selection"]["whole_book_candidate"])
        self.assertLessEqual(
            structure["trial_selection"]["max_rendered_pages_per_candidate"], 30
        )
        self.assertEqual([], structure["tables"][0]["header_rows"])
        self.assertEqual("center", structure["tables"][0]["visual"]["alignment"])
        self.assertEqual("none", structure["tables"][0]["visual"]["text_wrapping"])
        self.assertTrue(
            all(not image["resize"]["allow_upscale"] for image in structure["images"])
        )

    def test_large_synthetic_book_keeps_trial_samples_bounded(self) -> None:
        source = self.root / "large-synthetic.docx"
        document = Document()
        document.add_paragraph("Synthetic title")
        for chapter in range(1, 21):
            document.add_paragraph(f"第{chapter}章 Synthetic chapter {chapter}")
            for section in range(1, 11):
                document.add_paragraph(
                    f"{chapter}.{section} Synthetic section {chapter}-{section}"
                )
                document.add_paragraph("Synthetic body paragraph")
            table = document.add_table(rows=3, cols=3)
            for row in table.rows:
                for cell in row.cells:
                    cell.text = "Synthetic cell"
        for appendix in ("A", "B", "C"):
            document.add_paragraph(f"APPENDIX {appendix} Synthetic appendix")
        document.save(source)

        structure = candidate_structure_map(source)
        trial = structure["trial_selection"]
        self.assertGreaterEqual(len(structure["headings"]), 200)
        self.assertEqual(3, len(structure["appendices"]))
        self.assertLessEqual(len(trial["heading_samples"]), 8)
        self.assertLessEqual(len(trial["appendix_samples"]), 2)
        self.assertLessEqual(len(trial["table_samples"]), 2)
        self.assertFalse(trial["whole_book_candidate"])

    def test_schema_14_remains_readable(self) -> None:
        legacy = candidate_structure_map(self.synthetic_book())
        legacy["schema_version"] = "1.4"
        legacy["status"] = "approved"
        path = self.root / "legacy-map.json"
        path.write_text(json.dumps(legacy), encoding="utf-8")
        self.assertEqual("1.4", load_structure_map(path)["schema_version"])

    def test_prepare_resume_uses_cache_and_source_change_invalidates_it(self) -> None:
        source = self.synthetic_book()
        profile = SKILL / "examples" / "profiles" / "technical-textbook-layout.v0.2.5.draft.json"
        work = self.root / "run"
        command = [
            sys.executable,
            str(SCRIPTS / "run_monograph.py"),
            "prepare",
            str(source),
            "--profile",
            str(profile),
            "--work-dir",
            str(work),
        ]
        first = subprocess.run(command, capture_output=True, text=True, check=False)
        self.assertEqual(0, first.returncode, first.stderr)
        first_state = json.loads((work / "run-state.json").read_text(encoding="utf-8"))
        second = subprocess.run(
            [*command, "--resume"], capture_output=True, text=True, check=False
        )
        self.assertEqual(0, second.returncode, second.stderr)
        second_state = json.loads((work / "run-state.json").read_text(encoding="utf-8"))
        self.assertEqual(1, second_state["metrics"]["cache_hits"])
        self.assertEqual(first_state["run_id"], second_state["run_id"])

        document = Document(source)
        document.add_paragraph("Changed synthetic source")
        document.save(source)
        third = subprocess.run(
            [*command, "--resume"], capture_output=True, text=True, check=False
        )
        self.assertEqual(0, third.returncode, third.stderr)
        third_state = json.loads((work / "run-state.json").read_text(encoding="utf-8"))
        self.assertNotEqual(first_state["run_id"], third_state["run_id"])
        self.assertEqual(0, third_state["metrics"]["cache_hits"])

    def test_environment_exposes_portable_capability_contract(self) -> None:
        completed = subprocess.run(
            [sys.executable, str(SCRIPTS / "check_environment.py"), "--json"],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(0, completed.returncode, completed.stderr)
        result = json.loads(completed.stdout)
        expected = {
            "file_read",
            "python_execution",
            "docx_inspection",
            "profile_validation",
            "docx_editing",
            "font_discovery",
            "rendering",
            "target_word",
            "field_update",
            "multimodal_source_reading",
        }
        self.assertEqual(expected, set(result["portable_capabilities"]))
        self.assertIsNone(
            result["portable_capabilities"]["multimodal_source_reading"]["available"]
        )

    def test_controlled_writeback_accepts_only_field_result_changes(self) -> None:
        image = self.root / "pixel.png"
        image.write_bytes(PNG_1X1)
        baseline = self.root / "baseline.docx"
        document = Document()
        document.add_paragraph("Authored content")
        document.add_picture(str(image))
        add_page_field(document.add_paragraph(), "1")
        document.save(baseline)
        refreshed = self.root / "refreshed.docx"

        def approved_transform(name: str, data: bytes) -> bytes:
            if name == "word/document.xml":
                root = etree.fromstring(data)
                result = root.xpath(
                    ".//w:fldChar[@w:fldCharType='separate']/following::w:t[1]",
                    namespaces=NS,
                )[0]
                result.text = "2"
                return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            return data

        rewrite_package(baseline, refreshed, approved_transform)
        output = self.root / "output.docx"
        replaced = controlled_field_result_writeback(baseline, refreshed, output)
        self.assertIn("word/document.xml", replaced)
        with zipfile.ZipFile(baseline) as before, zipfile.ZipFile(output) as after:
            media = next(name for name in before.namelist() if name.startswith("word/media/"))
            self.assertEqual(before.read(media), after.read(media))
            self.assertIn(b">2<", after.read("word/document.xml"))

        rejected = self.root / "rejected.docx"

        def rejected_transform(name: str, data: bytes) -> bytes:
            if name != "word/document.xml":
                return data
            root = etree.fromstring(data)
            root.xpath(".//w:t", namespaces=NS)[0].text = "Changed authored content"
            return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

        rewrite_package(baseline, refreshed, rejected_transform)
        with self.assertRaises(Exception):
            controlled_field_result_writeback(baseline, refreshed, rejected)
        self.assertFalse(rejected.exists())

        def protected_transform(name: str, data: bytes) -> bytes:
            if name.startswith("word/media/"):
                return b"target-application-mutated-media"
            return approved_transform(name, data)

        rewrite_package(baseline, refreshed, protected_transform)
        with self.assertRaises(Exception):
            controlled_field_result_writeback(
                baseline, refreshed, self.root / "protected-rejected.docx"
            )

    def test_controlled_writeback_accepts_multi_paragraph_toc_cache(self) -> None:
        baseline = self.root / "toc-baseline.docx"
        document = Document()
        add_toc_field(document.add_paragraph())
        document.add_paragraph("Authored body")
        document.save(baseline)
        refreshed = self.root / "toc-refreshed.docx"

        def expand_toc(name: str, data: bytes) -> bytes:
            if name != "word/document.xml":
                return data
            root = etree.fromstring(data)
            body = root.xpath("/w:document/w:body", namespaces=NS)[0]
            paragraph = body.xpath("./w:p[.//w:instrText]", namespaces=NS)[0]
            end_run = paragraph.xpath(
                ".//w:r[w:fldChar[@w:fldCharType='end']]", namespaces=NS
            )[0]
            paragraph.remove(end_run)
            result = paragraph.xpath(
                ".//w:fldChar[@w:fldCharType='separate']/following::w:t[1]",
                namespaces=NS,
            )[0]
            result.text = "Chapter 1"
            second = OxmlElement("w:p")
            second_run = OxmlElement("w:r")
            second_text = OxmlElement("w:t")
            second_text.text = "Section 1.1"
            second_run.append(second_text)
            second.append(second_run)
            second.append(end_run)
            body.insert(body.index(paragraph) + 1, second)
            return etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )

        rewrite_package(baseline, refreshed, expand_toc)
        output = self.root / "toc-output.docx"
        replaced = controlled_field_result_writeback(baseline, refreshed, output)
        self.assertIn("word/document.xml", replaced)
        reloaded = Document(output)
        self.assertEqual("Authored body", reloaded.paragraphs[-1].text)
        self.assertIn("Section 1.1", [paragraph.text for paragraph in reloaded.paragraphs])

    def test_final_ready_manifest_requires_every_page_and_no_issues(self) -> None:
        manifest = self.root / "visual.json"
        manifest.write_text(
            json.dumps(
                {
                    "all_pages_inspected": True,
                    "target_layout_verified": True,
                    "page_count": 12,
                    "issues": [],
                }
            ),
            encoding="utf-8",
        )
        self.assertEqual(12, validate_visual_manifest(manifest, 12)["page_count"])
        with self.assertRaises(Exception):
            validate_visual_manifest(manifest, 11)
        self.assertFalse(
            has_target_layout_evidence(
                {
                    "target_software": None,
                    "target_pdf_source": None,
                    "target_layout_unverified": False,
                }
            )
        )
        self.assertTrue(
            has_target_layout_evidence(
                {
                    "target_software": "LibreOffice Writer",
                    "target_pdf_source": None,
                    "target_layout_unverified": False,
                }
            )
        )
        self.assertTrue(
            has_target_layout_evidence(
                {
                    "target_software": "Microsoft Word",
                    "target_pdf_source": "target.pdf",
                    "target_layout_unverified": False,
                }
            )
        )


if __name__ == "__main__":
    unittest.main()
