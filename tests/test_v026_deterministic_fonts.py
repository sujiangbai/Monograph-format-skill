from __future__ import annotations

import copy
import hashlib
import struct
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from lxml import etree

REPO = Path(__file__).resolve().parents[1]
SCRIPTS = REPO / "format-monograph" / "scripts"
sys.path.insert(0, str(SCRIPTS))

from _common import (  # noqa: E402
    FormatMonographError,
    apply_field_properties,
    apply_style_properties,
    apply_style_rule_to_paragraphs,
    apply_table_properties,
    ensure_paragraph_style,
    style_effective_font,
    style_name_for_selector,
)
from audit_docx import audit_field_rule, audit_paragraph_rule, audit_table_rule  # noqa: E402
from docx_pagination import (  # noqa: E402
    apply_pagination_sections,
    audit_pagination_sections,
    finalize_pagination_sections,
    _ensure_page_field,
    _page_field_count,
    _resolve_audit_boundary,
    pagination_inventory,
)
from finalize_docx import effective_font_failures  # noqa: E402
from structure_map import (  # noqa: E402
    apply_structure_map,
    approved_data_tables,
    approved_role_paragraphs,
    audit_caption_identifier_replacements,
    audit_structure_heading_operations,
    audit_structure_image_operations,
    audit_structure_table_operations,
    candidate_structure_map,
    prime_structure_map_locators,
    resolve_paragraph_locator,
    structure_content_inventory,
    structure_content_fingerprint,
    text_sha256,
)


def set_theme_fonts(element, ascii_theme: str, east_asia_theme: str) -> None:
    r_pr = element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    for name in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.attrib.pop(qn(f"w:{name}"), None)
    r_fonts.set(qn("w:asciiTheme"), ascii_theme)
    r_fonts.set(qn("w:hAnsiTheme"), ascii_theme)
    r_fonts.set(qn("w:eastAsiaTheme"), east_asia_theme)


def synthetic_png(width: int, height: int) -> bytes:
    def chunk(kind: bytes, payload: bytes) -> bytes:
        return (
            struct.pack(">I", len(payload))
            + kind
            + payload
            + struct.pack(">I", zlib.crc32(kind + payload) & 0xFFFFFFFF)
        )

    scanline = b"\x00" + b"\x66\x99\xcc" * width
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(scanline * height, 9))
        + chunk(b"IEND", b"")
    )


def replace_theme_fonts(path: Path) -> None:
    temp = path.with_suffix(".theme.docx")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(
        temp, "w", zipfile.ZIP_DEFLATED
    ) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "word/theme/theme1.xml":
                root = etree.fromstring(data)
                namespace = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                for family, typeface in (
                    ("majorFont", "Synthetic Heading Theme"),
                    ("minorFont", "Synthetic Body Theme"),
                ):
                    fonts = root.xpath(
                        f"//a:fontScheme/a:{family}/a:font[@script='Hans']",
                        namespaces=namespace,
                    )
                    if fonts:
                        fonts[0].set("typeface", typeface)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            target.writestr(info, data)
    temp.replace(path)


def add_page_field(paragraph) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def page_field_element(cache: str = "CXLI") -> OxmlElement:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = cache
    run.append(text)
    field.append(run)
    return field


class V026DeterministicFontTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_theme_fonts_are_resolved_then_removed_for_controlled_roles(self) -> None:
        source = self.root / "theme-source.docx"
        document = Document()
        body = document.add_paragraph("Synthetic body")
        heading = document.add_paragraph("Synthetic heading", style="Heading 2")
        set_theme_fonts(document.styles["Normal"].element, "minorHAnsi", "minorEastAsia")
        set_theme_fonts(document.styles["Heading 2"].element, "majorHAnsi", "majorEastAsia")
        set_theme_fonts(body.runs[0]._r, "minorHAnsi", "minorEastAsia")
        set_theme_fonts(heading.runs[0]._r, "majorHAnsi", "majorEastAsia")
        document.save(source)
        replace_theme_fonts(source)

        document = Document(source)
        self.assertEqual(
            "Synthetic Body Theme",
            style_effective_font(document, document.styles["Normal"], "eastAsia")[0],
        )
        self.assertEqual(
            "Synthetic Heading Theme",
            style_effective_font(document, document.styles["Heading 2"], "eastAsia")[0],
        )

        body_rule = {
            "id": "FMT-BODY-TEST",
            "selector": {"kind": "paragraph_role", "value": "body_text"},
            "properties": {
                "font_name_ascii": "Times New Roman",
                "font_name_east_asia": "å®‹ä½“",
                "font_name_complex_script": "Times New Roman",
            },
        }
        heading_rule = {
            "id": "FMT-HEAD-TEST",
            "selector": {"kind": "paragraph_role", "value": "level_2_section"},
            "properties": {
                "font_name_ascii": "Arial",
                "font_name_east_asia": "é»‘ä½“",
                "font_name_complex_script": "Arial",
            },
        }
        apply_style_rule_to_paragraphs(document, body_rule, [document.paragraphs[0]])
        apply_style_rule_to_paragraphs(document, heading_rule, [document.paragraphs[1]])
        output = self.root / "deterministic.docx"
        document.save(output)

        formatted = Document(output)
        for style_name, ascii_name, east_asia_name in (
            ("Normal", "Times New Roman", "å®‹ä½“"),
            ("Heading 2", "Arial", "é»‘ä½“"),
        ):
            r_fonts = formatted.styles[style_name].element.rPr.rFonts
            self.assertEqual(ascii_name, r_fonts.get(qn("w:ascii")))
            self.assertEqual(east_asia_name, r_fonts.get(qn("w:eastAsia")))
            for attribute in (
                "asciiTheme",
                "hAnsiTheme",
                "eastAsiaTheme",
                "cstheme",
            ):
                self.assertIsNone(r_fonts.get(qn(f"w:{attribute}")))
        self.assertFalse(audit_paragraph_rule(formatted, body_rule, [formatted.paragraphs[0]]))
        self.assertFalse(
            audit_paragraph_rule(formatted, heading_rule, [formatted.paragraphs[1]])
        )
        profile = {
            "rules": [
                {**body_rule, "status": "approved", "application": "automatic"},
                {**heading_rule, "status": "approved", "application": "automatic"},
            ]
        }
        self.assertFalse(effective_font_failures(output, profile))

    def test_all_controlled_text_styles_receive_mixed_script_fonts(self) -> None:
        source = self.root / "role-styles.docx"
        document = Document()
        cases = (
            ("Normal", {"kind": "paragraph_role", "value": "body_text"}, "å®‹ä½“", "Times New Roman"),
            ("Heading 1", {"kind": "paragraph_role", "value": "heading_1"}, "é»‘ä½“", "Arial"),
            ("Heading 2", {"kind": "paragraph_role", "value": "heading_2"}, "é»‘ä½“", "Arial"),
            ("Heading 3", {"kind": "paragraph_role", "value": "heading_3"}, "é»‘ä½“", "Arial"),
            ("Heading 4", {"kind": "paragraph_role", "value": "heading_4"}, "å®‹ä½“", "Times New Roman"),
            ("TOC 1", {"kind": "paragraph_role", "value": "toc_level_1"}, "å®‹ä½“", "Times New Roman"),
            ("TOC 2", {"kind": "paragraph_role", "value": "toc_level_2"}, "å®‹ä½“", "Times New Roman"),
            ("TOC 3", {"kind": "paragraph_role", "value": "toc_level_3"}, "å®‹ä½“", "Times New Roman"),
            ("Caption", {"kind": "caption_role", "value": "all"}, "å®‹ä½“", "Times New Roman"),
            ("Quote", {"kind": "paragraph_role", "value": "long_quote"}, "å®‹ä½“", "Times New Roman"),
            ("Bibliography", {"kind": "bibliography_role", "value": "entries"}, "å®‹ä½“", "Times New Roman"),
            ("Footnote Text", {"kind": "style_name", "value": "Footnote Text"}, "å®‹ä½“", "Times New Roman"),
        )
        for style_name, _selector, _east_asia, _ascii in cases:
            style = ensure_paragraph_style(document, style_name)
            set_theme_fonts(style.element, "minorHAnsi", "minorEastAsia")
        document.save(source)
        replace_theme_fonts(source)

        document = Document(source)
        profile_rules = []
        for index, (style_name, selector, east_asia, ascii_name) in enumerate(cases):
            paragraph = document.add_paragraph(f"Synthetic role {index}", style=style_name)
            rule = {
                "id": f"FMT-ROLE-{index}",
                "selector": selector,
                "properties": {
                    "font_name_ascii": ascii_name,
                    "font_name_east_asia": east_asia,
                    "font_name_complex_script": ascii_name,
                },
            }
            apply_style_rule_to_paragraphs(document, rule, [paragraph])
            profile_rules.append(
                {**rule, "status": "approved", "application": "automatic"}
            )
        output = self.root / "role-styles-formatted.docx"
        document.save(output)

        formatted = Document(output)
        for style_name, _selector, east_asia, ascii_name in cases:
            self.assertEqual(
                ascii_name,
                style_effective_font(formatted, formatted.styles[style_name], "ascii")[0],
            )
            self.assertEqual(
                east_asia,
                style_effective_font(formatted, formatted.styles[style_name], "eastAsia")[0],
            )
        self.assertFalse(
            effective_font_failures(output, {"rules": profile_rules})
        )

    def test_page_footer_rebuild_is_idempotent_and_protects_other_content(self) -> None:
        document = Document()
        footer = document.sections[0].footer
        paragraph = footer.paragraphs[0]
        add_page_field(paragraph)
        add_page_field(paragraph)
        self.assertTrue(_ensure_page_field(footer, 2))
        self.assertEqual(1, _page_field_count(footer._element))
        self.assertEqual(1, len(footer.paragraphs))
        self.assertFalse(_ensure_page_field(footer, 2))
        self.assertEqual(1, _page_field_count(footer._element))

        protected = Document().sections[0].footer
        protected.paragraphs[0].text = "Synthetic chapter footer"
        with self.assertRaises(FormatMonographError):
            _ensure_page_field(protected, 2)

        static_page = Document().sections[0].footer
        static_page.paragraphs[0].text = "315"
        self.assertTrue(
            _ensure_page_field(static_page, 2, replace_static_page_text=True)
        )
        self.assertEqual(1, _page_field_count(static_page._element))

        duplicate_path = self.root / "duplicate-footer.docx"
        duplicated = Document()
        add_page_field(duplicated.sections[0].footer.paragraphs[0])
        add_page_field(duplicated.sections[0].footer.paragraphs[0])
        duplicated.save(duplicate_path)
        inventory = pagination_inventory(duplicate_path)
        self.assertEqual(
            2, inventory["sections"][0]["footer_page_field_counts"]["default"]
        )

        legacy = Document().sections[0].footer
        legacy_paragraph = legacy.paragraphs[0]
        for container_name in ("drawing", "pict"):
            container = OxmlElement(f"w:{container_name}")
            text_box = OxmlElement("w:txbxContent")
            inner_paragraph = OxmlElement("w:p")
            inner_paragraph.append(page_field_element())
            text_box.append(inner_paragraph)
            container.append(text_box)
            if container_name == "pict":
                container.append(
                    etree.Element("{urn:schemas-microsoft-com:vml}imagedata")
                )
            legacy_paragraph._p.append(container)
        self.assertTrue(_ensure_page_field(legacy, 2))
        self.assertEqual(1, _page_field_count(legacy._element))
        self.assertFalse(legacy._element.xpath(".//w:drawing | .//w:pict"))

        linked_vml_footer = Document().sections[0].footer
        pict = OxmlElement("w:pict")
        text_box = OxmlElement("w:txbxContent")
        inner_paragraph = OxmlElement("w:p")
        inner_paragraph.append(page_field_element())
        text_box.append(inner_paragraph)
        pict.append(text_box)
        image_data = etree.Element("{urn:schemas-microsoft-com:vml}imagedata")
        image_data.set(qn("r:id"), "rIdSyntheticImage")
        pict.append(image_data)
        linked_vml_footer.paragraphs[0]._p.append(pict)
        self.assertFalse(_ensure_page_field(linked_vml_footer, 2))
        self.assertTrue(linked_vml_footer._element.xpath(".//w:pict"))

        image_footer = Document().sections[0].footer
        drawing = OxmlElement("w:drawing")
        drawing.append(OxmlElement("a:blip"))
        image_footer.paragraphs[0]._p.append(drawing)
        with self.assertRaises(FormatMonographError):
            _ensure_page_field(image_footer, 2)

    def test_table_fonts_are_explicit_and_effectively_audited(self) -> None:
        output = self.root / "table-fonts.docx"
        document = Document()
        table = document.add_table(rows=2, cols=2)
        for row_index, row in enumerate(table.rows):
            for cell_index, cell in enumerate(row.cells):
                ce÷_7¶‰žËkºwµçU¹Ð ‰Üé‘É…Ý¥¹œˆ¤¤4(€€€€€€€Õ¹¹Õµ‰•É•€ô‘½Õµ•¹Ð¹…‘‘}Á…É…É…Á  ‰Må¹Ñ¡•Ñ¥ŒÝ…Ù•™½É´ˆ¤4(€€€€€€€Õ¹¹Õµ‰•É•¹…±¥¹µ•¹Ð€ô]}1%9}AIIA ¹9QH4(4(€€€€€€€Á…¹•°€ô‘½Õµ•¹Ð¹…‘‘}Ñ…‰±”¡É½ÝÌôÈ°½±ÌôÈ¤4(€€€€€€€™½È•±°¥¸Á…¹•°¹É½ÝÍlÁt¹•±±Ìè4(€€€€€€€€€€€•±°¹Á…É…É…Á¡ÍlÁt¹}À¹…ÁÁ•¹¡=áµ±±•µ•¹Ð ‰Üé‘É…Ý¥¹œˆ¤¤4(€€€€€€€™½È¥¹‘•à°•±°¥¸•¹Õµ•É…Ñ”¡Á…¹•°¹É½ÝÍlÅt¹•±±Ì¤è4(€€€€€€€€€€€•±°¹Ñ•áÐ€ô˜‰A…¹•°í¥¹‘•à€¬€Åôˆ4(€€€€€€€€€€€•±°¹Á…É…É…Á¡ÍlÁt¹…±¥¹µ•¹Ð€ô]}1%9}AIIA ¹9QH4(4(€€€€€€€‘…Ñ„€ô‘½Õµ•¹Ð¹…‘‘}Ñ…‰±”¡É½ÝÌôÈ°½±ÌôÈ¤4(€€€€€€€‘…Ñ„¹•±° À°€À¤¹Ñ•áÐ€ô€‰!•…‘•Èˆ4(€€€€€€€‘…Ñ„¹•±° À°€Ä¤¹Ñ•áÐ€ô€‰!•…‘•Èˆ4(€€€€€€€±•…¹ÕÁ}•±°€ô‘…Ñ„¹•±° Ä°€À¤4(€€€€€€€±•…¹ÕÁ}•±°¹…‘‘}Á…É…É…Á  ‰½¹±ÕÍ¥½¸ˆ¤4(€€€€€€€‘…Ñ„¹•±° Ä°€Ä¤¹Ñ•áÐ€ô€‰I•ÍÕ±Ðˆ4(€€€€€€€‘½Õµ•¹Ð¹Í…Ù”¡Í½ÕÉ”¤4(4(€€€€€€€ÍÑÉÕÑÕÉ”€ô…¹‘¥‘…Ñ•}ÍÑÉÕÑÕÉ•}µ…À¡Í½ÕÉ”¤4(€€€€€€€ÍÑÉÕÑÕÉ•l‰ÍÑ…ÑÕÌ‰t€ô€‰…ÁÁÉ½Ù•ˆ4(€€€€€€€Á…¹•±}•¹ÑÉä€ôÍÑÉÕÑÕÉ•l‰Ñ…‰±•Ì‰ulÁt4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° ‰±…å½ÕÐˆ°Á…¹•±}•¹ÑÉål‰­¥¹‰t¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° ‰™¥ÕÉ•}Á…¹•°ˆ°Á…¹•±}•¹ÑÉål‰±…å½ÕÑ}ÁÕÉÁ½Í”‰t¤4(€€€€€€€Á…¹•±}•¹ÑÉål‰…ÁÁÉ½Ù•‰t€ôQÉÕ”4(€€€€€€€Á…¹•±}•¹ÑÉål‰Á…¥¹…Ñ¥½¹}½¹±ä‰t€ô…±Í”4(€€€€€€€Á…¹•±}•¹ÑÉål‰Ù¥ÍÕ…°‰ul‰…ÁÁÉ½Ù•‰t€ôQÉÕ”4(€€€€€€€™½ÈÉ½±”¥¸ÍÑÉÕÑÕÉ•l‰Á…É…É…Á¡}É½±•Ì‰tè4(€€€€€€€€€€€¥˜É½±•l‰É½±”‰t¥¸ì‰™¥ÕÉ•}…ÁÑ¥½¹}Õ¹¹Õµ‰•É•ˆ°€‰™¥ÕÉ•}Á…¹•±}±…‰•°‰ôè4(€€€€€€€€€€€€€€€É½±•l‰…ÁÁÉ½Ù•‰t€ôQÉÕ”4(4(€€€€€€€‘…Ñ…}•¹ÑÉä€ôÍÑÉÕÑÕÉ•l‰Ñ…‰±•Ì‰ulÅt4(€€€€€€€‘…Ñ…}•¹ÑÉä¹ÕÁ‘…Ñ”¡ì‰…ÁÁÉ½Ù•ˆèQÉÕ”°€‰­¥¹ˆè€‰‘…Ñ„ˆ°€‰¡•…‘•É}É½ÝÌˆèlÁuô¤4(€€€€€€€ÍÑÉÕÑÕÉ•l‰Ñ…‰±•}•±±}±•…¹ÕÁÌ‰t€ôl4(€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€‰Ñ…‰±”ˆè€Ä°4(€€€€€€€€€€€€€€€€‰É½Üˆè€Ä°4(€€€€€€€€€€€€€€€€‰•±°ˆè€À°4(€€€€€€€€€€€€€€€€‰…Ñ¥½¸ˆè€‰É•µ½Ù•}±•…‘¥¹}•µÁÑå}Á…É…É…Á¡Ìˆ°4(€€€€€€€€€€€€€€€€‰½Õ¹Ðˆè€Ä°4(€€€€€€€€€€€€€€€€‰Ñ…‰±•}Ñ•áÑ}Í¡„ÈÔØˆè‘…Ñ…}•¹ÑÉål‰Ñ…‰±•}Ñ•áÑ}Í¡„ÈÔØ‰t°4(€€€€€€€€€€€€€€€€‰•±±}Ñ•áÑ}Í¡„ÈÔØˆèÑ•áÑ}Í¡„ÈÔØ¡±•…¹ÕÁ}•±°¹Ñ•áÐ¤°4(€€€€€€€€€€€€€€€€‰É•ÍÕ±Ñ}•±±}Ñ•áÑ}Í¡„ÈÔØˆèÑ•áÑ}Í¡„ÈÔØ ‰½¹±ÕÍ¥½¸ˆ¤°4(€€€€€€€€€€€€€€€€‰…ÁÁÉ½Ù•ˆèQÉÕ”°4(€€€€€€€€€€€ô4(€€€€€€€t4(4(€€€€€€€™½Éµ…ÑÑ•€ô½Õµ•¹Ð¡Í½ÕÉ”¤4(€€€€€€€ÁÉ¥µ•}ÍÑÉÕÑÕÉ•}µ…Á}±½…Ñ½ÉÌ¡™½Éµ…ÑÑ•°ÍÑÉÕÑÕÉ”¤4(€€€€€€€…ÁÁ±å}ÍÑÉÕÑÕÉ•}µ…À¡™½Éµ…ÑÑ•°ÍÑÉÕÑÕÉ”¤4(€€€€€€€…ÁÁ±å}ÍÑÉÕÑÕÉ•}µ…À¡™½Éµ…ÑÑ•°ÍÑÉÕÑÕÉ”¤4(€€€€€€€½ÕÑÁÕÐ€ôÍ•±˜¹É½½Ð€¼€‰™¥ÕÉ”µÁ…¹•°µ±•…¹ÕÀµ™½Éµ…ÑÑ•¹‘½àˆ4(€€€€€€€™½Éµ…ÑÑ•¹Í…Ù”¡½ÕÑÁÕÐ¤4(4(€€€€€€€É•±½…‘•€ô½Õµ•¹Ð¡½ÕÑÁÕÐ¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…°¡]}Q	1}1%959P¹9QH°É•±½…‘•¹Ñ…‰±•ÍlÁt¹…±¥¹µ•¹Ð¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑQÉÕ” 4(€€€€€€€€€€€…±° 4(€€€€€€€€€€€€€€€•±°¹Ù•ÉÑ¥…±}…±¥¹µ•¹Ð€ôô]}11}YIQ%1}1%959P¹9QH4(€€€€€€€€€€€€€€€™½ÈÉ½Ü¥¸É•±½…‘•¹Ñ…‰±•ÍlÁt¹É½ÝÌ4(€€€€€€€€€€€€€€€™½È•±°¥¸É½Ü¹•±±Ì4(€€€€€€€€€€€€¤4(€€€€€€€€¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑQÉÕ” 4(€€€€€€€€€€€…±° 4(€€€€€€€€€€€€€€€Á…É…É…Á ¹ÍÑå±”¹¹…µ”€ôô€‰…ÁÑ¥½¸ˆ4(€€€€€€€€€€€€€€€™½È•±°¥¸É•±½…‘•¹Ñ…‰±•ÍlÁt¹É½ÝÍlÅt¹•±±Ì4(€€€€€€€€€€€€€€€™½ÈÁ…É…É…Á ¥¸•±°¹Á…É…É…Á¡Ì4(€€€€€€€€€€€€€€€¥˜Á…É…É…Á ¹Ñ•áÐ4(€€€€€€€€€€€€¤4(€€€€€€€€¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° ‰½¹±ÕÍ¥½¸ˆ°É•±½…‘•¹Ñ…‰±•ÍlÅt¹•±° Ä°€À¤¹Ñ•áÐ¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° 4(€€€€€€€€€€€ÍÑÉÕÑÕÉ•}½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡Í½ÕÉ”°ÍÑÉÕÑÕÉ”¤°4(€€€€€€€€€€€ÍÑÉÕÑÕÉ•}½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡½ÕÑÁÕÐ°ÍÑÉÕÑÕÉ”¤°4(€€€€€€€€¤4(€€€€€€€É•±½…‘•¹ÍÑå±•Íl‰…ÁÑ¥½¸‰t¹Á…É…É…Á¡}™½Éµ…Ð¹…±¥¹µ•¹Ð€ô€ 4(€€€€€€€€€€€]}1%9}AIIA ¹9QH4(€€€€€€€€¤4(€€€€€€€É•±½…‘•¹ÍÑå±•Íl‰…ÁÑ¥½¸‰t¹•±•µ•¹Ð¹™¥¹¡Å¸ ‰Üé¹…µ”ˆ¤¤¹Í•Ð 4(€€€€€€€€€€€Å¸ ‰ÜéÙ…°ˆ¤°€‰1½…±¥é•¥ÕÉ”1…‰•°ˆ4(€€€€€€€€¤4(€€€€€€€™½È•±°¥¸É•±½…‘•¹Ñ…‰±•ÍlÁt¹É½ÝÍlÅt¹•±±Ìè4(€€€€€€€€€€€™½ÈÁ…É…É…Á ¥¸•±°¹Á…É…É…Á¡Ìè4(€€€€€€€€€€€€€€€Á…É…É…Á ¹…±¥¹µ•¹Ð€ô9½¹”4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ…±Í”¡…Õ‘¥Ñ}ÍÑÉÕÑÕÉ•}Ñ…‰±•}½Á•É…Ñ¥½¹Ì¡É•±½…‘•°ÍÑÉÕÑÕÉ”¤¤4(€€€€€€€É½±•Ì€ôì4(€€€€€€€€€€€É½±•l‰É½±”‰t4(€€€€€€€€€€€™½ÈÉ½±”¥¸ÍÑÉÕÑÕÉ•l‰Á…É…É…Á¡}É½±•Ì‰t4(€€€€€€€€€€€¥˜É½±•l‰…ÁÁÉ½Ù•‰t4(€€€€€€€ô4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%¸ ‰™¥ÕÉ•}…ÁÑ¥½¹}Õ¹¹Õµ‰•É•ˆ°É½±•Ì¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%¸ ‰™¥ÕÉ•}Á…¹•±}±…‰•°ˆ°É½±•Ì¤4(4(€€€‘•˜Ñ•ÍÑ}¥µ…•}É•Í¥é¥¹}ÁÉ•Í•ÉÙ•Í}µ•‘¥…}…¹¡½ÉÍ}½É‘•É}…¹‘}Ñ…‰±•}Á½Í¥Ñ¥½¸¡Í•±˜¤€´ø9½¹”è4(€€€€€€€Í½ÕÉ”€ôÍ•±˜¹É½½Ð€¼€‰¥µ…”µ…¹¡½ÈµÍ½ÕÉ”¹‘½àˆ4(€€€€€€€¥µ…•}Á…Ñ €ôÍ•±˜¹É½½Ð€¼€‰Íå¹Ñ¡•Ñ¥Œµ¥µ…”¹Á¹œˆ4(€€€€€€€¥µ…•}Á…Ñ ¹ÝÉ¥Ñ•}‰åÑ•Ì¡Íå¹Ñ¡•Ñ¥}Á¹œ ØÀÀ°€ÐÀÀ¤¤4(4(€€€€€€€‘½Õµ•¹Ð€ô½Õµ•¹Ð ¤4(€€€€€€€‘½Õµ•¹Ð¹…‘‘}Á…É…É…Á  ‰Må¹Ñ¡•Ñ¥Œ¡…ÁÑ•Èˆ°ÍÑå±”ô‰!•…‘¥¹œ€Äˆ¤4(€€€€€€€ÍÑ…¹‘…±½¹”€ô‘½Õµ•¹Ð¹…‘‘}Á…É…É…Á  ¤4(€€€€€€€ÍÑ…¹‘…±½¹”¹…‘‘}ÉÕ¸ ¤¹…‘‘}Á¥ÑÕÉ”¡ÍÑÈ¡¥µ…•}Á…Ñ ¤°Ý¥‘Ñ õ%¹¡•Ì Ä¤¤4(€€€€€€€‘½Õµ•¹Ð¹…‘‘}Á…É…É…Á  ‰Må¹Ñ¡•Ñ¥Œ™½±±½Ý¥¹œÁ…É…É…Á ˆ¤4(€€€€€€€Á…¹•°€ô‘½Õµ•¹Ð¹…‘‘}Ñ…‰±”¡É½ÝÌôÈ°½±ÌôÈ¤4(€€€€€€€Á…¹•°¹•±° À°€À¤¹Á…É…É…Á¡ÍlÁt¹…‘‘}ÉÕ¸ ¤¹…‘‘}Á¥ÑÕÉ” 4(€€€€€€€€€€€ÍÑÈ¡¥µ…•}Á…Ñ ¤°Ý¥‘Ñ õ%¹¡•Ì Ä¤4(€€€€€€€€¤4(€€€€€€€Á…¹•°¹•±° À°€Ä¤¹Á…É…É…Á¡ÍlÁt¹…‘‘}ÉÕ¸ ¤¹…‘‘}Á¥ÑÕÉ” 4(€€€€€€€€€€€ÍÑÈ¡¥µ…•}Á…Ñ ¤°Ý¥‘Ñ õ%¹¡•Ì Ä¸È¤4(€€€€€€€€¤4(€€€€€€€Á…¹•°¹•±° Ä°€À¤¹Ñ•áÐ€ô€‰A…¹•°ˆ4(€€€€€€€Á…¹•°¹•±° Ä°€Ä¤¹Ñ•áÐ€ô€‰A…¹•°ˆ4(€€€€€€€‘½Õµ•¹Ð¹…‘‘}Á…É…É…Á  ‰Må¹Ñ¡•Ñ¥ŒÑ…¥°ˆ¤4(€€€€€€€‘½Õµ•¹Ð¹Í…Ù”¡Í½ÕÉ”¤4(4(€€€€€€€ÍÑÉÕÑÕÉ”€ô…¹‘¥‘…Ñ•}ÍÑÉÕÑÕÉ•}µ…À¡Í½ÕÉ”¤4(€€€€€€€ÍÑÉÕÑÕÉ•l‰ÍÑ…ÑÕÌ‰t€ô€‰…ÁÁÉ½Ù•ˆ4(€€€€€€€Á…¹•±}•¹ÑÉä€ôÍÑÉÕÑÕÉ•l‰Ñ…‰±•Ì‰ulÁt4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° ‰™¥ÕÉ•}Á…¹•°ˆ°Á…¹•±}•¹ÑÉål‰±…å½ÕÑ}ÁÕÉÁ½Í”‰t¤4(€€€€€€€Á…¹•±}•¹ÑÉål‰…ÁÁÉ½Ù•‰t€ôQÉÕ”4(€€€€€€€Á…¹•±}•¹ÑÉål‰Á½Í¥Ñ¥½¹}Á½±¥ä‰t€ô€‰ÁÉ•Í•ÉÙ•}…¹¡½Èˆ4(€€€€€€€Á…¹•±}•¹ÑÉål‰Ù¥ÍÕ…°‰ul‰…ÁÁÉ½Ù•‰t€ôQÉÕ”4(€€€€€€€™½È¥µ…”¥¸ÍÑÉÕÑÕÉ•l‰¥µ…•Ì‰tè4(€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° ‰ÁÉ•Í•ÉÙ•}…¹¡½Èˆ°¥µ…•l‰Á½Í¥Ñ¥½¹}Á½±¥ä‰t¤4(€€€€€€€€€€€¥˜¥µ…•l‰Á±…•µ•¹Ð‰t¥¸ì‰ÍÑ…¹‘…±½¹”ˆ°€‰Ñ…‰±•}™¥ÕÉ•}Á…¹•°‰ôè4(€€€€€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑQÉÕ”¡¥µ…•l‰ÍÕÁÁ½ÉÑ•‰t¤4(€€€€€€€€€€€€€€€¥µ…•l‰…ÁÁÉ½Ù•‰t€ôQÉÕ”4(€€€€€€€€€€€€€€€¥µ…•l‰É•Í¥é”‰ul‰…ÁÁÉ½Ù•‰t€ôQÉÕ”4(4(€€€€€€€™½Éµ…ÑÑ•€ô½Õµ•¹Ð¡Í½ÕÉ”¤4(€€€€€€€€Œ‘•É¥Ù••µÁÑäÁ…É…É…Á µ…äÍ¡…É”Ñ¡”¥µ…”Á…É…É…Á ÌÑ•áÐ¡…Í …¹4(€€€€€€€€ŒÍ¡¥™Ð¥ÑÌ¹Õµ•É¥Œ¥¹‘•à¸5•‘¥„Á±ÕÌÕ¹¡…¹•…ÕÑ¡½É•¹•¥¡‰½ÉÌµÕÍÐ4(€€€€€€€€ŒÍÑ¥±°É•Í½±Ù”Ñ¡”½É¥¥¹…°¥µ…”Ý¥Ñ¡½ÕÐµ½Ù¥¹œ¥Ð¸4(€€€€€€€™½Éµ…ÑÑ•¹Á…É…É…Á¡ÍlÅt¹}À¹…‘‘ÁÉ•Ù¥½ÕÌ¡=áµ±±•µ•¹Ð ‰ÜéÀˆ¤¤4(€€€€€€€‰½‘å}¡¥±‘É•¹}‰•™½É”€ô±¥ÍÐ¡™½Éµ…ÑÑ•¹•±•µ•¹Ð¹‰½‘ä¤4(€€€€€€€ÍÑ…¹‘…±½¹•}Á…É…É…Á €ô™½Éµ…ÑÑ•¹Á…É…É…Á¡ÍlÉt4(€€€€€€€Á…¹•±}Ñ…‰±”€ô™½Éµ…ÑÑ•¹Ñ…‰±•ÍlÁt4(€€€€€€€Á…¹•±}Á…É…É…Á¡Ì€ôl4(€€€€€€€€€€€Á…¹•±}Ñ…‰±”¹•±° À°€À¤¹Á…É…É…Á¡ÍlÁt°4(€€€€€€€€€€€Á…¹•±}Ñ…‰±”¹•±° À°€Ä¤¹Á…É…É…Á¡ÍlÁt°4(€€€€€€€t4(€€€€€€€ÍÑ…¹‘…±½¹•}Á…É•¹Ð€ôÍÑ…¹‘…±½¹•}Á…É…É…Á ¹}À¹•ÑÁ…É•¹Ð ¤4(€€€€€€€ÍÑ…¹‘…±½¹•}¥¹‘•à€ôÍÑ…¹‘…±½¹•}Á…É•¹Ð¹¥¹‘•à¡ÍÑ…¹‘…±½¹•}Á…É…É…Á ¹}À¤4(€€€€€€€Á…¹•±}Á…É•¹Ð€ôÁ…¹•±}Ñ…‰±”¹}Ñ‰°¹•ÑÁ…É•¹Ð ¤4(€€€€€€€Á…¹•±}¥¹‘•à€ôÁ…¹•±}Á…É•¹Ð¹¥¹‘•à¡Á…¹•±}Ñ…‰±”¹}Ñ‰°¤4(€€€€€€€Á…¹•±}±½…Ñ¥½¹Ì€ôl4(€€€€€€€€€€€€¡Á…É…É…Á ¹}À¹•ÑÁ…É•¹Ð ¤°Á…É…É…Á ¹}À¹•ÑÁ…É•¹Ð ¤¹¥¹‘•à¡Á…É…É…Á ¹}À¤¤4(€€€€€€€€€€€™½ÈÁ…É…É…Á ¥¸Á…¹•±}Á…É…É…Á¡Ì4(€€€€€€€t4(€€€€€€€½É¥¥¹…±}•áÑ•¹ÑÌ€ôì4(€€€€€€€€€€€¥µ…•l‰¥µ…”‰tè‘¥Ð¡¥µ…•l‰Í½ÕÉ•}•áÑ•¹Ñ}•µÔ‰t¤4(€€€€€€€€€€€™½È¥µ…”¥¸ÍÑÉÕÑÕÉ•l‰¥µ…•Ì‰t4(€€€€€€€€€€€¥˜¥µ…•l‰…ÁÁÉ½Ù•‰t4(€€€€€€€ô4(€€€€€€€Í¡¥™Ñ•‘}Í½ÕÉ”€ôÍ•±˜¹É½½Ð€¼€‰¥µ…”µ…¹¡½ÈµÍ¡¥™Ñ•µÍ½ÕÉ”¹‘½àˆ4(€€€€€€€™½Éµ…ÑÑ•¹Í…Ù”¡Í¡¥™Ñ•‘}Í½ÕÉ”¤4(4(€€€€€€€ÁÉ¥µ•}ÍÑÉÕÑÕÉ•}µ…Á}±½…Ñ½ÉÌ¡™½Éµ…ÑÑ•°ÍÑÉÕÑÕÉ”¤4(€€€€€€€É•ÍÕ±Ð€ô…ÁÁ±å}ÍÑÉÕÑÕÉ•}µ…À¡™½Éµ…ÑÑ•°ÍÑÉÕÑÕÉ”¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° Ì°¹•áÐ¡¡…¹•l‰Ñ…É•ÑÌ‰t™½È¡…¹”¥¸É•ÍÕ±Ð¥˜¡…¹•l‰­¥¹‰t€ôô€‰ÍÑÉÕÑÕÉ•}¥µ…•}É•Í¥é”ˆ¤¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…°¡‰½‘å}¡¥±‘É•¹}‰•™½É”°±¥ÍÐ¡™½Éµ…ÑÑ•¹•±•µ•¹Ð¹‰½‘ä¤¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%Ì¡ÍÑ…¹‘…±½¹•}Á…É•¹Ð°ÍÑ…¹‘…±½¹•}Á…É…É…Á ¹}À¹•ÑÁ…É•¹Ð ¤¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…°¡ÍÑ…¹‘…±½¹•}¥¹‘•à°ÍÑ…¹‘…±½¹•}Á…É•¹Ð¹¥¹‘•à¡ÍÑ…¹‘…±½¹•}Á…É…É…Á ¹}À¤¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%Ì¡Á…¹•±}Á…É•¹Ð°Á…¹•±}Ñ…‰±”¹}Ñ‰°¹•ÑÁ…É•¹Ð ¤¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…°¡Á…¹•±}¥¹‘•à°Á…¹•±}Á…É•¹Ð¹¥¹‘•à¡Á…¹•±}Ñ…‰±”¹}Ñ‰°¤¤4(€€€€€€€™½ÈÁ…É…É…Á °€¡Á…É•¹Ð°¥¹‘•à¤¥¸é¥À¡Á…¹•±}Á…É…É…Á¡Ì°Á…¹•±}±½…Ñ¥½¹Ì¤è4(€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%Ì¡Á…É•¹Ð°Á…É…É…Á ¹}À¹•ÑÁ…É•¹Ð ¤¤4(€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…°¡¥¹‘•à°Á…É•¹Ð¹¥¹‘•à¡Á…É…É…Á ¹}À¤¤4(4(€€€€€€€½ÕÑÁÕÐ€ôÍ•±˜¹É½½Ð€¼€‰¥µ…”µ…¹¡½Èµ™½Éµ…ÑÑ•¹‘½àˆ4(€€€€€€€™½Éµ…ÑÑ•¹Í…Ù”¡½ÕÑÁÕÐ¤4(€€€€€€€É•±½…‘•€ô½Õµ•¹Ð¡½ÕÑÁÕÐ¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ…±Í”¡…Õ‘¥Ñ}ÍÑÉÕÑÕÉ•}¥µ…•}½Á•É…Ñ¥½¹Ì¡É•±½…‘•°ÍÑÉÕÑÕÉ”¤¤4(€€€€€€€É•Í¥é•‘}•áÑ•¹ÑÌ€ôì4(€€€€€€€€€€€¥µ…•l‰¥µ…”‰tèÉ•±½…‘•¹Á…ÉÐ¹‘½Õµ•¹Ð¹•±•µ•¹Ð¹áÁ…Ñ  4(€€€€€€€€€€€€€€€€ˆ¸¼½Üé‘É…Ý¥¹œˆ4(€€€€€€€€€€€€¥m¥¹‘•át4(€€€€€€€€€€€€¹áÁ…Ñ  ˆ¸½ÝÀé¥¹±¥¹”½ÝÀé•áÑ•¹Ðˆ¥lÁt4(€€€€€€€€€€€™½È¥¹‘•à°¥µ…”¥¸•¹Õµ•É…Ñ” 4(€€€€€€€€€€€€€€€m¥Ñ•´™½È¥Ñ•´¥¸ÍÑÉÕÑÕÉ•l‰¥µ…•Ì‰t¥˜¥Ñ•µl‰…ÁÁÉ½Ù•‰ut4(€€€€€€€€€€€€¤4(€€€€€€€ô4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑQÉÕ” 4(€€€€€€€€€€€…±° 4(€€€€€€€€€€€€€€€¥¹Ð¡•áÑ•¹Ð¹•Ð ‰àˆ¤¤4(€€€€€€€€€€€€€€€€ðô¥¹Ð¡½É¥¥¹…±}•áÑ•¹ÑÍm¥µ…•}¥‘ul‰à‰t¤4(€€€€€€€€€€€€€€€…¹¥¹Ð¡•áÑ•¹Ð¹•Ð ‰äˆ¤¤4(€€€€€€€€€€€€€€€€ðô¥¹Ð¡½É¥¥¹…±}•áÑ•¹ÑÍm¥µ…•}¥‘ul‰ä‰t¤4(€€€€€€€€€€€€€€€™½È¥µ…•}¥°•áÑ•¹Ð¥¸É•Í¥é•‘}•áÑ•¹ÑÌ¹¥Ñ•µÌ ¤4(€€€€€€€€€€€€¤4(€€€€€€€€¤4(€€€€€€€Á…¹•±}¡•¥¡ÑÌ€ôl4(€€€€€€€€€€€¥¹Ð¡•áÑ•¹Ð¹•Ð ‰äˆ¤¤4(€€€€€€€€€€€™½È¥µ…•}¥°•áÑ•¹Ð¥¸É•Í¥é•‘}•áÑ•¹ÑÌ¹¥Ñ•µÌ ¤4(€€€€€€€€€€€¥˜¹•áÐ 4(€€€€€€€€€€€€€€€¥µ…”4(€€€€€€€€€€€€€€€™½È¥µ…”¥¸ÍÑÉÕÑÕÉ•l‰¥µ…•Ì‰t4(€€€€€€€€€€€€€€€¥˜¥µ…•l‰¥µ…”‰t€ôô¥µ…•}¥4(€€€€€€€€€€€€¥l‰Á±…•µ•¹Ð‰t4(€€€€€€€€€€€€ôô€‰Ñ…‰±•}™¥ÕÉ•}Á…¹•°ˆ4(€€€€€€€t4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° Ä°±•¸¡Í•Ð¡Á…¹•±}¡•¥¡ÑÌ¤¤¤4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° 4(€€€€€€€€€€€ÍÑÉÕÑÕÉ•}½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡Í¡¥™Ñ•‘}Í½ÕÉ”°ÍÑÉÕÑÕÉ”¤°4(€€€€€€€€€€€ÍÑÉÕÑÕÉ•}½¹Ñ•¹Ñ}™¥¹•ÉÁÉ¥¹Ð¡½ÕÑÁÕÐ°ÍÑÉÕÑÕÉ”¤°4(€€€€€€€€¤4(4(€€€‘•˜Ñ•ÍÑ}¡•…‘¥¹}¹Õµ‰•É¥¹}¥¹¡•É¥ÑÍ}µ¥á•‘}™½¹ÑÍ}Í¥é•}Ý•¥¡Ñ}…¹‘}é•É½}¥¹‘•¹Ð¡Í•±˜¤€´ø9½¹”è(€€€€€€€‘½Õµ•¹Ð€ô½Õµ•¹Ð ¤4(€€€€€€€ÍÁ•¥™¥…Ñ¥½¹Ì€ô€ 4(€€€€€€€€€€€€ ‰!•…‘¥¹œ€Äˆ°€‹¦îG’öLˆ°€‰Q¥µ•Ì9•ÜI½µ…¸ˆ°€Äà¤°4(€€€€€€€€€€€€ ‰!•…‘¥¹œ€Èˆ°€‹¦îG’öLˆ°€‰Q¥µ•Ì9•ÜI½µ…¸ˆ°€ÄÐ¤°4(€€€€€€€€€€€€ ‰!•…‘¥¹œ€Ìˆ°€‹¦îG’öLˆ°€‰Q¥µ•Ì9•ÜI½µ…¸ˆ°€ÄÈ¤°4(€€€€€€€€€€€€ ‰!•…‘¥¹œ€Ðˆ°€‹–º/’öLˆ°€‰Q¥µ•Ì9•ÜI½µ…¸ˆ°€ÄÀ¸Ô¤°4(€€€€€€€€¤4(€€€€€€€™½ÈÍÑå±•}¹…µ”°•…ÍÑ}…Í¥„°…Í¥¥}¹…µ”°Í¥é”¥¸ÍÁ•¥™¥…Ñ¥½¹Ìè4(€€€€€€€€€€€ÍÑå±”€ô•¹ÍÕÉ•}Á…É…É…Á¡}ÍÑå±”¡‘½Õµ•¹Ð°ÍÑå±•}¹…µ”¤4(€€€€€€€€€€€…ÁÁ±å}ÍÑå±•}ÁÉ½Á•ÉÑ¥•Ì 4(€€€€€€€€€€€€€€€ÍÑå±”°4(€€€€€€€€€€€€€€€ì4(€€€€€€€€€€€€€€€€€€€€‰™½¹Ñ}¹…µ•}•…ÍÑ}…Í¥„ˆè•…ÍÑ}…Í¥„°4(€€€€€€€€€€€€€€€€€€€€‰™½¹Ñ}¹…µ•}…Í¥¤ˆè…Í¥¥}¹…µ”°4(€€€€€€€€€€€€€€€€€€€€‰™½¹Ñ}¹…µ•}½µÁ±•á}ÍÉ¥ÁÐˆè…Í¥¥}¹…µ”°4(€€€€€€€€€€€€€€€€€€€€‰™½¹Ñ}Í¥é•}ÁÐˆèÍ¥é”°4(€€€€€€€€€€€€€€€€€€€€‰‰½±ˆèQÉÕ”°4(€€€€€€€€€€€€€€€€€€€€‰™¥ÉÍÑ}±¥¹•}¥¹‘•¹Ñ}¡…ÉÌˆè€È°4(€€€€€€€€€€€€€€€ô°4(€€€€€€€€€€€€¤4(€€€€€€€€€€€Á…É…É…Á €ô‘½Õµ•¹Ð¹…‘‘}Á…É…É…Á  ‰Må¹Ñ¡•Ñ¥Œ¡•…‘¥¹œˆ°ÍÑå±”õÍÑå±•}¹…µ”¤(€€€€€€€€€€€‘¥É•Ñ}¥¹€ôÁ…É…É…Á ¹}À¹•Ñ}½É}…‘‘}ÁAÈ ¤¹•Ñ}½É}…‘‘}¥¹ ¤(€€€€€€€€€€€™½È…ÑÑÉ¥‰ÕÑ”°Ù…±Õ”¥¸€ (€€€€€€€€€€€€€€€€ ‰±•™Ðˆ°€ˆÜÈÀˆ¤°(€€€€€€€€€€€€€€€€ ‰±•™Ñ¡…ÉÌˆ°€ˆÈÀÀˆ¤°(€€€€€€€€€€€€€€€€ ‰É¥¡Ðˆ°€ˆÌØÀˆ¤°(€€€€€€€€€€€€€€€€ ‰É¥¡Ñ¡…ÉÌˆ°€ˆÄÀÀˆ¤°(€€€€€€€€€€€€€€€€ ‰™¥ÉÍÑ1¥¹”ˆ°€ˆÈÐÀˆ¤°(€€€€€€€€€€€€€€€€ ‰¡…¹¥¹¡…ÉÌˆ°€ˆÄÀÀˆ¤°(€€€€€€€€€€€€¤è(€€€€€€€€€€€€€€€‘¥É•Ñ}¥¹¹Í•Ð¡Å¸¡˜‰Üéí…ÑÑÉ¥‰ÕÑ•ôˆ¤°Ù…±Õ”¤(€€€€€€€€€€€‘¥É•Ñ}¹Õµ}ÁÈ€ô=áµ±±•µ•¹Ð ‰Üé¹ÕµAÈˆ¤(€€€€€€€€€€€Á…É…É…Á ¹}À¹•Ñ}½É}…‘‘}ÁAÈ ¤¹…ÁÁ•¹¡‘¥É•Ñ}¹Õµ}ÁÈ¤(4(€€€€€€€ÁÉ½Á•ÉÑ¥•Ì€ôì4(€€€€€€€€€€€€‰É•‰Õ¥±‘}¡•…‘¥¹}¹Õµ‰•É¥¹œˆèQÉÕ”°4(€€€€€€€€€€€€‰¡•…‘¥¹}±•Ù•±Ìˆè€Ð°4(€€€€€€€€€€€€‰¡…ÁÑ•É}ÍÑ…ÉÐˆè€Ð°4(€€€€€€€ô4(€€€€€€€…ÁÁ±å}™¥•±‘}ÁÉ½Á•ÉÑ¥•Ì¡‘½Õµ•¹Ð°ÁÉ½Á•ÉÑ¥•Ì¤4(€€€€€€€½ÕÑÁÕÐ€ôÍ•±˜¹É½½Ð€¼€‰¡•…‘¥¹œµ¹Õµ‰•É¥¹œµ™½Éµ…Ð¹‘½àˆ4(€€€€€€€‘½Õµ•¹Ð¹Í…Ù”¡½ÕÑÁÕÐ¤4(€€€€€€€É•±½…‘•€ô½Õµ•¹Ð¡½ÕÑÁÕÐ¤4(4(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ…±Í” 4(€€€€€€€€€€€…Õ‘¥Ñ}™¥•±‘}ÉÕ±” 4(€€€€€€€€€€€€€€€É•±½…‘•°4(€€€€€€€€€€€€€€€ì‰ÁÉ½Á•ÉÑ¥•ÌˆèÁÉ½Á•ÉÑ¥•Íô°4(€€€€€€€€€€€€€€€¡…ÁÑ•É}ÍÑ…ÉÐôÐ°4(€€€€€€€€€€€€€€€Á…Ñ õ½ÕÑÁÕÐ°4(€€€€€€€€€€€€¤4(€€€€€€€€¤4(€€€€€€€™½ÈÁ…É…É…Á ¥¸É•±½…‘•¹Á…É…É…Á¡Ìè(€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%Í9½¹”¡Á…É…É…Á ¹Á…É…É…Á¡}™½Éµ…Ð¹™¥ÉÍÑ}±¥¹•}¥¹‘•¹Ð¤(€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%Í9½¹”¡Á…É…É…Á ¹}À¹ÁAÈ¹™¥¹¡Å¸ ‰Üé¹ÕµAÈˆ¤¤¤(€€€€€€€€€€€‘¥É•Ñ}¥¹€ôÁ…É…É…Á ¹}À¹ÁAÈ¹™¥¹¡Å¸ ‰Üé¥¹ˆ¤¤(€€€€€€€€€€€¥˜‘¥É•Ñ}¥¹¥Ì¹½Ð9½¹”è(€€€€€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑ…±Í”¡‘¥É•Ñ}¥¹¹…ÑÑÉ¥ˆ¤(€€€€€€€™½ÈÍÑå±•}¹…µ”°}•…ÍÑ}…Í¥„°}…Í¥¥}¹…µ”°}Í¥é”¥¸ÍÁ•¥™¥…Ñ¥½¹Ìè(€€€€€€€€€€€¥¹€ôÉ•±½…‘•¹ÍÑå±•ÍmÍÑå±•}¹…µ•t¹•±•µ•¹Ð¹ÁAÈ¹™¥¹¡Å¸ ‰Üé¥¹ˆ¤¤(€€€€€€€€€€€™½È…ÑÑÉ¥‰ÕÑ”¥¸€ (€€€€€€€€€€€€€€€€‰±•™Ðˆ°(€€€€€€€€€€€€€€€€‰±•™Ñ¡…ÉÌˆ°(€€€€€€€€€€€€€€€€‰É¥¡Ðˆ°(€€€€€€€€€€€€€€€€‰É¥¡Ñ¡…ÉÌˆ°(€€€€€€€€€€€€€€€€‰™¥ÉÍÑ1¥¹”ˆ°(€€€€€€€€€€€€€€€€‰™¥ÉÍÑ1¥¹•¡…ÉÌˆ°(€€€€€€€€€€€€¤è(€€€€€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑÅÕ…° ˆÀˆ°¥¹¹•Ð¡Å¸¡˜‰Üéí…ÑÑÉ¥‰ÕÑ•ôˆ¤¤¤((€€€€€€€™½È±•Ù•°¥¸É…¹” Ð¤è(€€€€€€€€€€€¹Õµ‰•É¥¹}±•Ù•°€ôÉ•±½…‘•¹ÍÑå±•Ím˜‰!•…‘¥¹œí±•Ù•°€¬€Åô‰t¹•±•µ•¹Ð¹ÁAÈ¹™¥¹ (€€€€€€€€€€€€€€€Å¸ ‰Üé¹ÕµAÈˆ¤(€€€€€€€€€€€€¤(€€€€€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%Í9½Ñ9½¹”¡¹Õµ‰•É¥¹}±•Ù•°¤(€€€€€€€ÍÑÉÕÑÕÉ”€ôì(€€€€€€€€€€€€‰Í¡•µ…}Ù•ÉÍ¥½¸ˆè€ˆÄ¸Ôˆ°(€€€€€€€€€€€€‰¹Õµ‰•É¥¹œˆèì‰…ÁÁÉ½Ù•ˆèQÉÕ•ô°(€€€€€€€€€€€€‰¡•…‘¥¹Ìˆèl(€€€€€€€€€€€€€€€ì(€€€€€€€€€€€€€€€€€€€€‰…ÁÁÉ½Ù•ˆèQÉÕ”°(€€€€€€€€€€€€€€€€€€€€‰±•Ù•°ˆè±•Ù•°°(€€€€€€€€€€€€€€€€€€€€‰Á…É…É…Á ˆè±•Ù•°€´€Ä°(€€€€€€€€€€€€€€€€€€€€‰Ñ•áÑ}Í¡„ÈÔØˆèÑ•áÑ}Í¡„ÈÔØ¡É•±½…‘•¹Á…É…É…Á¡Ím±•Ù•°€´€Åt¹Ñ•áÐ¤°(€€€€€€€€€€€€€€€ô(€€€€€€€€€€€€€€€™½È±•Ù•°¥¸É…¹” Ä°€Ô¤(€€€€€€€€€€€t°(€€€€€€€ô(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ…±Í”¡…Õ‘¥Ñ}ÍÑÉÕÑÕÉ•}¡•…‘¥¹}½Á•É…Ñ¥½¹Ì¡É•±½…‘•°ÍÑÉÕÑÕÉ”¤¤((€€€‘•˜Ñ•ÍÑ}Õ¹…ÁÁÉ½Ù•‘}‘¥É•Ñ}¡•…‘¥¹}¹Õµ‰•É¥¹}¥Í}ÁÉ•Í•ÉÙ•‘}…¹‘}É•Á½ÉÑ•¡Í•±˜¤€´ø9½¹”è(€€€€€€€‘½Õµ•¹Ð€ô½Õµ•¹Ð ¤(€€€€€€€Á…É…É…Á €ô‘½Õµ•¹Ð¹…‘‘}Á…É…É…Á  ˆÄ¸ÄMå¹Ñ¡•Ñ¥Œ¡•…‘¥¹œˆ°ÍÑå±”ô‰!•…‘¥¹œ€Èˆ¤(€€€€€€€…ÁÁ±å}™¥•±‘}ÁÉ½Á•ÉÑ¥•Ì (€€€€€€€€€€€‘½Õµ•¹Ð°(€€€€€€€€€€€ì(€€€€€€€€€€€€€€€€‰É•‰Õ¥±‘}¡•…‘¥¹}¹Õµ‰•É¥¹œˆèQÉÕ”°(€€€€€€€€€€€€€€€€‰¡•…‘¥¹}±•Ù•±Ìˆè€Ð°(€€€€€€€€€€€€€€€€‰¡…ÁÑ•É}ÍÑ…ÉÐˆè€Ä°(€€€€€€€€€€€ô°(€€€€€€€€¤(€€€€€€€ÍÑå±•}¹Õµ}ÁÈ€ô‘½Õµ•¹Ð¹ÍÑå±•Íl‰!•…‘¥¹œ€È‰t¹•±•µ•¹Ð¹ÁAÈ¹™¥¹¡Å¸ ‰Üé¹ÕµAÈˆ¤¤(€€€€€€€Á…É…É…Á ¹}À¹•Ñ}½É}…‘‘}ÁAÈ ¤¹…ÁÁ•¹¡½Áä¹‘••Á½Áä¡ÍÑå±•}¹Õµ}ÁÈ¤¤(€€€€€€€¹Õµ}¥€ôÍÑå±•}¹Õµ}ÁÈ¹™¥¹¡Å¸ ‰Üé¹Õµ%ˆ¤¤¹•Ð¡Å¸ ‰ÜéÙ…°ˆ¤¤(€€€€€€€…‰ÍÑÉ…Ñ}¥€ô¹•áÐ (€€€€€€€€€€€¥Ñ•´(€€€€€€€€€€€™½È¥Ñ•´¥¸‘½Õµ•¹Ð¹Á…ÉÐ¹¹Õµ‰•É¥¹}Á…ÉÐ¹•±•µ•¹Ð¹™¥¹‘…±°¡Å¸ ‰Üé¹Õ´ˆ¤¤(€€€€€€€€€€€¥˜¥Ñ•´¹•Ð¡Å¸ ‰Üé¹Õµ%ˆ¤¤€ôô¹Õµ}¥(€€€€€€€€¤¹™¥¹¡Å¸ ‰Üé…‰ÍÑÉ…Ñ9Õµ%ˆ¤¤¹•Ð¡Å¸ ‰ÜéÙ…°ˆ¤¤(€€€€€€€…‰ÍÑÉ…Ð€ô¹•áÐ (€€€€€€€€€€€¥Ñ•´(€€€€€€€€€€€™½È¥Ñ•´¥¸‘½Õµ•¹Ð¹Á…ÉÐ¹¹Õµ‰•É¥¹}Á…ÉÐ¹•±•µ•¹Ð¹™¥¹‘…±°¡Å¸ ‰Üé…‰ÍÑÉ…Ñ9Õ´ˆ¤¤(€€€€€€€€€€€¥˜¥Ñ•´¹•Ð¡Å¸ ‰Üé…‰ÍÑÉ…Ñ9Õµ%ˆ¤¤€ôô…‰ÍÑÉ…Ñ}¥(€€€€€€€€¤(€€€€€€€±•Ù•°€ô¹•áÐ (€€€€€€€€€€€¥Ñ•´(€€€€€€€€€€€™½È¥Ñ•´¥¸…‰ÍÑÉ…Ð¹™¥¹‘…±°¡Å¸ ‰Üé±Ù°ˆ¤¤(€€€€€€€€€€€¥˜¥Ñ•´¹•Ð¡Å¸ ‰Üé¥±Ù°ˆ¤¤€ôô€ˆÄˆ(€€€€€€€€¤(€€€€€€€±•Ù•°¹™¥¹¡Å¸ ‰ÜéÁAÈˆ¤¤¹™¥¹¡Å¸ ‰Üé¥¹ˆ¤¤¹Í•Ð¡Å¸ ‰Üé±•™Ðˆ¤°€ˆÜÈÀˆ¤(€€€€€€€ÍÑÉÕÑÕÉ”€ôì(€€€€€€€€€€€€‰Í¡•µ…}Ù•ÉÍ¥½¸ˆè€ˆÄ¸Ôˆ°(€€€€€€€€€€€€‰¹Õµ‰•É¥¹œˆèì‰…ÁÁÉ½Ù•ˆè…±Í•ô°(€€€€€€€€€€€€‰¡•…‘¥¹Ìˆèl(€€€€€€€€€€€€€€€ì(€€€€€€€€€€€€€€€€€€€€‰…ÁÁÉ½Ù•ˆèQÉÕ”°(€€€€€€€€€€€€€€€€€€€€‰±•Ù•°ˆè€È°(€€€€€€€€€€€€€€€€€€€€‰Á…É…É…Á ˆè€À°(€€€€€€€€€€€€€€€€€€€€‰Ñ•áÑ}Í¡„ÈÔØˆèÑ•áÑ}Í¡„ÈÔØ¡Á…É…É…Á ¹Ñ•áÐ¤°(€€€€€€€€€€€€€€€ô(€€€€€€€€€€€t°(€€€€€€€ô(€€€€€€€…ÁÁ±å}ÍÑÉÕÑÕÉ•}µ…À¡‘½Õµ•¹Ð°ÍÑÉÕÑÕÉ”¤(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%Í9½Ñ9½¹”¡Á…É…É…Á ¹}À¹ÁAÈ¹™¥¹¡Å¸ ‰Üé¹ÕµAÈˆ¤¤¤(€€€€€€€™…¥±ÕÉ•Ì€ô…Õ‘¥Ñ}ÍÑÉÕÑÕÉ•}¡•…‘¥¹}½Á•É…Ñ¥½¹Ì¡‘½Õµ•¹Ð°ÍÑÉÕÑÕÉ”¤(€€€€€€€Í•±˜¹…ÍÍ•ÉÑ%¸ (€€€€€€€€€€€€‰¡•…‘¥¹}‘¥É•Ñ}¹Õµ‰•É¥¹}É•ÅÕ¥É•Í}Å„ˆ°(€€€€€€€€€€€í™…¥±ÕÉ•l‰ÁÉ½Á•ÉÑä‰t™½È™…¥±ÕÉ”¥¸™…¥±ÕÉ•Íô°(€€€€€€€€¤(4(4)¥˜}}¹…µ•}|€ôô€‰}}µ…¥¹}|ˆè4(€€€Õ¹¥ÑÑ•ÍÐ¹µ…¥¸ ¤4