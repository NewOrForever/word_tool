import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.dom import minidom

from remove_word_headers_footers import WORD_NS, build_header_footer_xml, process_docx, ensure_header_footer_reference, remove_section_references
from test_footer_update import create_sample_docx


class PageNumberTests(unittest.TestCase):
    def test_section_edits_preserve_compatibility_prefixes(self):
        source = (f'<w:document xmlns:w="{WORD_NS}" '
                  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
                  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
                  'mc:Ignorable="w14"><w:body><w:p/><w:sectPr/></w:body></w:document>').encode()
        added, changed = ensure_header_footer_reference(source, "footer", "rIdFooter")
        self.assertTrue(changed)
        for data in (added, remove_section_references(added, "footer")[0]):
            root = minidom.parseString(data).documentElement
            self.assertEqual("http://schemas.microsoft.com/office/word/2010/wordml", root.getAttribute("xmlns:w14"))
            self.assertEqual("w14", root.getAttribute("mc:Ignorable"))

    def test_page_fields_with_text_and_independent_alignment(self):
        for existing in (True, False):
            with self.subTest(existing=existing), tempfile.TemporaryDirectory() as folder:
                source = Path(folder) / "pages.docx"
                create_sample_docx(source, include_header=True, include_footer=existing)
                with zipfile.ZipFile(source) as archive:
                    header_before = archive.read("word/header1.xml")
                result = process_docx(
                    source, "set_footer", None, True,
                    footer_text="第 {PAGE} 页 / 共 {NUMPAGES} 页",
                    footer_alignment="center",
                )
                self.assertTrue(result.changed)
                with zipfile.ZipFile(source) as archive:
                    self.assertEqual(header_before, archive.read("word/header1.xml"))
                    root = ET.fromstring(archive.read("word/footer1.xml"))
                instructions = [node.text.strip() for node in root.iter(f"{{{WORD_NS}}}instrText")]
                self.assertEqual(["PAGE", "NUMPAGES"], instructions)
                marks = [node.get(f"{{{WORD_NS}}}fldCharType") for node in root.iter(f"{{{WORD_NS}}}fldChar")]
                self.assertEqual(["begin", "separate", "end"] * 2, marks)
                self.assertEqual("center", root.find(f".//{{{WORD_NS}}}jc").get(f"{{{WORD_NS}}}val"))
                self.assertNotIn("{PAGE}", ET.tostring(root, encoding="unicode"))
                self.assertFalse(process_docx(
                    source, "set_footer", None, True,
                    footer_text="第 {PAGE} 页 / 共 {NUMPAGES} 页", footer_alignment="center",
                ).changed)

    def test_plain_text_and_unknown_tokens_are_preserved(self):
        value = "  文本 & <标记> {UNKNOWN}  "
        root = ET.fromstring(build_header_footer_xml("footer", value, "right"))
        self.assertEqual(value, "".join(node.text or "" for node in root.iter(f"{{{WORD_NS}}}t")))
        self.assertEqual([], list(root.iter(f"{{{WORD_NS}}}instrText")))

    def test_adjacent_and_repeated_fields(self):
        root = ET.fromstring(build_header_footer_xml("footer", "{PAGE}{NUMPAGES}{PAGE}", "left"))
        self.assertEqual(["PAGE", "NUMPAGES", "PAGE"], [node.text.strip() for node in root.iter(f"{{{WORD_NS}}}instrText")])


if __name__ == "__main__":
    unittest.main()
