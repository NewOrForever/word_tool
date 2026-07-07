#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from remove_word_headers_footers import process_docx


class HeaderFooterUpdateTests(unittest.TestCase):
    # 验证已有页眉内容会被替换
    def test_updates_existing_header_text(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "has_header.docx"
            create_sample_docx(source, include_header=True, include_footer=False)

            result = process_docx(source, "set_header", None, True, header_text="新的页眉")

            self.assertTrue(result.changed)
            self.assert_zip_contains(source, "word/header1.xml", "新的页眉")
            self.assert_zip_not_contains(source, "word/header1.xml", "旧页眉")
            self.assert_zip_contains(source, "word/document.xml", "headerReference")

    # 验证没有页眉时会自动新增页眉
    def test_adds_header_when_document_has_none(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "no_header.docx"
            create_sample_docx(source, include_header=False, include_footer=False)

            result = process_docx(source, "set_header", None, True, header_text="新增页眉")

            self.assertTrue(result.changed)
            self.assert_zip_contains(source, "word/header1.xml", "新增页眉")
            self.assert_zip_contains(source, "word/document.xml", "headerReference")
            self.assert_zip_contains(source, "word/_rels/document.xml.rels", "header1.xml")
            self.assert_zip_contains(source, "[Content_Types].xml", "/word/header1.xml")

    # 验证已有页脚内容会被替换
    def test_updates_existing_footer_text(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "has_footer.docx"
            create_sample_docx(source, include_header=False, include_footer=True)

            result = process_docx(source, "set_footer", None, True, footer_text="新的页脚")

            self.assertTrue(result.changed)
            self.assert_zip_contains(source, "word/footer1.xml", "新的页脚")
            self.assert_zip_not_contains(source, "word/footer1.xml", "旧页脚")
            self.assert_zip_contains(source, "word/document.xml", "footerReference")

    # 验证没有页脚时会自动新增页脚
    def test_adds_footer_when_document_has_none(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "no_footer.docx"
            create_sample_docx(source, include_header=False, include_footer=False)

            result = process_docx(source, "set_footer", None, True, footer_text="新增页脚")

            self.assertTrue(result.changed)
            self.assert_zip_contains(source, "word/footer1.xml", "新增页脚")
            self.assert_zip_contains(source, "word/document.xml", "footerReference")
            self.assert_zip_contains(source, "word/_rels/document.xml.rels", "footer1.xml")
            self.assert_zip_contains(source, "[Content_Types].xml", "/word/footer1.xml")

    # 验证可以一次修改页眉和页脚
    def test_updates_header_and_footer_together(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "both.docx"
            create_sample_docx(source, include_header=True, include_footer=True)

            result = process_docx(
                source,
                "set_both",
                None,
                True,
                header_text="一起改页眉",
                footer_text="一起改页脚",
            )

            self.assertTrue(result.changed)
            self.assert_zip_contains(source, "word/header1.xml", "一起改页眉")
            self.assert_zip_contains(source, "word/footer1.xml", "一起改页脚")

    # 判断压缩包内指定文件包含目标文本
    def assert_zip_contains(self, source: Path, name: str, text: str) -> None:
        with zipfile.ZipFile(source) as zf:
            self.assertIn(text, zf.read(name).decode("utf-8"))

    # 判断压缩包内指定文件不包含目标文本
    def assert_zip_not_contains(self, source: Path, name: str, text: str) -> None:
        with zipfile.ZipFile(source) as zf:
            self.assertNotIn(text, zf.read(name).decode("utf-8"))


# 创建测试用的最小 Word 文件
def create_sample_docx(path: Path, include_header: bool, include_footer: bool) -> None:
    content_types_header = (
        '  <Override PartName="/word/header1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\n'
        if include_header
        else ""
    )
    content_types_footer = (
        '  <Override PartName="/word/footer1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>\n'
        if include_footer
        else ""
    )
    document_header = '      <w:headerReference w:type="default" r:id="rIdHeader"/>\n' if include_header else ""
    document_footer = '      <w:footerReference w:type="default" r:id="rIdFooter"/>\n' if include_footer else ""
    rels_header = (
        '  <Relationship Id="rIdHeader" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header1.xml"/>\n'
        if include_header
        else ""
    )
    rels_footer = (
        '  <Relationship Id="rIdFooter" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer1.xml"/>\n'
        if include_footer
        else ""
    )

    files = {
        "[Content_Types].xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
{content_types_header}{content_types_footer}</Types>""",
        "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""",
        "word/document.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>正文内容</w:t></w:r></w:p>
    <w:sectPr>
{document_header}{document_footer}      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>""",
        "word/_rels/document.xml.rels": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels_header}{rels_footer}</Relationships>""",
    }
    if include_header:
        files["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>旧页眉</w:t></w:r></w:p></w:hdr>"""
    if include_footer:
        files["word/footer1.xml"] = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>旧页脚</w:t></w:r></w:p></w:ftr>"""

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content.encode("utf-8"))


if __name__ == "__main__":
    unittest.main()
