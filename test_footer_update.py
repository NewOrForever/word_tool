#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import tempfile
import unittest
import zipfile
from base64 import b64decode
from pathlib import Path
from xml.etree import ElementTree as ET

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

    # 验证修改已有页脚时内容会右对齐
    def test_updates_existing_footer_text_on_right_side(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "has_footer.docx"
            create_sample_docx(source, include_header=False, include_footer=True)

            result = process_docx(source, "set_footer", None, True, footer_text="右侧页脚")

            self.assertTrue(result.changed)
            self.assert_footer_is_right_aligned(source)

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

    # 验证新增页脚时内容会右对齐
    def test_adds_footer_text_on_right_side(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "no_footer.docx"
            create_sample_docx(source, include_header=False, include_footer=False)

            result = process_docx(source, "set_footer", None, True, footer_text="右侧页脚")

            self.assertTrue(result.changed)
            self.assert_footer_is_right_aligned(source)

    # 验证替换页脚文字后会清理旧页脚关联文件
    def test_updates_footer_text_removes_old_footer_relationships(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "has_footer_rels.docx"
            create_sample_docx(source, include_header=False, include_footer=True, include_footer_rels=True)

            result = process_docx(source, "set_footer", None, True, footer_text="新的页脚")

            self.assertTrue(result.changed)
            with zipfile.ZipFile(source) as zf:
                self.assertNotIn("word/_rels/footer1.xml.rels", zf.namelist())

    # 验证替换页脚后会删除不再使用的旧页脚图片
    def test_updates_footer_text_removes_unreferenced_footer_image(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "obsolete_footer_image.docx"
            create_sample_docx(source, include_header=False, include_footer=True, include_footer_rels=True)

            result = process_docx(source, "set_footer", None, True, footer_text="新的页脚")

            self.assertTrue(result.changed)
            with zipfile.ZipFile(source) as zf:
                self.assertNotIn("word/media/old-footer.png", zf.namelist())

    # 验证正文仍在使用的图片不会随旧页脚一起删除
    def test_updates_footer_text_preserves_image_used_by_document(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "shared_footer_image.docx"
            create_sample_docx(
                source,
                include_header=False,
                include_footer=True,
                include_footer_rels=True,
                share_footer_image_with_document=True,
            )

            result = process_docx(source, "set_footer", None, True, footer_text="新的页脚")

            self.assertTrue(result.changed)
            with zipfile.ZipFile(source) as zf:
                self.assertIn("word/media/old-footer.png", zf.namelist())

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

    # 验证页眉可以选择左中右三种位置
    def test_header_alignment_can_be_selected(self) -> None:
        for alignment in ("left", "center", "right"):
            with self.subTest(alignment=alignment), tempfile.TemporaryDirectory() as temp_dir:
                source = Path(temp_dir) / "header_alignment.docx"
                create_sample_docx(source, include_header=True, include_footer=False)

                result = process_docx(
                    source,
                    "set_header",
                    None,
                    True,
                    header_text="页眉位置",
                    header_alignment=alignment,
                )

                self.assertTrue(result.changed)
                self.assert_part_alignment(source, "word/header1.xml", alignment)

    # 验证页脚可以选择左中右三种位置
    def test_footer_alignment_can_be_selected(self) -> None:
        for alignment in ("left", "center", "right"):
            with self.subTest(alignment=alignment), tempfile.TemporaryDirectory() as temp_dir:
                source = Path(temp_dir) / "footer_alignment.docx"
                create_sample_docx(source, include_header=False, include_footer=True)

                result = process_docx(
                    source,
                    "set_footer",
                    None,
                    True,
                    footer_text="页脚位置",
                    footer_alignment=alignment,
                )

                self.assertTrue(result.changed)
                self.assert_part_alignment(source, "word/footer1.xml", alignment)

    # 验证页眉和页脚可以分别选择位置
    def test_header_and_footer_alignments_are_independent(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "independent_alignments.docx"
            create_sample_docx(source, include_header=True, include_footer=True)

            result = process_docx(
                source,
                "set_both",
                None,
                True,
                header_text="居中页眉",
                footer_text="左侧页脚",
                header_alignment="center",
                footer_alignment="left",
            )

            self.assertTrue(result.changed)
            self.assert_part_alignment(source, "word/header1.xml", "center")
            self.assert_part_alignment(source, "word/footer1.xml", "left")

    # 判断压缩包内指定文件包含目标文本
    def assert_zip_contains(self, source: Path, name: str, text: str) -> None:
        with zipfile.ZipFile(source) as zf:
            self.assertIn(text, zf.read(name).decode("utf-8"))

    # 判断压缩包内指定文件不包含目标文本
    def assert_zip_not_contains(self, source: Path, name: str, text: str) -> None:
        with zipfile.ZipFile(source) as zf:
            self.assertNotIn(text, zf.read(name).decode("utf-8"))

    # 判断页脚段落是否为右对齐
    def assert_footer_is_right_aligned(self, source: Path) -> None:
        self.assert_part_alignment(source, "word/footer1.xml", "right")

    # 判断指定页眉页脚段落是否使用目标位置
    def assert_part_alignment(self, source: Path, name: str, expected: str) -> None:
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        with zipfile.ZipFile(source) as zf:
            root = ET.fromstring(zf.read(name))
        alignment = root.find(".//w:p/w:pPr/w:jc", ns)
        self.assertIsNotNone(alignment)
        self.assertEqual(expected, alignment.attrib.get(f"{{{ns['w']}}}val"))


# 创建测试用的最小 Word 文件
def create_sample_docx(
    path: Path,
    include_header: bool,
    include_footer: bool,
    include_footer_rels: bool = False,
    share_footer_image_with_document: bool = False,
) -> None:
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
    document_image = (
        '    <w:p><w:r><w:pict><v:shape><v:imagedata r:id="rIdBodyImage"/></v:shape></w:pict></w:r></w:p>\n'
        if share_footer_image_with_document
        else ""
    )
    rels_document_image = (
        '  <Relationship Id="rIdBodyImage" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/old-footer.png"/>\n'
        if share_footer_image_with_document
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <w:p><w:r><w:t>正文内容</w:t></w:r></w:p>
{document_image}
    <w:sectPr>
{document_header}{document_footer}      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>""",
        "word/_rels/document.xml.rels": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels_header}{rels_footer}{rels_document_image}</Relationships>""",
    }
    if include_header:
        files["word/header1.xml"] = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>旧页眉</w:t></w:r></w:p></w:hdr>"""
    if include_footer:
        files["word/footer1.xml"] = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>旧页脚</w:t></w:r></w:p></w:ftr>"""
    if include_footer_rels:
        files["word/_rels/footer1.xml.rels"] = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="file:///D:/old/image.png" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/old-footer.png"/>
</Relationships>"""
        files["word/media/old-footer.png"] = b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            data = content if isinstance(content, bytes) else content.encode("utf-8")
            zf.writestr(name, data)


if __name__ == "__main__":
    unittest.main()
