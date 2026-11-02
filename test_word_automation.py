import os
import queue
import tempfile
import threading
import unittest
import zipfile
from base64 import b64decode
from pathlib import Path
from unittest.mock import patch
from xml.dom import minidom

from remove_word_headers_footers import collect_targets, process_docx


class WordFormatTests(unittest.TestCase):
    def test_folder_collects_legacy_and_modern_word_files(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            for name in ('old.DOC', 'new.docx', 'macro.docm', 'other.txt'):
                (root / name).touch()
            self.assertEqual({'old.DOC', 'new.docx', 'macro.docm'},
                             {p.name for p in collect_targets(root, False)})

    def test_missing_office_does_not_modify_source_or_leave_temporary_files(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'old.doc'
            source.write_bytes(b'original document')
            with patch('word_automation.office_application', side_effect=RuntimeError('未安装 Word')):
                result = process_docx(source, 'both', None, True)
            self.assertFalse(result.changed)
            self.assertIn('未安装 Word', result.message)
            self.assertEqual(b'original document', source.read_bytes())
            self.assertEqual([source], list(Path(folder).iterdir()))


@unittest.skipUnless(os.environ.get('WORD_INTEGRATION_TESTS') == '1',
                     '设置 WORD_INTEGRATION_TESTS=1 运行本机 Word 集成测试')
class WordIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import pythoncom
        from win32com.client import DispatchEx
        pythoncom.CoInitialize()
        cls.word = DispatchEx('Word.Application')
        cls.word.Visible = False
        cls.word.DisplayAlerts = 0
        cls.word.AutomationSecurity = 3

    @classmethod
    def tearDownClass(cls):
        import pythoncom
        cls.word.Quit(0)
        cls.word = None
        pythoncom.CoUninitialize()

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.picture = self.folder / 'pixel.png'
        self.picture.write_bytes(b64decode(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='))

    def create_document(self, suffix, automatic=False, pictures=True, restart=False):
        from docx import Document
        from docx.enum.section import WD_SECTION_START
        from docx.shared import Pt
        path = self.folder / (f'sample-{automatic}-{pictures}-{restart}' + suffix)
        modern_path = path.with_suffix('.docx')
        fixture = Document()
        first = fixture.add_paragraph('First page')
        if pictures:
            first.add_run().add_picture(str(self.picture), width=Pt(18), height=Pt(18))
        if automatic:
            for _ in range(95):
                fixture.add_paragraph('body line')
        elif restart:
            fixture.add_section(WD_SECTION_START.NEW_PAGE)
        else:
            fixture.add_page_break()
        last = fixture.add_paragraph('Last page')
        if pictures:
            last.add_run().add_picture(str(self.picture), width=Pt(18), height=Pt(18))
        end = fixture.add_paragraph('End text')
        if pictures:
            end.add_run().add_picture(str(self.picture), width=Pt(18), height=Pt(18))
        for section in fixture.sections:
            section.header.paragraphs[0].text = 'Keep header'
            section.footer.paragraphs[0].text = 'Old footer'
        fixture.save(modern_path)
        doc = self.word.Documents.Open(str(modern_path))
        try:
            if pictures:
                # 浮动图片与普通形状共存，普通形状不能被误删。
                anchor = doc.Paragraphs.Last.Range
                shape = doc.InlineShapes(doc.InlineShapes.Count).ConvertToShape()
                shape.Name = 'last-picture'
                shape.WrapFormat.Type = 3
                doc.Shapes.AddShape(1, 80, 30, 18, 18, anchor).Name = 'keep-rectangle'
            if suffix == '.doc':
                doc.SaveAs2(str(path), 0)
            else:
                doc.Save()
            doc.Repaginate()
            self.assertGreaterEqual(doc.ComputeStatistics(2), 2)
        finally:
            doc.Close(0)
        if suffix == '.doc':
            modern_path.unlink()
        if restart:
            # 直接设分节页码，避免 Word 对尚无 PAGE 域的 PageNumbers 集合阻塞。
            with zipfile.ZipFile(path) as archive:
                parts = {name: archive.read(name) for name in archive.namelist()}
            ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            root = minidom.parseString(parts['word/document.xml'])
            for section in root.getElementsByTagNameNS(ns, 'sectPr'):
                numbering = root.createElementNS(ns, 'w:pgNumType')
                numbering.setAttributeNS(ns, 'w:start', '1')
                columns = section.getElementsByTagNameNS(ns, 'cols')
                section.insertBefore(numbering, columns[0] if columns else None)
            parts['word/document.xml'] = root.toxml(encoding='utf-8')
            with zipfile.ZipFile(path, 'w') as archive:
                for name, data in parts.items():
                    archive.writestr(name, data)
        return path

    def test_last_page_pictures_removed_in_doc_and_docx(self):
        for suffix in ('.docx', '.doc'):
            with self.subTest(suffix=suffix):
                source = self.create_document(suffix)
                original = source.read_bytes()
                result = process_docx(source, 'last_page_images', self.folder / 'output', False)
                self.assertTrue(result.changed, result.message)
                self.assertEqual(original, source.read_bytes())
                self.assertEqual(suffix, result.output.suffix)
                doc = self.word.Documents.Open(str(result.output))
                try:
                    self.assertEqual(1, doc.InlineShapes.Count)
                    self.assertEqual(['keep-rectangle'], [s.Name for s in doc.Shapes])
                    self.assertIn('Last page', doc.Content.Text)
                    self.assertIn('End text', doc.Content.Text)
                    self.assertIn('Keep header', doc.Sections(1).Headers(1).Range.Text)
                    self.assertEqual(0 if suffix == '.doc' else 12, doc.SaveFormat)
                finally:
                    doc.Close(0)

    def test_automatic_pagination_and_restarted_page_numbers(self):
        for automatic, restart in ((True, False), (False, True)):
            with self.subTest(automatic=automatic, restart=restart):
                path = self.create_document('.docx', automatic=automatic, restart=restart)
                result = process_docx(path, 'last_page_images', None, True)
                self.assertTrue(result.changed, result.message)
                doc = self.word.Documents.Open(str(path))
                try:
                    self.assertEqual(1, doc.InlineShapes.Count)
                    self.assertEqual(1, doc.Shapes.Count)
                finally:
                    doc.Close(0)

    def test_no_pictures_leaves_original_bytes_unchanged(self):
        path = self.create_document('.docx', pictures=False)
        original = path.read_bytes()
        result = process_docx(path, 'last_page_images', None, True)
        self.assertFalse(result.changed)
        self.assertEqual(original, path.read_bytes())

    def test_blank_last_page_does_not_remove_previous_page_pictures(self):
        path = self.create_document('.docx')
        doc = self.word.Documents.Open(str(path))
        try:
            end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            end.InsertBreak(7)
            doc.Save()
        finally:
            doc.Close(0)
        original = path.read_bytes()
        result = process_docx(path, 'last_page_images', None, True)
        self.assertFalse(result.changed, result.message)
        self.assertEqual(original, path.read_bytes())

    def test_single_page_with_tracking_keeps_header_image_and_text(self):
        from docx import Document
        path = self.folder / 'single.docx'
        fixture = Document()
        fixture.add_paragraph('Keep text').add_run().add_picture(str(self.picture))
        fixture.sections[0].header.paragraphs[0].add_run().add_picture(str(self.picture))
        fixture.save(path)
        doc = self.word.Documents.Open(str(path))
        try:
            doc.TrackRevisions = True
            doc.Save()
        finally:
            doc.Close(0)
        result = process_docx(path, 'last_page_images', None, True)
        self.assertTrue(result.changed, result.message)
        doc = self.word.Documents.Open(str(path))
        try:
            self.assertEqual(0, doc.InlineShapes.Count)
            self.assertEqual(0, doc.Revisions.Count)
            self.assertTrue(doc.TrackRevisions)
            self.assertEqual(1, doc.Sections(1).Headers(1).Range.InlineShapes.Count)
            self.assertIn('Keep text', doc.Content.Text)
        finally:
            doc.Close(0)

    def test_table_images_keep_non_picture_content(self):
        from docx import Document
        from docx.shared import Pt
        path = self.create_document('.docx')
        fixture = Document(path)
        cell = fixture.add_table(rows=1, cols=1).cell(0, 0)
        cell.text = 'Keep cell text'
        cell.paragraphs[0].add_run().add_picture(str(self.picture), width=Pt(18), height=Pt(18))
        fixture.save(path)
        result = process_docx(path, 'last_page_images', None, True)
        self.assertTrue(result.changed, result.message)
        doc = self.word.Documents.Open(str(path))
        try:
            self.assertEqual(1, doc.InlineShapes.Count)
            self.assertEqual(1, doc.Tables.Count)
            self.assertIn('Keep cell text', doc.Tables(1).Cell(1, 1).Range.Text)
            self.assertEqual(1, doc.Shapes.Count)
            self.assertEqual(1, doc.Shapes(1).Type)
        finally:
            doc.Close(0)

    def test_gui_batch_worker_accepts_doc_and_docx(self):
        from word_header_footer_gui import HeaderFooterCleanerApp
        paths = [self.create_document(suffix) for suffix in ('.doc', '.docx')]
        app = HeaderFooterCleanerApp.__new__(HeaderFooterCleanerApp)
        app.message_queue = queue.Queue()
        worker = threading.Thread(target=app._run_processing, args=(
            self.folder, 'last_page_images', False, None, None, 'left', 'right'))
        worker.start()
        worker.join(60)
        self.assertFalse(worker.is_alive(), '后台处理未完成')
        events = list(app.message_queue.queue)
        self.assertIn(('done', '完成：成功 2 个，跳过 0 个'), events)
        for path in paths:
            doc = self.word.Documents.Open(str(path))
            try:
                self.assertEqual(1, doc.InlineShapes.Count)
            finally:
                doc.Close(0)

    def test_legacy_footer_fields_alignment_and_removal(self):
        path = self.create_document('.doc')
        result = process_docx(path, 'set_footer', None, True,
                              footer_text='Page {PAGE} of {NUMPAGES}', footer_alignment='center')
        self.assertTrue(result.changed, result.message)
        doc = self.word.Documents.Open(str(path))
        try:
            self.assertEqual(0, doc.SaveFormat)
            self.assertIn('Keep header', doc.Sections(1).Headers(1).Range.Text)
            footer = doc.Sections(1).Footers(1).Range
            self.assertEqual(['PAGE', 'NUMPAGES'], [f.Code.Text.strip() for f in footer.Fields])
            self.assertEqual(1, footer.ParagraphFormat.Alignment)
            self.assertEqual(2, doc.InlineShapes.Count)
        finally:
            doc.Close(0)
        result = process_docx(path, 'both', None, True)
        self.assertTrue(result.changed, result.message)
        doc = self.word.Documents.Open(str(path))
        try:
            for kind in (doc.Sections(1).Headers, doc.Sections(1).Footers):
                for index in (1, 2, 3):
                    self.assertEqual('', kind(index).Range.Text.strip())
        finally:
            doc.Close(0)


if __name__ == '__main__':
    unittest.main()
