import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.dom import minidom
from xml.etree import ElementTree as ET

from remove_word_headers_footers import WORD_NS, REL_ATTR, process_docx
from test_footer_update import create_sample_docx

W = '{' + WORD_NS + '}'


class SpecialPageTests(unittest.TestCase):
    def make_document(self, path):
        create_sample_docx(path, True, True)
        with zipfile.ZipFile(path) as z:
            parts = {n: z.read(n) for n in z.namelist()}
        dom = minidom.parseString(parts['word/document.xml'])
        dom.getElementsByTagNameNS(WORD_NS, 'sectPr')[0].appendChild(dom.createElementNS(WORD_NS, 'w:titlePg'))
        parts['word/document.xml'] = dom.toxml(encoding='utf-8')
        parts['word/settings.xml'] = f'<w:settings xmlns:w="{WORD_NS}"><w:evenAndOddHeaders/></w:settings>'.encode()
        parts['word/styles.xml'] = f'<w:styles xmlns:w="{WORD_NS}"><w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:pPr><w:pBdr><w:bottom w:val="single"/></w:pBdr></w:pPr></w:style></w:styles>'.encode()
        with zipfile.ZipFile(path, 'w') as z:
            for n, b in parts.items():
                z.writestr(n, b)
        return parts

    def assert_all_pages(self, path, kind, expected):
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read('word/document.xml'))
            relationships = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('word/_rels/document.xml.rels'))}
            for section in root.iter(W + 'sectPr'):
                refs = {r.get(W + 'type'): r.get(REL_ATTR) for r in section.findall(W + kind + 'Reference')}
                self.assertEqual({'default', 'first', 'even'}, set(refs))
                for rid in refs.values():
                    part = ET.fromstring(z.read('word/' + relationships[rid]))
                    self.assertEqual(expected, ''.join(t.text or '' for t in part.iter(W + 't')))
                    bottom = part.find('.//' + W + 'pBdr/' + W + 'bottom')
                    self.assertIsNotNone(bottom)
                    self.assertEqual('nil', bottom.get(W + 'val'))

    def test_update_reaches_first_even_and_default_without_changing_other_kind(self):
        for kind in ('header', 'footer'):
            with self.subTest(kind=kind), tempfile.TemporaryDirectory() as d:
                p = Path(d) / 'sample.docx'
                original = self.make_document(p)
                self.assertTrue(process_docx(p, 'set_' + kind, None, True, **{kind + '_text': 'NEW'}).changed)
                self.assert_all_pages(p, kind, 'NEW')
                with zipfile.ZipFile(p) as z:
                    other = 'footer' if kind == 'header' else 'header'
                    self.assertEqual(original['word/' + other + '1.xml'], z.read('word/' + other + '1.xml'))
                    self.assertEqual(original['word/styles.xml'], z.read('word/styles.xml'))
                    self.assertEqual(original['word/settings.xml'], z.read('word/settings.xml'))
                self.assertFalse(process_docx(p, 'set_' + kind, None, True, **{kind + '_text': 'NEW'}).changed)

    def test_remove_leaves_explicit_borderless_empty_headers_on_all_pages(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'sample.docx'
            self.make_document(p)
            self.assertTrue(process_docx(p, 'both', None, True).changed)
            self.assert_all_pages(p, 'header', '')
            self.assert_all_pages(p, 'footer', '')
            self.assertFalse(process_docx(p, 'both', None, True).changed)

    def test_legacy_missing_compatibility_prefixes_are_repaired(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / 'sample.docx'
            parts = self.make_document(p)
            dom = minidom.parseString(parts['word/document.xml'])
            dom.documentElement.setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
            dom.documentElement.setAttribute('mc:Ignorable', 'w14 w15 wp14')
            parts['word/document.xml'] = dom.toxml(encoding='utf-8')
            with zipfile.ZipFile(p, 'w') as z:
                for n,b in parts.items(): z.writestr(n,b)
            process_docx(p, 'set_footer', None, True, footer_text='NEW')
            with zipfile.ZipFile(p) as z:
                root = minidom.parseString(z.read('word/document.xml')).documentElement
                for prefix in ('w14', 'w15', 'wp14'):
                    self.assertTrue(root.getAttribute('xmlns:' + prefix))
