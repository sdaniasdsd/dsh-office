"""XML-level edge cases for boundaries that cannot be represented as whole blocks."""
import sys
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src" / "engine"))
from observations import NS, read_observations


class ObservationsTest(unittest.TestCase):
    def parse(self, xml, flags=None):
        ns = " ".join(f'xmlns:{key}="{value}"' for key, value in NS.items())
        body = ET.fromstring(f"<w:body {ns}>{xml}</w:body>")
        warnings = []
        return read_observations(body, flags or {}, warnings), warnings

    def test_continuous_section_does_not_create_page_break(self):
        result, _ = self.parse('''<w:p><w:pPr><w:sectPr/></w:pPr></w:p>
            <w:p/><w:sectPr><w:type w:val="continuous"/></w:sectPr>''')
        self.assertEqual(result["breaks"], [])
        self.assertEqual(result["sections"][1]["startBodyIndex"], 1)
        self.assertIsNone(result["sections"][1]["geometry"])

    def test_false_break_and_revision_markers_are_ignored(self):
        result, _ = self.parse('''<w:p><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr>
            <w:del><w:r><w:br w:type="page"/></w:r></w:del>
            <w:r><w:br w:type="column"/></w:r></w:p>''')
        self.assertEqual(result["breaks"], [])

    def test_nested_paragraph_break_does_not_target_entire_table(self):
        result, _ = self.parse('''<w:tbl><w:tr><w:tc><w:p><w:pPr><w:pageBreakBefore/>
            </w:pPr></w:p></w:tc></w:tr></w:tbl>''')
        self.assertIsNone(result["breaks"][0]["beforeBodyIndex"])
        self.assertIn('/w:tbl[1]/w:tr[1]/w:tc[1]/w:p[1]', result["breaks"][0]["path"])

    def test_alternate_content_selects_one_branch_and_keeps_actual_xpath_index(self):
        result, _ = self.parse('''<w:p><w:r><mc:AlternateContent>
            <mc:Choice Requires="unknown"><wp:anchor/></mc:Choice>
            <mc:Choice Requires="wps"><wp:anchor simplePos="1"><wp:simplePos x="-5" y="8"/>
            </wp:anchor></mc:Choice><mc:Fallback><wp:inline/></mc:Fallback>
            </mc:AlternateContent></w:r></w:p>''')
        self.assertEqual(len(result["floats"]), 1)
        item = result["floats"][0]
        self.assertIn('/mc:Choice[2]/', item["path"])
        self.assertEqual(item["anchor"]["relativeFromHorizontal"], "page")
        self.assertEqual(item["anchor"]["offsetX"], -5)

    def test_empty_document_terminal_section_has_no_dangling_start(self):
        result, _ = self.parse('<w:sectPr/>')
        self.assertIsNone(result["sections"][0]["startBodyIndex"])
        self.assertEqual(result["breaks"], [])


if __name__ == "__main__":
    unittest.main()
