"""Contract tests for dynamic citation data, not native add-in acceptance."""

import copy
import json
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from paperflow.zotero import (
    CT_NS,
    CUSTOM_NS,
    CUSTOM_REL,
    REL_NS,
    VT_NS,
    ZoteroClient,
    document_preferences,
    make_bibliography_field,
    make_citation_field,
    preference_properties,
    set_docx_preferences,
)


def fixture_item():
    """Synthetic metadata only; never imported into a real library."""
    return {"key": "ABCD1234", "uri": "http://zotero.org/users/123456/items/ABCD1234", "item_data": {"type": "article-journal", "title": "测试研究 α", "author": [{"family": "Test"}]}}


class ZoteroContractTests(unittest.TestCase):
    def test_citation_uri_and_unicode_are_preserved_without_mutating_input(self):
        item = fixture_item()
        item.update({"locator": "12–14", "label": "page", "suppress-author": True})
        original = copy.deepcopy(item)
        code = make_citation_field([item], citation_id="testCitation")
        payload = json.loads(code[code.index("{"):code.rindex("}") + 1])
        entry = payload["citationItems"][0]
        self.assertEqual(entry["uris"], [item["uri"]])
        self.assertEqual(entry["itemData"]["title"], "测试研究 α")
        self.assertEqual(entry["locator"], "12–14")
        self.assertTrue(entry["suppress-author"])
        self.assertEqual(item, original)
        self.assertNotIn("dontUpdate", payload["properties"])
        self.assertNotIn("plainCitation", payload["properties"])

    def test_citations_have_distinct_ids_and_reject_unlinked_data(self):
        self.assertNotEqual(make_citation_field([fixture_item()]), make_citation_field([fixture_item()]))
        for uri in ("http://zotero.org/users/0/items/ABCD1234", "https://example.com/items/ABCD1234", ""):
            item = fixture_item()
            item["uri"] = uri
            with self.assertRaises(ValueError):
                make_citation_field([item])
        with self.assertRaises(ValueError):
            make_citation_field([])

    def test_bibliography_is_dynamic_code_and_validates_uncited_uris(self):
        code = make_bibliography_field(uncited_uris=[fixture_item()["uri"]])
        self.assertTrue(code.startswith(" ADDIN ZOTERO_BIBL "))
        self.assertTrue(code.endswith(" CSL_BIBLIOGRAPHY "))
        payload = json.loads(code[code.index("{"):code.rindex("}") + 1])
        self.assertEqual(payload["uncited"], [[fixture_item()["uri"]]])
        self.assertEqual(payload["custom"], [])
        with self.assertRaises(ValueError):
            make_bibliography_field(uncited_uris=["http://zotero.org/users/0/items/ABCD1234"])

    def test_preferences_native_chunks_roundtrip_unicode(self):
        xml = document_preferences(locale="zh-CN", session_id="sessionTest")
        root = ET.fromstring(xml)
        self.assertEqual(root.attrib["data-version"], "3")
        self.assertEqual(root.find("prefs/pref[@name='fieldType']").attrib["value"], "Field")
        # Non-BMP codepoints must not be split into invalid surrogate halves.
        long_xml = '<data note="' + "文字📄" * 180 + '" />'
        chunks = preference_properties(long_xml)
        self.assertEqual("".join(chunks.values()), long_xml)
        self.assertTrue(all(len(v.encode("utf-16-le")) // 2 <= 255 for v in chunks.values()))
        self.assertEqual(next(iter(chunks)), "ZOTERO_PREF_1")

    def test_package_update_keeps_other_content_and_removes_stale_chunks(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "fixture.zip"
            custom = ET.Element(f"{{{CUSTOM_NS}}}Properties")
            for pid, name, value in ((2, "Owner", "保持"), (3, "ZOTERO_PREF_1", "old"), (4, "ZOTERO_PREF_9", "stale")):
                prop = ET.SubElement(custom, f"{{{CUSTOM_NS}}}property", {"pid": str(pid), "name": name})
                ET.SubElement(prop, f"{{{VT_NS}}}lpwstr").text = value
            with zipfile.ZipFile(path, "w") as package:
                package.writestr("word/document.xml", b"<untouched/>")
                package.writestr("word/embeddings/equation.bin", b"OLE-SENTINEL")
                package.writestr("docProps/custom.xml", ET.tostring(custom))
                package.writestr("[Content_Types].xml", f'<Types xmlns="{CT_NS}"/>')
                package.writestr("_rels/.rels", f'<Relationships xmlns="{REL_NS}"><Relationship Id="rId1" Type="keep" Target="word/document.xml"/></Relationships>')
            expected = document_preferences(session_id="stable")
            set_docx_preferences(path, expected)
            with zipfile.ZipFile(path) as package:
                self.assertEqual(package.read("word/document.xml"), b"<untouched/>")
                self.assertEqual(package.read("word/embeddings/equation.bin"), b"OLE-SENTINEL")
                props = ET.fromstring(package.read("docProps/custom.xml"))
                values = {p.attrib["name"]: p[0].text for p in props}
                self.assertEqual(values["Owner"], "保持")
                self.assertNotIn("ZOTERO_PREF_9", values)
                self.assertEqual("".join(values[f"ZOTERO_PREF_{i}"] for i in range(1, len(values))), expected)
                rels = ET.fromstring(package.read("_rels/.rels"))
                self.assertEqual(len([r for r in rels if r.get("Type") == CUSTOM_REL]), 1)
                self.assertEqual(len({r.get("Id") for r in rels}), len(rels))

    def test_client_restricts_reads_to_loopback_and_checks_keys_before_request(self):
        for origin in ("https://api.zotero.org", "http://example.com", "http://localhost:23119/api", "http://name:secret@localhost:23119"):
            with self.assertRaises(ValueError):
                ZoteroClient(origin)
        with self.assertRaises(ValueError):
            ZoteroClient().get_item("bibtex_citekey_2020")


if __name__ == "__main__":
    unittest.main()
