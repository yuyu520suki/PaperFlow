import unittest
from unittest.mock import Mock

from paperflow.doi import normalize_doi, resolve_dois
from paperflow.zotero import ZoteroError


def record(key, doi, *, item_type="journalArticle", title="Fixture"):
    return {"key": key, "data": {"DOI": doi, "title": title, "itemType": item_type}}


class DoiTests(unittest.TestCase):
    def test_normalize_only_doi_syntax_and_ascii_case(self):
        self.assertEqual(normalize_doi(" DOI: 10.1234.2/ÄBc. "), "10.1234.2/Äbc.")
        self.assertEqual(normalize_doi("https://dx.doi.org/10.1234/%C3%84%23B"), "10.1234/Ä#b")
        self.assertIsNone(normalize_doi("https://example.org/10.1234/abc"))
        self.assertIsNone(normalize_doi("https://doi.org/10.1234/%ZZ"))
        self.assertIsNone(normalize_doi("10.1234/abc and more"))

    def test_resolve_filters_metadata_and_retains_duplicate_candidates(self):
        client = Mock()
        client.search.side_effect = [
            [record("ABCDEFG2", "https://doi.org/10.1234/ABC"),
             record("ABCDEFG3", "10.1234/abcd"),
             record("ABCDEFG4", "10.1234/abc", item_type="attachment"),
             record("ABCDEFG5", "10.1234/abc", title="Second copy"),
             record("ABCDEFG2", "https://doi.org/10.1234/ABC")],
            [],
        ]
        result = resolve_dois(["10.1234/abc", "DOI: 10.1234/ABC", "10.1234/none"], client=client)
        self.assertEqual([item["key"] for item in result["10.1234/abc"]], ["ABCDEFG2", "ABCDEFG5"])
        self.assertEqual(result["10.1234/none"], [])
        self.assertEqual(client.search.call_count, 2)
        client.search.assert_any_call("10.1234/abc", limit=100, mode="everything")

    def test_potentially_truncated_search_never_reports_unique_match(self):
        client = Mock()
        client.search.return_value = [record("ABCDEFG2", "10.1234/abc")] + [record("ABCDEFG3", "10.1234/other")] * 99
        with self.assertRaisesRegex(ZoteroError, "reached its result limit"):
            resolve_dois(["10.1234/abc"], client=client)
