"""Regression checks for citation identity and Word field placement."""

import unittest

from paperflow.citekey_ast import inspect_citations, replace_citations


def _str(text):
    return {"t": "Str", "c": text}


def _cite_item(key, *, mode="NormalCitation", prefix=(), suffix=()):
    return {"citationId": key, "citationPrefix": list(prefix), "citationSuffix": list(suffix),
            "citationMode": {"t": mode}, "citationNoteNum": 17, "citationHash": 0}


def _cite(*items, latex_command=None):
    fallback = {"t": "RawInline", "c": ["latex", latex_command]} if latex_command else _str("[@key]")
    return {"t": "Cite", "c": [list(items), [fallback]]}


class CitekeyAstTests(unittest.TestCase):
    def test_groups_affixes_and_plain_note_citation_preserve_body_order(self):
        group = _cite(
            _cite_item("CaseSensitive", prefix=[_str("see")], suffix=[_str(", pp.\u00a021–24, emphasis added")]),
            _cite_item("casesensitive", mode="SuppressAuthor"),
        )
        note = {"t": "Note", "c": [{"t": "Para", "c": [_cite(_cite_item("InNote"))]}]}
        ast = {"blocks": [{"t": "Para", "c": [group, note]}], "meta": {}}
        clusters, issues = inspect_citations(ast, "markdown-implicit_figures")
        self.assertEqual(issues, [])
        self.assertEqual(clusters, [
            {"index": 1, "items": [
                {"citekey": "CaseSensitive", "prefix": "see", "locator": "21–24",
                 "label": "page", "suffix": ", emphasis added"},
                {"citekey": "casesensitive", "suppress-author": True},
            ]},
            {"index": 2, "items": [{"citekey": "InNote"}]},
        ])
        with self.assertRaisesRegex(ValueError, "cover every Pandoc citation"):
            replace_citations(ast, {1: "<w:r/>"})
        result = replace_citations(ast, {1: "<w:r>first</w:r>", 2: "<w:r>note</w:r>"})
        body_group, body_note = result["blocks"][0]["c"]
        self.assertEqual(body_group, {"t": "RawInline", "c": ["openxml", "<w:r>first</w:r>"]})
        self.assertEqual(body_note["c"][0]["c"][0], {"t": "RawInline", "c": ["openxml", "<w:r>note</w:r>"]})
        self.assertEqual(group["t"], "Cite")

    def test_unsafe_contexts_report_location_without_renumbering_later_items(self):
        heading = {"t": "Header", "c": [1, ["", [], []], [_cite(_cite_item("InHeading"))]]}
        rich = _cite(_cite_item("Rich", prefix=[{"t": "Emph", "c": [_str("see")]}]))
        narrative = _cite(_cite_item("Narrative", mode="AuthorInText"))
        ast = {
            "meta": {"nocite": {"t": "MetaString", "c": "@*"}},
            "blocks": [heading, {"t": "Para", "c": [rich, narrative, _cite(_cite_item("Valid"))]}],
        }
        clusters, issues = inspect_citations(ast, "markdown")
        self.assertEqual(clusters, [{"index": 4, "items": [{"citekey": "Valid"}]}])
        self.assertEqual([issue["index"] for issue in issues], [None, 1, 2, 3])
        self.assertEqual([issue["context"] for issue in issues], ["metadata.nocite", "heading", "body", "body"])
        with self.assertRaisesRegex(ValueError, "unsupported Pandoc contexts"):
            replace_citations(ast, {index: "<w:r/>" for index in range(1, 5)})

    def test_latex_uses_raw_command_and_does_not_treat_starred_or_textual_cites_as_normal(self):
        ast = {"blocks": [{"t": "Para", "c": [
            _cite(_cite_item("A", suffix=[_str("21")]), latex_command=r"\citep[21]{A}"),
            _cite(_cite_item("B"), latex_command=r"\citep*{B}"),
            _cite(_cite_item("C"), latex_command=r"\citealp{C}"),
            _cite(_cite_item("D"), latex_command=r"\parencite{D}"),
        ]}]}
        clusters, issues = inspect_citations(ast, "latex")
        self.assertEqual(clusters, [
            {"index": 1, "items": [{"citekey": "A", "locator": "21", "label": "page"}]},
            {"index": 4, "items": [{"citekey": "D"}]},
        ])
        self.assertEqual([issue["index"] for issue in issues], [2, 3])

    def test_ambiguous_locator_is_not_silently_mapped_to_suffix(self):
        ast = {"blocks": [{"t": "Para", "c": [
            _cite(_cite_item("A", suffix=[_str(", chap. 3")])),
            _cite(_cite_item("B", suffix=[_str(", p. 14"), _str(", 17")])),
        ]}]}
        clusters, issues = inspect_citations(ast, "markdown")
        self.assertFalse(clusters)
        self.assertEqual([issue["index"] for issue in issues], [1, 2])


if __name__ == "__main__":
    unittest.main()
