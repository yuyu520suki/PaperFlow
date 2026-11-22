import unittest

from paperflow.citekeys import _bibliography_location, _place_bibliography


def text_node(text):
    return {"t": "Str", "c": text}


def heading(text):
    return {"t": "Header", "c": [1, ["", [], []], [text_node(text)]]}


class CitekeyBibliographyTests(unittest.TestCase):
    def test_latex_bibliography_reuses_existing_heading_before_appendix(self):
        ast = {"blocks": [heading("References"), {"t": "RawBlock", "c": ["latex", r"\printbibliography"]},
                          heading("Appendix")]}
        location, issue = _bibliography_location(ast)
        self.assertIsNone(issue)
        _place_bibliography(ast, location)
        self.assertEqual(ast["blocks"], [heading("References"),
                                         {"t": "Para", "c": [text_node("[[BIBLIOGRAPHY]]")]},
                                         heading("Appendix")])

    def test_nested_bibliography_marker_is_not_silently_duplicated(self):
        ast = {"blocks": [{"t": "BlockQuote", "c": [{"t": "Para", "c": [text_node("[[BIBLIOGRAPHY]]")]}]}]}
        location, issue = _bibliography_location(ast)
        self.assertIsNone(location)
        self.assertIn("standalone", issue)


if __name__ == "__main__":
    unittest.main()
