"""Regression checks for the LaTeX-to-Word number and reference mapping."""

import unittest

from paperflow.latex_references import needs_heading_labels, prepare_latex_references


def _node(kind, content):
    return {"t": kind, "c": content}


def _raw(text):
    return _node("RawInline", ["latex", text])


def _raw_block(text):
    return _node("RawBlock", ["latex", text])


def _heading(label, title):
    return _node("Header", [1, [label, [], []], [_node("Str", title)]])


def _math(latex):
    return _node("Math", [{"t": "DisplayMath"}, latex])


def _figure(label):
    image = _node("Image", [["", [], []], [], ["response.png", ""]])
    caption = _node("Plain", [_node("Emph", [_node("Str", "Rich caption")])])
    return _node("Figure", [[label, [], []], [None, [caption]], [_node("Plain", [image])]])


class LatexReferenceTests(unittest.TestCase):
    def test_sequence_split_and_rich_caption_preserve_source_order(self):
        prefix = _node("Str", "Prose")
        live_citation = _node("RawInline", ["openxml", "<w:r/>"])
        empty_label = _math(r"\begin{equation}u=1\end{equation}")
        labeled = _math(r"\begin{equation}v=2\label{eq:one}\end{equation}")
        starred = _math(r"\begin{equation*}w=3\end{equation*}")
        figure = _figure("fig:one")
        ast = {"blocks": [_node("Para", [prefix, live_citation, empty_label, labeled, starred]), figure,
                          _node("Para", [_raw(r"\eqref{eq:one}"), _raw(r"\autoref{fig:one}")])]}
        result, report = prepare_latex_references(ast)
        self.assertEqual(report["issues"], [])
        self.assertEqual([(target["kind"], target["label"]) for target in report["targets"]],
                         [("Equation", None), ("Equation", "eq:one"), ("Figure", "fig:one")])
        unnamed, equation, picture = [target["bookmark"] for target in report["targets"]]
        original, first, second, unnumbered, converted_figure, refs = result["blocks"]
        self.assertEqual(original["c"], [prefix, live_citation])
        self.assertEqual(first["c"], [_node("Math", [{"t": "InlineMath"}, "u=1"]),
                                      _node("Str", f"[[EQNO {unnamed}]]")])
        self.assertEqual(second["c"], [_node("Math", [{"t": "InlineMath"}, "v=2"]),
                                       _node("Str", f"[[EQNO {equation}]]")])
        self.assertEqual(unnumbered["c"], [starred])
        image, caption = converted_figure["c"][1]
        self.assertEqual(image["c"][0]["t"], "Image")
        self.assertEqual(caption["c"], [_node("Str", f"[[CAPTIONNO Figure {picture}]]"),
                                        {"t": "Space"}, figure["c"][1][1][0]["c"][0]])
        self.assertEqual(refs["c"], [_node("Str", f"([[REF {equation}]])"),
                                     _node("Str", f"Figure [[REF {picture}]]")])
        self.assertEqual(labeled["c"][0]["t"], "DisplayMath")
        self.assertEqual(figure["t"], "Figure")

    def test_unverifiable_targets_and_contexts_stay_visible(self):
        target = _math(r"\begin{equation}v=2\label{eq:one}\end{equation}")
        duplicate = _figure("fig:dup")
        refs = [_raw(r"\ref{fig:dup}"), _raw(r"\pageref*{eq:one}"), _raw(r"\ref{missing}")]
        note = _node("Note", [_node("Para", [_raw(r"\ref{eq:one}")])])
        result, report = prepare_latex_references({"blocks": [_node("Para", [target]), duplicate,
                                                               _figure("fig:dup"), _node("Para", [*refs, note])]})
        self.assertEqual([entry["label"] for entry in report["targets"]], ["eq:one"])
        self.assertFalse(report["references"])
        issues = [issue["reason"] for issue in report["issues"]]
        self.assertEqual(sum("occurs more than once" in reason for reason in issues), 2)
        self.assertTrue(any("in note" in reason for reason in issues))
        converted_refs = result["blocks"][-1]["c"]
        self.assertTrue(all(item["c"].startswith("[LaTeX reference needs review:") for item in converted_refs[:3]))
        self.assertTrue(converted_refs[-1]["c"][0]["c"][0]["c"].startswith("[LaTeX reference needs review:"))

    def test_page_references_use_supported_targets_and_verified_heading_labels(self):
        heading = _heading("sec:one", "Experimental method")
        implicit = _heading("unlabeled", "Unlabeled")
        figure = _figure("fig:one")
        table = _node("Table", [["tab:one", [], []], [None, [_node("Plain", [_node("Str", "Results")])]],
                                [], [["", [], []], []], [], [["", [], []], []]])
        equation = _math(r"\begin{equation}v=2\label{eq:one}\end{equation}")
        ast = {"blocks": [heading, implicit, figure, table, _node("Para", [equation]),
                          _raw_block(r"\pageref{sec:one}"), _raw_block(r"\pageref{unlabeled}"),
                          _node("Para", [_raw(r"\pageref{fig:one}"), _raw(r"\pageref{tab:one}"),
                                         _raw(r"\pageref{eq:one}"), _raw(r"\ref{sec:one}")])]}
        self.assertTrue(needs_heading_labels(ast))
        result, report = prepare_latex_references(ast, explicit_heading_ids={"sec:one"})
        bookmarks = {entry["label"]: entry["bookmark"] for entry in report["targets"]}
        self.assertEqual([entry["kind"] for entry in report["targets"]], ["Heading", "Figure", "Table", "Equation"])
        self.assertEqual([entry["command"] for entry in report["references"]], ["pageref"] * 4)
        self.assertEqual(result["blocks"][0]["c"][2], [
            heading["c"][2][0], _node("Str", f"[[ANCHOR {bookmarks['sec:one']}]]"), *heading["c"][2][1:],
        ])
        self.assertEqual(result["blocks"][1], implicit)
        self.assertEqual(result["blocks"][-3], _node("Para", [_node("Str", f"[[PAGEREF {bookmarks['sec:one']}]]")]))
        self.assertTrue(result["blocks"][-2]["c"][0]["c"].startswith("[LaTeX reference needs review:"))
        self.assertEqual(result["blocks"][-1]["c"][:3], [
            _node("Str", f"[[PAGEREF {bookmarks[label]}]]") for label in ("fig:one", "tab:one", "eq:one")
        ])
        self.assertEqual(len(report["issues"]), 2)
        self.assertIn("Numbered heading", report["issues"][-1]["reason"])

    def test_duplicate_and_nested_heading_page_reference_targets_need_review(self):
        duplicate = [_heading("dup", "First"), _heading("dup", "Second")]
        nested = _node("BlockQuote", [_heading("nested", "Nested")])
        ast = {"blocks": [*duplicate, nested, _raw_block(r"\pageref{dup}"), _raw_block(r"\pageref{nested}")]}
        result, report = prepare_latex_references(ast)
        self.assertEqual(result["blocks"][:3], [*duplicate, nested])
        self.assertEqual(report["targets"], [])
        self.assertEqual(report["references"], [])
        self.assertTrue(all(block["c"][0]["c"].startswith("[LaTeX reference needs review:")
                            for block in result["blocks"][-2:]))

    def test_numbered_heading_references_share_a_bookmark_with_page_references(self):
        heading = _heading("sec:one", "Methods")
        ast = {"blocks": [heading, _raw_block(r"\ref{sec:one}"), _raw_block(r"\autoref{sec:one}"),
                          _raw_block(r"\pageref{sec:one}")]}
        result, report = prepare_latex_references(ast, numbered_heading_levels={1})
        self.assertFalse(report["issues"])
        (target,) = report["targets"]
        bookmark = target["bookmark"]
        self.assertEqual(result["blocks"][0]["c"][2][-1], _node("Str", f"[[ANCHOR {bookmark}]]"))
        self.assertEqual([block["c"][0]["c"] for block in result["blocks"][1:]], [
            f"[[HEADINGREF {bookmark}]]", f"Section [[HEADINGREF {bookmark}]]", f"[[PAGEREF {bookmark}]]",
        ])
        numeric_only = {"blocks": ast["blocks"][:2]}
        self.assertTrue(needs_heading_labels(numeric_only))
        result, report = prepare_latex_references(numeric_only, numbered_heading_levels={1})
        self.assertEqual(result["blocks"][0]["c"][2][-1], _node("Str", f"[[ANCHOR {bookmark}]]"))
        self.assertEqual([ref["command"] for ref in report["references"]], ["ref"])

    def test_numbering_ambiguity_keeps_heading_page_references(self):
        for classes, options, reason in (
            (["unnumbered"], {"numbered_heading_levels": {1}}, "unnumbered LaTeX heading"),
            ([], {"numbered_heading_levels": {1}, "ambiguous_heading_numbering": True}, "custom or included"),
        ):
            with self.subTest(reason=reason):
                heading = _heading("sec:one", "Methods")
                heading["c"][1][1] = classes
                ast = {"blocks": [heading, _raw_block(r"\ref{sec:one}"), _raw_block(r"\pageref{sec:one}")]}
                result, report = prepare_latex_references(ast, **options)
                self.assertIn(reason, report["issues"][0]["reason"])
                self.assertEqual([ref["command"] for ref in report["references"]], ["pageref"])
                self.assertTrue(result["blocks"][1]["c"][0]["c"].startswith("[LaTeX reference needs review:"))

    def test_indexes_reuse_template_or_keep_source_position_and_flag_ambiguous_commands(self):
        intro = _node("Para", [_node("Str", "Preamble")])
        between = _node("Para", [_node("Str", "Between indexes")])
        nested = _node("BlockQuote", [_raw_block(r"\tableofcontents")])
        ast = {"blocks": [intro, _raw_block(r"\tableofcontents"), _raw_block(r"\listoffigures{}"), between,
                          _raw_block(r"\listoftables"), _raw_block(r"\listoftables"), nested,
                          _raw_block(r"\listoffigures[other]")]}
        result, report = prepare_latex_references(ast, template_indexes={"TOC"})
        self.assertEqual([(entry["kind"], entry["mode"]) for entry in report["indexes"]], [
            ("TOC", "template"), ("LOF", "source"), ("LOT", "source"),
        ])
        self.assertEqual(result["blocks"][:6], [
            intro,
            _node("Para", [_node("Strong", [_node("Str", "List of Figures")])]),
            _node("Para", [_node("Str", "[[LOF]]")]), between,
            _node("Para", [_node("Strong", [_node("Str", "List of Tables")])]),
            _node("Para", [_node("Str", "[[LOT]]")]),
        ])
        self.assertIn("Duplicate", report["issues"][0]["reason"])
        self.assertEqual(len(report["issues"]), 3)
        self.assertTrue(result["blocks"][6]["c"][0]["c"].startswith("[LaTeX index needs review:"))
        self.assertTrue(result["blocks"][7]["c"][0]["c"][0]["c"].startswith("[LaTeX index needs review:"))
        self.assertTrue(result["blocks"][8]["c"][0]["c"].startswith("[LaTeX index needs review:"))

    def test_index_on_its_own_source_line_can_split_a_pandoc_paragraph(self):
        intro, follow = _node("Str", "Preamble"), _node("Str", "Main text")
        soft = {"t": "SoftBreak"}
        ast = {"blocks": [_node("Para", [intro, soft, _raw(r"\tableofcontents"), soft, follow])]}
        result, report = prepare_latex_references(ast)
        self.assertFalse(report["issues"])
        self.assertEqual(result["blocks"], [
            _node("Para", [intro]), _node("Para", [_node("Strong", [_node("Str", "Contents")])]),
            _node("Para", [_node("Str", "[[TOC]]")]), _node("Para", [follow]),
        ])

    def test_starred_caption_guard_leaves_floats_unnumbered(self):
        figure = _figure("fig:one")
        result, report = prepare_latex_references({"blocks": [figure, _node("Para", [_raw(r"\ref{fig:one}")])]},
                                                  ambiguous_float_numbering=True)
        self.assertEqual(result["blocks"][0], figure)
        self.assertFalse(report["targets"])
        self.assertFalse(report["references"])
        self.assertIn("Starred captions", report["issues"][0]["reason"])


if __name__ == "__main__":
    unittest.main()
