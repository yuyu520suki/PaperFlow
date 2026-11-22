import unittest

from paperflow.recovery_proposals import propose_blocks


class RecoveryProposalTests(unittest.TestCase):
    def test_literal_manuscript_proposes_structure_without_skipping_lines(self):
        lines = [
            "",
            "悬索结构非线性稳定性分析",
            "摘要",
            "这段文字来自原稿。",
            "1 Introduction",
            "1.1 Cable properties",
            "1.1.1 Modal response",
            "1. This sentence is a list item.",
            "",
            "Case\tLoad\tComment",
            "A\t4.2\t",
            "B\t8.5\tcheck",
            "$$",
            "k = E A / L",
            "$$",
            "参考文献",
        ]

        blocks = propose_blocks(lines)
        self.assertEqual([line for block in blocks for line in range(block["start_line"], block["end_line"] + 1)],
                         list(range(len(lines))))
        by_start = {block["start_line"]: block for block in blocks}
        self.assertEqual(by_start[1]["kind"], "title")
        self.assertEqual([by_start[line]["level"] for line in (2, 4, 5, 6, 15)], [1, 1, 2, 3, 1])
        self.assertEqual(by_start[7]["kind"], "paragraph")
        self.assertEqual(by_start[9]["rows"][1], ["A", "4.2", ""])
        self.assertEqual(by_start[12], {"start_line": 12, "end_line": 14, "kind": "review",
                                        "reason": "possible_formula"})

    def test_markdown_separator_is_not_a_cell_row_and_irregular_tables_stay_literal(self):
        lines = [
            "Mode | Frequency",
            "--- | ---:",
            "1 | 7.4",
            "2 | 9.1",
            "",
            "Case    Load    Value",
            "A       2.0     8.6",
            "B       4.0     9.3",
            "",
            "| A | B |",
            "| 2 | 3 | 4 |",
        ]

        by_start = {block["start_line"]: block for block in propose_blocks(lines)}
        self.assertEqual(by_start[0]["rows"], [["Mode", "Frequency"], ["1", "7.4"], ["2", "9.1"]])
        self.assertEqual(by_start[0]["header_rows"], 1)
        for start in (5, 9):
            self.assertEqual(by_start[start]["kind"], "review")
            self.assertEqual(by_start[start]["reason"], "ambiguous_table")
            self.assertNotIn("rows", by_start[start])

    def test_ambiguous_title_and_unclosed_formula_do_not_swallow_other_roles(self):
        lines = ["Journal of Applied Mechanics", "Elastodynamic stability", "Abstract", r"\[x^2", "结论", "仍然可读。"]
        by_start = {block["start_line"]: block for block in propose_blocks(lines)}

        self.assertEqual(by_start[0]["kind"], "paragraph")
        self.assertEqual(by_start[3]["end_line"], 3)
        self.assertEqual(by_start[4]["kind"], "heading")
        self.assertEqual(by_start[5]["kind"], "paragraph")


if __name__ == "__main__":
    unittest.main()
