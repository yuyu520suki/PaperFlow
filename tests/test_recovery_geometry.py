import unittest

from paperflow.recovery_geometry import detect_ocr_tables


def ocr_line(text, left, top, *, page=1):
    words = []
    position = left
    for token in text.split():
        right = position + len(token) * .007
        words.append({"text": token, "bbox": [position, top, right, top + .023]})
        position = right + .006
    return {"text": text, "page": page, "bbox": [left, top, words[-1]["bbox"][2], top + .023],
            "adapter": "windows-ocr", "words": words}


class RecoveryGeometryTests(unittest.TestCase):
    def test_column_major_output_does_not_consume_intervening_prose(self):
        source = [
            ocr_line("Case", .1, .28),
            ocr_line("Steel A", .1, .33),
            ocr_line("Steel B", .1, .38),
            ocr_line("End of measurements", .1, .46),
            ocr_line("Voltage", .43, .28),
            ocr_line("220.0", .43, .33),
            ocr_line("215.3", .43, .38),
            ocr_line("Current", .69, .28),
            ocr_line("1.51", .69, .33),
            ocr_line("1.44", .69, .38),
        ]
        proposal, = detect_ocr_tables(source)
        self.assertEqual(proposal["line_indices"], [0, 1, 2, 4, 5, 6, 7, 8, 9])
        self.assertEqual(proposal["rows"], [
            ["Case", "Voltage", "Current"],
            ["Steel A", "220.0", "1.51"],
            ["Steel B", "215.3", "1.44"],
        ])
        self.assertTrue(proposal["evidence"]["requires_review"])

    def test_one_ocr_line_per_row_preserves_original_cell_text(self):
        source = []
        for top, cells in zip((.25, .3, .35), (
            ("物理量", "测试值"), ("位移  最大值", "2.30"), ("速度 峰值", "1.25"),
        )):
            left = ocr_line(cells[0], .1, top)
            right = ocr_line(cells[1], .45, top)
            left["bbox"][2] = right["bbox"][2]
            left["words"].extend(right["words"])
            left["text"] = "   ".join(cells)
            source.append(left)
        proposal, = detect_ocr_tables(source)
        self.assertEqual(proposal["rows"][1], ["位移  最大值", "2.30"])

    def test_uncaptioned_parallel_prose_and_incomplete_tokens_stay_literal(self):
        prose = []
        for top in (.25, .30, .35):
            prose.extend((
                ocr_line("This is ordinary text", .10, top),
                ocr_line("A second prose column", .58, top),
            ))
        self.assertEqual(detect_ocr_tables(prose), [])

        source = [
            ocr_line(label, left, top)
            for top, row in zip((.25, .30, .35), (("Label", "Value"), ("A", "1.20"), ("B", "2.0")))
            for label, left in zip(row, (.1, .48))
        ]
        source[3]["words"][0]["text"] = "12.0"
        self.assertEqual(detect_ocr_tables(source), [])


if __name__ == "__main__":
    unittest.main()
