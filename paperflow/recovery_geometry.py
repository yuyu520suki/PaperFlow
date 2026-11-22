"""Propose reviewable tables from the positions of complete OCR tokens."""

from __future__ import annotations

from dataclasses import dataclass
import re
from statistics import median


_OCR_ADAPTERS = {"windows-ocr", "windows-ocr-image"}
_CAPTION = re.compile(r"^(?:(?:table|tab\.)\s+[A-Z]?\d+\b|表\s*[A-Z]?\d+)", re.IGNORECASE)
_NUMBER = re.compile(r"[+−-]?(?:\d+(?:[,.]\d+)*|[.,]\d+)(?:[eE][+−-]?\d+)?\s*[%‰]?")

_Box = tuple[float, float, float, float]


@dataclass(frozen=True)
class _Word:
    text: str
    box: _Box
    line_index: int
    start: int
    end: int


@dataclass(frozen=True)
class _Line:
    index: int
    text: str
    box: _Box
    words: tuple[_Word, ...] | None

    @property
    def center(self):
        return (self.box[1] + self.box[3]) / 2

    @property
    def height(self):
        return self.box[3] - self.box[1]


@dataclass(frozen=True)
class _Row:
    lines: tuple[_Line, ...]
    cells: tuple[str, ...]
    cell_boxes: tuple[_Box, ...]
    center: float
    height: float


def _read_box(value) -> _Box | None:
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        return None
    if any(type(item) not in (int, float) or not 0 <= item <= 1 for item in value):
        return None
    x0, y0, x1, y1 = value
    if 0 <= x0 < x1 <= 1 and 0 <= y0 < y1 <= 1:
        return x0, y0, x1, y1
    return None


def _tokens(raw_words, text, box, index) -> tuple[_Word, ...] | None:
    """Bind every token to its exact source text span before using geometry."""
    if not isinstance(raw_words, list) or not raw_words:
        return None
    positions = [index for index, char in enumerate(text) if not char.isspace()]
    compact = "".join(text[index] for index in positions)
    offset = 0
    tokens = []
    for item in raw_words:
        if not isinstance(item, dict) or not isinstance(item.get("text"), str):
            return None
        word = "".join(item["text"].split())
        bounds = _read_box(item.get("bbox"))
        if not word or bounds is None or not compact.startswith(word, offset):
            return None
        if any(lower - upper > .002 for lower, upper in (
            (box[0], bounds[0]), (box[1], bounds[1]),
            (bounds[2], box[2]), (bounds[3], box[3]),
        )):
            return None
        start, end = positions[offset], positions[offset + len(word) - 1] + 1
        if tokens and bounds[0] < tokens[-1].box[0]:
            return None
        tokens.append(_Word(text[start:end], bounds, index, start, end))
        offset += len(word)
    return tuple(tokens) if offset == len(compact) else None


def _visual_rows(lines):
    """Group OCR fragments sharing a baseline, even in column-major output."""
    rows = []
    for line in sorted(lines, key=lambda entry: (entry.center, entry.box[0], entry.index)):
        if not rows:
            rows.append([line])
            continue
        previous = rows[-1]
        center = median(entry.center for entry in previous)
        height = median(entry.height for entry in previous)
        if abs(line.center - center) <= .48 * min(height, line.height):
            previous.append(line)
        else:
            rows.append([line])
    return rows


def _cell_text(group, texts):
    chunks = []
    line_id = group[0].line_index
    start = group[0].start
    end = group[0].end
    for word in group[1:]:
        if word.line_index != line_id:
            chunks.append(texts[line_id][start:end])
            line_id, start = word.line_index, word.start
        end = word.end
    chunks.append(texts[line_id][start:end])
    return " ".join(chunks)


def _segment_row(lines, minimum_gap):
    center = median(line.center for line in lines)
    height = median(line.height for line in lines)
    if any(line.words is None or _CAPTION.match(line.text.strip()) for line in lines):
        return _Row(tuple(lines), (), (), center, height)
    words = sorted((word for line in lines for word in line.words), key=lambda word: word.box[0])
    groups = [[words[0]]]
    for previous, word in zip(words, words[1:]):
        gap = word.box[0] - previous.box[2]
        if gap < -.002 and previous.line_index != word.line_index:
            return _Row(tuple(lines), (), (), center, height)
        if gap >= minimum_gap:
            groups.append([])
        groups[-1].append(word)
    texts = {line.index: line.text for line in lines}
    cells = tuple(_cell_text(group, texts) for group in groups)
    boxes = tuple(
        (min(w.box[0] for w in group), min(w.box[1] for w in group),
         max(w.box[2] for w in group), max(w.box[3] for w in group))
        for group in groups
    )
    return _Row(tuple(lines), cells, boxes, center, height)


def _grid_matches(rows, minimum_gap, tolerance):
    width = len(rows[0].cells)
    if not 2 <= width <= 12 or any(len(row.cells) != width for row in rows):
        return False
    if any(len(cell) > 90 for row in rows for cell in row.cells):
        return False
    gaps = [second.center - first.center for first, second in zip(rows, rows[1:])]
    if gaps:
        typical_height = median(row.height for row in rows)
        if max(gaps) > 3.6 * typical_height or max(gaps) > 1.55 * min(gaps):
            return False
    for column in range(width):
        column_boxes = [row.cell_boxes[column] for row in rows]
        starts, ends = ([box[edge] for box in column_boxes] for edge in (0, 2))
        if min(max(starts) - min(starts), max(ends) - min(ends)) > tolerance:
            return False
        if column + 1 < width:
            next_start = min(row.cell_boxes[column + 1][0] for row in rows)
            if next_start - max(ends) < .72 * minimum_gap:
                return False
    return True


def _table_signal(rows, preceding_row):
    if any(
        sum(bool(_NUMBER.fullmatch(cell.strip())) for cell in column) >= max(2, len(rows) - 1)
        for column in zip(*(row.cells for row in rows))
    ):
        return "consistent_numeric_column"
    if preceding_row and len(preceding_row.lines) == 1:
        caption = preceding_row.lines[0]
        gap = rows[0].center - preceding_row.center
        if _CAPTION.match(caption.text.strip()) and gap <= 3.6 * max(rows[0].height, preceding_row.height):
            return "nearby_table_caption"
    return None


def _page_tables(page, lines):
    char_widths = [
        (word.box[2] - word.box[0]) / len("".join(word.text.split()))
        for line in lines if line.words
        for word in line.words
    ]
    if not char_widths:
        return []
    glyph_width = median(char_widths)
    minimum_gap = max(.014, 2.6 * glyph_width)
    tolerance = max(.009, min(.024, 1.1 * glyph_width))
    rows = [_segment_row(fragments, minimum_gap) for fragments in _visual_rows(lines)]
    proposals = []
    start = 0
    while start + 2 < len(rows):
        end = start + 3
        if not _grid_matches(rows[start:end], minimum_gap, tolerance):
            start += 1
            continue
        while end < len(rows) and _grid_matches(rows[start:end + 1], minimum_gap, tolerance):
            end += 1
        group = rows[start:end]
        signal = _table_signal(group, rows[start - 1] if start else None)
        if not signal:
            start = end
            continue
        row_indices = [sorted(line.index for line in row.lines) for row in group]
        line_indices = sorted(index for indices in row_indices for index in indices)
        proposals.append({
            "kind": "table",
            "start_line": line_indices[0],
            "end_line": line_indices[-1],
            "line_indices": line_indices,
            "rows": [list(row.cells) for row in group],
            "header_rows": 0,
            "evidence": {
                "page": page,
                "method": "ocr_column_alignment",
                "signal": signal,
                "row_source_line_indices": row_indices,
                "requires_review": True,
            },
        })
        start = end
    return proposals


def detect_ocr_tables(source_lines: list[dict]) -> list[dict]:
    """Suggest tables using 0-based original line IDs; ranges may have holes.

    Only validated OCR geometry participates. A suggestion uses complete visual
    rows with repeated gutters and requires human comparison to the source.
    """
    pages: dict[int, list[_Line]] = {}
    unlocatable_pages = set()
    for index, item in enumerate(source_lines):
        page, text = item.get("page"), item.get("text")
        if (item.get("adapter") not in _OCR_ADAPTERS or type(page) is not int
                or page < 1 or not isinstance(text, str) or not text.strip()):
            continue
        box = _read_box(item.get("bbox"))
        if box is None:
            unlocatable_pages.add(page)
            continue
        words = _tokens(item.get("words"), text, box, index)
        pages.setdefault(page, []).append(_Line(index, text, box, words))
    proposals = [
        proposal for page, lines in pages.items() if page not in unlocatable_pages
        for proposal in _page_tables(page, lines)
    ]
    return sorted(proposals, key=lambda proposal: proposal["start_line"])
