"""Suggest structure for literal lines without replacing their source text."""

from __future__ import annotations

import re


_HEADINGS = {
    "abstract", "introduction", "background", "literature review", "related work",
    "methods", "methodology", "materials and methods", "experimental methods",
    "results", "discussion", "results and discussion", "conclusion", "conclusions",
    "acknowledgements", "acknowledgments", "references", "bibliography", "appendix",
    "data availability", "funding", "摘要", "中文摘要", "英文摘要", "引言", "绪论", "研究背景",
    "相关工作", "文献综述", "研究方法", "材料与方法", "实验方法", "实验结果", "结果", "讨论",
    "结果与讨论", "结论", "总结", "致谢", "参考文献", "附录",
}
_DECIMAL_HEADING = re.compile(r"^(?P<number>[1-9][0-9]?(?:\.[1-9][0-9]?){0,2})[.、]?\s+(?P<title>\S.*)$")
_CHINESE_CHAPTER = re.compile(r"^第\s*[一二三四五六七八九十百零〇两0-9]+\s*章\s*(?P<title>\S.*)$")
_LABEL_TITLE = re.compile(r"^(?:title|论文题目|题目|题名)\s*[:：]\s*\S", re.I)
_TITLE_METADATA = re.compile(
    r"(?:\b(?:https?://|doi\s*:|issn\b|orcid\b|received\b|published\b|copyright\b|vol\.\s*\d)|@|[©])",
    re.I,
)
_LATEX_DISPLAY = re.compile(r"^\\begin\{(equation\*?|align\*?|gather\*?|multline\*?|aligned|cases)\}")
_MATH_NOTATION = re.compile(r"[∑∏∫∂√∞≤≥≠≈±]|\\(?:frac|sqrt|begin|end|sum|int|left|right)\b|\^\{?[-+0-9]")
_FORMULA_START = re.compile(r"^(?P<left>[\w.′'()\[\]{}+*/^_ \-]{1,35})\s*(?:=|≈|≠|≤|≥)\s*\S")
_TABLE_RULE = re.compile(r":?-{3,}:?")
_WIDE_GAP = re.compile(r" {2,}")


def _heading_key(text: str) -> str:
    return " ".join(text.strip().rstrip(":：").split()).casefold()


def _numbered_heading(text: str) -> tuple[tuple[int, ...], str] | None:
    match = _DECIMAL_HEADING.fullmatch(text)
    if match is None:
        return None
    return tuple(map(int, match["number"].split("."))), match["title"]


def _heading_like_text(text: str) -> bool:
    return (0 < len(text) <= 80 and not any(symbol in text for symbol in ".。!?！？;；,，|\t")
            and len(text.split()) <= 12 and any(character.isalpha() for character in text)
            and not _MATH_NOTATION.search(text) and "=" not in text)


def _heading_levels(lines: list[str]) -> dict[int, int]:
    headings: dict[int, int] = {}
    recognized_sections: set[tuple[int, ...]] = set()
    for index, line in enumerate(lines):
        if _heading_key(line) in _HEADINGS or _CHINESE_CHAPTER.fullmatch(line):
            headings[index] = 1
            continue
        numbered = _numbered_heading(line)
        if numbered is None:
            continue
        number, title = numbered
        known_title = _heading_key(title) in _HEADINGS
        # A short 2.1 heading under a recognized section is stronger evidence
        # than a free-standing “1. ...”, which is often a numbered list item.
        if known_title or (len(number) > 1 and number[:-1] in recognized_sections and _heading_like_text(title)):
            headings[index] = len(number)
            recognized_sections.add(number)
    return headings


def _title_index(lines: list[str], headings: dict[int, int]) -> int | None:
    nonempty = [index for index, line in enumerate(lines) if line]
    if not nonempty:
        return None
    first = nonempty[0]
    line = lines[first]
    if first in headings or _TITLE_METADATA.search(line) or _numbered_heading(line) is not None:
        return None
    if _LABEL_TITLE.match(line):
        return first
    # An adjacent standalone abstract heading is useful evidence. With
    # author/affiliation or journal matter in between, leave it for review.
    if (len(nonempty) >= 2 and _heading_key(lines[nonempty[1]]) in {"abstract", "摘要", "中文摘要"}
            and _heading_like_text(line) and sum(c.isalpha() for c in line) >= 6
            and (len(line.split()) >= 2 or re.search(r"[\u3400-\u9fff]", line))):
        return first
    return None


def _pipe_cells(line: str) -> list[str] | None:
    if "|" not in line or "\\|" in line or "`" in line:
        return None
    parts = line.removeprefix("|").removesuffix("|").split("|")
    if len(parts) < 2 or not any(part.strip() for part in parts):
        return None
    return [part.strip() for part in parts]


def _pipe_block(lines: list[str], start: int) -> tuple[int, list[list[str]] | None, int] | None:
    line = lines[start]
    first = _pipe_cells(line)
    next_cells = _pipe_cells(lines[start + 1]) if start + 1 < len(lines) else None
    divider = first is not None and next_cells is not None and all(_TABLE_RULE.fullmatch(cell) for cell in next_cells)
    bordered = line.startswith("|") and line.endswith("|") and line.count("|") >= 3
    if not (bordered or divider):
        return None
    end = start + 1
    rows = [first] if first is not None else []
    valid = first is not None
    while end < len(lines):
        line = lines[end]
        if "|" not in line or (bordered and not (line.startswith("|") and line.endswith("|"))):
            break
        cells = _pipe_cells(line)
        valid = valid and cells is not None and len(cells) == len(first)
        if cells is not None and not (divider and end == start + 1):
            rows.append(cells)
        end += 1
    if not divider and any(cell in {"=", "<", ">", "≤", "≥"} for row in rows for cell in row):
        valid = False
    return end - 1, rows if valid and len(rows) >= 2 else None, int(divider)


def _tab_cells(line: str) -> list[str] | None:
    if "\t" not in line or not line.strip():
        return None
    return [cell.strip() for cell in line.split("\t")]


def _tab_block(lines: list[str], start: int) -> tuple[int, list[list[str]] | None] | None:
    first = _tab_cells(lines[start])
    if first is None:
        return None
    rows = [first]
    end = start + 1
    while end < len(lines) and (cells := _tab_cells(lines[end])) is not None:
        rows.append(cells)
        end += 1
    valid = (len(rows) >= 2 and all(len(row) == len(first) for row in rows)
             and any(sum(bool(cell) for cell in row) >= 2 for row in rows)
             and not all(not row[0] for row in rows))
    return end - 1, rows if valid else None


def _possible_spaced_table(lines: list[str], start: int) -> int | None:
    width = len(_WIDE_GAP.split(lines[start]))
    if width < 2:
        return None
    end = start
    numeric_rows = 0
    while end < len(lines):
        cells = _WIDE_GAP.split(lines[end])
        if len(cells) != width or _MATH_NOTATION.search(lines[end]):
            break
        if any(re.fullmatch(r"[+\-−]?\d+(?:[.,]\d+)?(?:\s*%)?", cell) for cell in cells):
            numeric_rows += 1
        end += 1
        if end - start >= 4 and numeric_rows < 2:
            return None
    return end - 1 if end - start >= 2 and numeric_rows >= 2 else None


def _formula_end(lines: list[str], start: int) -> int | None:
    line = lines[start]
    if line.startswith("$$"):
        if line[2:].endswith("$$"):
            return start
        close = "$$"
    elif line.startswith(r"\["):
        close = r"\]"
    elif match := _LATEX_DISPLAY.match(line):
        close = rf"\end{{{match[1]}}}"
    else:
        equals = _FORMULA_START.match(line)
        if equals and len(equals["left"].split()) <= 3:
            return start
        if _MATH_NOTATION.search(line):
            return start
        return None
    # A missing closing delimiter must not claim the rest of the draft.
    for index in range(start, min(start + 80, len(lines))):
        candidate = lines[index][2:] if index == start and close == "$$" else lines[index]
        if candidate.endswith(close):
            return index
    return start


def propose_blocks(lines: list[str]) -> list[dict]:
    """Cover every zero-based source line once with a proposed block.

    Delimited tables carry proposed cell text. Heading/title suggestions retain
    their entire source line. Formula-like text and ambiguous tables are
    explicitly left for review. Whitespace trimming is only for classification.
    """
    classified = [line.strip() for line in lines]
    headings = _heading_levels(classified)
    title = _title_index(classified, headings)
    result = []
    index = 0
    while index < len(lines):
        kind = "paragraph"
        end = index
        detail = {}
        if (formula_end := _formula_end(classified, index)) is not None:
            kind, end = "review", formula_end
            detail = {"reason": "possible_formula"}
        elif (pipe := _pipe_block(classified, index)) is not None:
            end, rows, header_rows = pipe
            if rows is None:
                kind, detail = "review", {"reason": "ambiguous_table"}
            else:
                kind, detail = "table", {"rows": rows, "header_rows": header_rows}
        elif (tabs := _tab_block(lines, index)) is not None:
            end, rows = tabs
            if rows is None:
                kind, detail = "review", {"reason": "ambiguous_table"}
            else:
                kind, detail = "table", {"rows": rows, "header_rows": 0}
        elif (spaced_end := _possible_spaced_table(classified, index)) is not None:
            kind, end = "review", spaced_end
            detail = {"reason": "ambiguous_table"}
        elif index in headings:
            kind, detail = "heading", {"level": headings[index]}
        elif index == title:
            kind = "title"
        result.append({"start_line": index, "end_line": end, "kind": kind, **detail})
        index = end + 1
    return result
