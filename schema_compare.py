"""Read the tab/column structure of an Excel workbook and compare it with a reference."""

from __future__ import annotations

import difflib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import IO, Union

import openpyxl
from openpyxl.utils import get_column_letter

REFERENCE_PATH = Path(__file__).with_name("reference_schema.json")

_D = r"\d{1,2}[./_-]\d{1,2}(?:[./_-]\d{2,4})?"
_ISO = r"\d{4}[./_-]\d{1,2}[./_-]\d{1,2}"
_ONE = rf"(?:{_ISO}|{_D})"
DATE_PATTERN = re.compile(
    rf"(?<![\d.\-/_])(?:{_ONE}\s*(?:-|–|—|to)\s*{_ONE}|{_ONE})(?![\d.\-/_])",
    re.IGNORECASE,
)


def normalize_tab_name(name: str) -> str:
    """Remove dates / date ranges (e.g. '11.9.26', '11.1-11.7.26') from a tab name."""
    cleaned = DATE_PATTERN.sub(" ", str(name))
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" -_.()[]")


def _header_value(value) -> str:
    if value is None:
        return ""
    return str(value)


def read_structure(source: Union[str, Path, IO[bytes]]) -> list[dict]:
    """Return [{'name', 'key', 'state', 'columns'}] for every tab, in workbook order.

    Columns are taken from row 1. Blank header cells inside the row are kept as ''
    so positions are preserved; trailing blank cells are dropped.
    """
    wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    try:
        tabs = []
        for ws in wb.worksheets:
            columns: list[str] = []
            if hasattr(ws, "reset_dimensions"):
                # The stored dimension tag can be wrong; read every cell actually present in row 1.
                ws.reset_dimensions()
                rows = ws.iter_rows(min_row=1, max_row=1)
            else:
                rows = ws.iter_rows(min_row=1, max_row=1, max_col=ws.max_column)
            for row in rows:
                for cell in row:
                    col_idx = getattr(cell, "column", None)
                    if col_idx is None:
                        continue
                    while len(columns) < col_idx - 1:
                        columns.append("")
                    columns.append(_header_value(cell.value))
            while columns and columns[-1].strip() == "":
                columns.pop()
            tabs.append(
                {
                    "name": ws.title,
                    "key": normalize_tab_name(ws.title),
                    "state": ws.sheet_state,
                    "columns": columns,
                }
            )
        return tabs
    finally:
        wb.close()


def load_reference() -> dict:
    with open(REFERENCE_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def save_reference(source_file: Union[str, Path]) -> dict:
    ref = {"source_file": Path(source_file).name, "tabs": read_structure(source_file)}
    with open(REFERENCE_PATH, "w", encoding="utf-8") as fh:
        json.dump(ref, fh, indent=2, ensure_ascii=False)
    return ref


@dataclass
class Difference:
    tab: str
    kind: str
    detail: str
    location: str = ""
    expected: str = ""
    found: str = ""


@dataclass
class TabResult:
    reference_name: str | None
    uploaded_name: str | None
    status: str
    differences: list[Difference] = field(default_factory=list)
    column_rows: list[dict] = field(default_factory=list)


def _display(header: str) -> str:
    return "(blank)" if header.strip() == "" else header


def _col(i: int) -> str:
    return f"Column {get_column_letter(i + 1)} (#{i + 1})"


def _loose(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().casefold()


def _moved(expected_tags: list, found_tags: list) -> list:
    """Items present in both lists whose relative order changed (minimal set via SequenceMatcher)."""
    common = set(expected_tags) & set(found_tags)
    e = [t for t in expected_tags if t in common]
    f = [t for t in found_tags if t in common]
    moved = []
    for op, i1, i2, _j1, _j2 in difflib.SequenceMatcher(None, e, f, autojunk=False).get_opcodes():
        if op in ("delete", "replace"):
            moved.extend(e[i1:i2])
    return moved


def compare_columns(tab_label: str, expected: list[str], found: list[str]):
    """Return (differences, side_by_side_rows) for one tab.

    Headers are aligned by position (longest matching runs). Whatever is left unaligned is
    classified once each: out of order (same name elsewhere), renamed (different name in
    the same spot), missing (reference only) or extra (upload only).
    """
    blocks = []  # (unmatched expected indexes, unmatched found indexes) per changed region
    sm = difflib.SequenceMatcher(None, expected, found, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op != "equal":
            blocks.append((list(range(i1, i2)), list(range(j1, j2))))

    moved: list[tuple[int, int]] = []
    for b_e, _ in blocks:
        for i in list(b_e):
            if expected[i].strip() == "":
                continue
            for _, b_f in blocks:
                j = next((j for j in b_f if found[j] == expected[i]), None)
                if j is not None:
                    moved.append((i, j))
                    b_e.remove(i)
                    b_f.remove(j)
                    break

    renamed: list[tuple[int, int]] = []
    missing: list[int] = []
    extra: list[int] = []
    for b_e, b_f in blocks:
        n = min(len(b_e), len(b_f))
        renamed.extend(zip(b_e[:n], b_f[:n]))
        missing.extend(b_e[n:])
        extra.extend(b_f[n:])

    diffs: list[Difference] = []
    for i, j in renamed:
        exp, got = expected[i], found[j]
        kind = ("Column name differs (case/spacing only)" if _loose(exp) == _loose(got)
                else "Column name changed")
        where = f"at {_col(j)}" if i == j else f"at {_col(j)} (reference {_col(i)})"
        diffs.append(Difference(
            tab_label, kind,
            f"Expected '{_display(exp)}' but found '{_display(got)}' {where}",
            _col(j), _display(exp), _display(got)))
    for i in missing:
        diffs.append(Difference(
            tab_label, "Missing column",
            f"'{_display(expected[i])}' (reference {_col(i)}) is not in the uploaded tab",
            _col(i), _display(expected[i]), "(not present)"))
    for j in extra:
        diffs.append(Difference(
            tab_label, "Extra column",
            f"'{_display(found[j])}' at {_col(j)} is not in the reference tab",
            _col(j), "(not expected)", _display(found[j])))
    for i, j in moved:
        diffs.append(Difference(
            tab_label, "Column out of order",
            f"'{expected[i]}' is at {_col(j)} but the reference has it at {_col(i)}",
            _col(j), f"{expected[i]} @ {_col(i)}", f"{found[j]} @ {_col(j)}"))

    diffs.sort(key=lambda d: _sort_key(d.location))

    rows = []
    for k in range(max(len(expected), len(found))):
        exp = expected[k] if k < len(expected) else None
        got = found[k] if k < len(found) else None
        rows.append({
            "Position": _col(k),
            "Expected (reference)": "— (none)" if exp is None else _display(exp),
            "Found (uploaded)": "— (none)" if got is None else _display(got),
            "Same?": "✅" if exp == got else "❌",
        })
    return diffs, rows


def _sort_key(location: str):
    m = re.search(r"#(\d+)", location)
    return int(m.group(1)) if m else 0


def compare(reference_tabs: list[dict], uploaded_tabs: list[dict]) -> dict:
    results: list[TabResult] = []
    all_diffs: list[Difference] = []
    used: set[int] = set()

    pairs: list[tuple[dict, dict | None, bool]] = []
    for ref in reference_tabs:
        up = next((t for i, t in enumerate(uploaded_tabs)
                   if i not in used and t["key"] == ref["key"]), None)
        loose = False
        if up is None:
            up = next((t for i, t in enumerate(uploaded_tabs)
                       if i not in used and _loose(t["key"]) == _loose(ref["key"])), None)
            loose = up is not None
        if up is not None:
            used.add(uploaded_tabs.index(up))
        pairs.append((ref, up, loose))

    # Tab order among tabs present in both files.
    ref_seq = [(ref["key"], i) for i, (ref, up, _) in enumerate(pairs) if up is not None]
    up_index = {id(up): uploaded_tabs.index(up) for _, up, _ in pairs if up is not None}
    up_seq = sorted(ref_seq, key=lambda t: up_index[id(pairs[t[1]][1])])
    out_of_order = set(_moved(ref_seq, up_seq))

    for i, (ref, up, loose) in enumerate(pairs):
        label = ref["key"]
        if up is None:
            d = Difference(label, "Missing tab",
                           f"Tab '{ref['key']}' (reference tab '{ref['name']}', position {i + 1}) "
                           f"is not in the uploaded file",
                           f"Tab position {i + 1}", ref["key"], "(not present)")
            all_diffs.append(d)
            results.append(TabResult(ref["name"], None, "Missing tab", [d]))
            continue

        tab_diffs: list[Difference] = []
        if loose:
            tab_diffs.append(Difference(
                label, "Tab name differs (case/spacing only)",
                f"Expected tab '{ref['key']}' but found '{up['key']}' (actual tab name '{up['name']}')",
                "Tab name", ref["key"], up["key"]))
        if (ref["key"], i) in out_of_order:
            pos = up_index[id(up)] + 1
            tab_diffs.append(Difference(
                label, "Tab out of order",
                f"Tab '{ref['key']}' is at position {pos} in the uploaded file "
                f"but at position {i + 1} in the reference",
                "Tab position", f"position {i + 1}", f"position {pos}"))
        col_diffs, rows = compare_columns(label, ref["columns"], up["columns"])
        tab_diffs.extend(col_diffs)
        all_diffs.extend(tab_diffs)
        results.append(TabResult(ref["name"], up["name"],
                                 "Match" if not tab_diffs else "Different", tab_diffs, rows))

    for i, up in enumerate(uploaded_tabs):
        if i in used:
            continue
        label = up["key"] or up["name"]
        d = Difference(label, "Extra tab",
                       f"Tab '{up['name']}' (position {i + 1}) is not in the reference file",
                       f"Tab position {i + 1}", "(not expected)", up["name"])
        all_diffs.append(d)
        _, rows = compare_columns(label, [], up["columns"])
        results.append(TabResult(None, up["name"], "Extra tab", [d], rows))

    return {"tabs": results, "differences": all_diffs}
