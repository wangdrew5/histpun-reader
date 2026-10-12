"""Worksheet helpers for the two workbook pipelines.

Everything except opening the workbook works on any object with xlwings'
range API (the tests pass tests/fake_sheet.py). xlwings itself is imported
only when a real workbook is opened, so importing this module never starts
Excel automation.
"""
import sys
from pathlib import Path

from common import backup_workbook

ROW_SCAN_LIMIT = 20000


def get_or_open_workbook(path: Path):
    import xlwings as xw  # deferred: only needed once Excel is actually driven

    resolved = str(path.resolve())
    for app in xw.apps:
        for book in app.books:
            if book.fullname and str(Path(book.fullname).resolve()) == resolved:
                return app, book
    book = xw.Book(str(path))
    return book.app, book


def open_workbook_with_backup(workbook_path: Path, backup_dir: Path):
    """Attach to (or open) the workbook, then save and back it up.

    Saving comes first: the live workbook can have unsaved edits that exist
    only in Excel's memory, and the backup must include them."""
    _app, book = get_or_open_workbook(workbook_path)
    book.save()
    backup_path = backup_workbook(workbook_path, backup_dir)
    print(f"Backed up workbook to: {backup_path}")
    return book


def get_header_map(sheet) -> dict:
    values = sheet.range("A1").expand("right").value
    if not isinstance(values, list):
        values = [values]
    return {str(v).strip(): i + 1 for i, v in enumerate(values) if v}


def is_blank(v) -> bool:
    return v is None or v == ""


def read_block(sheet, first_row: int, first_col: int, last_row: int, last_col: int) -> list:
    """Always returns a list of row lists, even for a single cell/row/column
    (plain xlwings .value collapses those to a scalar or flat list)."""
    if last_row < first_row:
        return []
    return sheet.range((first_row, first_col), (last_row, last_col)).options(ndim=2).value


def find_append_row(sheet, header_map: dict) -> int:
    """First row below every existing record, judged by ALL named columns.

    Anchoring on one column (it used to be Crime) lets a final record whose
    Crime is blank -- but whose Name, File Name, Status Notes, etc. are filled
    -- look like free space and get overwritten.

    Excel's used range can't be trusted as the answer (Google Sheets exports
    carry formatting far below the real data), but it is always an upper
    bound on where data can be, so it limits how much has to be read.
    """
    ncols = max(header_map.values())
    last_used = sheet.used_range.last_cell.row
    scan_to = min(last_used, ROW_SCAN_LIMIT)
    data_cols = sorted(header_map.values())
    last_record = 1  # the header row
    for offset, row in enumerate(read_block(sheet, 2, 1, scan_to, ncols)):
        if any(not is_blank(row[c - 1]) for c in data_cols):
            last_record = 2 + offset
    if last_used > ROW_SCAN_LIMIT and last_record == ROW_SCAN_LIMIT:
        raise RuntimeError(f"Data reaches row {ROW_SCAN_LIMIT} (ROW_SCAN_LIMIT); can't confirm "
                           f"where the last record is. Raise ROW_SCAN_LIMIT in process_reports.py.")
    return last_record + 1


def occupied_cells(sheet, header_map: dict, first_row: int, n_rows: int) -> list:
    """(row, header) for every non-blank named cell in the n_rows starting at
    first_row. Any hit means writing there would overwrite someone's record."""
    if n_rows <= 0:
        return []
    names_by_col = {c: name for name, c in header_map.items()}
    block = read_block(sheet, first_row, 1, first_row + n_rows - 1, max(header_map.values()))
    return [
        (first_row + i, names_by_col[c])
        for i, row in enumerate(block)
        for c in sorted(names_by_col)
        if not is_blank(row[c - 1])
    ]


def set_cell(sheet, row, col, value):
    sheet.range((row, col)).value = value


def require_headers(sheet, required: list) -> dict:
    header_map = get_header_map(sheet)
    missing = [h for h in required if h not in header_map]
    if missing:
        sys.exit(f"Sheet is missing expected header column(s): {missing}. "
                 f"Found headers: {list(header_map.keys())}")
    return header_map
