#!/usr/bin/env python3
"""
Fills in the "Race" column from the OTHER (left-hand) half of each ledger
page -- the half that shows Complexion -- by pairing each new complexion
image with its crime-pipeline partner BY NUMBER (e.g. "74.jpg" pairs with
"75.jpg"), using the row range already recorded for that partner by
process_reports.py.

This does NOT scan the sheet for blank cells and guess where to put things,
and it does NOT depend on processing order -- each complexion image declares
its own partner via filename number, so images can be dropped in and
processed in any order.

The row range recorded in progress.json is never trusted on its own: before
extracting, and again right before writing, the worksheet's File Name column
must show the partner's filename on exactly those rows (and not on the rows
just outside them). If rows were inserted, deleted, or sorted since, the run
stops and says where that page's rows actually are now.

Writes are resumable. An existing Race value that agrees with the new
reading is kept; blank cells are filled; any disagreement stops the run
without writing that page. A page's extraction is cached in
progress_complexion.json until the page is done, so stopping and re-running
never pays to read the same image twice.

Usage:
    ./run_complexion_batch.sh                 # process all new complexion images
    ./run_complexion_batch.sh --limit 5        # validation run on the first 5
"""
import argparse
import sys
from pathlib import Path

from claude_cli import call_claude_vision
from common import (
    add_selection_args,
    configure_stdout,
    extract_number,
    list_images,
    load_config,
    load_progress,
    now_iso,
    record_rate_limit,
    save_progress,
    select_images,
)
from excel_sheet import (
    ROW_SCAN_LIMIT,
    is_blank,
    open_workbook_with_backup,
    read_block,
    require_headers,
    set_cell,
)
from prompts import COMPLEXION_PROMPT_TEMPLATE, COMPLEXION_SCHEMA

REQUIRED_HEADERS = ["Race", "Status Notes", "File Name"]


class AlignmentError(Exception):
    """The recorded row range can't be confirmed against the live worksheet."""


def _label(v) -> str:
    return "" if is_blank(v) else str(v).strip()


def _same_value(a, b) -> bool:
    return " ".join(str(a).split()).casefold() == " ".join(str(b).split()).casefold()


def _ranges(row_numbers: list) -> str:
    """[5, 6, 7, 9] -> '5-7, 9'"""
    parts, start, prev = [], None, None
    for n in row_numbers + [None]:
        if start is not None and (n is None or n != prev + 1):
            parts.append(f"{start}-{prev}" if prev != start else str(start))
            start = None
        if n is not None and start is None:
            start = n
        prev = n
    return ", ".join(parts)


def file_name_runs(sheet, header_map: dict, filename: str) -> list:
    """Every contiguous block of rows whose File Name is `filename`, as (first, last)."""
    col = header_map["File Name"]
    scan_to = min(sheet.used_range.last_cell.row, ROW_SCAN_LIMIT)
    rows = [n for n, row in enumerate(read_block(sheet, 2, col, scan_to, col), start=2)
            if _label(row[0]) == filename]
    runs = []
    for n in rows:
        if runs and runs[-1][1] == n - 1:
            runs[-1][1] = n
        else:
            runs.append([n, n])
    return [tuple(r) for r in runs]


def verify_alignment(sheet, header_map: dict, target: dict) -> tuple:
    """Confirms the partner page's recorded rows are still exactly where
    progress.json says: every row in the range carries the partner's
    filename, the range size equals the recorded row count, and the page
    doesn't continue past either end. Returns (first_row, last_row) or raises
    AlignmentError with what's actually in the sheet and how to fix it."""
    fname = target["filename"]
    start, end = target["excel_rows"]
    expected = target["rows_written"]
    col = header_map["File Name"]
    problems = []

    if end - start + 1 != expected:
        problems.append(f"that range spans {end - start + 1} rows but {expected} were recorded")
    if start < 2 or end < start:
        problems.append("that isn't a valid data range")
    else:
        lo = max(2, start - 1)
        labels = [_label(r[0]) for r in read_block(sheet, lo, col, end + 1, col)]
        inside = labels[start - lo:start - lo + (end - start + 1)]
        wrong = {}
        for i, v in enumerate(inside):
            if v != fname:
                wrong.setdefault(v or "(blank)", []).append(start + i)
        for v, rows in wrong.items():
            problems.append(f"File Name says {v} on rows {_ranges(rows)}")
        if start - 1 >= 2 and labels[0] == fname:
            problems.append(f"row {start - 1}, just above the range, is also {fname}")
        if labels[-1] == fname:
            problems.append(f"row {end + 1}, just below the range, is also {fname}")

    if not problems:
        return start, end

    lines = [f"progress.json records {fname} at rows {start}-{end} ({expected} rows), but "
             + "; ".join(problems) + "."]
    runs = file_name_runs(sheet, header_map, fname)
    if len(runs) == 1:
        a, b = runs[0]
        lines.append(f"The File Name column shows {fname} at rows {a}-{b} ({b - a + 1} rows).")
        if b - a + 1 == expected:
            lines.append(f"Rows were probably inserted, deleted, or sorted after {fname} was written. "
                         f"If rows {a}-{b} are the right records, set {fname}'s \"excel_rows\" in "
                         f"progress.json to [{a}, {b}] and re-run.")
        else:
            lines.append(f"That's {b - a + 1} rows, not the {expected} recorded, so rows were added to or "
                         f"removed from this page in the sheet. Check them against the image before "
                         f"updating {fname}'s \"excel_rows\" and \"rows_written\" in progress.json.")
    elif not runs:
        lines.append(f"No row's File Name is {fname}. Check the page's rows against the image, then "
                     f"fill in File Name for them or correct progress.json.")
    else:
        where = ", ".join(f"{a}-{b}" if a != b else str(a) for a, b in runs)
        lines.append(f"{fname} appears in {len(runs)} separate places ({where}), probably from a sort "
                     f"or a partial move. Put the page's rows back together, then update progress.json.")
    raise AlignmentError("\n".join(lines))


def plan_race_writes(existing: list, rows: list) -> dict:
    """Compares the Race cells already in the sheet with a fresh reading.
    Indexes into rows: "write" = blank cell, fill it; "agree" = same value
    already there; "kept" = cell filled but the reading is blank (illegible),
    so the existing value stands; "conflicts" = (index, existing, reading)
    where both are filled and differ."""
    plan = {"write": [], "agree": [], "kept": [], "conflicts": []}
    for i, (old, r) in enumerate(zip(existing, rows)):
        new = (r.get("complexion") or "").strip()
        if is_blank(old) or not str(old).strip():
            plan["write"].append(i)
        elif not new:
            plan["kept"].append(i)
        elif _same_value(old, new):
            plan["agree"].append(i)
        else:
            plan["conflicts"].append((i, old, new))
    return plan


def review_note_value(existing, r: dict):
    """New Status Notes value with this row's complexion note appended, or
    None if the note is already there (so resuming never doubles it)."""
    note_text = (r.get("notes") or "").strip()
    note = f"complexion: {note_text}" if note_text else "complexion uncertain"
    existing = "" if is_blank(existing) else str(existing).strip()
    if note in existing:
        return None
    if not existing or existing.upper() == "OK":
        return f"REVIEW: {note}"
    if existing.upper().startswith("REVIEW"):
        return f"{existing}; {note}"
    return f"{existing}; REVIEW: {note}"


def _row_status(r: dict) -> str:
    return r.get("status") if r.get("status") in ("ok", "review") else "review"


# --- Stages of one page, in the order process_page() runs them ---------------

def find_partner(img_path: Path, crime_by_number: dict):
    """The crime-pipeline page this complexion image pairs with (its number
    plus one), or None (printed) if there isn't one yet."""
    my_number = extract_number(img_path.stem)
    if my_number is None:
        print(f"[skip] {img_path.name}: couldn't find a number in this filename to pair with.")
        return None
    target = crime_by_number.get(my_number + 1)
    if target is None:
        print(f"[skip] {img_path.name}: partner page {my_number + 1} hasn't been processed by "
              f"the offence/parish/residence pipeline yet. Will retry next run.")
    return target


def cached_reading(prev: dict, target: dict):
    """A reading saved by an earlier run, reused only for the same partner."""
    rows = prev.get("extracted") if prev.get("paired_with") == target["filename"] else None
    return rows if isinstance(rows, list) else None


def check_alignment(sheet, header_map, target, name, progress, progress_path, rows=None) -> bool:
    """verify_alignment(), recording a "misaligned" stop (keeping any reading
    already made, since it's still valid once things are fixed)."""
    try:
        verify_alignment(sheet, header_map, target)
        return True
    except AlignmentError as exc:
        print(f"STOPPED: can't confirm where '{target['filename']}' is in the sheet. "
              f"Nothing was written for {name}.")
        print(str(exc))
        entry = {"status": "misaligned", "paired_with": target["filename"],
                 "message": str(exc), "timestamp": now_iso()}
        if rows is not None:
            entry["extracted"] = rows
        progress[name] = entry
        save_progress(progress_path, progress)
        return False


def read_page(img_path, target, extract, progress, progress_path, processed):
    """Sends the image to Claude. Returns its rows, or None after recording
    why the run has to stop (usage limit or error)."""
    name = img_path.name
    progress[name] = {"status": "in_progress", "started": now_iso(), "paired_with": target["filename"]}
    save_progress(progress_path, progress)

    def on_row(n, _img=name):
        print(f"    reading row {n} of {_img} ...")

    result = extract(img_path, on_row)
    if result["kind"] == "rate_limit":
        record_rate_limit(progress, progress_path, name, result["message"], processed)
        return None
    if result["kind"] == "error":
        print(f"ERROR: {result['message'][:500]}")
        print("Stopping the run here (rather than guessing the next pairing) so you can "
              "investigate this image before continuing.")
        progress[name] = {"status": "error", "message": result["message"], "timestamp": now_iso()}
        save_progress(progress_path, progress)
        return None
    return result["rows"]


def check_row_count(rows, target, name, progress, progress_path) -> bool:
    """A miscount is recorded without caching the reading: re-reading the
    image is how it gets fixed."""
    expected = target["rows_written"]
    if len(rows) == expected:
        return True
    start_row, end_row = target["excel_rows"]
    print(f"MISMATCH: '{name}' produced {len(rows)} row(s), but its paired page "
          f"'{target['filename']}' had {expected} row(s) (Excel rows {start_row}-{end_row}).")
    print("Not writing anything for this image. Stopping so you can check whether this is "
          "the right image / right order before continuing.")
    progress[name] = {
        "status": "error",
        "message": f"row count mismatch: got {len(rows)}, expected {expected} "
                   f"(paired with {target['filename']})",
        "timestamp": now_iso(),
    }
    save_progress(progress_path, progress)
    return False


def save_reading(rows, target, name, prev, progress, progress_path):
    """Saved before any cell is touched: if the run dies mid-page, the next
    run reuses this reading and finishes the page instead of re-reading."""
    progress[name] = {
        "status": "in_progress",
        "started": prev.get("started") or now_iso(),
        "paired_with": target["filename"],
        "excel_rows": list(target["excel_rows"]),
        "extracted": rows,
    }
    save_progress(progress_path, progress)


def record_conflicts(plan, rows, target, name, progress, progress_path):
    start_row, end_row = target["excel_rows"]
    print(f"CONFLICT: {len(plan['conflicts'])} Race cell(s) for '{target['filename']}' already "
          f"hold a different value than this reading of {name}:")
    for i, old, new in plan["conflicts"]:
        print(f"    row {start_row + i}: sheet has {old!r}, image reads {new!r}")
    print("Nothing was written for this page. Decide which is right, then either fix or clear "
          "those Race cells and re-run (the saved reading is reused, no new extraction), or, "
          f"if the sheet is already right, mark {name} as done in progress_complexion.json.")
    progress[name] = {
        "status": "conflict",
        "paired_with": target["filename"],
        "excel_rows": [start_row, end_row],
        "conflicts": [{"row": start_row + i, "sheet": old, "image": new}
                      for i, old, new in plan["conflicts"]],
        "extracted": rows,
        "timestamp": now_iso(),
    }
    save_progress(progress_path, progress)


def write_race_cells(sheet, header_map, start_row, end_row, rows, plan) -> int:
    """Fills blank Race cells, appends review notes once each, and returns
    how many rows were flagged. Existing values (agreeing, or kept because
    the reading was illegible) are never written to."""
    race_col = header_map["Race"]
    status_notes_col = header_map["Status Notes"]
    notes = [r[0] for r in read_block(sheet, start_row, status_notes_col, end_row, status_notes_col)]
    to_write = set(plan["write"])
    kept = set(plan["kept"])
    flagged = 0
    for i, r in enumerate(rows):
        row = start_row + i
        if i in to_write:
            set_cell(sheet, row, race_col, r.get("complexion", ""))
        if i in kept or _row_status(r) != "review":
            continue
        flagged += 1
        new_note = review_note_value(notes[i], r)
        if new_note is not None:
            set_cell(sheet, row, status_notes_col, new_note)
    return flagged


def mark_done(plan, rows, flagged, target, name, progress, progress_path):
    start_row, end_row = target["excel_rows"]
    entry = {
        "status": "done",
        "paired_with": target["filename"],
        "excel_rows": [start_row, end_row],
        "rows_written": len(rows),
        "flagged": flagged,
        "timestamp": now_iso(),
    }
    present = len(plan["agree"]) + len(plan["kept"])
    if present:
        entry["already_present"] = present
    progress[name] = entry
    save_progress(progress_path, progress)
    extra = f", {present} already filled and kept" if present else ""
    print(f"OK ({len(plan['write'])} rows filled into {start_row}-{end_row}{extra}, "
          f"{flagged} flagged for review)")


def process_page(img_path, target, sheet, book, header_map, progress, progress_path, extract, processed):
    """One complexion page. Returns the number of rows flagged for review, or
    None if the run has to stop (the reason is already printed and saved)."""
    name = img_path.name
    start_row, end_row = target["excel_rows"]
    prev = progress.get(name, {})
    cached = cached_reading(prev, target)

    # 1. Rows still where progress.json says? Checked before extracting, so a
    #    misaligned page costs no usage.
    if not check_alignment(sheet, header_map, target, name, progress, progress_path, cached):
        return None

    # 2. Read the image, or reuse the reading an earlier run saved.
    if cached is not None:
        rows = cached
        print(f"    using the reading saved by a previous run ({len(rows)} rows) -- not re-reading the image")
    else:
        rows = read_page(img_path, target, extract, progress, progress_path, processed)
        if rows is None:
            return None

    # 3. Same number of people as the partner page.
    if not check_row_count(rows, target, name, progress, progress_path):
        return None

    # 4. Save the reading before touching any cell.
    save_reading(rows, target, name, prev, progress, progress_path)

    # 5. Alignment again: the workbook is live and extraction takes minutes.
    if not check_alignment(sheet, header_map, target, name, progress, progress_path, rows):
        return None

    # 6. Compare the whole page with what's already there; any conflict
    #    stops before a single cell is written.
    race_col = header_map["Race"]
    existing = [r[0] for r in read_block(sheet, start_row, race_col, end_row, race_col)]
    plan = plan_race_writes(existing, rows)
    if plan["conflicts"]:
        record_conflicts(plan, rows, target, name, progress, progress_path)
        return None

    # 7. Write, save the workbook, and only then mark the page done.
    flagged = write_race_cells(sheet, header_map, start_row, end_row, rows, plan)
    book.save()
    mark_done(plan, rows, flagged, target, name, progress, progress_path)
    return flagged


def process_images(todo, sheet, book, header_map, crime_by_number, progress, progress_path, extract) -> dict:
    """The per-image loop. extract(img_path, on_row) returns a
    call_claude_vision()-style result dict. Unlike the crime pipeline, any
    problem stops the whole run. Returns {"processed", "flagged", "stopped"}."""
    processed = 0
    flagged_total = 0
    for img_path in todo:
        target = find_partner(img_path, crime_by_number)
        if target is None:
            continue
        start_row, end_row = target["excel_rows"]
        print(f"[{processed + 1}/{len(todo)}] {img_path.name} -> pairing with '{target['filename']}' "
              f"(rows {start_row}-{end_row}, expecting {target['rows_written']}) ...")
        flagged = process_page(img_path, target, sheet, book, header_map,
                               progress, progress_path, extract, processed)
        if flagged is None:
            return {"processed": processed, "flagged": flagged_total, "stopped": True}
        flagged_total += flagged
        processed += 1
    return {"processed": processed, "flagged": flagged_total, "stopped": False}


def crime_pages_by_number(crime_progress: dict) -> dict:
    crime_by_number = {}
    for name, entry in crime_progress.items():
        if entry.get("status") != "done" or "excel_rows" not in entry:
            continue
        num = extract_number(Path(name).stem)
        if num is not None:
            crime_by_number[num] = {"filename": name, **entry}
    return crime_by_number


def main():
    configure_stdout()
    ap = argparse.ArgumentParser()
    add_selection_args(ap, "config_complexion.json", "74.jpg")
    args = ap.parse_args()

    config = load_config(args.config)
    images_folder = Path(config["images_folder"])
    workbook_path = Path(config["workbook_path"])
    sheet_ref = config.get("sheet_name", 0)
    crime_progress_path = Path(config["crime_progress_file"])
    progress_path = Path(config["progress_file"])
    backup_dir = Path(config["backup_dir"])

    if not images_folder.is_dir():
        sys.exit(f"Images folder not found: {images_folder}")
    if not workbook_path.is_file():
        sys.exit(f"Workbook not found: {workbook_path}")
    if not crime_progress_path.is_file():
        sys.exit(f"Crime-pipeline progress file not found: {crime_progress_path}")

    crime_by_number = crime_pages_by_number(load_progress(crime_progress_path))
    progress = load_progress(progress_path)

    resuming = [f for f, v in progress.items()
                if isinstance(v, dict) and v.get("status") in ("in_progress", "conflict", "misaligned")]
    if resuming:
        print("Unfinished from a previous run (checked against the sheet before anything is written):")
        for f in resuming:
            print(f"  - {f} ({progress[f]['status']})")

    images = list_images(images_folder)
    todo = select_images(images, lambda p: progress.get(p.name, {}).get("status") == "done",
                         args.only, args.limit, images_folder,
                         f"'{{name}}' is already marked 'done' in {progress_path}. "
                         f"To force a redo, remove its entry from that file first.")
    if not todo:
        print("Nothing to do -- all complexion images already processed.")
        return

    print(f"Processing {len(todo)} of {len(images)} complexion image(s). "
          f"{len(crime_by_number)} crime-pipeline page(s) available to pair against.")

    book = open_workbook_with_backup(workbook_path, backup_dir)
    sheet = book.sheets[sheet_ref]
    header_map = require_headers(sheet, REQUIRED_HEADERS)

    def extract(img_path, on_row):
        return call_claude_vision(img_path, images_folder, COMPLEXION_SCHEMA,
                                  COMPLEXION_PROMPT_TEMPLATE, on_row=on_row)

    summary = process_images(todo, sheet, book, header_map, crime_by_number,
                             progress, progress_path, extract)
    if summary["stopped"]:
        return

    print(f"\nDone. Processed {summary['processed']} image(s) this run. "
          f"Total rows flagged for review: {summary['flagged']}.")
    print(f"Workbook saved: {workbook_path}")


if __name__ == "__main__":
    main()
