#!/usr/bin/env python3
"""
Batch-transcribes scanned convict-register pages into an open Excel workbook,
using the locally installed Claude Code CLI (subscription login, vision
built in via the Read tool) -- no OCR engine, no separately billed API calls.

Append-only: each new page goes below every existing record. Errors on one
image are recorded and the run moves on to the next image.

Usage:
    ./run_batch.sh                 # process all not-yet-done images
    ./run_batch.sh --limit 5       # validation run on the first 5 images
"""
import argparse
import sys
from pathlib import Path

from claude_cli import call_claude_vision
from common import (
    add_selection_args,
    configure_stdout,
    list_images,
    load_config,
    load_progress,
    now_iso,
    record_rate_limit,
    save_progress,
    select_images,
)
from excel_sheet import (
    find_append_row,
    occupied_cells,
    open_workbook_with_backup,
    require_headers,
    set_cell,
)
from prompts import EXTRACTION_SCHEMA, PROMPT_TEMPLATE

REQUIRED_HEADERS = ["Crime", "Parish", "Home Address", "Status Notes", "File Name"]


def write_page(sheet, header_map: dict, start_row: int, rows: list, image_name: str) -> int:
    """Writes one page's records into consecutive rows; returns how many were
    flagged for review."""
    flagged = 0
    for i, r in enumerate(rows):
        row = start_row + i
        status = r.get("status") if r.get("status") in ("ok", "review") else "review"
        set_cell(sheet, row, header_map["Crime"], r.get("crime", ""))
        set_cell(sheet, row, header_map["Parish"], r.get("parish", ""))
        set_cell(sheet, row, header_map["Home Address"], r.get("home_address", ""))
        set_cell(sheet, row, header_map["File Name"], image_name)
        if status == "review":
            note_text = (r.get("notes") or "").strip()
            set_cell(sheet, row, header_map["Status Notes"], f"REVIEW: {note_text}" if note_text else "REVIEW")
            flagged += 1
        else:
            set_cell(sheet, row, header_map["Status Notes"], "OK")
    return flagged


def confirm_unfinished(progress: dict):
    """Crime-page recovery stays manual: a page cut off mid-write has to be
    cleaned up by hand, because re-running appends it again below."""
    stuck = [f for f, v in progress.items() if v.get("status") == "in_progress"]
    if not stuck:
        return
    print("WARNING: these images were left 'in_progress' by a previous run that did not")
    print("finish cleanly. They may have partial rows already written to Excel -- check")
    print("the workbook manually before continuing:")
    for f in stuck:
        print("  -", f)
    if input("Continue anyway? [y/N] ").strip().lower() != "y":
        sys.exit(1)


def stop_if_destination_occupied(sheet, header_map, start_row, rows, name, progress, progress_path):
    """Re-reads the destination right before writing: the sheet is live in
    Excel and may have been edited during the minutes-long extraction."""
    occupied = occupied_cells(sheet, header_map, start_row, len(rows))
    if not occupied:
        return
    shown = ", ".join(f"row {r} {h}" for r, h in occupied[:8])
    more = f" (+{len(occupied) - 8} more)" if len(occupied) > 8 else ""
    message = (f"destination rows {start_row}-{start_row + len(rows) - 1} are not empty: "
               f"{shown}{more}")
    print(f"STOPPED: {message}.")
    print("Nothing was written for this image. The sheet probably changed while the run "
          "was going (rows added/moved below the data). Check those rows, then re-run -- "
          "the append position is recalculated at the start of every run.")
    progress[name] = {"status": "error", "message": message, "timestamp": now_iso()}
    save_progress(progress_path, progress)
    sys.exit(1)


def write_and_save(sheet, book, header_map, start_row, rows, name, progress, progress_path) -> int:
    """Writes the page and saves the workbook; on failure leaves the image
    in_progress (for the manual check above) and exits."""
    try:
        flagged = write_page(sheet, header_map, start_row, rows, name)
        book.save()
        return flagged
    except Exception as exc:
        print(f"EXCEL WRITE ERROR: {exc}")
        print(f"Image '{name}' left marked in_progress -- check the workbook "
              f"around rows {start_row}-{start_row + len(rows) - 1} before re-running.")
        save_progress(progress_path, progress)
        sys.exit(1)


def process_images(todo, sheet, book, header_map, start_row, progress, progress_path, extract) -> dict:
    """The per-image loop. extract(img_path, on_row) returns a
    call_claude_vision()-style result dict. Returns {"processed", "flagged",
    "stopped"}."""
    current_row = start_row
    processed = 0
    flagged_total = 0

    for img_path in todo:
        name = img_path.name
        print(f"[{processed + 1}/{len(todo)}] {name} ...")
        progress[name] = {"status": "in_progress", "started": now_iso()}
        save_progress(progress_path, progress)

        def on_row(n, _img=name):
            print(f"    reading row {n} of {_img} ...")

        result = extract(img_path, on_row)

        if result["kind"] == "rate_limit":
            record_rate_limit(progress, progress_path, name, result["message"], processed)
            return {"processed": processed, "flagged": flagged_total, "stopped": True}

        if result["kind"] == "error":
            # Unlike the complexion pipeline, one bad page doesn't stop the run.
            print("ERROR -- flagged for manual review, continuing to next image.")
            progress[name] = {"status": "error", "message": result["message"], "timestamp": now_iso()}
            save_progress(progress_path, progress)
            processed += 1
            continue

        rows = result["rows"]
        stop_if_destination_occupied(sheet, header_map, current_row, rows, name, progress, progress_path)
        flagged_total += write_and_save(sheet, book, header_map, current_row, rows, name, progress, progress_path)

        # Marked done only after the workbook save above succeeded.
        flagged_here = sum(1 for r in rows if r.get("status") == "review")
        progress[name] = {
            "status": "done",
            "rows_written": len(rows),
            "excel_rows": [current_row, current_row + len(rows) - 1],
            "flagged": flagged_here,
            "timestamp": now_iso(),
        }
        save_progress(progress_path, progress)
        current_row += len(rows)
        processed += 1
        print(f"OK ({len(rows)} rows, {flagged_here} flagged for review)")

    return {"processed": processed, "flagged": flagged_total, "stopped": False}


def main():
    configure_stdout()
    ap = argparse.ArgumentParser()
    add_selection_args(ap, "config.json", "45.jpg")
    args = ap.parse_args()

    config = load_config(args.config)
    images_folder = Path(config["images_folder"])
    workbook_path = Path(config["workbook_path"])
    sheet_ref = config.get("sheet_name", 0)
    progress_path = Path(config["progress_file"])
    backup_dir = Path(config["backup_dir"])

    if not images_folder.is_dir():
        sys.exit(f"Images folder not found: {images_folder}")
    if not workbook_path.is_file():
        sys.exit(f"Workbook not found: {workbook_path}")

    progress = load_progress(progress_path)
    confirm_unfinished(progress)

    def is_done(p):
        return progress.get(p.name, {}).get("status") == "done"

    images = list_images(images_folder)
    todo = select_images(images, is_done, args.only, args.limit, images_folder,
                         f"'{{name}}' is already marked 'done' in {progress_path}. "
                         f"To force a redo, remove its entry from that file first.")
    if not todo:
        print("Nothing to do -- all images already processed.")
        return

    print(f"Processing {len(todo)} of {len(images)} image(s) "
          f"(already done: {sum(1 for p in images if is_done(p))})")

    book = open_workbook_with_backup(workbook_path, backup_dir)
    sheet = book.sheets[sheet_ref]
    header_map = require_headers(sheet, REQUIRED_HEADERS)

    try:
        start_row = find_append_row(sheet, header_map)
    except RuntimeError as exc:
        sys.exit(str(exc))
    print(f"Appending new records starting at row {start_row} "
          f"(first row below every filled named column).")

    def extract(img_path, on_row):
        return call_claude_vision(img_path, images_folder, EXTRACTION_SCHEMA, PROMPT_TEMPLATE, on_row=on_row)

    summary = process_images(todo, sheet, book, header_map, start_row, progress, progress_path, extract)
    if summary["stopped"]:
        return

    print(f"\nDone. Processed {summary['processed']} image(s) this run. "
          f"Total rows flagged for review: {summary['flagged']}.")
    print(f"Workbook saved: {workbook_path}")


if __name__ == "__main__":
    main()
