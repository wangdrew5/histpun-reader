# How this pipeline works

Two independent scripts fill in your convict-register spreadsheet from scanned
ledger pages (a third only exports complexion readings to their own file), using the `claude` CLI's built-in vision (your subscription
login — no OCR engine, no separate API billing):

| Script | Reads from | Writes to | Command |
|---|---|---|---|
| `scripts/process_reports.py` | `images/` (right-hand pages: Offence/Parish/Residence) | Crime, Parish, Home Address, Status Notes, File Name | `./run_batch.sh` |
| `scripts/process_complexion.py` | `images_complexion/` (left-hand pages: Name/Age/Complexion) | Race, Status Notes (appended) | `./run_complexion_batch.sh` |
| `scripts/export_complexion.py` | `images_complexion/` | a separate `complexion_export.xlsx` (Image, Complexion) — never the main workbook | `./run_complexion_export.sh` |

The three scripts are just entry points. What they share lives in four
small modules next to them (see the file map), and no script imports
another. They all follow the same safety philosophy: **never guess, never
overwrite real data, stop rather than silently misalign.**

On purpose, they handle a failed image differently:

| | crime (`run_batch.sh`) | complexion (`run_complexion_batch.sh`) | export (`run_complexion_export.sh`) |
|---|---|---|---|
| Claude error on one image | record it, **continue** | **stop** the run | skip it; **stop after two in a row** |
| Crash in the middle of writing a page | **manual**: y/N prompt next run, clean up by hand | **automatic**: resumes from the saved reading | n/a (saves after each image) |

---

## 1. The shared modules

### Reading an image with Claude (`claude_cli.py`)

`call_claude_vision(image_path, add_dir, schema, prompt_template, on_row=...)`
is the one function all three scripts use to actually talk to Claude. It:

1. **Downscales the image first.** Full-resolution scans from your camera are
   6000px+ on the long edge — that's pure latency, not legibility. Every
   image gets resized to 2200px on its long edge before being sent (verified
   by eye that text is still fully readable at that size). Resized copies are
   cached in `scripts/.resize_cache/`, keyed by a hash of the *full source
   path* (not just the filename) so two different folders can both contain a
   file named e.g. `45.jpg` without colliding.

2. **Shells out to the `claude` CLI** in streaming JSON mode, with a specific
   set of flags chosen for speed and correctness:
   - `--json-schema` — forces the reply into the exact structure we need
     (see schemas below), so there's no fragile text-parsing.
   - `--allowedTools Read` — the model can only read the image, nothing else.
   - `--strict-mcp-config` / `--disable-slash-commands` — skips loading
     unrelated project integrations (Gmail, Drive, etc.) and skill
     definitions that add startup latency for zero benefit here.
   - `--effort low` — tested head-to-head against default effort: default
     effort spiraled into dozens of turns trying to zoom into the image via
     tools it didn't have permission to use, burned ~$1.20 in usage, and
     still hit a rate limit without finishing. Low effort finished cleanly in
     under two minutes with equivalent-quality output. For this kind of
     "read and transcribe" task, low effort is more *reliable*, not just
     faster.
   - `--exclude-dynamic-system-prompt-sections` — keeps the cached prompt
     prefix byte-identical across every call, so repeated calls increasingly
     hit Anthropic's prompt cache instead of re-paying for it.
   - `--no-session-persistence` — each image is a one-off; no need to save a
     resumable session to disk.

3. **Streams the response live.** Rather than blocking silently for minutes,
   it watches the raw JSON stream for each new `"row_number":` key as it's
   generated and calls `on_row(n)` — this is what prints
   `reading row 7 of 45.jpg ...` live in your terminal instead of leaving you
   staring at a frozen screen.

4. **Enforces one deadline over the whole call** (`CLI_TIMEOUT_SECONDS`,
   600s): launch, streaming, and exit. A CLI that stalls mid-stream is
   stopped instead of hanging the batch. stderr is read on its own thread so
   a noisy CLI can't fill the pipe and freeze. The CLI runs in its own
   process group; on a timeout, an error, or Ctrl+C, that whole group is
   terminated (then killed if it ignores that) before the script moves on.
   The streaming part lives in `run_cli_stream()`.

5. **Classifies the outcome** into one of three buckets: `ok` (got rows),
   `rate_limit` (subscription/session limit hit — this is detected by
   pattern-matching the CLI's own error text, and the human-readable reset
   message like *"You've hit your session limit · resets 8:40pm"* is
   extracted and shown to you directly), or `error` (anything else — a
   timeout, a malformed response, a crash).

### The prompt, and why it says what it says (`prompts.py`)

Both pipelines' prompts converged on the same hard-won rules after real
mistakes surfaced in your actual ledger:

- **"Most pages are ruled for 21 entries — a common count, not a required
  one."** Used as a one-time self-check: if the count isn't 21, re-scan once
  for a split or merged row, then report the true count. The model is told
  never to force a page to 21.
- **"Use the date/age column as ground truth for row count, not the
  crime/complexion text."** The original failure mode: a long crime
  description wrapping onto a second handwritten line got misread as a
  *second person*, splitting one real row into two. Fields like "When
  Received" or "Age" are always exactly one value per person and never wrap,
  so they're a much more reliable anchor for "how many rows are actually
  here" than the free-text columns.
- **"Ditto marks are legible, not a reason to flag."** Early runs flagged
  almost every row for review because standard ledger shorthand (`"`, `do.`,
  `Same`) was treated as uncertain. It isn't — it's a normal, confidently
  resolvable convention, so only *genuinely* illegible handwriting should
  produce a `review` status.
- **"Leave out repeat-offense counters like (2nd)/(3)."** Cosmetic cleanup
  requested partway through, so `Crime` reads as `"Burglary and Larceny"`
  rather than `"Burglary(2) and Larceny(1)"`.

### Writing to Excel (`excel_sheet.py`)

Both workbook scripts drive Excel live via `xlwings`/AppleScript rather than
writing a file on disk that you'd have to reopen. xlwings is only imported
when a workbook is actually opened, so the export and the tests never load
it. Apart from opening, every helper works on anything with xlwings' range
API (the tests pass an in-memory fake).

- `get_or_open_workbook()` looks for the workbook already open in Excel (by
  full path) and attaches to it; if it's not open, it opens it.
- Before anything else, `open_workbook_with_backup()` **saves the live
  workbook to disk, then backs it up**
  (`backups/…_backup_<timestamp>.xlsx`) — saving first matters because your
  workbook can have unsaved edits sitting only in Excel's memory, and a
  file-level copy would otherwise miss them.
- `find_append_row()` finds where to start appending: the first row below
  the last row with **anything in any named column**. (It used to look only
  at Crime, so a final record with a blank Crime looked like free space and
  could be overwritten.) Excel's own "used range" is not the answer on this
  sheet, because a Google Sheets export artifact inflates it with
  thousands of blank but formatted rows. It is only used as an upper limit
  on how far down to read.
- Right before each page is written, `occupied_cells()` re-reads the exact
  destination rows. If any named cell there is filled (for example someone
  typed below the data while the run was going), nothing is written and the
  run stops.
- Every cell write goes through `set_cell()`, a one-line wrapper — nothing
  fancier than direct `sheet.range((row, col)).value = ...` calls.

### Progress tracking / resume / dedup (`common.py`)

`common.py` also holds what every entry point shares: config loading,
`--limit`/`--only`/`--config` handling, finding images in natural number
order, the usage-limit stop message, and `configure_stdout()`. That last one
line-buffers output so progress shows live even when redirected to a log.
It runs from each `main()`, not at import.

`progress.json` (crime pipeline) and `progress_complexion.json` (complexion
pipeline) are the memory between runs. For every image, one entry records
its outcome:

```json
"53.jpg": {
  "status": "done",
  "rows_written": 21,
  "excel_rows": [484, 504],
  "flagged": 1,
  "timestamp": "2026-09-15T22:08:00"
}
```

- Anything already `"done"` is skipped on the next run — this is what makes
  `./run_batch.sh` safe to just re-run after adding new images, after a
  crash, or after a subscription-limit stop.
- Before reading an image, it is marked `"in_progress"` and the file is
  saved immediately. Cells are only written after a *complete* successful
  extraction, but a crash *during* the writes can still leave a page half
  written.
  - Crime pipeline: a leftover `"in_progress"` triggers the "left
    in-progress by a previous run — continue anyway? [y/N]" prompt. Check
    the rows at the end of the sheet first. A half-written page is *not*
    overwritten (the append position goes below it), so continuing without
    cleaning up would add that page a second time.
  - Complexion pipeline: resumes on its own and safely (see section 3).
- On a genuine subscription/rate-limit hit, the run stops immediately,
  prints the CLI's own reset-time message, and leaves everything already
  written intact.

---

## 2. `process_reports.py` — the offence/parish/residence pipeline

This is the simpler of the two: **append-only.** For each new image, in
filename order:

1. Send it to Claude with `EXTRACTION_SCHEMA` (`row_number`, `crime`,
   `parish`, `home_address`, `status`, `notes` per row).
2. Write each row into the next available spreadsheet rows, filling Crime,
   Parish, Home Address, File Name (the source image, for traceability), and
   Status Notes (`"OK"` or `"REVIEW: <reason>"`).
3. Record the exact Excel row range used (`excel_rows: [start, end]`) and row
   count (`rows_written`) in `progress.json`. The complexion pipeline uses
   this record to find the page again, and checks it against the File Name
   column before every write (see section 3).

`--limit N` processes only the first N not-yet-done images — useful for
validating a fix before letting it run over the whole folder.

---

## 3. `process_complexion.py` — the Race-column pipeline

This one is trickier because it has to fill in *existing* rows correctly,
not just append. It reuses `call_claude_vision` with a different schema
(`row_number`, `complexion`, `status`, `notes`) and a different prompt, but
the real design is in how it decides *where* to write:

### Pairing by filename number, not by order

Each ledger page's two photographed halves are numbered consecutively — the
complexion half is always **crime-page-number minus 1** (`74.jpg` pairs with
`75.jpg`). The script computes this directly from the filename
(`extract_number()` pulls the first integer out of the stem) and looks up
`75.jpg`'s recorded `excel_rows`/`rows_written` from `progress.json`.

This was a deliberate rewrite — the original version paired the *Nth
not-yet-paired complexion image* with the *Nth crime image*, in whatever
order they happened to be processed. That's fragile: any manual correction
to one page's row count/position (which came up more than once) drifted
every pairing after it. Number-based lookup has no such dependency — images
can be dropped in and processed in any order, out of sequence, with gaps,
and each one still finds its own correct partner independently.

### Never write without confirmed alignment

1. **The partner page must exist and be `"done"`** in the crime pipeline's
   `progress.json`. If not (e.g. `76.jpg`'s partner `77.jpg` hasn't been
   processed yet), that one image is skipped for this run, left untouched in
   `progress_complexion.json`, and retried next time. It does **not** block
   later images in the same run.
2. **The recorded rows must still be that page's rows** (`verify_alignment()`).
   Every row in `excel_rows` must have the partner's filename in File Name,
   the range must be exactly `rows_written` long, and the rows just above and
   below must *not* carry that filename. This runs *before* the image is sent
   to Claude, so a misaligned page costs no usage. It runs again right before
   writing, because the workbook is live and the extraction takes minutes.
   If it fails, the run stops and says where the File Name column shows that
   page now and what to change in `progress.json`. It never picks a new
   location on its own.
3. **The extracted row count must exactly equal the recorded row count.** A
   mismatch (e.g. 21 vs. 22) stops the run without writing that page.
4. **Existing Race values are compared, not just detected**
   (`plan_race_writes()`). For each row: blank → filled; same value already
   there (ignoring case and spacing) → kept; reading came back blank
   (illegible) → existing value kept; **different value** → conflict. Any
   conflict stops the run *before anything on that page is written* and
   lists each row, the sheet's value and the image's value. Fix or clear
   those cells and re-run, or mark the image done in
   `progress_complexion.json` if the sheet is already right.

### Resuming without re-reading images

Once a page has been read, its rows are saved in its
`progress_complexion.json` entry (`"extracted"`) *before* any cell is
written, and stay there until the page is `"done"`. A re-run after a crash,
a conflict, or a misalignment reuses that saved reading instead of calling
Claude again (as long as it's for the same partner page). A crash halfway
through a page is then finished safely: rows already written agree with the
saved reading and are kept, and the rest are filled in. The "left in
progress, continue anyway?" prompt is gone from this script for that reason.

To force a fresh reading, delete the image's entry. A row-count mismatch is
not cached, because re-reading is how that gets fixed.

Entry statuses: `done`; `in_progress` (reading saved, writes may be
partial); `conflict` (with a `conflicts` list); `misaligned` (with the
explanation); `error`; `pending` (stopped by a usage limit). Older entries
marked `skipped`, which used to be re-read on every run, are now resolved
once: compared, finished where they agree, and stopped on a conflict.

### Status Notes get appended to, never overwritten

If a complexion value is uncertain, its reviewer note is *appended* to
whatever's already in that row's Status Notes cell (which may already say
`"OK"` or carry a note from the crime pass) rather than replacing it — so a
row flagged for one reason doesn't silently lose an earlier flag for another
reason. An `"OK"` cell becomes `"REVIEW: complexion uncertain"`; an existing
`"REVIEW: ..."` cell becomes `"REVIEW: ...; complexion: ..."`.

---

## 4. What a "mismatch" or manual fix actually means

`process_complexion.py` no longer trusts `progress.json`'s row records
blindly: it checks them against the File Name column. When they don't
match — usually because rows were deleted, inserted, moved, or sorted
directly in Excel, or because an original extraction had a bug — it stops
and lets a human confirm, rather than write into what might be the wrong
rows. The stop message shows where that page's rows are now. Rows that have
no File Name (the hand-entered rows near the top of the sheet) can't be
confirmed, so they always stop.
This came up for real several times while building this out, and the fix
pattern was always the same:

1. Look at the actual scanned image directly to establish ground truth.
2. Compare against what's live in the sheet, row by row, at the boundary in
   question.
3. Either patch `progress.json`'s row numbers (if nothing physically moved,
   just the bookkeeping was wrong) or re-run the extraction fresh and
   overwrite the affected rows in place (if the original extraction itself
   was wrong) — while leaving every *other* tracked image's rows untouched
   whenever possible, to avoid a small fix cascading into a large one.

If you ever manually cut/paste or delete rows directly in Excel, the same
principle applies: tell me exactly what moved, and I'll reconcile
`progress.json`/`progress_complexion.json` to match reality rather than
assume anything.

Known gap: the check works page by page. If the rows *within* one page were
reordered among themselves (e.g. sorted by Name inside the page's block),
every row still says the right filename and it can't be detected.

---

## 5. File map

```
histpun reader/
  run_batch.sh                    one command: crime/parish/address pipeline
  run_complexion_batch.sh         one command: complexion pipeline
  run_complexion_export.sh        complexion readings to complexion_export.xlsx (no workbook)
  progress.json                   crime pipeline's memory (per-image status + row ranges)
  progress_complexion.json        complexion pipeline's memory
  backups/                        timestamped workbook backups, made before every run
  scripts/
    process_reports.py            entry point: crime pipeline (append pages)
    process_complexion.py         entry point: complexion pipeline (+ its alignment/recovery rules)
    export_complexion.py          entry point: standalone Image/Complexion export
    claude_cli.py                 image downscaling + the streaming `claude -p` call
    prompts.py                    the two schemas and prompts
    excel_sheet.py                worksheet helpers; xlwings loaded only when opening a workbook
    common.py                     config, stdout, image selection, progress files, backups
    config.json                   paths for the crime pipeline
    config_complexion.json        paths for the complexion pipeline
    .resize_cache/                downscaled images, cached across runs
  tests/                          regression tests (fake worksheet + fake CLI; no Excel, no Claude)
  .venv/                          isolated Python environment (xlwings, openpyxl, Pillow)
```

## 6. Tests

```
.venv/bin/python -m unittest discover -s tests
```

These tests use an in-memory stand-in for the worksheet (`tests/fake_sheet.py`),
temporary folders, and tiny Python processes standing in for the `claude`
CLI. They never open Excel, read the real workbook, images, or progress
files, or use any Claude usage. They take about 10 seconds.

| File | Covers |
|---|---|
| `test_append_position.py` | where the crime pipeline appends; occupied-destination check |
| `test_complexion.py` | alignment checks, conflicts, cached readings, crash recovery |
| `test_cli_stream.py` | the whole-call deadline, stderr draining, child cleanup, Ctrl+C |
| `test_orchestration.py` | the crime and export loops: error handling, write → save → done order |
| `test_boundaries.py` | no xlwings or stdout changes at import, no script importing a script, `--help` on all three launchers |
