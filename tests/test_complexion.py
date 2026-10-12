"""Issues 2 and 3: alignment checks and resumable complexion writes."""
import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_sheet import FakeBook, FakeSheet  # noqa: E402
from process_complexion import (  # noqa: E402
    AlignmentError,
    plan_race_writes,
    process_images,
    review_note_value,
    verify_alignment,
)
from common import load_progress  # noqa: E402

HEADERS = ["Name", "Race", "Crime", "Status Notes", "File Name"]


def crime_page(name, n):
    return [{"Crime": f"{name} crime {i}", "Status Notes": "OK", "File Name": name} for i in range(n)]


def reading(*values, review=()):
    return [{"row_number": i + 1, "complexion": v,
             "status": "review" if i in review else "ok",
             "notes": "faint" if i in review else ""} for i, v in enumerate(values)]


class FakeExtractor:
    def __init__(self, rows=None, kind="ok", side_effect=None):
        self.rows, self.kind, self.side_effect = rows, kind, side_effect
        self.calls = []

    def __call__(self, img_path, on_row):
        self.calls.append(img_path.name)
        if self.side_effect:
            self.side_effect()
        if self.kind != "ok":
            return {"kind": self.kind, "message": "boom"}
        return {"kind": "ok", "rows": self.rows}


class ComplexionTestBase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.progress_path = Path(self.tmp.name) / "progress_complexion.json"
        # 75.jpg -> rows 2-4, 77.jpg -> rows 5-7
        self.sheet = FakeSheet(HEADERS, crime_page("75.jpg", 3) + crime_page("77.jpg", 3))
        self.book = FakeBook()
        self.crime = {
            75: {"filename": "75.jpg", "status": "done", "excel_rows": [2, 4], "rows_written": 3},
            77: {"filename": "77.jpg", "status": "done", "excel_rows": [5, 7], "rows_written": 3},
        }
        self.progress = {}

    def tearDown(self):
        self.tmp.cleanup()

    def run_images(self, extractor, names=("74.jpg",), sheet=None):
        out = io.StringIO()
        with redirect_stdout(out):
            summary = process_images([Path(n) for n in names], sheet or self.sheet, self.book,
                                     self.sheet.header_map, self.crime, self.progress,
                                     self.progress_path, extractor)
        self.output = out.getvalue()
        return summary

    def race(self, rows):
        return [self.sheet.get(r, "Race") for r in rows]

    def saved(self, name="74.jpg"):
        return load_progress(self.progress_path)[name]


class AlignmentTest(ComplexionTestBase):
    def test_aligned_page_is_written(self):
        ext = FakeExtractor(reading("Black", "White", "Mulatto"))
        summary = self.run_images(ext)
        self.assertFalse(summary["stopped"])
        self.assertEqual(self.race([2, 3, 4]), ["Black", "White", "Mulatto"])
        self.assertEqual(self.race([5, 6, 7]), [None] * 3)
        entry = self.saved()
        self.assertEqual(entry["status"], "done")
        self.assertEqual(entry["excel_rows"], [2, 4])
        self.assertEqual(entry["rows_written"], 3)
        self.assertNotIn("extracted", entry)

    def test_inserted_rows_stop_before_any_extraction(self):
        self.sheet.insert_rows(2, 1)  # someone inserted a row above 75.jpg
        ext = FakeExtractor(reading("Black", "White", "Mulatto"))
        summary = self.run_images(ext)
        self.assertTrue(summary["stopped"])
        self.assertEqual(ext.calls, [])  # no usage spent
        self.assertEqual(self.sheet.writes, [])
        entry = self.saved()
        self.assertEqual(entry["status"], "misaligned")
        self.assertIn("rows 3-5", entry["message"])
        self.assertIn("[3, 5]", entry["message"])

    def test_deleted_row_is_detected(self):
        # 75.jpg lost a row: its block is now 2-3 and 77.jpg starts at 4
        self.sheet = FakeSheet(HEADERS, crime_page("75.jpg", 2) + crime_page("77.jpg", 3))
        with self.assertRaises(AlignmentError) as cm:
            verify_alignment(self.sheet, self.sheet.header_map, self.crime[75])
        self.assertIn("77.jpg on rows 4", str(cm.exception))
        self.assertIn("2 rows, not the 3 recorded", str(cm.exception))

    def test_page_extending_past_recorded_range_is_detected(self):
        self.sheet.set_record(5, {"File Name": "75.jpg"})
        with self.assertRaises(AlignmentError) as cm:
            verify_alignment(self.sheet, self.sheet.header_map, self.crime[75])
        self.assertIn("just below the range", str(cm.exception))

    def test_sorted_sheet_is_detected(self):
        a, b = self.sheet.get(3, "File Name"), self.sheet.get(6, "File Name")
        self.sheet.set_record(3, {"File Name": b})
        self.sheet.set_record(6, {"File Name": a})
        with self.assertRaises(AlignmentError) as cm:
            verify_alignment(self.sheet, self.sheet.header_map, self.crime[75])
        self.assertIn("3 separate places (2, 4, 6)", str(cm.exception))

    def test_recorded_span_disagreeing_with_row_count_is_detected(self):
        target = dict(self.crime[75], rows_written=4)
        with self.assertRaises(AlignmentError) as cm:
            verify_alignment(self.sheet, self.sheet.header_map, target)
        self.assertIn("spans 3 rows but 4 were recorded", str(cm.exception))

    def test_rows_without_file_names_cannot_be_confirmed(self):
        self.sheet = FakeSheet(HEADERS, [{"Crime": "x"}] * 3)
        with self.assertRaises(AlignmentError) as cm:
            verify_alignment(self.sheet, self.sheet.header_map, self.crime[75])
        self.assertIn("(blank) on rows 2-4", str(cm.exception))
        self.assertIn("No row's File Name is 75.jpg", str(cm.exception))

    def test_sheet_changed_during_extraction_keeps_reading_for_later(self):
        ext = FakeExtractor(reading("Black", "White", "Mulatto"),
                            side_effect=lambda: self.sheet.insert_rows(2, 1))
        summary = self.run_images(ext)
        self.assertTrue(summary["stopped"])
        self.assertEqual([w for w in self.sheet.writes], [])
        entry = self.saved()
        self.assertEqual(entry["status"], "misaligned")
        self.assertEqual([r["complexion"] for r in entry["extracted"]], ["Black", "White", "Mulatto"])

        # The user follows the message and fixes progress.json; the saved
        # reading is reused and lands on the moved rows.
        self.crime[75]["excel_rows"] = [3, 5]
        ext2 = FakeExtractor(reading("WRONG", "WRONG", "WRONG"))
        summary = self.run_images(ext2)
        self.assertFalse(summary["stopped"])
        self.assertEqual(ext2.calls, [])
        self.assertEqual(self.race([3, 4, 5]), ["Black", "White", "Mulatto"])
        self.assertEqual(self.saved()["status"], "done")

    def test_row_count_mismatch_writes_nothing(self):
        ext = FakeExtractor(reading("Black", "White"))
        summary = self.run_images(ext)
        self.assertTrue(summary["stopped"])
        self.assertEqual(self.sheet.writes, [])
        entry = self.saved()
        self.assertEqual(entry["status"], "error")
        self.assertNotIn("extracted", entry)  # a miscount is fixed by re-reading, so don't cache it


class RecoveryTest(ComplexionTestBase):
    def test_partial_page_is_completed_without_touching_agreeing_cells(self):
        self.sheet.set_record(2, {"Race": "Black"})
        ext = FakeExtractor(reading("black ", "White", "Mulatto"))
        self.run_images(ext)
        self.assertEqual(self.race([2, 3, 4]), ["Black", "White", "Mulatto"])
        written_rows = [rc[0] for rc, _ in self.sheet.writes]
        self.assertNotIn(2, written_rows)
        entry = self.saved()
        self.assertEqual(entry["status"], "done")
        self.assertEqual(entry["already_present"], 1)

    def test_conflict_stops_writes_nothing_and_does_not_reextract(self):
        self.sheet.set_record(3, {"Race": "Griffe"})
        ext = FakeExtractor(reading("Black", "White", "Mulatto"))
        summary = self.run_images(ext)
        self.assertTrue(summary["stopped"])
        self.assertEqual(self.sheet.writes, [])
        entry = self.saved()
        self.assertEqual(entry["status"], "conflict")
        self.assertEqual(entry["conflicts"], [{"row": 3, "sheet": "Griffe", "image": "White"}])
        self.assertIn("row 3: sheet has 'Griffe', image reads 'White'", self.output)

        # Re-running unchanged: same conflict, no new extraction.
        ext2 = FakeExtractor(reading("x", "x", "x"))
        self.run_images(ext2)
        self.assertEqual(ext2.calls, [])
        self.assertEqual(self.saved()["status"], "conflict")

        # The researcher clears the cell; the saved reading completes the page.
        self.sheet.set_record(3, {"Race": None})
        self.run_images(ext2)
        self.assertEqual(ext2.calls, [])
        self.assertEqual(self.race([2, 3, 4]), ["Black", "White", "Mulatto"])
        self.assertEqual(self.saved()["status"], "done")

    def test_illegible_reading_keeps_existing_value(self):
        self.sheet.set_record(4, {"Race": "Mulatto"})
        ext = FakeExtractor(reading("Black", "White", "", review={2}))
        self.run_images(ext)
        self.assertEqual(self.race([2, 3, 4]), ["Black", "White", "Mulatto"])
        self.assertEqual(self.sheet.get(4, "Status Notes"), "OK")  # existing value stands, no flag
        self.assertEqual(self.saved()["status"], "done")

    def test_legacy_skipped_entry_is_resolved_once(self):
        for r, v in zip((2, 3, 4), ("Black", "White", "Mulatto")):
            self.sheet.set_record(r, {"Race": v})
        self.progress["74.jpg"] = {"status": "skipped", "reason": "target rows already had Race values: [2]",
                                   "paired_with": "75.jpg"}
        ext = FakeExtractor(reading("Black", "White", "Mulatto"))
        self.run_images(ext)
        self.assertEqual(ext.calls, ["74.jpg"])
        self.assertEqual(self.sheet.writes, [])
        entry = self.saved()
        self.assertEqual(entry["status"], "done")  # no longer re-extracted every run
        self.assertEqual(entry["already_present"], 3)

    def test_crash_mid_write_resumes_from_saved_reading(self):
        self.sheet.writes_before_crash = 2  # row 2's Race + note, then die
        ext = FakeExtractor(reading("Black", "White", "Mulatto", review={0}))
        with self.assertRaises(KeyboardInterrupt):
            self.run_images(ext)
        # Race and note for row 2 made it; row 3+ didn't.
        self.assertEqual(self.race([2, 3, 4]), ["Black", None, None])
        entry = self.saved()
        self.assertEqual(entry["status"], "in_progress")
        self.assertEqual(len(entry["extracted"]), 3)

        self.sheet.writes_before_crash = None
        self.progress = load_progress(self.progress_path)
        ext2 = FakeExtractor(reading("x", "x", "x"))
        self.run_images(ext2)
        self.assertEqual(ext2.calls, [])
        self.assertEqual(self.race([2, 3, 4]), ["Black", "White", "Mulatto"])
        self.assertEqual(self.sheet.get(2, "Status Notes"), "REVIEW: complexion: faint")  # not doubled
        self.assertEqual(self.saved()["status"], "done")

    def test_extraction_error_is_not_cached(self):
        ext = FakeExtractor(kind="error")
        summary = self.run_images(ext)
        self.assertTrue(summary["stopped"])
        self.assertEqual(self.saved()["status"], "error")
        self.assertNotIn("extracted", self.saved())

    def test_cached_reading_for_a_different_partner_is_ignored(self):
        self.progress["74.jpg"] = {"status": "conflict", "paired_with": "other.jpg",
                                   "extracted": reading("a", "b", "c")}
        ext = FakeExtractor(reading("Black", "White", "Mulatto"))
        self.run_images(ext)
        self.assertEqual(ext.calls, ["74.jpg"])
        self.assertEqual(self.race([2, 3, 4]), ["Black", "White", "Mulatto"])

    def test_progress_file_stays_plain_json(self):
        self.run_images(FakeExtractor(reading("Black", "White", "Mulatto")))
        json.loads(self.progress_path.read_text())


class HelpersTest(unittest.TestCase):
    def test_plan(self):
        plan = plan_race_writes(["Black", None, "  ", "Griffe", "Dark"],
                                reading("BLACK", "White", "Yellow", "Brown", ""))
        self.assertEqual(plan["agree"], [0])
        self.assertEqual(plan["write"], [1, 2])
        self.assertEqual(plan["conflicts"], [(3, "Griffe", "Brown")])
        self.assertEqual(plan["kept"], [4])

    def test_review_note_is_appended_once(self):
        r = {"notes": "faint"}
        self.assertEqual(review_note_value("OK", r), "REVIEW: complexion: faint")
        self.assertEqual(review_note_value(None, r), "REVIEW: complexion: faint")
        self.assertEqual(review_note_value("REVIEW: parish", r), "REVIEW: parish; complexion: faint")
        self.assertEqual(review_note_value("checked by hand", r), "checked by hand; REVIEW: complexion: faint")
        self.assertIsNone(review_note_value("REVIEW: complexion: faint", r))
        self.assertEqual(review_note_value("OK", {}), "REVIEW: complexion uncertain")


if __name__ == "__main__":
    unittest.main()
