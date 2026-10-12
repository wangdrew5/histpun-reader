"""Issue 1: the crime pipeline's append position and occupied-destination guard."""
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fake_sheet import FakeSheet  # noqa: E402
from excel_sheet import find_append_row, occupied_cells  # noqa: E402
from process_reports import write_page  # noqa: E402

HEADERS = ["Name", "Race", "Crime", "Parish", "Home Address", "Status Notes", "File Name"]


def page(name, n, **extra):
    return [{"Crime": f"crime {i}", "Parish": "P", "Home Address": "A",
             "Status Notes": "OK", "File Name": name, **extra} for i in range(n)]


class FindAppendRowTest(unittest.TestCase):
    def test_empty_sheet_starts_below_header(self):
        self.assertEqual(find_append_row(FakeSheet(HEADERS), FakeSheet(HEADERS).header_map), 2)

    def test_final_record_with_blank_crime_is_not_treated_as_free(self):
        rows = page("1.jpg", 3)
        rows[-1]["Crime"] = ""  # illegible crime on the last record of the page
        sheet = FakeSheet(HEADERS, rows)
        # The old Crime-only scan returned 4 here, overwriting row 4.
        self.assertEqual(find_append_row(sheet, sheet.header_map), 5)

    def test_manual_record_with_only_a_name_counts(self):
        sheet = FakeSheet(HEADERS, page("1.jpg", 2) + [{"Name": "Hand-entered"}])
        self.assertEqual(find_append_row(sheet, sheet.header_map), 5)

    def test_formatting_below_data_does_not_move_the_append_row(self):
        sheet = FakeSheet(HEADERS, page("1.jpg", 2), extra_used_rows=5000)
        self.assertEqual(find_append_row(sheet, sheet.header_map), 4)

    def test_gap_inside_data_does_not_end_the_scan(self):
        sheet = FakeSheet(HEADERS, page("1.jpg", 2) + [{}] + page("3.jpg", 1))
        self.assertEqual(find_append_row(sheet, sheet.header_map), 6)


class OccupiedCellsTest(unittest.TestCase):
    def test_blank_destination(self):
        sheet = FakeSheet(HEADERS, page("1.jpg", 2))
        self.assertEqual(occupied_cells(sheet, sheet.header_map, 4, 21), [])

    def test_reports_any_filled_named_cell_in_destination(self):
        sheet = FakeSheet(HEADERS, page("1.jpg", 2))
        sheet.set_record(10, {"Name": "Typed in mid-run"})
        self.assertEqual(occupied_cells(sheet, sheet.header_map, 4, 21), [(10, "Name")])

    def test_zero_rows(self):
        sheet = FakeSheet(HEADERS)
        self.assertEqual(occupied_cells(sheet, sheet.header_map, 2, 0), [])


class WritePageTest(unittest.TestCase):
    def test_writes_consecutive_rows_and_counts_flags(self):
        sheet = FakeSheet(HEADERS, page("1.jpg", 1))
        rows = [
            {"crime": "Larceny", "parish": "St. Ann", "home_address": "X", "status": "ok", "notes": ""},
            {"crime": "", "parish": "", "home_address": "", "status": "review", "notes": "smudged"},
            {"crime": "Assault", "parish": "", "home_address": "", "status": "bogus", "notes": ""},
        ]
        flagged = write_page(sheet, sheet.header_map, 3, rows, "2.jpg")
        self.assertEqual(flagged, 2)
        self.assertEqual([sheet.get(r, "File Name") for r in (3, 4, 5)], ["2.jpg"] * 3)
        self.assertEqual(sheet.get(3, "Status Notes"), "OK")
        self.assertEqual(sheet.get(4, "Status Notes"), "REVIEW: smudged")
        self.assertEqual(sheet.get(5, "Status Notes"), "REVIEW")
        self.assertEqual(sheet.get(2, "Crime"), "crime 0")  # previous record untouched


if __name__ == "__main__":
    unittest.main()
