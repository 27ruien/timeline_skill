from __future__ import annotations

import unittest

from scripts.build_timeline import build_workbook


class TimelineExportOrderTests(unittest.TestCase):
    def test_export_keeps_first_seen_stage_order_and_font_sizes(self) -> None:
        workbook = build_workbook(
            {
                "project_name": "Stage order",
                "include_status": False,
                "language": "en",
                "tasks": [
                    {"model": "Development", "name": "Build A", "owners": ["Kivisense"], "start": "2026-06-08", "end": "2026-06-08"},
                    {"model": "Requirement", "name": "Scope", "owners": ["Brands"], "start": "2026-06-01", "end": "2026-06-01"},
                    {"model": "Proposal", "name": "Concept", "owners": ["Kivisense"], "start": "2026-06-02", "end": "2026-06-02"},
                    {"model": "Development", "name": "Build B", "owners": ["Kivisense"], "start": "2026-06-09", "end": "2026-06-09"},
                ],
            }
        )
        worksheet = workbook["Timeline"]

        stage_starts = [worksheet.cell(row, 1).value for row in range(5, 9) if worksheet.cell(row, 1).value]
        self.assertEqual(stage_starts, ["Development", "Requirement", "Proposal"])
        self.assertEqual(worksheet["B1"].font.sz, 16)
        self.assertEqual(worksheet["B1"].font.name, "Gotham")
        self.assertEqual(worksheet["A3"].font.sz, 11)
        self.assertEqual(worksheet["E4"].font.sz, 11)
        self.assertEqual(worksheet["A5"].font.sz, 11)
        self.assertIsNone(worksheet["A5"].font.scheme)
        self.assertEqual(worksheet["B5"].font.sz, 11)
        for coordinate in ("A3", "E4", "A5", "B5"):
            self.assertEqual(worksheet[coordinate].font.name, "Gotham")


if __name__ == "__main__":
    unittest.main()
