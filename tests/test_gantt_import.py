from __future__ import annotations

import json
import os
import subprocess
import sys
import unittest
from pathlib import Path

from openpyxl import load_workbook

from local_app import month_from_label, normalize_cell_fill, parse_imported_workbook, run_self_test


FIXTURE = Path(__file__).parent / "fixtures" / "CHAGEE_AR_Campaign_Timeline.xlsx"


class VisualGanttImportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.file_bytes = FIXTURE.read_bytes()
        cls.result = parse_imported_workbook(cls.file_bytes, filename=FIXTURE.name, import_year=2026)
        cls.tasks_by_name: dict[str, list[dict]] = {}
        for task in cls.result["tasks"]:
            cls.tasks_by_name.setdefault(task["name"], []).append(task)

    def task(self, name: str, group: str | None = None) -> dict:
        matches = self.tasks_by_name[name]
        if group is not None:
            matches = [task for task in matches if task["group"] == group]
        self.assertEqual(len(matches), 1, f"expected one {name!r} task in {group!r}")
        return matches[0]

    def test_builds_the_exact_non_contiguous_timeline_dates(self) -> None:
        dates = self.result["import_preview"]["column_date_map"]
        self.assertEqual(dates["F"], "2026-07-13")
        self.assertEqual(dates["T"], "2026-07-31")
        self.assertEqual(dates["U"], "2026-08-03")
        self.assertEqual(dates["AO"], "2026-08-31")
        self.assertEqual(dates["AP"], "2026-09-01")
        self.assertEqual(dates["BJ"], "2026-09-30")
        self.assertEqual(dates["BK"], "2026-10-08")
        self.assertEqual(dates["BO"], "2026-10-13")
        self.assertNotIn("2026-08-01", dates.values())
        self.assertNotIn("2026-10-11", dates.values())
        self.assertEqual(self.result["import_preview"]["date_range"], "2026-07-13 至 2026-10-13")

    def test_infers_july_and_keeps_the_import_year_as_date_only_strings(self) -> None:
        preview = self.result["import_preview"]
        self.assertFalse(self.result["yearInferred"])
        self.assertEqual(self.result["inferredYearSource"], "user-input")
        self.assertIn("July", preview["inferred_months"])
        self.assertEqual(preview["months"], ["July", "August", "September", "October"])
        self.assertTrue(all(len(task["start"]) == 10 and len(task["end"]) == 10 for task in self.result["tasks"]))
        self.assertTrue(all(not task["start"].startswith("1900-") for task in self.result["tasks"]))

    def test_month_labels_accept_chinese_and_english_names(self) -> None:
        self.assertEqual(month_from_label("7月"), 7)
        self.assertEqual(month_from_label("八月"), 8)
        self.assertEqual(month_from_label("SEP"), 9)

    def test_date_only_import_is_independent_of_timezone_environment(self) -> None:
        script = (
            "import json\n"
            "from pathlib import Path\n"
            "from local_app import parse_imported_workbook\n"
            f"result = parse_imported_workbook(Path({str(FIXTURE)!r}).read_bytes(), filename={FIXTURE.name!r}, import_year=2026)\n"
            "print(json.dumps({'map': result['import_preview']['column_date_map'], 'tasks': [(task['start'], task['end']) for task in result['tasks']]}))\n"
        )
        outputs = []
        for timezone in ("Asia/Shanghai", "America/Los_Angeles", "UTC"):
            environment = os.environ | {"TZ": timezone}
            raw = subprocess.check_output(
                [sys.executable, "-c", script],
                cwd=Path(__file__).parents[1],
                env=environment,
                text=True,
            )
            outputs.append(json.loads(raw))
        self.assertEqual(outputs[0], outputs[1])
        self.assertEqual(outputs[1], outputs[2])

    def test_recognizes_empty_value_colored_task_bars_and_merged_groups(self) -> None:
        workbook = load_workbook(FIXTURE, data_only=True)
        worksheet = workbook["Timeline"]
        self.assertIsNone(worksheet["F5"].value)
        self.assertEqual(normalize_cell_fill(worksheet["F5"]), "A9D08E")
        self.assertEqual(self.task("Project Scope & Timeline Confirmation")["group"], "Requirement")
        self.assertEqual(self.task("SIT")["group"], "Development")

    def test_expected_task_ranges_and_working_day_counts(self) -> None:
        expected = {
            ("Project Scope & Timeline Update", "Requirement"): ("2026-07-13", "2026-07-13", 1),
            ("Project Scope & Timeline Confirmation", "Requirement"): ("2026-07-13", "2026-07-14", 2),
            ("Confirmation #Mural", "Proposal"): ("2026-07-30", "2026-08-05", 5),
            ("Development #Vally Fair H5", "Development"): ("2026-07-27", "2026-08-28", 25),
            ("SIT", "Development"): ("2026-08-31", "2026-09-11", 10),
            ("UAT #Round 1 & Feedback", "UAT"): ("2026-09-14", "2026-09-16", 3),
            ("Launch", "Launch"): ("2026-09-30", "2026-09-30", 1),
        }
        for (name, group), (start, end, working_days) in expected.items():
            task = self.task(name, group)
            self.assertEqual((task["startDate"], task["endDate"], task["durationWorkingDays"]), (start, end, working_days))
            self.assertEqual(sum(segment["workingDays"] for segment in task["segments"]), working_days)

    def test_all_31_tasks_have_valid_dates(self) -> None:
        self.assertEqual(len(self.result["tasks"]), 31)
        self.assertEqual(self.result["import_preview"]["valid_date_task_count"], 31)
        self.assertTrue(all(task["start"] <= task["end"] for task in self.result["tasks"]))

    def test_standard_exported_workbook_still_imports(self) -> None:
        output_path = run_self_test()
        imported = parse_imported_workbook(output_path.read_bytes(), filename=output_path.name, import_year=2026)
        self.assertEqual(len(imported["tasks"]), 3)
        first = next(task for task in imported["tasks"] if task["name"] == "Project requirement")
        self.assertEqual(first["group"], "需求")
        self.assertEqual((first["start"], first["end"]), ("2026-06-01", "2026-06-05"))


if __name__ == "__main__":
    unittest.main()
