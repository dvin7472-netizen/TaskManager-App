import unittest
from datetime import date, datetime

from sprint_report import (
    SprintTask,
    _extract_json,
    _fallback_brief,
    _fallback_team_brief,
    burndown_series,
    parse_date,
    parse_time_spent,
)


class TimeLogParsingTests(unittest.TestCase):
    def test_parses_text_hours_and_minutes(self):
        result, errors = parse_time_spent(
            "2026-06-01: 4h, 2026-06-02: 30 minutes; 2026-06-03=1.5"
        )
        self.assertEqual(errors, [])
        self.assertEqual(result[date(2026, 6, 1)], 4)
        self.assertEqual(result[date(2026, 6, 2)], 0.5)
        self.assertEqual(result[date(2026, 6, 3)], 1.5)

    def test_parses_json_mapping_and_reports_bad_entry(self):
        result, errors = parse_time_spent(
            '{"2026-06-01": "2h", "not-a-date": "bad"}'
        )
        self.assertEqual(result, {date(2026, 6, 1): 2})
        self.assertEqual(len(errors), 1)

    def test_keeps_valid_time_log_entries_when_other_fragments_are_invalid(self):
        result, errors = parse_time_spent("2026-06-01: 2h, invalid fragment")
        self.assertEqual(result, {date(2026, 6, 1): 2})
        self.assertEqual(len(errors), 1)


class DateAndBurndownTests(unittest.TestCase):
    def test_date_parser_handles_excel_style_values(self):
        self.assertEqual(parse_date("06/01/2026"), date(2026, 6, 1))
        self.assertEqual(parse_date(datetime(2026, 6, 1, 12)), date(2026, 6, 1))
        self.assertIsNone(parse_date("not-a-date"))

    def test_burndown_uses_task_count_and_completion_date(self):
        tasks = [
            SprintTask("A", "done", "S1", "Done", None, None, date(2026, 6, 2)),
            SprintTask("A", "active", "S1", "In-progress", None, None, None),
        ]
        days, ideal, actual, effort = burndown_series(
            tasks, date(2026, 6, 1), date(2026, 6, 3)
        )
        self.assertEqual(days, [date(2026, 6, 1), date(2026, 6, 2), date(2026, 6, 3)])
        self.assertEqual(ideal, [2, 1, 0])
        self.assertEqual(actual, [2, 1, 1])
        self.assertEqual(effort, [0, 0, 0])


class BriefTests(unittest.TestCase):
    def _tasks(self):
        return [
            SprintTask("Ava", "Build API", "S1", "Done", None, None, date(2026, 6, 2)),
            SprintTask("Ava", "QA follow-up", "S1", "In-progress", None, None, None),
            SprintTask("Ben", "Integration tests", "S1", "Blocked", None, None, None),
        ]

    def test_extract_json_handles_fenced_block(self):
        payload = _extract_json('```json\n{"summary": "x", "actions": ["y"]}\n```')
        self.assertEqual(payload["summary"], "x")
        self.assertEqual(payload["actions"], ["y"])

    def test_extract_json_returns_none_on_non_json(self):
        self.assertIsNone(_extract_json("no json here"))
        self.assertIsNone(_extract_json(""))

    def test_fallback_brief_reports_risk_and_actions(self):
        brief = _fallback_brief(self._tasks())
        self.assertTrue(brief.summary)
        self.assertIn("blocked", brief.risk.lower())
        self.assertTrue(brief.actions)
        self.assertLessEqual(len(brief.actions), 3)

    def test_fallback_team_brief_names_blocked_developer(self):
        team = _fallback_team_brief(self._tasks())
        self.assertTrue(team.narrative)
        self.assertTrue(any("Ben" in risk for risk in team.risks))


if __name__ == "__main__":
    unittest.main()
