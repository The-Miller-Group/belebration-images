import sys
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).parent / "src"))
import celebrations


class TestScheduling(unittest.TestCase):
    def test_central_timezone(self):
        for date, zone in [
            (datetime(2026, 7, 1, 9, tzinfo=ZoneInfo("America/Chicago")), "CDT"),
            (datetime(2026, 11, 4, 9, tzinfo=ZoneInfo("America/Chicago")), "CST"),
        ]:
            self.assertEqual(date.tzname(), zone)

    def test_teams_window(self):
        for hour in range(24):
            allowed = 9 <= hour < 11
            self.assertEqual(allowed, hour in (9, 10))

    def test_teams_blocked_outside_window(self):
        for hour in (0, 8, 11, 14, 23):
            with self.subTest(hour=hour):
                with patch.object(celebrations, "datetime") as clock:
                    clock.now.return_value = datetime(
                        2026, 11, 4, hour,
                        tzinfo=ZoneInfo("America/Chicago")
                    )
                    with self.assertRaises(RuntimeError):
                        celebrations.post_teams("Test", "Test", "test")
    def test_teams_allowed_inside_window(self):
        for hour in (9, 10):
            with self.subTest(hour=hour):
                with patch.object(celebrations, "datetime") as clock:
                    clock.now.return_value = datetime(
                        2026, 11, 4, hour,
                        tzinfo=ZoneInfo("America/Chicago")
                    )
                    with patch.dict(celebrations.os.environ, {"TEAMS_WEBHOOK_URL": "https://example.invalid/test"}):
                        with patch.object(celebrations, "request_json") as send:
                            celebrations.post_teams("Test", "Test", "test")
                            send.assert_called_once()
    def test_duplicate_keys(self):
        base = "2026-11-04|anniversary|employee@example.com"
        self.assertNotEqual(base + "|teams", base + "|email")


if __name__ == "__main__":
    unittest.main()
