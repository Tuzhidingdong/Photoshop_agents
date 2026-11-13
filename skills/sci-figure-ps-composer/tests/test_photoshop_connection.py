from __future__ import annotations

import sys
import unittest
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPT_DIR))

from photoshop_connection import parse_tasklist_csv, probe_win32com


class PhotoshopConnectionHelpersTests(unittest.TestCase):
    def test_parse_tasklist_csv_filters_processes_and_preserves_memory_field(self):
        output = (
            '"Photoshop.exe","10832","Console","1","2,345,678 K"\n'
            '"notepad.exe","900","Console","1","12,000 K"\n'
            '"Photoshop.exe","10833","Console","1","3,000 K"\n'
        )
        processes = parse_tasklist_csv(output)
        self.assertEqual([item["pid"] for item in processes], [10832, 10833])
        self.assertEqual(processes[0]["memory_usage"], "2,345,678 K")

    @unittest.skipUnless(sys.platform != "win32", "This test expects the non-Windows guard.")
    def test_non_windows_probe_returns_unsupported_without_importing_com(self):
        result = probe_win32com()
        self.assertFalse(result["ok"])
        self.assertEqual(result["error_type"], "UnsupportedPlatform")


if __name__ == "__main__":
    unittest.main()
