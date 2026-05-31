"""Regression tests for pipeline Excel atomic-write paths."""
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parent


class TestPipelineAtomicWrites(unittest.TestCase):
    def test_backtest_replace_target_is_defined(self):
        source = (ROOT / "batch_backtest.py").read_text(encoding="utf-8")
        self.assertNotIn("os.replace(fn_tmp, fn)", source)
        self.assertIn("os.replace(fn_tmp, fn_final)", source)

    def test_weekly_temp_file_keeps_xlsx_extension(self):
        source = (ROOT / "batch_weekly.py").read_text(encoding="utf-8")
        self.assertNotIn("output_fn + '.tmp'", source)
        self.assertIn(".tmp.xlsx", source)


if __name__ == "__main__":
    unittest.main()
