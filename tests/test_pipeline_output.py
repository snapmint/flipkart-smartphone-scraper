"""Regression tests for Jenkins workspace cleanup and workbook selection."""

import os
from pathlib import Path
import re
import subprocess
import tempfile
import unittest


JENKINSFILE = Path(__file__).resolve().parents[1] / "Jenkinsfile"


def stage_script(stage):
    pipeline = JENKINSFILE.read_text()
    match = re.search(
        r"stage\('" + re.escape(stage) + r"'\).*?sh '''(.*?)'''",
        pipeline,
        re.DOTALL,
    )
    if match is None:
        raise AssertionError(f"Shell script not found for stage: {stage}")
    # Groovy triple-single-quoted strings unescape doubled backslashes.
    return match.group(1).replace("\\\\", "\\")


class PipelineOutputTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.workspace = Path(self.temp_dir.name)
        self.env = dict(
            os.environ,
            VENV_DIR=".venv",
            FLIPKART_DATASET="flipkart_mobile_scraper",
            AMAZON_DATASET="amazon_mobile_scraper",
        )
        python = self.workspace / ".venv/bin/python"
        python.parent.mkdir(parents=True)
        # Stub only Excel conversion; execute the real pipeline selection shell.
        python.write_text('#!/bin/sh\nset -eu\ncp "$3" "$4"\n')
        python.chmod(0o755)

    def run_stage(self, stage):
        return subprocess.run(
            ["/bin/sh", "-c", stage_script(stage)],
            cwd=self.workspace,
            env=self.env,
            capture_output=True,
            text=True,
        )

    def workbook(self, dataset, timestamp, content):
        path = self.workspace / f"{dataset}_mobile_{timestamp}.xlsx"
        path.write_text(content)
        return path

    def test_cleanup_removes_only_previous_scraper_workbooks(self):
        old_fk = self.workbook("flipkart", "20260928_080000", "old flipkart")
        old_az = self.workbook("amazon", "20260928_080000", "old amazon")
        unrelated = self.workspace / "report.xlsx"
        unrelated.write_text("keep")
        archived = self.workspace / "archive/flipkart_mobile_old.xlsx"
        archived.parent.mkdir()
        archived.write_text("keep")

        result = self.run_stage("Clean Previous Scraper Output")

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertFalse(old_fk.exists())
        self.assertFalse(old_az.exists())
        self.assertTrue(unrelated.exists())
        self.assertTrue(archived.exists())

    def test_reused_workspace_converts_only_current_run(self):
        for timestamp in ("20260928_080000", "20261006_080000"):
            result = self.run_stage("Clean Previous Scraper Output")
            self.assertEqual(result.returncode, 0, result.stderr)
            for dataset in ("flipkart", "amazon"):
                self.workbook(dataset, timestamp, f"{dataset}:{timestamp}")

            result = self.run_stage("Prepare Output")

            self.assertEqual(result.returncode, 0, result.stderr)
            for dataset in ("flipkart", "amazon"):
                csvs = list((self.workspace / f"output/{dataset}_mobile_scraper").glob("*.csv"))
                self.assertEqual(len(csvs), 1)
                self.assertEqual(csvs[0].read_text(), f"{dataset}:{timestamp}")

    def test_missing_workbook_fails_instead_of_reusing_old_output(self):
        for missing in ("flipkart", "amazon"):
            with self.subTest(dataset=missing):
                self.run_stage("Clean Previous Scraper Output")
                other = "amazon" if missing == "flipkart" else "flipkart"
                self.workbook(other, "20261006_080000", "current")

                result = self.run_stage("Prepare Output")

                self.assertNotEqual(result.returncode, 0)
                self.assertIn("found 0", result.stdout)

    def test_multiple_workbooks_fail_instead_of_selecting_first_match(self):
        for duplicate in ("flipkart", "amazon"):
            with self.subTest(dataset=duplicate):
                self.run_stage("Clean Previous Scraper Output")
                for dataset in ("flipkart", "amazon"):
                    self.workbook(dataset, "20261006_080000", "current")
                self.workbook(duplicate, "20260928_080000", "stale")

                result = self.run_stage("Prepare Output")

                self.assertNotEqual(result.returncode, 0)
                self.assertIn("found 2", result.stdout)

    def test_cleanup_precedes_scrapers_and_concurrent_builds_are_disabled(self):
        pipeline = JENKINSFILE.read_text()
        self.assertIn("disableConcurrentBuilds()", pipeline)
        for stage in ("Run Flipkart Scraper", "Run Amazon Scraper"):
            self.assertLess(
                pipeline.index("stage('Clean Previous Scraper Output')"),
                pipeline.index(f"stage('{stage}')"),
            )


if __name__ == "__main__":
    unittest.main()