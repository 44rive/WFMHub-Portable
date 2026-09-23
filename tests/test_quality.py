from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from wfmhub.config import ensure_user_config, load_config
from wfmhub.database import write_session
from wfmhub.models import _build_quality


REPO = Path(__file__).resolve().parents[1]


class QualityIssueTests(unittest.TestCase):
    def test_duplicate_findings_do_not_break_refresh(self):
        with tempfile.TemporaryDirectory() as folder:
            home = Path(folder) / "hub"
            (home / "config").mkdir(parents=True)
            for name in (
                "default.toml", "default_rules.toml", "default_metrics.toml",
                "default_analytics.toml", "default_reports.toml",
            ):
                shutil.copy2(REPO / "config" / name, home / "config" / name)
            shutil.copy2(REPO / "config" / "default_queue_mapping.csv", home / "config" / "default_queue_mapping.csv")
            shutil.copytree(REPO / "sql", home / "sql")
            ensure_user_config(home)
            config = load_config(home)

            with write_session(config) as conn:
                values = []
                for number in (1, 2):
                    values.append(
                        (
                            f"failure-{number}",
                            "schedule",
                            f"C:/extracts/attempt-{number}/schedule.txt",
                            "schedule.txt",
                            f"hash-{number}",
                            1,
                            datetime.now(),
                            datetime.now(),
                            False,
                            "ERROR",
                            0,
                            0,
                            "same parser failure",
                            "",
                            0,
                        )
                    )
                conn.executemany(
                    """INSERT INTO meta.source_file(
                           file_id, source_family, source_path, file_name, sha256,
                           size_bytes, modified_at, discovered_at, active, status,
                           row_count, rejected_count, error_message, scope_fingerprint,
                           scoped_out_count
                       ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    values,
                )

                _build_quality(conn, config, "duplicate-test", date(2026, 8, 1), date(2026, 8, 31))

                duplicate_rows = conn.execute(
                    """SELECT count(*) FROM meta.quality_issue
                       WHERE source_file='schedule.txt'
                         AND issue_type='Source load error'
                         AND details='same parser failure'"""
                ).fetchone()[0]
                self.assertEqual(duplicate_rows, 1)


if __name__ == "__main__":
    unittest.main()
