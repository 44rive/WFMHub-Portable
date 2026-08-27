from __future__ import annotations

import shutil
import tempfile
import unittest
from datetime import date
from pathlib import Path

from wfmhub.config import ensure_user_config, load_config, write_source_root
from wfmhub.custom_jobs import run_python_job, run_sql_job
from wfmhub.database import connect, migrate


REPO = Path(__file__).resolve().parents[1]


class CustomLabTests(unittest.TestCase):
    def make_config(self, folder: str):
        home = Path(folder) / "hub"
        (home / "config").mkdir(parents=True)
        shutil.copy2(REPO / "config" / "default.toml", home / "config" / "default.toml")
        shutil.copy2(REPO / "config" / "default_rules.toml", home / "config" / "default_rules.toml")
        shutil.copytree(REPO / "sql", home / "sql")
        config_file = ensure_user_config(home)
        write_source_root(config_file, Path(folder) / "source")
        config = load_config(home)
        migrate(config)
        return config

    def test_python_and_sql_jobs_receive_dates_and_write_outputs(self):
        with tempfile.TemporaryDirectory() as folder:
            config = self.make_config(folder)
            python_job = config.custom / "jobs" / "dates.py"
            python_job.parent.mkdir(parents=True, exist_ok=True)
            python_job.write_text(
                "def run(ctx):\n"
                "    result = ctx.query('SELECT ? AS start_date, ? AS end_date', [ctx.start, ctx.end])\n"
                "    return ctx.write_csv('dates', result)\n",
                encoding="utf-8",
            )
            sql_job = config.custom / "sql" / "dates.sql"
            sql_job.parent.mkdir(parents=True, exist_ok=True)
            sql_job.write_text("SELECT :start AS start_date, :end AS end_date;\n", encoding="utf-8")
            conn = connect(config, read_only=True)
            try:
                py_result = run_python_job(conn, config, python_job, date(2026, 8, 1), date(2026, 8, 31))
                sql_result = run_sql_job(conn, config, sql_job, date(2026, 8, 1), date(2026, 8, 31))
            finally:
                conn.close()
            self.assertTrue(Path(py_result.result).exists())
            self.assertTrue(Path(sql_result.result).exists())
            self.assertIn("2026-08-01", Path(py_result.result).read_text(encoding="utf-8-sig"))

    def test_sql_job_rejects_mutating_statement(self):
        with tempfile.TemporaryDirectory() as folder:
            config = self.make_config(folder)
            job = config.custom / "sql" / "bad.sql"
            job.parent.mkdir(parents=True, exist_ok=True)
            job.write_text("DELETE FROM mart.agent_pcs_day;\n", encoding="utf-8")
            conn = connect(config, read_only=True)
            try:
                with self.assertRaisesRegex(ValueError, "read-only"):
                    run_sql_job(conn, config, job, date(2026, 8, 1), date(2026, 8, 31))
            finally:
                conn.close()


if __name__ == "__main__":
    unittest.main()
