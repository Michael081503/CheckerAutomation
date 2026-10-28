"""Exercise in-process Cloud updates in an isolated Python interpreter."""

import os
import subprocess
import sys
import textwrap
import unittest
from pathlib import Path


class CloudStartupTests(unittest.TestCase):
    def test_cached_previous_release_modules_are_refreshed_before_checking(self):
        root = Path(__file__).resolve().parents[1]
        script = textwrap.dedent("""
            import io
            import sys
            from pathlib import Path
            from unittest.mock import patch

            import checker
            import web_checker
            from streamlit.testing.v1 import AppTest

            sys.path.insert(0, str(Path.cwd() / "tests"))
            from test_sbf_pl_viber import BLAST_HEADERS, pl_database, pl_row
            from test_web_checker import csv_upload

            class BrowserFile(io.BytesIO):
                def __init__(self, upload):
                    super().__init__(upload.data)
                    self.name = upload.name

            database = BrowserFile(pl_database([
                ["Ana Peña", "00123", "000001", "a@example.test", "09171234567", 1001],
            ]))
            files = [BrowserFile(csv_upload("viber.csv", [BLAST_HEADERS, pl_row(name="Ana Pe\ufffda")]))]

            def uploader(label, *args, **kwargs):
                return database if kwargs.get("key", "").startswith("database_upload_") else files

            def outdated_detector(*args, **kwargs):
                raise AssertionError("The previous release's upload adapter was reused")

            # Streamlit can retain imported modules while rerunning a new entrypoint.
            if sys.argv[1] == "engine-and-adapter":
                del checker.is_datagrid_pl_viber_blast
            if sys.argv[1] == "previous-comparison":
                checker.COMPARISON_VERSION = 0
            if sys.argv[1] == "previous-endorsement":
                checker.COMPARISON_VERSION = 1
                del checker.is_salad_endo_pair
            if sys.argv[1] == "previous-regular":
                checker.COMPARISON_VERSION = 2
            web_checker.AUTO_DETECTION_VERSION = 7
            if sys.argv[1] == "previous-comparison":
                web_checker.AUTO_DETECTION_VERSION = 8
            if sys.argv[1] == "previous-endorsement":
                web_checker.AUTO_DETECTION_VERSION = 9
            if sys.argv[1] == "previous-regular":
                web_checker.AUTO_DETECTION_VERSION = 10
            web_checker._key_column = outdated_detector

            with patch("streamlit.file_uploader", side_effect=uploader):
                app = AppTest.from_file(str(Path.cwd() / "streamlit_app.py"), default_timeout=20)
                app.run()
                assert not app.exception, [item.message for item in app.exception]
                app.button(key="run_check").click().run()
                assert not app.exception, [item.message for item in app.exception]
                run = app.session_state["check_run"]
                assert run.results[0].database_match_column == "Debtor ID"
                assert run.results[0].count("MATCHED") == 1
                assert web_checker.AUTO_DETECTION_VERSION == 11
                assert checker.COMPARISON_VERSION == 3
                assert callable(checker.is_salad_endo_pair)

                # Healthy modules and the completed result survive normal reruns.
                options_type = checker.Options
                detect = web_checker.prepare_uploaded_check
                app.run()
                assert not app.exception, [item.message for item in app.exception]
                assert checker.Options is options_type
                assert web_checker.prepare_uploaded_check is detect
                assert app.session_state["check_run"] is run
        """)
        for cached in ("engine-and-adapter", "adapter-only", "previous-comparison", "previous-endorsement", "previous-regular"):
            with self.subTest(cached=cached):
                result = subprocess.run(
                    [sys.executable, "-c", script, cached], cwd=root,
                    capture_output=True, text=True, timeout=60,
                    env={**os.environ, "OPENBLAS_NUM_THREADS": "1", "OMP_NUM_THREADS": "1"},
                )
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
