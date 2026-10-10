import tempfile
import unittest
from pathlib import Path
import hashlib

import app


class FakeUpload:
    def __init__(self, name="fake_manual.csv", content=b"uploaded csv"):
        self.name = name
        self.content = content

    def getvalue(self):
        return self.content


class FakeProgress:
    def __init__(self):
        self.updates = []

    def progress(self, value, text=None):
        self.updates.append((value, text))


class FakeStreamlit:
    def __init__(self, upload=None, button_click=True, trigger_upload_change=False):
        self.session_state = {}
        self.downloads = []
        self.messages = []
        self.progress_bar = FakeProgress()
        self.upload = upload or FakeUpload()
        self.button_click = button_click
        self.trigger_upload_change = trigger_upload_change

    def set_page_config(self, **kwargs): pass
    def markdown(self, *args, **kwargs): pass
    def title(self, *args, **kwargs): pass
    def write(self, *args, **kwargs): pass
    def success(self, *args, **kwargs): self.messages.append(("success", args))
    def error(self, *args, **kwargs): self.messages.append(("error", args))
    def warning(self, *args, **kwargs): self.messages.append(("warning", args))
    def info(self, *args, **kwargs): self.messages.append(("info", args))
    def text(self, *args, **kwargs): pass
    def subheader(self, *args, **kwargs): pass
    def caption(self, *args, **kwargs): pass
    def expander(self, *args, **kwargs): return _Context()
    def file_uploader(self, *args, **kwargs):
        if self.trigger_upload_change and kwargs.get("on_change"):
            kwargs["on_change"](self.session_state)
        return self.upload
    def button(self, *args, **kwargs): return self.button_click
    def progress(self, *args, **kwargs): return self.progress_bar
    def spinner(self, *args, **kwargs): return _Context()
    def download_button(self, *args, **kwargs): self.downloads.append((args, kwargs)); return True


class _Context:
    def __enter__(self): return self
    def __exit__(self, *args): return False


class Stage1FrontendTests(unittest.TestCase):
    @staticmethod
    def _detection(_path=None):
        return {"sheets": ["DATA"], "total_records": 3,
                "category_counts": {"rectangular": 3}, "columns": {"DATA": ["Subject", "DUCT SIZE"]},
                "warnings": [], "required_information_available": True}

    def test_generate_flow_shows_summary_and_four_downloads(self):
        with tempfile.TemporaryDirectory() as temp_name:
            root = Path(temp_name)
            artifact = root / "artifact.bin"
            artifact.write_bytes(b"artifact")
            result = {
                "records_processed": 3, "records_with_writes": 3, "zero_write_records": 0,
                "cells_written": 21, "missing_source_data_records": 0, "unsupported_records": 0,
                "other_unresolved_records": 0, "unresolved_records": 0,
                "formula_changes": 0, "structure_preserved": True,
                "job_id": "testjob", "output": artifact, "report": artifact,
                "audit": artifact, "unresolved": artifact,
            }
            received = {}

            def backend(input_file, template_file, output_dir):
                received["input"] = Path(input_file).read_bytes()
                received["template"] = template_file
                received["output_dir"] = output_dir
                return result

            fake_st = FakeStreamlit()
            app.main(fake_st, backend, self._detection,)
            self.assertEqual(received["input"], b"uploaded csv")
            self.assertEqual(len(fake_st.downloads), 4)
            self.assertTrue(any(level == "success" for level, _ in fake_st.messages))
            self.assertEqual(fake_st.session_state["stage1_last_result"], result)

    def test_selecting_new_upload_clears_previous_result_and_downloads(self):
        stale = {"job_id": "old-run"}
        fake_st = FakeStreamlit(upload=FakeUpload("new.csv", b"new input"),
                                button_click=False, trigger_upload_change=True)
        fake_st.session_state["stage1_last_result"] = stale
        fake_st.session_state["stage1_active_upload_fingerprint"] = "old-fingerprint"

        def backend(*args):
            self.fail("generation should not run until the user clicks Generate")

        app.main(fake_st, backend, self._detection)
        self.assertNotIn("stage1_last_result", fake_st.session_state)
        self.assertEqual(fake_st.downloads, [])
        self.assertNotIn(("success", ("Processing completed",)), fake_st.messages)

    def test_failed_new_run_cannot_expose_an_old_download(self):
        fake_st = FakeStreamlit(upload=FakeUpload("new.csv", b"new input"),
                                button_click=True, trigger_upload_change=True)
        fake_st.session_state["stage1_last_result"] = {"job_id": "old-run"}
        fake_st.session_state["stage1_active_upload_fingerprint"] = "old-fingerprint"

        def backend(*args):
            raise RuntimeError("injected failure")

        app.main(fake_st, backend, self._detection)
        self.assertNotIn("stage1_last_result", fake_st.session_state)
        self.assertEqual(fake_st.downloads, [])
        self.assertTrue(any(level == "error" for level, _ in fake_st.messages))

    def test_downloads_use_current_run_workbook_and_report_bytes(self):
        with tempfile.TemporaryDirectory() as temp_name:
            run_dir = Path(temp_name) / "job_run-current"
            run_dir.mkdir()
            expected = {
                "output": b"current workbook bytes",
                "report": b"Processing run ID: run-current\nCurrent validation report",
                "audit": b"current audit",
                "unresolved": b"current unresolved",
            }
            paths = {}
            for key, payload in expected.items():
                paths[key] = run_dir / f"{key}.bin"
                paths[key].write_bytes(payload)
            result = {
                "job_id": "run-current", "run_id": "run-current",
                "output": paths["output"], "report": paths["report"],
                "audit": paths["audit"], "unresolved": paths["unresolved"],
                "output_sha256": hashlib.sha256(expected["output"]).hexdigest(),
                "records_processed": 1, "records_with_writes": 1,
                "cells_written": 6, "unresolved_records": 0,
                "unsupported_records": 0, "formula_changes": 0,
                "structure_preserved": True, "category_counts": {},
            }
            fake_st = FakeStreamlit()
            app.main(fake_st, lambda *args: result, self._detection)
            downloaded = {kwargs["key"].rsplit("_", 1)[-1]: kwargs["data"]
                          for _, kwargs in fake_st.downloads}
            self.assertEqual(downloaded, expected)
            self.assertIn("run-current", expected["report"].decode())


if __name__ == "__main__":
    unittest.main()
