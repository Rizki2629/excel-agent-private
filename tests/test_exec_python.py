import os
import tempfile
import unittest
from unittest.mock import patch

import exec_python
from workbook_target import TargetWorkbookError


TARGET = {"sessionId": "s1", "workbookName": "Nilai.xlsm", "workbookPath": r"C:\Data\Nilai.xlsm"}


class ExecPythonTests(unittest.TestCase):
    @patch("exec_python.assert_same_workbook")
    @patch("exec_python.resolve_workbook")
    def test_execute_envelope_exposes_bound_workbook(self, resolve, assert_same):
        xl, wb = object(), object()
        resolve.return_value = (xl, wb)
        result = exec_python.execute_envelope({"code": "result = (xl, wb)", "target": TARGET})
        self.assertTrue(result["ok"])
        resolve.assert_called_once_with(TARGET)
        self.assertEqual(assert_same.call_count, 2)

    @patch("exec_python.assert_same_workbook")
    @patch("exec_python.resolve_workbook")
    def test_progress_function_writes_to_file(self, resolve, assert_same):
        resolve.return_value = (object(), object())
        with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
            progress_file = f.name
        try:
            payload = {
                "code": 'progress("Membuat header...")',
                "target": TARGET,
                "progressFile": progress_file,
            }
            result = exec_python.execute_envelope(payload)
            self.assertTrue(result["ok"])
            with open(progress_file, "r", encoding="utf-8") as f:
                self.assertEqual(f.read(), "Membuat header...")
        finally:
            os.unlink(progress_file)

    @patch("exec_python.assert_same_workbook")
    @patch("exec_python.resolve_workbook")
    def test_progress_without_file_does_not_fail(self, resolve, assert_same):
        resolve.return_value = (object(), object())
        result = exec_python.execute_envelope({
            "code": 'progress("tanpa file")',
            "target": TARGET,
        })
        self.assertTrue(result["ok"])

    def test_execute_envelope_rejects_missing_code_or_target(self):
        for payload in ({}, {"code": "pass"}, {"code": "", "target": TARGET}):
            with self.subTest(payload=payload):
                result = exec_python.execute_envelope(payload)
                self.assertFalse(result["ok"])

    @patch("exec_python.resolve_workbook")
    def test_execute_returns_target_error_explicitly(self, resolve):
        resolve.side_effect = TargetWorkbookError("target hilang")
        result = exec_python.execute_envelope({"code": "pass", "target": TARGET})
        self.assertFalse(result["ok"])
        self.assertTrue(result["targetError"])
        self.assertIn("target hilang", result["error"])

    @patch("exec_python.assert_same_workbook")
    @patch("exec_python.resolve_workbook")
    def test_execute_code_error_is_not_target_error(self, resolve, assert_same):
        resolve.return_value = (object(), object())
        result = exec_python.execute_envelope({"code": "raise RuntimeError('kode gagal')", "target": TARGET})
        self.assertFalse(result["ok"])
        self.assertFalse(result.get("targetError", False))
        self.assertIn("kode gagal", result["error"])

    @patch("exec_python.assert_same_workbook")
    @patch("exec_python.resolve_workbook")
    def test_printed_error_from_caught_exception_is_execution_failure(self, resolve, assert_same):
        resolve.return_value = (object(), object())
        code = "try:\n    raise RuntimeError('formula ditolak')\nexcept Exception as e:\n    print(f'Error: {e}')"

        result = exec_python.execute_envelope({"code": code, "target": TARGET})

        self.assertFalse(result["ok"])
        self.assertIn("formula ditolak", result["error"])
        self.assertIn("Error: formula ditolak", result["output"])

    @patch("exec_python.resolve_workbook")
    def test_execute_returns_target_error_when_identity_changes_after_exec(self, resolve):
        resolve.return_value = (object(), object())
        with patch("exec_python.assert_same_workbook", side_effect=[None, TargetWorkbookError("identitas berubah")]) as assert_same:
            result = exec_python.execute_envelope({"code": "pass", "target": TARGET})
        self.assertFalse(result["ok"])
        self.assertTrue(result["targetError"])
        self.assertEqual(assert_same.call_count, 2)

    @patch("exec_python.resolve_workbook")
    def test_execute_returns_target_error_when_rot_enumeration_fails(self, resolve):
        resolve.side_effect = TargetWorkbookError("Enumerasi aplikasi Excel gagal.")
        result = exec_python.execute_envelope({"code": "pass", "target": TARGET})
        self.assertFalse(result["ok"])
        self.assertTrue(result["targetError"])
        self.assertIn("Enumerasi", result["error"])


if __name__ == "__main__":
    unittest.main()
