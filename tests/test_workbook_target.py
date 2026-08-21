import unittest
from types import SimpleNamespace
from unittest.mock import patch

from workbook_target import (
    TargetWorkbookError,
    assert_same_workbook,
    iter_excel_applications,
    normalize_windows_path,
    resolve_workbook,
    validate_target,
)


class FakeWorkbook:
    def __init__(self, full_name):
        self.FullName = full_name
        self.Name = full_name.replace("/", "\\").rsplit("\\", 1)[-1]


class FakeApplication:
    def __init__(self, hwnd, paths):
        self.Hwnd = hwnd
        self.Workbooks = [FakeWorkbook(path) for path in paths]


class FakeComWorkbooks:
    def __init__(self, paths):
        self._items = [FakeWorkbook(path) for path in paths]
        self.Count = len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class BrokenWorkbooksApplication:
    Hwnd = 40

    @property
    def Workbooks(self):
        raise RuntimeError("RPC server unavailable")


class BrokenEnumerator:
    def Next(self, count):
        raise RuntimeError("ROT enumeration failed")


class BrokenRunningObjectTable:
    def EnumRunning(self):
        raise RuntimeError("ROT setup failed")


TARGET = {
    "sessionId": "sesi-1",
    "workbookName": "Nilai.xlsm",
    "workbookPath": r"C:\Data\Nilai.xlsm",
}


class WorkbookTargetTests(unittest.TestCase):
    def test_normalize_windows_path_supports_unc_file_uri(self):
        self.assertEqual(
            normalize_windows_path("file://server/share/Nilai.xlsm"),
            r"\\server\share\nilai.xlsm",
        )

    def test_normalize_windows_path_supports_local_drive_file_uri(self):
        self.assertEqual(
            normalize_windows_path("file:///C:/Data/Nilai.xlsm"),
            r"c:\data\nilai.xlsm",
        )

    def test_normalize_windows_path_supports_onedrive_url(self):
        self.assertEqual(
            normalize_windows_path(
                "https://D.DOCS.LIVE.NET/A652565C59200A84/Documents/Daftar%20hadir.xlsx"
            ),
            "https://d.docs.live.net/a652565c59200a84/Documents/Daftar hadir.xlsx",
        )

    def test_resolve_onedrive_matches_encoded_target_to_com_fullname(self):
        target = {
            "sessionId": "onedrive",
            "workbookName": "Daftar hadir.xlsx",
            "workbookPath": "https://d.docs.live.net/A652/Documents/Daftar%20hadir.xlsx",
        }
        app = FakeApplication(50, ["https://d.docs.live.net/a652/Documents/Daftar hadir.xlsx"])
        xl, wb = resolve_workbook(target, applications=[app])
        self.assertEqual(xl.Hwnd, 50)
        self.assertEqual(wb.Name, "Daftar hadir.xlsx")

    def test_validate_target_normalizes_path(self):
        target = validate_target(TARGET)
        self.assertEqual(target["workbookPath"], r"c:\data\nilai.xlsm")

    def test_validate_target_rejects_relative_or_missing_values(self):
        for bad in ({}, {**TARGET, "sessionId": ""}, {**TARGET, "workbookPath": "Book1"}):
            with self.subTest(bad=bad), self.assertRaises(TargetWorkbookError):
                validate_target(bad)

    def test_resolve_workbook_across_applications(self):
        apps = [FakeApplication(10, [r"C:\Data\Lain.xlsx"]), FakeApplication(20, [r"C:\Data\Nilai.xlsm"])]
        xl, wb = resolve_workbook(TARGET, applications=apps)
        self.assertEqual(xl.Hwnd, 20)
        self.assertEqual(wb.Name, "Nilai.xlsm")

    def test_resolve_workbook_supports_com_count_item_collection(self):
        app = FakeApplication(30, [])
        app.Workbooks = FakeComWorkbooks([TARGET["workbookPath"]])
        xl, wb = resolve_workbook(TARGET, applications=[app])
        self.assertEqual(xl.Hwnd, 30)
        self.assertEqual(wb.Name, "Nilai.xlsm")

    def test_resolve_rejects_zero_or_duplicate_matches(self):
        with self.assertRaisesRegex(TargetWorkbookError, "tidak ditemukan"):
            resolve_workbook(TARGET, applications=[])
        duplicate = [FakeApplication(10, [TARGET["workbookPath"]]), FakeApplication(20, [TARGET["workbookPath"]])]
        with self.assertRaisesRegex(TargetWorkbookError, "lebih dari satu"):
            resolve_workbook(TARGET, applications=duplicate)

    def test_resolve_wraps_workbooks_rpc_failure_as_target_error(self):
        with self.assertRaisesRegex(TargetWorkbookError, "Workbooks"):
            resolve_workbook(TARGET, applications=[BrokenWorkbooksApplication()])

    def test_iter_excel_applications_wraps_next_failure_as_target_error(self):
        pythoncom = SimpleNamespace(
            GetRunningObjectTable=lambda: SimpleNamespace(
                EnumRunning=lambda: BrokenEnumerator()
            )
        )
        win32com_client = SimpleNamespace()
        with patch.dict(
            "sys.modules",
            {
                "pythoncom": pythoncom,
                "win32com": SimpleNamespace(client=win32com_client),
                "win32com.client": win32com_client,
            },
        ):
            with self.assertRaisesRegex(TargetWorkbookError, "enumerasi"):
                list(iter_excel_applications())

    def test_iter_excel_applications_wraps_rot_setup_failures_as_target_error(self):
        failures = [
            lambda: (_ for _ in ()).throw(RuntimeError("ROT unavailable")),
            lambda: BrokenRunningObjectTable(),
        ]
        for get_rot in failures:
            with self.subTest(get_rot=get_rot):
                pythoncom = SimpleNamespace(GetRunningObjectTable=get_rot)
                win32com_client = SimpleNamespace()
                with patch.dict(
                    "sys.modules",
                    {
                        "pythoncom": pythoncom,
                        "win32com": SimpleNamespace(client=win32com_client),
                        "win32com.client": win32com_client,
                    },
                ):
                    with self.assertRaisesRegex(TargetWorkbookError, "enumerasi"):
                        list(iter_excel_applications())

    def test_assert_same_workbook_detects_rename(self):
        wb = FakeWorkbook(r"C:\Data\Baru.xlsm")
        with self.assertRaisesRegex(TargetWorkbookError, "berubah"):
            assert_same_workbook(wb, normalize_windows_path(TARGET["workbookPath"]))


if __name__ == "__main__":
    unittest.main()
