import io
import re
import json
import os
import sys
import traceback

from workbook_target import TargetWorkbookError, assert_same_workbook, resolve_workbook, validate_target

# Konstanta Excel yang sering digunakan AI dalam kode win32com.
# Tanpa ini, kode seperti ws.Cells(...).End(xlUp) akan gagal karena xlUp tidak terdefinisi.
EXCEL_CONSTANTS = {
    # Navigasi
    "xlUp": -4162,
    "xlDown": -4121,
    "xlToLeft": -4159,
    "xlToRight": -4161,
    # Alignment horizontal
    "xlCenter": -4108,
    "xlLeft": -4131,
    "xlRight": -4152,
    "xlJustify": -4130,
    "xlDistributed": -4117,
    "xlFill": -4122,
    "xlCenterAcrossSelection": -4108,
    "xlGeneral": 1,
    # Alignment vertikal
    "xlTop": -4160,
    "xlBottom": -4107,
    "xlVAlignCenter": -4108,
    "xlVAlignTop": -4160,
    "xlVAlignBottom": -4107,
    "xlVAlignJustify": -4130,
    "xlVAlignDistributed": -4117,
    # Border line style
    "xlContinuous": 1,
    "xlDash": -4115,
    "xlDashDot": 4,
    "xlDashDotDot": 5,
    "xlDot": -4118,
    "xlDouble": -4119,
    "xlSlantDashDot": 13,
    "xlLineStyleNone": -4142,
    # Border weight
    "xlHairline": 1,
    "xlThin": 2,
    "xlMedium": -4138,
    "xlThick": 4,
    # Border index
    "xlEdgeLeft": 7,
    "xlEdgeTop": 8,
    "xlEdgeBottom": 9,
    "xlEdgeRight": 10,
    "xlInsideVertical": 11,
    "xlInsideHorizontal": 12,
    "xlDiagonalDown": 5,
    "xlDiagonalUp": 6,
    # Fill / Interior
    "xlSolid": 1,
    "xlGray16": 17,
    "xlGray25": -4124,
    "xlGray50": -4125,
    "xlGray75": -4126,
    "xlAutomatic": -4105,
    "xlNone": -4142,
    # Color index
    "xlColorIndexAutomatic": -4105,
    "xlColorIndexNone": -4142,
    # Paste special
    "xlValues": -4163,
    "xlFormulas": -4123,
    "xlComments": -4144,
    "xlFormats": -4122,
    "xlAll": -4104,
    "xlAllExceptBorders": 7,
    # Fill direction
    "xlByRows": 1,
    "xlByColumns": 2,
    # Page setup
    "xlLandscape": 2,
    "xlPortrait": 1,
    "xlPaperA4": 9,
    "xlPaperLetter": 1,
    "xlPaperLegal": 5,
    "xlPaperFolio": 14,
    # Page break
    "xlPageBreakAutomatic": -4105,
    "xlPageBreakManual": -4135,
    "xlPageBreakNone": -4142,
    # Print titles
    "xlRepeatRows": 1,
    "xlRepeatColumns": 2,
    # Number format
    "xlTextFormat": "@",
    "xlGeneralFormat": "General",
    # Chart type (umum)
    "xlColumnClustered": 51,
    "xlBarClustered": 57,
    "xlLine": 4,
    "xlPie": 5,
    "xlArea": 1,
    "xlXYScatter": -4169,
    # List object
    "xlListObject": 1,
    "xlSrcRange": 1,
    "xlYes": 1,
    "xlNo": 2,
    # Sort
    "xlAscending": 1,
    "xlDescending": 2,
    "xlSortOnCellValues": 0,
    "xlSortOnCellColor": 1,
    "xlSortOnFontColor": 2,
    "xlSortOnIcon": 3,
    # Filter
    "xlFilterValues": 7,
    "xlFilterCellColor": 8,
    "xlFilterFontColor": 9,
    "xlFilterIcon": 10,
    # Orientation
    "xlHorizontal": 1,
    "xlVertical": 2,
    # Wrap text
    "xlWrapText": 1,
    # Shrink to fit
    "xlShrinkToFit": 1,
    # Merge
    "xlMergeCells": 1,
    # Protection
    "xlProtect": 1,
    "xlUnprotect": 2,
    # Visible
    "xlSheetVisible": -1,
    "xlSheetHidden": 0,
    "xlSheetVeryHidden": 2,
}


def _make_progress_writer(progress_file):
    """Buat fungsi progress yang menulis pesan ke file temp."""
    def progress(msg):
        if not progress_file:
            return
        try:
            with open(progress_file, "w", encoding="utf-8") as f:
                f.write(str(msg))
        except Exception:
            pass
    return progress


def execute_envelope(payload):
    captured = io.StringIO()
    real_stdout = sys.stdout
    try:
        if not isinstance(payload, dict):
            raise ValueError("Payload eksekusi harus berupa objek JSON.")
        code = payload.get("code")
        if not isinstance(code, str) or not code.strip():
            raise ValueError("Kode Python kosong.")
        raw_target = payload.get("target")
        target = validate_target(raw_target)
        xl, wb = resolve_workbook(raw_target)
        assert_same_workbook(wb, target["workbookPath"])
        progress_file = payload.get("progressFile", "")
        progress = _make_progress_writer(progress_file)
        env = {"xl": xl, "wb": wb, "progress": progress, "__builtins__": __builtins__}
        env.update(EXCEL_CONSTANTS)
        # Tangkap print() dari kode AI: berguna sebagai bahan diagnostik saat
        # eksekusi gagal sehingga AI bisa memperbaiki kodenya sendiri.
        sys.stdout = captured
        try:
            exec(compile(code, "<agent>", "exec"), env)
        finally:
            sys.stdout = real_stdout
        assert_same_workbook(wb, target["workbookPath"])
        result = {"ok": True, "message": "Berhasil dieksekusi pada " + target["workbookName"]}
        output = captured.getvalue()
        if output.strip():
            result["output"] = output[-4000:]
            error_line = re.search(r"(?im)^\s*(?:Error|Traceback)\s*:\s*(.+)$", output)
            if error_line:
                result["ok"] = False
                result["error"] = error_line.group(1).strip()
                result["traceback"] = output[-4000:]
        return result
    except TargetWorkbookError as exc:
        sys.stdout = real_stdout
        return {
            "ok": False,
            "targetError": True,
            "error": str(exc),
            "traceback": traceback.format_exc(),
        }
    except Exception as exc:
        sys.stdout = real_stdout
        result = {"ok": False, "error": str(exc), "traceback": traceback.format_exc()}
        output = captured.getvalue()
        if output.strip():
            result["output"] = output[-4000:]
        return result


def main():
    try:
        payload = json.loads(sys.stdin.read())
    except Exception as exc:
        print(json.dumps({"ok": False, "error": "Envelope JSON tidak valid: " + str(exc)}))
        return
    print(json.dumps(execute_envelope(payload)))

if __name__ == "__main__":
    main()
