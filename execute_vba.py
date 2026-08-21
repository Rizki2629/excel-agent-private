"""Execute VBA code in the active Excel workbook via COM automation."""
import sys, json, re

TMP_MODULE = "ExcelAgentTmp"
TMP_PROC = "ExcelAgentRun"


def strip_wrapper(code):
    """Buang Sub/Function wrapper jika AI tetap menuliskannya."""
    lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    start = None
    end = None
    head = re.compile(r"^\s*(public\s+|private\s+)?(sub|function)\s+\w+", re.I)
    tail = re.compile(r"^\s*end\s+(sub|function)\s*$", re.I)

    for i, line in enumerate(lines):
        if start is None and head.match(line):
            start = i
        elif tail.match(line):
            end = i

    if start is not None and end is not None and end > start:
        lines = lines[start + 1:end]

    # Buang Option Explicit (tidak valid di tengah module yang digenerate)
    lines = [l for l in lines if not re.match(r"^\s*option\s+explicit\s*$", l, re.I)]

    return "\n".join(lines).strip()


def run(vba_code):
    try:
        import win32com.client
    except ImportError:
        return {"ok": False, "error": "pywin32 belum terinstall. Jalankan: pip install pywin32"}

    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return {"ok": False, "error": "Excel tidak sedang berjalan."}

    wb = excel.ActiveWorkbook
    if wb is None:
        return {"ok": False, "error": "Tidak ada workbook aktif."}

    try:
        vbproj = wb.VBProject
    except Exception:
        return {
            "ok": False,
            "error": (
                "Akses VBProject diblokir. Aktifkan: File > Options > Trust Center > "
                "Trust Center Settings > Macro Settings > centang "
                "'Trust access to the VBA project object model'."
            ),
        }

    body = strip_wrapper(vba_code)
    if not body:
        return {"ok": False, "error": "Kode VBA kosong setelah diproses."}

    mod = None
    try:
        # Hapus module sisa run sebelumnya
        try:
            vbproj.VBComponents.Remove(vbproj.VBComponents(TMP_MODULE))
        except Exception:
            pass

        mod = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.Name = TMP_MODULE

        # Inner sub berisi kode asli; wrapper menangkap runtime error dan
        # mengembalikannya sebagai string agar tidak muncul dialog modal
        # yang membekukan COM.
        wrapped = (
            "Public Function {proc}() As String\n"
            "    On Error GoTo Gagal\n"
            "    Call {proc}Body\n"
            "    {proc} = \"\"\n"
            "    Exit Function\n"
            "Gagal:\n"
            "    {proc} = \"Error \" & Err.Number & \": \" & Err.Description\n"
            "End Function\n"
            "\n"
            "Private Sub {proc}Body()\n"
            "{body}\n"
            "End Sub"
        ).format(proc=TMP_PROC, body=body)
        mod.CodeModule.AddFromString(wrapped)

        excel.DisplayAlerts = False
        err = excel.Run("{0}.{1}".format(TMP_MODULE, TMP_PROC))
        excel.DisplayAlerts = True

        if err:
            return {"ok": False, "error": str(err)}
        return {"ok": True, "message": "Kode dijalankan di " + str(wb.Name)}

    except Exception as e:
        msg = str(e)
        m = re.search(r"'([^']*)'\s*$", msg)
        return {"ok": False, "error": m.group(1) if m else msg}

    finally:
        try:
            excel.DisplayAlerts = True
        except Exception:
            pass
        if mod is not None:
            try:
                vbproj.VBComponents.Remove(mod)
            except Exception:
                pass


if __name__ == "__main__":
    result = run(sys.stdin.read())
    print(json.dumps(result, ensure_ascii=False))
