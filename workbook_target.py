import ntpath
from urllib.parse import unquote, urlparse


class TargetWorkbookError(RuntimeError):
    pass


def normalize_windows_path(value):
    if not isinstance(value, str) or not value.strip():
        raise TargetWorkbookError("Path workbook target kosong.")
    value = value.strip()
    if value.lower().startswith(("https://", "http://")):
        parsed = urlparse(value)
        host = parsed.netloc.lower()
        if parsed.scheme.lower() != "https" or not (
            host == "d.docs.live.net" or host.endswith(".sharepoint.com")
        ):
            raise TargetWorkbookError("URL workbook cloud tidak didukung.")
        path_parts = unquote(parsed.path).split("/")
        if host == "d.docs.live.net" and len(path_parts) > 1:
            path_parts[1] = path_parts[1].lower()
        return "https://" + host + "/".join(path_parts)
    if value.lower().startswith("file:"):
        parsed = urlparse(value)
        path = unquote(parsed.path)
        value = "//" + parsed.netloc + path if parsed.netloc else path.lstrip("/")
    normalized = ntpath.normcase(ntpath.normpath(value))
    if not ntpath.isabs(normalized):
        raise TargetWorkbookError("Workbook harus disimpan sebelum agent dijalankan.")
    return normalized


def validate_target(target):
    if not isinstance(target, dict):
        raise TargetWorkbookError("Metadata target workbook tidak valid.")
    result = {}
    for key in ("sessionId", "workbookName", "workbookPath"):
        value = target.get(key)
        if not isinstance(value, str) or not value.strip():
            raise TargetWorkbookError("Metadata target workbook tidak lengkap: " + key)
        result[key] = value.strip()
    result["workbookPath"] = normalize_windows_path(result["workbookPath"])
    target_name = (
        unquote(urlparse(result["workbookPath"]).path).rsplit("/", 1)[-1]
        if result["workbookPath"].lower().startswith("https://")
        else ntpath.basename(result["workbookPath"])
    )
    if target_name.lower() != result["workbookName"].lower():
        raise TargetWorkbookError("Nama dan path workbook target tidak cocok.")
    return result


def iter_excel_applications():
    import pythoncom
    import win32com.client

    try:
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
    except Exception as exc:
        raise TargetWorkbookError(
            "Gagal menyiapkan enumerasi aplikasi Excel. Coba lagi saat Excel siap."
        ) from exc
    seen = set()
    while True:
        try:
            monikers = enum.Next(1)
        except Exception as exc:
            raise TargetWorkbookError(
                "Gagal melakukan enumerasi aplikasi Excel. Coba lagi saat Excel siap."
            ) from exc
        if not monikers:
            break
        moniker = monikers[0]
        try:
            running = rot.GetObject(moniker)
            obj = win32com.client.Dispatch(running.QueryInterface(pythoncom.IID_IDispatch))
            app = obj if getattr(obj, "Name", "") == "Microsoft Excel" else obj.Application
            hwnd = int(app.Hwnd)
            if hwnd not in seen:
                seen.add(hwnd)
                yield app
        except Exception:
            continue


def _workbooks(application):
    try:
        collection = application.Workbooks
    except Exception as exc:
        raise TargetWorkbookError(
            "Koleksi Workbooks Excel tidak dapat dibaca. Coba lagi saat Excel siap."
        ) from exc
    try:
        return [collection.Item(index) for index in range(1, collection.Count + 1)]
    except AttributeError:
        try:
            return list(collection)
        except Exception as exc:
            raise TargetWorkbookError(
                "Koleksi Workbooks Excel tidak dapat dibaca. Coba lagi saat Excel siap."
            ) from exc
    except Exception as exc:
        raise TargetWorkbookError(
            "Koleksi Workbooks Excel tidak dapat dibaca. Coba lagi saat Excel siap."
        ) from exc


def resolve_workbook(target, applications=None):
    target = validate_target(target)
    matches = []
    for application in applications if applications is not None else iter_excel_applications():
        for workbook in _workbooks(application):
            try:
                if normalize_windows_path(str(workbook.FullName)) == target["workbookPath"]:
                    matches.append((application, workbook))
            except Exception:
                continue
    if not matches:
        raise TargetWorkbookError("Workbook target tidak ditemukan. Buka ulang task pane pada workbook tersebut.")
    if len(matches) != 1:
        raise TargetWorkbookError("Workbook target ditemukan lebih dari satu kali. Tutup salinan duplikat.")
    return matches[0]


def assert_same_workbook(workbook, normalized_path):
    try:
        current = normalize_windows_path(str(workbook.FullName))
    except Exception as exc:
        raise TargetWorkbookError("Workbook target sudah ditutup.") from exc
    if current != normalized_path:
        raise TargetWorkbookError("Identitas workbook target berubah. Buka ulang task pane.")
