import ntpath
import os
import shutil
import tempfile
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
        # Normalisasi: gabungkan host + path tanpa double slash
        full_path = "/".join(p for p in path_parts if p)
        return "https://" + host + "/" + full_path
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


def _safe_dispatch(raw_or_dispatch_obj):
    import win32com.client
    import win32com.client.dynamic

    try:
        return win32com.client.Dispatch(raw_or_dispatch_obj)
    except Exception:
        try:
            return win32com.client.dynamic.Dispatch(raw_or_dispatch_obj)
        except Exception:
            # Jika gen_py corrupt, bersihkan gen_py cache dan coba lagi
            try:
                gen_path = os.path.join(tempfile.gettempdir(), "gen_py")
                if os.path.exists(gen_path):
                    shutil.rmtree(gen_path, ignore_errors=True)
                return win32com.client.dynamic.Dispatch(raw_or_dispatch_obj)
            except Exception:
                return None


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
            raw = running.QueryInterface(pythoncom.IID_IDispatch)
            obj = _safe_dispatch(raw)
            if obj is None:
                continue
            app = obj if getattr(obj, "Name", "") == "Microsoft Excel" else getattr(obj, "Application", None)
            if app is None:
                continue
            hwnd = int(app.Hwnd)
            if hwnd not in seen:
                seen.add(hwnd)
                yield app
        except Exception:
            continue

    # Fallback jika ROT tidak memuat Excel (timing / single instance aktif)
    if not seen:
        try:
            active_app = _safe_dispatch(win32com.client.GetActiveObject("Excel.Application"))
            if active_app is not None:
                hwnd = int(active_app.Hwnd)
                if hwnd not in seen:
                    seen.add(hwnd)
                    yield active_app
        except Exception:
            pass


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
    name_matches = []
    target_norm_path = target["workbookPath"].lower()
    target_norm_name = target["workbookName"].strip().lower()

    for application in applications if applications is not None else iter_excel_applications():
        for workbook in _workbooks(application):
            try:
                wb_full = normalize_windows_path(str(workbook.FullName)).lower()
                wb_name = str(getattr(workbook, "Name", "")).strip().lower()
                if wb_full == target_norm_path:
                    matches.append((application, workbook))
                elif target_norm_name and (
                    wb_name == target_norm_name
                    or wb_full.endswith("/" + target_norm_name)
                    or wb_full.endswith("\\" + target_norm_name)
                ):
                    name_matches.append((application, workbook))
            except Exception:
                continue

    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise TargetWorkbookError("Workbook target ditemukan lebih dari satu kali. Tutup salinan duplikat.")

    # Fallback ke pencocokan nama file jika path OneDrive sync vs web URL sedikit berbeda
    if len(name_matches) == 1:
        return name_matches[0]
    if len(name_matches) > 1:
        raise TargetWorkbookError("Workbook target ditemukan lebih dari satu kali. Tutup salinan duplikat.")

    raise TargetWorkbookError("Workbook target tidak ditemukan. Buka ulang task pane pada workbook tersebut.")


def assert_same_workbook(workbook, normalized_path):
    try:
        current = normalize_windows_path(str(workbook.FullName))
    except Exception as exc:
        raise TargetWorkbookError("Workbook target sudah ditutup.") from exc
    if current.lower() != normalized_path.lower():
        # Fallback: jika nama file tetap sama (misal URL OneDrive berubah saat AutoSave)
        try:
            current_name = str(getattr(workbook, "Name", "")).strip().lower()
            target_name = (
                unquote(urlparse(normalized_path).path).rsplit("/", 1)[-1].lower()
                if normalized_path.lower().startswith("https://")
                else ntpath.basename(normalized_path).lower()
            )
            if current_name and current_name == target_name:
                return
        except Exception:
            pass
        raise TargetWorkbookError("Identitas workbook target berubah. Buka ulang task pane.")
