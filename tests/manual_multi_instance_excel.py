import argparse
import json
import ntpath
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from workbook_target import iter_excel_applications, resolve_workbook


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook_path")
    args = parser.parse_args()
    apps = list(iter_excel_applications())
    inventory = []
    for app in apps:
        inventory.append({
            "hwnd": int(app.Hwnd),
            "workbooks": [str(app.Workbooks.Item(i).FullName) for i in range(1, app.Workbooks.Count + 1)],
        })
    target = {
        "sessionId": "probe-manual",
        "workbookName": ntpath.basename(args.workbook_path),
        "workbookPath": args.workbook_path,
    }
    app, workbook = resolve_workbook(target, applications=apps)
    print(json.dumps({"inventory": inventory, "resolvedHwnd": int(app.Hwnd), "resolvedPath": str(workbook.FullName)}, indent=2))


if __name__ == "__main__":
    main()
