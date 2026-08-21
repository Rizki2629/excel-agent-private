"""
Excel tools for the CLI agent.
Uses COM automation to read/write Excel.
"""
import json
import sys
import os

try:
    import win32com.client
    HAS_COM = True
except ImportError:
    HAS_COM = False


class ExcelTools:
    """Collection of Excel operations callable by the agent."""

    def __init__(self):
        self.excel = None
        self.connect()

    def connect(self):
        """Connect to running Excel instance."""
        if not HAS_COM:
            raise RuntimeError("pywin32 not installed. Run: pip install pywin32")
        try:
            self.excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            raise RuntimeError("Excel tidak sedang berjalan. Buka Excel dulu.")

    def get_context(self):
        """Get current workbook context for AI."""
        wb = self.excel.ActiveWorkbook
        ws = self.excel.ActiveSheet

        context = {
            "workbook": {
                "name": wb.Name,
                "sheetCount": wb.Worksheets.Count,
                "sheets": []
            },
            "activeSheet": {
                "name": ws.Name
            }
        }

        # List all sheets
        for i in range(1, wb.Worksheets.Count + 1):
            s = wb.Worksheets(i)
            used_rows = s.UsedRange.Rows.Count
            used_cols = s.UsedRange.Columns.Count
            context["workbook"]["sheets"].append({
                "name": s.Name,
                "rows": used_rows,
                "cols": used_cols,
                "visible": s.Visible == -1  # xlSheetVisible = -1
            })

        # Active sheet details
        used = ws.UsedRange
        if used.Rows.Count > 1 or used.Columns.Count > 1:
            context["activeSheet"]["range"] = used.Address
            context["activeSheet"]["rowCount"] = used.Rows.Count
            context["activeSheet"]["colCount"] = used.Columns.Count

            # Headers (row 1)
            if used.Rows.Count > 0:
                headers = []
                for c in range(1, min(used.Columns.Count + 1, 30)):
                    headers.append(str(ws.Cells(1, c).Value or ""))
                context["activeSheet"]["headers"] = headers

            # Sample data (first 5 rows)
            sample = []
            max_rows = min(used.Rows.Count, 6)
            max_cols = min(used.Columns.Count, 20)
            for r in range(1, max_rows + 1):
                row = []
                for c in range(1, max_cols + 1):
                    val = ws.Cells(r, c).Value
                    row.append(val)
                sample.append(row)
            context["activeSheet"]["sampleData"] = sample

            # Formulas
            formulas = []
            for r in range(1, min(used.Rows.Count + 1, 50)):
                for c in range(1, min(used.Columns.Count + 1, 20)):
                    cell = ws.Cells(r, c)
                    if cell.HasFormula:
                        formulas.append({
                            "cell": ws.Cells(r, c).Address(False, False),
                            "formula": cell.Formula
                        })
            context["activeSheet"]["formulas"] = formulas[:20]

        return context

    def read_range(self, sheet_name=None, range_addr=None):
        """Read data from a range."""
        ws = self.excel.ActiveSheet if not sheet_name else self.excel.Worksheets(sheet_name)

        if range_addr:
            rng = ws.Range(range_addr)
        else:
            rng = ws.UsedRange

        data = []
        for r in range(1, rng.Rows.Count + 1):
            row = []
            for c in range(1, rng.Columns.Count + 1):
                val = rng.Cells(r, c).Value
                row.append(val)
            data.append(row)

        return {
            "ok": True,
            "sheet": ws.Name,
            "range": rng.Address,
            "data": data,
            "rows": rng.Rows.Count,
            "cols": rng.Columns.Count
        }

    def write_range(self, sheet_name, range_addr, values):
        """Write values to a range."""
        ws = self.excel.ActiveSheet if not sheet_name else self.excel.Worksheets(sheet_name)
        rng = ws.Range(range_addr)
        rng.Value = values
        return {"ok": True, "message": f"Tulis ke {range_addr} berhasil"}

    def write_vba(self, vba_code):
        """Execute VBA code via COM."""
        wb = self.excel.ActiveWorkbook

        # Remove existing temp module if any
        try:
            wb.VBProject.VBComponents.Remove(wb.VBProject.VBComponents("TempAgentModule"))
        except Exception:
            pass

        # Add new module
        module = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule = 1
        module.Name = "TempAgentModule"
        module.CodeModule.AddFromString(vba_code)

        return {"ok": True, "message": "VBA code injected successfully"}

    def search(self, text, sheet_name=None):
        """Search for text in worksheet."""
        ws = self.excel.ActiveSheet if not sheet_name else self.excel.Worksheets(sheet_name)
        used = ws.UsedRange

        results = []
        found = used.Find(text, LookAt=1)  # xlPart = 1
        if found:
            first_addr = found.Address
            while True:
                results.append({
                    "address": found.Address(False, False),
                    "value": found.Value
                })
                found = used.FindNext(found)
                if found is None or found.Address == first_addr:
                    break

        return {"ok": True, "results": results, "count": len(results)}

    def list_sheets(self):
        """List all sheets with metadata."""
        wb = self.excel.ActiveWorkbook
        sheets = []
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            sheets.append({
                "name": ws.Name,
                "index": i,
                "visible": ws.Visible == -1,
                "rows": ws.UsedRange.Rows.Count,
                "cols": ws.UsedRange.Columns.Count
            })
        return {"ok": True, "sheets": sheets}

    def get_cell(self, cell_addr, sheet_name=None):
        """Get detailed info about a single cell."""
        ws = self.excel.ActiveSheet if not sheet_name else self.excel.Worksheets(sheet_name)
        cell = ws.Range(cell_addr)

        return {
            "ok": True,
            "address": cell.Address(False, False),
            "value": cell.Value,
            "formula": cell.Formula if cell.HasFormula else None,
            "numberFormat": cell.NumberFormat,
            "font": {
                "name": cell.Font.Name,
                "size": cell.Font.Size,
                "bold": cell.Font.Bold,
                "italic": cell.Font.Italic,
                "color": cell.Font.Color
            },
            "fill": {
                "color": cell.Interior.Color if cell.Interior.ColorIndex != -4142 else None
            }
        }

    def screenshot(self):
        """Take screenshot of Excel window."""
        try:
            from PIL import ImageGrab
            import win32gui

            hwnd = self.excel.Hwnd
            rect = win32gui.GetWindowRect(hwnd)
            img = ImageGrab.grab(bbox=rect)
            path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'excel_screenshot.png')
            img.save(path)
            return {"ok": True, "path": path}
        except Exception as e:
            return {"ok": False, "error": str(e)}


# Tool definitions for the agent
TOOLS = [
    {
        "name": "get_excel_context",
        "description": "Get current workbook context including sheet names, headers, sample data, and formulas",
        "parameters": {}
    },
    {
        "name": "read_range",
        "description": "Read data from a specific sheet and range",
        "parameters": {
            "sheet_name": {"type": "string", "description": "Sheet name (optional, defaults to active)"},
            "range_addr": {"type": "string", "description": "Range address like 'A1:D10' (optional, defaults to used range)"}
        }
    },
    {
        "name": "write_range",
        "description": "Write values to a range in a sheet",
        "parameters": {
            "sheet_name": {"type": "string", "description": "Sheet name"},
            "range_addr": {"type": "string", "description": "Range address"},
            "values": {"type": "array", "description": "2D array of values"}
        }
    },
    {
        "name": "write_vba",
        "description": "Execute VBA code in Excel",
        "parameters": {
            "vba_code": {"type": "string", "description": "VBA code to execute"}
        }
    },
    {
        "name": "search_excel",
        "description": "Search for text in a worksheet",
        "parameters": {
            "text": {"type": "string", "description": "Text to search for"},
            "sheet_name": {"type": "string", "description": "Sheet name (optional)"}
        }
    },
    {
        "name": "list_sheets",
        "description": "List all sheets in the workbook",
        "parameters": {}
    },
    {
        "name": "get_cell",
        "description": "Get detailed info about a single cell",
        "parameters": {
            "cell_addr": {"type": "string", "description": "Cell address like 'A1'"},
            "sheet_name": {"type": "string", "description": "Sheet name (optional)"}
        }
    },
    {
        "name": "screenshot",
        "description": "Take a screenshot of the Excel window",
        "parameters": {}
    }
]
