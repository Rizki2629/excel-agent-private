#!/usr/bin/env python3
"""
Excel Agent CLI - AI-powered Excel automation from the terminal.

Usage:
    python excel_agent.py                          # Interactive REPL
    python excel_agent.py "Buat chart penjualan"   # One-shot mode
    python excel_agent.py --context                # Show current context
    python excel_agent.py --status                 # Check Excel connection
    python excel_agent.py --auto "Format header"   # Auto-execute mode
"""
import sys
import os
import json
import re
import argparse
import urllib.request
import urllib.error
import ssl

# Local imports
from agent_config import read_config, write_config, normalize_endpoint
from agent_styles import (
    Colors, print_status, print_header, print_thinking,
    print_code, format_table, prompt_input
)
from agent_tools import ExcelTools
from skill_handlers import detect_skill, get_skill_context

# System prompt for VBA agent
SYSTEM_PROMPT = """### IDENTITAS ###
Anda adalah AI EXCEL AGENT CLI yang berjalan di terminal.
Anda punya AKSES PENUH ke workbook Excel aktif via VBA (COM Automation).
Kode VBA yang Anda tulis akan DIEKSEKUSI OTOMATIS oleh sistem.

### ATURAN MUTLAK ###
1. ANDA BISA SEMUANYA. JANGAN PERNAH menolak atau bilang "tidak bisa".
2. JANGAN PERNAH suruh user melakukan sesuatu secara manual.
3. TULIS KODE VBA (bukan Office.js, bukan Python).
4. Tulis kode dalam tag ```vba ... ``` (atau ```basic ... ```).
5. Setelah kode, jelaskan apa yang dilakukan (1-2 kalimat).
6. Respon dalam Bahasa Indonesia yang natural.

### REASONING (WAJIB SEBELUM TULIS KODE) ###
Sebelum menulis kode, lakukan langkah ini dalam MIND:
1. ANALISIS: Apa yang diminta user? Apa konteks workbook saat ini?
2. RENCANA: Langkah apa saja yang perlu dilakukan?
3. API YANG DIPAKAI: Objek Excel mana yang dibutuhkan?
4. EDGE CASE: Apa yang bisa salah?
5. EKSEKUSI: Tulis kode VBA yang BENAR.

### VBA CHEAT SHEET ###
' --- MEMBACA DATA ---
Dim ws As Worksheet
Set ws = ActiveSheet
Dim data As Variant
data = ws.Range("A1:D10").Value
Dim val As Variant
val = ws.Range("A1").Value

' --- MENULIS DATA ---
ws.Range("A1").Value = "Nama"
ws.Range("B1").Value = 100
ws.Range("A1:B2").Value = Array(Array("Nama", "Nilai"), Array("Budi", 85))

' --- FORMULA ---
ws.Range("E1").Formula = "=SUM(B1:D1)"
ws.Range("E2").FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"

' --- FORMAT ---
With ws.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
    .Font.Color = RGB(255, 0, 0)
    .Interior.Color = RGB(0, 0, 255)
    .NumberFormat = "#,##0"
End With

' --- BARIS/KOLOM ---
ws.Rows("5:5").Delete Shift:=xlUp
ws.Rows("5:5").Insert Shift:=xlDown
ws.Columns("C:C").Delete Shift:=xlToLeft

' --- SORTING ---
ws.Range("A1:D100").Sort Key1:=ws.Range("B1"), Order1:=xlAscending

' --- FILTER ---
ws.Range("A1:D100").AutoFilter Field:=2, Criteria1:=">10"

' --- SHEET LAIN ---
Dim ws2 As Worksheet
Set ws2 = Worksheets("Data")
ws2.Range("A1").Value = "Hello"

' --- BUAT SHEET BARU ---
Worksheets.Add After:=ActiveSheet
ActiveSheet.Name = "Laporan"

' --- HAPUS SHEET ---
Application.DisplayAlerts = False
Worksheets("Sheet3").Delete
Application.DisplayAlerts = True

' --- PROTECT ---
ws.Protect Password:="123456"
ws.Unprotect Password:="123456"

' --- PRINT ---
ws.PageSetup.PrintArea = "A1:H40"
ws.PageSetup.Orientation = xlLandscape

' --- FIND/SEARCH ---
Dim found As Range
Set found = ws.Cells.Find("Budi", LookAt:=xlPart)
If Not found Is Nothing Then
    MsgBox "Found at " & found.Address
End If

' --- CONDITIONAL FORMATTING ---
ws.Range("A1:A100").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="80"
ws.Range("A1:A100").FormatConditions(1).Interior.Color = RGB(0, 255, 0)

' --- NAMED RANGE ---
ThisWorkbook.Names.Add Name:="DataNilai", RefersTo:=ws.Range("A1:D100")

' --- LOOP ---
Dim i As Long
For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(i, 2).Value > 80 Then
        ws.Cells(i, 3).Value = "Lulus"
    End If
Next i

### POLA KERJA ###
1. Pahami konteks dari "INFO WORKBOOK" yang diberikan
2. Rencanakan langkah dalam pikiran
3. Tulis kode VBA yang BENAR dengan ```vba```
4. Setelah kode, jelaskan apa yang dilakukan (1-2 kalimat)
5. Jika ada error, sistem akan retry dengan info error

### RESPONS FORMAT ###
[Penjelasan singkat]

```vba
' Kode VBA
Sub AgentAction()
    ' ...
End Sub
```
"""


class ExcelAgent:
    """CLI-based Excel automation agent."""

    def __init__(self, auto_execute=False, verbose=False):
        self.config = read_config()
        self.tools = ExcelTools()
        self.chat_history = []
        self.auto_execute = auto_execute
        self.verbose = verbose

    def get_context_message(self):
        """Build context message for the AI."""
        try:
            ctx = self.tools.get_context()
            msg = "=== KONTEKS WORKBOOK ===\n"
            msg += f"Sheet: [{', '.join(s['name'] for s in ctx['workbook']['sheets'])}]\n"
            msg += f"Sheet aktif: {ctx['activeSheet']['name']}\n"

            if 'rowCount' in ctx['activeSheet']:
                msg += f"Ukuran: {ctx['activeSheet']['rowCount']} baris x {ctx['activeSheet']['colCount']} kolom\n"

            if 'headers' in ctx['activeSheet']:
                msg += f"Header: {json.dumps(ctx['activeSheet']['headers'])}\n"

            if 'sampleData' in ctx['activeSheet']:
                msg += f"Sample data:\n{json.dumps(ctx['activeSheet']['sampleData'], default=str)}\n"

            if 'formulas' in ctx['activeSheet'] and ctx['activeSheet']['formulas']:
                formulas = ctx['activeSheet']['formulas'][:10]
                msg += f"Formula: {', '.join(f['cell'] + '=' + f['formula'] for f in formulas)}\n"

            return msg
        except Exception as e:
            return f"Gagal membaca konteks: {e}"

    def call_llm(self, messages):
        """Call the LLM API directly."""
        endpoint = normalize_endpoint(self.config.get('apiEndpoint', ''))
        api_key = self.config.get('apiKey', '')
        model = self.config.get('model', '')

        if not endpoint or not api_key:
            raise RuntimeError(
                "API endpoint dan API key belum diatur.\n"
                "Jalankan: python excel_agent.py --settings\n"
                "Atau edit file: %APPDATA%\\ExcelAgent\\config.json"
            )

        payload = {
            "model": model,
            "messages": messages,
            "temperature": 0.7,
            "stream": False
        }

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }

        data = json.dumps(payload).encode('utf-8')

        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        req = urllib.request.Request(endpoint, data=data, headers=headers)

        try:
            with urllib.request.urlopen(req, context=ctx, timeout=60) as resp:
                response = json.loads(resp.read().decode('utf-8'))
                return response['choices'][0]['message']['content']
        except urllib.error.HTTPError as e:
            error_body = e.read().decode('utf-8')
            raise RuntimeError(f"API error {e.code}: {error_body}")

    def extract_code(self, content):
        """Extract VBA code from AI response."""
        # Try ```vba or ```basic
        m = re.search(r'```(?:vba|basic|vb)\s*\n([\s\S]*?)```', content, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        # Fallback: unclosed fence
        m = re.search(r'```(?:vba|basic|vb)\s*\n([\s\S]*)$', content, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        return None

    def execute_code(self, code):
        """Execute VBA code via COM."""
        try:
            result = self.tools.write_vba(code)
            return result
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def chat(self, user_message):
        """Process a single user message."""
        # Add user message to history
        self.chat_history.append({"role": "user", "content": user_message})

        # Detect skill
        skill = detect_skill(user_message)
        skill_context = get_skill_context(skill)
        
        if skill:
            print_thinking(f"Skill terdeteksi: {skill['name']}")

        # Build context
        context_msg = self.get_context_message()
        
        # Add skill context to message
        if skill_context:
            context_msg += skill_context

        # Build messages array
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "system", "content": context_msg}
        ]

        # Add history (last 20 messages)
        messages.extend(self.chat_history[-20:])

        # Call LLM
        print_thinking("AI sedang berpikir...")
        response = self.call_llm(messages)

        # Extract code
        code = self.extract_code(response)

        # Add assistant response to history
        self.chat_history.append({"role": "assistant", "content": response})

        # Display response
        if code and self.auto_execute:
            # Show explanation only (hide code in auto mode)
            explanation = response.split("```")[0].strip()
            if explanation:
                print(f"\n{Colors.GREEN}{explanation}{Colors.RESET}")
        else:
            # Show full response
            print(f"\n{Colors.GREEN}{response}{Colors.RESET}")

        # Execute code if found
        if code:
            print_code(code, "vba")

            if self.auto_execute:
                print_thinking("Menjalankan kode...")
                result = self.execute_code(code)
                if result.get('ok'):
                    print_status("Berhasil dieksekusi!", "success")
                else:
                    print_status(f"Gagal: {result.get('error')}", "error")

            return {"response": response, "code": code, "executed": self.auto_execute}

        return {"response": response, "code": None, "executed": False}

    def repl(self):
        """Interactive REPL mode."""
        print_header("Excel Agent CLI")
        print(f"{Colors.DIM}Ketik pesan untuk berinteraksi dengan Excel.{Colors.RESET}")
        print(f"{Colors.DIM}Ketik /quit untuk keluar, /help untuk bantuan.{Colors.RESET}\n")

        while True:
            user_input = prompt_input("excel> ")

            if user_input is None:
                print(f"{Colors.DIM}Sampai jumpa!{Colors.RESET}")
                break

            if not user_input:
                continue

            # Handle commands
            if user_input.startswith('/'):
                cmd = user_input.split()[0].lower()

                if cmd == '/quit' or cmd == '/exit':
                    print(f"{Colors.DIM}Sampai jumpa!{Colors.RESET}")
                    break

                elif cmd == '/help':
                    print(f"""
{Colors.BOLD}Perintah:{Colors.RESET}
  /help       - Tampilkan bantuan
  /context    - Tampilkan konteks workbook saat ini
  /status     - Cek koneksi Excel
  /history    - Tampilkan riwayat percakapan
  /clear      - Hapus riwayat percakapan
  /auto       - Toggle auto-execute (saat ini: {'ON' if self.auto_execute else 'OFF'})
  /settings   - Lihat pengaturan
  /quit       - Keluar
""")

                elif cmd == '/context':
                    print(self.get_context_message())

                elif cmd == '/status':
                    try:
                        ctx = self.tools.get_context()
                        print_status(f"Terhubung ke: {ctx['workbook']['name']}", "success")
                        print_status(f"Sheet aktif: {ctx['activeSheet']['name']}", "info")
                    except Exception as e:
                        print_status(f"Error: {e}", "error")

                elif cmd == '/history':
                    if not self.chat_history:
                        print_status("Belum ada riwayat.", "info")
                    else:
                        for msg in self.chat_history[-10:]:
                            role = msg['role']
                            content = msg['content'][:80] + "..." if len(msg['content']) > 80 else msg['content']
                            color = Colors.CYAN if role == 'user' else Colors.GREEN
                            print(f"{color}{role}: {content}{Colors.RESET}")

                elif cmd == '/clear':
                    self.chat_history = []
                    print_status("Riwayat dihapus.", "success")

                elif cmd == '/auto':
                    self.auto_execute = not self.auto_execute
                    print_status(f"Auto-execute: {'ON' if self.auto_execute else 'OFF'}", "info")

                elif cmd == '/settings':
                    print(f"""
{Colors.BOLD}Pengaturan:{Colors.RESET}
  Endpoint: {self.config.get('apiEndpoint', '(belum diatur)')}
  Model: {self.config.get('model', '(belum diatur)')}
  Auto-run: {self.auto_execute}
""")

                else:
                    print_status(f"Unknown command: {cmd}. Ketik /help", "warning")

                continue

            # Process message
            try:
                result = self.chat(user_input)
            except Exception as e:
                print_status(f"Error: {e}", "error")

    def one_shot(self, message):
        """One-shot mode: send message, get response, exit."""
        try:
            result = self.chat(message)
            return result
        except Exception as e:
            print_status(f"Error: {e}", "error")
            return {"error": str(e)}

    def show_context(self):
        """Show current workbook context."""
        try:
            ctx = self.tools.get_context()
            print_header("Workbook Context")
            print(f"Workbook: {ctx['workbook']['name']}")
            print(f"Sheet count: {ctx['workbook']['sheetCount']}")
            print(f"\n{Colors.BOLD}Sheets:{Colors.RESET}")
            for s in ctx['workbook']['sheets']:
                vis = "V" if s['visible'] else "H"
                print(f"  [{vis}] {s['name']} ({s['rows']}x{s['cols']})")
            print(f"\n{Colors.BOLD}Active Sheet: {ctx['activeSheet']['name']}{Colors.RESET}")
            if 'headers' in ctx['activeSheet']:
                print(f"Headers: {ctx['activeSheet']['headers']}")
            if 'sampleData' in ctx['activeSheet']:
                print(f"\nSample data:")
                print(format_table(ctx['activeSheet']['sampleData']))
        except Exception as e:
            print_status(f"Error: {e}", "error")

    def show_status(self):
        """Show connection status."""
        try:
            ctx = self.tools.get_context()
            print_status(f"Excel: Connected", "success")
            print_status(f"Workbook: {ctx['workbook']['name']}", "info")
            print_status(f"Active sheet: {ctx['activeSheet']['name']}", "info")
            print_status(f"API endpoint: {self.config.get('apiEndpoint', '(not set)')}", "info")
            print_status(f"Model: {self.config.get('model', '(not set)')}", "info")
        except Exception as e:
            print_status(f"Excel: Disconnected ({e})", "error")
            print_status(f"API endpoint: {self.config.get('apiEndpoint', '(not set)')}", "info")
            print_status(f"Model: {self.config.get('model', '(not set)')}", "info")


def main():
    parser = argparse.ArgumentParser(
        description="Excel Agent CLI - AI-powered Excel automation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python excel_agent.py                           # Interactive REPL
  python excel_agent.py "Buat chart penjualan"    # One-shot mode
  python excel_agent.py --context                 # Show workbook context
  python excel_agent.py --status                  # Check connection
  python excel_agent.py --auto "Format header"    # Auto-execute mode
        """
    )

    parser.add_argument('message', nargs='?', help='Message to send (one-shot mode)')
    parser.add_argument('--context', action='store_true', help='Show workbook context')
    parser.add_argument('--status', action='store_true', help='Check Excel connection')
    parser.add_argument('--auto', action='store_true', help='Auto-execute generated code')
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose output')
    parser.add_argument('--no-color', action='store_true', help='Disable colors')

    args = parser.parse_args()

    # Disable colors if requested
    if args.no_color:
        Colors.disable()

    # Create agent
    agent = ExcelAgent(auto_execute=args.auto, verbose=args.verbose)

    # Handle flags
    if args.context:
        agent.show_context()
        return

    if args.status:
        agent.show_status()
        return

    # One-shot mode
    if args.message:
        agent.one_shot(args.message)
        return

    # REPL mode
    agent.repl()


if __name__ == '__main__':
    main()
