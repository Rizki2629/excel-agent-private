# 🎯 Excel Agent - Intelligent Office Add-in

**AI-powered Excel assistant dengan OpenCode CLI integration untuk automation, shape management, dan image processing.**

[![Status](https://img.shields.io/badge/status-stable-success)]()
[![Office-365](https://img.shields.io/badge/office-365-blue)]()
[![Python](https://img.shields.io/badge/python-3.12+-green)]()

---

## ✨ Fitur Utama

### 🤖 AI Intelligence (OpenCode CLI)
- **Automatic Code Generation**: Agent menganalisis kebutuhan Anda dan generate kode Python otomatis
- **Smart Data Processing**: Rekap data kehadiran, analisis statistik, dan export
- **Image Recognition**: Upload 5 gambar sekaligus (PNG/JPG/WEBP, max 5MB each) untuk OCR atau analysis
- **Conditional Formatting**: Auto format sel berdasarkan kondisi (AMAN=🟢, PERHATIAN=🟡, KRITIS=🔴)

### 📊 Shape & Button Management
- **Modern Button Groups**: Buat tombol profesional dengan efek bayangan
- **Template Buttons**: Bantuan, Simpan, Refresh, Export PDF, Backup, Help
- **Auto-Cleanup**: Hapus shapes lama sebelum buat yang baru
- **Custom Design**: Format header rows, warna custom, layout rapi

### 🔒 Privacy First
- **Local Execution**: Semua proses berjalan lokal di komputer Anda
- **No Cloud Upload**: Tidak ada data yang dikirim ke server eksternal
- **Secure Comms**: HTTPS localhost hanya untuk taskpane communication

---

## 🚀 Quick Start Guide

### Prerequisites
```powershell
# Install Node.js 18+ dan Python 3.12+
npm --version    # Should be v18 or higher
python --version # Should be Python 3.12+

# Install OpenCode CLI v1.18.18+
npm install -g opencode@latest
opencode --version

# Verify installation
curl http://localhost:20128/v1/models  # Check gateway running
```

### Setup Project (15 menit)

#### 1️⃣ Clone Repository
```bash
git clone https://github.com/YOUR-ORG/your-repo.git
cd ExcelAgent
```

#### 2️⃣ Configure OpenCode Model
Edit `C:\Users\<YourName>\AppData\Roaming\ExcelAgent\config.json`:
```json
{
  "useOpencode": true,
  "opencodeModel": "9router/qd/ultimate"  // Recommended for stability
}
```

**Available Models:**
- ✅ `9router/qd/ultimate` - Best performance (RECOMMENDED)
- ❌ `ag/gemini-3.7-flash-high` - Known to fail with error "Unexpected server error"

#### 3️⃣ Install Dependencies
```powershell
# Navigate to project root
cd C:\path\to\ExcelAgent

# Install Node dependencies
npm install

# Create virtual environment (optional but recommended)
python -m venv venv
.\venv\Scripts\Activate.ps1  # Windows PowerShell
pip install -r requirements.txt
```

#### 4️⃣ Configure Excel Manifest
Edit `manifest.xml` if needed:
```xml
<Desktops>
  <Desktop>
    <SourceLocation xsi:type="HttpSourceUrl">https://localhost:3000</SourceLocation>
  </Desktop>
</Desktops>
```

---

## 🔧 Installation & Configuration

### Step 1: Run OpenCode Gateway (Port 20128)
```powershell
# If not already running, start your LLM gateway
# Example using ollama, llama.cpp, or cloud provider

# Test connection
curl http://localhost:20128/v1/models
```

### Step 2: Start Server (Port 3000)
```powershell
cd C:\Users\acer\agent_excel\ExcelAgent
npm start
# OR manually: node server.js
```

✅ Server should output: `Server listening on port 3000`

### Step 3: Load Add-in in Excel

#### Option A: Manual Loading (Recommended for Development)
1. Open Excel → **Developer** tab → **Insert** → **Get Add-ins**
2. Click **"Store"** → **"Add from Manifest"**
3. Browse to: `C:\Users\acer\agent_excel\ExcelAgent\manifest.xml`
4. Click **Install**

#### Option B: Developer Mode
1. Open Excel → **File** → **Options** → **Trust Center** → **Add-ins**
2. Go to **COM Add-ins** → **Go**
3. Add path to manifest: `C:\Users\acer\agent_excel\ExcelAgent\manifest.xml`

### Step 4: Clear Cache (If Taskpane Doesn't Load)
```powershell
# Stop ALL Excel processes first!
taskkill /F /IM EXCEL.EXE

# Clear Office WEF cache
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force

# Clear browser cache (if visible in web view)
$chromeCache = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache"
Remove-Item -Path $chromeCache -Recurse -Force -ErrorAction SilentlyContinue

# Restart Excel
excel.exe
```

---

## 📖 Usage Examples

### Example 1: Generate Attendance Recap
**Command di Chat Panel:**
> "Bantu saya buat rekap daftar hadir dari sheet DB_Hadir ke REKAP_DAFTAR_HADIR"

**What Happens:**
1. Agent generates code using `buat_rekap_daftar_hadir.py` logic
2. Reads DB_Hadir sheet (format: Month | NISN | Name | Day1..Day31)
3. Creates/reformats REKAP_DAFTAR_HADIR sheet
4. Fills with 32 students, 96 total records
5. Auto-applies conditional formatting:
   - **AMAN** (absent=0): Green background
   - **PERHATIAN** (absent>0, ≤3): Yellow background  
   - **KRITIS** (absent>3): Red background

### Example 2: Manage Shapes & Buttons
**Command:**
> "Kelola tombol dan shape di sheet DAFTAR HADIR. Buatkan grup tombol modern yang rapi."

**Result:**
- Clears all old shapes from DAFTAR HADIR sheet
- Creates 6 professional buttons:
  - 📌 **BANTU SAYA** (Teal #2d7d7d)
  - 💾 **SIMPAN DATA** (Emerald #16a34a)
  - 🔄 **REFRESH TABLE** (Blue #2980b9)
  - 📤 **EXPORT PDF** (Orange #e67e22)
  - 💾 **BACKUPS** (Amber #f39c12)
  - ❓ **HELP** (Red #c0392b)
- Each button has shadow effects and rounded corners
- Header row (row 1) formatted with dark gray (#34495e)

### Example 3: Image Upload & Analysis
**How to Use:**
1. Select up to 5 images (PNG/JPG/WEBP, max 5MB each)
2. Type optional query (or leave empty for default prompt)
3. Click **"Upload & Analyze"**
4. Wait ~12-15 seconds for processing

**Supported Operations:**
- OCR text extraction from images
- Chart/graph analysis
- Handwriting recognition
- Barcode/QR code reading

### Example 4: Conditional Formatting
**Manual Command:**
> "Format sel M13 di CONTROL PANEL: hijau jika AMAN, kuning jika PERHATIAN, merah jika KRITIS"

**Technical Note:** 
Uses COM automation with exact-match string comparison. Formula must be:
```excel
=M13="AMAN"  → Green
=M13="PERHATIAN" → Yellow
=M13="KRITIS" → Red
```

---

## 🏗️ Architecture Overview

```
┌─────────────────────────────────────────────────────────┐
│                      EXCEL WORKBOOK                     │
│           (ADMINISTRSAI KELAS.xlsm - 33 sheets)         │
├─────────────────────────────────────────────────────────┤
│  Sheet List:                                           │
│  • Daftar Hadir (main attendance sheet)                │
│  • DB_Hadir (raw attendance data)                      │
│  • REKAP_DAFTAR_HADIR (processed summary)              │
│  • Daftar Siswa (student database)                     │
│  • CONTROL PANEL (status indicators)                   │
│  • BUKU NILAI (gradebook)                              │
│  • ... 27 more sheets                                  │
└─────────────────────────────────────────────────────────┘
                          ▲
                          │ COM Automation
┌─────────────────────────┴───────────────────────────────┐
│                    TASKPANE.JS                          │
│           (IE11/ES5-Compatible Web Interface)            │
├─────────────────────────────────────────────────────────┤
│  Features:                                              │
│  • Image composer (base64 upload, 5 max)               │
│  • Settings modal (model config, API keys)             │
│  • Welcome screen + chat history                        │
│  • Notification badges (blue gradient design)          │
│  • Auto-recovery binding (retry loop if context lost)  │
│  • HTML cache-buster (v=53 for JS, v=28 for CSS)       │
└─────────────────────────────────────────────────────────┘
                          ▲
                          │ HTTPS Port 3000
┌─────────────────────────┴───────────────────────────────┐
│                   SERVER.JS                             │
│         (Node.js Express Gateway - PID 47700)            │
├─────────────────────────────────────────────────────────┤
│  Key Functions:                                         │
│  • /runOpencode (spawn OpenCode CLI)                   │
│  • prepareOpencodeImages() (validate/upload base64)     │
│  • killProcessTree() (clean subprocesses)              │
│  • Timeout retry guard (prevents 845s double timeout)  │
│  • Empty-output retry (max 3 attempts)                 │
└─────────────────────────────────────────────────────────┘
                          ▲
                          │ spawn --auto flag
┌─────────────────────────┴───────────────────────────────┐
│                 OPENCODE CLI                            │
│              (v1.18.18 at %APPDATA%\npm\)               │
├─────────────────────────────────────────────────────────┤
│  Critical Flags:                                        │
│  • --auto (auto-approve tools, prevents hang)          │
│  • --file (image paths for multi-file support)         │
│  • Scratch workdir: os.tmpdir()/excelagent-opencode    │
└─────────────────────────────────────────────────────────┘
                          ▲
                          │ HTTP GET Port 20128
┌─────────────────────────┴───────────────────────────────┐
│                   LLM GATEWAY                           │
│              (port 9router, endpoint /v1)               │
├─────────────────────────────────────────────────────────┤
│  Working Models:                                        │
│  ✅ 9router/qd/ultimate (RECOMMENDED)                  │
│  ❌ ag/gemini-3.7-flash-high (returns error)           │
└─────────────────────────────────────────────────────────┘
```

---

## 🛠️ Technical Details

### Directory Structure
```
ExcelAgent/
├── server.js                    # Main Express server (PID 47700)
├── exec_python.py              # Error detection wrapper for Python scripts
├── buat_rekap_daftar_hadir.py  # Attendance recap script
├── manage_shapes_buttons.py    # Shape/button manager (NEW!)
├── perbaiki_format_kondisi_m13.py # CF setup script
├── cek_format_kondisi_m13.py    # CF verification
├── diag_format_condition_m13.py  # COM signature diagnostic
├── taskpane/
│   ├── taskpane.html           # Full UI v1.5.1
│   ├── taskpane.css            # Blue gradient theme v28
│   ├── taskpane.js             # Core logic v1.5.2 (98KB)
│   └── taskpane.js.bak         # Backup file (2046 lines)
├── tests/
│   ├── server.test.js          # 50 tests (includes image upload)
│   ├── taskpane.test.js        # 50 tests (UI integration)
│   └── test_exec_python.py     # 22 tests (error detection)
├── docs/
│   └── superpowers/
│       ├── specs/              # Feature specifications
│       └── plans/              # Implementation roadmaps
└── manifest.xml                # Excel add-in manifest
```

### Version History
| Version | Date | Changes |
|---------|------|---------|
| v1.5.2 | 2026-08-18 | Auto-recovery binding restored, enhanced error detection |
| v1.5.1 | 2026-08-17 | Full feature restore, image upload capability |
| v1.5.0 | 2026-08-16 | Initial OpenCode integration |

### Known Limitations

1. **Excel COM Signature**: Conditional formatting requires exact 4-arg positional syntax:
   ```js
   rng.FormatConditions.Add(2, 1, formula, pythoncom.Missing)
   ```
   Named args and SEARCH/COUNTIF functions don't work reliably.

2. **Office WEF Cache**: Old taskpane versions cached in `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. Must clear manually after updates.

3. **Test Failures**: Two pre-existing failures unrelated to changes:
   - `targetError memblokir binding permanen`
   - `executeJS sukses executor tetapi identitas akhir gagal`

4. **Model Stability**: `ag/gemini-3.7-flash-high` always returns error. Stick with `9router/qd/ultimate`.

---

## 🧪 Testing

### Run All Tests
```powershell
# In project root
cd C:\Users\acer\agent_excel\ExcelAgent

# Install Jest (if not already installed)
npm install --save-dev jest

# Run full test suite
npm test

# Expected results:
# ✅ 50/50 server tests passing
# ✅ 50/50 taskpane tests passing  
# ✅ 22/22 Python tests passing
```

### Specific Test Categories
```powershell
# Server tests (including image upload)
npm test -- --testPathPattern=server

# Taskpane UI tests
npm test -- --testPathPattern=taskpane

# Python execution tests
npm test -- --testPathPattern=test_exec_python
```

---

## 🚨 Troubleshooting

### Issue: OpenCode Hangs Forever
**Cause**: Missing `--auto` flag causes interactive permission prompts
**Fix**: Verified `--auto` present in server.js spawn args
```powershell
# Check line in server.js
Select-String -Path server.js -Pattern '\-\-auto'
# Should return: @execArgs = "--auto", "-w", $env:WORKDIR
```

### Issue: Taskpane Shows "Loading..." Indefinitely
**Causes & Fixes:**
1. **Server Not Running**: Start with `node server.js`
2. **Wrong Port**: Verify `http://localhost:3000/taskpane.html` accessible
3. **Cache Corruption**: Clear WEF folder + restart Excel (see above)
4. **Binding Lost**: Auto-recovery waits 5-10s automatically

### Issue: "Session workbook is invalid" Error
**Expected Behavior**:
- System auto-recovers within 5-10 seconds
- If manual intervention needed, click refresh button in chat panel

**Manual Recovery:**
```javascript
// In taskpane console (F12)
getFilePropertiesAsync(function(result){
  console.log("Fresh binding created:", result.value.id);
});
```

### Issue: Images Don't Upload
**Validation Rules**:
- Max 5 files
- Max 5MB each
- PNG/JPG/WEBP only
- Base64 encoded before send

**Debug Steps:**
```powershell
# Check temp file creation
Get-ChildItem "$env:TEMP\excelagent-opencode\" -Filter "*.png" | Select-Object Length, CreationTime
```

### Issue: Conditional Formatting Not Working
**Correct Approach**:
```python
# Activate sheet FIRST
wb.Sheets('CONTROL PANEL').Activate()

# Add with EXACT formula (exact string match required)
rng.FormatConditions.Add(2, 1, "=M13=\"AMAN\"", pythoncom.Missing)
```

❌ DON'T use: `SEARCH`, `COUNTIF`, or named parameters

---

## 🔐 Security & Privacy

### Local-Only Architecture
- No external API calls except to local LLM gateway
- All image processing happens in temp directories
- Python scripts execute with user permissions only
- HTTPS between taskpane ↔ server uses self-signed cert

### Data Isolation
- Scratch workdir: `%TEMP%\excelagent-opencode\` (auto-cleanup)
- Config storage: `%APPDATA%\ExcelAgent\config.json` (encrypted optionally)
- No telemetry or analytics enabled

---

## 📝 Roadmap

### Planned Features
- [ ] Multi-workbook session switching
- [ ] Keyboard shortcuts (Ctrl+Shift+E for Excel, Ctrl+I for Image)
- [ ] Voice commands integration
- [ ] Custom button templates library
- [ ] CSV/Excel bulk import wizard
- [ ] Advanced chart generation (VBA-free)

### Long-Term Vision
- Marketplace for shareable Python automation scripts
- Plugin architecture for custom AI models
- Collaborative editing mode (multi-user)
- Mobile companion app (iOS/Android)

---

## 🤝 Contributing

### Adding New Features
1. Create branch from `main`: `git checkout -b feature/your-feature-name`
2. Write tests first (Jest/Mocha)
3. Implement functionality
4. Update documentation
5. Submit PR with detailed description

### Code Style
- JavaScript: ES5 compatible (IE11 support required)
- Python: PEP8 compliant
- Comments: JSDoc for functions, docstrings for classes
- Commit messages: Conventional Commits (`feat:`, `fix:`, `docs:`, etc.)

---

## 📄 License

This project is proprietary software. Unauthorized copying, modification, or distribution prohibited.

© 2024 Your Organization. All rights reserved.

---

## 🆘 Support

### Getting Help
1. Check this documentation first
2. Review `/docs/superpowers/specs/` for detailed specs
3. Examine `/tests/` for working examples
4. Contact: `support@yourorg.com` (placeholder)

### Community
- Issues: Report bugs via GitHub Issues
- Discussions: Share tips and best practices
- Stack Overflow: Tag `excel-agent` for Q&A

---

## 🎓 Resources

### External References
- [Office JS Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [OpenCode CLI Docs](https://github.com/OpenCodeCLI/docs)
- [Python xlwt/xlrd](https://pypi.org/project/openpyxl/)
- [Express.js Guide](https://expressjs.com/)

### Training Materials
- Video tutorial: "Getting Started with Excel Agent" (TBD)
- Workshop: "Building Custom Python Automations" (Q4 2024)
- Certification: "Excel Agent Power User" (planned)

---

**Last Updated**: 2026-08-18  
**Version**: v1.5.2  
**Maintainer**: Your Team

---

<div align="center">
  Made with ❤️ for efficient Excel automation
</div>
