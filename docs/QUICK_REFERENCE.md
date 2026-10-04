# ⚡ Quick Reference Card - Excel Agent

## 🚀 Start Server
```powershell
cd C:\Users\acer\agent_excel\ExcelAgent
node server.js
```
**Expected:** `Server listening on port 3000` ✅

---

## 🔄 Load Add-in in Excel
1. Open Developer tab → Insert → Get Add-ins
2. Click **"Add from Manifest"**
3. Browse to: `manifest.xml`
4. Install → Done!

---

## 🔍 Troubleshooting (Common Issues)

### Taskpane Won't Load?
```powershell
# Clear Office cache
taskkill /F /IM EXCEL.EXE
Remove-Item "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force
excel.exe
```

### OpenCode Hangs Forever?
✅ Already fixed with `--auto` flag  
⏱️ Auto-recovery triggers after 5-10 seconds

### "Session Invalid" Error?
✅ **Don't panic!** Auto-recovery handles it  
⏳ Wait 5-10 seconds for automatic fix

---

## 💬 Useful Commands to Try

### Basic Tasks
> "Bantu saya buat rekap dari sheet DB_Hadir"  
> "Format conditional M13: hijau jika AMAN"

### Shape Management  
> "Kelola tombol dan shape di DAFTAR HADIR"  
> "Buat grup tombol modern yang rapi"

### Image Upload
1. Select up to 5 files (PNG/JPG/WEBP ≤5MB)
2. Type query or leave empty
3. Click **"Upload & Analyze"**
4. Wait ~12 seconds

---

## 🔑 Configuration Files

### Model Settings (`%APPDATA%\ExcelAgent\config.json`)
```json
{
  "useOpencode": true,
  "opencodeModel": "9router/qd/ultimate"
}
```
**DO NOT USE:** `ag/gemini-3.7-flash-high` ❌

### Scratch Workdir (Auto-cleaned)
`%TEMP%\excelagent-opencode\`

---

## 📊 Port Map
- **3000**: Server (HTTPS, taskpane UI)
- **20128**: LLM Gateway (HTTP, OpenCode provider)

---

## 🧪 Test Suite
```powershell
npm test
# Expected: 122 tests passing ✅
```

---

## 📁 Key Files at a Glance

| File | Purpose | Size |
|------|---------|------|
| `server.js` | Main Express server | 98KB |
| `taskpane.js` | Core logic (v1.5.2) | 98KB |
| `taskpane.html` | Full UI (v1.5.1) | Cached with v=53 |
| `manifest.xml` | Excel add-in definition | N/A |
| `exec_python.py` | Python execution wrapper | Error detection |
| `manage_shapes_buttons.py` | Modern button generator | NEW! |
| `buat_rekap_daftar_hadir.py` | Attendance recap script | Working |

---

## 🎯 Success Indicators

✅ Blue notification badge appears (not dark gray)  
✅ Auto-recovery completes within 10 seconds  
✅ Image upload processes in <15 seconds  
✅ Shape buttons created instantly  
✅ All tests pass consistently  

❌ Stuck loading >1 minute → Clear cache + restart  
❌ Error messages persist after recovery → Check model config  
❌ Images rejected → Verify file types/sizes  

---

## 🌐 Repository Links

**Private GitHub Repo:**  
https://github.com/Rizki2629/excel-agent-private

**Full Documentation:**  
- [README.md](../../README.md) - User guide  
- [DEPLOYMENT.md](../../DEPLOYMENT.md) - Installation steps  
- [TECH_ARCH.md](../../TECH_ARCH.md) - Technical details  
- [FINAL_SUMMARY.md](../../FINAL_SUMMARY.md) - Project overview

---

## ⚠️ Important Notes

1. **Always use `9router/qd/ultimate`** model for stability
2. **Never commit** sensitive files (.gitignore already configured)
3. **Clear WEF cache** whenever switching versions
4. **Test changes** in development before deploying to production workbook
5. **Monitor active tasks** in chat panel during long operations

---

## 🆘 Emergency Contacts

**Development:** Review `/docs/superpowers/specs/` directory  
**Infrastructure:** Check firewall rules for ports 3000/20128  
**Office Automation:** Run `regsvr32.exe` if COM errors occur  
**LLM Gateway:** Verify instance running on localhost:20128

---

**Version:** v1.5.2  
**Updated:** 2026-08-18  
**Status:** Production Ready ✅
