# 🎯 EXCEL AGENT - QUICK START GUIDE
## **HANYA BUTUH 3 KLIK!** ✨

---

## ⭐ CARA TERMUDAH: LAUNCH_ALL.bat (RECOMMENDED!)

### Langkah 1: Buka Folder Project
```
C:\Users\acer\agent_excel\ExcelAgent\
```

### Langkah 2: Klik Double **LAUNCH_ALL.bat**
![Bat file icon](https://img.shields.io/badge/icon-click-orange)

### Langkah 3: Tunggu!
✅ Server starts  
✅ Excel opens  
✅ Ready to use in 10 seconds!

**Done!** ✅

---

## 📋 PANDUAN LENGKAP SEMUA SCRIPT

### 🚀 **LAUNCH_ALL.bat** (PAKAI INI SAJA!)
**Icon:** 🚀 Rocket  
**Fungsi:** Mulai semuanya dalam SATU KLIK  
**Kapan dipakai:** Setiap hari saat mau kerja

```
Cara pakai: Double-click → Done! ✅
Time: ~10 detik total
```

---

### 🎮 **MENU_CONTROL.bat** (Kontrol Manual)
**Icon:** 🎮 Game controller  
**Tampilan menu:**
```
[1] 🚀 LAUNCH ALL (Pakai ini aja!)
[2] 🖥️  Start Server Only
[3] 📁 Load Add-in ke Excel
[4] 🗑️  Clear All Cache
[5] ⏹️  Stop Everything
[6] 🔍 Cek Status
[Q] Quit
```

**Kapan dipakai:** 
- Debugging masalah
- Kontrol granular tiap service
- Check status server/Excel

---

### 🖥️ **START_SERVER.bat** (Server Saja)
**Icon:** 🖥️ Computer  
**Fungsi:** Jalankan server saja (tanpa buka Excel)  
**Kapan dipakai:**
- Restart server tanpa restart Excel
- Test server issues
- Background mode

---

### 📁 **LOAD_ADDIN.bat** (Buka Excel Saja)
**Icon:** 📁 File folder  
**Fungsi:** Buka workbook Excel Anda  
**Prerequisite:** Server HARUS sudah jalan dulu!

---

### 🗑️ **CLEAR_CACHE.bat** (Reset Cache)
**Icon:** 🗑️ Trash bin  
**Fungsi:** Hapus semua cache Office + Chrome  
**Kapan dipakai:**
- Taskpane tidak muncul ❌
- Error "Session invalid" persisten
- After update code baru
- Excel loading lama banget

---

### ⏹️ **STOP_ALL.bat** (Matikan Semua)
**Icon:** ⏹️ Stop sign  
**Fungsi:** Stop server + optional close Excel  
**Kapan dipakai:**
- Selesai work
- Before deploy changes
- Troubleshooting stuck processes

---

### 🔧 **INSTALL_PERMANENT.bat** (Install Permanen)
**Icon:** 🔧 Wrench  
**Fungsi:** Register add-in agar auto-load  
**Kapan dipakai:** Sekali saja saat pertama setup

---

## 🔄 FLOW DIAGRAMS

### Flow 1: Daily Usage (Setiap Hari)
```
Start Day
    ↓
LAUNCH_ALL.bat ← KLIK DI SINI! 🚀
    ↓
Excel opens automatically
    ↓
Use Excel Agent normally
    ↓
Close Excel when done
    ↓
End Day ✅
```

---

### Flow 2: First Time Setup
```
First Installation
    ↓
RUN INSTALL_PERMANENT.bat (sekali aja)
    ↓
Follow Excel prompts
    ↓
THEN RUN LAUNCH_ALL.bat
    ↓
Taskpane loads in Excel
    ↓
Ready to use! ✅
```

---

### Flow 3: Troubleshooting Loop
```
Problem Detected?
    ↓
CHECK STATUS [Option 6 in Menu]
    ↓
Identify Issue:
├─ Server down? → Run START_SERVER.bat
├─ Cache problem? → Run CLEAR_CACHE.bat
└─ Stuck process? → Run STOP_ALL.bat
    ↓
Restart Fresh
    ↓
Run LAUNCH_ALL.bat again
    ↓
Test functionality
    ↓
Working? ✅
Not working? → Repeat troubleshooting
```

---

## 🎨 VISUAL CHEAT SHEET

| Icon | Script | Description | Use Case |
|------|--------|-------------|----------|
| 🚀 | LAUNCH_ALL | One-click complete start | **Daily use (recommended)** |
| 🎮 | MENU_CONTROL | Interactive panel | Advanced control |
| 🖥️ | START_SERVER | Server only | Background mode |
| 📁 | LOAD_ADDIN | Excel only | Launch workbook |
| 🗑️ | CLEAR_CACHE | Reset everything | Fix problems |
| ⏹️ | STOP_ALL | Shutdown | Clean exit |
| 🔧 | INSTALL_PERMANENT | Auto-load setup | Initial installation |

---

## 💡 PRO TIPS

### Tip 1: Desktop Shortcut ⭐
Right-click `LAUNCH_ALL.bat` → Send to → Desktop  
Now you have ONE CLICK launch from anywhere!

---

### Tip 2: Pin to Taskbar
Drag `LAUNCH_ALL.bat` to Windows taskbar  
→ Single click access with one touch!

---

### Tip 3: Keyboard Shortcut
With AutoHotkey, set `Ctrl+Shift+E` to launch:
```autohotkey
^!e::
Run, "C:\Users\acer\agent_excel\ExcelAgent\LAUNCH_ALL.bat"
return
```

---

## 🆘 EMERGENCY FIXES

### Scenario: Excel Agent Tidak Muncul
**Solution:**
1. Run `CLEAR_CACHE.bat`
2. Close all windows
3. Run `LAUNCH_ALL.bat`
4. In Excel → ADD-INS tab → Click "Excel Agent"

---

### Scenario: Error "Session Invalid"
**Solution:**
1. Wait 5-10 seconds (auto-recovery kicks in)
2. If still error → Run `MENU_CONTROL.bat` → Option [6]
3. Check if server running
4. If not → Run `START_SERVER.bat`

---

### Scenario: Slow Performance
**Solution:**
1. Run `STOP_ALL.bat`
2. Run `CLEAR_CACHE.bat`
3. Run `LAUNCH_ALL.bat` fresh
4. Should be faster now

---

## 📊 SCRIPT STATISTICS

| Script | Lines | Size | Complexity |
|--------|-------|------|------------|
| LAUNCH_ALL.bat | 50+ lines | Medium | Easy to use |
| MENU_CONTROL.bat | 100+ lines | Complex | Flexible |
| START_SERVER.bat | 30+ lines | Simple | Quick start |
| CLEAR_CACHE.bat | 60+ lines | Medium | Maintenance |
| STOP_ALL.bat | 40+ lines | Simple | Graceful exit |

---

## 🎯 RECOMMENDED WORKFLOW

### For Beginners:
```
ONLY USE LAUNCH_ALL.bat
Everything else handled automatically! ✅
```

### For Power Users:
```
MENU_CONTROL.bat for granular control
Customize based on specific needs
Schedule automated tasks
```

---

## 📞 SUPPORT CHAIN

Have issues with batch scripts?

1. Check this guide first ✅
2. Review BATCH_FILES_GUIDE.md for details
3. Check repository README.md
4. Open GitHub issue: https://github.com/Rizki2629/excel-agent-private/issues

---

## ✨ SUMMARY

You have **7 batch scripts** covering all scenarios:

**Most Important:**
🚀 **LAUNCH_ALL.bat** ← Your primary tool!

**Support Tools:**
🎮 MENU_CONTROL.bat (advanced control)  
🖥️ START_SERVER.bat (server only)  
📁 LOAD_ADDIN.bat (Excel launcher)  
🗑️ CLEAR_CACHE.bat (reset utility)  
⏹️ STOP_ALL.bat (shutdown)  
🔧 INSTALL_PERMANENT.bat (setup)  

**Just remember:** `LAUNCH_ALL.bat` does EVERYTHING! 🚀

---

**Version:** v1.0  
**Last Updated:** 2026-08-18  
**Status:** Production Ready ✅

<div align="center">

### 🎉 ONE CLICK TO START YOUR WORKDAY!

#### Double-click LAUNCH_ALL.bat
#### And you're ready to go! 🚀

</div>
