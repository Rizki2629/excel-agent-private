# 📦 Deployment Guide - Excel Agent Office Add-in

## 🚀 Quick Deployment Checklist

### 1️⃣ Prerequisites Installation
```powershell
# Check Node.js version (must be >= 18)
npm --version

# Check Python version (must be >= 3.12)
python --version

# Install OpenCode CLI globally
npm install -g opencode@latest
opencode --version
```

**Expected Output:**
- Node: `v18.x.x` or higher ✅
- Python: `Python 3.12.x` or higher ✅
- OpenCode: `v1.18.18+` ✅

---

### 2️⃣ Clone Repository
```bash
# Option A: HTTPS
git clone https://github.com/Rizki2629/excel-agent-private.git
cd excel-agent-private

# Option B: SSH (if configured)
git clone git@github.com:Rizki2629/excel-agent-private.git
cd excel-agent-private
```

---

### 3️⃣ Configure LLM Gateway (Port 20128)

#### Step A: Ensure OpenCode gateway is running
```powershell
# Check if gateway running on port 20128
netstat -ano | findstr :20128

# If not running, start it according to your provider's instructions
# Common providers:
# - Ollama (local): ollama serve
# - llama.cpp: server.exe
# - Cloud provider: Follow provider docs
```

#### Step B: Test connection
```powershell
curl http://localhost:20128/v1/models
```

**Expected Response:** JSON array of available models

---

### 4️⃣ Configure Excel Agent Settings

#### Create config directory
```powershell
$ConfigPath = "$env:APPDATA\ExcelAgent"
New-Item -ItemType Directory -Path $ConfigPath -Force | Out-Null
```

#### Create/edit config.json
```json
{
  "useOpencode": true,
  "opencodeModel": "9router/qd/ultimate"
}
```

**Important:** DO NOT use `ag/gemini-3.7-flash-high` - it returns errors!

---

### 5️⃣ Install Dependencies
```powershell
cd C:\path\to\excel-agent-private\npm install
pip install -r requirements.txt  # If Python dependencies exist
```

---

### 6️⃣ Start Server
```powershell
node server.js
```

**Expected Output:**
```
Server listening on port 3000
```

---

### 7️⃣ Load Add-in in Excel

#### Manual Loading (Development Mode)
1. Open Excel → Developer tab → Insert → Get Add-ins
2. Click **"Store"** → **"Add from Manifest"**
3. Browse to: `C:\path\to\excel-agent-private\manifest.xml`
4. Click **Install**

#### Production Deployment
For enterprise environments, see **Enterprise Deployment** section below.

---

### 8️⃣ Clear Cache (First Time Only)
```powershell
# Stop ALL Excel processes
taskkill /F /IM EXCEL.EXE

# Clear Office WEF cache
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force

# Restart Excel
excel.exe
```

---

## 🧪 Post-Deployment Verification

### Test 1: Server Connection
Open browser: `https://localhost:3000/taskpane.html`

✅ Should load taskpane UI  
❌ Error? Check firewall/port bindings

---

### Test 2: Auto-Recovery Binding
1. Type any message in chat panel
2. Wait up to 10 seconds
3. Check notification badge

✅ Blue badge appears  
❌ Stuck loading? Clear cache + restart

---

### Test 3: Image Upload
1. Select 1-5 images (PNG/JPG/WEBP, ≤5MB each)
2. Click **"Upload & Analyze"**
3. Wait ~12-15 seconds

✅ Model reads text from image  
❌ Files rejected? Check file types/sizes

---

### Test 4: Shape Management
Send command:  
> "Kelola tombol dan shape di sheet DAFTAR HADIR"

✅ 6 modern buttons created  
❌ Button missing? Check COM automation permissions

---

## 🔐 Security Configuration

### Private Repository Access
The repository is set to **private**. Team members need:

1. **GitHub Account** with organization access
2. **Read access** via team member invitation

### Environment Variables (Optional)
Create `.env.local` for sensitive data:
```bash
OPENCODE_TOKEN=your_token_here
EXCEL_PASSWORD=encrypted_password
```

---

## 🛠 Troubleshooting

### Issue: OpenCode Hangs Forever
**Symptom:** Command never completes, 100% CPU usage

**Solution:**
1. Verify `--auto` flag in server.js spawn args
2. Check scratch workdir exists: `$env:TEMP\excelagent-opencode`
3. Kill stuck process: `taskkill /F /PID <PID>`

---

### Issue: Taskpane Shows Empty Chat
**Symptom:** Welcome screen visible, no messages sent

**Solutions:**
1. Check console (F12): `Error: xhrPost failed` → Network issue
2. Reload page: `Alt+F5`
3. Clear cache + restart (see Step 8 above)

---

### Issue: Conditional Formatting Not Working
**Symptom:** M13 cell shows error instead of colored background

**Solution:**
Verify exact-match formulas (not SEARCH/COUNTIF):
```python
rng.FormatConditions.Add(2, 1, "=M13=\"AMAN\"", pythoncom.Missing)
```

Also ensure:
- Sheet activated first: `wb.Sheets('CONTROL PANEL').Activate()`
- Formula uses double quotes escaped

---

## 📊 Monitoring & Health Checks

### Server Logs
```powershell
Get-Content logs/server.log -Wait -Tail 50
```

Key indicators:
- `[INFO] Server listening on port 3000` ✅
- `[ERROR] OpenCode timeout after X seconds` ❌
- `[WARN] Process killed by parent` ⚠️

---

### Performance Metrics
- **Taskpane Load Time**: Target < 2 seconds
- **Auto-Recovery Time**: Target 5-10 seconds
- **Image Processing Time**: Target 12-15 seconds
- **Shape Generation Time**: Target 3-5 seconds

---

## 🔄 Updates & Maintenance

### Update to Latest Version
```bash
git pull origin master
npm install --only=prod  # Reinstall dependencies
```

### Database Reset
```powershell
# Clear all session data
Get-ChildItem "$env:TEMP\excelagent-opencode\" -Recurse | Remove-Item -Force -Recurse
```

---

## 🏢 Enterprise Deployment Options

### Option A: Office Store Publishing
1. Package as `.appx` using Microsoft packaging tools
2. Submit to Office Store (requires app registration)
3. Users install via Store URL

### Option B: Centralized Management
1. Deploy manifest via GPO (Group Policy)
2. Pre-configured with org-wide settings
3. Silent installation script provided

### Option C: VSTO Solution
1. Compile to .NET assembly
2. MSI installer creation
3. System-wide deployment via SCCM/Intune

---

## 📞 Support & Contact

### For Issues
1. Check this deployment guide
2. Review `/docs/superpowers/specs/` for technical details
3. Run tests: `npm test`

### Escalation Path
1. Local admin → Review logs
2. IT Support → Network/firewall check
3. Development → GitHub Issues

---

## 🎓 Training Resources

### Beginner Course (30 min)
- Understanding AI vs traditional macros
- Basic chat commands
- How to upload images for analysis

### Advanced Course (2 hours)
- Creating custom Python scripts
- Extending agent capabilities
- Building reusable button templates

### Admin Course (1 day)
- Deployment planning
- Security configuration
- Monitoring & troubleshooting

---

**Last Updated**: 2026-08-18  
**Version**: v1.5.2  
**Document Maintainer**: Your DevOps Team
