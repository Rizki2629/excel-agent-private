# Excel Agent - Task Pane Version

AI-powered assistant for Excel with a modern task pane interface, similar to Claude.

## Features

- **Task Pane UI**: Appears on the right side of Excel (like Claude)
- **Dark Theme**: Modern dark interface
- **Chat Interface**: Type commands in Indonesian
- **Python COM Automation**: Uses local Python COM automation for Excel actions
- **Settings Panel**: Configure API key and model

## Quick Start

### Prerequisites

- Node.js 14+ installed
- Microsoft Excel 2016+ (Windows)
- OpenAI API key (or compatible API)

### Installation

1. Install dependencies:
```bash
cd ExcelAgent
npm install
```

2. Start the server:
```bash
npm start
```

3. Open Excel and load the add-in (see below)

### Loading the Add-in in Excel

#### Method 1: Developer Add-in (Recommended)

1. Open Excel
2. Go to **File > Options > Trust Center > Trust Center Settings**
3. Click **Trusted Add-in Catalogs**
4. Add this URL: `https://localhost:3000`
5. Click **Add catalog**
6. Go to **Insert > Office Add-ins > My Add-ins**
7. Click **Shared Folder** tab
8. Click **Add** next to Excel Agent

#### Method 2: Sideload (Development)

1. Install Office Add-in CLI:
```bash
npm install -g office-addin-cli
```

2. Sideload the add-in:
```bash
npm run sideload
```

## Usage

### Target Workbook

Task pane terikat ke workbook yang membukanya. Mengaktifkan workbook atau jendela Excel lain tidak mengubah target agent.

- Simpan workbook ke drive lokal atau path UNC/network sebelum membuka task pane. URL cloud `http://` atau `https://` belum didukung.
- Setelah workbook diganti nama atau dipindah, buka ulang task pane.
- Setiap workbook memiliki riwayat percakapan sendiri.
- Beberapa instance Excel didukung selama setiap target memiliki path penuh yang unik.

1. After loading, the Excel Agent task pane appears on the right
2. Click the settings icon (⚙️) to enter your API key
3. Type commands in Indonesian, such as:
   - "Buat chart penjualan bulanan"
   - "Format sel A1:F10 dengan warna biru"
   - "Ekspor data ke PDF"
4. Press Enter to send
5. The AI will respond with explanations and Python COM actions
6. Click "Terapkan ke Excel" to apply the actions

## Example Commands

| Command | Description |
|---------|-------------|
| Buat chart penjualan bulanan | Creates a monthly sales chart |
| Format sel A1:F10 dengan warna biru | Formats cells with blue color |
| Hitung rata-rata kolom B | Calculates average of column B |
| Ekspor data ke PDF | Exports data to PDF |
| Sorting data berdasarkan kolom A | Sorts data by column A |

## Configuration

### API Settings

Click the settings icon (⚙️) in the task pane to configure:

- **API Key**: Your OpenAI API key
- **API Endpoint**: API endpoint URL
- **Model**: GPT model to use (GPT-4, GPT-4 Turbo, GPT-3.5 Turbo)

API key disimpan oleh server lokal di profil pengguna dan tidak dikirim kembali ke task pane. Kolom API key kosong berarti key tersimpan tetap dipakai; isi kolom hanya saat ingin menggantinya. Cache `localStorage` lama dimigrasikan sekali lalu secret dihapus dari cache browser.

### Batas Keamanan Executor

Python hasil generasi berjalan dengan izin proses pengguna Windows. Binding `wb` mencegah perpindahan target yang tidak disengaja ketika fokus Excel berubah, tetapi bukan sandbox dan bukan isolasi keamanan terhadap kode berbahaya. Gunakan hanya endpoint/model yang dipercaya dan tinjau kode sebelum menjalankan saat `autoRun` dimatikan.

Penyimpanan riwayat menyerialkan perubahan dari satu task pane agar respons lama tidak menimpa snapshot baru. Dua task pane yang membuka workbook sama masih dapat melakukan overwrite snapshot terakhir; buka satu task pane per workbook untuk menghindari konflik tersebut.

### Supported APIs

- OpenAI (GPT-4, GPT-3.5)
- Azure OpenAI
- Any OpenAI-compatible API

## File Structure

```
ExcelAgent/
├── manifest.xml           # Office Add-in manifest
├── package.json           # Node.js dependencies
├── server.js              # HTTPS server for development
├── taskpane/
│   ├── taskpane.html      # Task pane UI
│   ├── taskpane.css       # Styling
│   └── taskpane.js        # JavaScript logic
└── scripts/
    └── install-addin.js   # Installation helper
```

## Troubleshooting

### "Add-in not showing"

1. Make sure server is running: `npm start`
2. Check if https://localhost:3000 is accessible
3. Restart Excel

### "Security warning"

1. Go to File > Options > Trust Center
2. Click Trust Center Settings
3. Go to Trust Add-in Catalogs
4. Check "Allow in-product catalog notifications"

### "API Error"

1. Check your API key in settings
2. Make sure you have credits in your API account
3. Check network connection

## Development

To modify the task pane:

1. Edit files in `taskpane/` directory
2. Refresh the task pane in Excel (right-click > Refresh)
3. Or restart Excel and reload the add-in

## Differences from VBA Version

| Feature | Task Pane | VBA Version |
|---------|-----------|-------------|
| UI Technology | HTML/CSS/JS | VBA UserForm |
| Appearance | Modern dark theme | Basic Windows form |
| Placement | Right side panel | Floating window |
| Installation | Requires server | Direct Excel file |
| Offline | No (needs server) | Yes |

## Support

For issues or questions, please check:
1. This README file
2. The main project README.md
3. Office Add-in documentation: https://learn.microsoft.com/en-us/office/dev/add-ins/
