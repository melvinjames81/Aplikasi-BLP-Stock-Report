# BLP Stock Report Generator

Aplikasi desktop untuk membuat laporan stock baru dengan auto-matching file dan FEFO sales processing.

## 🚀 Fitur Utama

- **Auto-match Files**: Otomatis mencocokkan file input ke sheet yang sesuai
- **FEFO Processing**: Proses First-Expire-First-Out untuk stock management
- **Multi-Sheet Support**: Proses multiple sheets sekaligus
- **Stock Masuk & Odoo**: Integrasi data dari Stock Masuk dan Odoo
- **Formula Auto-fill**: Automatic formula calculation untuk setiap row
- **Responsive UI**: Interface yang user-friendly dengan dark theme

## 📋 Requirements

```
Python 3.8+
openpyxl >= 3.0.0
tkinter (included dengan Python)
```

## ⚙️ Instalasi

### 1. Install Dependencies
```bash
pip install openpyxl
```

### 2. Jalankan Aplikasi
```bash
python buat_laporan_stock_gui.py
```

Atau gunakan batch file:
```bash
JALANKAN_BUAT_LAPORAN.bat
```

## 📖 Cara Penggunaan

### Step 1: Pilih File Sumber
- Klik tombol 📂 pada "File Sumber (.xlsx)"
- Pilih file Excel yang berisi template stock report

### Step 2: Load Sheets
- Klik tombol "📋 Load Sheets"
- Aplikasi akan menampilkan semua sheet dalam file

### Step 3: Pilih File Input (Opsional)
- Untuk setiap sheet, bisa memilih:
  - **SM (Stock Masuk)**: File input untuk stock masuk
  - **Odoo**: File export dari Odoo
- Jika kosong, sheet akan tetap diproses dengan data existing

### Step 4: Pilih Sheet untuk Diproses
- Centang sheet yang ingin diproses
- Jika tidak ada yang dicentang, semua sheet akan diproses

### Step 5: Set File Output (Opsional)
- Klik 📂 pada "File Output (opsional)"
- Jika kosong, file akan di-save dengan nama otomatis

### Step 6: Proses
- Klik tombol "🚀 PROSES LAPORAN"
- Monitor progress di log output

## 🎨 UI Features

- **Dark Theme**: Tema gelap untuk mata yang nyaman
- **Real-time Log**: Lihat progress proses secara real-time
- **Responsive Design**: Responsive terhadap resize window
- **Status Bar**: Progress indicator selama proses

## 🔧 Build Executable

### Generate .exe (PyInstaller)
```bash
build_exe.bat
```

### Generate Installer (Inno Setup)
```bash
build_installer.bat
```

## 📁 File Structure

```
Aplikasi-BLP-Stock-Report/
├── buat_laporan_stock_gui.py      # GUI Main App (RECOMMENDED)
├── buat_laporan_stock_baru.py     # CLI Version
├── BLPStockReport.spec            # PyInstaller Spec
├── installer.iss                  # Inno Setup Config
├── build_exe.bat                  # Build executable
├── build_installer.bat            # Build installer
└── JALANKAN_BUAT_LAPORAN.bat     # Run script
```

## 🐛 Troubleshooting

### Window tidak muncul
- Pastikan Python path sudah correct
- Coba run dari terminal: `python buat_laporan_stock_gui.py`

### File tidak ditemukan
- Pastikan path file benar
- Jangan gunakan special characters di filename

### Proses error
- Pastikan format Excel sesuai template
- Cek column headers sudah benar
- Lihat log output untuk detail error

## 📝 Notes

- Backup file Anda sebelum proses
- Support file format: .xlsx, .xls, .csv
- Process berjalan di background thread (UI tidak freeze)

## 👨‍💻 Developer

Melvin James - Stock Report Automation

## 📄 License

Open Source
