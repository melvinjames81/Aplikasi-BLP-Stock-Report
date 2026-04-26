@echo off
setlocal
cd /d "%~dp0"

echo =============================================
echo  BUILD BLP STOCK REPORT (.exe)
echo =============================================
echo.

:: Cek Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan!
    pause
    exit /b 1
)

:: Install/upgrade PyInstaller
echo [1/2] Menginstall PyInstaller...
pip install --upgrade pyinstaller
echo.

:: Build exe
echo [2/2] Building executable...
pyinstaller --noconfirm --onedir --windowed ^
    --name "BLPStockReport" ^
    --add-data "buat_laporan_stock_baru.py;." ^
    buat_laporan_stock_gui.py

echo.
if exist "dist\BLPStockReport\BLPStockReport.exe" (
    echo =============================================
    echo  BUILD BERHASIL!
    echo  Output: dist\BLPStockReport\BLPStockReport.exe
    echo =============================================
) else (
    echo [ERROR] Build gagal! Periksa error di atas.
)

echo.
pause
