@echo off
setlocal
cd /d "%~dp0"

echo =============================================
echo  BUILD INSTALLER BLP STOCK REPORT
echo =============================================
echo.

:: ── Step 1: Build EXE dengan PyInstaller ──
echo [STEP 1/2] Building executable dengan PyInstaller...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan!
    pause
    exit /b 1
)

pip install --upgrade pyinstaller >nul 2>&1

pyinstaller --noconfirm --onedir --windowed ^
    --name "BLPStockReport" ^
    --add-data "buat_laporan_stock_baru.py;." ^
    buat_laporan_stock_gui.py

if not exist "dist\BLPStockReport\BLPStockReport.exe" (
    echo.
    echo [ERROR] PyInstaller build gagal!
    pause
    exit /b 1
)

echo.
echo [OK] EXE berhasil dibuat: dist\BLPStockReport\BLPStockReport.exe
echo.

:: ── Step 2: Compile Installer dengan Inno Setup ──
echo [STEP 2/2] Compiling installer dengan Inno Setup...
echo.

set ISCC=""

:: Cari Inno Setup di lokasi umum
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
)
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
)

if %ISCC%=="" (
    echo [WARN] Inno Setup tidak ditemukan secara otomatis.
    echo        Silakan buka installer.iss di Inno Setup Compiler secara manual.
    echo        Atau install Inno Setup 6 dari: https://jrsoftware.org/isdl.php
    echo.
    echo        File EXE sudah siap di: dist\BLPStockReport\
    pause
    exit /b 0
)

%ISCC% installer.iss

if exist "Output\BLPStockReport_Setup.exe" (
    echo.
    echo =============================================
    echo  INSTALLER BERHASIL DIBUAT!
    echo  File: Output\BLPStockReport_Setup.exe
    echo =============================================
) else (
    echo.
    echo [ERROR] Inno Setup compile gagal!
)

echo.
pause
