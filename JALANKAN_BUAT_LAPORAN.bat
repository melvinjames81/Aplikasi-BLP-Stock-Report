@echo off
setlocal
cd /d "%~dp0"

:: Cek Python terinstall
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan. Silakan install Python terlebih dahulu.
    pause
    exit /b 1
)

echo =============================================
echo  PEMBUAT FILE LAPORAN STOCK BULAN BARU
echo =============================================
echo.
set /p sumber=Masukkan nama/path file sumber: 
set /p output=Masukkan nama/path file output (boleh dikosongkan): 
echo.
set /p odoo_masuk=Masukkan path file Odoo Stock Masuk (boleh dikosongkan): 
set /p odoo_sales=Masukkan path file Odoo Sales (boleh dikosongkan): 
set /p commercial=Masukkan path file Commercial/Sales (boleh dikosongkan): 
set /p comm_sheet=Masukkan nama sheet Commercial (boleh dikosongkan): 

echo.

:: Build command
set CMD=python buat_laporan_stock_baru.py "%sumber%"

if not "%output%"=="" set CMD=%CMD% "%output%"
if "%output%"=="" set CMD=%CMD% ""

if not "%odoo_masuk%"=="" set CMD=%CMD% --odoo-masuk "%odoo_masuk%"
if not "%odoo_sales%"=="" set CMD=%CMD% --odoo-sales "%odoo_sales%"
if not "%commercial%"=="" set CMD=%CMD% --commercial "%commercial%"
if not "%comm_sheet%"=="" set CMD=%CMD% --commercial-sheet "%comm_sheet%"

%CMD%

echo.
pause
