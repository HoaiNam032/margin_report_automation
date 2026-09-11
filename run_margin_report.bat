@echo off

REM --- Đổi encoding console & Python sang UTF-8 ---
set PYTHONIOENCODING=utf-8
chcp 65001 >NUL 

REM ==== Lấy thư mục chứa file .bat (vd: D:\report_excel_python_1\) ====
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM ==== Tạo folder logs cạnh script ====
set "LOGDIR=%SCRIPT_DIR%logs"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"

REM ==== Tạo tên file log theo ngày ====
set TODAY=%DATE:~6,4%-%DATE:~3,2%-%DATE:~0,2%
set "LOGFILE=%LOGDIR%\margin_report_%TODAY%.log"

echo =============================== >> "%LOGFILE%"
echo Run at %DATE% %TIME% >> "%LOGFILE%"
echo Current directory: %CD% >> "%LOGFILE%"

REM ==== ĐƯỜNG DẪN PYTHON TRONG VENV (CHỈNH Ở DÒNG NÀY NẾU CẦN) ====
set "PYTHON_EXE=D:\BreakevenPriceProject\venv\Scripts\python.exe"

IF NOT EXIST "%PYTHON_EXE%" (
    echo Python venv not found at "%PYTHON_EXE%" >> "%LOGFILE%"
    echo Python venv not found at "%PYTHON_EXE%"
    exit /b 1
)

REM ==== Chạy script build_margin_report.py ====
"%PYTHON_EXE%" "%SCRIPT_DIR%build_margin_report.py" >> "%LOGFILE%" 2>&1