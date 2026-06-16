@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

REM === Thu muc du an: keo-tha thu muc vao file .bat, hoac de mac dinh la thu muc chua file .bat ===
set "TARGET=%~1"
if "%TARGET%"=="" set "TARGET=%~dp0"
REM Bo dau gach cheo nguoc cuoi cung (tranh loi escape dau ngoac kep tren Windows)
if defined TARGET if "%TARGET:~-1%"=="\" set "TARGET=%TARGET:~0,-1%"

echo ============================================================
echo   TONG HOP BAO GIA - Cubic Architects
echo ============================================================
echo Thu muc dang xu ly: %TARGET%
echo.

REM === Tim Python ===
set "PY="
where py >nul 2>&1 && set "PY=py"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )
if not defined PY (
  echo [LOI] Chua cai dat Python tren may nay.
  echo Vui long cai Python 3 tai: https://www.python.org/downloads/
  echo Khi cai nho tick chon "Add Python to PATH".
  echo.
  pause
  exit /b 1
)

REM === Cai thu vien can thiet (chi lan dau, can mang) ===
%PY% -m pip install --quiet --disable-pip-version-check openpyxl >nul 2>&1

REM === Chay chuong trinh ===
%PY% "%~dp0bao_gia_index.py" "%TARGET%" "%~dp0Tong_hop_bao_gia.xlsx"
if errorlevel 1 (
  echo.
  echo [LOI] Co loi khi chay. Vui long chup man hinh va gui lai.
  pause
  exit /b 1
)

echo.
echo === XONG! Dang mo file ket qua: Tong_hop_bao_gia.xlsx ===
start "" "%~dp0Tong_hop_bao_gia.xlsx"
endlocal
