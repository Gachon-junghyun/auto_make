@echo off
set MYDIR=%~dp0
set PYDIR=%MYDIR%_python
set MARKER=%PYDIR%\setup_done.txt

echo.
echo  Susanna Design - Auto Generator
echo.

:: 1. Local Python already installed?
if exist "%PYDIR%\python.exe" goto check_pkg

:: 2. System Python with tkinter?
py -c "import tkinter" >nul 2>&1
if not errorlevel 1 (
    set PYTHON=py
    goto check_pkg
)

python -c "import tkinter" >nul 2>&1
if not errorlevel 1 (
    set PYTHON=python
    goto check_pkg
)

:: 3. Download full Python installer (~25MB, includes tkinter)
echo Python not found. Downloading Python installer (~25MB)...
echo This is a ONE-TIME setup. Please wait.
echo.

powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile '%MYDIR%py_setup.exe' -UseBasicParsing }"

if not exist "%MYDIR%py_setup.exe" goto dl_fail

echo Installing Python to local folder (takes ~1 min)...
"%MYDIR%py_setup.exe" /quiet InstallAllUsers=0 TargetDir="%PYDIR%" PrependPath=0 Include_test=0 Include_doc=0
del "%MYDIR%py_setup.exe"

if not exist "%PYDIR%\python.exe" (
    echo [ERROR] Python install failed.
    pause
    exit /b 1
)

set PYTHON=%PYDIR%\python.exe
goto check_pkg

:dl_fail
echo [ERROR] Download failed. Check internet connection.
pause
exit /b 1

:: PYTHON 경로 미설정시 fallback
:check_pkg
if not defined PYTHON set PYTHON=%PYDIR%\python.exe

:: 패키지 설치 여부 확인
%PYTHON% -c "import openpyxl, pptx, win32com" >nul 2>&1
if not errorlevel 1 goto run

echo [1/2] Installing packages...
%PYTHON% -m pip install openpyxl Pillow python-pptx pywin32 --quiet --no-warn-script-location
if errorlevel 1 goto pkg_fail
echo done > "%MARKER%"

:run
echo [2/2] Starting program...
echo.
%PYTHON% "%MYDIR%geonjeokseo_jadonghua.py"
if errorlevel 1 goto crash
goto end

:pkg_fail
echo [ERROR] Package install failed. Check internet.
pause
exit /b 1

:crash
echo.
echo [ERROR] Program crashed. Contact developer.
pause

:end
