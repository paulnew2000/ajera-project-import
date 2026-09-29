@echo off
setlocal
title Ajera Project Import
cd /d "%~dp0"

where python >nul 2>nul || (
  echo Python is not installed. Install it from https://www.python.org/downloads/ ^(tick "Add python.exe to PATH"^) and run this again.
  pause & exit /b 1
)
python -c "import httpx, openpyxl, dotenv, mcp" 2>nul || (
  echo First run: installing the few Python packages this tool needs...
  python -m pip install --user -r requirements.txt || (pause & exit /b 1)
)
if not exist ".env" (
  copy ".env.example" ".env" >nul
  echo.
  echo A settings file ^(.env^) was created. Notepad will open it now:
  echo fill in your Ajera API URL, user name and password, save, close Notepad.
  pause
  notepad ".env"
)

:menu
cls
echo  ==============================================
echo            AJERA PROJECT IMPORT
echo  ==============================================
echo   1  Check connection  (which database am I on?)
echo   2  Make a blank template from my Ajera
echo   3  CHECK a workbook   (safe - changes nothing)
echo   4  CREATE projects from a workbook
echo   5  Show projects this tool has created
echo   6  Open the settings file (.env)
echo   Q  Quit
echo.
set "choice="
set /p choice="  Choose: "
if /i "%choice%"=="1" python cli.py connect & pause & goto menu
if /i "%choice%"=="2" python cli.py template & pause & goto menu
if /i "%choice%"=="3" set "action=check" & goto withfile
if /i "%choice%"=="4" set "action=create" & goto withfile
if /i "%choice%"=="5" python cli.py log & pause & goto menu
if /i "%choice%"=="6" notepad ".env" & goto menu
if /i "%choice%"=="q" exit /b 0
goto menu

:withfile
call :pick
if defined file python cli.py %action% "%file%"
pause
goto menu

:pick
set "file="
for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Excel workbook (*.xlsx)|*.xlsx'; $d.InitialDirectory = (Get-Location).Path; if ($d.ShowDialog() -eq 'OK') { $d.FileName }"`) do set "file=%%F"
if not defined file echo No file chosen.
exit /b 0
