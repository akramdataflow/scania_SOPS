@echo off
setlocal
cd /d "%~dp0"

:: ???? ??? ????? ??? ?? ?????
set TARGET=app.py

echo ================================================================
echo  DEBUG MODE - Python launcher
echo  Target: %TARGET%
echo  Folder: %cd%
echo ================================================================

where py 2>nul
where python 2>nul
echo.

set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

echo Running: py -3 -X faulthandler -u "%TARGET%"
echo ------------------------------------------------
python -3 -X faulthandler -u "%TARGET%"
set ERR=%ERRORLEVEL%

echo ------------------------------------------------
echo Exit code: %ERR%
echo (Keep this window open to read errors)
pause
exit /b %ERR%
