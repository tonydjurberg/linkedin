@echo off
setlocal
cd /d "%~dp0"

if exist "ProspectHunter.exe" (
    start "" /b "ProspectHunter.exe"
    exit /b 0
)

echo ProspectHunter.exe was not found.
echo.
echo If this is the source package, build the EXE first.
echo The GitHub Actions build creates ProspectHunter.exe.
pause
