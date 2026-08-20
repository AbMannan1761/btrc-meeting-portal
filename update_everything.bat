@echo off
chcp 65001 > nul
echo ========================================================
echo   BTRC Meeting Database - Google Sheets Auto-Sync
echo ========================================================
echo.
echo 🌐 Fetching live data from your Google Sheet...

where python >nul 2>nul
if %errorlevel% equ 0 (
    python sync_from_google_sheet.py
    goto sync_done
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    py sync_from_google_sheet.py
    goto sync_done
)

if exist "%LOCALAPPDATA%\Python\bin\python.exe" (
    "%LOCALAPPDATA%\Python\bin\python.exe" sync_from_google_sheet.py
    goto sync_done
)

echo [ERROR] Python not found. Please ensure Python is installed and in PATH.
pause
exit /b 1

:sync_done
echo.
echo 🚀 Uploading updates to GitHub Live Web Portal...
git add index.html meetings_db.json extracted_agendas_live.xlsx "cm word copy" BTRC_Meeting_Portal_Package/ .github/
git commit -m "Auto-synced latest meeting data from Google Sheet"
git push origin main

echo.
echo ========================================================
echo ✅ DONE! Live Web Portal, Excel, and Word files are synced!
echo 🌐 Website: https://abmannan1761.github.io/btrc-meeting-portal/
echo ========================================================
pause
