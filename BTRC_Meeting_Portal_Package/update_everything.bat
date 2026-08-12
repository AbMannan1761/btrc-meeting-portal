@echo off
chcp 65001 > nul
echo ========================================================
echo   BTRC Meeting Database - Google Sheets Auto-Sync
echo ========================================================
echo.
echo 🌐 Fetching live data from your Google Sheet...
py sync_from_google_sheet.py

echo.
echo 🚀 Uploading updates to GitHub Live Web Portal...
git add index.html
git commit -m "Auto-synced latest meeting data from Google Sheet"
git push origin main

echo.
echo ========================================================
echo ✅ DONE! Live Web Portal, Excel, and Word files are synced!
echo ========================================================
pause
