@echo off
cd /d "%~dp0"

echo.
echo  Hotel Sahara — Pushing to GitHub...
echo  =====================================

git add .

:: Auto-generate commit message with date and time
set TIMESTAMP=%DATE:~-4%-%DATE:~3,2%-%DATE:~0,2% %TIME:~0,8%
git commit -m "Update: %TIMESTAMP%"

git push

echo.
echo  ✅ Done! Changes are live on GitHub.
echo  Render will redeploy automatically in ~30 seconds.
echo.
pause
