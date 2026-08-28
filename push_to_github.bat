@echo off
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
echo ========================================================
echo   GitHub 저장소로 프로젝트 업로드 (git push)
echo ========================================================
echo.
git push -u origin main
echo.
pause
