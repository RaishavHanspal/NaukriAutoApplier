@echo off
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo === Getting latest changes from GitHub ===
git pull --rebase --autostash -q
if errorlevel 1 echo WARNING: git pull failed - continuing with the local copy

echo.
echo === Applying to jobs (max 10 per profile, max 45 minutes) ===
python -u naukri_assistant.py run
set RC=%ERRORLEVEL%
if not "%RC%"=="0" echo WARNING: the run ended with exit code %RC% - see run.log

echo.
echo === Saving results to GitHub ===
for /f "delims=" %%d in ('powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm'"') do set NOW=%%d
git add -A
git diff --cached --quiet
if errorlevel 1 git commit -q -m "Local run %NOW%"
rem local data wins over anything the GitHub workflow may have changed meanwhile
git pull --rebase -X theirs -q
git push -q
if errorlevel 1 (
    echo ERROR: git push failed - results are committed locally and will be pushed next run.
) else (
    echo Pushed to GitHub.
)

echo.
pause
