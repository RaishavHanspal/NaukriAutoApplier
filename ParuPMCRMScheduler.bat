@echo off
rem Scheduled run that applies ONLY to project management / project coordinator / CRM roles
rem (keywords and title filter in config_pm_crm.json). Started by the Windows Task Scheduler
rem task "Naukri PM-CRM applier"; can also be double-clicked. No pause, so the scheduled run can finish.
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo === Getting latest changes from GitHub ===
git pull --rebase --autostash -q
if errorlevel 1 echo WARNING: git pull failed - continuing with the local copy

echo.
echo === Applying to project management / coordinator / CRM jobs ===
python -u naukri_assistant.py run --config config_pm_crm.json
set RC=%ERRORLEVEL%
if not "%RC%"=="0" echo WARNING: the run ended with exit code %RC% - see run.log

echo.
echo === Saving results to GitHub ===
for /f "delims=" %%d in ('powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm'"') do set NOW=%%d
git add -A
git diff --cached --quiet
if errorlevel 1 git commit -q -m "PM-CRM run %NOW%"
git pull --rebase -X theirs -q
git push -q
if errorlevel 1 (echo ERROR: git push failed - results are committed locally and will be pushed next run.) else (echo Pushed to GitHub.)
exit /b %RC%
