@echo off
setlocal
cd /d %~dp0

for /f "tokens=1-3 delims=-" %%a in ("%date%") do set YMD=%%a%%b%%c

set DEST=%~dp0banyoung
set ZIP=%~dp0banyoung_%YMD%.zip

if exist "%DEST%" rd /s /q "%DEST%"
if exist "%ZIP%"  del /q "%ZIP%"
mkdir "%DEST%"

copy /Y "serve.py"          "%DEST%" > nul
copy /Y "reader_ssf.py"     "%DEST%" > nul
copy /Y "reader_sso.py"     "%DEST%" > nul
copy /Y "logger.py"         "%DEST%" > nul
copy /Y "common.py"         "%DEST%" > nul
copy /Y "db_manager.py"     "%DEST%" > nul
copy /Y "db_initializer.py" "%DEST%" > nul
copy /Y "view_log.bat"      "%DEST%" > nul

xcopy /Y /I /Q "templates" "%DEST%\templates"
xcopy /Y /I /Q "static"    "%DEST%\static"

powershell -NoProfile -Command "Compress-Archive -Path '%DEST%' -DestinationPath '%ZIP%'"

echo Done: %ZIP%
pause
