@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
python -u "%SCRIPT_DIR%voice_auto_vosk.py" %*
endlocal
