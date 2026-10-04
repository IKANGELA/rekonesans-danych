@echo off
REM Podglad aplikacji na tym komputerze: http://127.0.0.1:8765
REM Zamknij to okno, aby wylaczyc serwer.
cd /d "%~dp0"
start "" http://127.0.0.1:8765
node narzedzia\serwer.mjs . 8765
pause
