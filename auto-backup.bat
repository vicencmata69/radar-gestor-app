@echo off
REM Backup automàtic del Radar-Gestor
REM Doble clic per executar manualment, o programa amb Task Scheduler de Windows.

cd /d "%~dp0"
node scripts\backup.js
if errorlevel 1 (
  echo.
  echo ERROR: el backup ha fallat. Comprova .env i la connexio.
  pause
)
