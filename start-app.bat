@echo off
REM ============================================================
REM  RADAR-GESTOR - arrencada robusta del servidor local
REM  Doble clic per treballar a http://localhost:5173/
REM ============================================================
setlocal EnableDelayedExpansion
title Radar-Gestor (local)
cd /d "%~dp0"

echo.
echo ================================================
echo   RADAR-GESTOR - Servidor local
echo ================================================
echo.

REM --- 1) Comprovar que node/npm existeixen --------------------
where npm >nul 2>&1
if errorlevel 1 (
  echo [ERROR] No s'ha trobat npm/Node.js al sistema.
  echo Instal.la Node.js des de https://nodejs.org i torna-ho a provar.
  echo.
  pause
  exit /b 1
)

REM --- 2) Alliberar el port 5173 si esta ocupat ----------------
REM  Mata qualsevol proces que estigui escoltant/usant el 5173
REM  (evita el cas d'un servidor orfe penjat que bloqueja l'app).
echo Comprovant si el port 5173 esta lliure...
set "KILLED="
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5173"') do (
  if not "%%P"=="0" (
    REM Nomes matem si el proces es node.exe (mai Firefox o altres apps)
    for /f "delims=" %%N in ('tasklist /FI "PID eq %%P" /NH 2^>nul ^| findstr /I "node.exe"') do (
      taskkill /PID %%P /F >nul 2>&1
      echo  - Tancat servidor node orfe al port 5173 ^(PID %%P^)
      set "KILLED=1"
    )
  )
)
if not defined KILLED echo  - Port 5173 lliure.
REM Petita pausa perque el SO alliberi el socket
ping -n 2 127.0.0.1 >nul

REM --- 3) Comprovar dependencies -------------------------------
if not exist "node_modules" (
  echo.
  echo Primer arrencada: instal.lant dependencies, pot trigar uns minuts...
  call npm install
  if errorlevel 1 (
    echo [ERROR] La instal.lacio de dependencies ha fallat.
    pause
    exit /b 1
  )
)

REM --- 4) Arrencar el servidor --------------------------------
echo.
echo ================================================
echo   URL: http://localhost:5173/
echo   Manten aquesta finestra oberta mentre treballes.
echo   Per aturar: tanca aquesta finestra o prem Ctrl+C.
echo ================================================
echo.

call npm run dev

REM --- 5) Si arriba aqui, el servidor s'ha aturat -------------
echo.
echo ================================================
echo  El servidor s'ha aturat.
echo  Si ha estat per un error, llegeix el missatge de sobre.
echo  Pots tornar a obrir l'app fent doble clic a aquesta drecera.
echo ================================================
echo.
pause >nul
