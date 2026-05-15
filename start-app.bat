@echo off
REM Arrenca l'aplicació Radar-Gestor en local
REM Doble clic per obrir el navegador a http://localhost:5173/
REM Per aturar el servidor: prem Ctrl+C en aquesta finestra (o tanca-la)

title Radar-Gestor (local)
cd /d "%~dp0"
echo.
echo ================================================
echo   RADAR-GESTOR — Servidor de desenvolupament
echo ================================================
echo.
echo   URL: http://localhost:5173/
echo.
echo   El navegador s'obrira automaticament.
echo   Mantingues aquesta finestra oberta mentre fas servir l'app.
echo   Per aturar: prem Ctrl+C o tanca aquesta finestra.
echo.
echo ================================================
echo.

npm run dev

REM Si npm peta, mantenim la finestra oberta perque vegis l'error
echo.
echo El servidor s'ha aturat. Prem qualsevol tecla per tancar...
pause >nul
