@echo off
echo Starting Livedoc Variables POC server (Python HTTP)...
echo.
echo Server: http://localhost:3000
echo.
echo Sideload steps:
echo   1. Open PowerPoint
echo   2. Insert ^> Get Add-ins ^> My Add-ins ^> Upload My Add-in
echo   3. Browse to: %~dp0manifest.xml
echo   4. Click "Variables" in the Home tab ribbon
echo.
echo Press Ctrl+C to stop.
echo.
cd /d "%~dp0"
python -m http.server 3000
pause
