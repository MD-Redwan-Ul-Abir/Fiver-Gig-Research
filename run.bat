@echo off
:: ─────────────────────────────────────────────────────────────────────────────
::  Fiverr Research Hub — Launcher  (Windows)
::  Double-click this file OR run from Command Prompt / PowerShell.
::  Requires: Docker Desktop for Windows (https://www.docker.com/products/docker-desktop)
:: ─────────────────────────────────────────────────────────────────────────────

echo.
echo   Fiverr Research Hub — Starting...
echo.

:: ── Pre-flight: create required folders and config files ─────────────────────
if not exist "Excel and Images\gig_images" (
    mkdir "Excel and Images\gig_images"
    echo   Created folder: Excel and Images\gig_images
)

if not exist api_config.json (
    echo {} > api_config.json
    echo   Created api_config.json  (add your API key inside the app^)
)

if not exist hub_config.json (
    echo {} > hub_config.json
    echo   Created hub_config.json
)

echo.
echo   Starting container...
echo   Once ready, open:  http://localhost:6080/vnc.html
echo.

docker-compose up
pause
