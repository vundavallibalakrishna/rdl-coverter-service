@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is not available on PATH. 1>&2
  exit /b 1
)

rem scripts\start-production.js loads RDL_ENV_FILE (default: .env.production), applies production
rem defaults for values that remain unset, and then starts the service. Telemetry remains on the console.
node scripts\start-production.js
set "RDL_EXIT_CODE=%ERRORLEVEL%"
endlocal & exit /b %RDL_EXIT_CODE%
