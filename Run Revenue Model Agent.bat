@echo off
REM Revenue Model Monthly Rollover Agent
REM Double-click to run with defaults, or call from cmd with flags:
REM   "Run Revenue Model Agent.bat" --only-step 5
REM   "Run Revenue Model Agent.bat" --skip-step1
REM   "Run Revenue Model Agent.bat" --file "Revenue Model - 04212026 (Internal).xlsm"

cd /d "%~dp0"

REM --- Preflight: config.json must exist and be edited for this machine ---
if not exist "config.json" (
    echo.
    echo ERROR: config.json is missing.
    echo Copy config.example.json to config.json and edit the paths.
    echo See SETUP.md for details.
    echo.
    pause
    exit /b 1
)

REM --- Preflight: Python launcher must be on PATH ---
where py >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: 'py' launcher not found on PATH.
    echo Install Python from python.org and tick "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

echo ===============================================
echo  Revenue Model Rollover Agent
echo  Working dir: %CD%
echo ===============================================
echo.

py -m agent.main %*
set EXITCODE=%ERRORLEVEL%

echo.
echo ===============================================
if %EXITCODE% EQU 0 (
    echo  Agent finished successfully [exit 0]
) else (
    echo  Agent FAILED [exit %EXITCODE%] - check newest log in agent\logs\
)
echo ===============================================
echo.
pause
exit /b %EXITCODE%
