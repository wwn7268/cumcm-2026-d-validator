@echo off
cd /d "%~dp0"
where python >nul 2>nul
if not errorlevel 1 (
    python app.py
) else (
    py -3 app.py
)
if errorlevel 1 (
    echo.
    echo Application could not start. See usage instructions for Python dependencies.
    pause
)
