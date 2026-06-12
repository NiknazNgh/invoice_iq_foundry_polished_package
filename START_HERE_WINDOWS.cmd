@echo off
setlocal
cd /d "%~dp0"
title Invoice IQ Agent - Foundry Polished Demo

echo ============================================================
echo Invoice IQ Agent - Foundry Polished Hackathon Demo
echo Installs packages, creates sample invoices, and opens Streamlit.
echo ============================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PYTHON_CMD=py
) else (
    set PYTHON_CMD=python
)

echo Using Python command: %PYTHON_CMD%
echo.

if not exist .env (
    echo .env not found. Creating .env from .env.example...
    copy .env.example .env >nul
)

%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: Package install failed.
    pause
    exit /b 1
)

echo.
echo Creating / refreshing 15 synthetic invoices...
%PYTHON_CMD% make_sample_invoices.py
if errorlevel 1 (
    echo.
    echo ERROR: Sample invoice generation failed.
    pause
    exit /b 1
)

echo.
echo IMPORTANT: Make sure you already ran: az login --use-device-code
echo.
echo Starting Streamlit Foundry app...
echo When the browser opens, upload PDFs from the sample_invoices folder.
echo.
%PYTHON_CMD% -m streamlit run invoice_iq_foundry_demo.py
pause
