@echo off
setlocal enabledelayedexpansion
title DocTranslate — Word & PDF Translator

echo.
echo  =====================================================
echo    DocTranslate — Word ^& PDF Translator
echo  =====================================================
echo.
echo  Architecture:
echo    C# WPF app  — GUI, DOCX (Word COM), PDF (PdfSharp)
echo    Python      — Argos Translate (offline, CPU-optimised)
echo.

:: ── Check .NET 8 ──────────────────────────────────────────────────────────────
where dotnet >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] .NET 8 SDK not found.
    echo  Download from: https://dotnet.microsoft.com/download/dotnet/8.0
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('dotnet --version 2^>^&1') do set DOTNET_VER=%%v
echo  Found: .NET %DOTNET_VER%

:: ── Check Python ──────────────────────────────────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found.
    echo  Install Python 3.10+ from: https://www.python.org/downloads/
    echo  Tick "Add Python to PATH" during install.
    pause & exit /b 1
)
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PY_VER=%%v
echo  Found: Python %PY_VER%

:: ── Install Python translation deps ──────────────────────────────────────────
echo.
echo  Installing Python translation engine (Argos Translate)...
python -m pip install -r requirements_translator.txt --quiet --disable-pip-version-check
if errorlevel 1 (
    echo  [ERROR] Failed to install Python dependencies.
    echo  Try: pip install -r requirements_translator.txt
    pause & exit /b 1
)
echo  Python dependencies ready.

:: ── Check Microsoft Word (optional, enhances DOCX quality) ───────────────────
echo.
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" >nul 2>&1
if errorlevel 1 (
    echo  [INFO] Microsoft Word not found.
    echo         DOCX translation will use OpenXML fallback (good quality).
    echo         Install Word for best-in-class DOCX fidelity.
) else (
    echo  [INFO] Microsoft Word detected — using Word COM for DOCX (best quality).
)

:: ── Build C# app ──────────────────────────────────────────────────────────────
echo.
echo  Building DocTranslate...
cd DocTranslate
dotnet build -c Release --nologo -v quiet 2>&1
if errorlevel 1 (
    echo  [ERROR] Build failed. See errors above.
    pause & exit /b 1
)
cd ..
echo  Build successful.

:: ── Launch ────────────────────────────────────────────────────────────────────
echo.
echo  Launching DocTranslate...
echo.
start "" "DocTranslate\bin\Release\net8.0-windows\DocTranslate.exe"

endlocal
