@echo off
setlocal
title DocTranslate — Release Build

echo.
echo  =====================================================
echo    DocTranslate — Release Build
echo  =====================================================
echo.

:: ── Publish C# as single-file exe ─────────────────────────────────────────────
echo  Publishing C# app...
cd DocTranslate
dotnet publish -c Release -r win-x64 --self-contained false ^
    /p:PublishSingleFile=true ^
    -o ..\dist\DocTranslate ^
    --nologo -v quiet
if errorlevel 1 (
    echo  [ERROR] Publish failed.
    pause & exit /b 1
)
cd ..

:: ── Copy Python microservice ───────────────────────────────────────────────────
echo  Copying Python translator...
xcopy /E /I /Y translator dist\DocTranslate\translator >nul

:: ── Copy requirements ─────────────────────────────────────────────────────────
copy /Y requirements_translator.txt dist\DocTranslate\ >nul

:: ── Write launcher bat for the dist folder ────────────────────────────────────
echo @echo off                                          > dist\run.bat
echo python -m pip install -r requirements_translator.txt --quiet >> dist\run.bat
echo start "" "DocTranslate\DocTranslate.exe"          >> dist\run.bat

echo.
echo  =====================================================
echo    Build complete: dist\DocTranslate\
echo.
echo    Distribute the entire dist\ folder.
echo    Users run dist\run.bat to install deps + launch.
echo  =====================================================
echo.

pause
endlocal
