@echo off
setlocal EnableExtensions
title SORAN Excel Extractor
color 0B

set "SCRIPT_DIR=%~dp0"
set "PYTHON_SCRIPT=%SCRIPT_DIR%process_excel_files.py"
set "DEFAULT_INPUT=%SCRIPT_DIR%source"
set "DEFAULT_OUTPUT=%SCRIPT_DIR%output"

cls
echo.
echo  ==============================================================
echo.
echo     SSSSS    OOOOO   RRRRR    AAAAA   N   N
echo    SS       OO   OO  RR  RR  AA   AA  NN  N
echo     SSSSS   OO   OO  RRRRR   AAAAAAA  N N N
echo         SS  OO   OO  RR  RR  AA   AA  N  NN
echo    SSSSSS    OOOOO   RR   RR AA   AA  N   N
echo.
echo  ==============================================================
echo.
echo                 SORAN EXCEL EXTRACTION TOOL
echo.
echo  Default Input Folder : "%DEFAULT_INPUT%"
echo  Default Output Folder: "%DEFAULT_OUTPUT%"
echo.

if not exist "%PYTHON_SCRIPT%" (
    echo [ERROR] Python script not found:
    echo "%PYTHON_SCRIPT%"
    echo.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    where python3 >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python was not found in PATH.
        echo Install Python and try again.
        echo.
        pause
        exit /b 1
    ) else (
        set "PYTHON_EXE=python3"
    )
) else (
    set "PYTHON_EXE=python"
)

if "%~1"=="" (
    set "INPUT_DIR=%DEFAULT_INPUT%"
    set "OUTPUT_DIR=%DEFAULT_OUTPUT%"
) else (
    set "INPUT_DIR=%~1"
    if "%~2"=="" (
        set "OUTPUT_DIR=%DEFAULT_OUTPUT%"
    ) else (
        set "OUTPUT_DIR=%~2"
    )
)

if not exist "%INPUT_DIR%" (
    echo [ERROR] Input folder not found:
    echo "%INPUT_DIR%"
    echo.
    pause
    exit /b 1
)

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

echo  Running...
echo  Input : "%INPUT_DIR%"
echo  Output: "%OUTPUT_DIR%"
echo.

"%PYTHON_EXE%" "%PYTHON_SCRIPT%" --input "%INPUT_DIR%" --output "%OUTPUT_DIR%"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo [DONE] SORAN finished successfully.
) else (
    echo [FAILED] SORAN finished with exit code %EXIT_CODE%.
)
echo.
pause
exit /b %EXIT_CODE%
