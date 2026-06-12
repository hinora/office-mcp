@echo off
REM ============================================================
REM  Build standalone MCP .exe files for Windows
REM  ============================================================
REM  This script creates:
REM    dist\wps-excel-mcp.exe — WPS Excel MCP server
REM    dist\outlook-mcp.exe   — Outlook MCP server
REM
REM  Requirements: Python 3.10+ must be installed on the BUILD
REM  machine (but NOT on the target/user machine).
REM
REM  Usage: build_exe.bat [--wps|--outlook|--all]
REM ============================================================

echo.
echo ========================================
echo   Building MCP standalone .exe files
echo ========================================
echo.

REM Check for Python
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python not found. Please install Python 3.10+ first.
    pause
    exit /b 1
)

REM Clean previous builds
if exist "build\" rmdir /s /q "build"
if exist "dist\"  rmdir /s /q "dist"
if exist "wps_excel_mcp.spec" del /q "wps_excel_mcp.spec"

REM Install/upgrade dependencies
echo [1/3] Installing dependencies...
python -m pip install -r requirements.txt --quiet
python -m pip install pyinstaller --quiet

REM Build the exes
echo [2/3] Building .exe files with PyInstaller...
python build_exe.py %*

REM Done
echo.
echo [3/3] Done!
echo.
if exist "dist\wps-excel-mcp.exe" (
    echo ✅  dist\wps-excel-mcp.exe is ready.
    echo.
    echo To use the WPS Excel MCP server, add this to your MCP client config:
    echo.
    echo   {
    echo     "wps-excel-mcp": {
    echo       "command": "D:\\work\\wps-mcp\\dist\\wps-excel-mcp.exe"
    echo     }
    echo   }
    echo.
)
if exist "dist\outlook-mcp.exe" (
    echo ✅  dist\outlook-mcp.exe is ready.
    echo.
    echo To use the Outlook MCP server, add this to your MCP client config:
    echo.
    echo   {
    echo     "outlook-mcp": {
    echo       "command": "D:\\work\\wps-mcp\\dist\\outlook-mcp.exe"
    echo     }
    echo   }
    echo.
)
if not exist "dist\wps-excel-mcp.exe" if not exist "dist\outlook-mcp.exe" (
    echo ❌  Build failed. Check errors above.
)
pause
