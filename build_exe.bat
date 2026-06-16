@echo off
REM ============================================================
REM  Build standalone MCP .exe files for Windows
REM  ============================================================
REM  This script creates:
REM    dist\wps-excel-mcp.exe — WPS Excel MCP server
REM    dist\wps-word-mcp.exe  — WPS Word MCP server
REM    dist\wps-slide-mcp.exe — WPS Slide MCP server
REM    dist\outlook-mcp.exe   — Outlook MCP server
REM    dist\mcp-meta.exe      — MCP Meta server (web-search, web-fetch)
REM    dist\whatsapp-mcp.exe  — WhatsApp MCP server
REM
REM  Requirements: Python 3.10+ must be installed on the BUILD
REM  machine (but NOT on the target/user machine).
REM
REM  Usage: build_exe.bat [--wps|--word|--slide|--outlook|--meta|--whatsapp|--all]
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
if exist "dist\wps-word-mcp.exe" (
    echo ✅  dist\wps-word-mcp.exe is ready.
    echo.
    echo To use the WPS Word MCP server, add this to your MCP client config:
    echo.
    echo   {
    echo     "wps-word-mcp": {
    echo       "command": "D:\\work\\wps-mcp\\dist\\wps-word-mcp.exe"
    echo     }
    echo   }
    echo.
)
if exist "dist\wps-slide-mcp.exe" (
    echo ✅  dist\wps-slide-mcp.exe is ready.
    echo.
    echo To use the WPS Slide MCP server, add this to your MCP client config:
    echo.
    echo   {
    echo     "wps-slide-mcp": {
    echo       "command": "D:\\work\\wps-mcp\\dist\\wps-slide-mcp.exe"
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
if exist "dist\mcp-meta.exe" (
    echo ✅  dist\mcp-meta.exe is ready.
    echo.
    echo To use the MCP Meta server, add this to your MCP client config:
    echo.
    echo   {
    echo     "mcp-meta": {
    echo       "command": "D:\\work\\wps-mcp\\dist\\mcp-meta.exe",
    echo       "env": {
    echo         "BRAVE_API_KEY": "your-brave-api-key-here",
    echo         "SEARCH_PROVIDER": "brave"
    echo       }
    echo     }
    echo   }
    echo.
)
if exist "dist\whatsapp-mcp.exe" (
    echo ✅  dist\whatsapp-mcp.exe is ready.
    echo.
    echo To use the WhatsApp MCP server, add this to your MCP client config:
    echo.
    echo   {
    echo     "whatsapp-mcp": {
    echo       "command": "D:\\work\\wps-mcp\\dist\\whatsapp-mcp.exe"
    echo     }
    echo   }
    echo.
)
if not exist "dist\wps-excel-mcp.exe" if not exist "dist\wps-word-mcp.exe" if not exist "dist\wps-slide-mcp.exe" if not exist "dist\outlook-mcp.exe" if not exist "dist\mcp-meta.exe" if not exist "dist\whatsapp-mcp.exe" (
    echo ❌  Build failed. Check errors above.
)
pause
