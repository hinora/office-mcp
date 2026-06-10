@echo off
REM ============================================================
REM  Build standalone wps-mcp.exe for Windows
REM  ============================================================
REM  This script creates dist\wps-mcp.exe — a single file that
REM  runs the WPS MCP server without any Python installation.
REM
REM  Requirements: Python 3.10+ must be installed on the BUILD
REM  machine (but NOT on the target/user machine).
REM ============================================================

echo.
echo ========================================
echo   Building wps-mcp.exe (standalone)
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
if exist "wps_mcp.spec" del /q "wps_mcp.spec"

REM Install/upgrade dependencies
echo [1/3] Installing dependencies...
python -m pip install -r requirements.txt --quiet
python -m pip install pyinstaller --quiet

REM Build the exe
echo [2/3] Building .exe with PyInstaller...
python build_exe.py

REM Done
echo.
echo [3/3] Done!
echo.
if exist "dist\wps-mcp.exe" (
    echo ✅  dist\wps-mcp.exe is ready.
    echo.
    echo To use the MCP server, add this to your MCP client config:
    echo.
    echo   {
    echo     "wps-mcp": {
    echo       "command": "D:\\work\\wps-mcp\\dist\\wps-mcp.exe"
    echo     }
    echo   }
    echo.
) else (
    echo ❌  Build failed. Check errors above.
)
pause
