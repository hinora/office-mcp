"""
Build standalone Windows .exe for wps-mcp using PyInstaller.

Usage:
    python build_exe.py          # Build single-file exe (dist/wps-mcp.exe)
    python build_exe.py --clean  # Clean build (remove build/ and dist/ first)

Output:
    dist/wps-mcp.exe  — Standalone executable, no Python install required.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
SPEC_FILE = PROJECT_ROOT / "wps_mcp.spec"
EXE_NAME = "wps-mcp.exe"


def clean() -> None:
    """Remove previous build artifacts."""
    for p in [DIST_DIR, BUILD_DIR, SPEC_FILE]:
        if p.exists():
            if p.is_dir():
                shutil.rmtree(p)
            else:
                p.unlink()
    print("[CLEAN] Removed build/ and dist/")


def build() -> None:
    """Run PyInstaller to build the standalone exe."""
    # Ensure pyinstaller is installed
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("Installing pyinstaller...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "pyinstaller"]
        )

    # Hidden imports required by pywin32 / COM
    hidden_imports = [
        "pythoncom",
        "win32com",
        "win32com.client",
        "win32com.client.dynamic",
        "win32com.server",
        "win32com.server.policy",
        "win32api",
        "win32event",
        "win32process",
        "win32timezone",
        "pywintypes",
    ]

    # Collections (huge stdlib parts we don't need — exclude them)
    exclude_modules = [
        "tkinter",
        "matplotlib",
        "numpy",
        "pandas",
        "PIL",
        "cv2",
        "scipy",
        "sqlalchemy",
        "IPython",
        "jupyter",
        "notebook",
        "django",
        "flask",
        "pytest",
        "setuptools",
        "distutils",
    ]

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",                  # Single .exe output
        "--console",                  # Keep console (stdio MCP needs it)
        "--name", "wps-mcp",
        "--add-data", f"{PROJECT_ROOT / 'src'}{os.pathsep}src",
        *[f"--hidden-import={m}" for m in hidden_imports],
        *[f"--exclude-module={m}" for m in exclude_modules],
        "--clean",
        "--noconfirm",
        str(PROJECT_ROOT / "src" / "wps_mcp" / "server.py"),
    ]

    # On Windows, path separator is ';'
    cmd[cmd.index("--add-data") + 1] = (
        f"{PROJECT_ROOT / 'src'}{os.pathsep}src"
    )

    print(f"[BUILD] Running PyInstaller...")
    print(f"[BUILD] Command: {' '.join(cmd)}")
    subprocess.check_call(cmd)

    # Verify output
    exe_path = DIST_DIR / EXE_NAME
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"\n✅ Build successful: {exe_path} ({size_mb:.1f} MB)")
    else:
        print(f"\n❌ Build failed: {EXE_NAME} not found in dist/")
        sys.exit(1)


def main() -> None:
    if "--clean" in sys.argv or "-c" in sys.argv:
        clean()

    if "--clean-only" in sys.argv:
        return

    build()


if __name__ == "__main__":
    main()
