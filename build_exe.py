"""
Build standalone Windows .exe for WPS Excel MCP, WPS Word MCP, and Outlook MCP using PyInstaller.

Usage:
    python build_exe.py             # Build all exes
    python build_exe.py --wps       # Build only wps-excel-mcp.exe
    python build_exe.py --word      # Build only wps-word-mcp.exe
    python build_exe.py --outlook   # Build only outlook-mcp.exe
    python build_exe.py --clean     # Clean build (remove build/ and dist/ first)

Output:
    dist/wps-excel-mcp.exe  — Standalone WPS Excel MCP server
    dist/wps-word-mcp.exe   — Standalone WPS Word MCP server
    dist/outlook-mcp.exe    — Standalone Outlook MCP server
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
SPEC_FILE = PROJECT_ROOT / "wps_excel_mcp.spec"
EXE_NAME = "wps-excel-mcp.exe"
WORD_EXE_NAME = "wps-word-mcp.exe"
OUTLOOK_EXE_NAME = "outlook-mcp.exe"

# Common hidden imports required by pywin32 / COM
HIDDEN_IMPORTS_COMMON = [
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
EXCLUDE_MODULES = [
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


def clean() -> None:
    """Remove previous build artifacts."""
    for p in [DIST_DIR, BUILD_DIR, SPEC_FILE]:
        if p.exists():
            if p.is_dir():
                shutil.rmtree(p)
            else:
                p.unlink()
    print("[CLEAN] Removed build/ and dist/")


def ensure_pyinstaller() -> None:
    """Ensure pyinstaller is installed."""
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("Installing pyinstaller...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "pyinstaller"]
        )


def build(target: str = "all") -> None:
    """Run PyInstaller to build the standalone exe.

    Args:
        target: 'wps' for wps-excel-mcp, 'word' for wps-word-mcp,
                'outlook' for outlook-mcp, 'all' for all three.
    """
    ensure_pyinstaller()

    targets_to_build = []
    if target == "all":
        targets_to_build = ["wps", "word", "outlook"]
    else:
        targets_to_build = [target]

    for t in targets_to_build:
        _build_one(t)

    # List all built exes
    print()
    for t in targets_to_build:
        exe_map = {"wps": EXE_NAME, "word": WORD_EXE_NAME, "outlook": OUTLOOK_EXE_NAME}
        exe_name = exe_map.get(t, "")
        if exe_name and (DIST_DIR / exe_name).exists():
            size_mb = (DIST_DIR / exe_name).stat().st_size / (1024 * 1024)
            print(f"[OK] {exe_name} ({size_mb:.1f} MB)")


def _build_one(target: str) -> None:
    """Build a single MCP exe."""
    if target == "wps":
        exe_name = EXE_NAME
        server_path = PROJECT_ROOT / "src" / "wps_excel_mcp" / "server.py"
        project_hidden = [
            "wps_excel_mcp",
            "wps_excel_mcp.wps_client",
            "wps_excel_mcp.tools",
        ]
    elif target == "word":
        exe_name = WORD_EXE_NAME
        server_path = PROJECT_ROOT / "src" / "wps_word_mcp" / "server.py"
        project_hidden = [
            "wps_word_mcp",
            "wps_word_mcp.word_client",
            "wps_word_mcp.tools",
        ]
    elif target == "outlook":
        exe_name = OUTLOOK_EXE_NAME
        server_path = PROJECT_ROOT / "src" / "outlook_mcp" / "server.py"
        project_hidden = [
            "outlook_mcp",
            "outlook_mcp.outlook_client",
        ]
    else:
        print(f"[ERROR] Unknown target: {target}")
        sys.exit(1)

    hidden_imports = HIDDEN_IMPORTS_COMMON + project_hidden

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--console",
        "--name", exe_name.replace(".exe", ""),
        "--add-data", f"{PROJECT_ROOT / 'src'}{os.pathsep}src",
        *[f"--hidden-import={m}" for m in hidden_imports],
        *[f"--exclude-module={m}" for m in EXCLUDE_MODULES],
        "--clean",
        "--noconfirm",
        str(server_path),
    ]

    # On Windows, path separator is ';'
    cmd[cmd.index("--add-data") + 1] = (
        f"{PROJECT_ROOT / 'src'}{os.pathsep}src"
    )

    print(f"\n[BUILD] Building {exe_name}...")
    print(f"[BUILD] Source: {server_path}")
    subprocess.check_call(cmd)

    # Verify output
    exe_path = DIST_DIR / exe_name
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"[BUILD] OK: {exe_name} ({size_mb:.1f} MB)")
    else:
        print(f"[BUILD] FAILED: {exe_name} not found in dist/")
        sys.exit(1)


def main() -> None:
    target = "all"
    do_clean = False

    for arg in sys.argv[1:]:
        if arg == "--clean":
            do_clean = True
        elif arg in ("--wps", "--wps-excel"):
            target = "wps"
        elif arg in ("--word", "--wps-word"):
            target = "word"
        elif arg in ("--outlook",):
            target = "outlook"
        elif arg in ("--all",):
            target = "all"

    if do_clean:
        clean()

    build(target)


if __name__ == "__main__":
    main()
