"""Build CalListenerSetup.exe with PyInstaller.

Run on a Windows machine that has Python + pywinauto + pyinstaller
installed (see `pip install -e .[windows,dev]`), or let GitHub Actions
do it via .github/workflows/build.yml.

Produces:
    dist/CalListener.exe                  (the single-file daemon)

Drop that .exe on any Windows laptop, double-click it, and:
  1. On first run a tkinter dialog asks for the 3 settings.
  2. The daemon writes %APPDATA%\\CalListener\\secrets.json.
  3. The daemon offers to register itself as a Scheduled Task
     (CalListener) that runs at logon and restarts on crash.
  4. From then on, every logon starts the listener automatically.

No zip files, no PowerShell, no separate Python install.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist"
BUILD = ROOT / "build"


def main():
    # Clean
    for d in (DIST, BUILD):
        if d.exists():
            shutil.rmtree(d)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--name", "CalListener",
        "--windowed",                          # no console window
        "--collect-all", "cal_listener",       # bundle all submodules
        # pywinauto + pywin32 do TONS of dynamic submodule loading.
        # --collect-all on each pulls the whole tree, which is the
        # only reliable way to make them work after PyInstaller bundling.
        "--collect-all", "pywinauto",
        "--collect-all", "comtypes",
        "--hidden-import", "psutil",
        "--hidden-import", "win32api",
        "--hidden-import", "win32con",
        "--hidden-import", "win32gui",
        "--hidden-import", "win32process",
        "--hidden-import", "pythoncom",
        "--icon", "NONE",
        str(ROOT / "cal_listener" / "__main__.py"),
    ]
    print("> " + " ".join(cmd))
    subprocess.check_call(cmd, cwd=str(ROOT))

    out = DIST / "CalListener.exe"
    if not out.exists():
        sys.exit(f"build failed: {out} not produced")
    print(f"\nBuilt: {out}  ({out.stat().st_size / 1024 / 1024:.1f} MB)")


if __name__ == "__main__":
    main()
