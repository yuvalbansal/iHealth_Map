# Build for Windows

import subprocess
import sys
import webbrowser
import time
from pathlib import Path
import os

def main():
    # Resolve base directory (works inside PyInstaller)
    if hasattr(sys, "_MEIPASS"):
        base_dir = Path(sys._MEIPASS)
        exe_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
        exe_dir = base_dir

    # Locate bundled Python interpreter (Windows only)
    python_path = exe_dir / "_internal" / "python.exe"

    if not python_path.exists():
        raise RuntimeError(f"Bundled Python not found at: {python_path}")

    # Prevent recursive relaunch
    if os.environ.get("IHEALTHMAP_RUNNING"):
        return
    os.environ["IHEALTHMAP_RUNNING"] = "1"

    cmd = [
        str(python_path),
        "-m",
        "streamlit",
        "run",
        str(base_dir / "app.py"),
        "--server.port", "8501",
        "--server.headless", "true",
    ]

    subprocess.Popen(cmd, creationflags=subprocess.CREATE_NO_WINDOW)

    time.sleep(2)
    webbrowser.open("http://localhost:8501")

if __name__ == "__main__":
    main()
