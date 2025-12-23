# Build for Linux

import subprocess
import webbrowser
import time
from pathlib import Path
import sys

def main():
    base_dir = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(__file__).parent

    cmd = [
        "python3",
        "-m",
        "streamlit",
        "run",
        str(base_dir / "app.py"),
        "--server.port", "8501",
        "--server.headless", "true",
    ]

    subprocess.Popen(cmd)
    time.sleep(2)
    webbrowser.open("http://localhost:8501")

if __name__ == "__main__":
    main()
