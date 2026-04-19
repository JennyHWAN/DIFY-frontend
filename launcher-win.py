"""
Launcher entry point for the PyInstaller-built Windows executable.
Double-clicking the .exe runs this file, which starts the Streamlit server
and opens the browser automatically.
"""

import sys
import os
import webbrowser
import threading
import time


def _open_browser():
    """Wait briefly for the server to start, then open the browser."""
    time.sleep(3)
    webbrowser.open("http://localhost:8501")


if __name__ == "__main__":
    # When frozen by PyInstaller, all bundled files live under sys._MEIPASS
    if getattr(sys, "frozen", False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    app_file = os.path.join(base_dir, "app.py")

    # Load .env from the directory next to the executable (user-editable)
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else base_dir
    env_file = os.path.join(exe_dir, ".env")
    if os.path.exists(env_file):
        from dotenv import load_dotenv
        load_dotenv(env_file, override=True)

    # Open the browser in the background after a short delay
    threading.Thread(target=_open_browser, daemon=True).start()

    # Launch Streamlit
    from streamlit.web import cli as stcli

    sys.argv = [
        "streamlit",
        "run",
        app_file,
        "--server.headless=true",
        "--server.port=8501",
        "--browser.gatherUsageStats=false",
        "--server.enableCORS=false",
    ]
    sys.exit(stcli.main())
