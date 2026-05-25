"""
Launcher entry point for the PyInstaller-built Windows executable.
Double-clicking the .exe runs this file, which starts the Streamlit server
and opens the browser automatically.
"""

import sys
import os

# Force pure-Python protobuf before any import of google.protobuf (including
# transitive imports via streamlit).  The bundled _message.pyd C-extension
# crashes with STATUS_ACCESS_VIOLATION (c0000005) when Streamlit sends a
# WebSocket delta right after json.loads allocates the large MAIN-Output
# node_finished payload.  The pure-Python implementation is slower but stable.
os.environ.setdefault("PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION", "python")

import webbrowser
import threading
import time
import traceback


def _open_browser():
    """Wait briefly for the server to start, then open the browser."""
    time.sleep(3)
    webbrowser.open("http://localhost:8501")


def main():
    # When frozen by PyInstaller, all bundled files live under sys._MEIPASS
    if getattr(sys, "frozen", False):
        base_dir = sys._MEIPASS
        exe_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        exe_dir = base_dir

    app_file = os.path.join(base_dir, "app.py")

    # Load API keys: compiled-in module when frozen, .env file in development.
    if getattr(sys, "frozen", False):
        import _bundled_config
        os.environ.setdefault("DIFY_API_BASE_URL", _bundled_config.API_BASE_URL)
        os.environ.setdefault("DIFY_API_KEY_MAIN", _bundled_config.API_KEY_MAIN)
        os.environ.setdefault("DIFY_API_KEY_SUB1", _bundled_config.API_KEY_SUB1)
        os.environ.setdefault("DIFY_API_KEY_SUB2", _bundled_config.API_KEY_SUB2)
    else:
        env_file = os.path.join(base_dir, ".env")
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
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.port=8501",
        "--browser.gatherUsageStats=false",
        "--server.enableCORS=false",
    ]
    sys.exit(stcli.main())


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        exe_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
        log_file = os.path.join(exe_dir, "launch_error.log")
        error_text = traceback.format_exc()
        with open(log_file, "w") as f:
            f.write(error_text)
        print("\n--- CRASH ---")
        print(error_text)
        print(f"Error also saved to: {log_file}")
        input("\nPress Enter to exit...")
