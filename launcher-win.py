"""
Launcher entry point for the PyInstaller-built Windows executable.
Double-clicking the .exe runs this file, which starts the Streamlit server
and opens the browser automatically.
"""

import sys
import os

# On Windows, Streamlit/Tornado defaults to the asyncio Proactor event loop,
# whose socket-accept path raises noisy "OSError: [WinError 64] The specified
# network name is no longer available" (and can wedge) when a client connection
# drops abruptly (VPN/Wi-Fi change, sleep/wake, tab refresh, corporate proxy).
# The Selector event loop policy avoids this. It must be set before the Streamlit
# server starts, so it lives here in the entry point rather than in app.py
# (which streamlit only runs *after* the server's accept loop already exists).
if sys.platform == "win32":
    import asyncio
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

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
    # Velopack must process its install/update/uninstall hooks before anything
    # else: when invoked with --veloapp-* args (during install or after an
    # update) it runs the hook and exits fast. No-op when running in dev or when
    # the app was not installed via the Velopack Setup.exe. The actual update
    # check is user-driven from the app's sidebar (notify, don't force-restart),
    # so there is no automatic download/apply here.
    try:
        import velopack
        velopack.App().run()
    except Exception:
        pass

    # When frozen by PyInstaller, all bundled files live under sys._MEIPASS
    if getattr(sys, "frozen", False):
        base_dir = sys._MEIPASS
        exe_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        exe_dir = base_dir

    app_file = os.path.join(base_dir, "app.py")

    # Trace launches into update.log (same file app._apply_update writes) so a
    # post-update relaunch is visible: if this line appears after "apply: updater
    # spawned", Velopack restarted us successfully.
    if getattr(sys, "frozen", False):
        try:
            with open(os.path.join(exe_dir, "update.log"), "a", encoding="utf-8") as _f:
                _f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} launcher: started, argv={sys.argv[1:]}\n")
        except Exception:
            pass

    # Load API keys: compiled-in module when frozen, .env file in development.
    if getattr(sys, "frozen", False):
        import _bundled_config
        os.environ.setdefault("DIFY_API_BASE_URL", _bundled_config.API_BASE_URL)
        os.environ.setdefault("DIFY_API_KEY_MAIN", _bundled_config.API_KEY_MAIN)
        os.environ.setdefault("DIFY_API_KEY_SUB1", _bundled_config.API_KEY_SUB1)
        os.environ.setdefault("DIFY_API_KEY_SUB2", _bundled_config.API_KEY_SUB2)
        # Also load an optional .env next to the .exe so end users can configure
        # per-machine settings (e.g. TEMPLATE_SOURCE / TEMPLATE_BASE_PATH for the
        # OneDrive-synced templates). override=False keeps the baked API keys above.
        env_file = os.path.join(exe_dir, ".env")
        if os.path.exists(env_file):
            from dotenv import load_dotenv
            load_dotenv(env_file, override=False)
        # Apply the template/Feishu settings baked from the build-time .env as
        # defaults. setdefault runs *after* the runtime .env, so a per-machine .env
        # still wins for the non-secret settings, while the baked values (incl. the
        # Feishu app secret) make the exe work out of the box like dev.
        for _k, _v in getattr(_bundled_config, "RUNTIME_ENV", {}).items():
            os.environ.setdefault(_k, _v)
    else:
        env_file = os.path.join(base_dir, ".env")
        if os.path.exists(env_file):
            from dotenv import load_dotenv
            load_dotenv(env_file, override=True)

    # Open the browser — unless this start is a post-update relaunch. After an
    # update Velopack restarts us on the same :8501; the user's existing tab
    # (left on the old, now-killed server) auto-reconnects to the new one, so
    # opening another window just leaves two. _apply_update drops an
    # `update_restart.flag` right before it exits; consume it and skip the open.
    _post_update = False
    if getattr(sys, "frozen", False):
        flag = os.path.join(exe_dir, "update_restart.flag")
        if os.path.exists(flag):
            _post_update = True
            try:
                os.remove(flag)
            except Exception:
                pass
    if not _post_update:
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
