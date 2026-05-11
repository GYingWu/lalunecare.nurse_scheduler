import os
import sys
import webbrowser

from streamlit.web import cli as stcli


def main() -> None:
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    app_path = os.path.join(base_dir, "streamlit_app.py")
    webbrowser.open("http://localhost:8501")
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--server.headless",
        "true",
    ]
    raise SystemExit(stcli.main())


if __name__ == "__main__":
    main()
