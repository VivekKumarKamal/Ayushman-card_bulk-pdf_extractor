import streamlit.web.cli as stcli
import sys, os

def resolve_path(path):
    basedir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(basedir, path)

if __name__ == "__main__":
    # 1. Set headless to false to ensure browser pops up
    # 2. Set developmentMode to false to hide "Usage Stats" warnings
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("app.py"),
        "--global.developmentMode=false",
    ]
    sys.exit(stcli.main())