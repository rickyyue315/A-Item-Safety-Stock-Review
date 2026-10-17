"""Entrypoint for Streamlit Community Cloud (streamlit.app).

Community Cloud defaults to "streamlit_app.py" as the main file path.
This shim executes the real app (app.py) so the default works with no
deploy-time configuration. You can still point the main file path at
app.py directly; both behave identically.
"""
import runpy
from pathlib import Path

runpy.run_path(str(Path(__file__).with_name("app.py")))
