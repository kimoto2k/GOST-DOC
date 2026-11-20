"""Smoke test: Streamlit should render its main interface without exceptions."""
from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_app_renders_homepage():
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    at = AppTest.from_file(str(app_path), default_timeout=25).run()
    assert not at.exception, [str(exc) for exc in at.exception]
    assert at.title is not None
