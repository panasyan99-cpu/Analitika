from __future__ import annotations

from pathlib import Path


def test_streamlit_app_starts_without_exception() -> None:
    from streamlit.testing.v1 import AppTest

    app = AppTest.from_file(str(Path("streamlit_app.py").resolve()), default_timeout=30).run()
    assert not app.exception
