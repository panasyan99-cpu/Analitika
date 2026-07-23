"""Test-only Streamlit fallback.

The production application installs Streamlit from requirements.txt. The build
container used for repository validation may not have it, so tests that exercise
pure business logic receive a small no-op module instead. When Streamlit is
installed this file does nothing.
"""
from __future__ import annotations

import sys
import types
import pytest
from contextlib import nullcontext

try:  # pragma: no cover - real package is preferred in CI/deployment
    import streamlit  # noqa: F401
except ModuleNotFoundError:  # pragma: no cover
    class _CachedDecorator:
        def __call__(self, *args, **kwargs):
            if args and callable(args[0]) and len(args) == 1 and not kwargs:
                return self._decorate(args[0])
            return self._decorate

        @staticmethod
        def _decorate(func):
            func.clear = lambda: None
            return func

    class _ColumnConfig:
        def __getattr__(self, name):
            return lambda *args, **kwargs: {"type": name, "args": args, **kwargs}

    class _Context:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        def __getattr__(self, name):
            return lambda *args, **kwargs: None

    class _DummyStreamlit(types.ModuleType):
        def __init__(self):
            super().__init__("streamlit")
            self.session_state = {}
            self.secrets = {}
            self.query_params = {}
            self.column_config = _ColumnConfig()
            self.cache_data = _CachedDecorator()
            self.cache_resource = _CachedDecorator()
            self.sidebar = _Context()

        @staticmethod
        def dialog(*args, **kwargs):
            return lambda func: func

        @staticmethod
        def fragment(func=None, *args, **kwargs):
            if callable(func):
                return func
            return lambda wrapped: wrapped

        @staticmethod
        def columns(spec, *args, **kwargs):
            count = spec if isinstance(spec, int) else len(spec)
            return [_Context() for _ in range(count)]

        @staticmethod
        def container(*args, **kwargs):
            return _Context()

        @staticmethod
        def expander(*args, **kwargs):
            return _Context()

        @staticmethod
        def spinner(*args, **kwargs):
            return nullcontext()

        @staticmethod
        def button(*args, **kwargs):
            return False

        @staticmethod
        def download_button(*args, **kwargs):
            return False

        @staticmethod
        def file_uploader(*args, **kwargs):
            return None

        @staticmethod
        def selectbox(label, options, index=0, **kwargs):
            values = list(options)
            return values[index] if values else None

        @staticmethod
        def segmented_control(label, options, default=None, **kwargs):
            return default if default is not None else (list(options)[0] if options else None)

        @staticmethod
        def radio(label, options, index=0, **kwargs):
            values = list(options)
            return values[index] if values else None

        @staticmethod
        def toggle(*args, value=False, **kwargs):
            return value

        @staticmethod
        def checkbox(*args, value=False, **kwargs):
            return value

        @staticmethod
        def number_input(*args, value=0, **kwargs):
            return value

        @staticmethod
        def text_input(*args, value="", **kwargs):
            return value

        @staticmethod
        def data_editor(data, *args, **kwargs):
            return data

        @staticmethod
        def rerun():
            return None

        def __getattr__(self, name):
            return lambda *args, **kwargs: None

    module = _DummyStreamlit()
    sys.modules["streamlit"] = module

    testing = types.ModuleType("streamlit.testing")
    testing_v1 = types.ModuleType("streamlit.testing.v1")

    class _UnavailableAppTest:
        @classmethod
        def from_file(cls, *args, **kwargs):
            pytest.skip("Streamlit AppTest is unavailable in this validation environment.")

    testing_v1.AppTest = _UnavailableAppTest
    sys.modules["streamlit.testing"] = testing
    sys.modules["streamlit.testing.v1"] = testing_v1
