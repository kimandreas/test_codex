from test_codex.cli import build_greeting


def test_build_greeting_uses_name() -> None:
    assert build_greeting("Ada") == "Hello, Ada!"

