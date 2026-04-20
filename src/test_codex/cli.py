"""Command-line interface for test_codex."""

from __future__ import annotations

import argparse


def build_greeting(name: str = "world") -> str:
    """Return a friendly greeting for the provided name."""
    clean_name = name.strip() or "world"
    return f"Hello, {clean_name}!"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Print a friendly greeting.")
    parser.add_argument(
        "name",
        nargs="?",
        default="world",
        help="Name to greet. Defaults to 'world'.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    print(build_greeting(args.name))


if __name__ == "__main__":
    main()

