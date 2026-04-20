# test_codex

A tiny Python CLI starter project created with Codex.

## What It Does

The CLI prints a friendly greeting. You can use the default name or pass your own.

## Setup

Create and activate a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Install the project with test dependencies:

```bash
python -m pip install -e ".[test]"
```

## Run The CLI

```bash
test-codex
```

Or greet someone by name:

```bash
test-codex Ada
```

You can also run it without installing:

```bash
python -m test_codex.cli Ada
```

## Run Tests

```bash
PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 pytest
```

That environment variable keeps pytest focused on this project's tests and avoids auto-loading unrelated plugins from global tools such as ROS.
