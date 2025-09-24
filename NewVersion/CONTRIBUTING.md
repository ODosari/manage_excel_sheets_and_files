# Contributing

## Setup
```bash
python -m venv .venv && source .venv/bin/activate
pip install -U pip
pip install -e ".[dev]"
pre-commit install
```

## Dev workflow
- Run `ruff check .` and `mypy .`
- Run tests: `pytest -q`
- Submit PRs with a descriptive title and linked issue.

## Releasing
- Bump version in `excelmgr/__init__.py` and `pyproject.toml`
- Tag: `git tag vX.Y.Z && git push --tags`
- CI builds wheels and (optionally) publishes.
