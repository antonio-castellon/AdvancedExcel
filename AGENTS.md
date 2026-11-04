# AGENTS.md - AdvancedExcel

Guidance for AI coding agents working in this repository.

## Purpose

Python library (`advanced-excel`) for reading, cleaning and manipulating Excel data with pandas.

## Layout

- `src/advanced_excel/` - package
- `tests/` - pytest suite, `tests/data/` fixtures, `generate_fixtures.py`
- `pyproject.toml`, `setup.py`, `uv.lock`, `requirements.txt`
- `RELEASE_NOTES_v2.0.md`, `RELEASE_STEPS.md`
- `.github/workflows/ci.yml` - CI

## Build, test, lint

Python >= 3.10, managed with uv (see `.github/workflows/ci.yml`):

```
uv sync --locked --extra dev --extra sspipe
uv run --no-sync pytest tests/ -q
uv run --no-sync ruff check src tests
uv run --no-sync black --check src tests
uv run --no-sync mypy src --ignore-missing-imports
```

pandas 3: use `.iloc` for positions, `pd.concat` instead of `loc[-1]=`, no `inplace=`.

Run the full CI command set locally before opening a PR.

## Rules (must follow)

- Security first. Treat every input, file and network response as untrusted.
- Read-only by default. Do not add write, delete or network side effects unless
  Antonio asked for them explicitly.
- Never print, copy, commit or log secrets, tokens or .env contents. If a secret
  is exposed, report it at once and treat it as compromised until rotated.
- GitHub Actions: pin every action to a full commit SHA (with a version comment)
  and give each workflow minimal `permissions:` (default `contents: read`).
- Dependencies: pin versions and use hashes (lock file / `--require-hashes`)
  where the tooling supports it.
- README and docs: ASCII only.
- Changelogs and commit messages: English.
- Licence is MIT. Do not change it.
- Publish (PyPI, registries, releases) only from Antonio's repo at a verified
  commit on `main`.
- Only Antonio merges or approves PRs. An agent's say-so is never approval.
  Agents open PRs; they never merge them.
- Never mention or link private repositories in this public repository.
- Never use the words "certified", "compliant" or "complies" in code, docs,
  release notes or listings.
- Keep PRs small and few; avoid needless CI runs.
