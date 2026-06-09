# v2.0.0 Release Steps (Run these locally)

This document contains ready-to-run commands for the v2.0.0 release of AdvancedExcel.

## Prerequisites (run once)

```powershell
# Make sure you have the tools
pip install build twine

# Authenticate to GitHub (if not already)
gh auth login
```

## 1. Create the Pull Request (from your local clone)

```powershell
cd C:\DEV\TE\AdvancedExcel

gh pr create `
  --base main `
  --head refactor/split-modules `
  --title "Refactor: snake_case API + modern package structure (v2.0) + real Excel test fixtures" `
  --body-file PR_BODY_v2.md
```

After creation, GitHub will give you the PR number (e.g. #42).

## 2. Merge the PR

You can merge from the GitHub web UI (recommended for reviews) **or** via CLI:

```powershell
# Replace 42 with the actual PR number
gh pr merge 42 --merge --delete-branch
```

Wait for the merge to complete and the `main` branch to be updated.

## 3. Create GitHub Release v2.0.0

```powershell
cd C:\DEV\TE\AdvancedExcel

# Make sure you're on main and up to date
git checkout main
git pull origin main

# Create the release using the prepared notes
gh release create v2.0.0 `
  --title "v2.0.0 - Snake Case API & Real Excel Test Fixtures" `
  --notes-file RELEASE_NOTES_v2.0.md `
  --target main
```

This will also automatically attach the tag.

## 4. Build & Publish to PyPI

```powershell
cd C:\DEV\TE\AdvancedExcel

# Clean previous builds (optional but recommended)
Remove-Item -Recurse -Force dist, build, *.egg-info -ErrorAction SilentlyContinue

# Build source distribution + wheel
python -m build

# Upload to PyPI (you will be prompted for token if not configured)
twine upload dist/*

# Or upload to TestPyPI first for verification:
# twine upload --repository testpypi dist/*
```

### PyPI Token Setup (recommended)

Create a token at https://pypi.org/manage/account/token/ with scope "Entire account" or scoped to the `advanced-excel` project.

Then either:
- Use `twine` interactive prompt, or
- Set environment variables before upload:
  ```powershell
  $env:TWINE_USERNAME = "__token__"
  $env:TWINE_PASSWORD = "pypi-XXXXXXXXXXXXXXXXXXXXXXXX"
  twine upload dist/*
  ```

## Verification After Publish

```bash
pip install --upgrade advanced-excel
python -c "from advanced_excel import AdvancedExcel; print(AdvancedExcel.__version__)"
```

You should see `2.0.0`.

## Optional: Update the main branch tag if needed

The tag `v2.0.0` should already point to the merge commit on `main`.

---

All supporting files created:
- `PR_BODY_v2.md` — used in step 1
- `RELEASE_NOTES_v2.0.md` — used in step 3
- `dist/advanced_excel-2.0.0.tar.gz` and `dist/advanced_excel-2.0.0-py3-none-any.whl` — ready for step 4

Good luck with the release! Let me know if you hit any issues.