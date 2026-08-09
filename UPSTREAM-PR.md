# Contributing Back to Upstream (exceljs/exceljs)

## Overview

This repo is a fork of [exceljs/exceljs](https://github.com/exceljs/exceljs).  
Fork-specific work lives on `development`. The `master` branch is kept clean for upstream PRs.

---

## One-time Setup

```bash
git remote add upstream https://github.com/exceljs/exceljs.git
git fetch upstream --tags
```

---

## Workflow: Prepare master for upstream PR

### 1. Create / reset local master from upstream

```bash
git fetch upstream
git checkout master || git checkout -b master upstream/master
git reset --hard upstream/master
```

### 2. Merge development

```bash
git merge development --no-edit
```

### 3. Remove fork-only files (not meant for upstream)

```bash
# Delete new files that only belong in this fork
git rm -f \
  .codesight/default.json \
  .github/workflows/sync-master.yml \
  CONTRIBUTING.md \
  package-lock.json

# Restore upstream versions of files the fork modified
git checkout upstream/master -- \
  .github/workflows/tests.yml \
  .gitignore \
  .npmrc \
  README.md
```

### 4. Restore upstream identity fields in package.json

Edit `package.json` and set:

| Field | Upstream value |
|---|---|
| `name` | `"exceljs"` |
| `version` | `"4.4.0"` (match upstream latest) |
| `author` | `{ "name": "Guyon Roche", "email": "guyon@live.com" }` |
| `repository.url` | `"https://github.com/exceljs/exceljs.git"` |
| `bugs.url` | `"https://github.com/exceljs/exceljs/issues"` |
| `forkedFrom` | **remove entirely** |

### 5. Commit and force-push master

```bash
git add package.json
git commit -m "chore: restore upstream package.json identity fields for PR"
git commit -m "chore: remove fork-only files from upstream PR scope"  # if separate
git push origin master --force-with-lease
```

### 6. Verify PR diff contains only real changes

```bash
git diff upstream/master --stat
```

Expected files only:
- `lib/exceljs.browser.js`
- `lib/stream/xlsx/workbook-reader.js`
- `lib/utils/parse-sax.js`
- `package.json` (deps only)
- `spec/integration/workbook/workbook.spec.js`

### 7. Open Pull Request

https://github.com/ArfanKhalilMughal/exceljs/compare/exceljs:master...ArfanKhalilMughal:master

---

## Files that ONLY belong in `development` (never in upstream PR)

| File | Reason |
|---|---|
| `.codesight/default.json` | Fork tooling |
| `.github/workflows/sync-master.yml` | Fork CI — auto-sync with upstream |
| `.github/workflows/tests.yml` | Fork CI overrides (Node 22, OIDC publish) |
| `.gitignore` | Fork-specific ignores |
| `.npmrc` | Fork npm config |
| `CONTRIBUTING.md` | Fork contribution guide |
| `README.md` | Fork-specific documentation |
| `package-lock.json` | Fork lockfile |
| `package.json` identity fields | `name`, `version`, `author`, `repository`, `bugs`, `forkedFrom` |
