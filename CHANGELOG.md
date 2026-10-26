# Changelog

## Unreleased

- Fix: opening a legacy data file that contained a single command loaded nothing
- Fix: Enter no longer stops running the selected command after opening or importing a different data file
- Fix: saving a data file now writes UTF-8 and always writes `Commands` as an array
- Change: the exe no longer requests elevation by default; build with `./build.ps1 -Task Build -RequireAdmin` to restore it
- Change: sources moved to `src/`, shipped files to `assets/`, and `build.ps1` replaces the psake script
- Remove: the unused `SnippetsPath` setting
- Add: Pester tests, PSScriptAnalyzer settings, GitHub Actions CI and tag-driven releases
- Add: README, contributing guide and third-party notices

## 1.4.2

First version in this repository.
