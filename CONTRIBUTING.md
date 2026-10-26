# Contributing

1. Fork and branch from `main`.
2. Make your change and add or update tests in `tests/` when it touches logic that can run without the UI.
3. Run `./build.ps1 -Task Test` and make sure it passes.
4. If the change is visible to users, add a line to `CHANGELOG.md`.
5. Open a pull request describing what changed and why.

## How the code fits together

`build.ps1 -Task Combine` concatenates `src/Functions/*.ps1` in file-name order, embeds the XAML and C# as here-strings, then appends `src/Main.ps1`. That single script is what runs and what PS2EXE compiles, so:

- Functions are plain definitions; nothing in `src/Functions` should run code at load time.
- All app state lives in `$script:State`, `$script:Settings` and `$script:UI`.
- A control used from code needs a `Name` in the XAML. The tests fail if `$script:UI.<name>` has no matching control.
- New settings need an entry in `assets/defaultsettings.json`, `Load-Settings`, `Save-Settings` and `Create-DefaultSettings`.

Keep logic that does not need a window in small functions that take plain inputs, so it can be tested.

## Style

- Four-space indentation. PowerShell, C# and XAML are checked out with CRLF line endings (see `.gitattributes`).
- Comments only where the reason is not obvious from the code.
- Keep `PSScriptAnalyzerSettings.psd1` exclusions to a minimum and explain any new one.
- Commit messages are one line in the form `area: what changed`.
