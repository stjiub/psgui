<#
.SYNOPSIS
Build tasks for PSGUI.

.DESCRIPTION
Combine   Assemble src/ into a single dist/PSGUI.ps1 and copy runtime assets next to it.
Build     Combine, then compile dist/PSGUI.exe with PS2EXE.
Analyze   Run PSScriptAnalyzer over src/.
Test      Analyze, then run the Pester tests.
Run       Combine, then launch dist/PSGUI.ps1 in Windows PowerShell (STA).

.EXAMPLE
./build.ps1 -Task Test -Coverage

.EXAMPLE
./build.ps1 -Task Build -RequireAdmin
#>
[CmdletBinding()]
param(
    [ValidateSet('Combine', 'Build', 'Analyze', 'Test', 'Run')]
    [string]$Task = 'Combine',

    [switch]$Coverage,

    # Ask for elevation when the compiled exe starts
    [switch]$RequireAdmin
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$src = Join-Path $root 'src'
$assets = Join-Path $root 'assets'
$dist = Join-Path $root 'dist'

function Get-AppVersion {
    $match = Select-String -Path (Join-Path $src 'Main.ps1') -Pattern '^\$script:Version\s*=\s*"([^"]+)"'
    if (-not $match) { throw 'Could not find $script:Version in src/Main.ps1' }
    $match.Matches[0].Groups[1].Value
}

function Invoke-Combine {
    New-Item -ItemType Directory -Path $dist -Force | Out-Null
    $out = Join-Path $dist 'PSGUI.ps1'

    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($file in Get-ChildItem (Join-Path $src 'Functions') -Filter '*.ps1' | Sort-Object Name) {
        $parts.Add((Get-Content $file.FullName -Raw))
    }
    foreach ($embed in @(
        @{ Variable = 'MainWindowXaml';    File = 'MainWindow.xaml' }
        @{ Variable = 'CommandWindowXaml'; File = 'CommandWindow.xaml' }
        @{ Variable = 'Win32API';          File = 'Win32API.cs' }
    )) {
        $parts.Add("`n# Embedded $($embed.File)`n`$script:$($embed.Variable) = @`"`n$(Get-Content (Join-Path $src $embed.File) -Raw)`n`"@")
    }
    $parts.Add((Get-Content (Join-Path $src 'Main.ps1') -Raw))

    Set-Content -Path $out -Value ($parts -join "`n") -Encoding UTF8
    Copy-Item (Join-Path $assets '*') -Destination $dist -Recurse -Force
    Write-Host "Combined script: $out" -ForegroundColor Green
}

function Invoke-Build {
    Invoke-Combine
    if (-not (Get-Module -ListAvailable ps2exe)) {
        Install-Module ps2exe -Scope CurrentUser -Force
    }
    $params = @{
        InputFile  = Join-Path $dist 'PSGUI.ps1'
        OutputFile = Join-Path $dist 'PSGUI.exe'
        IconFile   = Join-Path $assets 'icon.ico'
        Title      = 'PSGUI'
        Version    = Get-AppVersion
        NoConsole  = $true
    }
    if ($RequireAdmin) { $params.RequireAdmin = $true }
    Invoke-ps2exe @params
}

function Invoke-Analyze {
    $settings = Join-Path $root 'PSScriptAnalyzerSettings.psd1'
    $results = Get-ChildItem $src -Filter '*.ps1' -Recurse |
        ForEach-Object { Invoke-ScriptAnalyzer -Path $_.FullName -Settings $settings }
    if ($results) {
        $results | Format-Table RuleName, Severity, ScriptName, Line, Message -AutoSize -Wrap
        throw "PSScriptAnalyzer found $(@($results).Count) issue(s)."
    }
    Write-Host 'PSScriptAnalyzer: no issues' -ForegroundColor Green
}

function Invoke-Test {
    Import-Module Pester -MinimumVersion 5.5 -ErrorAction Stop
    $ErrorActionPreference = 'Continue'
    $config = New-PesterConfiguration
    $config.Run.Path = Join-Path $root 'tests'
    $config.Run.PassThru = $true
    $config.Output.Verbosity = 'Normal'
    $config.TestResult.Enabled = $true
    $config.TestResult.OutputPath = Join-Path $root 'TestResults/testResults.xml'
    if ($Coverage) {
        $config.CodeCoverage.Enabled = $true
        $config.CodeCoverage.Path = Join-Path $src 'Functions'
        $config.CodeCoverage.OutputPath = Join-Path $root 'coverage.xml'
        $config.CodeCoverage.OutputFormat = 'JaCoCo'
        # Most of src/Functions is WPF event handling that needs a live window; raise this as logic is extracted
        $config.CodeCoverage.CoveragePercentTarget = 10
    }
    $result = Invoke-Pester -Configuration $config
    if ($result.FailedCount -gt 0) {
        throw "$($result.FailedCount) test(s) failed."
    }
}

switch ($Task) {
    'Combine' { Invoke-Combine }
    'Build'   { Invoke-Build }
    'Analyze' { Invoke-Analyze }
    'Test'    { Invoke-Analyze; Invoke-Test }
    'Run'     {
        Invoke-Combine
        Start-Process powershell.exe -ArgumentList '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $dist 'PSGUI.ps1')
    }
}
