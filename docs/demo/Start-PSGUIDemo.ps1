<#
.SYNOPSIS
Launches PSGUI with sample commands in a throwaway profile, for screenshots and demos.

.DESCRIPTION
Uses a temporary %APPDATA%, so your real settings, history, favorites and data file are untouched.
The sample commands come from the PSGUIDemo module, which is added to PSModulePath for PSGUI and
for the shells it starts. A spare drive letter is mapped to the temporary folder so the window
title shows a short path instead of your profile path.

.PARAMETER Drive
Drive letter to map to the temporary folder. Use an unused letter.

.PARAMETER PrepareOnly
Set everything up and print the paths without launching PSGUI.

.EXAMPLE
./docs/demo/Start-PSGUIDemo.ps1
#>
[CmdletBinding()]
param(
    [ValidatePattern('^[A-Za-z]$')]
    [string]$Drive = 'P',

    [switch]$PrepareOnly
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '../..')
$demoRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'PSGUIDemo'
$driveName = "$($Drive.ToUpper()):"

& (Join-Path $repoRoot 'build.ps1') -Task Combine | Out-Null

if (Test-Path $demoRoot) { Remove-Item $demoRoot -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $demoRoot 'PSGUI') -Force | Out-Null
Copy-Item (Join-Path $PSScriptRoot 'demo-data.json') (Join-Path $demoRoot 'PSGUI/data.json')

if (Test-Path "$driveName\") { throw "Drive $driveName is already in use. Pass a different -Drive." }
subst $driveName $demoRoot
try {
    $env:APPDATA = "$driveName\"
    $env:PSModulePath = "$(Join-Path $PSScriptRoot '.');$env:PSModulePath"

    $script = Join-Path $repoRoot 'dist/PSGUI.ps1'
    if ($PrepareOnly) {
        "APPDATA      = $env:APPDATA"
        "Data file    = $env:APPDATA" + 'PSGUI\data.json'
        "Demo module  = $(Get-Module -ListAvailable PSGUIDemo | Select-Object -ExpandProperty Path)"
        "Launch with  = powershell.exe -STA -ExecutionPolicy Bypass -File $script"
        return
    }
    Start-Process powershell.exe -ArgumentList '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$script`"" -Wait
}
finally {
    subst $driveName /d | Out-Null
    Remove-Item $demoRoot -Recurse -Force -ErrorAction SilentlyContinue
}
