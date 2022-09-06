param(
    [Parameter(Mandatory = $true)]
    [string]$Version
)

$module = Get-Module -ListAvailable PS2Exe

if (-not $module) {
    Install-Module PS2Exe
}

$iconFile = Get-ChildItem $PSScriptRoot -Filter "*.ico"

ps2exe $PSScriptRoot\PSGUI.ps1 -requireAdmin -title "PSGUI" -noConsole -version $Version -iconFile $iconFile.FullName