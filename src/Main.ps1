# App Version
$script:Version = "1.4.2"
$script:AppTitle = "PSGUI - v$($script:Version)"

# Constants
$script:GWL_STYLE = -16
$script:WS_BORDERLESS = 0x800000  # WS_POPUP without WS_BORDER, WS_CAPTION, etc.
$script:WS_OVERLAPPEDWINDOW = 0x00CF0000

# Settings - will be loaded from defaultsettings.json
$script:Settings = @{}

$script:State = @{
    CurrentDataFile = $null
    CurrentCommand = $null
    FavoritesHighestOrder = 0
    TabsReadOnly = $true
    RunCommandAttached = $script:Settings.DefaultRunCommandAttached
    ExtraColumnsVisibility = "Collapsed"
    ExtraColumns = @("Id", "Command", "PostCommand", "SkipParameterSelect", "PreCommand", "Transcript", "PSTask", "PSTaskMode", "PSTaskVisibilityLevel", "ShellOverride", "LogParameterNames")
    SubGridExpandedHeight = 300
    HasUnsavedChanges = $false
    CurrentCommandListId = $null
    RecycleBin = [System.Collections.Generic.Queue[object]]::new()
    RecycleBinMaxSize = 10
    CommandHistory = [System.Collections.Generic.List[object]]::new()
    OpenCommandWindows = [System.Collections.Generic.List[object]]::new()
    Username = $env:USERNAME
    DragDrop = @{
        DraggedItem = $null
        LastHighlightedRow = $null
        IsBottomBorder = $false
    }
}

# Initialize status timer variable
$script:StatusTimer = $null

# Determine app pathing whether running as PS script or EXE
if ($PSScriptRoot) { 
    $script:Path =  $PSScriptRoot
} 
else { 
    $script:Path = Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0]))
}

$script:ApplicationPaths = @{
    MainWindowXamlFile = Join-Path $script:Path "MainWindow.xaml"
    CommandWindowXamlFile = Join-Path $script:Path "CommandWindow.xaml"
    MaterialDesignThemes = Join-Path $script:Path "Assembly\MaterialDesignThemes.Wpf.dll"
    MaterialDesignColors = Join-Path $script:Path "Assembly\MaterialDesignColors.dll"
    DefaultDataFile = Join-Path $script:Path "data.json"
    DefaultSettingsFile = Join-Path $script:Path "defaultsettings.json"
    SettingsFilePath = Join-Path $env:APPDATA "PSGUI\settings.json"
    IconFile = Join-Path $script:Path "icon.ico"
    Win32APIFile = Join-Path $script:Path "Win32API.cs"
}

# Load default settings from defaultsettings.json
if (Test-Path $script:ApplicationPaths.DefaultSettingsFile) {
    try {
        $defaultSettings = Get-Content $script:ApplicationPaths.DefaultSettingsFile | ConvertFrom-Json
        foreach ($key in $defaultSettings.PSObject.Properties.Name) {
            $value = $defaultSettings.$key
            # Expand environment variables in paths
            if ($value -is [string] -and $value -match '%\w+%') {
                $script:Settings[$key] = [Environment]::ExpandEnvironmentVariables($value)
            } else {
                $script:Settings[$key] = $value
            }
        }
    }
    catch {
        Write-Warning "Failed to load default settings from defaultsettings.json: $_"
    }
}

# Load necessary assemblies
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsFormsIntegration
[Void][System.Reflection.Assembly]::LoadFrom($script:ApplicationPaths.MaterialDesignThemes)
[Void][System.Reflection.Assembly]::LoadFrom($script:ApplicationPaths.MaterialDesignColors)

# Load Win32 API functions - use embedded if available, otherwise load from file
if ($script:Win32API) {
    Add-Type -TypeDefinition $script:Win32API
}
else {
    Add-Type -Path $script:ApplicationPaths.Win32APIFile
}

Start-MainWindow