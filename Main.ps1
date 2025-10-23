# App Version
$script:Version = "1.4.0"
$script:AppTitle = "PSGUI - v$($script:Version)"

# Constants
$script:GWL_STYLE = -16
$script:WS_BORDERLESS = 0x800000  # WS_POPUP without WS_BORDER, WS_CAPTION, etc.
$script:WS_OVERLAPPEDWINDOW = 0x00CF0000

# Settings
$script:Settings = @{
    DefaultShell = "powershell"
    DefaultShellArgs = "-ExecutionPolicy Bypass -NoExit -Command `" & { [System.Console]::Title = 'PS' } `""
    DefaultRunCommandAttached = $true
    OpenShellAtStart = $false
    StatusTimeout = 3
    DefaultLogsPath = Join-Path $env:APPDATA "PSGUI"
    SettingsPath = Join-Path $env:APPDATA "PSGUI\settings.json"
    FavoritesPath = Join-Path $env:APPDATA "PSGUI\favorites.json"
    ShowDebugTab = $false
    DefaultDataFile = Join-Path $env:APPDATA "PSGUI\data.json"
    CommandHistoryLimit = 50
}

$script:State = @{
    CurrentDataFile = $null
    CurrentCommand = $null
    HighestId = 0
    FavoritesHighestOrder = 0
    TabsReadOnly = $true
    RunCommandAttached = $script:Settings.DefaultRunCommandAttached
    ExtraColumnsVisibility = "Collapsed"
    ExtraColumns = @("Id", "Command", "SkipParameterSelect", "PreCommand")
    SubGridExpandedHeight = 300
    HasUnsavedChanges = $false
    CurrentCommandListId = $null
    RecycleBin = [System.Collections.Generic.Queue[object]]::new()
    RecycleBinMaxSize = 10
    CommandHistory = [System.Collections.Generic.List[object]]::new()
    OpenCommandWindows = [System.Collections.Generic.List[object]]::new()
    DragDrop = @{
        DraggedItem = $null
        LastHighlightedRow = $null
        IsBottomBorder = $false
    }
}

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
    SettingsFilePath = Join-Path $env:APPDATA "PSGUI\settings.json"
    IconFile = Join-Path $script:Path "icon.ico"
    Win32APIFile = Join-Path $script:Path "Win32API.cs"
}

# Load necessary assemblies
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsFormsIntegration
[Void][System.Reflection.Assembly]::LoadFrom($script:ApplicationPaths.MaterialDesignThemes)
[Void][System.Reflection.Assembly]::LoadFrom($script:ApplicationPaths.MaterialDesignColors)

# Load Win32 API functions from Detached file
Add-Type -Path $script:ApplicationPaths.Win32APIFile

Start-MainWindow