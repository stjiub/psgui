$global:App = [hashtable]::Synchronized(@{})
$App.Path = $PSScriptRoot
$App.MainWindowXamlFile = Join-Path $App.Path "MainWindow.xaml"
$App.CommandWindowXamlFile = Join-Path $App.Path "CommandWindow.xaml"
$App.DebugWindowXamlFile = Join-Path $App.Path "DebugWindow.xaml"
$App.MaterialDesignThemes = Join-Path $App.Path "Assembly\MaterialDesignThemes.Wpf.dll"
$App.MaterialDesignColors = Join-Path $App.Path "Assembly\MaterialDesignColors.dll"
$App.DefaultConfigFile = Join-Path $App.Path "data.json"

$App.UI = $null
$App.Command = @{}
$App.LastAddedId = 0
$App.LastCommandId = 0
$App.MainTabsReadOnly = $true
$App.AlternatingRowBackgroundColor = "#FF252526"

$runspace = [RunspaceFactory]::CreateRunspace()
$runspace.ApartmentState = "STA"
$runspace.ThreadOptions = "ReuseThread"      
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("App",$App)          
$App.PS = [PowerShell]::Create().AddScript({
    Add-Type -AssemblyName PresentationFramework
    [Void][System.Reflection.Assembly]::LoadFrom($App.MaterialDesignThemes)
    [Void][System.Reflection.Assembly]::LoadFrom($App.MaterialDesignColors)

    . "$($App.Path)\MainWindow.ps1"

    MainWindow
})
$App.PS.Runspace = $runspace
$App.Result = $App.PS.BeginInvoke()