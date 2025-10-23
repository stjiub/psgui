# Set app settings from loaded settings
function Initialize-Settings {
    Load-Settings
    # Update UI elements with loaded settings
    $script:UI.TxtDefaultShell.Text = $script:Settings.DefaultShell
    $script:UI.TxtDefaultShellArgs.Text = $script:Settings.DefaultShellArgs
    $script:UI.ChkRunCommandAttached.IsChecked = $script:Settings.DefaultRunCommandAttached
    $script:UI.ChkOpenShellAtStart.IsChecked = $script:Settings.OpenShellAtStart
    $script:UI.TxtDefaultLogsPath.Text = $script:Settings.DefaultLogsPath
    $script:UI.TxtDefaultDataFile.Text = $script:Settings.DefaultDataFile
    $script:UI.TxtCommandHistoryLimit.Text = $script:Settings.CommandHistoryLimit
    $script:UI.ChkShowDebugTab.IsChecked = $script:Settings.ShowDebugTab

    # Set the Debug tab visibility based on setting
    if ($script:Settings.ShowDebugTab) {
        $script:UI.LogTabControl.Items[0].Visibility = "Visible"
    } else {
        $script:UI.LogTabControl.Items[0].Visibility = "Collapsed"
    }
}

# Set the default app settings
function Create-DefaultSettings {
    $defaultSettings = @{
        DefaultShell = $script:Settings.DefaultShell
        DefaultShellArgs = $script:Settings.DefaultShellArgs
        DefaultRunCommandAttached = $script:Settings.DefaultRunCommandAttached
        OpenShellAtStart = $script:Settings.OpenShellAtStart
        DefaultLogsPath = $script:Settings.DefaultLogsPath
        SettingsPath = $script:Settings.SettingsPath
        FavoritesPath = $script:Settings.FavoritesPath
        ShowDebugTab = $script:Settings.ShowDebugTab
        DefaultDataFile = $script:Settings.DefaultDataFile
        CommandHistoryLimit = $script:Settings.CommandHistoryLimit
    }
    return $defaultSettings
}

# Show dialog window for settings
function Show-SettingsDialog {
    $script:UI.Overlay.Visibility = "Visible"
    $script:UI.SettingsDialog.Visibility = "Visible"

    # Populate current settings
    $script:UI.TxtDefaultShell.Text = $script:Settings.DefaultShell
    $script:UI.TxtDefaultShellArgs.Text = $script:Settings.DefaultShellArgs
    $script:UI.ChkRunCommandAttached.IsChecked = $script:Settings.DefaultRunCommandAttached
    $script:UI.ChkOpenShellAtStart.IsChecked = $script:Settings.OpenShellAtStart
    $script:UI.TxtDefaultLogsPath.Text = $script:Settings.DefaultLogsPath
    $script:UI.TxtDefaultDataFile.Text = $script:Settings.DefaultDataFile
    $script:UI.TxtCommandHistoryLimit.Text = $script:Settings.CommandHistoryLimit
    $script:UI.TxtSettingsPath.Text = $script:Settings.SettingsPath
    $script:UI.TxtFavoritesPath.Text = $script:Settings.FavoritesPath
}


function Hide-SettingsDialog {
    $script:UI.SettingsDialog.Visibility = "Hidden"
    $script:UI.Overlay.Visibility = "Collapsed"
}

function Apply-Settings {
    $script:Settings.DefaultShell = $script:UI.TxtDefaultShell.Text
    $script:Settings.DefaultShellArgs = $script:UI.TxtDefaultShellArgs.Text
    $script:Settings.DefaultRunCommandAttached = $script:UI.ChkRunCommandAttached.IsChecked
    $script:Settings.OpenShellAtStart = $script:UI.ChkOpenShellAtStart.IsChecked
    $script:Settings.DefaultLogsPath = $script:UI.TxtDefaultLogsPath.Text
    $script:Settings.DefaultDataFile = $script:UI.TxtDefaultDataFile.Text
    $script:Settings.SettingsPath = $script:UI.TxtSettingsPath.Text
    $script:Settings.FavoritesPath = $script:UI.TxtFavoritesPath.Text
    $script:Settings.ShowDebugTab = $script:UI.ChkShowDebugTab.IsChecked

    # Validate and set CommandHistoryLimit
    $historyLimit = 50  # Default value
    if ([int]::TryParse($script:UI.TxtCommandHistoryLimit.Text, [ref]$historyLimit)) {
        # Ensure it's within reasonable bounds
        if ($historyLimit -lt 1) { $historyLimit = 1 }
        if ($historyLimit -gt 1000) { $historyLimit = 1000 }
        $script:Settings.CommandHistoryLimit = $historyLimit
    } else {
        $script:Settings.CommandHistoryLimit = 50
        $script:UI.TxtCommandHistoryLimit.Text = "50"
    }

    # Apply Debug tab visibility change immediately
    if ($script:Settings.ShowDebugTab) {
        $script:UI.LogTabControl.Items[0].Visibility = "Visible"
    } else {
        $script:UI.LogTabControl.Items[0].Visibility = "Collapsed"
    }

    Save-Settings
    Hide-SettingsDialog
}

# Load settings from file
function Load-Settings {
    Ensure-SettingsFileExists
    $settings = Get-Content $script:Settings.SettingsPath | ConvertFrom-Json

    # Apply loaded settings to script variables
    $script:Settings.DefaultShell = $settings.DefaultShell
    $script:Settings.DefaultShellArgs = $settings.DefaultShellArgs
    $script:Settings.DefaultRunCommandAttached = $settings.DefaultRunCommandAttached
    $script:Settings.OpenShellAtStart = $settings.OpenShellAtStart
    $script:Settings.DefaultLogsPath = $settings.DefaultLogsPath
    $script:Settings.SettingsPath = $settings.SettingsPath
    $script:Settings.FavoritesPath = $settings.FavoritesPath

    # Handle the case where the setting might not exist in older config files
    if (Get-Member -InputObject $settings -Name "ShowDebugTab" -MemberType Properties) {
        $script:Settings.ShowDebugTab = $settings.ShowDebugTab
    }
    if (Get-Member -InputObject $settings -Name "DefaultDataFile" -MemberType Properties) {
        $script:Settings.DefaultDataFile = $settings.DefaultDataFile
    }
    if (Get-Member -InputObject $settings -Name "CommandHistoryLimit" -MemberType Properties) {
        $script:Settings.CommandHistoryLimit = $settings.CommandHistoryLimit
    }
}

# Save settings to file
function Save-Settings {
    try {
        $settingsDir = Split-Path $script:Settings.SettingsPath -Parent
        if (-not (Test-Path $settingsDir)) {
            New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
        }

        $settings = @{
            DefaultShell = $script:Settings.DefaultShell
            DefaultShellArgs = $script:Settings.DefaultShellArgs
            DefaultRunCommandAttached = $script:Settings.DefaultRunCommandAttached
            OpenShellAtStart = $script:Settings.OpenShellAtStart
            DefaultLogsPath = $script:Settings.DefaultLogsPath
            SettingsPath = $script:Settings.SettingsPath
            FavoritesPath = $script:Settings.FavoritesPath
            ShowDebugTab = $script:Settings.ShowDebugTab
            DefaultDataFile = $script:Settings.DefaultDataFile
            CommandHistoryLimit = $script:Settings.CommandHistoryLimit
        }

        $settings | ConvertTo-Json | Set-Content $script:Settings.SettingsPath
        Write-Status "Settings saved"
    }
    catch {
        Write-Status "Failed to save settings"
        Write-Log "Failed to save settings: $_"
    }
}

# Check if settings file exists and if not create it with default settings
function Ensure-SettingsFileExists {
    $settingsDir = Split-Path $script:Settings.SettingsPath -Parent
    if (-not (Test-Path $settingsDir)) {
        New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
    }
    if (-not (Test-Path $script:Settings.SettingsPath)) {
        $defaultSettings = Create-DefaultSettings
        $defaultSettings | ConvertTo-Json | Set-Content $script:Settings.SettingsPath
    }
}

function Invoke-BrowseLogs {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.SelectedPath = $script:UI.TxtDefaultLogsPath.Text
    if ($dialog.ShowDialog() -eq 'OK') {
        $script:UI.TxtDefaultLogsPath.Text = $dialog.SelectedPath
    }
}

function Invoke-BrowseSettings {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.FileName = $script:UI.TxtSettingsPath.Text
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    if ($dialog.ShowDialog()) {
        $script:UI.TxtSettingsPath.Text = $dialog.FileName
    }
}

function Invoke-BrowseFavorites {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.FileName = $script:UI.TxtFavoritesPath.Text
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    if ($dialog.ShowDialog()) {
        $script:UI.TxtFavoritesPath.Text = $dialog.FileName
    }
}

function Invoke-BrowseDataFile {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.FileName = $script:UI.TxtDefaultDataFile.Text
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    if ($dialog.ShowDialog()) {
        $script:UI.TxtDefaultDataFile.Text = $dialog.FileName
    }
}