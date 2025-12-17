# Load default settings from defaultsettings.json
function Load-DefaultSettings {
    try {
        if (Test-Path $script:ApplicationPaths.DefaultSettingsFile) {
            $defaultSettings = Get-Content $script:ApplicationPaths.DefaultSettingsFile | ConvertFrom-Json

            # Expand environment variables in paths
            $expandedSettings = @{}
            foreach ($key in $defaultSettings.PSObject.Properties.Name) {
                $value = $defaultSettings.$key
                if ($value -is [string] -and $value -match '%\w+%') {
                    $expandedSettings[$key] = [Environment]::ExpandEnvironmentVariables($value)
                } else {
                    $expandedSettings[$key] = $value
                }
            }

            return $expandedSettings
        }
    }
    catch {
        Write-Log "Failed to load default settings: $_"
    }
    return $null
}

# Set app settings from loaded settings
function Initialize-Settings {
    Load-Settings

    # Ensure logs directory exists
    if (-not (Test-Path $script:Settings.DefaultLogsPath)) {
        try {
            New-Item -ItemType Directory -Path $script:Settings.DefaultLogsPath -Force | Out-Null
        }
        catch {
            # Silently continue if directory creation fails during init
        }
    }

    # Update UI elements with loaded settings
    $script:UI.TxtDefaultShell.Text = $script:Settings.DefaultShell
    $script:UI.TxtDefaultShellArgs.Text = $script:Settings.DefaultShellArgs
    $script:UI.ChkRunCommandAttached.IsChecked = $script:Settings.DefaultRunCommandAttached
    $script:UI.ChkOpenShellAtStart.IsChecked = $script:Settings.OpenShellAtStart
    $script:UI.TxtDefaultLogsPath.Text = $script:Settings.DefaultLogsPath
    $script:UI.TxtDefaultDataFile.Text = $script:Settings.DefaultDataFile
    $script:UI.TxtCommandHistoryLimit.Text = $script:Settings.CommandHistoryLimit
    $script:UI.TxtStatusTimeout.Text = $script:Settings.StatusTimeout
    $script:UI.ChkShowDebugTab.IsChecked = $script:Settings.ShowDebugTab
    $script:UI.ChkUseProfile.IsChecked = $script:Settings.UseProfile
    $script:UI.TxtProfilePath.Text = $script:Settings.ProfilePath
    $script:UI.ChkEnablePSTask.IsChecked = $script:Settings.EnablePSTask

    # Set the Debug tab visibility based on setting
    $debugTab = $script:UI.Window.FindName("TabControlShell").Items | Where-Object { $_.Header -eq "Debug" }
    if ($debugTab) {
        if ($script:Settings.ShowDebugTab) {
            $debugTab.Visibility = "Visible"
        } else {
            $debugTab.Visibility = "Collapsed"
        }
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
        DefaultHistoryPath = $script:Settings.DefaultHistoryPath
        ShowDebugTab = $script:Settings.ShowDebugTab
        DefaultDataFile = $script:Settings.DefaultDataFile
        CommandHistoryLimit = $script:Settings.CommandHistoryLimit
        StatusTimeout = $script:Settings.StatusTimeout
        SaveHistory = $script:Settings.SaveHistory
        UseProfile = $script:Settings.UseProfile
        ProfilePath = $script:Settings.ProfilePath
        EnablePSTask = $script:Settings.EnablePSTask
    }
    return $defaultSettings
}

# Show dialog window for settings
function Show-SettingsDialog {
    # Save the current shell visibility state and hide it to prevent embedded windows from appearing over the dialog
    if ($script:UI.Shell -and $script:UI.Shell.Visibility -eq "Visible") {
        $script:State.ShellVisibleBeforeDialog = $true
        Toggle-ShellGrid
    } else {
        $script:State.ShellVisibleBeforeDialog = $false
    }

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
    $script:UI.TxtStatusTimeout.Text = $script:Settings.StatusTimeout
    $script:UI.TxtSettingsPath.Text = $script:Settings.SettingsPath
    $script:UI.TxtFavoritesPath.Text = $script:Settings.FavoritesPath
    $script:UI.TxtDefaultHistoryPath.Text = $script:Settings.DefaultHistoryPath
    $script:UI.ChkSaveHistory.IsChecked = $script:Settings.SaveHistory
    $script:UI.ChkUseProfile.IsChecked = $script:Settings.UseProfile
    $script:UI.TxtProfilePath.Text = $script:Settings.ProfilePath
    $script:UI.ChkEnablePSTask.IsChecked = $script:Settings.EnablePSTask
}

# Hide the settings dialog grid
function Hide-SettingsDialog {
    $script:UI.SettingsDialog.Visibility = "Hidden"
    $script:UI.Overlay.Visibility = "Collapsed"

    # Restore the shell to its previous state
    if ($script:State.ShellVisibleBeforeDialog -and $script:UI.Shell.Visibility -eq "Collapsed") {
        Toggle-ShellGrid
    }
}

# Apply any settings changes to application
function Apply-Settings {
    # Store the old default data file path to check if it changed
    $oldDefaultDataFile = $script:Settings.DefaultDataFile
    $newDefaultDataFile = $script:UI.TxtDefaultDataFile.Text

    # Check if EnablePSTask is being turned on
    $oldEnablePSTask = $script:Settings.EnablePSTask
    $newEnablePSTask = $script:UI.ChkEnablePSTask.IsChecked

    if ($newEnablePSTask -and -not $oldEnablePSTask) {
        # Check if PSTask module is installed
        if (-not (Test-PSTaskModule)) {
            $result = [System.Windows.MessageBox]::Show(
                "PSTask module is not installed. Please install it first using:`n`nInstall-Module -Name PSTask`n`nThen restart the application. Would you like to cancel enabling PSTask?",
                "PSTask Module Not Found",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )

            if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                $script:UI.ChkEnablePSTask.IsChecked = $false
                return
            }
        }
    }

    $script:Settings.DefaultShell = $script:UI.TxtDefaultShell.Text
    $script:Settings.DefaultShellArgs = $script:UI.TxtDefaultShellArgs.Text
    $script:Settings.DefaultRunCommandAttached = $script:UI.ChkRunCommandAttached.IsChecked
    $script:Settings.OpenShellAtStart = $script:UI.ChkOpenShellAtStart.IsChecked
    $script:Settings.DefaultLogsPath = $script:UI.TxtDefaultLogsPath.Text
    $script:Settings.DefaultDataFile = $newDefaultDataFile
    $script:Settings.SettingsPath = $script:UI.TxtSettingsPath.Text
    $script:Settings.FavoritesPath = $script:UI.TxtFavoritesPath.Text
    $script:Settings.DefaultHistoryPath = $script:UI.TxtDefaultHistoryPath.Text
    $script:Settings.SaveHistory = $script:UI.ChkSaveHistory.IsChecked
    $script:Settings.ShowDebugTab = $script:UI.ChkShowDebugTab.IsChecked
    $script:Settings.UseProfile = $script:UI.ChkUseProfile.IsChecked
    $script:Settings.ProfilePath = $script:UI.TxtProfilePath.Text
    $script:Settings.EnablePSTask = $newEnablePSTask

    # Validate and set CommandHistoryLimit
    $historyLimit = 50
    if ([int]::TryParse($script:UI.TxtCommandHistoryLimit.Text, [ref]$historyLimit)) {
        # Ensure it's within reasonable bounds
        if ($historyLimit -lt 1) { $historyLimit = 1 }
        if ($historyLimit -gt 1000) { $historyLimit = 1000 }
        $script:Settings.CommandHistoryLimit = $historyLimit
    } else {
        $script:Settings.CommandHistoryLimit = 50
        $script:UI.TxtCommandHistoryLimit.Text = "50"
    }

    # Validate and set StatusTimeout
    $statusTimeout = 6
    if ([int]::TryParse($script:UI.TxtStatusTimeout.Text, [ref]$statusTimeout)) {
        # Ensure it's within reasonable bounds
        if ($statusTimeout -lt 1) { $statusTimeout = 1 }
        if ($statusTimeout -gt 60) { $statusTimeout = 60 }
        $script:Settings.StatusTimeout = $statusTimeout
    } else {
        $script:Settings.StatusTimeout = 6
        $script:UI.TxtStatusTimeout.Text = "6"
    }

    # Apply Debug tab visibility change immediately
    $debugTab = $script:UI.Window.FindName("TabControlShell").Items | Where-Object { $_.Header -eq "Debug" }
    if ($debugTab) {
        if ($script:Settings.ShowDebugTab) {
            $debugTab.Visibility = "Visible"
        } else {
            $debugTab.Visibility = "Collapsed"
        }
    }

    Save-Settings
    Hide-SettingsDialog

    # Check if default data file changed and is different from currently opened file
    if ($oldDefaultDataFile -ne $newDefaultDataFile -and $script:State.CurrentDataFile -ne $newDefaultDataFile) {
        # Ask user if they want to open the new default data file
        $result = [System.Windows.MessageBox]::Show(
            "The default data file has been changed to:`n$newDefaultDataFile`n`nWould you like to open this file now?",
            "Open New Default Data File?",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            # Check for unsaved changes before opening
            if (Confirm-SaveBeforeAction "opening the new default data file") {
                $script:State.CurrentDataFile = $newDefaultDataFile
                Initialize-DataFile $newDefaultDataFile
                Load-NewDataFile -FilePath $newDefaultDataFile
                Update-WindowTitle
                Write-Status "Opened new default data file: $newDefaultDataFile"
            }
        }
    }
}

# Load settings from file with fallback to default settings
function Load-Settings {
    # Load default settings first
    $defaultSettings = Load-DefaultSettings

    Ensure-SettingsFileExists
    $settings = Get-Content $script:Settings.SettingsPath | ConvertFrom-Json

    # Helper function to get setting value with fallback to default
    function Get-SettingValue {
        param($settingsObj, $propertyName, $defaultValue)
        if (Get-Member -InputObject $settingsObj -Name $propertyName -MemberType Properties) {
            return $settingsObj.$propertyName
        }
        return $defaultValue
    }

    # Apply loaded settings to script variables with fallback to defaults
    $script:Settings.DefaultShell = Get-SettingValue $settings "DefaultShell" $defaultSettings.DefaultShell
    $script:Settings.DefaultShellArgs = Get-SettingValue $settings "DefaultShellArgs" $defaultSettings.DefaultShellArgs
    $script:Settings.DefaultRunCommandAttached = Get-SettingValue $settings "DefaultRunCommandAttached" $defaultSettings.DefaultRunCommandAttached
    $script:Settings.OpenShellAtStart = Get-SettingValue $settings "OpenShellAtStart" $defaultSettings.OpenShellAtStart
    $script:Settings.DefaultLogsPath = Get-SettingValue $settings "DefaultLogsPath" $defaultSettings.DefaultLogsPath
    $script:Settings.SettingsPath = Get-SettingValue $settings "SettingsPath" $defaultSettings.SettingsPath
    $script:Settings.FavoritesPath = Get-SettingValue $settings "FavoritesPath" $defaultSettings.FavoritesPath
    $script:Settings.DefaultHistoryPath = Get-SettingValue $settings "DefaultHistoryPath" $defaultSettings.DefaultHistoryPath
    $script:Settings.ShowDebugTab = Get-SettingValue $settings "ShowDebugTab" $defaultSettings.ShowDebugTab
    $script:Settings.DefaultDataFile = Get-SettingValue $settings "DefaultDataFile" $defaultSettings.DefaultDataFile
    $script:Settings.CommandHistoryLimit = Get-SettingValue $settings "CommandHistoryLimit" $defaultSettings.CommandHistoryLimit
    $script:Settings.StatusTimeout = Get-SettingValue $settings "StatusTimeout" $defaultSettings.StatusTimeout
    $script:Settings.SaveHistory = Get-SettingValue $settings "SaveHistory" $defaultSettings.SaveHistory
    $script:Settings.UseProfile = Get-SettingValue $settings "UseProfile" $defaultSettings.UseProfile
    $script:Settings.ProfilePath = Get-SettingValue $settings "ProfilePath" $defaultSettings.ProfilePath
    $script:Settings.EnablePSTask = Get-SettingValue $settings "EnablePSTask" $defaultSettings.EnablePSTask
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
            DefaultHistoryPath = $script:Settings.DefaultHistoryPath
            ShowDebugTab = $script:Settings.ShowDebugTab
            DefaultDataFile = $script:Settings.DefaultDataFile
            CommandHistoryLimit = $script:Settings.CommandHistoryLimit
            StatusTimeout = $script:Settings.StatusTimeout
            SaveHistory = $script:Settings.SaveHistory
            UseProfile = $script:Settings.UseProfile
            ProfilePath = $script:Settings.ProfilePath
            EnablePSTask = $script:Settings.EnablePSTask
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

# Generic function to browse for a file or folder
function Invoke-BrowsePath {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.TextBox]$TextBox,

        [Parameter(Mandatory=$false)]
        [string]$Filter = "All files (*.*)|*.*",

        [Parameter(Mandatory=$false)]
        [string]$DefaultExt = "",

        [Parameter(Mandatory=$false)]
        [switch]$IsFolder
    )

    if ($IsFolder) {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.SelectedPath = $TextBox.Text
        if ($dialog.ShowDialog() -eq 'OK') {
            $TextBox.Text = $dialog.SelectedPath
        }
    }
    else {
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.FileName = $TextBox.Text
        $dialog.Filter = $Filter
        if ($DefaultExt) {
            $dialog.DefaultExt = $DefaultExt
        }
        if ($dialog.ShowDialog()) {
            $TextBox.Text = $dialog.FileName
        }
    }
}

# Test if PSTask module is installed
function Test-PSTaskModule {
    try {
        $module = Get-Module -ListAvailable -Name PSTask
        return $null -ne $module
    }
    catch {
        return $false
    }
}