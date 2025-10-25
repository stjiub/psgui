# Determine the current highest Id that exists in the collection
# NOTE: This function is deprecated and kept for backward compatibility only
# New code should use Get-UniqueCommandId instead
function Get-HighestId {
    param (
        [System.Object[]]$json
    )

    if (-not $json -or $json.Count -eq 0) {
        return 0
    }

    $maxId = ($json | Measure-Object -Property Id -Maximum).Maximum
    if ($maxId) {
        return $maxId
    } else {
        return 0
    }
}

# Generate a unique command ID using GUID
function Get-UniqueCommandId {
    return [System.Guid]::NewGuid().ToString()
}

# Generate a unique command list ID using GUID
function Get-CommandListId {
    return [System.Guid]::NewGuid().ToString()
}

# Copy a string to the system clipboard
function Copy-ToClipboard {
    param (
        [string]$string,
        [MaterialDesignThemes.Wpf.Snackbar]$snackbar
    )

    Write-Log "Copied to clipboard: $string"
    Write-Status "Copied to clipboard"
    Set-ClipBoard -Value $string
}

# Show a GUI error popup so important or application breaking errors can be seen
function Show-ErrorMessageBox {
    param (
        [string]$message
    )

    Write-Error $message
    [System.Windows.MessageBox]::Show($message, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}

# Show error message using multiple channels (log, status, and optional popup)
function Write-ErrorMessage {
    param (
        [string]$message,
        [switch]$ShowPopup
    )

    Write-Log "ERROR: $message"
    Write-Status "Error: $message"

    if ($ShowPopup) {
        Show-ErrorMessageBox $message
    }
}

# Show error dialog, status update, and log entry for parameter loading failures
function Show-ParameterLoadError {
    param (
        [string]$commandName,
        [string]$errorMessage
    )

    $fullMessage = "Failed to load parameters for command '$commandName':`n$errorMessage"

    # Show popup dialog using existing function
    Show-ErrorMessageBox $fullMessage

    # Update status bar
    Write-Status "Parameter load failed for '$commandName'"

    # Additional log entry (Show-ErrorMessageBox already logs via Write-Error)
    Write-Log "Parameter Load Error: $fullMessage"
}

# Write text to the LogBox
function Write-Log {
    param (
        [string]$output
    )

    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.LogBox.AppendText("$output`n")}, "Normal")
}

# Write text to the status bar
function Write-Status {
    param (
        [string]$output
    )

    $script:UI.Window.Dispatcher.Invoke([action]{
        $script:UI.StatusBox.Text = $output

        # Stop any existing timer
        if ($script:StatusTimer) {
            $script:StatusTimer.Stop()
            $script:StatusTimer = $null
        }

        # Only set a timer if the message is not "Ready"
        if ($output -ne "Ready") {
            # Create a new timer that only fires once
            $script:StatusTimer = New-Object System.Windows.Threading.DispatcherTimer
            $script:StatusTimer.Interval = [TimeSpan]::FromSeconds($script:Settings.StatusTimeout)

            # Store timer reference for the event handler
            $timer = $script:StatusTimer

            $script:StatusTimer.Add_Tick({
                $script:UI.StatusBox.Text = "Ready"
                # Stop and clean up the timer (check for null first)
                if ($timer) {
                    $timer.Stop()
                }
                $script:StatusTimer = $null
            })
            $script:StatusTimer.Start()
        }
    }, "Normal")

    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.LogBox.AppendText("$output`n")}, "Normal")
}

# Update the data file indicator text
function Update-WindowTitle {
    $unsavedIndicator = if ($script:State.HasUnsavedChanges) { "*" } else { "" }
    $script:UI.Window.Title = "$unsavedIndicator$script:AppTitle - $($script:State.CurrentDataFile)"
}

# Mark that we have unsaved changes
function Set-UnsavedChanges {
    param([bool]$hasChanges = $true)

    $script:State.HasUnsavedChanges = $hasChanges
    Update-WindowTitle
}

# Check if user wants to save before proceeding with an action
function Confirm-SaveBeforeAction {
    param([string]$actionName = "continue")

    if ($script:State.HasUnsavedChanges) {
        $result = [System.Windows.MessageBox]::Show(
            "You have unsaved changes. Do you want to save them before $actionName?",
            "Unsaved Changes",
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($result -eq [System.Windows.MessageBoxResult]::Cancel) {
            return $false
        }
        elseif ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Save-DataFile -FilePath $script:State.CurrentDataFile -Data ($script:UI.Tabs["All"].Content.ItemsSource)
            Set-UnsavedChanges $false
        }
    }
    return $true
}

# Build PowerShell arguments with profile support
function Get-PowerShellArguments {
    param (
        [string]$Command,
        [switch]$NoExit
    )

    # Build the base command with profile support
    $finalCommand = ""

    # Add profile loading if enabled
    if ($script:Settings.UseProfile -and $script:Settings.ProfilePath) {
        # Expand environment variables in the profile path
        $expandedProfilePath = [Environment]::ExpandEnvironmentVariables($script:Settings.ProfilePath)

        # Only add the profile if the file exists
        if (Test-Path $expandedProfilePath) {
            # Escape single quotes in the profile path for PowerShell
            $escapedProfilePath = $expandedProfilePath -replace "'", "''"
            # Source the profile first, then run the command
            $finalCommand = ". '$escapedProfilePath'; "
        }
    }

    # Append the user command
    if ($Command) {
        $finalCommand += $Command
    }

    # Build argument list
    $argList = "-ExecutionPolicy Bypass"

    # Add -NoExit if specified
    if ($NoExit) {
        $argList += " -NoExit"
    }

    # Add the command
    if ($finalCommand) {
        $argList += " `" & { $finalCommand } `""
    }

    return $argList
}