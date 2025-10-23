function Add-CommandToHistory {
    param(
        [Parameter(Mandatory=$true)]
        [Command]$Command,

        [Parameter(Mandatory=$false)]
        [System.Windows.Controls.Grid]$Grid
    )

    try {
        # Extract parameter values from the grid if provided
        $parameterValues = @{}
        if ($Grid -and $Command.Parameters) {
            foreach ($param in $Command.Parameters) {
                $paramName = $param.Name.VariablePath
                $control = $Grid.Children | Where-Object { $_.Name -eq $paramName }

                if ($control) {
                    if ($control -is [System.Windows.Controls.CheckBox]) {
                        $parameterValues[$paramName] = $control.IsChecked
                    }
                    elseif ($control -is [System.Windows.Controls.ComboBox]) {
                        $parameterValues[$paramName] = $control.SelectedItem
                    }
                    elseif ($control -is [System.Windows.Controls.TextBox]) {
                        $parameterValues[$paramName] = $control.Text
                    }
                }
            }
        }

        # Create history entry
        $historyEntry = [PSCustomObject]@{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            CommandName = $Command.Root
            FullCommand = $Command.Full
            CleanCommand = $Command.CleanCommand
            PreCommand = $Command.PreCommand
            ParameterSummary = (Get-ParameterSummaryFromCommand -Command $Command)
            CommandObject = $Command
            ParameterValues = $parameterValues
            LogPath = $Command.LogPath
        }

        # Add to beginning of history list
        $script:State.CommandHistory.Insert(0, $historyEntry)

        # Trim history if it exceeds the limit
        while ($script:State.CommandHistory.Count -gt $script:Settings.CommandHistoryLimit) {
            $script:State.CommandHistory.RemoveAt($script:State.CommandHistory.Count - 1)
        }

        # Update the UI
        Update-CommandHistoryGrid

        # Return the history entry so it can be associated with tabs
        return $historyEntry

    }
    catch {
        Write-Log "Error adding command to history: $_"
        return $null
    }
}

function Get-ParameterSummaryFromCommand {
    param(
        [Command]$Command
    )

    # Use CleanCommand if available, otherwise fall back to Full
    $commandToProcess = if ($Command.CleanCommand) { $Command.CleanCommand } else { $Command.Full }

    if (-not $commandToProcess) {
        return "(No parameters)"
    }

    # Extract just the parameters part from the clean command
    $rootCmd = $Command.Root

    if ($Command.PreCommand) {
        # Remove the PreCommand part
        $commandToProcess = $commandToProcess -replace [regex]::Escape($Command.PreCommand + "; "), ""
    }

    # Remove the root command to get just parameters
    $paramsPart = $commandToProcess -replace [regex]::Escape($rootCmd), ""
    $paramsPart = $paramsPart.Trim()

    if ([string]::IsNullOrWhiteSpace($paramsPart)) {
        return "(No parameters)"
    }

    # Truncate if too long
    if ($paramsPart.Length -gt 100) {
        $paramsPart = $paramsPart.Substring(0, 97) + "..."
    }

    return $paramsPart
}

function Update-CommandHistoryGrid {
    try {
        $grid = $script:UI.Window.FindName("CommandHistoryGrid")
        if ($grid) {
            $grid.ItemsSource = $null
            $grid.ItemsSource = $script:State.CommandHistory
            Update-HistoryLogHighlighting
        }
    }
    catch {
        Write-Log "Error updating command history grid: $_"
    }
}

function Update-HistoryLogHighlighting {
    try {
        $grid = $script:UI.Window.FindName("CommandHistoryGrid")
        if (-not $grid) { return }

        # Force the grid to generate all row containers
        $grid.UpdateLayout()

        # Wait for row generation to complete
        $grid.Dispatcher.Invoke([Action]{
            foreach ($item in $script:State.CommandHistory) {
                try {
                    $row = $grid.ItemContainerGenerator.ContainerFromItem($item)
                    if ($row -is [System.Windows.Controls.DataGridRow]) {
                        # Check if this history entry has a log file
                        if ($item.LogPath -and -not [string]::IsNullOrWhiteSpace($item.LogPath)) {
                            $row.Tag = "HasLog"
                        }
                        else {
                            $row.Tag = $null
                        }
                    }
                }
                catch {
                    # Silently continue if row container is not ready
                }
            }
        }, [System.Windows.Threading.DispatcherPriority]::Background)
    }
    catch {
        Write-Log "Error updating history log highlighting: $_"
    }
}

function Reopen-CommandFromHistory {
    param(
        [Parameter(Mandatory=$true)]
        $HistoryEntry
    )

    try {
        # Get the command object from the history entry
        $command = $HistoryEntry.CommandObject

        if (-not $command) {
            Write-Log "No command data found in history entry"
            Write-Status "Error: No command data in history"
            return
        }

        # If SkipParameterSelect is true, just rerun the command
        if ($command.SkipParameterSelect) {
            Run-Command $command $script:State.RunCommandAttached
            Write-Status "Command rerun from history"
            return
        }

        # Create a new CommandWindow
        $commandWindow = New-CommandWindow -Command $command

        if ($commandWindow) {
            # Rebuild the command grid with the parameters
            if ($command.Parameters) {
                Build-CommandGrid -CommandWindow $commandWindow -Parameters $command.Parameters

                # Pre-fill the parameter values from the history
                $paramValues = $HistoryEntry.ParameterValues
                if ($paramValues) {
                    foreach ($key in $paramValues.Keys) {
                        $control = $commandWindow.CommandGrid.Children | Where-Object { $_.Name -eq $key }
                        if ($control) {
                            if ($control -is [System.Windows.Controls.CheckBox]) {
                                $control.IsChecked = $paramValues[$key]
                            }
                            elseif ($control -is [System.Windows.Controls.ComboBox]) {
                                $control.SelectedItem = $paramValues[$key]
                            }
                            elseif ($control -is [System.Windows.Controls.TextBox]) {
                                $control.Text = $paramValues[$key]
                            }
                        }
                    }
                }
            }

            # Show the window
            $commandWindow.Window.Show()
        }

        Write-Status "Command reopened from history"

    }
    catch {
        Write-Log "Error reopening command from history: $_"
        Write-Status "Error reopening command from history"
    }
}

function Remove-CommandFromHistory {
    param(
        [Parameter(Mandatory=$true)]
        $SelectedItems
    )

    try {
        if (-not $SelectedItems -or $SelectedItems.Count -eq 0) {
            Write-Status "No commands selected"
            return
        }

        $count = $SelectedItems.Count
        $message = if ($count -eq 1) {
            "Are you sure you want to remove this command from history?"
        } else {
            "Are you sure you want to remove $count commands from history?"
        }

        $result = [System.Windows.MessageBox]::Show(
            $message,
            "Remove From History",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            foreach ($item in $SelectedItems) {
                $script:State.CommandHistory.Remove($item)
            }
            Update-CommandHistoryGrid
            Write-Status "Removed $count command(s) from history"
        }
    }
    catch {
        Write-Log "Error removing command from history: $_"
        Write-Status "Error removing command from history"
    }
}

function Clear-CommandHistory {
    try {
        $result = [System.Windows.MessageBox]::Show(
            "Are you sure you want to clear the command history?",
            "Clear Command History",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $script:State.CommandHistory.Clear()
            Update-CommandHistoryGrid
            Write-Status "Command history cleared"
        }
    }
    catch {
        Write-Log "Error clearing command history: $_"
    }
}

function Copy-HistoryCommandToClipboard {
    param(
        [Parameter(Mandatory=$true)]
        $HistoryEntry
    )

    try {
        $command = $HistoryEntry.CommandObject
        if ($command -and $command.CleanCommand) {
            Copy-ToClipboard -String $command.CleanCommand
            Write-Status "Command copied to clipboard"
        }
        else {
            Write-Status "No command to copy"
        }
    }
    catch {
        Write-Log "Error copying command to clipboard: $_"
        Write-Status "Error copying command to clipboard"
    }
}

function Open-HistoryCommandLog {
    param(
        [Parameter(Mandatory=$true)]
        $HistoryEntry
    )

    try {
        if ($HistoryEntry.LogPath -and -not [string]::IsNullOrWhiteSpace($HistoryEntry.LogPath)) {
            if (Test-Path $HistoryEntry.LogPath) {
                New-LogMonitorTab -FilePath $HistoryEntry.LogPath -TabControl $script:UI.LogTabControl

                # Switch to the Logs tab
                $logsTab = $script:UI.TabControlShell.Items | Where-Object { $_.Header -eq "Logs" }
                if ($logsTab) {
                    $script:UI.TabControlShell.SelectedItem = $logsTab
                }

                Write-Status "Opened log file"
            }
            else {
                Write-ErrorMessage "Log file not found at: $($HistoryEntry.LogPath)"
            }
        }
        else {
            Write-ErrorMessage "No log file associated with this command"
        }
    }
    catch {
        Write-ErrorMessage "Failed to open log file: $_"
    }
}

function Initialize-CommandHistoryUI {
    try {
        # Get UI elements
        $grid = $script:UI.Window.FindName("CommandHistoryGrid")
        $menuRerun = $script:UI.Window.FindName("MenuHistoryRerun")
        $menuCopy = $script:UI.Window.FindName("MenuHistoryCopyToClipboard")
        $menuOpenLog = $script:UI.Window.FindName("MenuHistoryOpenLog")
        $menuRemove = $script:UI.Window.FindName("MenuHistoryRemove")
        $menuClear = $script:UI.Window.FindName("MenuHistoryClear")

        # Set up double-click handler for grid
        if ($grid) {
            $grid.Add_MouseDoubleClick({
                $selectedItem = $script:UI.Window.FindName("CommandHistoryGrid").SelectedItem
                if ($selectedItem) {
                    Reopen-CommandFromHistory -HistoryEntry $selectedItem
                }
            })
        }

        # Set up context menu handlers
        if ($menuRerun) {
            $menuRerun.Add_Click({
                $selectedItem = $script:UI.Window.FindName("CommandHistoryGrid").SelectedItem
                if ($selectedItem) {
                    Reopen-CommandFromHistory -HistoryEntry $selectedItem
                } else {
                    Write-Status "Please select a command from history"
                }
            })
        }

        if ($menuCopy) {
            $menuCopy.Add_Click({
                $selectedItem = $script:UI.Window.FindName("CommandHistoryGrid").SelectedItem
                if ($selectedItem) {
                    Copy-HistoryCommandToClipboard -HistoryEntry $selectedItem
                } else {
                    Write-Status "Please select a command from history"
                }
            })
        }

        if ($menuOpenLog) {
            $menuOpenLog.Add_Click({
                $selectedItem = $script:UI.Window.FindName("CommandHistoryGrid").SelectedItem
                if ($selectedItem) {
                    Open-HistoryCommandLog -HistoryEntry $selectedItem
                } else {
                    Write-Status "Please select a command from history"
                }
            })
        }

        if ($menuRemove) {
            $menuRemove.Add_Click({
                $selectedItems = $script:UI.Window.FindName("CommandHistoryGrid").SelectedItems
                if ($selectedItems -and $selectedItems.Count -gt 0) {
                    # Convert to array to avoid modification during iteration
                    $itemsArray = @($selectedItems)
                    Remove-CommandFromHistory -SelectedItems $itemsArray
                } else {
                    Write-Status "Please select command(s) from history"
                }
            })
        }

        if ($menuClear) {
            $menuClear.Add_Click({
                Clear-CommandHistory
            })
        }

        # Initialize grid
        Update-CommandHistoryGrid

    }
    catch {
        Write-Log "Error initializing command history UI: $_"
    }
}
