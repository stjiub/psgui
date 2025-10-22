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
            PreCommand = $Command.PreCommand
            ParameterSummary = (Get-ParameterSummaryFromCommand -Command $Command)
            CommandObject = $Command
            ParameterValues = $parameterValues
        }

        # Add to beginning of history list
        $script:State.CommandHistory.Insert(0, $historyEntry)

        # Trim history if it exceeds the limit
        while ($script:State.CommandHistory.Count -gt $script:Settings.CommandHistoryLimit) {
            $script:State.CommandHistory.RemoveAt($script:State.CommandHistory.Count - 1)
        }

        # Update the UI
        Update-CommandHistoryGrid

    } 
    catch {
        Write-Log "Error adding command to history: $_"
    }
}

function Get-ParameterSummaryFromCommand {
    param(
        [Command]$Command
    )

    if (-not $Command.Full) {
        return "(No parameters)"
    }

    # Extract just the parameters part from the full command
    $fullCmd = $Command.Full
    $rootCmd = $Command.Root

    if ($Command.PreCommand) {
        # Remove the PreCommand part
        $fullCmd = $fullCmd -replace [regex]::Escape($Command.PreCommand + "; "), ""
    }

    # Remove the root command to get just parameters
    $paramsPart = $fullCmd -replace [regex]::Escape($rootCmd), ""
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
        }
    } 
    catch {
        Write-Log "Error updating command history grid: $_"
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

        # Clear the command grid before rebuilding
        Clear-Grid $script:UI.CommandGrid

        # Rebuild the command grid with the parameters
        if ($command.Parameters) {
            Build-CommandGrid -Grid $script:UI.CommandGrid -Parameters $command.Parameters

            # Pre-fill the parameter values from the history
            $paramValues = $HistoryEntry.ParameterValues
            if ($paramValues) {
                foreach ($key in $paramValues.Keys) {
                    $control = $script:UI.CommandGrid.Children | Where-Object { $_.Name -eq $key }
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

        # Set as current command and show dialog
        $script:State.CurrentCommand = $command
        $script:UI.BoxCommandName.Text = $command.Root
        Show-CommandDialog

        Write-Status "Command reopened from history"

    } 
    catch {
        Write-Log "Error reopening command from history: $_"
        Write-Status "Error reopening command from history"
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

function Initialize-CommandHistoryUI {
    try {
        # Get UI elements
        $grid = $script:UI.Window.FindName("CommandHistoryGrid")
        $btnReopen = $script:UI.Window.FindName("BtnReopenHistoryCommand")
        $btnClear = $script:UI.Window.FindName("BtnClearHistory")

        # Set up double-click handler for grid
        if ($grid) {
            $grid.Add_MouseDoubleClick({
                $selectedItem = $script:UI.Window.FindName("CommandHistoryGrid").SelectedItem
                if ($selectedItem) {
                    Reopen-CommandFromHistory -HistoryEntry $selectedItem
                }
            })
        }

        # Set up button handlers
        if ($btnReopen) {
            $btnReopen.Add_Click({
                $selectedItem = $script:UI.Window.FindName("CommandHistoryGrid").SelectedItem
                if ($selectedItem) {
                    Reopen-CommandFromHistory -HistoryEntry $selectedItem
                } else {
                    Write-Status "Please select a command from history"
                }
            })
        }

        if ($btnClear) {
            $btnClear.Add_Click({
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
