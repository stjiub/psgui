# Define the RowData object. This is the object that is used on all the Main window tabitem grids
class RowData {
    [int]$Id
    [string]$Name
    [string]$Description
    [string]$Category
    [string]$Command
    [bool]$SkipParameterSelect
    [string]$PreCommand
    [bool]$Log
}

class FavoriteRowData : RowData {
    [int]$Order

    FavoriteRowData([RowData]$rowData, [int]$order) {
        $this.Id = $rowData.Id
        $this.Name = $rowData.Name
        $this.Description = $rowData.Description
        $this.Category = $rowData.Category
        $this.Command = $rowData.Command
        $this.SkipParameterSelect = $rowData.SkipParameterSelect
        $this.PreCommand = $rowData.PreCommand
        $this.Log = $rowData.Log
        $this.Order = $order
    }
}

# Define the Command object. This is used by the CommandWindow to construct the grid and run the command
class Command {
    [string]$Root
    [string]$Full
    [string]$CleanCommand
    [string]$PreCommand
    [System.Object[]]$Parameters
    [bool]$SkipParameterSelect
    [bool]$Log
    [string]$LogPath
}
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

        # Save history to file if enabled
        Save-CommandHistory

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
            Save-CommandHistory
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
            Save-CommandHistory
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

function Save-CommandHistory {
    try {
        if (-not $script:Settings.SaveHistory) {
            Write-Log "SaveHistory is disabled, skipping save"
            return
        }

        $historyPath = $script:Settings.DefaultHistoryPath
        Write-Log "Attempting to save command history to: $historyPath"

        if ([string]::IsNullOrWhiteSpace($historyPath)) {
            Write-Log "ERROR: DefaultHistoryPath is null or empty"
            return
        }

        # Ensure directory exists
        $historyDir = Split-Path -Parent $historyPath
        if (-not (Test-Path $historyDir)) {
            Write-Log "Creating history directory: $historyDir"
            New-Item -ItemType Directory -Path $historyDir -Force | Out-Null
        }

        # Convert history entries to a serializable format
        $serializedHistory = @()
        foreach ($entry in $script:State.CommandHistory) {
            if (-not $entry.CommandObject) {
                Write-Log "WARNING: Skipping history entry with null CommandObject"
                continue
            }

            # Convert ParameterValues hashtable to a simple object for JSON serialization
            $paramValuesObj = $null
            if ($entry.ParameterValues) {
                $paramValuesObj = @{}
                foreach ($key in $entry.ParameterValues.Keys) {
                    # Convert key to string to ensure JSON compatibility
                    $paramValuesObj[$key.ToString()] = $entry.ParameterValues[$key]
                }
            }

            $serializedEntry = @{
                Timestamp = $entry.Timestamp
                CommandName = $entry.CommandName
                FullCommand = $entry.FullCommand
                CleanCommand = $entry.CleanCommand
                PreCommand = $entry.PreCommand
                ParameterSummary = $entry.ParameterSummary
                LogPath = $entry.LogPath
                # Store command properties separately since Command objects don't serialize well
                CommandData = @{
                    Root = $entry.CommandObject.Root
                    Full = $entry.CommandObject.Full
                    CleanCommand = $entry.CommandObject.CleanCommand
                    PreCommand = $entry.CommandObject.PreCommand
                    SkipParameterSelect = $entry.CommandObject.SkipParameterSelect
                    Log = $entry.CommandObject.Log
                    LogPath = $entry.CommandObject.LogPath
                }
                ParameterValues = $paramValuesObj
            }
            $serializedHistory += $serializedEntry
        }

        Write-Log "Serializing $($serializedHistory.Count) history entries"

        # Save to JSON file
        $serializedHistory | ConvertTo-Json -Depth 10 | Set-Content -Path $historyPath -Encoding UTF8
        Write-Log "Command history saved successfully to $historyPath"
    }
    catch {
        Write-Log "ERROR saving command history: $_"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
    }
}

function Load-CommandHistory {
    try {
        if (-not $script:Settings.SaveHistory) {
            Write-Log "SaveHistory is disabled, skipping load"
            return
        }

        $historyPath = $script:Settings.DefaultHistoryPath
        Write-Log "Attempting to load command history from: $historyPath"

        if ([string]::IsNullOrWhiteSpace($historyPath)) {
            Write-Log "ERROR: DefaultHistoryPath is null or empty"
            return
        }

        if (-not (Test-Path $historyPath)) {
            Write-Log "No command history file found at $historyPath (this is normal on first run)"
            return
        }

        # Load history from JSON file
        $serializedHistory = Get-Content -Path $historyPath -Raw | ConvertFrom-Json

        if (-not $serializedHistory) {
            Write-Log "History file is empty or invalid"
            return
        }

        # Convert back to history entry objects
        $script:State.CommandHistory.Clear()

        foreach ($serialized in $serializedHistory) {
            # Recreate Command object
            $command = New-Object Command
            $command.Root = $serialized.CommandData.Root
            $command.Full = $serialized.CommandData.Full
            $command.CleanCommand = $serialized.CommandData.CleanCommand
            $command.PreCommand = $serialized.CommandData.PreCommand
            $command.SkipParameterSelect = $serialized.CommandData.SkipParameterSelect
            $command.Log = $serialized.CommandData.Log
            $command.LogPath = $serialized.CommandData.LogPath

            # Recreate history entry
            $historyEntry = [PSCustomObject]@{
                Timestamp = $serialized.Timestamp
                CommandName = $serialized.CommandName
                FullCommand = $serialized.FullCommand
                CleanCommand = $serialized.CleanCommand
                PreCommand = $serialized.PreCommand
                ParameterSummary = $serialized.ParameterSummary
                CommandObject = $command
                ParameterValues = $serialized.ParameterValues
                LogPath = $serialized.LogPath
            }

            $script:State.CommandHistory.Add($historyEntry)
        }

        # Trim to history limit
        while ($script:State.CommandHistory.Count -gt $script:Settings.CommandHistoryLimit) {
            $script:State.CommandHistory.RemoveAt($script:State.CommandHistory.Count - 1)
        }

        Update-CommandHistoryGrid
        Write-Log "Loaded $($script:State.CommandHistory.Count) command(s) from history successfully"
    }
    catch {
        Write-Log "ERROR loading command history: $_"
        Write-Log "Stack trace: $($_.ScriptStackTrace)"
    }
}
# Handle the Main Window Add Button click event to add a new RowData object to the collection
function Add-CommandRow {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $newRow = New-Object RowData
    $newRow.Id = ++$script:State.HighestId
    $tab = $tabs["All"]
    $grid = $tab.Content
    $grid.ItemsSource.Add($newRow)
    Set-UnsavedChanges $true
    $tabControl.SelectedItem = $tab
    # We don't want to change the tabs read only status if they are already in edit mode
    if ($script:State.TabsReadOnly) {
        Set-TabsReadOnlyStatus -Tabs $tabs
        Set-TabsExtraColumnsVisibility -Tabs $tabs
    }
    # Select the new row and set it as the current item
    $grid.SelectedItem = $newRow
    $grid.CurrentItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.Focus()
    # Update the layout to ensure the selection is processed
    $grid.UpdateLayout()
    # Set the current cell to the Name column of the new row
    $nameColumn = $grid.Columns | Where-Object { $_.Header -eq "Name" } | Select-Object -First 1
    if ($nameColumn) {
        $grid.CurrentCell = New-Object System.Windows.Controls.DataGridCellInfo($newRow, $nameColumn)
    }
    $grid.BeginEdit()
}

# Handle the Duplicate Command to create a copy of the selected command row
function Duplicate-CommandRow {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $grid = $tabControl.SelectedItem.Content
    $selectedItem = $grid.SelectedItem

    if (-not $selectedItem) {
        Write-Status "No command selected to duplicate"
        return
    }

    # Create a new row with a new ID
    $newRow = New-Object RowData
    $newRow.Id = ++$script:State.HighestId

    # Copy all properties except Id
    $newRow.Name = $selectedItem.Name
    $newRow.Description = $selectedItem.Description
    $newRow.Category = $selectedItem.Category
    $newRow.Command = $selectedItem.Command
    $newRow.SkipParameterSelect = $selectedItem.SkipParameterSelect
    $newRow.PreCommand = $selectedItem.PreCommand

    # Add to All tab
    $allTab = $tabs["All"]
    $allGrid = $allTab.Content
    $allGrid.ItemsSource.Add($newRow)

    # If the item has a category, add to category tab as well
    if ($newRow.Category) {
        $categoryTab = $tabs[$newRow.Category]
        if ($categoryTab) {
            $categoryTab.Content.ItemsSource.Add($newRow)
        }
    }

    Set-UnsavedChanges $true

    # We don't want to change the tabs read only status if they are already in edit mode
    if ($script:State.TabsReadOnly) {
        Set-TabsReadOnlyStatus -Tabs $tabs
        Set-TabsExtraColumnsVisibility -Tabs $tabs
    }

    # Select the new row and set it as the current item
    $grid.SelectedItem = $newRow
    $grid.CurrentItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.Focus()
    $grid.UpdateLayout()

    Write-Status "Command duplicated"
}

# Handle the Main Window Remove Button click event to remove one or multiple RowData objects from the collection
function Remove-CommandRow {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $allGrid = $tabs["All"].Content
    $allData = $allGrid.ItemsSource
    $grid = $tabControl.SelectedItem.Content

    # We want to make a copy of the selected items to avoid issues
    # with the collection being modified while still enumerating
    $selectedItems = @()
    foreach ($item in $grid.SelectedItems) {
        $selectedItems += $item
    }

    # Show confirmation dialog if there are items to delete
    if ($selectedItems.Count -gt 0) {
        $itemText = if ($selectedItems.Count -eq 1) { "command" } else { "commands" }
        $message = "Are you sure you want to delete the selected $($selectedItems.Count) $($itemText)?"
        $result = [System.Windows.MessageBox]::Show($message, "Confirm Delete", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }
    }

    # Create a snapshot of the deleted items for the recycle bin
    $deletedBatch = @{
        Timestamp = Get-Date
        Items = @()
    }

    foreach ($item in $selectedItems) {
        $id = $item.Id

        # Create a deep copy of the item for the recycle bin
        $itemCopy = New-Object RowData
        $itemCopy.Id = $item.Id
        $itemCopy.Name = $item.Name
        $itemCopy.Description = $item.Description
        $itemCopy.Category = $item.Category
        $itemCopy.Command = $item.Command
        $itemCopy.SkipParameterSelect = $item.SkipParameterSelect
        $itemCopy.PreCommand = $item.PreCommand

        $deletedBatch.Items += $itemCopy

        # If item has a category then remove from category's tab and remove the tab
        # if it was the only item of that category
        $category = $item.Category
        if ($category) {
            $categoryGrid = $tabs[$category].Content
            $categoryData = $categoryGrid.ItemsSource
            $categoryIndex = Get-GridIndexOfId -Grid $categoryGrid -Id $id
            $categoryData.RemoveAt($categoryIndex)
            if ($categoryData.Count -eq 0) {
                $tabControl.Items.Remove($tabs[$category])
                $tabs.Remove($category)
            }
        }
        $allIndex = Get-GridIndexOfId -Grid $allGrid -Id $Id
        $allData.RemoveAt($allIndex)
    }

    if ($selectedItems.Count -gt 0) {
        # Add deleted items to recycle bin
        $script:State.RecycleBin.Enqueue($deletedBatch)

        # Maintain max size by removing oldest items
        while ($script:State.RecycleBin.Count -gt $script:State.RecycleBinMaxSize) {
            [void]$script:State.RecycleBin.Dequeue()
        }

        Set-UnsavedChanges $true
        $itemText = if ($selectedItems.Count -eq 1) { "command" } else { "commands" }
        Write-Status "Deleted $($selectedItems.Count) $($itemText) (can be restored with Undo Delete)"
    }
}

# Restore the last deleted command(s) from the recycle bin
function Restore-DeletedCommand {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    if ($script:State.RecycleBin.Count -eq 0) {
        Write-Status "No deleted commands to restore"
        [System.Windows.MessageBox]::Show("The recycle bin is empty. There are no deleted commands to restore.", "Recycle Bin Empty", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    # Get the most recent deletion batch
    $deletedBatch = $script:State.RecycleBin.Dequeue()
    $restoredCount = 0

    $allTab = $tabs["All"]
    $allGrid = $allTab.Content
    $allData = $allGrid.ItemsSource

    foreach ($item in $deletedBatch.Items) {
        # Check if an item with this ID already exists (to prevent duplicates)
        $existingIndex = Get-GridIndexOfId -Grid $allGrid -Id $item.Id
        if ($existingIndex -ge 0) {
            Write-Log "Skipping restore of command ID $($item.Id) - already exists"
            continue
        }

        # Add to All tab
        $allData.Add($item)

        # If the item has a category, add to category tab (create if needed)
        if ($item.Category) {
            $categoryTab = $tabs[$item.Category]
            if (-not $categoryTab) {
                # Create new category tab
                $itemsSource = New-Object System.Collections.ObjectModel.ObservableCollection[RowData]
                $categoryTab = New-DataTab -Name $item.Category -ItemsSource $itemsSource -TabControl $tabControl
                $tabs.Add($item.Category, $categoryTab)

                # Assign event handlers to the new tab
                $categoryTab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
                $categoryTab.Content.Add_PreviewKeyDown({
                    param($sender,$e)
                    if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                        Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
                    }
                    elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
                        # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
                        if (-not $script:State.TabsReadOnly) {
                            $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                            $e.Handled = $true
                        }
                        # If Edit Mode is disabled (TabsReadOnly is true), run the command
                        else {
                            $e.Handled = $true
                            Invoke-MainRunClick -TabControl $script:UI.TabControl
                        }
                    }
                })
                Sort-TabControl -TabControl $tabControl
            }
            $categoryTab.Content.ItemsSource.Add($item)
        }

        $restoredCount++
    }

    if ($restoredCount -gt 0) {
        Set-UnsavedChanges $true
        Update-FavoriteHighlighting
        $itemText = if ($restoredCount -eq 1) { "command" } else { "commands" }
        Write-Status "Restored $restoredCount $($itemText)"
    }
}

# Handle the Main Edit Button click event to enable or disable editing of the grids
function Toggle-EditMode {
    param (
        [hashtable]$tabs
    )

    # Commit any pending edits before toggling columns
    foreach ($tab in $tabs.GetEnumerator()) {
        $grid = $tab.Value.Content
        if ($grid) {
            # End any edit in progress
            $grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
            # Update the layout to ensure changes are processed
            $grid.UpdateLayout()
        }
    }

    Set-TabsReadOnlyStatus -Tabs $tabs
    Set-TabsExtraColumnsVisibility -Tabs $tabs
}

function Toggle-CommandFavorite {
    $selectedTab = $script:UI.TabControl.SelectedItem
    $grid = $selectedTab.Content
    $selectedItem = $grid.SelectedItem

    if ($selectedItem) {
        try {
            $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
            $existingFavorite = $favorites | Where-Object { $_.Id -eq $selectedItem.Id }

            if ($existingFavorite) {
                [void]$favorites.Remove($existingFavorite)
                Save-Favorites -Favorites $favorites
                Update-FavoriteHighlighting
                Write-Status "Removed from favorites"
            }
            else {
                $script:State.FavoritesHighestOrder++
                $favoriteRow = [FavoriteRowData]::new($selectedItem, $script:State.FavoritesHighestOrder)
                [void]$favorites.Add($favoriteRow)
                Save-Favorites -Favorites $favorites
                Update-FavoriteHighlighting
                Write-Status "Added to favorites"
            }
        }
        catch {
            Write-Status "Failed to add/remove favorite"
            Write-Log "Failed to add/remove favorite: $_"
        }
    }
}

# Handle the Cell Edit ending event to make sure all tabs are updated properly for cell changes
function Invoke-CellEditEndingHandler {
    param (
        $sender,
        $e,
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $editedObject = $e.Row.Item
    $category = $editedObject.Category
    $columnHeader = $e.Column.Header
    $id = $editedObject.Id

    # If there is no category yet then there is nothing to sync between tabs unless we are editing Category
    if ((-not $category) -and ($columnHeader -ne "Category")) {
        return
    }
    elseif (-not $category) {
        $newObject = $true
    }

    $allGrid = $tabs["All"].Content
    $allData = $allGrid.ItemsSource
    $allIndex = Get-GridIndexOfId -Grid $allGrid -Id $id

    # Sync only the changed property between tabs
    if (-not $newObject) {
        $categoryGrid = $tabs[$category].Content
        $categoryData = $categoryGrid.ItemsSource
        $categoryIndex = Get-GridIndexOfId -Grid $categoryGrid -Id $id
        
        # Only update the specific property that was edited, preserve other values
        if ($categoryIndex -ge 0) {
            $propertyName = $e.Column.Header
            $propertyValue = $editedObject.GetType().GetProperty($propertyName).GetValue($editedObject)
            $categoryData[$categoryIndex].GetType().GetProperty($propertyName).SetValue($categoryData[$categoryIndex], $propertyValue)
        }
    }

    # Update the specific edited property in the All tab
    $propertyName = $e.Column.Header
    $propertyValue = $editedObject.GetType().GetProperty($propertyName).GetValue($editedObject)
    $allData[$allIndex].GetType().GetProperty($propertyName).SetValue($allData[$allIndex], $propertyValue)

    # Mark as having unsaved changes
    Set-UnsavedChanges $true

    # Update the category tab if the Category property changes
    if (($columnHeader -eq "Category") -and -not ([String]::IsNullOrWhiteSpace($e.EditingElement.Text))) {
        $newCategory = $e.EditingElement.Text
        $editedObject.Category = $newCategory

        if (-not $newObject) {
            $categoryData.RemoveAt($categoryIndex)
            if ($categoryData.Count -eq 0) {
                $tabControl.Items.Remove($tabs[$category])
                $tabs.Remove($category)
            }
        }

        # Add the object to the new category tab
        $newTab = $tabs[$newCategory]
        if (-not $newTab) {
            $itemsSource = New-Object System.Collections.ObjectModel.ObservableCollection[RowData]
            $newTab = New-DataTab -Name $newCategory -ItemsSource $itemsSource -TabControl $tabControl
            $tabs.Add($newCategory, $newTab)

            # Assign the CellEditEnding event to the new tab
            $newTab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
            $newTab.Content.Add_PreviewKeyDown({
                param($sender,$e)
                if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                    Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
                }
                elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
                    # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
                    if (-not $script:State.TabsReadOnly) {
                        $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                        $e.Handled = $true
                    }
                    # If Edit Mode is disabled (TabsReadOnly is true), run the command
                    else {
                        $e.Handled = $true
                        Invoke-MainRunClick -TabControl $script:UI.TabControl
                    }
                }
            })
        }
        $newTab.Content.ItemsSource.Add($editedObject)
        Sort-TabControl -TabControl $tabControl
    }
}
# Handle the Main Run Button click event to run the selected command/launch the CommandWindow
function Invoke-MainRunClick {
    param (
        [System.Windows.Controls.TabControl]$tabControl
    )

    $grid = $tabControl.SelectedItem.Content
    $selection = $grid.SelectedItems
    $command = New-Object Command
    $command.Full = ""
    $command.Root = $selection.Command
    $command.PreCommand = $selection.PreCommand
    $command.SkipParameterSelect = $selection.SkipParameterSelect
    $command.Log = $selection.Log

    Write-Log "Command created - Root: $($command.Root), Log: $($command.Log), SkipParameterSelect: $($command.SkipParameterSelect)"

    if ($command.Root) {
        if ($selection.SkipParameterSelect) {
            $command.Full = ""
            $command.CleanCommand = ""

            # Add log command if logging is enabled
            if ($command.Log) {
                $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
                $command.LogPath = "$($script:Settings.DefaultLogsPath)\$timestamp-$($command.Root).log"
                $command.Full = "Start-Transcript -Path `"$($command.LogPath)`""
                $command.Full += "; "
            }

            # Add PreCommand if it exists
            if ($command.PreCommand) {
                $command.Full += $command.PreCommand + "; "
                $command.CleanCommand += $command.PreCommand + "; "
            }

            $command.Full += $command.Root
            $command.CleanCommand += $command.Root

            # Add Stop-Transcript if logging is enabled
            if ($command.Log) {
                $command.Full += "; Stop-Transcript"
            }

            # Add to command history (no grid since parameters were skipped)
            $historyEntry = Add-CommandToHistory -Command $command

            Run-Command -Command $command -RunAttached $script:State.RunCommandAttached -HistoryEntry $historyEntry
        }
        else {
            Start-CommandWindow -Command $command
        }
    }
}

function Toggle-ShellGrid {
    if ($script:UI.Shell.Visibility -eq "Visible") {
        # Store current height before collapsing
        $script:State.SubGridExpandedHeight = $script:UI.Window.FindName("ShellRow").Height.Value

        # Collapse the Sub grid
        $script:UI.Window.FindName("ShellRow").Height = New-Object System.Windows.GridLength(0)
        $script:UI.Shell.Visibility = "Collapsed"

        # Update toggle button state
        if ($script:UI.BtnToggleShell) {
            $script:UI.BtnToggleShell.IsChecked = $false
        }
    }
    else {
        # Restore previous height and visibility
        $script:UI.Window.FindName("ShellRow").Height = New-Object System.Windows.GridLength($script:State.SubGridExpandedHeight)
        $script:UI.Shell.Visibility = "Visible"

        # Update toggle button state
        if ($script:UI.BtnToggleShell) {
            $script:UI.BtnToggleShell.IsChecked = $true
        }
    }
}

function Toggle-CommonParametersGrid {
    param(
        [System.Windows.Window]$CommandWindow
    )

    $commonGrid = $CommandWindow.FindName("CommonParametersGrid")
    $toggleButton = $CommandWindow.FindName("BtnToggleCommonParameters")
    $toggleIcon = $CommandWindow.FindName("IconToggleCommonParameters")

    if ($commonGrid.Visibility -eq "Collapsed") {
        # Build the common parameters grid if it hasn't been built yet
        if ($commonGrid.Children.Count -eq 0) {
            Build-CommonParametersGrid -CommandWindow $CommandWindow
        }

        # Show the grid
        $commonGrid.Visibility = "Visible"
        $toggleIcon.Kind = "ChevronUp"
        $toggleButton.ToolTip = "Hide Common Parameters"
    }
    else {
        # Hide the grid
        $commonGrid.Visibility = "Collapsed"
        $toggleIcon.Kind = "ChevronDown"
        $toggleButton.ToolTip = "Show Common Parameters"
    }
}

function Build-CommonParametersGrid {
    param(
        [System.Windows.Window]$CommandWindow
    )

    $grid = $CommandWindow.FindName("CommonParametersGrid")

    # Define common PowerShell parameters
    $commonParameters = @(
        @{ Name = "Verbose"; Type = "Switch"; Description = "Displays detailed information about the operation" }
        @{ Name = "Debug"; Type = "Switch"; Description = "Displays programmer-level detail about the operation" }
        @{ Name = "ErrorAction"; Type = "ValidateSet"; Values = @("", "Continue", "Ignore", "Inquire", "SilentlyContinue", "Stop", "Suspend"); Description = "Determines how the cmdlet responds to errors" }
        @{ Name = "WarningAction"; Type = "ValidateSet"; Values = @("", "Continue", "Inquire", "SilentlyContinue", "Stop"); Description = "Determines how the cmdlet responds to warnings" }
        @{ Name = "InformationAction"; Type = "ValidateSet"; Values = @("", "Continue", "Ignore", "Inquire", "SilentlyContinue", "Stop", "Suspend"); Description = "Determines how the cmdlet responds to information messages" }
        @{ Name = "ErrorVariable"; Type = "String"; Description = "Stores errors in the specified variable" }
        @{ Name = "WarningVariable"; Type = "String"; Description = "Stores warnings in the specified variable" }
        @{ Name = "InformationVariable"; Type = "String"; Description = "Stores information messages in the specified variable" }
        @{ Name = "OutVariable"; Type = "String"; Description = "Stores output objects in the specified variable" }
        @{ Name = "OutBuffer"; Type = "String"; Description = "Determines the number of objects to buffer before calling the next cmdlet" }
        @{ Name = "PipelineVariable"; Type = "String"; Description = "Stores the current pipeline object in the specified variable" }
    )

    for ($i = 0; $i -lt $commonParameters.Count; $i++) {
        $param = $commonParameters[$i]

        # Add row definition
        $rowDefinition = New-Object System.Windows.Controls.RowDefinition
        [void]$grid.RowDefinitions.Add($rowDefinition)

        # Create label
        $label = New-Label -Content $param.Name -HAlign "Left" -VAlign "Center"
        Add-ToGrid -Grid $grid -Element $label
        Set-GridPosition -Element $label -Row $i -Column 0

        # Add tooltip with description
        $label.ToolTip = New-ToolTip -Content $param.Description

        # Create appropriate control based on type
        if ($param.Type -eq "Switch") {
            $control = New-CheckBox -Name "Common_$($param.Name)" -IsChecked $false
        }
        elseif ($param.Type -eq "ValidateSet") {
            $control = New-ComboBox -Name "Common_$($param.Name)" -ItemsSource $param.Values -SelectedItem ""
        }
        else {
            $control = New-TextBox -Name "Common_$($param.Name)" -Text ""
        }

        Add-ToGrid -Grid $grid -Element $control
        Set-GridPosition -Element $control -Row $i -Column 2
    }
}



# Process the CommandWindow dialog grid to show command parameter list
function Start-CommandWindow([Command]$command) {

    # If we are rerunning the command then the parameters are already saved
    if (-not $command.Parameters) {
        # Show loading indicator immediately
        Show-LoadingIndicator -Message "Loading parameters for $($command.Root)..."

        # Force UI update to show the loading indicator
        $script:UI.Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

        # Create a runspace to process parameters asynchronously
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = "STA"
        $runspace.ThreadOptions = "ReuseThread"
        $runspace.Open()

        # Create PowerShell instance
        $powershell = [powershell]::Create()
        $powershell.Runspace = $runspace

        # Add script to extract parameters
        [void]$powershell.AddScript({
            param($commandName)

            try {
                # Get command type
                $type = (Get-Command $commandName -ErrorAction Stop).CommandType
                if (($type -ne "Function") -and ($type -ne "Script")) {
                    return @{
                        Success = $false
                        Error = "Command type '$type' not supported"
                        Type = $type
                    }
                }

                # Parse parameters
                $scriptBlock = (Get-Command $commandName -ErrorAction Stop).ScriptBlock
                if (-not $scriptBlock) {
                    return @{
                        Success = $false
                        Error = "Command does not have a script block"
                    }
                }

                $parsed = [System.Management.Automation.Language.Parser]::ParseInput($scriptBlock.ToString(), [ref]$null, [ref]$null)
                $parameters = $parsed.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)

                return @{
                    Success = $true
                    Parameters = $parameters
                    Type = $type
                }
            }
            catch {
                return @{
                    Success = $false
                    Error = $_.Exception.Message
                }
            }
        }).AddArgument($command.Root)

        # Begin async invocation
        $asyncResult = $powershell.BeginInvoke()

        # Capture variables needed in timer callback
        $commandRef = $command

        # Create timer to poll for completion
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(100)
        $timer.Tag = @{
            AsyncResult = $asyncResult
            PowerShell = $powershell
            Runspace = $runspace
            Command = $commandRef
        }

        $timer.Add_Tick({
            param($sender, $e)

            $data = $sender.Tag

            if ($data.AsyncResult.IsCompleted) {
                $sender.Stop()

                try {
                    # Get results
                    $result = $data.PowerShell.EndInvoke($data.AsyncResult)

                    if ($result.Success) {
                        # Update UI on dispatcher thread
                        $script:UI.Window.Dispatcher.Invoke([action]{
                            $data.Command.Parameters = $result.Parameters

                            # Hide loading and show dialog
                            Hide-LoadingIndicator
                            Show-CommandWindow -Command $data.Command
                        }, "Normal")
                    }
                    else {
                        # Handle error
                        $script:UI.Window.Dispatcher.Invoke([action]{
                            Hide-LoadingIndicator
                            Show-ParameterLoadError -CommandName $data.Command.Root -ErrorMessage $result.Error
                        }, "Normal")
                    }
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    $script:UI.Window.Dispatcher.Invoke([action]{
                        Hide-LoadingIndicator
                        Show-ParameterLoadError -CommandName $data.Command.Root -ErrorMessage $errorMsg
                    }, "Normal")
                }
                finally {
                    # Cleanup
                    $data.PowerShell.Dispose()
                    $data.Runspace.Close()
                    $data.Runspace.Dispose()
                }
            }
        })

        $timer.Start()
    }
    else {
        # Parameters already loaded, show dialog immediately
        Show-CommandWindow -Command $command
    }
}

# Display a new CommandWindow
function Show-CommandWindow {
    param(
        [Command]$Command
    )

    # Create a new CommandWindow instance
    $commandWindow = New-CommandWindow -Command $Command

    if ($commandWindow) {
        # Build the command grid with parameters
        if ($Command.Parameters) {
            Build-CommandGrid -CommandWindow $commandWindow -Parameters $Command.Parameters
        }

        # Show the window (non-modal so multiple can be open)
        $commandWindow.Window.Show()
    }
}

# Hide/Close a specific CommandWindow (kept for compatibility, but windows self-close)
function Hide-CommandWindow() {
    # This function is deprecated - windows now close themselves
    # Kept for backward compatibility
}

# Show loading indicator with optional custom message
function Show-LoadingIndicator {
    param (
        [string]$Message = "Loading command parameters..."
    )

    $script:UI.Window.Dispatcher.Invoke([action]{
        $script:UI.LoadingText.Text = $Message
        $script:UI.Overlay.Visibility = "Visible"
        $script:UI.LoadingIndicator.Visibility = "Visible"
        $script:UI.Window.Cursor = [System.Windows.Input.Cursors]::Wait
    }, "Send")
}

# Hide loading indicator
function Hide-LoadingIndicator {
    $script:UI.Window.Dispatcher.Invoke([action]{
        $script:UI.LoadingIndicator.Visibility = "Collapsed"
        $script:UI.Overlay.Visibility = "Collapsed"
        $script:UI.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    }, "Send")
}

# Construct the CommandWindow grid to show the correct content for each parameter
function Build-CommandGrid {
    param(
        $CommandWindow,
        [System.Object[]]$Parameters
    )

    $grid = $CommandWindow.CommandGrid
    for ($i = 0; $i -lt $parameters.Count; $i++) {
        $param = $parameters[$i]
        $paramName = $param.Name.VariablePath

        # Because there isn't a static number of rows and we need to iterate over the row index
        # we need to manually add a row for each parameter
        $rowDefinition = New-Object System.Windows.Controls.RowDefinition
        [void]$Grid.RowDefinitions.Add($rowDefinition)

        # In instances such as when the parameter is an array the value is stored
        # in DefaultValue rather than DefaultValue.Value
        if ((-not $param.DefaultValue.Value) -and ($param.DefaultValue)) {
            $paramDefault = $param.DefaultValue
        }
        else {
            $paramDefault = $param.DefaultValue.Value
        }
        $isMandatory = [System.Convert]::ToBoolean($param.Attributes.NamedArguments.Argument.VariablePath.Userpath)

        $label = New-Label -Content $paramName -HAlign "Left" -VAlign "Center"
        Add-ToGrid -Grid $Grid -Element $label
        Set-GridPosition -Element $label -Row $i -Column 0

        # Set asterisk next to values that are mandatory
        if ($isMandatory) {
            $asterisk = New-Label -Content "*" -HAlign "Right" -VAlign "Center"
            $asterisk.Foreground = "Red"
            Add-ToGrid -Grid $Grid -Element $asterisk
            Set-GridPosition -Element $asterisk -Row $i -Column 1
        }

        if (Test-AttributeType -Parameter $param -TypeName "ValidateSet") {
            # Get valid values from validate set and create dropdown box of them
            $validValues = Get-ValidateSetValues -Parameter $param
            $paramSource = $validValues -split "','"
            $box = New-ComboBox -Name $paramName -ItemsSource $paramSource -SelectedItem $paramDefault
        }
        elseif (Test-AttributeType -Parameter $param -TypeName "switch") {
            # If switch is true by default then check the box
            if ($param.DefaultValue) {
                $box = New-CheckBox -Name $paramName -IsChecked $true
            }
            else {
                $box = New-CheckBox -Name $paramName -IsChecked $false
            }
        }
        else {
            # Fill text box with any default values
            $box = New-TextBox -Name $paramName -Text $paramDefault
        }
        Add-ToGrid -Grid $Grid -Element $box
        Set-GridPosition -Element $box -Row $i -Column 2
    }
}

function Clear-Grid([System.Windows.Controls.Grid]$grid) {
    $grid.Children.Clear()
    $grid.RowDefinitions.Clear()
}

function Compile-Command {
    param(
        [Command]$Command,
        $CommandWindow
    )

    $grid = $CommandWindow.CommandGrid

    # Build the full command string
    $command.Full = ""
    $command.CleanCommand = ""

    # Add log command if logging is enabled
    if ($command.Log) {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $command.LogPath = "$($script:Settings.DefaultLogsPath)\$timestamp-$($command.Root).log"
        $command.Full = "Start-Transcript -Path `"$($command.LogPath)`""
        $command.Full += "; "
    }

    # Add PreCommand if it exists
    if ($command.PreCommand) {
        $command.Full += $command.PreCommand + "; "
        $command.CleanCommand += $command.PreCommand + "; "
    }

    $args = @{}
    $command.Full += "$($command.Root)"
    $command.CleanCommand += "$($command.Root)"

    foreach ($param in $command.Parameters) {
        $isSwitch = $false
        $paramName = $param.Name.VariablePath
        $selection = $grid.Children | Where-Object { $_.Name -eq $paramName }

        if (Test-AttributeType -Parameter $param -TypeName "ValidateSet") {
            if ($selection.SelectedItem) {
                $args[$paramName] = $selection.SelectedItem.ToString()
            }
        }
        elseif (Test-AttributeType -Parameter $param -TypeName "switch") {
            if ($selection.IsChecked) {
                $isSwitch = $true
            }
            elseif (-not ($selection.IsChecked) -and ($param.DefaultValue)) {
                # If switch isn't checked and by default it is, then explicitly set
                # the switch value to false
                $paramName = $paramName.ToString() + ':$false'
                $isSwitch = $true
            }
        }
        else {
            $args[$paramName] = $selection.Text
        }

        if ($isSwitch) {
            $command.Full += " -$paramName"
            $command.CleanCommand += " -$paramName"
        }
        elseif (-not [String]::IsNullOrWhiteSpace($args[$paramName])) {
            $command.Full += " -$paramName `"$($args[$paramName])`""
            $command.CleanCommand += " -$paramName `"$($args[$paramName])`""
        }
    }

    # Add common PowerShell parameters if the common parameters grid exists and is visible
    $commonGrid = $CommandWindow.CommandGrid.Parent.FindName("CommonParametersGrid")
    if ($commonGrid -and $commonGrid.Visibility -eq "Visible") {
        # Common parameter names to process
        $commonParamNames = @("Verbose", "Debug", "ErrorAction", "WarningAction", "InformationAction",
                              "ErrorVariable", "WarningVariable", "InformationVariable",
                              "OutVariable", "OutBuffer", "PipelineVariable")

        foreach ($paramName in $commonParamNames) {
            $controlName = "Common_$paramName"
            $control = $commonGrid.Children | Where-Object { $_.Name -eq $controlName }

            if ($control) {
                # Handle switch parameters (Verbose, Debug)
                if ($control -is [System.Windows.Controls.CheckBox]) {
                    if ($control.IsChecked) {
                        $command.Full += " -$paramName"
                        $command.CleanCommand += " -$paramName"
                    }
                }
                # Handle ValidateSet parameters (ErrorAction, WarningAction, InformationAction)
                elseif ($control -is [System.Windows.Controls.ComboBox]) {
                    if ($control.SelectedItem -and -not [String]::IsNullOrWhiteSpace($control.SelectedItem)) {
                        $command.Full += " -$paramName `"$($control.SelectedItem)`""
                        $command.CleanCommand += " -$paramName `"$($control.SelectedItem)`""
                    }
                }
                # Handle string parameters (variables and buffers)
                elseif ($control -is [System.Windows.Controls.TextBox]) {
                    if (-not [String]::IsNullOrWhiteSpace($control.Text)) {
                        $command.Full += " -$paramName `"$($control.Text)`""
                        $command.CleanCommand += " -$paramName `"$($control.Text)`""
                    }
                }
            }
        }
    }

    # Add Stop-Transcript if logging is enabled
    if ($command.Log) {
        $command.Full += "; Stop-Transcript"
    }
}

# Handle the Command Run Button click event to compile the inputted values for each parameter into a command string to be executed
function Invoke-CommandRunClick {
    param (
        [System.Windows.Window]$CommandWindow,
        [bool]$RunAttached
    )

    # Get command and grid from the window
    $command = $CommandWindow.Tag.Command
    $commandWindowHash = @{
        CommandGrid = $CommandWindow.FindName("CommandGrid")
    }

    Compile-Command -Command $command -CommandWindow $commandWindowHash

    # Add to command history
    $historyEntry = Add-CommandToHistory -Command $command -Grid $commandWindowHash.CommandGrid

    # Close the window
    $CommandWindow.Close()

    Run-Command -Command $command -RunAttached $RunAttached -HistoryEntry $historyEntry
}

function Invoke-CommandCopyToClipboard {
    param (
        [System.Windows.Window]$CommandWindow
    )

    # Get command and grid from the window
    $currentCommand = $CommandWindow.Tag.Command
    $commandWindowHash = @{
        CommandGrid = $CommandWindow.FindName("CommandGrid")
    }

    if ($currentCommand) {
        Compile-Command -Command $currentCommand -CommandWindow $commandWindowHash
        Copy-ToClipboard -String $currentCommand.CleanCommand
    }
}

# Execute a command string
function Run-Command {
    param (
        [Command]$command,
        [bool]$runAttached,
        [PSCustomObject]$historyEntry = $null
    )

    Write-Log "Running: $($command.Root)"
    Write-Log "Full Command: $($command.Full)"
    Write-Log "Log Enabled: $($command.Log)"

    # Ensure log directory exists if logging is enabled
    if ($command.Log) {
        try {
            if (-not (Test-Path $script:Settings.DefaultLogsPath)) {
                New-Item -ItemType Directory -Path $script:Settings.DefaultLogsPath -Force | Out-Null
                Write-Log "Created log directory: $($script:Settings.DefaultLogsPath)"
            }
            else {
                Write-Log "Log directory exists: $($script:Settings.DefaultLogsPath)"
            }
        }
        catch {
            Write-Log "Warning: Could not create log directory: $_"
        }
    }

    # We must escape any quotation marks passed or it will cause problems being passed through Start-Process
    $escapedCommand = $command.Full -replace '"', '\"'
    Write-Log "Escaped Command: $escapedCommand"

    if ($runAttached) {
        # Show the shell grid if it's not visible
        if ($script:UI.Shell.Visibility -ne "Visible") {
            Toggle-ShellGrid
        }

        # Show loading indicator while PowerShell window is being created
        Show-LoadingIndicator -Message "Starting PowerShell: $($command.Root)..."

        # Force UI update to show the loading indicator
        $script:UI.Window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

        # Switch to the Shell tab in TabControlShell if not already selected
        $shellTab = $script:UI.TabControlShell.Items | Where-Object { $_.Header -eq "Shell" }
        if ($shellTab -and $script:UI.TabControlShell.SelectedItem -ne $shellTab) {
            $script:UI.TabControlShell.SelectedItem = $shellTab
        }

        # Synchronously create the process tab (the 2-second wait is already in New-ProcessTab)
        # Note: New-ProcessTab automatically selects the newly created tab
        New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs "-ExecutionPolicy Bypass -NoExit `" & { $escapedCommand } `"" -TabName $command.Root -HistoryEntry $historyEntry

        # Hide loading indicator after tab is created
        Hide-LoadingIndicator
    }
    else {
        Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit `" & { $escapedCommand } `""
    }
}

# Determine the PowerShell command type (Function,Script,Cmdlet)
function Get-CommandType {
    param (
        [string]$command
    )

    try {
        return (Get-Command $command -ErrorAction Stop).CommandType
    }
    catch {
        throw "Command '$command' not found: $($_.Exception.Message)"
    }
}

# Parse the command's script block to extract parameter info
function Get-ScriptBlockParameters {
    param (
        [string]$command
    )

    try {
        $scriptBlock = (Get-Command $command -ErrorAction Stop).ScriptBlock
        if (-not $scriptBlock) {
            throw "Command '$command' does not have a script block (may be a compiled cmdlet)"
        }
        $parsed = [System.Management.Automation.Language.Parser]::ParseInput($scriptBlock.ToString(), [ref]$null, [ref]$null)
        return $parsed.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
    }
    catch {
        throw "Failed to parse parameters for '$command': $($_.Exception.Message)"
    }
}

# Determine if a parameter contains a certain attribute type (Switch,ValidateSet)
function Test-AttributeType {
    param (
        [System.Management.Automation.Language.ParameterAst]$parameter,
        [string]$typeName
    )

    return $parameter.Attributes | Where-Object { $_.TypeName.FullName -eq $typeName }
}

# Retrieve the list of values from a parameter's ValidateSet
function Get-ValidateSetValues {
    param (
        [System.Management.Automation.Language.ParameterAst]$parameter
    )

    $validValues = [System.Collections.ArrayList]@("")
    $values = ($parameter.Attributes | Where-Object { $_.TypeName.FullName -eq 'ValidateSet' }).PositionalArguments
    foreach ($value in $values) {
        $valueStr = $($value.ToString()).Replace("'","").Replace("`"","")
        [void]$validValues.Add($valueStr)
    }
    return $validValues
}
# Load and process main application window
function Start-MainWindow {
    # We create a new window and load all the window elements to variables of 
    # the same name and assign the window and all its elements under $script:UI 
    # e.g. $script:UI.Window, $script:UI.TabControl
    try {
        $script:UI = New-Window -File $script:ApplicationPaths.MainWindowXamlFile -ErrorAction Stop
    }
    catch {
        Show-ErrorMessageBox "Failed to create window from $($script:ApplicationPaths.MainWindowXamlFile): $_"
        exit(1)
    }

    Initialize-Settings

    # Use the configured DefaultDataFile from settings
    $script:State.CurrentDataFile = $script:Settings.DefaultDataFile
    Initialize-DataFile $script:State.CurrentDataFile
    $json = Load-DataFile $script:State.CurrentDataFile

    # Ensure we always have a valid collection, even if file is empty
    if (-not $json) {
        $json = [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
    }
    $script:State.HighestId = Get-HighestId -Json $json
    $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json)

    # Create tabs and grids
    $script:UI.Tabs = @{}
    $allTab = New-DataTab -Name "All" -ItemsSource $itemsSource -TabControl $script:UI.TabControl
    $allTab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $allTab.Content.Add_PreviewKeyDown({
        param($sender,$e)
        if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
            Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
        }
        elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
            # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
            if (-not $script:State.TabsReadOnly) {
                $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                $e.Handled = $true
            }
            # If Edit Mode is disabled (TabsReadOnly is true), run the command
            else {
                $e.Handled = $true
                Invoke-MainRunClick -TabControl $script:UI.TabControl
            }
        }
    })
    [void]$script:UI.Tabs.Add("All", $allTab)

    $favItemsSource = [System.Collections.ObjectModel.ObservableCollection[FavoriteRowData]]::new()
    $loadedFavorites = Load-Favorites -AllData $json
    foreach ($fav in $loadedFavorites) {
        $favItemsSource.Add($fav)
    }
    $favTab = New-DataTab -Name "*" -ItemsSource $favItemsSource -TabControl $script:UI.TabControl
    $favTab.Content.Add_CellEditEnding({
        param($sender,$e)
        if ($e.Column.Header -eq "Order") {
            # Special handling for Order changes
            $favorites = $sender.ItemsSource
            Save-Favorites -Favorites $favorites
        } else {
            Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
        }
    })
    $favTab.Content.Add_PreviewKeyDown({
        param($sender,$e)
        if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
            Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
        }
        elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
            # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
            if (-not $script:State.TabsReadOnly) {
                $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                $e.Handled = $true
            }
            # If Edit Mode is disabled (TabsReadOnly is true), run the command
            else {
                $e.Handled = $true
                Invoke-MainRunClick -TabControl $script:UI.TabControl
            }
        }
    })

    # Add drag/drop event handlers for reordering favorites
    Initialize-FavoritesDragDrop -Grid $favTab.Content

    [void]$script:UI.Tabs.Add("Favorites", $favTab)
    if ($favItemsSource.Count -eq 0) {
        $script:UI.TabControl.SelectedItem = $allTab
    }

    foreach ($category in ($json | Select-Object -ExpandProperty Category -Unique)) {
        $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json | Where-Object { $_.Category -eq $category })
        $tab = New-DataTab -Name $category -ItemsSource $itemsSource -TabControl $script:UI.TabControl
        $tab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs }) # We need to assign the cell edit handler to each tab's grid so that it works for all tabs
        $tab.Content.Add_PreviewKeyDown({
            param($sender,$e)
            if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
                # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
                if (-not $script:State.TabsReadOnly) {
                    $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                    $e.Handled = $true
                }
                # If Edit Mode is disabled (TabsReadOnly is true), run the command
                else {
                    $e.Handled = $true
                    Invoke-MainRunClick -TabControl $script:UI.TabControl
                }
            }
        })
        $script:UI.Tabs.Add($category, $tab)
    }
    Sort-TabControl -TabControl $script:UI.TabControl

    # Initialize favorite highlighting after all tabs are created
    Update-FavoriteHighlighting

    Register-EventHandlers

    # Set content and display the window
    $script:UI.Window.DataContext = $script:UI.Tabs
    $script:UI.Window.Dispatcher.InvokeAsync{ $script:UI.Window.ShowDialog() }.Wait() | Out-Null
}

# Register all GUI events
function Register-EventHandlers {
    # Dock Panel Menu Buttons
    $script:UI.BtnMenuAdd.Add_Click({ Add-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuDuplicate.Add_Click({ Duplicate-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuRemove.Add_Click({ Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuUndoDelete.Add_Click({ Restore-DeletedCommand -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuSave.Add_Click({ Save-DataFile -FilePath $script:State.CurrentDataFile -Data ($script:UI.Tabs["All"].Content.ItemsSource) })
    $script:UI.BtnMenuSaveAs.Add_Click({ Save-DataFileAs })
    $script:UI.BtnMenuOpen.Add_Click({ Open-DataFile })
    $script:UI.BtnMenuImport.Add_Click({ Invoke-ImportDataFileDialog })
    $script:UI.BtnMenuEdit.Add_Click({ Toggle-EditMode -Tabs $script:UI.Tabs })
    $script:UI.BtnToggleEditMode.Add_Click({ Toggle-EditMode -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuFavorite.Add_Click({ Toggle-CommandFavorite })
    $script:UI.BtnMenuSettings.Add_Click({ Show-SettingsDialog })
    $script:UI.BtnMenuRunOpen.Add_Click({
        Invoke-MainRunClick -TabControl $script:UI.TabControl
    })
    $script:UI.BtnMenuRunDetached.Add_Click({
        $script:State.RunCommandAttached = $false
        Invoke-MainRunClick -TabControl $script:UI.TabControl
    })
    $script:UI.BtnMenuRunAttached.Add_Click({
        $script:State.RunCommandAttached = $true
        Invoke-MainRunClick -TabControl $script:UI.TabControl -Attached $true
    })
    $script:UI.BtnMenuRunRerunLast.Add_Click({
        if ($script:State.CommandHistory -and $script:State.CommandHistory.Count -gt 0) {
            $lastHistoryEntry = $script:State.CommandHistory[0]
            Reopen-CommandFromHistory -HistoryEntry $lastHistoryEntry
        }
        else {
            Write-Status "No command history available"
        }
    })

    # Main Buttons
    $script:UI.BtnMainRun.Add_Click({
        $script:State.RunCommandAttached = $script:Settings.DefaultRunCommandAttached
        Invoke-MainRunClick -TabControl $script:UI.TabControl
    })
    $script:UI.BtnToggleShell.Add_Click({ Toggle-ShellGrid })

    # Command dialog button events - Now handled per-window in New-CommandWindow

    # Settings dialog button events
    $script:UI.BtnBrowseLogs.Add_Click({ Invoke-BrowseLogs })
    $script:UI.BtnBrowseDataFile.Add_Click({ Invoke-BrowseDataFile })
    $script:UI.BtnBrowseSettings.Add_Click({ Invoke-BrowseSettings })
    $script:UI.BtnBrowseFavorites.Add_Click({ Invoke-BrowseFavorites })
    $script:UI.BtnBrowseHistory.Add_Click({ Invoke-BrowseHistory })
    $script:UI.BtnApplySettings.Add_Click({ Apply-Settings })
    $script:UI.BtnCloseSettings.Add_Click({ Hide-SettingsDialog })

    # Main Tab Control events
    $script:UI.TabControl.Add_SelectionChanged({
        param($sender, $e)
        Handle-TabSelection -SelectedTab $sender.SelectedItem
        Update-MainRunButtonText
        # Re-apply the filter when tab changes
        if ($script:UI.TxtSearchFilter) {
            Invoke-GridFilter -SearchText $script:UI.TxtSearchFilter.Text
        }
    })

    # Search filter text box event
    $script:UI.TxtSearchFilter.Add_TextChanged({
        param($sender, $e)
        Invoke-GridFilter -SearchText $sender.Text
    })

    # Process Tab events
    # Add context menu to the "+" tab
    $addTabContextMenu = New-Object System.Windows.Controls.ContextMenu
    $addTabContextMenu.FontSize = 12

    # New PS Session menu item
    $menuNewSession = New-Object System.Windows.Controls.MenuItem
    $menuNewSession.Header = "New PS Session"
    $menuNewSession.FontSize = 12

    # Create icon for New PS Session
    $iconNew = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconNew.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Plus
    $iconNew.Width = 16
    $iconNew.Height = 16
    $iconNew.Margin = New-Object System.Windows.Thickness(0)
    $menuNewSession.Icon = $iconNew

    $menuNewSession.Add_Click({
        New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs $script:Settings.DefaultShellArgs
    })
    $addTabContextMenu.Items.Add($menuNewSession)

    # Attach PS Session menu item
    $menuAttachSession = New-Object System.Windows.Controls.MenuItem
    $menuAttachSession.Header = "Attach PS Session"
    $menuAttachSession.FontSize = 12

    # Create icon for Attach PS Session
    $iconAttach = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconAttach.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Import
    $iconAttach.Width = 16
    $iconAttach.Height = 16
    $iconAttach.Margin = New-Object System.Windows.Thickness(0)
    $menuAttachSession.Icon = $iconAttach

    $menuAttachSession.Add_Click({
        Show-AttachWindow
    })
    $addTabContextMenu.Items.Add($menuAttachSession)

    $script:UI.PSAddTab.ContextMenu = $addTabContextMenu

    # Right-click on PSAddTab shows context menu
    $script:UI.PSAddTab.Add_PreviewMouseRightButtonDown({
        param($sender, $e)
        $sender.ContextMenu.IsOpen = $true
        $e.Handled = $true
    })

    # Left-click on PSAddTab also shows context menu (instead of immediately creating tab)
    $script:UI.PSAddTab.Add_PreviewMouseLeftButtonDown({
        param($sender, $e)
        $sender.ContextMenu.IsOpen = $true
        $e.Handled = $true
    })
    $script:UI.PSTabControl.Add_SelectionChanged({
        param($sender, $eventArgs)
        $selectedTab = $script:UI.PSTabControl.SelectedItem
        if (($selectedTab) -and ($selectedTab -ne $script:UI.PSAddTab)) {
            $psHandle = $selectedTab.Tag["Handle"]
            [Win32]::SetFocus($psHandle)
        }
    })
    $script:UI.Window.Add_GotFocus({
        if ($script:UI.PSTabControl.SelectedItem -ne $script:UI.PSAddTab) {
            $psHandle = $script:UI.PSTabControl.SelectedItem.Tag["Handle"]
            [Win32]::SetFocus($psHandle)
        }
    })
    
    # Log Tab events
    $script:UI.LogAddTab.Add_PreviewMouseLeftButtonDown({ Open-LogFile })

    # Command History events
    Initialize-CommandHistoryUI
    Load-CommandHistory

    $script:UI.Window.Add_Loaded({
        $script:UI.Window.Icon = $script:ApplicationPaths.IconFile
        Update-WindowTitle

        if (-not $script:Settings.OpenShellAtStart) {
            Toggle-ShellGrid
        }
    })

    $script:UI.Window.Add_Closing({ param($sender, $e) Invoke-WindowClosing -Sender $sender -E $e })
}

# Update the main Run button text and menu items visibility based on the selected command
function Update-MainRunButtonText {
    if (-not $script:UI.TabControl.SelectedItem) {
        $script:UI.BtnMainRun.Content = "Run"
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Visible
        return
    }

    $grid = $script:UI.TabControl.SelectedItem.Content
    if (-not $grid -or -not $grid.SelectedItem) {
        $script:UI.BtnMainRun.Content = "Run"
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Visible
        return
    }

    $selectedItem = $grid.SelectedItem

    if ($selectedItem.SkipParameterSelect) {
        # Show Run (Attached) or Run (Detached) based on DefaultRunCommandAttached setting
        if ($script:Settings.DefaultRunCommandAttached) {
            $script:UI.BtnMainRun.Content = "Run (Attached)"
        }
        else {
            $script:UI.BtnMainRun.Content = "Run (Detached)"
        }
        # Hide Open menu item, show Run items
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Collapsed
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Visible
    }
    else {
        # Show "Open" when SkipParameterSelect is false
        $script:UI.BtnMainRun.Content = "Open"
        # Show Open menu item, hide Run items
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Collapsed
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

function Invoke-WindowClosing {
    param($sender, $e)

    # Check for unsaved changes before closing
    if ($script:State.HasUnsavedChanges) {
        $result = [System.Windows.MessageBox]::Show(
            "You have unsaved changes. Do you want to save them before closing?",
            "Unsaved Changes",
            [System.Windows.MessageBoxButton]::YesNoCancel,
            [System.Windows.MessageBoxImage]::Question
        )

        if ($result -eq [System.Windows.MessageBoxResult]::Cancel) {
            $e.Cancel = $true
            return
        }
        elseif ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Save-DataFile -FilePath $script:State.CurrentDataFile -Data ($script:UI.Tabs["All"].Content.ItemsSource)
        }
    }

    $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
    Save-Favorites -Favorites $favorites

    # Close all open CommandWindows
    foreach ($window in $script:State.OpenCommandWindows) {
        try {
            if ($window -and -not $window.IsClosed) {
                $window.Close()
            }
        }
        catch {
            Write-Log "Error closing CommandWindow: $_"
        }
    }

    foreach ($tab in $script:UI.PSTabControl.Items) {
        if ($tab -ne $script:UI.PSAddTab) {
            $process = $tab.Tag["Process"]
            if ($process -ne $null) {
                try {
                    $process.Kill()
                }
                catch {
                    Write-Log "Error closing process: $_"
                }
            }
        }
    }
    Write-Log "PSGUI is shutting down..."
}
# Create a blank data file if it doesn't already exist
function Initialize-DataFile {
    param (
        [string]$filePath
    )

    if (-not (Test-Path $filePath)) {
        try {
            # Ensure the directory exists
            $directory = Split-Path -Path $filePath -Parent
            if (-not (Test-Path $directory)) {
                New-Item -ItemType Directory -Path $directory -Force | Out-Null
            }

            # Create file with empty structure including CommandListId
            $newFileStructure = @{
                CommandListId = Get-CommandListId
                Commands = @()
            }
            $newFileStructure | ConvertTo-Json -Depth 3 | Set-Content -Path $filePath -Encoding UTF8
            Write-Log "Created new data file: $filePath"
        }
        catch {
            Show-ErrorMessageBox("Failed to create configuration file at path: $filePath - $_")
            exit(1)
        }
    }
}

# Load an existing data file
function Load-DataFile {
    param (
        [string]$filePath
    )

    try {
        [string]$contentRaw = (Get-Content $filePath -Raw -ErrorAction Stop)
        if ($contentRaw) {
            $contentJson = $contentRaw | ConvertFrom-Json

            # Handle both old format (array) and new format (object with CommandListId)
            $commandsArray = $null
            if ($contentJson -is [Array]) {
                # Old format - array of commands
                $commandsArray = $contentJson
                # Generate and store a new CommandListId for old files
                $script:State.CurrentCommandListId = Get-CommandListId
                Write-Log "Loaded legacy data file format, generated new CommandListId"
            } else {
                # New format - object with CommandListId and Commands
                $commandsArray = $contentJson.Commands
                $script:State.CurrentCommandListId = $contentJson.CommandListId
                if (-not $script:State.CurrentCommandListId) {
                    # Generate ID if missing
                    $script:State.CurrentCommandListId = Get-CommandListId
                    Write-Log "CommandListId missing from file, generated new one"
                }
            }

            # Convert JSON objects to RowData objects
            $rowDataCollection = [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
            if ($commandsArray) {
                foreach ($item in $commandsArray) {
                    $rowData = [RowData]::new()
                    $rowData.Id = $item.Id
                    $rowData.Name = $item.Name
                    $rowData.Description = $item.Description
                    $rowData.Category = $item.Category
                    $rowData.Command = $item.Command
                    $rowData.SkipParameterSelect = $item.SkipParameterSelect
                    $rowData.PreCommand = $item.PreCommand
                    $rowData.Log = $item.Log
                    $rowDataCollection.Add($rowData)
                }
            }
            return $rowDataCollection
        }
        else {
            Write-Verbose "Data file $filePath is empty."
            # Generate CommandListId for empty files
            $script:State.CurrentCommandListId = Get-CommandListId
            return [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
        }
    }
    catch {
        Write-Error "Failed to load data from: $filePath"
        Write-Log "Failed to load data: $_"
        # Generate CommandListId even for failed loads
        $script:State.CurrentCommandListId = Get-CommandListId
        return [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
    }
}

# Save the data collection to the data file
function Save-DataFile {
    param (
        [string]$filePath,
        [System.Collections.ObjectModel.ObservableCollection[Object]]$data
    )

    try {
        # Filter out unpopulated rows and convert to plain objects for JSON serialization
        $populatedRows = $data | Where-Object { $_.Name -ne $null } | ForEach-Object {
            @{
                Id = $_.Id
                Name = $_.Name
                Description = $_.Description
                Category = $_.Category
                Command = $_.Command
                SkipParameterSelect = $_.SkipParameterSelect
                PreCommand = $_.PreCommand
                Log = $_.Log
            }
        }

        # Save in new format with CommandListId
        $fileStructure = @{
            CommandListId = $script:State.CurrentCommandListId
            Commands = $populatedRows
        }

        $json = ConvertTo-Json $fileStructure -Depth 3
        Set-Content -Path $filePath -Value $json
        Set-UnsavedChanges $false
        Write-Status "Data saved"
    }
    catch {
        Write-Error "Failed to save data to: $filePath"
        Write-Log "Failed to save data: $_"
        throw
    }
}

function Open-DataFile {
    # Check for unsaved changes first
    if (-not (Confirm-SaveBeforeAction "opening a new data file")) {
        return
    }

    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.InitialDirectory = Split-Path $script:State.CurrentDataFile -Parent
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.FilterIndex = 1

    if ($dialog.ShowDialog()) {
        # Load new data file
        $script:State.CurrentDataFile = $dialog.FileName
        Load-NewDataFile -FilePath $script:State.CurrentDataFile
        Set-UnsavedChanges $false
        Update-WindowTitle
        Write-Status "Opened data file: $($dialog.FileName)"
    }
}

function Save-DataFileAs {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.InitialDirectory = Split-Path $script:State.CurrentDataFile -Parent
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    $dialog.FileName = [System.IO.Path]::GetFileNameWithoutExtension($script:State.CurrentDataFile) + "_copy.json"

    if ($dialog.ShowDialog()) {
        # Generate a new CommandListId for the saved file
        $originalCommandListId = $script:State.CurrentCommandListId
        $script:State.CurrentCommandListId = Get-CommandListId

        try {
            # Save data to the new file with new CommandListId
            Save-DataFile -FilePath $dialog.FileName -Data ($script:UI.Tabs["All"].Content.ItemsSource)

            # Update current data file path to the new file
            $script:State.CurrentDataFile = $dialog.FileName
            Set-UnsavedChanges $false
            Update-WindowTitle
            Write-Status "Data saved as: $($dialog.FileName)"

            # Clear and reload favorites for the new CommandListId (will be empty initially)
            $favItemsSource = [System.Collections.ObjectModel.ObservableCollection[FavoriteRowData]]::new()
            $script:UI.Tabs["Favorites"].Content.ItemsSource = $favItemsSource
            Update-FavoriteHighlighting
        }
        catch {
            # Restore original CommandListId if save failed
            $script:State.CurrentCommandListId = $originalCommandListId
            Show-ErrorMessageBox "Failed to save file as: $_"
        }
    }
}

function Invoke-ImportDataFileDialog {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.InitialDirectory = Split-Path $script:State.CurrentDataFile -Parent
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.FilterIndex = 1

    if ($dialog.ShowDialog()) {
        Import-DataFile -FilePath $dialog.FileName
        Write-Status "Imported data from: $($dialog.FileName)"
    }
}

function Load-NewDataFile {
    param (
        [string]$filePath
    )

    try {
        $json = Load-DataFile $filePath
        $script:State.HighestId = Get-HighestId -Json $json

        # Clear existing tabs except Favorites
        $tabsToRemove = @()
        foreach ($tab in $script:UI.TabControl.Items) {
            if ($tab.Header -ne "*" -and $tab.Header -ne "All") {
                $tabsToRemove += $tab
            }
        }
        foreach ($tab in $tabsToRemove) {
            $script:UI.TabControl.Items.Remove($tab)
            $script:UI.Tabs.Remove($tab.Header)
        }

        # Update All tab with new data
        $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json)
        $script:UI.Tabs["All"].Content.ItemsSource = $itemsSource

        # Recreate category tabs
        foreach ($category in ($json | Select-Object -ExpandProperty Category -Unique | Where-Object { $_ -ne $null -and $_ -ne "" })) {
            $categoryItemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json | Where-Object { $_.Category -eq $category })
            $tab = New-DataTab -Name $category -ItemsSource $categoryItemsSource -TabControl $script:UI.TabControl
            $tab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
            $tab.Content.Add_PreviewKeyDown({ param($sender,$e) if ($e.Key -eq [System.Windows.Input.Key]::Delete) { Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs } })
            $script:UI.Tabs.Add($category, $tab)
        }

        # Reload favorites based on new data
        $favItemsSource = [System.Collections.ObjectModel.ObservableCollection[FavoriteRowData]]::new()
        $loadedFavorites = Load-Favorites -AllData $json
        foreach ($fav in $loadedFavorites) {
            $favItemsSource.Add($fav)
        }
        $script:UI.Tabs["Favorites"].Content.ItemsSource = $favItemsSource

        # Reinitialize drag/drop for favorites grid
        Initialize-FavoritesDragDrop -Grid $script:UI.Tabs["Favorites"].Content

        Sort-TabControl -TabControl $script:UI.TabControl
    }
    catch {
        Show-ErrorMessageBox "Failed to load data file: $_"
    }
}

function Import-DataFile {
    param (
        [string]$filePath
    )

    try {
        $importedJson = Load-DataFile $filePath
        if (-not $importedJson -or $importedJson.Count -eq 0) {
            Write-Status "No data found in file to import"
            return
        }

        $allData = $script:UI.Tabs["All"].Content.ItemsSource
        if ($null -eq $allData) {
            Write-Status "Error: All tab data source not found"
            return
        }

        $importCount = 0
        foreach ($item in $importedJson) {
            if (-not $item) { continue }

            # Check if item with same ID already exists
            $existingItem = $allData | Where-Object { $_.Id -eq $item.Id }
            if ($existingItem) {
                # Update the highest ID to avoid conflicts
                $item.Id = ++$script:State.HighestId
            } else {
                # Update highest ID if this ID is higher
                if ($item.Id -gt $script:State.HighestId) {
                    $script:State.HighestId = $item.Id
                }
            }

            # Add to All tab
            $allData.Add($item)
            $importCount++

            # Add to category tab if category exists
            if ($item.Category -and $item.Category -ne "") {
                $categoryTab = $script:UI.Tabs[$item.Category]
                if (-not $categoryTab) {
                    # Create new category tab
                    $categoryItemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
                    $categoryTab = New-DataTab -Name $item.Category -ItemsSource $categoryItemsSource -TabControl $script:UI.TabControl
                    $categoryTab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
                    $categoryTab.Content.Add_PreviewKeyDown({ param($sender,$e) if ($e.Key -eq [System.Windows.Input.Key]::Delete) { Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs } })
                    $script:UI.Tabs.Add($item.Category, $categoryTab)
                }
                if ($categoryTab -and $categoryTab.Content -and $categoryTab.Content.ItemsSource) {
                    $categoryTab.Content.ItemsSource.Add($item)
                }
            }
        }

        # Refresh the grid items to ensure UI updates
        $script:UI.Tabs["All"].Content.Items.Refresh()

        # Update favorite highlighting in case any imported items are favorites
        Update-FavoriteHighlighting

        Set-UnsavedChanges $true
        Sort-TabControl -TabControl $script:UI.TabControl
        Write-Log "Imported $importCount command(s) from: $filePath"
    }
    catch {
        Show-ErrorMessageBox "Failed to import data file: $_"
    }
}
function Save-Favorites {
    param (
        [System.Collections.ObjectModel.ObservableCollection[Object]]$favorites
    )

    try {
        $favoritesDir = Split-Path $script:Settings.FavoritesPath -Parent
        if (-not (Test-Path $favoritesDir)) {
            New-Item -ItemType Directory -Path $favoritesDir -Force | Out-Null
        }

        # Load existing favorites file or create new structure
        $allFavorites = @{}
        if (Test-Path $script:Settings.FavoritesPath) {
            try {
                $existingContent = Get-Content $script:Settings.FavoritesPath | ConvertFrom-Json
                # Handle both old format (array) and new format (object with command list IDs)
                if ($existingContent -is [Array]) {
                    # Convert old format - all favorites go under a default ID
                    $allFavorites["default"] = $existingContent
                } else {
                    # Convert PSCustomObject to hashtable for proper manipulation
                    $existingContent.PSObject.Properties | ForEach-Object {
                        $allFavorites[$_.Name] = $_.Value
                    }
                }
            }
            catch {
                Write-Log "Failed to parse existing favorites file, creating new one"
            }
        }

        # Save favorites for current command list
        $currentListId = $script:State.CurrentCommandListId
        if ($currentListId) {
            $favoriteData = $favorites | Select-Object Id, Order
            $allFavorites[$currentListId] = $favoriteData
        }

        # Convert hashtable to PSCustomObject for proper JSON serialization
        $outputObject = New-Object PSObject
        $allFavorites.GetEnumerator() | ForEach-Object {
            $outputObject | Add-Member -MemberType NoteProperty -Name $_.Key -Value $_.Value
        }

        $outputObject | ConvertTo-Json -Depth 3 | Set-Content $script:Settings.FavoritesPath
        Write-Status "Favorites saved"
    }
    catch {
        Write-Status "Failed to save favorites"
        Write-Log "Failed to save favorites: $_"
    }
}

function Load-Favorites {
    param (
        [System.Collections.ObjectModel.ObservableCollection[RowData]]$allData
    )

    try {
        if (Test-Path $script:Settings.FavoritesPath) {
            $allFavorites = Get-Content $script:Settings.FavoritesPath | ConvertFrom-Json
            $favorites = @()

            # Get favorites for current command list
            $currentListId = $script:State.CurrentCommandListId
            $favoriteData = $null

            # Handle both old format (array) and new format (object with command list IDs)
            if ($allFavorites -is [Array]) {
                # Old format - use as default
                $favoriteData = $allFavorites
            } else {
                # New format - get favorites for current command list
                if ($currentListId -and $allFavorites.PSObject.Properties[$currentListId]) {
                    $favoriteData = $allFavorites.$currentListId
                }
            }

            if ($favoriteData) {
                foreach ($fav in $favoriteData | Sort-Object Order) {
                    $rowData = $allData | Where-Object { $_.Id -eq $fav.Id }
                    if ($rowData) {
                        $favoriteRow = [FavoriteRowData]::new($rowData, $fav.Order)
                        $favorites += $favoriteRow
                        if ($fav.Order -gt $script:State.FavoritesHighestOrder) {
                            $script:State.FavoritesHighestOrder = $fav.Order
                        }
                    }
                }
            }
            return $favorites
        }
    }
    catch {
        Write-Log "Failed to load favorites: $_"
        return @()
    }
    return @()
}


# Update favorite highlighting across all tabs except the Favorites tab
function Update-FavoriteHighlighting {
    $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
    $favoriteIds = @($favorites | ForEach-Object { $_.Id })

    foreach ($tabEntry in $script:UI.Tabs.GetEnumerator()) {
        $tabName = $tabEntry.Key
        $tab = $tabEntry.Value

        # Skip the Favorites tab since it only contains favorites
        if ($tabName -eq "Favorites") { continue }

        $grid = $tab.Content
        if ($grid -and $grid.Items) {
            # Use Dispatcher to ensure UI updates happen on the UI thread
            $script:UI.Window.Dispatcher.Invoke([action]{
                foreach ($item in $grid.Items) {
                    $container = $grid.ItemContainerGenerator.ContainerFromItem($item)
                    if ($container -is [System.Windows.Controls.DataGridRow]) {
                        if ($favoriteIds -contains $item.Id) {
                            $container.Tag = "IsFavorite"
                        }
                        else {
                            $container.Tag = $null
                        }
                    }
                }
            }, "Normal")
        }
    }
}

# Initialize drag and drop functionality for the Favorites grid
function Initialize-FavoritesDragDrop {
    param (
        [System.Windows.Controls.DataGrid]$grid
    )

    # Enable drag/drop on the grid
    $grid.AllowDrop = $true

    # Handle mouse down to capture the item being dragged
    $grid.Add_PreviewMouseLeftButtonDown({
        param($sender, $e)

        # Disable drag-and-drop when in edit mode
        if (-not $script:State.TabsReadOnly) {
            return
        }

        $row = Get-DataGridRowFromPoint -Grid $sender -Point ($e.GetPosition($sender))
        if ($row -and $row.Item) {
            $script:State.DragDrop.DraggedItem = $row.Item
        }
    })

    # Handle mouse move to initiate drag operation
    $grid.Add_MouseMove({
        param($sender, $e)

        # Disable drag-and-drop when in edit mode
        if (-not $script:State.TabsReadOnly) {
            return
        }

        if ($e.LeftButton -eq [System.Windows.Input.MouseButtonState]::Pressed -and
            $script:State.DragDrop.DraggedItem -ne $null) {

            $dragData = New-Object System.Windows.DataObject([System.Windows.DataFormats]::Serializable, $script:State.DragDrop.DraggedItem)
            [System.Windows.DragDrop]::DoDragDrop($sender, $dragData, [System.Windows.DragDropEffects]::Move)
        }
    })

    # Handle drag over to show drop feedback
    $grid.Add_DragOver({
        param($sender, $e)

        # Disable drag-and-drop when in edit mode
        if (-not $script:State.TabsReadOnly) {
            $e.Effects = [System.Windows.DragDropEffects]::None
            $e.Handled = $true
            return
        }

        $position = $e.GetPosition($sender)
        $row = Get-DataGridRowFromPoint -Grid $sender -Point $position

        if ($row) {
            # Normal drop - highlight top border of target row
            # Only update highlighting if the row object or border type changed
            if ($row -ne $script:State.DragDrop.LastHighlightedRow -or $script:State.DragDrop.IsBottomBorder -eq $true) {
                Clear-DropHighlight
                $script:State.DragDrop.IsBottomBorder = $false
                Set-DropHighlight -Row $row -IsBottomBorder $false
            }
            $e.Effects = [System.Windows.DragDropEffects]::Move
        }
        elseif (Test-IsPositionBelowLastRow -Grid $sender -Position $position) {
            # Drop after last item - highlight bottom border of last row
            $lastRow = $sender.ItemContainerGenerator.ContainerFromIndex($sender.Items.Count - 1)

            # Only update if we're not already highlighting this row's bottom border
            if ($lastRow -ne $script:State.DragDrop.LastHighlightedRow -or $script:State.DragDrop.IsBottomBorder -eq $false) {
                Clear-DropHighlight
                $script:State.DragDrop.IsBottomBorder = $true
                Set-DropHighlight -Row $lastRow -IsBottomBorder $true
            }
            $e.Effects = [System.Windows.DragDropEffects]::Move
        }
        else {
            $e.Effects = [System.Windows.DragDropEffects]::None
        }
        $e.Handled = $true
    })

    # Handle drag leave to clear feedback
    $grid.Add_DragLeave({
        param($sender, $e)
        Clear-DropHighlight
        $script:State.DragDrop.IsBottomBorder = $false
    })

    # Handle drop to reorder items
    $grid.Add_Drop({
        param($sender, $e)

        # Disable drag-and-drop when in edit mode
        if (-not $script:State.TabsReadOnly) {
            $e.Handled = $true
            return
        }

        Clear-DropHighlight

        if ($script:State.DragDrop.DraggedItem -ne $null) {
            $itemsSource = $sender.ItemsSource
            $draggedItem = $script:State.DragDrop.DraggedItem
            $position = $e.GetPosition($sender)

            $targetRow = Get-DataGridRowFromPoint -Grid $sender -Point $position
            $targetItem = $null
            $isDropAfterLast = $false

            # Determine drop target and whether dropping after last item
            if (Test-IsPositionBelowLastRow -Grid $sender -Position $position) {
                # Dropping after last item
                $targetItem = $sender.Items[$sender.Items.Count - 1]
                $isDropAfterLast = $true
            }
            elseif ($targetRow) {
                # Normal drop on a row
                $targetItem = $targetRow.Item
            }

            if ($targetItem -and ($draggedItem -ne $targetItem -or $isDropAfterLast)) {
                $draggedOrder = $draggedItem.Order
                $targetOrder = $targetItem.Order

                if ($isDropAfterLast) {
                    # Move to end - set Order to current maximum
                    $maxOrder = ($itemsSource | Measure-Object -Property Order -Maximum).Maximum

                    # Shift all items after the dragged item up by one
                    foreach ($item in $itemsSource) {
                        if ($item.Order -gt $draggedOrder) {
                            $item.Order--
                        }
                    }
                    $draggedItem.Order = $maxOrder
                }
                elseif ($draggedOrder -lt $targetOrder) {
                    # Moving down - shift items between old and new position up
                    foreach ($item in $itemsSource) {
                        if ($item.Order -gt $draggedOrder -and $item.Order -le $targetOrder) {
                            $item.Order--
                        }
                    }
                    $draggedItem.Order = $targetOrder
                }
                elseif ($draggedOrder -gt $targetOrder) {
                    # Moving up - shift items between new and old position down
                    foreach ($item in $itemsSource) {
                        if ($item.Order -ge $targetOrder -and $item.Order -lt $draggedOrder) {
                            $item.Order++
                        }
                    }
                    $draggedItem.Order = $targetOrder
                }

                # Refresh the sort to reflect new order
                $sender.Items.SortDescriptions.Clear()
                $sortDescription = New-Object System.ComponentModel.SortDescription("Order", [System.ComponentModel.ListSortDirection]::Ascending)
                $sender.Items.SortDescriptions.Add($sortDescription)
                $sender.Items.Refresh()

                # Save favorites and keep selection
                Save-Favorites -Favorites $itemsSource
                $sender.SelectedItem = $draggedItem
                $sender.ScrollIntoView($draggedItem)
            }
        }

        # Reset drag state
        $script:State.DragDrop.DraggedItem = $null
        $script:State.DragDrop.IsBottomBorder = $false
        $e.Handled = $true
    })
}

# Get the DataGridRow at a specific point in the grid
function Get-DataGridRowFromPoint {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [System.Windows.Point]$point
    )

    # First try hit testing to find row from actual content
    $element = $grid.InputHitTest($point)
    while ($element -ne $null) {
        if ($element -is [System.Windows.Controls.DataGridRow]) {
            return $element
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }

    # If hit testing didn't find a row, iterate through rows to find the closest by Y position
    # This handles padding/margin areas between row content
    if ($grid.Items.Count -gt 0) {
        for ($i = 0; $i -lt $grid.Items.Count; $i++) {
            $row = $grid.ItemContainerGenerator.ContainerFromIndex($i)
            if ($row) {
                $rowPosition = $row.TranslatePoint([System.Windows.Point]::new(0, 0), $grid)
                $rowBottom = $rowPosition.Y + $row.ActualHeight

                # Check if point is within this row's vertical bounds
                if ($point.Y -ge $rowPosition.Y -and $point.Y -lt $rowBottom) {
                    return $row
                }
            }
        }
    }

    return $null
}

# Check if the mouse position is below the last row (for dropping at the end)
function Test-IsPositionBelowLastRow {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [System.Windows.Point]$position
    )

    if ($grid.Items.Count -eq 0) {
        return $false
    }

    $lastRow = $grid.ItemContainerGenerator.ContainerFromIndex($grid.Items.Count - 1)
    if ($lastRow) {
        $lastRowTop = $lastRow.TranslatePoint([System.Windows.Point]::new(0, 0), $grid).Y
        return $position.Y -gt $lastRowTop
    }

    return $false
}

# Set visual feedback for drop target
function Set-DropHighlight {
    param (
        [System.Windows.Controls.DataGridRow]$row,
        [bool]$isBottomBorder = $false
    )

    if ($row) {
        # Use app theme color from XAML resources
        $row.BorderBrush = $script:UI.Window.FindResource("AppPrimaryBrush")
        if ($isBottomBorder) {
            # Highlight bottom border for "drop after last item"
            $row.BorderThickness = New-Object System.Windows.Thickness(0, 0, 0, 2)
        } else {
            # Highlight top border for normal drops
            $row.BorderThickness = New-Object System.Windows.Thickness(0, 2, 0, 0)
        }
        $script:State.DragDrop.LastHighlightedRow = $row
    }
}

# Clear drop target visual feedback
function Clear-DropHighlight {
    if ($script:State.DragDrop.LastHighlightedRow) {
        $script:State.DragDrop.LastHighlightedRow.BorderThickness = New-Object System.Windows.Thickness(0)
        $script:State.DragDrop.LastHighlightedRow = $null
    }
}
# Create new datagrid element for the main window
function New-DataGrid {
    param (
        [string]$name,
        [System.Collections.ObjectModel.ObservableCollection[Object]]$itemsSource
    )

    $grid = New-DataGridBase -Name $name -ItemsSource $itemsSource
    
    $isFavorites = $name -eq "*"
    $propertyType = Get-GridPropertyType -Name $name -ItemsSource $itemsSource
    
    Add-GridColumns -Grid $grid -PropertyType $propertyType -IsFavorites $isFavorites
    Set-GridExtraColumnsVisibility -Grid $grid -TabHeader $name
    Set-GridSorting -Grid $grid -IsFavorites $isFavorites
    Add-GridValidation -Grid $grid -IsFavorites $isFavorites
    
    return $grid
}

function New-DataGridBase {
    param (
        [string]$name,
        [System.Collections.ObjectModel.ObservableCollection[Object]]$itemsSource
    )

    $grid = New-Object System.Windows.Controls.DataGrid
    $grid.Name = $name.Replace("*", "_")
    $grid.Margin = New-Object System.Windows.Thickness(5)
    $grid.ItemsSource = $itemsSource
    $grid.CanUserAddRows = $false
    $grid.IsReadOnly = $script:State.TabsReadOnly

    # Create context menu
    $contextMenu = New-Object System.Windows.Controls.ContextMenu
    $contextMenuStyle = $script:UI.Window.FindResource("GridContextMenuStyle")
    $contextMenu.Style = $contextMenuStyle
    $menuItemStyle = $script:UI.Window.FindResource("GridContextMenuItemStyle")
    $iconStyle = $script:UI.Window.FindResource("ContextMenuIconStyle")

    if ($name -eq "*") {
        # Favorites tab - simplified menu items (drag-and-drop handles reordering)
        $openMenuItem = New-Object System.Windows.Controls.MenuItem
        $openMenuItem.Header = "Open"
        $openMenuItem.Style = $menuItemStyle
        $openIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $openIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::FileDocumentEditOutline
        $openIcon.Style = $iconStyle
        $openMenuItem.Icon = $openIcon
        $openMenuItem.Add_Click({
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($openMenuItem)

        $runAttachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runAttachedMenuItem.Header = "Run (Attached)"
        $runAttachedMenuItem.Style = $menuItemStyle
        $runAttachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runAttachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Play
        $runAttachedIcon.Style = $iconStyle
        $runAttachedMenuItem.Icon = $runAttachedIcon
        $runAttachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runAttachedMenuItem)

        $runDetachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runDetachedMenuItem.Header = "Run (Detached)"
        $runDetachedMenuItem.Style = $menuItemStyle
        $runDetachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runDetachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::OpenInNew
        $runDetachedIcon.Style = $iconStyle
        $runDetachedMenuItem.Icon = $runDetachedIcon
        $runDetachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $false
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runDetachedMenuItem)

        # Add event handler to update run/open visibility when context menu opens
        $contextMenu.Add_Opened({
            param($sender, $e)
            $currentGrid = $script:UI.TabControl.SelectedItem.Content
            $selectedItem = $currentGrid.SelectedItem
            if ($selectedItem) {
                # Update Run/Open menu item visibility based on SkipParameterSelect
                $openItem = $sender.Tag.OpenMenuItem
                $runAttachedItem = $sender.Tag.RunAttachedMenuItem
                $runDetachedItem = $sender.Tag.RunDetachedMenuItem

                if ($selectedItem.SkipParameterSelect) {
                    # Show Run (Attached) and Run (Detached), hide Open
                    $openItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Visible
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Visible
                } else {
                    # Show Open, hide Run (Attached) and Run (Detached)
                    $openItem.Visibility = [System.Windows.Visibility]::Visible
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
        })

        # Store references for dynamic visibility
        $contextMenu.Tag = @{
            OpenMenuItem = $openMenuItem
            RunAttachedMenuItem = $runAttachedMenuItem
            RunDetachedMenuItem = $runDetachedMenuItem
        }

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $favoriteMenuItem = New-Object System.Windows.Controls.MenuItem
        $favoriteMenuItem.Header = "Remove from Favorites"
        $favoriteMenuItem.Style = $menuItemStyle
        $favIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $favIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::StarOff
        $favIcon.Style = $iconStyle
        $favoriteMenuItem.Icon = $favIcon
        $favoriteMenuItem.Add_Click({ Toggle-CommandFavorite })
        [void]$contextMenu.Items.Add($favoriteMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $duplicateMenuItem = New-Object System.Windows.Controls.MenuItem
        $duplicateMenuItem.Header = "Duplicate Command"
        $duplicateMenuItem.Style = $menuItemStyle
        $duplicateIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $duplicateIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::ContentCopy
        $duplicateIcon.Style = $iconStyle
        $duplicateMenuItem.Icon = $duplicateIcon
        $duplicateMenuItem.Add_Click({ Duplicate-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($duplicateMenuItem)
    } else {
        # Regular tabs - standard menu items
        $openMenuItem = New-Object System.Windows.Controls.MenuItem
        $openMenuItem.Header = "Open"
        $openMenuItem.Style = $menuItemStyle
        $openIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $openIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::FileDocumentEditOutline
        $openIcon.Style = $iconStyle
        $openMenuItem.Icon = $openIcon
        $openMenuItem.Add_Click({
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($openMenuItem)

        $runAttachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runAttachedMenuItem.Header = "Run (Attached)"
        $runAttachedMenuItem.Style = $menuItemStyle
        $runAttachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runAttachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Play
        $runAttachedIcon.Style = $iconStyle
        $runAttachedMenuItem.Icon = $runAttachedIcon
        $runAttachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runAttachedMenuItem)

        $runDetachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runDetachedMenuItem.Header = "Run (Detached)"
        $runDetachedMenuItem.Style = $menuItemStyle
        $runDetachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runDetachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::OpenInNew
        $runDetachedIcon.Style = $iconStyle
        $runDetachedMenuItem.Icon = $runDetachedIcon
        $runDetachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $false
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runDetachedMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $favoriteMenuItem = New-Object System.Windows.Controls.MenuItem
        $favoriteMenuItem.Header = "Add to Favorites"
        $favoriteMenuItem.Style = $menuItemStyle
        $favIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $favIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Star
        $favIcon.Style = $iconStyle
        $favoriteMenuItem.Icon = $favIcon
        $favoriteMenuItem.Add_Click({ Toggle-CommandFavorite })

        # Store reference to favorite menu item and run/open items so we can update them
        $contextMenu.Tag = @{
            FavoriteMenuItem = $favoriteMenuItem
            IconStyle = $iconStyle
            OpenMenuItem = $openMenuItem
            RunAttachedMenuItem = $runAttachedMenuItem
            RunDetachedMenuItem = $runDetachedMenuItem
        }

        # Add event handler to update the favorite menu item text/icon and run/open visibility when context menu opens
        $contextMenu.Add_Opened({
            param($sender, $e)
            $currentGrid = $script:UI.TabControl.SelectedItem.Content
            $selectedItem = $currentGrid.SelectedItem
            if ($selectedItem -and $script:UI.Tabs["Favorites"]) {
                $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
                $existingFavorite = $favorites | Where-Object { $_.Id -eq $selectedItem.Id }

                # Get the favorite menu item from the context menu's tag
                $favMenuItem = $sender.Tag.FavoriteMenuItem
                $style = $sender.Tag.IconStyle

                # Create new icon each time to avoid reference issues
                $newFavIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
                $newFavIcon.Style = $style

                if ($existingFavorite) {
                    $favMenuItem.Header = "Remove from Favorites"
                    $newFavIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::StarOff
                } else {
                    $favMenuItem.Header = "Add to Favorites"
                    $newFavIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Star
                }

                $favMenuItem.Icon = $newFavIcon

                # Update Run/Open menu item visibility based on SkipParameterSelect
                $openItem = $sender.Tag.OpenMenuItem
                $runAttachedItem = $sender.Tag.RunAttachedMenuItem
                $runDetachedItem = $sender.Tag.RunDetachedMenuItem

                if ($selectedItem.SkipParameterSelect) {
                    # Show Run (Attached) and Run (Detached), hide Open
                    $openItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Visible
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Visible
                } else {
                    # Show Open, hide Run (Attached) and Run (Detached)
                    $openItem.Visibility = [System.Windows.Visibility]::Visible
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
        })

        [void]$contextMenu.Items.Add($favoriteMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $addMenuItem = New-Object System.Windows.Controls.MenuItem
        $addMenuItem.Header = "Add Command"
        $addMenuItem.Style = $menuItemStyle
        $addIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $addIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::AddBox
        $addIcon.Style = $iconStyle
        $addMenuItem.Icon = $addIcon
        $addMenuItem.Add_Click({ Add-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($addMenuItem)

        $duplicateMenuItem = New-Object System.Windows.Controls.MenuItem
        $duplicateMenuItem.Header = "Duplicate Command"
        $duplicateMenuItem.Style = $menuItemStyle
        $duplicateIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $duplicateIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::ContentCopy
        $duplicateIcon.Style = $iconStyle
        $duplicateMenuItem.Icon = $duplicateIcon
        $duplicateMenuItem.Add_Click({ Duplicate-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($duplicateMenuItem)

        $removeMenuItem = New-Object System.Windows.Controls.MenuItem
        $removeMenuItem.Header = "Remove Command"
        $removeMenuItem.Style = $menuItemStyle
        $removeIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $removeIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::TrashCan
        $removeIcon.Style = $iconStyle
        $removeMenuItem.Icon = $removeIcon
        $removeMenuItem.Add_Click({ Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($removeMenuItem)

        # Update Tag to include remove menu item
        $contextMenu.Tag.RemoveMenuItem = $removeMenuItem
    }

    # Update context menu when it opens to show/hide and update text for remove menu item
    $contextMenu.Add_Opened({
        param($sender, $e)

        # Only update remove menu item for regular tabs (not Favorites tab)
        if ($sender.Tag.RemoveMenuItem) {
            $currentGrid = $script:UI.TabControl.SelectedItem.Content
            $selectedCount = $currentGrid.SelectedItems.Count
            $removeItem = $sender.Tag.RemoveMenuItem

            if ($selectedCount -eq 0) {
                $removeItem.Visibility = [System.Windows.Visibility]::Collapsed
            } else {
                $removeItem.Visibility = [System.Windows.Visibility]::Visible
                if ($selectedCount -eq 1) {
                    $removeItem.Header = "Remove Command"
                } else {
                    $removeItem.Header = "Remove $selectedCount Commands"
                }
            }
        }
    })

    $grid.ContextMenu = $contextMenu
    $grid.AutoGenerateColumns = $false

    # Add selection changed event to update the Run button text
    $grid.Add_SelectionChanged({
        Update-MainRunButtonText
    })

    # Apply the favorite row style (skip for Favorites tab since it only contains favorites)
    if ($name -ne "*") {
        $rowStyle = $script:UI.Window.FindResource("FavoriteRowStyle")
        $grid.RowStyle = $rowStyle

        # Add event handler to set favorite highlighting when rows are loaded
        $grid.Add_LoadingRow({
            param($sender, $e)
            if ($script:UI.Tabs -and $script:UI.Tabs["Favorites"]) {
                $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
                $favoriteIds = @($favorites | ForEach-Object { $_.Id })

                $rowItem = $e.Row.Item
                if ($favoriteIds -contains $rowItem.Id) {
                    $e.Row.Tag = "IsFavorite"
                }
                else {
                    $e.Row.Tag = $null
                }
            }
        })
    }

    return $grid
}

function Get-GridPropertyType {
    param (
        [string]$name,
        [System.Collections.ObjectModel.ObservableCollection[Object]]$itemsSource
    )
    
    $isFavorites = $name -eq "*"
    if ($isFavorites) {
        return [FavoriteRowData]
    }
    return [RowData]
}

function New-GridColumn {
    param (
        [string]$propertyName,
        [bool]$isFavorites
    )

    # Create a checkbox column for SkipParameterSelect and Log
    if ($propertyName -eq "SkipParameterSelect" -or $propertyName -eq "Log") {
        $column = New-Object System.Windows.Controls.DataGridCheckBoxColumn
        $column.Header = $propertyName
        $binding = New-Object System.Windows.Data.Binding $propertyName
        $binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
        $binding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $column.Binding = $binding
    }
    else {
        $column = New-Object System.Windows.Controls.DataGridTextColumn
        $column.Header = $propertyName
        $column.Binding = New-Object System.Windows.Data.Binding $propertyName
    }

    if ($propertyName -eq "Order") {
        $column.IsReadOnly = $false
        $column.Visibility = $script:State.ExtraColumnsVisibility
    }

    return $column
}

function Add-GridColumns {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [type]$propertyType,
        [bool]$isFavorites
    )

    $properties = $propertyType.GetProperties()
    foreach ($prop in $properties) {
        # Skip the Order property for non-Favorites tabs
        if (-not $isFavorites -and $prop.Name -eq "Order") {
            continue
        }
        
        $column = New-GridColumn -PropertyName $prop.Name -IsFavorites $isFavorites
        $grid.Columns.Add($column)
    }
}

function Add-GridValidation {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [bool]$isFavorites
    )
    
    if ($isFavorites) {
        $grid.Add_CellEditEnding({
            param($sender, $e)
            if ($e.Column.Header -eq "Order") {
                try {
                    $newValue = [int]($e.EditingElement.Text)
                    if ($newValue -lt 1) {
                        $e.Cancel = $true
                        return
                    }
                }
                catch {
                    $e.Cancel = $true
                    return
                }
            }
        })
    }
}

function Set-GridSorting {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [bool]$isFavorites
    )
    
    if (-not $isFavorites) {
        Sort-GridByColumn -Grid $grid -ColumnName "Name"
    } else {
        Sort-GridByColumn -Grid $grid -ColumnName "Order"
    }
}

# Create a new tabitem that contains a datagrid and assign to the main tabcontrol
function New-DataTab {
    param (
        [string]$name,
        [System.Collections.ObjectModel.ObservableCollection[Object]]$itemsSource,
        [System.Windows.Controls.TabControl]$tabControl
    )

    $grid = New-DataGrid -Name $name -ItemsSource $itemsSource
    $tab = New-Tab -Name $name
    $tab.Content = $grid
    [void]$tabControl.Items.Add($tab)
    return $tab
}


# Add a WPF element to a grid
function Add-ToGrid {
    param (
        [System.Windows.Controls.Grid]$grid,
        $element
    )

    [void]$grid.Children.Add($element)
}

# Determine a grid row index of a specific command id on a particular datagrid
function Get-GridIndexOfId {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [int]$id
    )

    $itemsSource = $grid.ItemsSource
    $index = -1
    for ($i = 0; $i -lt $itemsSource.Count; $i++) {
        if ($itemsSource[$i].Id -eq $id) {
            $index = $i
            break
        }
    }
    return $index
}

# Assign the row/column position of a WPF element to a grid
function Set-GridPosition {
    param (
        [System.Windows.Controls.Control]$element,
        [int]$row,
        [int]$column,
        [int]$columnSpan
    )

    if ($row) {
        [System.Windows.Controls.Grid]::SetRow($element, $row)
    }
    if ($column) {
        [System.Windows.Controls.Grid]::SetColumn($element, $column)
    }
    if ($columnSpan) {
        [System.Windows.Controls.Grid]::SetColumnSpan($element, $columnSpan)
    }   
}

# Enable or disable editing of all main datagrids and update the visual status of the edit button to match
function Set-TabsReadOnlyStatus {
    param (
        [hashtable]$tabs
    )

    $script:UI.BtnMenuEdit.IsChecked = $script:State.TabsReadOnly
    $script:State.TabsReadOnly = (-not $script:State.TabsReadOnly)

    # Sync both toggle buttons (MaterialDesign will handle icon switching automatically)
    $script:UI.BtnToggleEditMode.IsChecked = (-not $script:State.TabsReadOnly)

    foreach ($tab in $tabs.GetEnumerator()) {
        $tab.Value.Content.IsReadOnly = $script:State.TabsReadOnly
    }
}

# Show or hide the 'extra columns' on all tabs' grids
function Set-TabsExtraColumnsVisibility {
    param (
        [hashtable]$tabs
    )

    $script:State.ExtraColumnsVisibility = if ($script:State.ExtraColumnsVisibility -eq "Visible") { "Collapsed" } else { "Visible" }
    foreach ($tab in $tabs.GetEnumerator()) {
        Set-GridExtraColumnsVisibility -Grid $tab.Value.Content -TabHeader $tab.Value.Header
    }
}

# Show or hide the 'extra columns' on a single grid
function Set-GridExtraColumnsVisibility {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [string]$tabHeader
    )
    
    foreach ($column in $grid.Columns) {
        # Handle regular extra columns
        foreach ($extraCol in $script:State.ExtraColumns) {
            if ($column.Header -eq $extraCol) {
                $column.Visibility = $script:State.ExtraColumnsVisibility
            }
        }
        
        # Special handling for Order column in Favorites tab
        if ($tabHeader -eq "*" -and $column.Header -eq "Order") {
            $column.Visibility = $script:State.ExtraColumnsVisibility
        }
    }
}

function Update-OrderColumnVisibility {
    param (
        [System.Windows.Controls.TabItem]$selectedTab
    )

    if ($selectedTab.Header -eq "*") {
        $grid = $selectedTab.Content
        $orderColumn = $grid.Columns | Where-Object { $_.Header -eq "Order" }
        if ($orderColumn) {
            $orderColumn.Visibility = $script:State.ExtraColumnsVisibility
        }
    }
}

function Handle-TabSelection {
    param (
        [System.Windows.Controls.TabItem]$selectedTab
    )

    Update-OrderColumnVisibility -SelectedTab $selectedTab
}

# Sort the order of the tabs in tab control alphabetically by their header
function Sort-TabControl {
    param (
        [System.Windows.Controls.TabControl]$tabControl
    )

    # Remember which tab was selected
    $selectedTab = $tabControl.SelectedItem
    
    $favTabItem = $tabControl.Items | Where-Object { $_.Header -eq "*" }
    $allTabItem = $tabControl.Items | Where-Object { $_.Header -eq "All" }
    $sortedTabItems = $tabControl.Items | Where-Object { $_.Header -ne "*" -and $_.Header -ne "All" } | Sort-Object -Property { $_.Header.ToString() }
    
    $tabControl.Items.Clear()
    [void]$tabControl.Items.Add($favTabItem)
    [void]$tabControl.Items.Add($allTabItem)
    foreach ($tabItem in $sortedTabItems) {
        [void]$tabControl.Items.Add($tabItem)
    }
    
    # Restore the selected tab
    $tabControl.SelectedItem = $selectedTab
}

# Sort a grid alphabetically by a specific column
function Sort-GridByColumn {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [string]$columnName
    )

    $grid.Items.SortDescriptions.Clear()
    $sort = New-Object System.ComponentModel.SortDescription($columnName, [System.ComponentModel.ListSortDirection]::Ascending)
    $grid.Items.SortDescriptions.Add($sort)
    $grid.Items.Refresh()
}

# Filter the current tab's grid based on search text
function Invoke-GridFilter {
    param (
        [string]$searchText
    )

    $selectedTab = $script:UI.TabControl.SelectedItem
    if (-not $selectedTab) {
        return
    }

    $grid = $selectedTab.Content
    if (-not $grid) {
        return
    }

    # Get the default view and clear any existing filter
    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($grid.ItemsSource)

    if ([string]::IsNullOrWhiteSpace($searchText)) {
        # If search is empty, clear the filter
        $view.Filter = $null
    }
    else {
        # Create a wildcard pattern (case-insensitive)
        $pattern = "*$searchText*"

        # Set the filter predicate - need to use GetNewClosure() to capture the pattern variable
        $view.Filter = {
            param($item)

            # Check if Name matches
            if ($item.Name -and ($item.Name -like $pattern)) {
                return $true
            }

            # Check if Command matches
            if ($item.Command -and ($item.Command -like $pattern)) {
                return $true
            }

            return $false
        }.GetNewClosure()
    }

    $grid.Items.Refresh()
}
function New-LogMonitorTab {
    param (
        [string]$filePath,
        [System.Windows.Controls.TabControl]$tabControl
    )
    
    try {
        # Create new tab
        $fileName = Split-Path $filePath -Leaf
        $tab = New-Tab -Name $fileName
        
        # Create containing grid with ClipToBounds
        $grid = New-Object System.Windows.Controls.Grid
        $grid.ClipToBounds = $true

        # Create textbox for log content
        $textBox = New-Object System.Windows.Controls.TextBox
        $textBox.IsReadOnly = $true
        $textBox.VerticalScrollBarVisibility = "Visible"
        $textBox.HorizontalScrollBarVisibility = "Auto"
        $textBox.TextWrapping = "NoWrap"
        $textBox.Foreground = "Black"
        
        # Load initial content
        $content = Get-Content -Path $filePath -Raw
        $textBox.Text = $content

        # Add TextBox to Grid
        $grid.Children.Add($textBox)
        
        # Store file path and controls in tab's Tag for potential future use
        $tab.Tag = @{
            FilePath = $filePath
            TextBox = $textBox
            Grid = $grid
        }
        
        $tab.Content = $grid
        
        # Add close button functionality with middle-click
        $tab.Add_PreviewMouseDown({
            param($sender, $e)
            if ($e.MiddleButton -eq 'Pressed') {
                $script:UI.LogTabControl.Items.Remove($sender)
                $e.Handled = $true
            }
        })

        # Add close functionality with right-click
        $tab.Add_PreviewMouseRightButtonDown({
            param($sender, $eventArgs)
            if ($eventArgs.ChangedButton -eq 'Right') {
                $script:UI.LogTabControl.Items.Remove($sender)
                Write-Log "Closed log file: $($sender.Tag.FilePath)"
                $eventArgs.Handled = $true
            }
        })
        
        # Insert the new tab before the "+" tab
        $addTabIndex = $tabControl.Items.Count - 1
        $tabControl.Items.Insert($addTabIndex, $tab)
        $tabControl.SelectedItem = $tab
        
        Write-Log "Opened log file: $filePath"
    }
    catch {
        Show-ErrorMessageBox "Failed to create log monitor tab: $_"
    }
}

function Open-LogFile {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.InitialDirectory = $script:Settings.DefaultLogsPath
    $dialog.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $dialog.FilterIndex = 1

    if ($dialog.ShowDialog()) {
        New-LogMonitorTab -FilePath $dialog.FileName -TabControl $script:UI.LogTabControl
    }
}

# Create a styled context menu for PowerShell tabs
function New-PSTabContextMenu {
    param (
        [System.Windows.Controls.TabItem]$Tab
    )

    $contextMenu = New-Object System.Windows.Controls.ContextMenu

    # Apply the same style as History context menu
    $contextMenu.FontSize = 12

    # Close Tab menu item
    $menuCloseTab = New-Object System.Windows.Controls.MenuItem
    $menuCloseTab.Header = "Close Tab"
    $menuCloseTab.FontSize = 12

    # Create icon for Close Tab
    $iconClose = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconClose.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Close
    $iconClose.Width = 16
    $iconClose.Height = 16
    $iconClose.Margin = New-Object System.Windows.Thickness(0)
    $menuCloseTab.Icon = $iconClose

    $menuCloseTab.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag -and $tab.Tag["Process"]) {
            try {
                $script:UI.PSTabControl.Items.Remove($tab)
                $tab.Tag["Process"].Kill()
                Write-Status "PowerShell tab closed"
            }
            catch {
                Write-ErrorMessage "Failed to close PowerShell tab: $_"
            }
        }
    })
    [void]$contextMenu.Items.Add($menuCloseTab)

    # Close All Tabs menu item
    $menuCloseAllTabs = New-Object System.Windows.Controls.MenuItem
    $menuCloseAllTabs.Header = "Close All Tabs"
    $menuCloseAllTabs.FontSize = 12

    # Create icon for Close All Tabs
    $iconCloseAll = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconCloseAll.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::CloseBoxMultiple
    $iconCloseAll.Width = 16
    $iconCloseAll.Height = 16
    $iconCloseAll.Margin = New-Object System.Windows.Thickness(0)
    $menuCloseAllTabs.Icon = $iconCloseAll

    $menuCloseAllTabs.Add_Click({
        param($menuSender, $menuArgs)
        # Get all tabs except the "+" add tab
        $tabsToClose = @($script:UI.PSTabControl.Items | Where-Object { $_ -ne $script:UI.PSAddTab })

        foreach ($tab in $tabsToClose) {
            if ($tab.Tag -and $tab.Tag["Process"]) {
                try {
                    $script:UI.PSTabControl.Items.Remove($tab)
                    $tab.Tag["Process"].Kill()
                }
                catch {
                    Write-ErrorMessage "Failed to close PowerShell tab: $_"
                }
            }
        }

        if ($tabsToClose.Count -gt 0) {
            Write-Status "Closed $($tabsToClose.Count) PowerShell tab(s)"
        }
    })
    [void]$contextMenu.Items.Add($menuCloseAllTabs)

    # Detach Tab menu item
    $menuDetachTab = New-Object System.Windows.Controls.MenuItem
    $menuDetachTab.Header = "Detach Tab"
    $menuDetachTab.FontSize = 12

    # Create icon for Detach Tab
    $iconDetach = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconDetach.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Export
    $iconDetach.Width = 16
    $iconDetach.Height = 16
    $iconDetach.Margin = New-Object System.Windows.Thickness(0)
    $menuDetachTab.Icon = $iconDetach

    $menuDetachTab.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag) {
            $script:UI.PSTabControl.SelectedItem = $tab
            Detach-CurrentTab
            Write-Status "PowerShell tab detached"
        }
    })
    [void]$contextMenu.Items.Add($menuDetachTab)

    # Open Log menu item
    $menuOpenLog = New-Object System.Windows.Controls.MenuItem
    $menuOpenLog.Header = "Open Log"
    $menuOpenLog.FontSize = 12

    # Create icon for Open Log
    $iconLog = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconLog.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::FileDocumentOutline
    $iconLog.Width = 16
    $iconLog.Height = 16
    $iconLog.Margin = New-Object System.Windows.Thickness(0)
    $menuOpenLog.Icon = $iconLog

    $menuOpenLog.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag -and $tab.Tag["HistoryEntry"]) {
            $historyEntry = $tab.Tag["HistoryEntry"]
            if ($historyEntry.LogPath -and (Test-Path $historyEntry.LogPath)) {
                New-LogMonitorTab -FilePath $historyEntry.LogPath -TabControl $script:UI.LogTabControl

                # Switch to the Logs tab
                $logsTab = $script:UI.TabControlShell.Items | Where-Object { $_.Header -eq "Logs" }
                if ($logsTab) {
                    $script:UI.TabControlShell.SelectedItem = $logsTab
                }

                Write-Status "Opened log file"
            }
            else {
                Write-ErrorMessage "Log file not found or command was not logged"
            }
        }
        else {
            Write-ErrorMessage "No log associated with this tab"
        }
    })
    [void]$contextMenu.Items.Add($menuOpenLog)

    # Go to History menu item
    $menuGoToHistory = New-Object System.Windows.Controls.MenuItem
    $menuGoToHistory.Header = "Go to History"
    $menuGoToHistory.FontSize = 12

    # Create icon for Go to History
    $iconHistory = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconHistory.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::History
    $iconHistory.Width = 16
    $iconHistory.Height = 16
    $iconHistory.Margin = New-Object System.Windows.Thickness(0)
    $menuGoToHistory.Icon = $iconHistory

    $menuGoToHistory.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag -and $tab.Tag["HistoryEntry"]) {
            # Switch to History tab
            $historyTab = $script:UI.TabControlShell.Items | Where-Object { $_.Header -eq "History" }
            if ($historyTab) {
                $script:UI.TabControlShell.SelectedItem = $historyTab

                # Select the history entry in the grid
                $historyGrid = $script:UI.Window.FindName("CommandHistoryGrid")
                if ($historyGrid) {
                    $historyEntry = $tab.Tag["HistoryEntry"]
                    $historyGrid.SelectedItem = $historyEntry
                    $historyGrid.ScrollIntoView($historyEntry)
                    Write-Status "Jumped to command history"
                }
            }
        }
        else {
            Write-ErrorMessage "No history entry associated with this tab"
        }
    })
    [void]$contextMenu.Items.Add($menuGoToHistory)

    return $contextMenu
}

# Create a new embedded process under a Tab Control
function New-ProcessTab {
    param (
        $tabControl,
        $process,
        $processArgs,
        $tabName = "PS_$($tabControl.Items.Count)",
        [PSCustomObject]$historyEntry = $null
    )

    $proc = Start-Process $process -WindowStyle Hidden -PassThru -ArgumentList $processArgs

    Start-Sleep -Seconds 2

    # Find the window handle of the PowerShell process using process ID
    $psHandle = [Win32]::FindWindowByProcessId($proc.Id)
    if ($psHandle -eq [IntPtr]::Zero) {
        Write-Log "Failed to retrieve the PowerShell window handle for process ID: $($proc.Id)."
        return
    }

    $tab = New-Tab -Name $tabName
    $tabData = @{}
    $tabData["Handle"] = $psHandle
    $tabData["Process"] = $proc
    $tabData["HistoryEntry"] = $historyEntry
    $tab.Tag = $tabData

    # Create a WindowsFormsHost and a Panel to host the PowerShell window
    $windowsFormsHost = New-Object System.Windows.Forms.Integration.WindowsFormsHost
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.BackColor = [System.Drawing.Color]::Black
    $windowsFormsHost.Child = $panel
    $tab.Content = $windowsFormsHost

    # Add the TabItem to the TabControl before the "New Tab" tab
    $tabControl.Items.Insert($tabControl.Items.Count - 1, $tab)
    $tabControl.Dispatcher.Invoke([action]{$tabControl.SelectedItem = $tab})
    #$tabControl.SelectedItem = $tab

    # Remove window frame (title bar, borders) by modifying window style
    $currentStyle = [Win32]::GetWindowLong($psHandle, $script:GWL_STYLE)
    [Win32]::SetWindowLong($psHandle, $script:GWL_STYLE, $currentStyle -band -0x00C00000)  # Remove WS_CAPTION and WS_THICKFRAME

    # Re-parent the PowerShell window to the panel
    [Win32]::SetParent($psHandle, $panel.Handle)
    [Win32]::ShowWindow($psHandle, 5)  # 5 = SW_SHOW
    [Win32]::MoveWindow($psHandle, 0, 0, $panel.Width, $panel.Height, $true)
    
    # Handle resizing
    $panel.Add_SizeChanged({
        param($sender, $eventArgs)
        $handle = $script:UI.PSTabControl.SelectedItem.Tag["Handle"]
        if ($handle -ne [IntPtr]::Zero) {
            [Win32]::MoveWindow($handle, 0, 0, $sender.Width, $sender.Height, $true)
        }
        else {
            Write-Log "Invalid window handle in SizeChanged event."
        }
    })

    # Add context menu to the tab header
    $tab.ContextMenu = New-PSTabContextMenu -Tab $tab

    # Handle middle-click to detach tab
    $tab.Add_PreviewMouseDown({
        param($sender, $e)
        if ($e.MiddleButton -eq 'Pressed') {
            Detach-CurrentTab
            $e.Handled = $true
        }
    })
}

# Detach and unparent an embedded process so it is running outside of PSGUI
function Detach-CurrentTab {
    $selectedTab = $script:UI.PSTabControl.SelectedItem
    if ($selectedTab -and $selectedTab -ne $script:UI.PSAddTab) {
        $psHandle = $selectedTab.Tag["Handle"]
        $proc = $selectedTab.Tag["Process"]

        # Restore window styles
        $style = [Win32]::GetWindowLong($psHandle, $script:GWL_STYLE)
        $style = $style -bor $script:WS_OVERLAPPEDWINDOW
        [Win32]::SetWindowLong($psHandle, $script:GWL_STYLE, $style)

        # Detach from parent
        [Win32]::SetParent($psHandle, [IntPtr]::Zero)

        # Show window
        [Win32]::ShowWindow($psHandle, 1)  # 1 = SW_SHOWNORMAL
        [Win32]::SetWindowPos($psHandle, [IntPtr]::Zero, 100, 100, 800, 600, 0x0040)  # SWP_SHOWWINDOW

        # Remove the tab
        $script:UI.PSTabControl.Items.Remove($selectedTab)
    }
}

# Popup window to select Detached PowerShell windows to attach
function Show-AttachWindow {
    $attachWindow = New-Object System.Windows.Window
    $attachWindow.Title = "Attach PowerShell Window"
    $attachWindow.Width = 400
    $attachWindow.Height = 300
    $attachWindow.WindowStartupLocation = "CenterOwner"
    $attachWindow.Owner = $script:UI.Window

    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $attachWindow.Content = $stackPanel

    $listBox = New-Object System.Windows.Controls.ListBox
    $listBox.Margin = New-Object System.Windows.Thickness(10)
    $stackPanel.Children.Add($listBox)

    # Find PowerShell windows
    $psWindows = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and ($_.ProcessName -eq "powershell" -or $_.ProcessName -eq "pwsh") }
    foreach ($window in $psWindows) {
        $item = New-Object System.Windows.Controls.ListBoxItem
        $item.Content = "$($window.Id): $($window.MainWindowTitle)"
        $item.Tag = $window
        $listBox.Items.Add($item)
    }

    $btnAttach = New-Object System.Windows.Controls.Button
    $btnAttach.Content = "Attach"
    $btnAttach.Margin = New-Object System.Windows.Thickness(10)
    $stackPanel.Children.Add($btnAttach)

    $btnAttach.Add_Click({
        $selectedItem = $listBox.SelectedItem
        if ($selectedItem) {
            $proc = $selectedItem.Tag
            Attach-DetachedWindow -Process $proc
            $attachWindow.Close()
        }
    })

    $attachWindow.ShowDialog()
}

# Attach and reparent an Detached window as an embedded tab
function Attach-DetachedWindow {
    param (
        [System.Diagnostics.Process]$Process
    )

    $psHandle = $Process.MainWindowHandle
    
    # Remove window frame
    $style = [Win32]::GetWindowLong($psHandle, $script:GWL_STYLE)
    $style = $style -band -bnot $script:WS_OVERLAPPEDWINDOW
    [Win32]::SetWindowLong($psHandle, $script:GWL_STYLE, $style)

    $tab = New-Tab -Name "PS_$($script:UI.PSTabControl.Items.Count)"
    $tabData = @{}
    $tabData["Handle"] = $psHandle
    $tabData["Process"] = $Process
    $tab.Tag = $tabData

    $windowsFormsHost = New-Object System.Windows.Forms.Integration.WindowsFormsHost
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $panel.BackColor = [System.Drawing.Color]::Black
    $windowsFormsHost.Child = $panel
    $tab.Content = $windowsFormsHost

    $script:UI.PSTabControl.Items.Insert($script:UI.PSTabControl.Items.Count - 1, $tab)
    $script:UI.PSTabControl.SelectedItem = $tab

    # Re-parent the PowerShell window
    [Win32]::SetParent($psHandle, $panel.Handle)
    [Win32]::ShowWindow($psHandle, 5)  # 5 = SW_SHOW
    [Win32]::MoveWindow($psHandle, 0, 0, $panel.Width, $panel.Height, $true)

    # Handle resizing
    $panel.Add_SizeChanged({
        param($sender, $eventArgs)
        $handle = $script:UI.PSTabControl.SelectedItem.Tag["Handle"]
        if ($handle -ne [IntPtr]::Zero) {
            [Win32]::MoveWindow($handle, 0, 0, $sender.Width, $sender.Height, $true)
        }
        else {
            Write-Log "Invalid window handle in SizeChanged event."
        }
    })

    # Add context menu to the tab header
    $tab.ContextMenu = New-PSTabContextMenu -Tab $tab

    # Handle middle-click to detach tab
    $tab.Add_PreviewMouseDown({
        param($sender, $e)
        if ($e.MiddleButton -eq 'Pressed') {
            Detach-CurrentTab
            $e.Handled = $true
        }
    })
}
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
    $script:UI.TxtStatusTimeout.Text = $script:Settings.StatusTimeout
    $script:UI.TxtSettingsPath.Text = $script:Settings.SettingsPath
    $script:UI.TxtFavoritesPath.Text = $script:Settings.FavoritesPath
    $script:UI.TxtDefaultHistoryPath.Text = $script:Settings.DefaultHistoryPath
    $script:UI.ChkSaveHistory.IsChecked = $script:Settings.SaveHistory
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
    $script:Settings.DefaultHistoryPath = $script:UI.TxtDefaultHistoryPath.Text
    $script:Settings.SaveHistory = $script:UI.ChkSaveHistory.IsChecked
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

    # Validate and set StatusTimeout
    $statusTimeout = 6  # Default value
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
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.FileName = $script:UI.TxtSettingsPath.Text
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    if ($dialog.ShowDialog()) {
        $script:UI.TxtSettingsPath.Text = $dialog.FileName
    }
}

function Invoke-BrowseFavorites {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.FileName = $script:UI.TxtFavoritesPath.Text
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    if ($dialog.ShowDialog()) {
        $script:UI.TxtFavoritesPath.Text = $dialog.FileName
    }
}

function Invoke-BrowseHistory {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.FileName = $script:UI.TxtDefaultHistoryPath.Text
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    if ($dialog.ShowDialog()) {
        $script:UI.TxtDefaultHistoryPath.Text = $dialog.FileName
    }
}

function Invoke-BrowseDataFile {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.FileName = $script:UI.TxtDefaultDataFile.Text
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dialog.DefaultExt = ".json"
    if ($dialog.ShowDialog()) {
        $script:UI.TxtDefaultDataFile.Text = $dialog.FileName
    }
}
# Create a new WPF window from an XML file and load all WPF elements and return them under one variable
function New-Window {
    param (
        [string]$filePath
    )

    try {
        [xml]$xaml = (Get-Content $filePath)
        $window = New-Object System.Collections.Hashtable
        $nodeReader = [System.Xml.XmlNodeReader]::New($xaml)
        $xamlReader = [Windows.Markup.XamlReader]::Load($nodeReader)
        [void]$window.Add('Window', $xamlReader)
        $elements = $xaml.SelectNodes("//*[@Name]")
        foreach ($element in $elements) {
            $varName = $element.Name
            $varValue = $window.Window.FindName($Element.Name)
            [void]$window.Add($varName, $varValue)
        }
        return $window
    }
    catch {
        Show-ErrorMessageBox("Error building Xaml data or loading window data.`n$_")
        exit
    }
}

# Create a new CommandWindow instance with its own state
function New-CommandWindow {
    param (
        [Command]$command
    )

    try {
        # Load the CommandWindow XAML
        $commandWindow = New-Window -FilePath $script:ApplicationPaths.CommandWindowXamlFile

        # Set the owner to the main window for proper modal behavior
        $commandWindow.Window.Owner = $script:UI.Window
        $commandWindow.Window.Icon = $script:ApplicationPaths.IconFile
        $commandWindow.Window.Title = $command.Root

        # Store the command object in the window's Tag for later access
        $commandWindow.Window.Tag = @{
            Command = $command
        }

        # Register event handlers for this specific window

        $commandWindow.BtnCommandRunAttached.Add_Click({
            param($sender, $e)
            $window = $sender.Parent
            while ($window -and $window -isnot [System.Windows.Window]) {
                $window = $window.Parent
            }
            if ($window) {
                Invoke-CommandRunClick -CommandWindow $window -RunAttached $true
            }
        })

        $commandWindow.BtnCommandRunDetached.Add_Click({
            param($sender, $e)
            $window = $sender.Parent
            while ($window -and $window -isnot [System.Windows.Window]) {
                $window = $window.Parent
            }
            if ($window) {
                Invoke-CommandRunClick -CommandWindow $window -RunAttached $false
            }
        })

        $commandWindow.BtnCommandCopyToClipboard.Add_Click({
            param($sender, $e)
            $window = $sender.Parent
            while ($window -and $window -isnot [System.Windows.Window]) {
                $window = $window.Parent
            }
            if ($window) {
                Invoke-CommandCopyToClipboard -CommandWindow $window
            }
        })

        $commandWindow.BtnCommandHelp.Add_Click({
            param($sender, $e)
            $window = $sender.Parent
            while ($window -and $window -isnot [System.Windows.Window]) {
                $window = $window.Parent
            }
            if ($window) {
                $cmd = $window.Tag.Command
                if ($cmd) {
                    Get-Help -Name $cmd.Root -ShowWindow
                }
            }
        })

        $commandWindow.BtnToggleCommonParameters.Add_Click({
            param($sender, $e)
            $window = $sender.Parent
            while ($window -and $window -isnot [System.Windows.Window]) {
                $window = $window.Parent
            }
            if ($window) {
                Toggle-CommonParametersGrid -CommandWindow $window
            }
        })

        # Handle window closing to remove from tracking
        $commandWindow.Window.Add_Closed({
            param($sender, $e)
            $script:State.OpenCommandWindows.Remove($sender)
        })

        # Add to tracking collection
        $script:State.OpenCommandWindows.Add($commandWindow.Window)

        return $commandWindow
    }
    catch {
        Show-ErrorMessageBox("Error creating CommandWindow: $_")
        return $null
    }
}

# Create new tabitem element
function New-Tab {
    param (
        [string]$name
    )

    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header = $name
    return $tabItem
}

# Create a new text label element
function New-Label {
    param (
        [string]$content,
        [string]$halign,
        [string]$valign
    )

    $label = New-Object System.Windows.Controls.Label
    $label.Content = $content
    $label.HorizontalAlignment = $halign
    $label.VerticalAlignment = $valign
    $label.Margin = New-Object System.Windows.Thickness(3)
    return $label
}

# Create a new tooltip element
function New-ToolTip {
    param (
        [string]$content
    )

    $tooltip = New-Object System.Windows.Controls.ToolTip
    $tooltip.Content = $content
    return $tooltip
}

# Create a new combo box element
function New-ComboBox {
    param (
        [string]$name,
        [System.String[]]$itemsSource,
        [string]$selectedItem
    )

    $comboBox = New-Object System.Windows.Controls.ComboBox
    $comboBox.Name = $name
    $comboBox.Margin = New-Object System.Windows.Thickness(5)
    $comboBox.ItemsSource = $itemsSource
    $comboBox.SelectedItem = $selectedItem
    return $comboBox
}

# Create a new text box element
function New-TextBox {
    param (
        [string]$name,
        [string]$text
    )

    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Name = $name
    $textBox.Margin = New-Object System.Windows.Thickness(5)
    $textBox.Text = $text
    return $textBox
}

# Create a new check box element
function New-CheckBox {
    param (
        [string]$name,
        [bool]$isChecked
    )

    $checkbox = New-Object System.Windows.Controls.CheckBox
    $checkbox.Name = $name
    $checkbox.IsChecked = $isChecked
    return $checkbox
}

# Create a new button element
function New-Button {
    param (
        [string]$content,
        [string]$halign,
        [int]$width
    )
    
    $button = New-Object System.Windows.Controls.Button
    $button.Content = $content
    $button.Margin = New-Object System.Windows.Thickness(10)
    $button.HorizontalAlignment = $halign
    $button.Width = $width
    $button.IsDefault = $true
    return $button
}
# Determine the current highest Id that exists in the collection
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
# App Version
$script:Version = "1.4.0"
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
    HighestId = 0
    FavoritesHighestOrder = 0
    TabsReadOnly = $true
    RunCommandAttached = $script:Settings.DefaultRunCommandAttached
    ExtraColumnsVisibility = "Collapsed"
    ExtraColumns = @("Id", "Command", "SkipParameterSelect", "PreCommand", "Log")
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

# Load Win32 API functions from Detached file
Add-Type -Path $script:ApplicationPaths.Win32APIFile

Start-MainWindow
