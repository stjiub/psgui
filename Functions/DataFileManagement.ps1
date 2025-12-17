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

                    # Migrate old integer IDs to GUIDs
                    if ($item.Id -is [int] -or $item.Id -match '^\d+$') {
                        $rowData.Id = Get-UniqueCommandId
                        Write-Log "Migrated integer ID $($item.Id) to GUID: $($rowData.Id)"
                    } else {
                        $rowData.Id = $item.Id
                    }

                    $rowData.Name = $item.Name
                    $rowData.Description = $item.Description
                    $rowData.Category = $item.Category
                    $rowData.Command = $item.Command
                    $rowData.SkipParameterSelect = $item.SkipParameterSelect
                    $rowData.PreCommand = $item.PreCommand
                    $rowData.PostCommand = $item.PostCommand

                    # Migrate old Log property to new Transcript/PSTask checkboxes
                    if ($null -ne $item.Log) {
                        # Old data format - migrate from Log string to checkboxes
                        if ($item.Log -is [string]) {
                            $rowData.Transcript = ($item.Log -eq "Transcript")
                            $rowData.PSTask = ($item.Log -eq "PSTask")
                            Write-Log "Migrated Log value from string ($($item.Log)) to Transcript/PSTask checkboxes for command: $($item.Name)"
                        }
                        # Very old format - bool converted to Transcript
                        elseif ($item.Log -is [bool]) {
                            $rowData.Transcript = $item.Log
                            $rowData.PSTask = $false
                            Write-Log "Migrated Log value from bool ($($item.Log)) to Transcript checkbox for command: $($item.Name)"
                        }
                    }
                    else {
                        # New data format - use Transcript/PSTask directly
                        $rowData.Transcript = if ($null -ne $item.Transcript) { $item.Transcript } else { $false }
                        $rowData.PSTask = if ($null -ne $item.PSTask) { $item.PSTask } else { $false }
                        $rowData.PSTaskMode = $item.PSTaskMode
                        $rowData.PSTaskVisibilityLevel = $item.PSTaskVisibilityLevel
                    }

                    $rowData.ShellOverride = $item.ShellOverride
                    $rowData.LogParameterNames = $item.LogParameterNames
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
                PostCommand = $_.PostCommand
                Transcript = $_.Transcript
                PSTask = $_.PSTask
                PSTaskMode = $_.PSTaskMode
                PSTaskVisibilityLevel = $_.PSTaskVisibilityLevel
                ShellOverride = $_.ShellOverride
                LogParameterNames = $_.LogParameterNames
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

        # Update last modification time to prevent syncing our own changes
        if (Test-Path $filePath) {
            $fileInfo = Get-Item $filePath
            $script:State.DataFileLastModTime = $fileInfo.LastWriteTime
        }
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

            # Check if item with same ID already exists (should be rare with GUIDs)
            $existingItem = $allData | Where-Object { $_.Id -eq $item.Id }
            if ($existingItem) {
                # Generate a new unique ID to avoid conflicts
                $item.Id = Get-UniqueCommandId
                Write-Log "ID conflict detected during import, generated new ID: $($item.Id)"
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

function Sync-DataFile {
    try {
        $dataFilePath = $script:State.CurrentDataFile

        if ([string]::IsNullOrWhiteSpace($dataFilePath)) {
            return
        }

        if (-not (Test-Path $dataFilePath)) {
            return
        }

        # Get current file modification time
        $fileInfo = Get-Item $dataFilePath
        $currentModTime = $fileInfo.LastWriteTime

        # Initialize last mod time if not set
        if (-not $script:State.DataFileLastModTime) {
            $script:State.DataFileLastModTime = $currentModTime
            return
        }

        # Check if file has been modified externally
        if ($currentModTime -le $script:State.DataFileLastModTime) {
            return
        }

        # Skip sync if we have unsaved changes (user's changes take priority)
        if ($script:State.HasUnsavedChanges) {
            return
        }

        # Load the external data file
        $externalData = Load-DataFile $dataFilePath

        if (-not $externalData) {
            $script:State.DataFileLastModTime = $currentModTime
            return
        }

        # Get current data
        $currentData = $script:UI.Tabs["All"].Content.ItemsSource

        # Create hashtables for comparison by ID
        $currentById = @{}
        foreach ($item in $currentData) {
            if ($item.Id) {
                $currentById[$item.Id] = $item
            }
        }

        $externalById = @{}
        foreach ($item in $externalData) {
            if ($item.Id) {
                $externalById[$item.Id] = $item
            }
        }

        # Track changes
        $itemsAdded = 0
        $itemsRemoved = 0
        $itemsUpdated = 0

        # Add or update items from external file (updates happen in background)
        foreach ($externalItem in $externalData) {
            if (-not $externalItem.Id) { continue }

            $currentItem = $currentById[$externalItem.Id]
            if ($currentItem) {
                # Item exists - update properties if different
                $hasChanges = $false
                if ($currentItem.Name -ne $externalItem.Name) { $currentItem.Name = $externalItem.Name; $hasChanges = $true }
                if ($currentItem.Description -ne $externalItem.Description) { $currentItem.Description = $externalItem.Description; $hasChanges = $true }
                if ($currentItem.Category -ne $externalItem.Category) { $currentItem.Category = $externalItem.Category; $hasChanges = $true }
                if ($currentItem.Command -ne $externalItem.Command) { $currentItem.Command = $externalItem.Command; $hasChanges = $true }
                if ($currentItem.SkipParameterSelect -ne $externalItem.SkipParameterSelect) { $currentItem.SkipParameterSelect = $externalItem.SkipParameterSelect; $hasChanges = $true }
                if ($currentItem.PreCommand -ne $externalItem.PreCommand) { $currentItem.PreCommand = $externalItem.PreCommand; $hasChanges = $true }
                if ($currentItem.PostCommand -ne $externalItem.PostCommand) { $currentItem.PostCommand = $externalItem.PostCommand; $hasChanges = $true }
                if ($currentItem.Transcript -ne $externalItem.Transcript) { $currentItem.Transcript = $externalItem.Transcript; $hasChanges = $true }
                if ($currentItem.PSTask -ne $externalItem.PSTask) { $currentItem.PSTask = $externalItem.PSTask; $hasChanges = $true }
                if ($currentItem.PSTaskMode -ne $externalItem.PSTaskMode) { $currentItem.PSTaskMode = $externalItem.PSTaskMode; $hasChanges = $true }
                if ($currentItem.PSTaskVisibilityLevel -ne $externalItem.PSTaskVisibilityLevel) { $currentItem.PSTaskVisibilityLevel = $externalItem.PSTaskVisibilityLevel; $hasChanges = $true }
                if ($currentItem.ShellOverride -ne $externalItem.ShellOverride) { $currentItem.ShellOverride = $externalItem.ShellOverride; $hasChanges = $true }
                if ($currentItem.LogParameterNames -ne $externalItem.LogParameterNames) { $currentItem.LogParameterNames = $externalItem.LogParameterNames; $hasChanges = $true }

                if ($hasChanges) {
                    $itemsUpdated++
                }
            }
            else {
                # New item - add it to All tab
                $currentData.Add($externalItem)
                $itemsAdded++

                # Also add to category tab if it exists
                if ($externalItem.Category -and $externalItem.Category -ne "") {
                    $categoryTab = $script:UI.Tabs[$externalItem.Category]
                    if ($categoryTab -and $categoryTab.Content -and $categoryTab.Content.ItemsSource) {
                        $categoryTab.Content.ItemsSource.Add($externalItem)
                    }
                }
            }
        }

        # Remove items that exist locally but not in external file
        $itemsToRemove = @()
        foreach ($currentItem in $currentData) {
            if ($currentItem.Id -and -not $externalById.ContainsKey($currentItem.Id)) {
                $itemsToRemove += $currentItem
            }
        }
        foreach ($item in $itemsToRemove) {
            # Remove from All tab
            $currentData.Remove($item)
            $itemsRemoved++

            # Also remove from category tab if it has one
            if ($item.Category -and $item.Category -ne "") {
                $categoryTab = $script:UI.Tabs[$item.Category]
                if ($categoryTab -and $categoryTab.Content -and $categoryTab.Content.ItemsSource) {
                    $categoryTab.Content.ItemsSource.Remove($item)
                }
            }
        }

        # Log changes (updates happen silently in background)
        if ($itemsAdded -gt 0 -or $itemsRemoved -gt 0 -or $itemsUpdated -gt 0) {
            Write-Log "Data file synced in background: $itemsAdded added, $itemsRemoved removed, $itemsUpdated updated"
        }

        # Update last modification time
        $script:State.DataFileLastModTime = $currentModTime
    }
    catch {
        Write-Log "ERROR syncing data file: $_"
    }
}

function Start-DataFileSyncTimer {
    try {
        # Only start if sync interval is greater than 0
        if ($script:Settings.DataFileSyncIntervalSeconds -le 0) {
            return
        }

        # Initialize the last modification time tracker
        $dataFilePath = $script:State.CurrentDataFile
        if ((Test-Path $dataFilePath)) {
            $fileInfo = Get-Item $dataFilePath
            $script:State.DataFileLastModTime = $fileInfo.LastWriteTime
        }

        # Create a DispatcherTimer for periodic sync
        $script:DataFileSyncTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:DataFileSyncTimer.Interval = [TimeSpan]::FromSeconds($script:Settings.DataFileSyncIntervalSeconds)

        $script:DataFileSyncTimer.Add_Tick({
            Sync-DataFile
        })

        $script:DataFileSyncTimer.Start()
    }
    catch {
        Write-Log "ERROR starting data file sync timer: $_"
    }
}

function Stop-DataFileSyncTimer {
    try {
        if ($script:DataFileSyncTimer) {
            $script:DataFileSyncTimer.Stop()
            $script:DataFileSyncTimer = $null
        }
    }
    catch {
        Write-Log "ERROR stopping data file sync timer: $_"
    }
}