# App Version
$script:Version = "1.3.7"
$script:AppTitle = "PSGUI - v$($script:Version)"

# Constants
$script:GWL_STYLE = -16
$script:WS_BORDERLESS = 0x800000  # WS_POPUP without WS_BORDER, WS_CAPTION, etc.
$script:WS_OVERLAPPEDWINDOW = 0x00CF0000

# Settings
$script:Settings = @{
    DefaultShell = "powershell"
    DefaultShellArgs = "-ExecutionPolicy Bypass -NoExit -Command `" & { [System.Console]::Title = 'PS' } `""
    DefaultRunCommandInternal = $true
    OpenShellAtStart = $false
    StatusTimeout = 3
    DefaultLogsPath = Join-Path $env:APPDATA "PSGUI"
    SettingsPath = Join-Path $env:APPDATA "PSGUI\settings.json"
    FavoritesPath = Join-Path $env:APPDATA "PSGUI\favorites.json"
    ShowDebugTab = $false
    DefaultDataFile = Join-Path $env:APPDATA "PSGUI\data.json"
}

# Initialize variables and load resources for application 
function Initialize-Application() {
    # Determine app pathing whether running as PS script or EXE
    $script:Path = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0])) }

    $script:ApplicationPaths = @{
        MainWindowXamlFile = Join-Path $script:Path "MainWindow.xaml"
        MaterialDesignThemes = Join-Path $script:Path "Assembly\MaterialDesignThemes.Wpf.dll"
        MaterialDesignColors = Join-Path $script:Path "Assembly\MaterialDesignColors.dll"
        DefaultDataFile = Join-Path $script:Path "data.json"
        SettingsFilePath = Join-Path $env:APPDATA "PSGUI\settings.json"
        IconFile = Join-Path $script:Path "icon.ico"
        Win32APIFile = Join-Path $script:Path "Win32API.cs"
    }

    $script:State = @{
        CurrentDataFile = $null
        CurrentCommand = $null
        LastCommand = $null
        HighestId = 0
        FavoritesHighestOrder = 0
        TabsReadOnly = $true
        RunCommandInternal = $script:Settings.DefaultRunCommandInternal
        ExtraColumnsVisibility = "Collapsed"
        ExtraColumns = @("Id", "Command", "SkipParameterSelect", "PreCommand")
        SubGridExpandedHeight = 300
        HasUnsavedChanges = $false
        CurrentCommandListId = $null
        DragDrop = @{
            DraggedItem = $null
            LastHighlightedRow = $null
            IsBottomBorder = $false
        }
    }

    # Load necessary assemblies
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsFormsIntegration
    [Void][System.Reflection.Assembly]::LoadFrom($script:ApplicationPaths.MaterialDesignThemes)
    [Void][System.Reflection.Assembly]::LoadFrom($script:ApplicationPaths.MaterialDesignColors)

    # Load Win32 API functions from external file
    Add-Type -Path $script:ApplicationPaths.Win32APIFile
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
        elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem -and -not $sender.IsInEditMode) {
            $e.Handled = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        }
    })
    $script:UI.Tabs.Add("All", $allTab)

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
        elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem -and -not $sender.IsInEditMode) {
            $e.Handled = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        }
    })

    # Add drag/drop event handlers for reordering favorites
    Initialize-FavoritesDragDrop -Grid $favTab.Content

    $script:UI.Tabs.Add("Favorites", $favTab)
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
            elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem -and -not $sender.IsInEditMode) {
                $e.Handled = $true
                Invoke-MainRunClick -TabControl $script:UI.TabControl
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
    $script:UI.BtnMenuRemove.Add_Click({ Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuSave.Add_Click({ Save-DataFile -FilePath $script:State.CurrentDataFile -Data ($script:UI.Tabs["All"].Content.ItemsSource) })
    $script:UI.BtnMenuSaveAs.Add_Click({ Save-DataFileAs })
    $script:UI.BtnMenuOpen.Add_Click({ Open-DataFile })
    $script:UI.BtnMenuImport.Add_Click({ Invoke-ImportDataFileDialog })
    $script:UI.BtnMenuEdit.Add_Click({ Edit-DataFile -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuFavorite.Add_Click({ Toggle-CommandFavorite })
    $script:UI.BtnMenuSettings.Add_Click({ Show-SettingsDialog })
    $script:UI.BtnMenuToggleSub.Add_Click({ Toggle-SubGrid })
    $script:UI.BtnMenuRunExternal.Add_Click({ 
        $script:State.RunCommandInternal = $false
        Invoke-MainRunClick -TabControl $script:UI.TabControl 
    })
    $script:UI.BtnMenuRunInternal.Add_Click({ 
        $script:State.RunCommandInternal = $true
        Invoke-MainRunClick -TabControl $script:UI.TabControl 
    })
    $script:UI.BtnMenuRunReopenLast.Add_Click({ if ($script:State.LastCommand) { Invoke-CommandDialog -Command $script:State.LastCommand } })
    $script:UI.BtnMenuRunRerunLast.Add_Click({ if ($script:State.LastCommand) { Run-Command -Command $script:State.LastCommand } })
    $script:UI.BtnMenuRunCopyToClipboard.Add_Click({ if ($script:State.LastCommand) { Copy-ToClipboard -String $script:State.LastCommand.Full } })

    # Main Buttons
    $script:UI.BtnMainRun.Add_Click({ Invoke-MainRunClick -TabControl $script:UI.TabControl })

    # Command dialog button events
    $script:UI.BtnCommandClose.Add_Click({ Hide-CommandDialog })
    $script:UI.BtnCommandRun.Add_Click({ Invoke-CommandRunClick -Command $script:State.CurrentCommand -Grid $script:UI.CommandGrid })
    $script:UI.BtnCommandCopyToClipboard.Add_Click({ Invoke-CommandCopyToClipboard -CurrentCommand $script:State.CurrentCommand -Grid $script:UI.CommandGrid })
    $script:UI.BtnCommandHelp.Add_Click({ Get-Help -Name $script:State.CurrentCommand.Root -ShowWindow })

    # Settings dialog button events
    $script:UI.BtnBrowseLogs.Add_Click({ Invoke-BrowseLogs })
    $script:UI.BtnBrowseDataFile.Add_Click({ Invoke-BrowseDataFile })
    $script:UI.BtnBrowseSettings.Add_Click({ Invoke-BrowseSettings })
    $script:UI.BtnBrowseFavorites.Add_Click({ Invoke-BrowseFavorites })
    $script:UI.BtnApplySettings.Add_Click({ Apply-Settings })
    $script:UI.BtnCloseSettings.Add_Click({ Hide-SettingsDialog })

    # Main Tab Control events
    $script:UI.TabControl.Add_SelectionChanged({
        param($sender, $e)
        Handle-TabSelection -SelectedTab $sender.SelectedItem
    })

    # Process Tab events
    $script:UI.BtnPSDetachTab.Add_Click({ Detach-CurrentTab })
    $script:UI.BtnPSAttachTab.Add_Click({ Show-AttachWindow })
    $script:UI.PSAddTab.Add_PreviewMouseLeftButtonDown({ New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs $script:Settings.DefaultShellArgs })
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

    $script:UI.Window.Add_Loaded({
        $script:UI.Window.Icon = $script:ApplicationPaths.IconFile
        Update-WindowTitle

        if ($script:Settings.OpenShellAtStart) {
            New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs $script:Settings.DefaultShellArgs
        }
    })

    $script:UI.Window.Add_Closing({ param($sender, $e) Invoke-WindowClosing -Sender $sender -E $e })
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

    foreach ($item in $selectedItems) {
        $id = $item.Id

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
        Set-UnsavedChanges $true
    }
}

# Handle the Main Edit Button click event to enable or disable editing of the grids
function Edit-DataFile {
    param (
        [hashtable]$tabs
    )

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

# Handle the Main Run Button click event to run the selected command/launch the CommandDialog
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

    if ($command.Root) {
        if ($selection.SkipParameterSelect) {
            $script:State.LastCommand = $command
            if ($command.PreCommand) {
                $command.Full = $command.PreCommand + "; "
            }
            $command.Full += $command.Root
            Run-Command $command $script:State.RunCommandInternal
        }
        else {
            Start-CommandDialog -Command $command
        }
    }
}

function Toggle-SubGrid {    
    if ($script:UI.Sub.Visibility -eq "Visible") {
        # Store current height before collapsing
        $script:State.SubGridExpandedHeight = $script:UI.Window.FindName("SubGridRow").Height.Value
        
        # Collapse the Sub grid
        $script:UI.Window.FindName("SubGridRow").Height = New-Object System.Windows.GridLength(0)
        $script:UI.Sub.Visibility = "Collapsed"
    } 
    else {
        # Restore previous height and visibility
        $script:UI.Window.FindName("SubGridRow").Height = New-Object System.Windows.GridLength($script:State.SubGridExpandedHeight)
        $script:UI.Sub.Visibility = "Visible"
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
                elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem -and -not $sender.IsInEditMode) {
                    $e.Handled = $true
                    Invoke-MainRunClick -TabControl $script:UI.TabControl
                }
            })
        }
        $newTab.Content.ItemsSource.Add($editedObject)
        Sort-TabControl -TabControl $tabControl
    }
}

# Process the CommandDialog dialog grid to show command parameter list
function Start-CommandDialog([Command]$command) {

    # If we are rerunning the command then the parameters are already saved
    if (-not $command.Parameters) {
        Clear-Grid $script:UI.CommandGrid

        try {
            # We only want to process the command if it is a PS script or function
            $type = Get-CommandType -Command $command.Root
            if (($type -ne "Function") -and ($type -ne "External Script")) {
                return
            }

            # Parse the command for parameters to build command grid with
            $command.Parameters = Get-ScriptBlockParameters -Command $command.Root
            Build-CommandGrid -Grid $script:UI.CommandGrid -Parameters $command.Parameters
        }
        catch {
            Show-ParameterLoadError -CommandName $command.Root -ErrorMessage $_.Exception.Message
            return
        }
    }

    # Assign the command as the current command so that BtnCommandRun can obtain it
    $script:State.CurrentCommand = $command

    $script:UI.BoxCommandName.Text = $command.Root
    Show-CommandDialog
}

# Display the hidden CommandDialog grid
function Show-CommandDialog {
    $script:UI.Overlay.Visibility = "Visible"
    $script:UI.CommandDialog.Visibility = "Visible"
}

# Hide the CommandDialog grid and clear for reuse
function Hide-CommandDialog() {
    $script:UI.CommandDialog.Visibility = "Hidden"
    $script:UI.Overlay.Visibility = "Collapsed"
}

# Construct the CommandDialog grid to show the correct content for each parameter
function Build-CommandGrid([System.Windows.Controls.Grid]$grid, [System.Object[]]$parameters) {
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
        $label.ToolTip = New-ToolTip -Content ""

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

function Compile-Command([Command]$command, [System.Windows.Controls.Grid]$grid) {
    # Clear if it existed from rerunning previous command
    if ($command.PreCommand) {
        $command.Full = $command.PreCommand + "; "
    }
    else {
        $command.Full = ""
    }
    $args = @{}
    $command.Full += "$($command.Root)"

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
        }
        elseif (-not [String]::IsNullOrWhiteSpace($args[$paramName])) {
            $command.Full += " -$paramName `"$($args[$paramName])`""
        }
    }
}

# Handle the Command Run Button click event to compile the inputted values for each parameter into a command string to be executed
function Invoke-CommandRunClick {
    param (
        [Command]$command,
        [System.Windows.Controls.Grid]$grid
    )

    Compile-Command -Command $command -Grid $grid
    $script:State.LastCommand = $command
    Run-Command $command $script:Settings.DefaultRunCommandInternal
    Hide-CommandDialog
}

function Invoke-CommandCopyToClipboard {
    param (
        [command]$currentCommand,
        [System.Windows.Controls.Grid]$grid
    )

    if ($currentCommand) {
        Compile-Command -Command $currentCommand -Grid $grid 
        Copy-ToClipboard -String $currentCommand.Full
    }
}

# Execute a command string
function Run-Command {
    param (
        [Command]$command
    )    

    Write-Log "Running: $($command.Root)"

    # We must escape any quotation marks passed or it will cause problems being passed through Start-Process
    $escapedCommand = $command.Full -replace '"', '\"'

    if ($script:State.RunCommandInternal) {
        New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs "-ExecutionPolicy Bypass -NoExit `" & { $escapedCommand } `"" -TabName $command.Root
    }
    else {
        Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit `" & { $escapedCommand } `""
    }
    $script:State.RunCommandInternal = $script:Settings.DefaultRunCommandInternal
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

# Create new tabitem element
function New-Tab {
    param (
        [string]$name
    )

    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header = $name
    return $tabItem
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
    $iconStyle = $script:UI.Window.FindResource("ContextMenuIconStyle")

    if ($name -eq "*") {
        # Favorites tab - simplified menu items (drag-and-drop handles reordering)
        $runAttachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runAttachedMenuItem.Header = "Run (Attached)"
        $runAttachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runAttachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Play
        $runAttachedIcon.Style = $iconStyle
        $runAttachedMenuItem.Icon = $runAttachedIcon
        $runAttachedMenuItem.Add_Click({
            $script:State.RunCommandInternal = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runAttachedMenuItem)

        $runDetachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runDetachedMenuItem.Header = "Run (Detached)"
        $runDetachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runDetachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::OpenInNew
        $runDetachedIcon.Style = $iconStyle
        $runDetachedMenuItem.Icon = $runDetachedIcon
        $runDetachedMenuItem.Add_Click({
            $script:State.RunCommandInternal = $false
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runDetachedMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $favoriteMenuItem = New-Object System.Windows.Controls.MenuItem
        $favoriteMenuItem.Header = "Remove from Favorites"
        $favIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $favIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::StarOff
        $favIcon.Style = $iconStyle
        $favoriteMenuItem.Icon = $favIcon
        $favoriteMenuItem.Add_Click({ Toggle-CommandFavorite })
        [void]$contextMenu.Items.Add($favoriteMenuItem)
    } else {
        # Regular tabs - standard menu items
        $runAttachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runAttachedMenuItem.Header = "Run (Attached)"
        $runAttachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runAttachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Play
        $runAttachedIcon.Style = $iconStyle
        $runAttachedMenuItem.Icon = $runAttachedIcon
        $runAttachedMenuItem.Add_Click({
            $script:State.RunCommandInternal = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runAttachedMenuItem)

        $runDetachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runDetachedMenuItem.Header = "Run (Detached)"
        $runDetachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runDetachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::OpenInNew
        $runDetachedIcon.Style = $iconStyle
        $runDetachedMenuItem.Icon = $runDetachedIcon
        $runDetachedMenuItem.Add_Click({
            $script:State.RunCommandInternal = $false
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runDetachedMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $favoriteMenuItem = New-Object System.Windows.Controls.MenuItem
        $favoriteMenuItem.Header = "Add to Favorites"
        $favIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $favIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Star
        $favIcon.Style = $iconStyle
        $favoriteMenuItem.Icon = $favIcon
        $favoriteMenuItem.Add_Click({ Toggle-CommandFavorite })

        # Store reference to favorite menu item so we can update it
        $contextMenu.Tag = @{
            FavoriteMenuItem = $favoriteMenuItem
            IconStyle = $iconStyle
        }

        # Add event handler to update the favorite menu item text/icon when context menu opens
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
            }
        })

        [void]$contextMenu.Items.Add($favoriteMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $addMenuItem = New-Object System.Windows.Controls.MenuItem
        $addMenuItem.Header = "Add Command"
        $addIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $addIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::AddBox
        $addIcon.Style = $iconStyle
        $addMenuItem.Icon = $addIcon
        $addMenuItem.Add_Click({ Add-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($addMenuItem)

        $removeMenuItem = New-Object System.Windows.Controls.MenuItem
        $removeMenuItem.Header = "Remove Command"
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

    $column = New-Object System.Windows.Controls.DataGridTextColumn
    $column.Header = $propertyName
    $column.Binding = New-Object System.Windows.Data.Binding $propertyName
    
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

# Show a GUI error popup so important or application breaking errors can be seen
function Show-ErrorMessageBox {
    param (
        [string]$message
    )

    Write-Error $message
    [System.Windows.MessageBox]::Show($message, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
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

    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.StatusBox.Text = $output}, "Normal")
    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.LogBox.AppendText("$output`n")}, "Normal")
    #Start-Sleep -Seconds $script:StatusTimeout; 
    #$script:UI.Window.Dispatcher.Invoke([action]{$script.UI.StatusBox.Text = ""}, "Normal")
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

# Create a new embedded process under a Tab Control
function New-ProcessTab {
    param (
        $tabControl,
        $process,
        $processArgs,
        $tabName = "PS_$($tabControl.Items.Count)"
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

    # Handle tab closure to terminate the PowerShell process
    $tab.Add_PreviewMouseRightButtonDown({
        param($sender, $eventArgs)
        if ($eventArgs.ChangedButton -eq 'Right') {
            $script:UI.PSTabControl.Items.Remove($sender)
            $sender.Tag["Process"].Kill()
        }
    })
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

# Popup window to select external PowerShell windows to attach
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
            Attach-ExternalWindow -Process $proc
            $attachWindow.Close()
        }
    })

    $attachWindow.ShowDialog()
}

# Attach and reparent an external window as an embedded tab
function Attach-ExternalWindow {
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

    # Handle tab closure
    $tab.Add_PreviewMouseRightButtonDown({
        param($sender, $eventArgs)
        if ($eventArgs.ChangedButton -eq 'Right') {
            $script:UI.PSTabControl.Items.Remove($sender)
            Detach-CurrentTab
        }
    })
}

# Set app settings from loaded settings
function Initialize-Settings {
    Load-Settings
    # Update UI elements with loaded settings
    $script:UI.TxtDefaultShell.Text = $script:Settings.DefaultShell
    $script:UI.TxtDefaultShellArgs.Text = $script:Settings.DefaultShellArgs
    $script:UI.ChkRunCommandInternal.IsChecked = $script:Settings.DefaultRunCommandInternal
    $script:UI.ChkOpenShellAtStart.IsChecked = $script:Settings.OpenShellAtStart
    $script:UI.TxtDefaultLogsPath.Text = $script:Settings.DefaultLogsPath
    $script:UI.TxtDefaultDataFile.Text = $script:Settings.DefaultDataFile
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
        RunCommandInternal = $script:Settings.DefaultRunCommandInternal
        OpenShellAtStart = $script:Settings.OpenShellAtStart
        DefaultLogsPath = $script:Settings.DefaultLogsPath
        SettingsPath = $script:Settings.SettingsPath
        FavoritesPath = $script:Settings.FavoritesPath
        ShowDebugTab = $script:Settings.ShowDebugTab
        DefaultDataFile = $script:Settings.DefaultDataFile
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
    $script:UI.ChkRunCommandInternal.IsChecked = $script:Settings.DefaultRunCommandInternal
    $script:UI.ChkOpenShellAtStart.IsChecked = $script:Settings.OpenShellAtStart
    $script:UI.TxtDefaultLogsPath.Text = $script:Settings.DefaultLogsPath
    $script:UI.TxtDefaultDataFile.Text = $script:Settings.DefaultDataFile
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
    $script:Settings.DefaultRunCommandInternal = $script:UI.ChkRunCommandInternal.IsChecked
    $script:Settings.OpenShellAtStart = $script:UI.ChkOpenShellAtStart.IsChecked
    $script:Settings.DefaultLogsPath = $script:UI.TxtDefaultLogsPath.Text
    $script:Settings.DefaultDataFile = $script:UI.TxtDefaultDataFile.Text
    $script:Settings.SettingsPath = $script:UI.TxtSettingsPath.Text
    $script:Settings.FavoritesPath = $script:UI.TxtFavoritesPath.Text
    $script:Settings.ShowDebugTab = $script:UI.ChkShowDebugTab.IsChecked

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
    $script:Settings.DefaultRunCommandInternal = $settings.RunCommandInternal
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
            RunCommandInternal = $script:Settings.DefaultRunCommandInternal
            OpenShellAtStart = $script:Settings.OpenShellAtStart
            DefaultLogsPath = $script:Settings.DefaultLogsPath
            SettingsPath = $script:Settings.SettingsPath
            FavoritesPath = $script:Settings.FavoritesPath
            ShowDebugTab = $script:Settings.ShowDebugTab
            DefaultDataFile = $script:Settings.DefaultDataFile
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

        $row = Get-DataGridRowFromPoint -Grid $sender -Point ($e.GetPosition($sender))
        if ($row -and $row.Item) {
            $script:State.DragDrop.DraggedItem = $row.Item
        }
    })

    # Handle mouse move to initiate drag operation
    $grid.Add_MouseMove({
        param($sender, $e)

        if ($e.LeftButton -eq [System.Windows.Input.MouseButtonState]::Pressed -and
            $script:State.DragDrop.DraggedItem -ne $null) {

            $dragData = New-Object System.Windows.DataObject([System.Windows.DataFormats]::Serializable, $script:State.DragDrop.DraggedItem)
            [System.Windows.DragDrop]::DoDragDrop($sender, $dragData, [System.Windows.DragDropEffects]::Move)
        }
    })

    # Handle drag over to show drop feedback
    $grid.Add_DragOver({
        param($sender, $e)

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

# Define the RowData object. This is the object that is used on all the Main window tabitem grids
class RowData {
    [int]$Id
    [string]$Name
    [string]$Description
    [string]$Category
    [string]$Command
    [bool]$SkipParameterSelect
    [string]$PreCommand
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
        $this.Order = $order
    }
}

# Define the Command object. This is used by the CommandDialog to construct the grid and run the command
class Command {
    [string]$Root
    [string]$Full
    [string]$PreCommand
    [System.Object[]]$Parameters
}

# Launch app
Initialize-Application
Start-MainWindow