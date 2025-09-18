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
}

# Initialize variables and load resources for application 
function Initialize-Application() {
    # Determine app pathing whether running as PS script or EXE
    $script:Path = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0])) }

    $script:ApplicationPaths = @{
        MainWindowXamlFile = Join-Path $script:Path "MainWindow.xaml"
        MaterialDesignThemes = Join-Path $script:Path "Assembly\MaterialDesignThemes.Wpf.dll"
        MaterialDesignColors = Join-Path $script:Path "Assembly\MaterialDesignColors.dll"
        DefaultConfigFile = Join-Path $script:Path "data.json"
        SettingsFilePath = Join-Path $env:APPDATA "PSGUI\settings.json"
        IconFile = Join-Path $script:Path "icon.ico"
        Win32APIFile = Join-Path $script:Path "Win32API.cs"
    }

    $script:State = @{
        CurrentConfigFile = $null
        CurrentCommand = $null
        LastCommand = $null
        HighestId = 0
        FavoritesHighestOrder = 0
        TabsReadOnly = $true
        RunCommandInternal = $script:Settings.DefaultRunCommandInternal
        ExtraColumnsVisibility = "Collapsed"
        ExtraColumns = @("Id", "Command", "SkipParameterSelect", "PreCommand")
        UpDownButtons = @{
            Up = $null
            Down = $null
        }
        AddRemoveButtons = @{
            Add = $null
            Remove = $null
        }
        SubGridExpandedHeight = 300
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

    $script:State.CurrentConfigFile = $script:ApplicationPaths.DefaultConfigFile
    $json = Load-DataFile $script:State.CurrentConfigFile
    $script:State.HighestId = Get-HighestId -Json $json
    $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json)

    # Create tabs and grids
    $script:UI.Tabs = @{}
    $allTab = New-DataTab -Name "All" -ItemsSource $itemsSource -TabControl $script:UI.TabControl
    $allTab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
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
    $script:UI.Tabs.Add("Favorites", $favTab)
    if ($favItemsSource -eq $null) {
        $script:UI.TabControl.SelectedItem = $allTab
    }

    foreach ($category in ($json | Select-Object -ExpandProperty Category -Unique)) {
        $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json | Where-Object { $_.Category -eq $category })
        $tab = New-DataTab -Name $category -ItemsSource $itemsSource -TabControl $script:UI.TabControl
        $tab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs }) # We need to assign the cell edit handler to each tab's grid so that it works for all tabs
        $script:UI.Tabs.Add($category, $tab) 
    }
    Sort-TabControl -TabControl $script:UI.TabControl
    
    Register-EventHandlers

    # Set content and display the window
    $script:UI.Window.DataContext = $script:UI.Tabs
    $script:UI.Window.Dispatcher.InvokeAsync{ $script:UI.Window.ShowDialog() }.Wait() | Out-Null
}

# Register all GUI events
function Register-EventHandlers {
    # Main button events
    $script:UI.BtnMainAdd.Add_Click({ Invoke-MainAddClick -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMainRemove.Add_Click({ Invoke-MainRemoveClick -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMainMoveUp.Add_Click({ Move-FavoriteItem -Direction "Up" })
    $script:UI.BtnMainMoveDown.Add_Click({ Move-FavoriteItem -Direction "Down" })
    $script:UI.BtnMainSave.Add_Click({ Save-DataFile -FilePath $script:State.CurrentConfigFile -Data ($script:UI.Tabs["All"].Content.ItemsSource) })
    $script:UI.BtnMainEdit.Add_Click({ Invoke-MainEditClick -Tabs $script:UI.Tabs })
    $script:UI.BtnMainFavorite.Add_Click({  Invoke-MainFavoriteClick })
    $script:UI.BtnMainSettings.Add_Click({ Show-SettingsDialog })
    $script:UI.BtnMainRun.Add_Click({ Invoke-MainRunClick -TabControl $script:UI.TabControl })
    $script:UI.BtnMainRunMenu.Add_Click({ $script:UI.ContextMenuMainRunMenu.IsOpen = $true })

    $script:UI.BtnToggleSub.Add_Click({ Invoke-ToggleSubGrid })

    # Run Menu Item events
    $script:UI.MenuItemMainRunExternal.Add_Click({ 
        $script:State.RunCommandInternal = $false
        Invoke-MainRunClick -TabControl $script:UI.TabControl 
    })
    $script:UI.MenuItemMainRunInternal.Add_Click({ 
        $script:State.RunCommandInternal = $true
        Invoke-MainRunClick -TabControl $script:UI.TabControl 
    })
    $script:UI.MenuItemMainRunReopenLast.Add_Click({ if ($script:State.LastCommand) { Start-CommandDialog -Command $script:State.LastCommand } })
    $script:UI.MenuItemMainRunRerunLast.Add_Click({ if ($script:State.LastCommand) { Run-Command -Command $script:State.LastCommand } })
    $script:UI.MenuItemMainRunCopyToClipboard.Add_Click({ if ($script:State.LastCommand) { Copy-ToClipboard -String $script:State.LastCommand.Full } })

    # Command dialog button events
    $script:UI.BtnCommandClose.Add_Click({ Hide-CommandDialog })
    $script:UI.BtnCommandRun.Add_Click({ Invoke-CommandRunClick -Command $script:State.CurrentCommand -Grid $script:UI.CommandGrid })
    $script:UI.BtnCommandCopyToClipboard.Add_Click({ Invoke-CommandCopyToClipboard -CurrentCommand $script:State.CurrentCommand -Grid $script:UI.CommandGrid })
    $script:UI.BtnCommandHelp.Add_Click({ Get-Help -Name $script:State.CurrentCommand.Root -ShowWindow })

    # Settings dialog button events
    $script:UI.BtnBrowseLogs.Add_Click({ Invoke-BrowseLogs })
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
        $script:UI.Window.Title = $script:AppTitle

        if ($script:Settings.OpenShellAtStart) {
            New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs $script:Settings.DefaultShellArgs
        }
    })

    $script:UI.Window.Add_Closing({ param($sender, $e) Invoke-WindowClosing -Sender $sender -E $e })
}

function Invoke-WindowClosing {
    param($sender, $e)

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
function Invoke-MainAddClick {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $newRow = New-Object RowData
    $newRow.Id = ++$script:State.HighestId
    $tab = $tabs["All"]
    $grid = $tab.Content
    $grid.ItemsSource.Add($newRow)
    $tabControl.SelectedItem = $tab
    # We don't want to change the tabs read only status if they are already in edit mode
    if ($script:State.TabsReadOnly) {
        Set-TabsReadOnlyStatus -Tabs $tabs
        Set-TabsExtraColumnsVisibility -Tabs $tabs
    }
    $grid.SelectedItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.BeginEdit()
}

# Handle the Main Window Remove Button click event to remove one or multiple RowData objects from the collection
function Invoke-MainRemoveClick {
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
}

# Handle the Main Edit Button click event to enable or disable editing of the grids
function Invoke-MainEditClick {
    param (
        [hashtable]$tabs
    )

    Set-TabsReadOnlyStatus -Tabs $tabs
    Set-TabsExtraColumnsVisibility -Tabs $tabs
}

function Invoke-MainFavoriteClick {
    $selectedTab = $script:UI.TabControl.SelectedItem
    $grid = $selectedTab.Content
    $selectedItem = $grid.SelectedItem

    if ($selectedItem) {
        $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
        $existingFavorite = $favorites | Where-Object { $_.Id -eq $selectedItem.Id }

        if ($existingFavorite) {
            [void]$favorites.Remove($existingFavorite)
            Write-Status "Removed from favorites"
        }
        else {
            $script:State.FavoritesHighestOrder++
            $favoriteRow = [FavoriteRowData]::new($selectedItem, $script:State.FavoritesHighestOrder)
            [void]$favorites.Add($favoriteRow)
            Write-Status "Added to favorites"
        }
        Save-Favorites -Favorites $favorites
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

function Invoke-ToggleSubGrid {    
    if ($script:UI.Sub.Visibility -eq "Visible") {
        # Store current height before collapsing
        $script:State.SubGridExpandedHeight = $script:UI.Window.FindName("SubGridRow").Height.Value
        
        # Collapse the Sub grid
        $script:UI.Window.FindName("SubGridRow").Height = New-Object System.Windows.GridLength(0)
        $script:UI.Sub.Visibility = "Collapsed"
        
        # Change the icon to indicate collapsed state
        $script:UI.BtnToggleSub.Content = New-Object MaterialDesignThemes.Wpf.PackIcon -Property @{
            Kind = "ArrowCollapseUp"
            Foreground = "Black"
        }
    } else {
        # Restore previous height and visibility
        $script:UI.Window.FindName("SubGridRow").Height = New-Object System.Windows.GridLength($script:State.SubGridExpandedHeight)
        $script:UI.Sub.Visibility = "Visible"
        
        # Change the icon to indicate expanded state
        $script:UI.BtnToggleSub.Content = New-Object MaterialDesignThemes.Wpf.PackIcon -Property @{
            Kind = "ArrowCollapseDown"
            Foreground = "Black"
        }
    }
}

function Invoke-TabControlSelectionChanged {
    param($sender, $e)
        
    $selectedTab = $sender.SelectedItem
    if ($selectedTab.Header -eq "*") {
        # Show Up/Down buttons, hide Add/Remove buttons
        $script:UI.BtnMainMoveUp.Visibility = "Visible"
        $script:UI.BtnMainMoveDown.Visibility = "Visible"
        $script:UI.BtnMainAdd.Visibility = "Collapsed"
        $script:UI.BtnMainRemove.Visibility = "Collapsed"
    } else {
        # Hide Up/Down buttons, show Add/Remove buttons
        $script:UI.BtnMainMoveUp.Visibility = "Collapsed"
        $script:UI.BtnMainMoveDown.Visibility = "Collapsed"
        $script:UI.BtnMainAdd.Visibility = "Visible"
        $script:UI.BtnMainRemove.Visibility = "Visible"
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
        # We only want to process the command if it is a PS script or function
        $type = Get-CommandType -Command $command.Root
        if (($type -ne "Function") -and ($type -ne "External Script")) {
            return
        }

        # Parse the command for parameters to build command grid with
        $command.Parameters = Get-ScriptBlockParameters -Command $command.Root
        Build-CommandGrid -Grid $script:UI.CommandGrid -Parameters $command.Parameters
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
    
    return (Get-Command $command).CommandType 
}

# Parse the command's script block to extract parameter info
function Get-ScriptBlockParameters {
    param (
        [string]$command
    )

    $scriptBlock = (Get-Command $command).ScriptBlock
    $parsed = [System.Management.Automation.Language.Parser]::ParseInput($scriptBlock.ToString(), [ref]$null, [ref]$null)
    return $parsed.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
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

    $script:UI.BtnMainEdit.IsChecked = $script:State.TabsReadOnly
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

function Update-ButtonVisibility {
    param (
        [string]$tabHeader
    )
    
    if ($tabHeader -eq "*") {
        # Show Up/Down buttons, hide Add/Remove buttons
        $script:UI.BtnMainMoveUp.Visibility = "Visible"
        $script:UI.BtnMainMoveDown.Visibility = "Visible"
        $script:UI.BtnMainAdd.Visibility = "Collapsed"
        $script:UI.BtnMainRemove.Visibility = "Collapsed"
    } else {
        # Hide Up/Down buttons, show Add/Remove buttons
        $script:UI.BtnMainMoveUp.Visibility = "Collapsed"
        $script:UI.BtnMainMoveDown.Visibility = "Collapsed"
        $script:UI.BtnMainAdd.Visibility = "Visible"
        $script:UI.BtnMainRemove.Visibility = "Visible"
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

    Update-ButtonVisibility -TabHeader $selectedTab.Header
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
    return ($json | Measure-Object -Property Id -Maximum).Maximum
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
    $grid.AutoGenerateColumns = $false
    
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
            New-Item -Path $filePath -ItemType "File" | Out-Null
        }
        catch {
            Show-ErrorMessageBox("Failed to create configuration file at path: $filePath")
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
            [array]$contentJson = $contentRaw | ConvertFrom-Json
            
            # Convert JSON objects to RowData objects
            $rowDataCollection = [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
            foreach ($item in $contentJson) {
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
            return $rowDataCollection
        }
        else {
            Write-Verbose "Data file $filePath is empty."
            return [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
        }
    }
    catch {
        Write-Error "Failed to load data from: $filePath"
        Write-Log "Failed to load data: $_"
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
        
        $json = ConvertTo-Json $populatedRows
        Set-Content -Path $filePath -Value $json
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
function Save-Favorites {
    param (
        [System.Collections.ObjectModel.ObservableCollection[Object]]$favorites
    )
    
    try {
        $favoritesDir = Split-Path $script:Settings.FavoritesPath -Parent
        if (-not (Test-Path $favoritesDir)) {
            New-Item -ItemType Directory -Path $favoritesDir -Force | Out-Null
        }

        # Only save the Id and Order to keep the file minimal
        $favoriteData = $favorites | Select-Object Id, Order
        $favoriteData | ConvertTo-Json | Set-Content $script:Settings.FavoritesPath
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
            $favoriteData = Get-Content $script:Settings.FavoritesPath | ConvertFrom-Json
            $favorites = @()
            
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
            return $favorites
        }
    }
    catch {
        Write-Log "Failed to load favorites: $_"
        return @()
    }
    return @()
}

function Move-FavoriteItem {
    param (
        [string]$direction
    )
    
    $grid = $script:UI.Tabs["Favorites"].Content
    $selectedItem = $grid.SelectedItem
    if (-not $selectedItem) { return }
    
    $itemsSource = $grid.ItemsSource
    $currentIndex = $itemsSource.IndexOf($selectedItem)
    $count = $itemsSource.Count
    
    # Calculate new index based on direction
    $newIndex = if ($direction -eq "Up") {
        if ($currentIndex -eq 0) { $count - 1 } else { $currentIndex - 1 }
    } else {
        if ($currentIndex -eq ($count - 1)) { 0 } else { $currentIndex + 1 }
    }
    
    # Store current orders
    $currentOrder = $selectedItem.Order
    $swapWithItem = $itemsSource[$newIndex]
    $swapWithOrder = $swapWithItem.Order
    
    # Swap order numbers
    $selectedItem.Order = $swapWithOrder
    $swapWithItem.Order = $currentOrder
    
    # Remove from current position and insert at new position
    $itemsSource.RemoveAt($currentIndex)
    $itemsSource.Insert($newIndex, $selectedItem)
    
    # Keep the moved item selected
    $grid.SelectedItem = $selectedItem
    $grid.ScrollIntoView($selectedItem)
    
    Save-Favorites -Favorites $itemsSource
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