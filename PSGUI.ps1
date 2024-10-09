param(
    [string]$StartCommand
)

# App Version
$script:Version = "1.2.1"

# Default configurable settings
$script:DefaultShell = "powershell"
$script:DefaultShellArgs = "-ExecutionPolicy Bypass -NoExit -Command `" & { [System.Console]::Title = 'PS' } `""
$script:DefaultRunCommandInternal = $true
$script:OpenShellAtStart = $false
$script:StatusTimeout = 3
$script:SettingsFilePath = Join-Path $env:APPDATA "PSGUI\settings.json"


$script:ExtraColumnsVisibility = "Collapsed"
$script:ExtraColumns = @("Id", "Command", "SkipParameterSelect", "PreCommand")

# Constants
$script:GWL_STYLE = -16
$script:WS_BORDERLESS = 0x800000  # WS_POPUP without WS_BORDER, WS_CAPTION, etc.
$script:WS_OVERLAPPEDWINDOW = 0x00CF0000

# Initialize variables and load resources for application 
function Initialize() {
    # Determine app pathing whether running as PS script or EXE
    if ($PSScriptRoot) {
        $script:Path = $PSScriptRoot
    }
    else {
        $script:Path = Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0]))
    }

    $script:CurrentCommand = $null
    $script:LastCommand = $null
    $script:HighestId = 0
    $script:TabsReadOnly = $true
    $script:RunCommandInternal = $script:DefaultRunCommandInternal
    $script:MainWindowXamlFile = Join-Path $script:Path "MainWindow.xaml"
    $script:MaterialDesignThemes = Join-Path $script:Path "Assembly\MaterialDesignThemes.Wpf.dll"
    $script:MaterialDesignColors = Join-Path $script:Path "Assembly\MaterialDesignColors.dll"
    $script:DefaultConfigFile = Join-Path $script:Path "data.json"
    $script:IconFile = Join-Path $script:Path "icon.ico"
    Add-Type -AssemblyName PresentationCore, PresentationFramework
    Add-Type -AssemblyName WindowsFormsIntegration
    [Void][System.Reflection.Assembly]::LoadFrom($script:MaterialDesignThemes)
    [Void][System.Reflection.Assembly]::LoadFrom($script:MaterialDesignColors)

    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class Win32 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

    public static IntPtr FindWindowByProcessId(int processId) {
        IntPtr foundHandle = IntPtr.Zero;

        EnumWindows(delegate (IntPtr hWnd, IntPtr lParam) {
            int windowProcessId;
            GetWindowThreadProcessId(hWnd, out windowProcessId);
            if (windowProcessId == processId) {
                foundHandle = hWnd;
                return false;  // Stop enumerating
            }
            return true;  // Continue enumerating
        }, IntPtr.Zero);

        return foundHandle;
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetParent(IntPtr hWndChild, IntPtr hWndNewParent);
    
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetFocus(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}
"@
}

# Load and process main application window
function MainWindow {
    # We create a new window and load all the window elements to variables of 
    # the same name and assign the window and all its elements under $script:UI 
    # e.g. $script:UI.Window, $script:UI.TabControl
    try {
        $script:UI = NewWindow -File $script:MainWindowXamlFile -ErrorAction Stop
    }
    catch {
        Show-ErrorMessageBox "Failed to create window from $($script:MainWindowXamlFile): $_"
        exit(1)
    }

    InitializeSettings

    $script:CurrentConfigFile = $script:DefaultConfigFile
    $json = LoadConfig $script:CurrentConfigFile
    $script:HighestId = GetHighestId -Json $json
    $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json)

    # The "All" tab is the primary tab and so it must be created first
    $allTab = NewDataTab -Name "All" -ItemsSource $itemsSource -TabControl $script:UI.TabControl
    $script:UI.Tabs = @{}
    $script:UI.Tabs.Add("All", $allTab)

    # Generate tabs and grids for each category
    foreach ($category in ($json | Select-Object -ExpandProperty Category -Unique)) {
        $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json | Where-Object { $_.Category -eq $category })
        $tab = NewDataTab -Name $category -ItemsSource $itemsSource -TabControl $script:UI.TabControl
        $tab.Content.Add_CellEditEnding({ param($sender,$e) CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs }) # We need to assign the cell edit handler to each tab's grid so that it works for all tabs
        $script:UI.Tabs.Add($category, $tab) 
    }
    SortTabControl -TabControl $script:UI.TabControl
    
    # Register Main button events
    $script:UI.BtnMainAdd.Add_Click({ BtnMainAddClick -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMainRemove.Add_Click({ BtnMainRemoveClick -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMainSave.Add_Click({ BtnMainSaveClick -File $script:CurrentConfigFile -Data ($script:UI.Tabs["All"].Content.ItemsSource) })
    $script:UI.BtnMainEdit.Add_Click({ BtnMainEditClick -Tabs $script:UI.Tabs })
    $script:UI.BtnMainSettings.Add_Click({ OpenSettingsDialog })
    $script:UI.BtnMainRun.Add_Click({ BtnMainRunClick -TabControl $script:UI.TabControl })
    $script:UI.BtnMainRunMenu.Add_Click({ $script:UI.ContextMenuMainRunMenu.IsOpen = $true })
    $script:UI.MenuItemMainRunExternal.Add_Click({ 
        $script:RunCommandInternal = $false
        BtnMainRunClick -TabControl $script:UI.TabControl 
    })
    $script:UI.MenuItemMainRunInternal.Add_Click({ 
        $script:RunCommandInternal = $true
        BtnMainRunClick -TabControl $script:UI.TabControl 
    })
    $script:UI.MenuItemMainRunReopenLast.Add_Click({ if ($script:LastCommand) { CommandDialog -Command $script:LastCommand } })
    $script:UI.MenuItemMainRunRerunLast.Add_Click({ if ($script:LastCommand) { RunCommand -Command $script:LastCommand.Full } })
    $script:UI.MenuItemMainRunCopyToClipboard.Add_Click({ if ($script:LastCommand) { CopyToClipBoard -String $script:LastCommand.Full -SnackBar $script:UI.Snackbar } })

    # Register Command dialog button events
    $script:UI.BtnCommandClose.Add_Click({ CloseCommandDialog })
    $script:UI.BtnCommandRun.Add_Click({ BtnCommandRunClick -Command $script:CurrentCommand -Grid $script:UI.CommandGrid })
    $script:UI.BtnCommandCopyToClipboard.Add_Click({ BtnCommandCopyToClipboard -CurrentCommand $script:CurrentCommand -Grid $script:UI.CommandGrid -SnackBar $script:UI.Snackbar })
    $script:UI.BtnCommandHelp.Add_Click({ Get-Help -Name $script:CurrentCommand.Root -ShowWindow })

    # Register Settings dialog button events
    $script:UI.BtnApplySettings.Add_Click({ ApplySettings })
    $script:UI.BtnCloseSettings.Add_Click({ CloseSettingsDialog })

    # Register Process Tab events
    $script:UI.BtnDetach.Add_Click({ DetachCurrentTab })
    $script:UI.BtnAttach.Add_Click({ ShowAttachWindow })
    $script:UI.PSAddTab.Add_PreviewMouseLeftButtonDown({ NewProcessTab -TabControl $script:UI.PSTabControl -Process $script:DefaultShell -ProcessArgs $script:DefaultShellArgs })
    $script:UI.PSTabControl.Add_SelectionChanged({
        param($sender, $eventArgs)
        $selectedTab = $script:UI.PSTabControl.SelectedItem
        if (($selectedTab) -and ($selectedTab -ne $script:UI.PSAddTab)) {
            $psHandle = $selectedTab.Tag["Handle"]
            #[Win32]::SetFocus($psHandle)
        }
    })
    $script:UI.Window.Add_GotFocus({
        if ($script:UI.PSTabControl.SelectedItem -ne $script:UI.PSAddTab) {
            $psHandle = $script:UI.PSTabControl.SelectedItem.Tag["Handle"]
            #[Win32]::SetFocus($psHandle)
        }
    })

    # Set content and display the window
    $script:UI.Window.Add_Loaded({ 
        $script:UI.Window.Icon = $script:IconFile
        $script:UI.Window.Title = "PSGUI - v$($script:Version)"
        if ($StartCommand) {
            $command = New-Object Command
            $command.Full = ""
            $command.Root = $script:StartCommand
            CommandDialog -Command $command
        }
        if ($script:OpenShellAtStart) {
            NewProcessTab -TabControl $script:UI.PSTabControl -Process $script:DefaultShell -ProcessArgs $script:DefaultShellArgs
        }
    })

    $script:UI.Window.DataContext = $script:UI.Tabs
    $script:UI.Window.Dispatcher.InvokeAsync{ $script:UI.Window.ShowDialog() }.Wait() | Out-Null
}

# Handle the Main Window Add Button click event to add a new RowData object to the collection
function BtnMainAddClick([System.Windows.Controls.TabControl]$tabControl, [hashtable]$tabs) {
    $newRow = New-Object RowData
    $newRow.Id = ++$script:HighestId
    $tab = $tabs["All"]
    $grid = $tab.Content
    $grid.ItemsSource.Add($newRow)
    $tabControl.SelectedItem = $tab
    # We don't want to change the tabs read only status if they are already in edit mode
    if ($script:TabsReadOnly) {
        SetTabsReadOnlyStatus -Tabs $tabs
        SetTabsExtraColumnsVisibility -Tabs $tabs
    }
    $grid.SelectedItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.BeginEdit()
}

# Handle the Main Window Remove Button click event to remove one or multiple RowData objects from the collection
function BtnMainRemoveClick([System.Windows.Controls.TabControl]$tabControl, [hashtable]$tabs) {
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
            $categoryIndex = GetGridIndexOfId -Grid $categoryGrid -Id $id
            $categoryData.RemoveAt($categoryIndex)
            if ($categoryData.Count -eq 0) {
                $tabControl.Items.Remove($tabs[$category])
                $tabs.Remove($category)
            }
        }
        $allIndex = GetGridIndexOfId -Grid $allGrid -Id $Id
        $allData.RemoveAt($allIndex)        
    }
}

# Handle the Main Save Button click event to save the current RoWData collection to the data file
function BtnMainSaveClick([string]$filePath, [System.Collections.ObjectModel.ObservableCollection[RowData]]$data, [MaterialDesignThemes.Wpf.Snackbar]$snackbar) {
    try {
        SaveConfig -FilePath $filePath -Data $data
        WriteStatus "Configuration saved"
    }
    catch {
        WriteStatus "Configuration save failed"
    }
}

# Handle the Main Edit Button click event to enable or disable editing of the grids
function BtnMainEditClick([hashtable]$tabs) {
    SetTabsReadOnlyStatus -Tabs $tabs
    SetTabsExtraColumnsVisibility -Tabs $tabs
}

# Handle the Main Log Button click event to view or hide the log grid
function BtnMainLogClick() {
    switch ($script:UI.LogGrid.Visibility) {
        "Visible" { $script:UI.LogGrid.Visibility = "Collapsed" }
        "Collapsed" { $script:UI.LogGrid.Visibility = "Visible" }
    }
}

# Handle the Main Run Button click event to run the selected command/launch the CommandDialog
function BtnMainRunClick([System.Windows.Controls.TabControl]$tabControl) {
    $grid = $tabControl.SelectedItem.Content
    $selection = $grid.SelectedItems
    $command = New-Object Command
    $command.Full = ""
    $command.Root = $selection.Command
    $command.PreCommand = $selection.PreCommand

    if ($command.Root) {
        if ($selection.SkipParameterSelect) {
            $script:LastCommand = $command
            if ($command.PreCommand) {
                $command.Full = $command.PreCommand + "; "
            }
            $command.Full += $command.Root
            RunCommand $command.Full $script:RunCommandInternal
        }
        else {
            CommandDialog -Command $command
        }
    }
}

# Handle the Cell Edit ending event to make sure all tabs are updated properly for cell changes
function CellEditEndingHandler($sender, $e, [System.Windows.Controls.TabControl]$tabControl, [hashtable]$tabs) {
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
    $allIndex = GetGridIndexOfId -Grid $allGrid -Id $id

    # Sync values between All and Category tabs
    if (-not $newObject) {
        $categoryGrid = $tabs[$category].Content
        $categoryData = $categoryGrid.ItemsSource
        $categoryIndex = GetGridIndexOfId -Grid $categoryGrid -Id $id
        $categoryData[$categoryIndex] = $editedObject
    }
    $allData[$allIndex] = $editedObject

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
            $newTab = NewDataTab -Name $newCategory -ItemsSource $itemsSource -TabControl $tabControl
            $tabs.Add($newCategory, $newTab)

            # Assign the CellEditEnding event to the new tab. We must use $script level vars here for TabControl and Tabs because the way events are handled if we use the local version
            # they will no longer exist when the event actually triggers
            $newTab.Content.Add_CellEditEnding({ param($sender,$e) CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        }
        $newTab.Content.ItemsSource.Add($editedObject)
        SortTabControl -TabControl $tabControl
    }
}

# Process the CommandDialog dialog grid to show command parameter list
function CommandDialog([Command]$command) {

    # If we are rerunning the command then the parameters are already saved
    if (-not $command.Parameters) {
        ClearGrid $script:UI.CommandGrid
        # We only want to process the command if it is a PS script or function
        $type = GetCommandType -Command $command.Root
        if (($type -ne "Function") -and ($type -ne "External Script")) {
            return
        }

        # Parse the command for parameters to build command grid with
        $command.Parameters = GetScriptBlockParameters -Command $command.Root
        BuildCommandGrid -Grid $script:UI.CommandGrid -Parameters $command.Parameters
    }
    
    # Assign the command as the current command so that BtnCommandRun can obtain it
    $script:CurrentCommand = $command

    $script:UI.BoxCommandName.Text = $command.Root
    OpenCommandDialog
}

# Construct the CommandDialog grid to show the correct content for each parameter
function BuildCommandGrid([System.Windows.Controls.Grid]$grid, [System.Object[]]$parameters) {
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

        $label = NewLabel -Content $paramName -HAlign "Left" -VAlign "Center"
        AddToGrid -Grid $Grid -Element $label
        SetGridPosition -Element $label -Row $i -Column 0
        $label.ToolTip = NewToolTip -Content ""

        # Set asterisk next to values that are mandatory
        if ($isMandatory) {
            $asterisk = NewLabel -Content "*" -HAlign "Right" -VAlign "Center"
            $asterisk.Foreground = "Red"
            AddToGrid -Grid $Grid -Element $asterisk
            SetGridPosition -Element $asterisk -Row $i -Column 1
        }

        if (ContainsAttributeType -Parameter $param -TypeName "ValidateSet") {
            # Get valid values from validate set and create dropdown box of them
            $validValues = GetValidateSetValues -Parameter $param
            $paramSource = $validValues -split "','"
            $box = NewComboBox -Name $paramName -ItemsSource $paramSource -SelectedItem $paramDefault
        }
        elseif (ContainsAttributeType -Parameter $param -TypeName "switch") {
            # If switch is true by default then check the box
            if ($param.DefaultValue) {
                $box = NewCheckBox -Name $paramName -IsChecked $true
            }
            else {
                $box = NewCheckBox -Name $paramName -IsChecked $false
            }
        }
        else {
            # Fill text box with any default values
            $box = NewTextBox -Name $paramName -Text $paramDefault
        }
        AddToGrid -Grid $Grid -Element $box
        SetGridPosition -Element $box -Row $i -Column 2
    }
}

# Display the hidden CommandDialog grid
function OpenCommandDialog {
    $script:UI.Overlay.Visibility = "Visible"
    $script:UI.CommandDialog.Visibility = "Visible"
}

# Hide the CommandDialog grid and clear for reuse
function CloseCommandDialog() {
    $script:UI.CommandDialog.Visibility = "Hidden"
    $script:UI.Overlay.Visibility = "Collapsed"
}

function ClearGrid([System.Windows.Controls.Grid]$grid) {
    $grid.Children.Clear()
    $grid.RowDefinitions.Clear()
}

function CompileCommand([Command]$command, [System.Windows.Controls.Grid]$grid) {
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
        
        if (ContainsAttributeType -Parameter $param -TypeName "ValidateSet") {
            if ($selection.SelectedItem) {
                $args[$paramName] = $selection.SelectedItem.ToString()
            }
        }
        elseif (ContainsAttributeType -Parameter $param -TypeName "switch") {
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
function BtnCommandRunClick([Command]$command, [System.Windows.Controls.Grid]$grid) {
    CompileCommand -Command $command -Grid $grid
    $script:LastCommand = $command
    RunCommand $command.Full $script:DefaultRunCommandInternal
    CloseCommandDialog
}

function BtnCommandCopyToClipboard([command]$currentCommand, [System.Windows.Controls.Grid]$grid, [MaterialDesignThemes.Wpf.Snackbar]$snackbar) {
    if ($currentCommand) {
        CompileCommand -Command $currentCommand -Grid $grid 
        CopyToClipBoard -String $currentCommand.Full -SnackBar $snackbar
    }
}

# Execute a command string in an external PowerShell window
function RunCommand([string]$command) {      
    WriteLog "Running: $command"
    # We must escape any quotation marks passed or it will cause problems being passed through Start-Process
    $command = $command -replace '"', '\"'

    if ($script:RunCommandInternal) {
        NewProcessTab -TabControl $script:UI.PSTabControl -Process $script:DefaultShell -ProcessArgs "-ExecutionPolicy Bypass -NoExit `" & { $command } `""
    }
    else {
        Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit `" & { $command } `""
    }
    $script:RunCommandInternal = $script:DefaultRunCommandInternal
}

# Determine the PowerShell command type (Function,Script,Cmdlet)
function GetCommandType([string]$command) {
    $type = (Get-Command $command).CommandType 
    return $type
}

# Parse the command's script block to extract parameter info
function GetScriptBlockParameters([string]$command) {
    $scriptBlock = (Get-Command $command).ScriptBlock
    $parsed = [System.Management.Automation.Language.Parser]::ParseInput($scriptBlock.ToString(), [ref]$null, [ref]$null)
    $parameters = $parsed.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)

    return $parameters
}

# Determine if a parameter contains a certain attribute type (Switch,ValidateSet)
function ContainsAttributeType([System.Management.Automation.Language.ParameterAst]$parameter, [string]$typeName) {
    foreach ($attribute in $parameter.Attributes) {
        if ($attribute.TypeName.FullName -eq $typeName) {
            return $true
        }
    }
    return $false
}

# Retrieve the list of values from a parameter's ValidateSet 
function GetValidateSetValues([System.Management.Automation.Language.ParameterAst]$parameter) {
    # Start with an empty string value so that we can "deselect" values when
    # displayed in the drop-down box
    $validValues = [System.Collections.ArrayList]@("")

    # Check if the parameter has ValidateSet attribute
    foreach ($attribute in $parameter.Attributes) {
        if ($attribute.TypeName.FullName -eq 'ValidateSet') {
            $values = $attribute.PositionalArguments
            break
        }
    }
    foreach ($value in $values) {
        # We need to convert from AST object to string so we can remove extra quotes
        $valueStr = $($value.ToString()).Replace("'","").Replace("`"","")
        [void]$validValues.Add($valueStr)
    }

    return $validValues
}

# Add a WPF element to a grid
function AddToGrid([System.Windows.Controls.Grid]$grid, $element) {
    [void]$grid.Children.Add($element)
}

# Determine a grid row index of a specific command id on a particular datagrid
function GetGridIndexOfId([System.Windows.Controls.DataGrid]$grid, [int]$id) {
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
function SetGridPosition([System.Windows.Controls.Control]$element, [int]$row, [int]$column, [int]$columnSpan) {
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
function SetTabsReadOnlyStatus([hashtable]$tabs) {
    $script:UI.BtnMainEdit.IsChecked = $script:TabsReadOnly
    $script:TabsReadOnly = (-not $script:TabsReadOnly)
    foreach ($tab in $tabs.GetEnumerator()) {
        $grid = $tab.Value.Content
        $grid.IsReadOnly = $script:TabsReadOnly
    }
}

# Show or hide the 'extra columns' on all tabs' grids
function SetTabsExtraColumnsVisibility([hashtable]$tabs) {
    switch ($script:ExtraColumnsVisibility) {
        "Visible" { $script:ExtraColumnsVisibility = "Collapsed" }
        "Collapsed" { $script:ExtraColumnsVisibility = "Visible" }
    }
    foreach ($tab in $tabs.GetEnumerator()) {
        $grid = $tab.Value.Content
        SetGridExtraColumnsVisibility -Grid $grid
    }
}

# Show or hide the 'extra columns' on a single grid
function SetGridExtraColumnsVisibility([System.Windows.Controls.DataGrid]$grid) {
    foreach ($column in $grid.Columns) {
        foreach ($extraCol in $script:ExtraColumns) {
            if ($column.Header -eq $extraCol) {
                $column.Visibility = $script:ExtraColumnsVisibility
            }
        }
    }
}

# Sort the order of the tabs in tab control alphabetically by their header
function SortTabControl([System.Windows.Controls.TabControl]$tabControl) {
    $tabItems = $tabControl.Items
    $allTabItem = $tabItems | Where-Object { $_.Header -eq "All" }
    $sortedTabItems = $tabItems | Where-Object { $_.Header -ne "All" } | Sort-Object -Property { $_.Header.ToString() }
    $tabControl.Items.Clear()
    [void]$tabControl.Items.Add($allTabItem)
    foreach ($tabItem in $sortedTabItems) {
        [void]$tabControl.Items.Add($tabItem)
    }
}

# Sort a grid alphabetically by a specific column
function SortGridByColumn([System.Windows.Controls.DataGrid]$grid, [string]$columnName) {
    $grid.Items.SortDescriptions.Clear()
    $sort = New-Object System.ComponentModel.SortDescription($columnName, [System.ComponentModel.ListSortDirection]::Ascending)
    $grid.Items.SortDescriptions.Add($sort)
    $grid.Items.Refresh()
}

# Determine the current highest Id that exists in the collection
function GetHighestId([System.Object[]]$json) {
    $highest = 0
    foreach ($value in ($json | Select-Object -ExpandProperty id -Unique)) {
        if ($value -gt $greatest) {
            $highest = $value
        }
    }
    return $highest
}

# Copy a string to the system clipboard
function CopyToClipBoard([string]$string, [MaterialDesignThemes.Wpf.Snackbar]$snackbar) {
    WriteLog "Copied to clipboard: $string"
    NewSnackBar -Snackbar $snackbar -Text "Copied to clipboard"
    Set-ClipBoard -Value $string
}

# Create a new WPF window from an XML file and load all WPF elements and return them under one variable
function NewWindow([string]$filePath) {
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

# Create new popup snackbar message
function NewSnackBar([MaterialDesignThemes.Wpf.Snackbar]$snackbar, [string]$text, [string]$caption) {
    try {
        $queue = New-Object MaterialDesignThemes.Wpf.SnackbarMessageQueue
        $queue.DiscardDuplicates = $true
        if ($caption) {       
            $queue.Enqueue($text, $caption, {$null}, $null, $false, $false, [TimeSpan]::FromHours(9999))
        }
        else {
            $queue.Enqueue($text, $null, $null, $null, $false, $false, $null)
        }
        $snackbar.MessageQueue = $queue
    }
    catch {
        Write-Error "No MessageQueue was declared in the window.`n$_"
    }
}

# Create new tabitem element
function NewTab([string]$name) {
    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header = $name
    return $tabItem
}

# Create new datagrid element for the main window
function NewDataGrid([string]$name, [System.Collections.ObjectModel.ObservableCollection[RowData]]$itemsSource) {
    $grid = New-Object System.Windows.Controls.DataGrid
    $grid.Name = $name
    $grid.Margin = New-Object System.Windows.Thickness(5)
    $grid.ItemsSource = $itemsSource
    $grid.CanUserAddRows = $false
    $grid.IsReadOnly = $script:TabsReadOnly

    # Rather than autogenerate columns we want to manually create them based on the properties of RowData
    # as autogenerated columns cannot have their visibility set
    $grid.AutoGenerateColumns = $false
    $rowType = [RowData]
    $properties = $rowType.GetProperties()
    foreach ($prop in $properties) {
        $column = New-Object System.Windows.Controls.DataGridTextColumn
        $column.Header = $prop.Name
        $column.Binding = New-Object System.Windows.Data.Binding $prop.Name
        $grid.Columns.Add($column)
    }
    SetGridExtraColumnsVisibility -Grid $grid
    SortGridByColumn -Grid $grid -ColumnName "Name"
    return $grid
}

# Create a new tabitem that contains a datagrid and assign to the main tabcontrol
function NewDataTab([string]$name, [System.Collections.ObjectModel.ObservableCollection[RowData]]$itemsSource, [System.Windows.Controls.TabControl]$tabControl) {
    $grid = NewDataGrid -Name $name -ItemsSource $itemsSource
    $tab = NewTab -Name $name
    $tab.Content = $grid
    [void]$tabControl.Items.Add($tab)
    return $tab
}

# Create a new text label element
function NewLabel([string]$content, [string]$halign, [string]$valign) {
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $content
    $label.HorizontalAlignment = $halign
    $label.VerticalAlignment = $valign
    $label.Margin = New-Object System.Windows.Thickness(3)
    return $label
}

# Create a new tooltip element
function NewToolTip([string]$content) {
    $tooltip = New-Object System.Windows.Controls.ToolTip
    $tooltip.Content = $content
    return $tooltip
}

# Create a new combo box element
function NewComboBox([string]$name, [System.String[]]$itemsSource, [string]$selectedItem) {
    $comboBox = New-Object System.Windows.Controls.ComboBox
    $comboBox.Name = $name
    $comboBox.Margin = New-Object System.Windows.Thickness(5)
    $comboBox.ItemsSource = $itemsSource
    $comboBox.SelectedItem = $selectedItem
    return $comboBox
}

# Create a new text box element
function NewTextBox([string]$name, [string]$text) {
    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Name = $name
    $textBox.Margin = New-Object System.Windows.Thickness(5)
    $textBox.Text = $text
    return $textBox
}

# Create a new check box element
function NewCheckBox([string]$name, [bool]$isChecked) {
    $checkbox = New-Object System.Windows.Controls.CheckBox
    $checkbox.Name = $name
    $checkbox.IsChecked = $isChecked
    return $checkbox
}

# Create a new button element
function NewButton([string]$content, [string]$halign, [int]$width) {
    $button = New-Object System.Windows.Controls.Button
    $button.Content = $content
    $button.Margin = New-Object System.Windows.Thickness(10)
    $button.HorizontalAlignment = $halign
    $button.Width = $width
    $button.IsDefault = $true
    return $button
}

function ResetRunCommandDefaultLocation(){
    $script:RunCommandInternal = $script:DefaultRunCommandInternal
}

# Create a blank data file if it doesn't already exist
function InitializeConfig([string]$filePath) {
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
function LoadConfig([string]$filePath) {
    try {
        [string]$contentRaw = (Get-Content $filePath -Raw -ErrorAction Stop)
        if ($contentRaw) {
            [array]$contentJson = $contentRaw | ConvertFrom-Json
            return $contentJson
        }
        else {
            Write-Verbose "Config file $filePath is empty."
            return
        }
    }
    catch {
        Show-ErrorMessageBox("Failed to load configuration from: $filePath")
        return
    }
}

# Save the data collection to the data file
function SaveConfig([string]$filePath, [System.Collections.ObjectModel.ObservableCollection[RowData]]$data) {
    try {
        $populatedRows = $data | Where-Object { $_.Name -ne $null }
        $json = ConvertTo-Json $populatedRows
        Set-Content -Path $filePath -Value $json
    }
    catch {
        Show-ErrorMessageBox("Failed to save configuration to: $filePath")
        return
    }
}

# Show a GUI error popup so important or application breaking errors can be seen
function Show-ErrorMessageBox([string]$message) {
    Write-Error $message
    [System.Windows.MessageBox]::Show($message, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}

# Write text to the LogBox
function WriteLog([string]$output) {
    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.LogBox.AppendText("$output`n")}, "Normal")
}

function WriteStatus([string]$output) {
    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.StatusBox.Text = $output}, "Normal")
    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.LogBox.AppendText("$output`n")}, "Normal")
    #Start-Sleep -Seconds $script:StatusTimeout; 
    #$script:UI.Window.Dispatcher.Invoke([action]{$script.UI.StatusBox.Text = ""}, "Normal")
}

function NewProcessTab($tabControl, $process, $processArgs) {

    $proc = Start-Process $process -WindowStyle Hidden -PassThru -ArgumentList $processArgs
    
    Start-Sleep -Seconds 2

    # Find the window handle of the PowerShell process using process ID
    $psHandle = [Win32]::FindWindowByProcessId($proc.Id)
    if ($psHandle -eq [IntPtr]::Zero) {
        WriteLog "Failed to retrieve the PowerShell window handle for process ID: $($proc.Id)."
        return
    }

    $tab = NewTab -Name "PS_$($tabControl.Items.Count)"
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
    $currentStyle = [Win32]::GetWindowLong($psHandle, $GWL_STYLE)
    [Win32]::SetWindowLong($psHandle, $GWL_STYLE, $currentStyle -band -0x00C00000)  # Remove WS_CAPTION and WS_THICKFRAME

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
            WriteLog "Invalid window handle in SizeChanged event."
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
            DetachPowerShellWindow -tab $sender
            $e.Handled = $true
        }
    })
}

function DetachCurrentTab {
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

function ShowAttachWindow {
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
            AttachExternalWindow -Process $proc
            $attachWindow.Close()
        }
    })

    $attachWindow.ShowDialog()
}

function AttachExternalWindow {
    param (
        [System.Diagnostics.Process]$Process
    )

    $psHandle = $Process.MainWindowHandle
    
    # Remove window frame
    $style = [Win32]::GetWindowLong($psHandle, $script:GWL_STYLE)
    $style = $style -band -bnot $script:WS_OVERLAPPEDWINDOW
    [Win32]::SetWindowLong($psHandle, $script:GWL_STYLE, $style)

    $tab = NewTab -Name "PS_$($script:UI.PSTabControl.Items.Count)"
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
            WriteLog "Invalid window handle in SizeChanged event."
        }
    })

    # Handle tab closure
    $tab.Add_PreviewMouseRightButtonDown({
        param($sender, $eventArgs)
        if ($eventArgs.ChangedButton -eq 'Right') {
            $script:UI.PSTabControl.Items.Remove($sender)
            DetachCurrentTab
        }
    })
}

function InitializeSettings {
    LoadSettings
    # Update UI elements with loaded settings
    $script:UI.TxtDefaultShell.Text = $script:DefaultShell
    $script:UI.TxtDefaultShellArgs.Text = $script:DefaultShellArgs
    $script:UI.ChkRunCommandInternal.IsChecked = $script:DefaultRunCommandInternal
    $script:UI.ChkOpenShellAtStart.IsChecked = $script:OpenShellAtStart
}

function CreateDefaultSettings {
    $defaultSettings = @{
        DefaultShell = $script:DefaultShell
        DefaultShellArgs = $script:DefaultShellArgs
        RunCommandInternal = $script:DefaultRunCommandInternal
        OpenShellAtStart = $script:OpenShellAtStart
    }
    return $defaultSettings
}

function OpenSettingsDialog {
    $script:UI.Overlay.Visibility = "Visible"
    $script:UI.SettingsDialog.Visibility = "Visible"
    
    # Populate current settings
    $script:UI.TxtDefaultShell.Text = $script:DefaultShell
    $script:UI.TxtDefaultShellArgs.Text = $script:DefaultShellArgs
    $script:UI.ChkRunCommandInternal.IsChecked = $script:DefaultRunCommandInternal
    $script:UI.ChkOpenShellAtStart.IsChecked = $script:OpenShellAtStart
}

function CloseSettingsDialog {
    $script:UI.SettingsDialog.Visibility = "Hidden"
    $script:UI.Overlay.Visibility = "Collapsed"
}

function ApplySettings {
    $script:DefaultShell = $script:UI.TxtDefaultShell.Text
    $script:DefaultShellArgs = $script:UI.TxtDefaultShellArgs.Text
    $script:DefaultRunCommandInternal = $script:UI.ChkRunCommandInternal.IsChecked
    $script:OpenShellAtStart = $script:UI.ChkOpenShellAtStart.IsChecked

    SaveSettings
    CloseSettingsDialog
}

# Load settings from file
function LoadSettings {
    EnsureSettingsFileExists
    $settings = Get-Content $script:SettingsFilePath | ConvertFrom-Json

    # Apply loaded settings to script variables
    $script:DefaultShell = $settings.DefaultShell
    $script:DefaultShellArgs = $settings.DefaultShellArgs
    $script:DefaultRunCommandInternal = $settings.RunCommandInternal
    $script:OpenShellAtStart = $settings.OpenShellAtStart
}

# Save settings to file
function SaveSettings {
    try {
        $settings = @{
            DefaultShell = $script:DefaultShell
            DefaultShellArgs = $script:DefaultShellArgs
            RunCommandInternal = $script:DefaultRunCommandInternal
            OpenShellAtStart = $script:OpenShellAtStart
        }
        $settings | ConvertTo-Json | Set-Content $script:SettingsFilePath
        WriteStatus "Settings saved"
    }
    catch {
        WriteStatus "Failed to save settings"
    }
}

function EnsureSettingsFileExists {
    $settingsDir = Split-Path $script:SettingsFilePath -Parent
    if (-not (Test-Path $settingsDir)) {
        New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
    }
    if (-not (Test-Path $script:SettingsFilePath)) {
        $defaultSettings = CreateDefaultSettings
        $defaultSettings | ConvertTo-Json | Set-Content $script:SettingsFilePath
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

# Define the Command object. This is used by the CommandDialog to construct the grid and run the command
class Command {
    [string]$Root
    [string]$Full
    [string]$PreCommand
    [System.Object[]]$Parameters
}

# Run the application
Initialize
MainWindow