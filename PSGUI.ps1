$App = @{}

# Determine app pathing whether running as PS script or EXE
if ($PSScriptRoot) {
    $App.Path = $PSScriptRoot
}
else {
    $App.Path = Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0]))
}

# Source files
$App.MainWindowXamlFile = Join-Path $App.Path "MainWindow.xaml"
$App.CommandWindowXamlFile = Join-Path $App.Path "CommandWindow.xaml"
$App.DebugWindowXamlFile = Join-Path $App.Path "DebugWindow.xaml"
$App.MaterialDesignThemes = Join-Path $App.Path "Assembly\MaterialDesignThemes.Wpf.dll"
$App.MaterialDesignColors = Join-Path $App.Path "Assembly\MaterialDesignColors.dll"
$App.DefaultConfigFile = Join-Path $App.Path "data.json"

# Initializing variables
$App.UI = $null
$App.Command = @{}
$App.LastAddedId = 0
$App.LastCommandId = 0
$App.MainTabsReadOnly = $true

# Get display resolution for initial window scaling
$App.MaxDisplayResolution = Get-CimInstance CIM_VideoController | Select SystemName, CurrentHorizontalResolution, CurrentVerticalResolution

# Load external resources
Add-Type -AssemblyName PresentationFramework
[Void][System.Reflection.Assembly]::LoadFrom($App.MaterialDesignThemes)
[Void][System.Reflection.Assembly]::LoadFrom($App.MaterialDesignColors)

class RowData {
    [int] $Id
    [string] $Name
    [string] $Description
    [string] $Category
    [string] $Command
    [bool] $SkipParameterSelect
    [string] $PreCommand
}

function MainWindow {
    # We create a new window and load all the window elements to variables of 
    # the same name and assign the window and all its elements under $App.UI 
    # e.g. $App.UI.Window, $App.UI.TabControl
    try {
        $App.UI = NewWindow -File $App.MainWindowXamlFile -ErrorAction Stop
    }
    catch {
        Show-ErrorMessageBox "Failed to create window from $($App.MainWindowXamlFile): $_"
        exit(1)
    }

    $App.UI.Window.add_Loaded({
        $App.UI.Window.Icon = Join-Path $App.Path "icon.ico"
    })

    $App.CurrentConfigFile = $App.DefaultConfigFile
    $json = LoadConfig $App.CurrentConfigFile

    # The "All" tab is the primary tab and so it must be created first
    $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json)
    $allTab = NewDataTab -Name "All" -ItemsSource $itemsSource -TabControl $App.UI.TabControl
    $App.UI.Tabs = @{}
    $App.UI.Tabs.Add("All", $allTab)

    # Determine current highest id number
    foreach ($value in ($json | Select-Object -ExpandProperty id -Unique)) {
        if ($value -gt $App.LastAddedId) {
            $App.LastAddedId = $value
        }
    }

    # Generate tabs and grids for each category
    foreach ($category in ($json | Select-Object -ExpandProperty Category -Unique)) {
        $itemsSource = [System.Collections.ObjectModel.ObservableCollection[RowData]]($json | Where-Object { $_.Category -eq $category })
        $tab = NewDataTab -Name $category -ItemsSource $itemsSource -TabControl $App.UI.TabControl
        $App.UI.Tabs.Add($category, $tab)
    }
    
    # We need to assign the cell edit handler to each tab's grid so that it works for all tabs
    foreach ($tab in $App.UI.Tabs.Values) {
        $tab.Content.Add_CellEditEnding({ 
            param($sender,$e) 
            CellEditingHandler -Sender $sender -E $e -TabControl $App.UI.TabControl -Tabs $App.UI.Tabs
        })
    }
    
    # Register button events
    $App.UI.BtnAdd.Add_Click({ BtnAddClick -TabControl $App.UI.TabControl -Tabs $App.UI.Tabs })
    $App.UI.BtnRemove.Add_Click({ BtnRemoveClick -TabControl $App.UI.TabControl -Tabs $App.UI.Tabs })
    $App.UI.BtnSave.Add_Click({ BtnSaveClick -File $App.CurrentConfigFile -Data ($App.UI.Tabs["All"].Content.ItemsSource) -SnackBar $App.UI.Snackbar })
    $App.UI.BtnEdit.Add_Click({ BtnEditClick -Tabs $App.UI.Tabs })
    $App.UI.BtnRun.Add_Click({ BtnMainRunClick -TabControl $App.UI.TabControl })

    $App.UI.BtnCommandClose.Add_Click({ CloseCommandDialog })
    $App.UI.BtnCommandRun.Add_Click({
        BtnCommandRunClick -Command $App.Command.Root -CommandEx $App.Command.Full -Parameters $App.Command.Parameters -Grid $App.UI.CommandGrid
    })

    # Set content and display the window
    $App.UI.Window.DataContext = $App.UI.Tabs
    $App.UI.Window.Dispatcher.InvokeAsync{ $App.UI.Window.ShowDialog() }.Wait() | Out-Null
}

function BtnAddClick($TabControl, $Tabs) {
    $newRow = [RowData]::New()
    $newRow.Id = ++$App.LastAddedId
    $tab = $Tabs["All"]
    $grid = $tab.Content
    $grid.ItemsSource.Add($newRow)
    $TabControl.SelectedItem = $tab
    SetTabsReadOnly -Tabs $Tabs
    $grid.SelectedItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.BeginEdit()
}

function BtnRemoveClick($TabControl, $Tabs) {
    $allGrid = $Tabs["All"].Content
    $allData = $allGrid.ItemsSource

    $grid = $TabControl.SelectedItem.Content

    # We want to make a copy of the selected items to avoid issues 
    # with the collection being modified while still enumerating
    $selectedItems = @()
    foreach ($item in $grid.SelectedItems) {
        $selectedItems += $item
    }

    foreach ($item in $selectedItems) {
        $category = $item.Category
        $id = $item.Id

        $categoryGrid = $Tabs[$category].Content
        $categoryData = $categoryGrid.ItemsSource
        $categoryIndex = GetGridIndexOfId -Grid $categoryGrid -Id $id
        $allIndex = GetGridIndexOfId -Grid $allGrid -Id $Id

        $allData.RemoveAt($allIndex)
        $categoryData.RemoveAt($categoryIndex)

        if ($categoryData.Count -eq 0) {
            $TabControl.Items.Remove($Tabs[$category])
            $Tabs.Remove($category)
        }
    }
}

function BtnSaveClick($File, $Data, $Snackbar) {
    try {
        SaveConfig $File $Data
        NewSnackBar -Snackbar $Snackbar -Text "Configuration saved"
    }
    catch {
        NewSnackBar -Snackbar $Snackbar -Text "Configuration save failed"
    }
}

function BtnEditClick($Tabs) {
    SetTabsReadOnly -Tabs $Tabs
}

function BtnMainRunClick($TabControl) {
    $grid = $TabControl.SelectedItem.Content
    $selection = $grid.SelectedItems
    $App.Command.Full = ""
    $App.Command.Root = $selection.Command

    if ($App.Command.Root) {
        if ($selection.PreCommand) {
            $App.Command.Full = $selection.PreCommand + "; "
        }

        if ($selection.SkipParameterSelect) {
            RunCommand $App.Command.Full
        }
        else {
            CommandDialog
        }
    }
}

function CellEditingHandler($sender, $e, $TabControl, $Tabs) {
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

    $allGrid = $Tabs["All"].Content
    $allData = $allGrid.ItemsSource
    $allIndex = GetGridIndexOfId -Grid $allGrid -Id $id

    # Sync values between All and Category tabs
    if (-not $newObject) {
        $categoryGrid = $Tabs[$category].Content
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
                $TabControl.Items.Remove($Tabs[$category])
                $Tabs.Remove($category)
            }
        }

        # Add the object to the new category tab
        $newTab = $Tabs[$newCategory]
        if (-not $newTab) {
            $itemsSource = New-Object System.Collections.ObjectModel.ObservableCollection[RowData]
            $newTab = NewDataTab -Name $newCategory -ItemsSource $itemsSource -TabControl $TabControl
            $Tabs.Add($newCategory, $newTab)
            $Tab.Content.Add_CellEditEnding({ 
                param($sender,$e) 
                CellEditingHandler -Sender $sender -E $e -TabControl $TabControl -Tabs $tabs
            })
        }
        $newTab.Content.ItemsSource.Add($editedObject)
    }
}

function CommandDialog() {
    # We must make a copy of the empty grid object
    #$emptyGrid = CopyObject -InputObject $App.UI.CommandGrid

    $App.UI.BoxCommandName.Text = $App.Command.Root

    # We only want to process the command if it is a PS script or function
    $type = GetCommandType -Command $App.Command.Root
    if (($type -ne "Function") -and ($type -ne "External Script")) {
        return
    }

    # Parse the command for parameters and build the grid with them
    $App.Command.Parameters = GetScriptBlockParameters -Command $App.Command.Root
    BuildCommandGrid -Grid $App.UI.CommandGrid -Parameters $App.Command.Parameters

    OpenCommandDialog
}

function BtnCommandRunClick($Command, $CommandEx, $Parameters, $Grid) {
    $args = @{}
    $CommandEx += "$($Command)"

    foreach ($param in $Parameters) {
        $isSwitch = $false
        $paramName = $param.Name.VariablePath
        $selection = $Grid.Children | Where-Object { $_.Name -eq $paramName }
        
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
            $CommandEx += " -$paramName"
        }
        elseif (-not [String]::IsNullOrWhiteSpace($args[$paramName])) {
            $CommandEx += " -$paramName `"$($args[$paramName])`""
        }
    }

    WriteDebug $CommandEx
    RunCommand $CommandEx
    CloseCommandDialog
}

function OpenCommandDialog {
    $App.UI.Main.Opacity = "0.5"
    $App.UI.CommandDialog.Visibility = "Visible"
}

function CloseCommandDialog() {
    $App.UI.Main.Opacity = "100"
    $App.UI.CommandDialog.Visibility = "Hidden"
    $App.UI.CommandGrid.Children.Clear()
    $App.UI.CommandGrid.RowDefinitions.Clear()
}

function GetCommandType($Command) {
    $type = (Get-Command $Command).CommandType 
    return $type
}

function GetScriptBlockParameters($Command) {
    $scriptBlock = (Get-Command $Command).ScriptBlock
    $parsed = [System.Management.Automation.Language.Parser]::ParseInput($ScriptBlock.ToString(), [ref]$null, [ref]$null)
    $parameters = $parsed.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)

    return $parameters
}

function ContainsAttributeType($Parameter, $TypeName) {
    foreach ($attribute in $Parameter.Attributes) {
        if ($attribute.TypeName.FullName -eq $TypeName) {
            return $true
        }
    }

    return $false
}

function GetValidateSetValues($Parameter) {
    # Start with an empty string value so that we can "deselect" values when
    # displayed in the drop-down box
    $validValues = [System.Collections.ArrayList]@("")

    # Check if the parameter has ValidateSet attribute
    foreach ($attribute in $Parameter.Attributes) {
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

function BuildCommandGrid($Grid, $Parameters) {
    for ($i = 0; $i -lt $Parameters.Count; $i++) {
        # Because there isn't a static number of rows and we need to iterate over the row index
        # we need to manually add a row for each parameter
        $rowDefinition = [System.Windows.Controls.RowDefinition]::New()
        [void]$Grid.RowDefinitions.Add($rowDefinition)

        $param = $Parameters[$i]
        $paramName = $param.Name.VariablePath
        $paramDefault = $param.DefaultValue.Value
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

function RunCommand($CommandEx) {      
    # We must escape any quotation marks passed or it will cause problems being passed through Start-Process
    $CommandEx = $CommandEx -replace '"', '\"'
    Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit `" & { $CommandEx } `""
}

function NewWindow($File) {
    try {
        [xml]$xaml = (Get-Content $File)
        $window = [System.Collections.Hashtable]::New()
        $nodeReader = [System.Xml.XmlNodeReader]::New($xaml)
        $xamlReader = [Windows.Markup.XamlReader]::Load($nodeReader)
        [void]$window.Add('Window', $xamlReader)

        $elements = $xaml.SelectNodes("//*[@Name]")
        foreach ($element in $elements) {
            $VarName = $element.Name
            $VarValue = $window.Window.FindName($Element.Name)
            [void]$window.Add($VarName, $VarValue)
        }
        
        return $window
    } 
    catch {
        Show-ErrorMessageBox("Error building Xaml data or loading window data.`n$_")
        exit
    }
}

function NewSnackBar {
    param (
        $Snackbar,
        $Text,
        $ButtonCaption
    )
    try {
        # Create queue for snackbar (popup messages)
        $MessageQueue = [MaterialDesignThemes.Wpf.SnackbarMessageQueue]::new()
        $MessageQueue.DiscardDuplicates = $true

        if ($ButtonCaption) {       
            $MessageQueue.Enqueue($Text, $ButtonCaption, {$null}, $null, $false, $false, [TimeSpan]::FromHours( 9999 ))
        }
        else {
            $MessageQueue.Enqueue($Text, $null, $null, $null, $false, $false, $null)
        }
        $Snackbar.MessageQueue = $MessageQueue
    }
    catch {
        Write-Error "No MessageQueue was declared in the window.`n$_"
    }
}

function BtnCloseWindowClick($Window) {
    $Window.Close()
}

function BtnMinimizeWindowClick($Window) {
    $Window.WindowState = 'Minimized'
}

function BtnMaximizeWindowClick($Window) {
    if ($Window.WindowState -eq 'Normal') {
        $Window.WindowState = 'Maximized'
    } 
    else {
        $Window.WindowState = 'Normal'
    }
}

function NewTab($Name) {
    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header = $Name

    return $tabItem
}

function NewDataGrid($Name, $ItemsSource) {
    $grid = New-Object System.Windows.Controls.DataGrid
    $grid.Name = $Name
    #$grid.AlternatingRowBackground = $App.AlternatingRowBackgroundColor
    $grid.Margin = New-Object System.Windows.Thickness(5)
    $grid.ItemsSource = $ItemsSource
    $grid.AutoGenerateColumns = $true
    $grid.CanUserAddRows = $false
    $grid.IsReadOnly = $App.MainTabsReadOnly
    return $grid
}

function NewDataTab($Name, $ItemsSource, $TabControl) {
    $grid = NewDataGrid -Name $Name -ItemsSource $ItemsSource
    $tab = NewTab -Name $Name
    $tab.Content = $grid
    [void]$TabControl.Items.Add($tab)
    return $tab
}

function AddToGrid($Grid, $Element) {
    [void]$Grid.Children.Add($Element)
}

function GetGridIndexOfId($Grid, $Id) {
    $itemsSource = $Grid.ItemsSource

    $index = -1
    for ($i = 0; $i -lt $itemsSource.Count; $i++) {
        if ($itemsSource[$i].Id -eq $Id) {
            $index = $i
            break
        }
    }

    return $index
}

function SetGridPosition($Element, $Row, $Column, $ColumnSpan) {
    if ($Row) {
        [System.Windows.Controls.Grid]::SetRow($Element, $Row)
    }
    if ($Column) {
        [System.Windows.Controls.Grid]::SetColumn($Element, $Column)
    }
    if ($ColumnSpan) {
        [System.Windows.Controls.Grid]::SetColumnSpan($Element, $ColumnSpan)
    }   
}

function NewLabel($Content, $HAlign, $VAlign) {
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $Content
    $label.HorizontalAlignment = $HAlign
    $label.VerticalAlignment = $VAlign
    $label.Margin = New-Object System.Windows.Thickness(3)

    return $label
}

function NewToolTip($Content) {
    $tooltip = New-Object System.Windows.Controls.ToolTip
    $tooltip.Content = $Content

    return $tooltip
}

function NewComboBox($Name, $ItemsSource, $SelectedItem) {
    $comboBox = New-Object System.Windows.Controls.ComboBox
    $comboBox.Name = $Name
    $comboBox.Margin = New-Object System.Windows.Thickness(5)
    $comboBox.ItemsSource = $ItemsSource
    $comboBox.SelectedItem = $SelectedItem
    
    return $comboBox
}

function NewTextBox($Name, $Text) {
    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Name = $Name
    $textBox.Margin = New-Object System.Windows.Thickness(5)
    $textBox.Text = $Text

    return $textBox
}

function NewCheckBox($Name, $IsChecked) {
    $checkbox = [System.Windows.Controls.CheckBox]::New()
    $checkbox.Name = $Name
    $checkbox.IsChecked = $IsChecked

    return $checkbox
}

function NewButton($Content, $HAlign, $Width) {
    $button = New-Object System.Windows.Controls.Button
    $button.Content = $Content
    $button.Margin = New-Object System.Windows.Thickness(10)
    $button.HorizontalAlignment = $HAlign
    $button.Width = $Width
    $button.IsDefault = $true

    return $button
}

function SetTabsReadOnly($Tabs) {
    $App.UI.BtnEdit.IsChecked = $App.MainTabsReadOnly

    $App.MainTabsReadOnly = (-not $App.MainTabsReadOnly)

    foreach ($tab in $Tabs.GetEnumerator()) {
        $grid = $tab.Value.Content
        $grid.IsReadOnly = $App.MainTabsReadOnly
    }
}

function InitializeConfig($File) {
    if (-not (Test-Path $File)) {
        try {
            New-Item -Path $File -ItemType "File" | Out-Null
        }
        catch {
            Show-ErrorMessageBox("Failed to create configuration file at path: $File")
            exit(1)
        }
    }
}

function LoadConfig($File) {
    try {
        [string]$contentRaw = (Get-Content $File -Raw -ErrorAction Stop)
        if ($contentRaw) {
            [array]$contentJson = $contentRaw | ConvertFrom-Json
            return $contentJson
        }
        else {
            Write-Verbose "Config file $file is empty."
            return
        }
    }
    catch {
        Show-ErrorMessageBox("Failed to load configuration from: $File")
        return
    }
}

function SaveConfig($File, $Data) {
    try {
        $populatedRows = $Data | Where-Object { $_.Name -ne $null }
        $json = ConvertTo-Json $populatedRows
        Set-Content -Path $File -Value $json
    }
    catch {
        Show-ErrorMessageBox("Failed to save configuration to: $File")
        return
    }
}

function Show-ErrorMessageBox($Message) {
    Write-Error $Message
    [System.Windows.MessageBox]::Show($Message, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}

function CopyObject($InputObject) {
    $SerialObject = [System.Management.Automation.PSSerializer]::Serialize($InputObject)
    return [System.Management.Automation.PSSerializer]::Deserialize($SerialObject)
}

function WriteDebug($Output) {
    #if ($App.UI.DebugGrid.Visibility = "Visible") {
        $App.UI.Window.Dispatcher.Invoke([action]{$App.UI.DebugBox.AppendText("$Output`n")}, "Normal")
    #}
}

MainWindow