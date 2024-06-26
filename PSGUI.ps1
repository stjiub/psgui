$script:Version = "1.1.1"
$script:ExtraColumnsVisibility = "Collapsed"
$script:ExtraColumns = @("Id", "Command", "SkipParameterSelect", "PreCommand")

function Initialize() {
    # Determine app pathing whether running as PS script or EXE
    if ($PSScriptRoot) {
        $script:Path = $PSScriptRoot
    }
    else {
        $script:Path = Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0]))
    }

    $script:CurrentCommand = $null
    $script:HighestId = 0
    $script:TabsReadOnly = $true
    $script:MainWindowXamlFile = Join-Path $script:Path "MainWindow.xaml"
    $script:MaterialDesignThemes = Join-Path $script:Path "Assembly\MaterialDesignThemes.Wpf.dll"
    $script:MaterialDesignColors = Join-Path $script:Path "Assembly\MaterialDesignColors.dll"
    $script:DefaultConfigFile = Join-Path $script:Path "data.json"
    $script:IconFile = Join-Path $script:Path "icon.ico"
    Add-Type -AssemblyName PresentationFramework
    [Void][System.Reflection.Assembly]::LoadFrom($script:MaterialDesignThemes)
    [Void][System.Reflection.Assembly]::LoadFrom($script:MaterialDesignColors)
}

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
        $tab.Content.Add_CellEditEnding({ param($sender,$e) CellEditingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs }) # We need to assign the cell edit handler to each tab's grid so that it works for all tabs
        $script:UI.Tabs.Add($category, $tab) 
    }
    SortTabControl -TabControl $script:UI.TabControl
    
    # Register button events
    $script:UI.BtnAdd.Add_Click({ BtnAddClick -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnRemove.Add_Click({ BtnRemoveClick -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnSave.Add_Click({ BtnSaveClick -File $script:CurrentConfigFile -Data ($script:UI.Tabs["All"].Content.ItemsSource) -SnackBar $script:UI.Snackbar })
    $script:UI.BtnEdit.Add_Click({ BtnEditClick -Tabs $script:UI.Tabs })
    $script:UI.BtnLog.Add_Click({ BtnLogClick })
    $script:UI.BtnRun.Add_Click({ BtnMainRunClick -TabControl $script:UI.TabControl })
    $script:UI.BtnCommandClose.Add_Click({ CloseCommandDialog })
    $script:UI.BtnCommandRun.Add_Click({ BtnCommandRunClick -Command $script:CurrentCommand -Grid $script:UI.CommandGrid })
    $script:UI.BtnCommandHelp.Add_Click({ Get-Help -Name $script:CurrentCommand.Root -ShowWindow })

    # Set content and display the window
    $script:UI.Window.Add_Loaded({ 
        $script:UI.Window.Icon = $script:IconFile
        $script:UI.Window.Title = "PSGUI - v$($script:Version)"
    })
    $script:UI.Window.DataContext = $script:UI.Tabs
    $script:UI.Window.Dispatcher.InvokeAsync{ $script:UI.Window.ShowDialog() }.Wait() | Out-Null
}

function BtnAddClick([System.Windows.Controls.TabControl]$tabControl, [hashtable]$tabs) {
    $newRow = New-Object RowData
    $newRow.Id = ++$script:HighestId
    $tab = $tabs["All"]
    $grid = $tab.Content
    $grid.ItemsSource.Add($newRow)
    $tabControl.SelectedItem = $tab
    # We don't want to change the tabs read only status if they are already in edit mode
    if ($script:TabsReadOnly) {
        SetTabsReadOnlyStatus -Tabs $tabs
    }
    $grid.SelectedItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.BeginEdit()
}

function BtnRemoveClick([System.Windows.Controls.TabControl]$tabControl, [hashtable]$tabs) {
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
        $category = $item.Category
        $id = $item.Id

        $categoryGrid = $tabs[$category].Content
        $categoryData = $categoryGrid.ItemsSource
        $categoryIndex = GetGridIndexOfId -Grid $categoryGrid -Id $id
        $allIndex = GetGridIndexOfId -Grid $allGrid -Id $Id

        $allData.RemoveAt($allIndex)
        $categoryData.RemoveAt($categoryIndex)

        if ($categoryData.Count -eq 0) {
            $tabControl.Items.Remove($tabs[$category])
            $tabs.Remove($category)
        }
    }
}

function BtnSaveClick([string]$filePath, [System.Collections.ObjectModel.ObservableCollection[RowData]]$data, [MaterialDesignThemes.Wpf.Snackbar]$snackbar) {
    try {
        SaveConfig -FilePath $filePath -Data $data
        NewSnackBar -Snackbar $snackbar -Text "Configuration saved"
    }
    catch {
        NewSnackBar -Snackbar $snackbar -Text "Configuration save failed"
    }
}

function BtnEditClick([hashtable]$tabs) {
    SetTabsReadOnlyStatus -Tabs $tabs
    SetTabsExtraColumnsVisibility -Tabs $tabs
}

function BtnLogClick() {
    switch ($script:UI.LogGrid.Visibility) {
        "Visible" { $script:UI.LogGrid.Visibility = "Collapsed" }
        "Collapsed" { $script:UI.LogGrid.Visibility = "Visible" }
    }
}

function BtnMainRunClick([System.Windows.Controls.TabControl]$tabControl) {
    $grid = $tabControl.SelectedItem.Content
    $selection = $grid.SelectedItems
    $command = New-Object Command
    $command.Full = ""
    $command.Root = $selection.Command

    if ($command.Root) {
        if ($selection.PreCommand) {
            $command.Full = $selection.PreCommand + "; "
        }

        if ($selection.SkipParameterSelect) {
            RunCommand $command.Full
        }
        else {
            CommandDialog -Command $command
        }
    }
}

function CellEditingHandler($sender, $e, [System.Windows.Controls.TabControl]$tabControl, [hashtable]$tabs) {
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
            $newTab.Content.Add_CellEditEnding({ param($sender,$e) CellEditingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        }
        $newTab.Content.ItemsSource.Add($editedObject)
        SortTabControl -TabControl $tabControl
    }
}

function CommandDialog([Command]$command) {
    # We only want to process the command if it is a PS script or function
    $type = GetCommandType -Command $command.Root
    if (($type -ne "Function") -and ($type -ne "External Script")) {
        return
    }

    # Parse the command for parameters and build the grid with them
    $command.Parameters = GetScriptBlockParameters -Command $command.Root
    BuildCommandGrid -Grid $script:UI.CommandGrid -Parameters $command.Parameters

    # Assign the command as the current command so that BtnCommandRun can obtain it
    $script:CurrentCommand = $command

    $script:UI.BoxCommandName.Text = $command.Root
    OpenCommandDialog
}

function BtnCommandRunClick([Command]$command, [System.Windows.Controls.Grid]$grid) {
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

    WriteLog $command.Full
    RunCommand $command.Full
    CloseCommandDialog
}

function OpenCommandDialog {
    $script:UI.Main.Opacity = "0.5"
    $script:UI.CommandDialog.Visibility = "Visible"
}

function CloseCommandDialog() {
    $script:UI.Main.Opacity = "100"
    $script:UI.CommandDialog.Visibility = "Hidden"
    $script:UI.CommandGrid.Children.Clear()
    $script:UI.CommandGrid.RowDefinitions.Clear()
}

function GetCommandType([string]$command) {
    $type = (Get-Command $command).CommandType 
    return $type
}

function GetScriptBlockParameters([string]$command) {
    $scriptBlock = (Get-Command $command).ScriptBlock
    $parsed = [System.Management.Automation.Language.Parser]::ParseInput($scriptBlock.ToString(), [ref]$null, [ref]$null)
    $parameters = $parsed.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)

    return $parameters
}

function ContainsAttributeType([System.Management.Automation.Language.ParameterAst]$parameter, [string]$typeName) {
    foreach ($attribute in $parameter.Attributes) {
        if ($attribute.TypeName.FullName -eq $typeName) {
            return $true
        }
    }
    return $false
}

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

function RunCommand([string]$command) {      
    # We must escape any quotation marks passed or it will cause problems being passed through Start-Process
    $command = $command -replace '"', '\"'
    Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit `" & { $command } `""
}

function AddToGrid([System.Windows.Controls.Grid]$grid, $element) {
    [void]$grid.Children.Add($element)
}

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

function SetTabsReadOnlyStatus([hashtable]$tabs) {
    $script:UI.BtnEdit.IsChecked = $script:TabsReadOnly
    $script:TabsReadOnly = (-not $script:TabsReadOnly)
    foreach ($tab in $tabs.GetEnumerator()) {
        $grid = $tab.Value.Content
        $grid.IsReadOnly = $script:TabsReadOnly
    }
}

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

function SetGridExtraColumnsVisibility([System.Windows.Controls.DataGrid]$grid) {
    foreach ($column in $grid.Columns) {
        foreach ($extraCol in $script:ExtraColumns) {
            if ($column.Header -eq $extraCol) {
                $column.Visibility = $script:ExtraColumnsVisibility
            }
        }
    }
}

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

function SortGridByColumn([System.Windows.Controls.DataGrid]$grid, [string]$columnName) {
    $grid.Items.SortDescriptions.Clear()
    $sort = New-Object System.ComponentModel.SortDescription($columnName, [System.ComponentModel.ListSortDirection]::Ascending)
    $grid.Items.SortDescriptions.Add($sort)
    $grid.Items.Refresh()
}

function GetHighestId([System.Object[]]$json) {
    $highest = 0
    foreach ($value in ($json | Select-Object -ExpandProperty id -Unique)) {
        if ($value -gt $greatest) {
            $highest = $value
        }
    }
    return $highest
}

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

function NewTab([string]$name) {
    $tabItem = New-Object System.Windows.Controls.TabItem
    $tabItem.Header = $name
    return $tabItem
}

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

function NewDataTab([string]$name, [System.Collections.ObjectModel.ObservableCollection[RowData]]$itemsSource, [System.Windows.Controls.TabControl]$tabControl) {
    $grid = NewDataGrid -Name $name -ItemsSource $itemsSource
    $tab = NewTab -Name $name
    $tab.Content = $grid
    [void]$tabControl.Items.Add($tab)
    return $tab
}

function NewLabel([string]$content, [string]$halign, [string]$valign) {
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $content
    $label.HorizontalAlignment = $halign
    $label.VerticalAlignment = $valign
    $label.Margin = New-Object System.Windows.Thickness(3)
    return $label
}

function NewToolTip([string]$content) {
    $tooltip = New-Object System.Windows.Controls.ToolTip
    $tooltip.Content = $content
    return $tooltip
}

function NewComboBox([string]$name, [System.String[]]$itemsSource, [string]$selectedItem) {
    $comboBox = New-Object System.Windows.Controls.ComboBox
    $comboBox.Name = $name
    $comboBox.Margin = New-Object System.Windows.Thickness(5)
    $comboBox.ItemsSource = $itemsSource
    $comboBox.SelectedItem = $selectedItem
    return $comboBox
}

function NewTextBox([string]$name, [string]$text) {
    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Name = $name
    $textBox.Margin = New-Object System.Windows.Thickness(5)
    $textBox.Text = $text
    return $textBox
}

function NewCheckBox([string]$name, [bool]$isChecked) {
    $checkbox = New-Object System.Windows.Controls.CheckBox
    $checkbox.Name = $name
    $checkbox.IsChecked = $isChecked
    return $checkbox
}

function NewButton([string]$content, [string]$halign, [int]$width) {
    $button = New-Object System.Windows.Controls.Button
    $button.Content = $content
    $button.Margin = New-Object System.Windows.Thickness(10)
    $button.HorizontalAlignment = $halign
    $button.Width = $width
    $button.IsDefault = $true
    return $button
}

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

function Show-ErrorMessageBox([string]$message) {
    Write-Error $message
    [System.Windows.MessageBox]::Show($message, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}

function WriteLog([string]$output) {
    $script:UI.Window.Dispatcher.Invoke([action]{$script:UI.LogBox.AppendText("$output`n")}, "Normal")
}

class RowData {
    [int]$Id
    [string]$Name
    [string]$Description
    [string]$Category
    [string]$Command
    [bool]$SkipParameterSelect
    [string]$PreCommand
}

class Command {
    [string]$Root
    [string]$Full
    [System.Object[]]$Parameters
}

Initialize
MainWindow