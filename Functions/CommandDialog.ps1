# Handle the Main Run Button click event to run the selected command/launch the CommandDialog
function Invoke-MainRunClick {
    param (
        [System.Windows.Controls.TabControl]$tabControl
    )

    if (($script:State.RunCommandAttached) -and ($script:UI.Shell.Visibility -ne "Visible"))
    { 
        Toggle-ShellGrid 
    }

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
            Run-Command $command $script:State.RunCommandAttached
        }
        else {
            Start-CommandDialog -Command $command
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
    } 
    else {
        # Restore previous height and visibility
        $script:UI.Window.FindName("ShellRow").Height = New-Object System.Windows.GridLength($script:State.SubGridExpandedHeight)
        $script:UI.Shell.Visibility = "Visible"
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
            if (($type -ne "Function") -and ($type -ne "Detached Script")) {
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
    Run-Command $command $script:Settings.DefaultRunCommandAttached
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

    if ($script:State.RunCommandAttached) {
        New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs "-ExecutionPolicy Bypass -NoExit `" & { $escapedCommand } `"" -TabName $command.Root
    }
    else {
        Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit `" & { $escapedCommand } `""
    }
    $script:State.RunCommandAttached = $script:Settings.DefaultRunCommandAttached
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