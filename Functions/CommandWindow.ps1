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
    $command.PostCommand = $selection.PostCommand
    $command.SkipParameterSelect = $selection.SkipParameterSelect
    $command.Log = $selection.Log
    $command.ShellOverride = $selection.ShellOverride

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

            # Add PostCommand if it exists
            if ($command.PostCommand) {
                $command.Full += "; " + $command.PostCommand
                $command.CleanCommand += "; " + $command.PostCommand
            }

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

    # Add PostCommand if it exists
    if ($command.PostCommand) {
        $command.Full += "; " + $command.PostCommand
        $command.CleanCommand += "; " + $command.PostCommand
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

    # Determine which shell to use - ShellOverride takes precedence over DefaultShell
    $shellToUse = if ([string]::IsNullOrWhiteSpace($command.ShellOverride)) {
        $script:Settings.DefaultShell
    } else {
        $command.ShellOverride
    }
    Write-Log "Shell to use: $shellToUse (Override: $($command.ShellOverride), Default: $($script:Settings.DefaultShell))"

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

        # Asynchronously create the process tab - the loading indicator will stay animated
        # Note: New-ProcessTab automatically selects the newly created tab
        $psArgs = Get-PowerShellArguments -Command $escapedCommand -NoExit
        Write-Log "Running attached command: Shell=$shellToUse, Args=$psArgs"
        New-ProcessTab -TabControl $script:UI.PSTabControl -Process $shellToUse -ProcessArgs $psArgs -TabName $command.Root -HistoryEntry $historyEntry -OnComplete {
            # Hide loading indicator after tab is created
            Hide-LoadingIndicator
        }
    }
    else {
        $psArgs = Get-PowerShellArguments -Command $escapedCommand -NoExit
        Start-Process -FilePath $shellToUse -ArgumentList $psArgs
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
