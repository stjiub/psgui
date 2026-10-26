# Command Injection Functions for sending commands to active PowerShell tabs

# Check if there is an active PowerShell tab with an attached session
function Test-ActivePowerShellTab {
    if (-not $script:UI.PSTabControl) {
        return $false
    }

    $selectedTab = $script:UI.PSTabControl.SelectedItem

    # Check if we have a valid tab selected (not the "+" add tab)
    if (-not $selectedTab -or $selectedTab -eq $script:UI.PSAddTab) {
        return $false
    }

    # Check if the tab has a valid process
    if (-not $selectedTab.Tag -or -not $selectedTab.Tag["Process"] -or -not $selectedTab.Tag["Handle"]) {
        return $false
    }

    # Check if the process is still running
    $process = $selectedTab.Tag["Process"]
    if ($process.HasExited) {
        return $false
    }

    return $true
}

# Inject a command into the active PowerShell tab using SendKeys
function Invoke-CommandInjection {
    param (
        [string]$CommandString
    )

    if (-not (Test-ActivePowerShellTab)) {
        Write-ErrorMessage "No active PowerShell tab available for command injection"
        return
    }

    try {
        # Load System.Windows.Forms assembly if not already loaded
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

        # Get the active PowerShell tab
        $selectedTab = $script:UI.PSTabControl.SelectedItem
        $psHandle = $selectedTab.Tag["Handle"]

        Write-Log "PowerShell window handle: $psHandle"

        # Bring window to foreground and set focus
        [Win32]::ShowWindow($psHandle, 5)  # SW_SHOW
        [Win32]::SetForegroundWindow($psHandle)
        [Win32]::SetFocus($psHandle)

        # Give it a moment to focus
        Start-Sleep -Milliseconds 500

        # Verify focus
        $focusedHandle = [Win32]::GetForegroundWindow()
        Write-Log "Current foreground window: $focusedHandle"

        # Send the command string followed by Enter
        # For SendKeys, we need to escape special characters that have meaning in SendKeys syntax
        # Characters that need escaping: + ^ % ~ ( ) { } [ ]
        $sb = New-Object System.Text.StringBuilder
        foreach ($char in $CommandString.ToCharArray()) {
            switch ($char) {
                '+' { [void]$sb.Append('{+}') }
                '^' { [void]$sb.Append('{^}') }
                '%' { [void]$sb.Append('{%}') }
                '~' { [void]$sb.Append('{~}') }
                '(' { [void]$sb.Append('{(}') }
                ')' { [void]$sb.Append('{)}') }
                '{' { [void]$sb.Append('{{}') }
                '}' { [void]$sb.Append('{}}') }
                '[' { [void]$sb.Append('{[}') }
                ']' { [void]$sb.Append('{]}') }
                default { [void]$sb.Append($char) }
            }
        }
        $escapedCommand = $sb.ToString()

        Write-Log "Original command: $CommandString"
        Write-Log "Escaped command: $escapedCommand"

        [System.Windows.Forms.SendKeys]::SendWait($escapedCommand)
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")

        Write-Status "Command injected into PowerShell tab: $($selectedTab.Header)"
        Write-Log "Injected command: $CommandString"
    }
    catch {
        Write-ErrorMessage "Failed to inject command: $_"
        Write-Log "Command injection error: $_"
    }
}

# Inject a command from the main grid
function Invoke-MainInjectClick {
    param (
        [System.Windows.Controls.TabControl]$tabControl
    )

    if (-not (Test-ActivePowerShellTab)) {
        Write-ErrorMessage "No active PowerShell tab available. Please create or select a PowerShell tab first."
        return
    }

    $grid = $tabControl.SelectedItem.Content
    $selection = $grid.SelectedItems
    $command = New-Object Command
    $command.Full = ""
    $command.Root = $selection.Command
    $command.PreCommand = $selection.PreCommand
    $command.PostCommand = $selection.PostCommand
    $command.SkipParameterSelect = $selection.SkipParameterSelect
    $command.Transcript = $selection.Transcript
    $command.PSTask = $selection.PSTask
    $command.PSTaskMode = $selection.PSTaskMode
    $command.PSTaskVisibilityLevel = $selection.PSTaskVisibilityLevel
    $command.ShellOverride = $selection.ShellOverride

    Write-Log "Preparing command for injection - Root: $($command.Root), SkipParameterSelect: $($command.SkipParameterSelect)"

    if ($command.Root) {
        if ($selection.SkipParameterSelect) {
            # Build the command string
            $command.Full = ""
            $command.CleanCommand = ""

            # Note: We don't add transcript logging for injected commands
            # since they're being run in an existing shell

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

            # Inject the command
            Invoke-CommandInjection -CommandString $command.CleanCommand
        }
        else {
            # For commands with parameters, we need to open the CommandWindow
            # and let the user fill in parameters before injecting
            Write-ErrorMessage "Cannot inject commands with parameters. Use 'Inject' button in the Command Window instead."
        }
    }
}

# Inject a command from the CommandWindow
function Invoke-CommandWindowInjectClick {
    param (
        [System.Windows.Window]$CommandWindow
    )

    if (-not (Test-ActivePowerShellTab)) {
        Write-ErrorMessage "No active PowerShell tab available. Please create or select a PowerShell tab first."
        return
    }

    # Get command and grid from the window
    $command = $CommandWindow.Tag.Command
    $commandWindowHash = @{
        CommandGrid = $CommandWindow.FindName("CommandGrid")
    }

    # Compile the command with parameters
    Compile-Command -Command $command -CommandWindow $commandWindowHash

    # Close the window
    $CommandWindow.Close()

    # Inject the command (use CleanCommand to avoid transcript commands)
    Invoke-CommandInjection -CommandString $command.CleanCommand
}
