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

    # Handle tab closure
    $tab.Add_PreviewMouseRightButtonDown({
        param($sender, $eventArgs)
        if ($eventArgs.ChangedButton -eq 'Right') {
            $script:UI.PSTabControl.Items.Remove($sender)
            Detach-CurrentTab
        }
    })
}