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

# Create a styled context menu for PowerShell tabs
function New-PSTabContextMenu {
    param (
        [System.Windows.Controls.TabItem]$Tab
    )

    $contextMenu = New-Object System.Windows.Controls.ContextMenu

    # Apply the same style as History context menu
    $contextMenu.FontSize = 12

    # Close Tab menu item
    $menuCloseTab = New-Object System.Windows.Controls.MenuItem
    $menuCloseTab.Header = "Close Tab"
    $menuCloseTab.FontSize = 12

    # Create icon for Close Tab
    $iconClose = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconClose.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Close
    $iconClose.Width = 16
    $iconClose.Height = 16
    $iconClose.Margin = New-Object System.Windows.Thickness(0)
    $menuCloseTab.Icon = $iconClose

    $menuCloseTab.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag -and $tab.Tag["Process"]) {
            try {
                $script:UI.PSTabControl.Items.Remove($tab)
                $tab.Tag["Process"].Kill()
                Write-Status "PowerShell tab closed"
            }
            catch {
                Write-ErrorMessage "Failed to close PowerShell tab: $_"
            }
        }
    })
    [void]$contextMenu.Items.Add($menuCloseTab)

    # Close All Tabs menu item
    $menuCloseAllTabs = New-Object System.Windows.Controls.MenuItem
    $menuCloseAllTabs.Header = "Close All Tabs"
    $menuCloseAllTabs.FontSize = 12

    # Create icon for Close All Tabs
    $iconCloseAll = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconCloseAll.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::CloseBoxMultiple
    $iconCloseAll.Width = 16
    $iconCloseAll.Height = 16
    $iconCloseAll.Margin = New-Object System.Windows.Thickness(0)
    $menuCloseAllTabs.Icon = $iconCloseAll

    $menuCloseAllTabs.Add_Click({
        param($menuSender, $menuArgs)
        # Get all tabs except the "+" add tab
        $tabsToClose = @($script:UI.PSTabControl.Items | Where-Object { $_ -ne $script:UI.PSAddTab })

        foreach ($tab in $tabsToClose) {
            if ($tab.Tag -and $tab.Tag["Process"]) {
                try {
                    $script:UI.PSTabControl.Items.Remove($tab)
                    $tab.Tag["Process"].Kill()
                }
                catch {
                    Write-ErrorMessage "Failed to close PowerShell tab: $_"
                }
            }
        }

        if ($tabsToClose.Count -gt 0) {
            Write-Status "Closed $($tabsToClose.Count) PowerShell tab(s)"
        }
    })
    [void]$contextMenu.Items.Add($menuCloseAllTabs)

    # Detach Tab menu item
    $menuDetachTab = New-Object System.Windows.Controls.MenuItem
    $menuDetachTab.Header = "Detach Tab"
    $menuDetachTab.FontSize = 12

    # Create icon for Detach Tab
    $iconDetach = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconDetach.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Export
    $iconDetach.Width = 16
    $iconDetach.Height = 16
    $iconDetach.Margin = New-Object System.Windows.Thickness(0)
    $menuDetachTab.Icon = $iconDetach

    $menuDetachTab.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag) {
            $script:UI.PSTabControl.SelectedItem = $tab
            Detach-CurrentTab
            Write-Status "PowerShell tab detached"
        }
    })
    [void]$contextMenu.Items.Add($menuDetachTab)

    # Open Log menu item
    $menuOpenLog = New-Object System.Windows.Controls.MenuItem
    $menuOpenLog.Header = "Open Log"
    $menuOpenLog.FontSize = 12

    # Create icon for Open Log
    $iconLog = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconLog.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::FileDocumentOutline
    $iconLog.Width = 16
    $iconLog.Height = 16
    $iconLog.Margin = New-Object System.Windows.Thickness(0)
    $menuOpenLog.Icon = $iconLog

    $menuOpenLog.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag -and $tab.Tag["HistoryEntry"]) {
            $historyEntry = $tab.Tag["HistoryEntry"]
            if ($historyEntry.LogPath -and (Test-Path $historyEntry.LogPath)) {
                New-LogMonitorTab -FilePath $historyEntry.LogPath -TabControl $script:UI.LogTabControl

                # Switch to the Logs tab
                $logsTab = $script:UI.TabControlShell.Items | Where-Object { $_.Header -eq "Logs" }
                if ($logsTab) {
                    $script:UI.TabControlShell.SelectedItem = $logsTab
                }

                Write-Status "Opened log file"
            }
            else {
                Write-ErrorMessage "Log file not found or command was not logged"
            }
        }
        else {
            Write-ErrorMessage "No log associated with this tab"
        }
    })
    [void]$contextMenu.Items.Add($menuOpenLog)

    # Go to History menu item
    $menuGoToHistory = New-Object System.Windows.Controls.MenuItem
    $menuGoToHistory.Header = "Go to History"
    $menuGoToHistory.FontSize = 12

    # Create icon for Go to History
    $iconHistory = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconHistory.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::History
    $iconHistory.Width = 16
    $iconHistory.Height = 16
    $iconHistory.Margin = New-Object System.Windows.Thickness(0)
    $menuGoToHistory.Icon = $iconHistory

    $menuGoToHistory.Add_Click({
        param($menuSender, $menuArgs)
        $tab = $menuSender.Parent.PlacementTarget
        if ($tab -and $tab.Tag -and $tab.Tag["HistoryEntry"]) {
            # Switch to History tab
            $historyTab = $script:UI.TabControlShell.Items | Where-Object { $_.Header -eq "History" }
            if ($historyTab) {
                $script:UI.TabControlShell.SelectedItem = $historyTab

                # Select the history entry in the grid
                $historyGrid = $script:UI.Window.FindName("CommandHistoryGrid")
                if ($historyGrid) {
                    $historyEntry = $tab.Tag["HistoryEntry"]
                    $historyGrid.SelectedItem = $historyEntry
                    $historyGrid.ScrollIntoView($historyEntry)
                    Write-Status "Jumped to command history"
                }
            }
        }
        else {
            Write-ErrorMessage "No history entry associated with this tab"
        }
    })
    [void]$contextMenu.Items.Add($menuGoToHistory)

    return $contextMenu
}

# Create a new embedded process under a Tab Control
function New-ProcessTab {
    param (
        $tabControl,
        $process,
        $processArgs,
        $tabName = "PS_$($tabControl.Items.Count)",
        [PSCustomObject]$historyEntry = $null
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
    $tabData["HistoryEntry"] = $historyEntry
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

    # Add context menu to the tab header
    $tab.ContextMenu = New-PSTabContextMenu -Tab $tab

    # Handle middle-click to detach tab
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

    # Add context menu to the tab header
    $tab.ContextMenu = New-PSTabContextMenu -Tab $tab

    # Handle middle-click to detach tab
    $tab.Add_PreviewMouseDown({
        param($sender, $e)
        if ($e.MiddleButton -eq 'Pressed') {
            Detach-CurrentTab
            $e.Handled = $true
        }
    })
}