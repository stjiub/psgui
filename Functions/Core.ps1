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
        elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
            # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
            if (-not $script:State.TabsReadOnly) {
                $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                $e.Handled = $true
            }
            # If Edit Mode is disabled (TabsReadOnly is true), run the command
            else {
                $e.Handled = $true
                Invoke-MainRunClick -TabControl $script:UI.TabControl
            }
        }
    })
    [void]$script:UI.Tabs.Add("All", $allTab)

    $favItemsSource = [System.Collections.ObjectModel.ObservableCollection[FavoriteRowData]]::new()
    $loadedFavorites = Load-Favorites -AllData $json
    foreach ($fav in $loadedFavorites) {
        [void]$favItemsSource.Add($fav)
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
        elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
            # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
            if (-not $script:State.TabsReadOnly) {
                $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                $e.Handled = $true
            }
            # If Edit Mode is disabled (TabsReadOnly is true), run the command
            else {
                $e.Handled = $true
                Invoke-MainRunClick -TabControl $script:UI.TabControl
            }
        }
    })

    # Add drag/drop event handlers for reordering favorites
    Initialize-FavoritesDragDrop -Grid $favTab.Content

    [void]$script:UI.Tabs.Add("Favorites", $favTab)
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
            elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem) {
                # If Edit Mode is enabled (TabsReadOnly is false), commit any pending edits
                if (-not $script:State.TabsReadOnly) {
                    $sender.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
                    $e.Handled = $true
                }
                # If Edit Mode is disabled (TabsReadOnly is true), run the command
                else {
                    $e.Handled = $true
                    Invoke-MainRunClick -TabControl $script:UI.TabControl
                }
            }
        })
        [void]$script:UI.Tabs.Add($category, $tab)
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
    $script:UI.BtnMenuDuplicate.Add_Click({ Duplicate-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuRemove.Add_Click({ Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuUndoDelete.Add_Click({ Restore-DeletedCommand -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuSave.Add_Click({ Save-DataFile -FilePath $script:State.CurrentDataFile -Data ($script:UI.Tabs["All"].Content.ItemsSource) })
    $script:UI.BtnMenuSaveAs.Add_Click({ Save-DataFileAs })
    $script:UI.BtnMenuOpen.Add_Click({ Open-DataFile })
    $script:UI.BtnMenuImport.Add_Click({ Invoke-ImportDataFileDialog })
    $script:UI.BtnMenuEdit.Add_Click({ Toggle-EditMode -Tabs $script:UI.Tabs })
    $script:UI.BtnToggleEditMode.Add_Click({ Toggle-EditMode -Tabs $script:UI.Tabs })
    $script:UI.BtnMenuFavorite.Add_Click({ Toggle-CommandFavorite })
    $script:UI.BtnMenuSettings.Add_Click({ Show-SettingsDialog })
    $script:UI.BtnMenuRunOpen.Add_Click({
        Invoke-MainRunClick -TabControl $script:UI.TabControl
    })
    $script:UI.BtnMenuRunDetached.Add_Click({
        $script:State.RunCommandAttached = $false
        Invoke-MainRunClick -TabControl $script:UI.TabControl
    })
    $script:UI.BtnMenuRunAttached.Add_Click({
        $script:State.RunCommandAttached = $true
        Invoke-MainRunClick -TabControl $script:UI.TabControl -Attached $true
    })
    $script:UI.BtnMenuRunInject.Add_Click({
        Invoke-MainInjectClick -TabControl $script:UI.TabControl
    })
    $script:UI.BtnMenuRunRerunLast.Add_Click({
        if ($script:State.CommandHistory -and $script:State.CommandHistory.Count -gt 0) {
            $lastHistoryEntry = $script:State.CommandHistory[0]
            Reopen-CommandFromHistory -HistoryEntry $lastHistoryEntry
        }
        else {
            Write-Status "No command history available"
        }
    })

    # Main Buttons
    $script:UI.BtnMainRun.Add_Click({
        $script:State.RunCommandAttached = $script:Settings.DefaultRunCommandAttached
        Invoke-MainRunClick -TabControl $script:UI.TabControl
    })
    $script:UI.BtnToggleShell.Add_Click({ Toggle-ShellGrid })

    # Command dialog button events - Now handled per-window in New-CommandWindow

    # Settings dialog button events
    $script:UI.BtnBrowseLogs.Add_Click({ Invoke-BrowsePath -TextBox $script:UI.TxtDefaultLogsPath -IsFolder })
    $script:UI.BtnBrowseDataFile.Add_Click({ Invoke-BrowsePath -TextBox $script:UI.TxtDefaultDataFile -Filter "JSON files (*.json)|*.json|All files (*.*)|*.*" -DefaultExt ".json" })
    $script:UI.BtnBrowseSettings.Add_Click({ Invoke-BrowsePath -TextBox $script:UI.TxtSettingsPath -Filter "JSON files (*.json)|*.json|All files (*.*)|*.*" -DefaultExt ".json" })
    $script:UI.BtnBrowseFavorites.Add_Click({ Invoke-BrowsePath -TextBox $script:UI.TxtFavoritesPath -Filter "JSON files (*.json)|*.json|All files (*.*)|*.*" -DefaultExt ".json" })
    $script:UI.BtnBrowseHistory.Add_Click({ Invoke-BrowsePath -TextBox $script:UI.TxtDefaultHistoryPath -Filter "JSON files (*.json)|*.json|All files (*.*)|*.*" -DefaultExt ".json" })
    $script:UI.BtnBrowseProfilePath.Add_Click({ Invoke-BrowsePath -TextBox $script:UI.TxtProfilePath -Filter "PowerShell Profile (*.ps1)|*.ps1|All files (*.*)|*.*" -DefaultExt ".ps1" })
    $script:UI.BtnApplySettings.Add_Click({ Apply-Settings })
    $script:UI.BtnCloseSettings.Add_Click({ Hide-SettingsDialog })

    # Main Tab Control events
    $script:UI.TabControl.Add_SelectionChanged({
        param($sender, $e)
        Handle-TabSelection -SelectedTab $sender.SelectedItem
        Update-MainRunButtonText
        # Re-apply the filter when tab changes
        if ($script:UI.TxtSearchFilter) {
            Invoke-GridFilter -SearchText $script:UI.TxtSearchFilter.Text
        }
    })

    # Search filter text box event
    $script:UI.TxtSearchFilter.Add_TextChanged({
        param($sender, $e)
        Invoke-GridFilter -SearchText $sender.Text
    })

    # Process Tab events
    # Add context menu to the "+" tab
    $addTabContextMenu = New-Object System.Windows.Controls.ContextMenu
    $addTabContextMenu.FontSize = 12

    # New PS Session menu item
    $menuNewSession = New-Object System.Windows.Controls.MenuItem
    $menuNewSession.Header = "New PS Session"
    $menuNewSession.FontSize = 12

    # Create icon for New PS Session
    $iconNew = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconNew.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Plus
    $iconNew.Width = 16
    $iconNew.Height = 16
    $iconNew.Margin = New-Object System.Windows.Thickness(0)
    $menuNewSession.Icon = $iconNew

    $menuNewSession.Add_Click({
        New-ProcessTab -TabControl $script:UI.PSTabControl -Process $script:Settings.DefaultShell -ProcessArgs $script:Settings.DefaultShellArgs
    })
    [void]$addTabContextMenu.Items.Add($menuNewSession)

    # Attach PS Session menu item
    $menuAttachSession = New-Object System.Windows.Controls.MenuItem
    $menuAttachSession.Header = "Attach PS Session"
    $menuAttachSession.FontSize = 12

    # Create icon for Attach PS Session
    $iconAttach = New-Object MaterialDesignThemes.Wpf.PackIcon
    $iconAttach.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Import
    $iconAttach.Width = 16
    $iconAttach.Height = 16
    $iconAttach.Margin = New-Object System.Windows.Thickness(0)
    $menuAttachSession.Icon = $iconAttach

    $menuAttachSession.Add_Click({
        Show-AttachWindow
    })
    [void]$addTabContextMenu.Items.Add($menuAttachSession)

    $script:UI.PSAddTab.ContextMenu = $addTabContextMenu

    # Right-click on PSAddTab shows context menu
    $script:UI.PSAddTab.Add_PreviewMouseRightButtonDown({
        param($sender, $e)
        $sender.ContextMenu.IsOpen = $true
        $e.Handled = $true
    })

    # Left-click on PSAddTab also shows context menu (instead of immediately creating tab)
    $script:UI.PSAddTab.Add_PreviewMouseLeftButtonDown({
        param($sender, $e)
        $sender.ContextMenu.IsOpen = $true
        $e.Handled = $true
    })
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

    # Command History events
    Initialize-CommandHistoryUI
    Load-CommandHistory

    $script:UI.Window.Add_Loaded({
        $script:UI.Window.Icon = $script:ApplicationPaths.IconFile
        Update-WindowTitle

        if (-not $script:Settings.OpenShellAtStart) {
            Toggle-ShellGrid
        }
    })

    $script:UI.Window.Add_Closing({ param($sender, $e) Invoke-WindowClosing -Sender $sender -E $e })
}

# Update the main Run button text and menu items visibility based on the selected command
function Update-MainRunButtonText {
    if (-not $script:UI.TabControl.SelectedItem) {
        $script:UI.BtnMainRun.Content = "Run"
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunInject.Visibility = [System.Windows.Visibility]::Collapsed
        return
    }

    $grid = $script:UI.TabControl.SelectedItem.Content
    if (-not $grid -or -not $grid.SelectedItem) {
        $script:UI.BtnMainRun.Content = "Run"
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunInject.Visibility = [System.Windows.Visibility]::Collapsed
        return
    }

    $selectedItem = $grid.SelectedItem

    if ($selectedItem.SkipParameterSelect) {
        # Show Run (Attached) or Run (Detached) based on DefaultRunCommandAttached setting
        if ($script:Settings.DefaultRunCommandAttached) {
            $script:UI.BtnMainRun.Content = "Run (Attached)"
        }
        else {
            $script:UI.BtnMainRun.Content = "Run (Detached)"
        }
        # Hide Open menu item, show Run and Inject items
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Collapsed
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunInject.Visibility = [System.Windows.Visibility]::Visible
    }
    else {
        # Show "Open" when SkipParameterSelect is false
        $script:UI.BtnMainRun.Content = "Open"
        # Show Open menu item, hide Run and Inject items
        $script:UI.BtnMenuRunOpen.Visibility = [System.Windows.Visibility]::Visible
        $script:UI.BtnMenuRunAttached.Visibility = [System.Windows.Visibility]::Collapsed
        $script:UI.BtnMenuRunDetached.Visibility = [System.Windows.Visibility]::Collapsed
        $script:UI.BtnMenuRunInject.Visibility = [System.Windows.Visibility]::Collapsed
    }
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

    # Close all open CommandWindows
    foreach ($window in $script:State.OpenCommandWindows) {
        try {
            if ($window -and -not $window.IsClosed) {
                $window.Close()
            }
        }
        catch {
            Write-Log "Error closing CommandWindow: $_"
        }
    }

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