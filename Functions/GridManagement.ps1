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

    # Create context menu
    $contextMenu = New-Object System.Windows.Controls.ContextMenu
    $contextMenuStyle = $script:UI.Window.FindResource("GridContextMenuStyle")
    $contextMenu.Style = $contextMenuStyle
    $menuItemStyle = $script:UI.Window.FindResource("GridContextMenuItemStyle")
    $iconStyle = $script:UI.Window.FindResource("ContextMenuIconStyle")

    if ($name -eq "*") {
        # Favorites tab - simplified menu items (drag-and-drop handles reordering)
        $openMenuItem = New-Object System.Windows.Controls.MenuItem
        $openMenuItem.Header = "Open"
        $openMenuItem.Style = $menuItemStyle
        $openIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $openIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::FileDocumentEditOutline
        $openIcon.Style = $iconStyle
        $openMenuItem.Icon = $openIcon
        $openMenuItem.Add_Click({
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($openMenuItem)

        $runAttachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runAttachedMenuItem.Header = "Run (Attached)"
        $runAttachedMenuItem.Style = $menuItemStyle
        $runAttachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runAttachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Play
        $runAttachedIcon.Style = $iconStyle
        $runAttachedMenuItem.Icon = $runAttachedIcon
        $runAttachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runAttachedMenuItem)

        $runDetachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runDetachedMenuItem.Header = "Run (Detached)"
        $runDetachedMenuItem.Style = $menuItemStyle
        $runDetachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runDetachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::OpenInNew
        $runDetachedIcon.Style = $iconStyle
        $runDetachedMenuItem.Icon = $runDetachedIcon
        $runDetachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $false
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runDetachedMenuItem)

        # Add event handler to update run/open visibility when context menu opens
        $contextMenu.Add_Opened({
            param($sender, $e)
            $currentGrid = $script:UI.TabControl.SelectedItem.Content
            $selectedItem = $currentGrid.SelectedItem
            if ($selectedItem) {
                # Update Run/Open menu item visibility based on SkipParameterSelect
                $openItem = $sender.Tag.OpenMenuItem
                $runAttachedItem = $sender.Tag.RunAttachedMenuItem
                $runDetachedItem = $sender.Tag.RunDetachedMenuItem

                if ($selectedItem.SkipParameterSelect) {
                    # Show Run (Attached) and Run (Detached), hide Open
                    $openItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Visible
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Visible
                } else {
                    # Show Open, hide Run (Attached) and Run (Detached)
                    $openItem.Visibility = [System.Windows.Visibility]::Visible
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
        })

        # Store references for dynamic visibility
        $contextMenu.Tag = @{
            OpenMenuItem = $openMenuItem
            RunAttachedMenuItem = $runAttachedMenuItem
            RunDetachedMenuItem = $runDetachedMenuItem
        }

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $favoriteMenuItem = New-Object System.Windows.Controls.MenuItem
        $favoriteMenuItem.Header = "Remove from Favorites"
        $favoriteMenuItem.Style = $menuItemStyle
        $favIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $favIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::StarOff
        $favIcon.Style = $iconStyle
        $favoriteMenuItem.Icon = $favIcon
        $favoriteMenuItem.Add_Click({ Toggle-CommandFavorite })
        [void]$contextMenu.Items.Add($favoriteMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $duplicateMenuItem = New-Object System.Windows.Controls.MenuItem
        $duplicateMenuItem.Header = "Duplicate Command"
        $duplicateMenuItem.Style = $menuItemStyle
        $duplicateIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $duplicateIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::ContentCopy
        $duplicateIcon.Style = $iconStyle
        $duplicateMenuItem.Icon = $duplicateIcon
        $duplicateMenuItem.Add_Click({ Duplicate-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($duplicateMenuItem)
    } else {
        # Regular tabs - standard menu items
        $openMenuItem = New-Object System.Windows.Controls.MenuItem
        $openMenuItem.Header = "Open"
        $openMenuItem.Style = $menuItemStyle
        $openIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $openIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::FileDocumentEditOutline
        $openIcon.Style = $iconStyle
        $openMenuItem.Icon = $openIcon
        $openMenuItem.Add_Click({
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($openMenuItem)

        $runAttachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runAttachedMenuItem.Header = "Run (Attached)"
        $runAttachedMenuItem.Style = $menuItemStyle
        $runAttachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runAttachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Play
        $runAttachedIcon.Style = $iconStyle
        $runAttachedMenuItem.Icon = $runAttachedIcon
        $runAttachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $true
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runAttachedMenuItem)

        $runDetachedMenuItem = New-Object System.Windows.Controls.MenuItem
        $runDetachedMenuItem.Header = "Run (Detached)"
        $runDetachedMenuItem.Style = $menuItemStyle
        $runDetachedIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $runDetachedIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::OpenInNew
        $runDetachedIcon.Style = $iconStyle
        $runDetachedMenuItem.Icon = $runDetachedIcon
        $runDetachedMenuItem.Add_Click({
            $script:State.RunCommandAttached = $false
            Invoke-MainRunClick -TabControl $script:UI.TabControl
        })
        [void]$contextMenu.Items.Add($runDetachedMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $favoriteMenuItem = New-Object System.Windows.Controls.MenuItem
        $favoriteMenuItem.Header = "Add to Favorites"
        $favoriteMenuItem.Style = $menuItemStyle
        $favIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $favIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Star
        $favIcon.Style = $iconStyle
        $favoriteMenuItem.Icon = $favIcon
        $favoriteMenuItem.Add_Click({ Toggle-CommandFavorite })

        # Store reference to favorite menu item and run/open items so we can update them
        $contextMenu.Tag = @{
            FavoriteMenuItem = $favoriteMenuItem
            IconStyle = $iconStyle
            OpenMenuItem = $openMenuItem
            RunAttachedMenuItem = $runAttachedMenuItem
            RunDetachedMenuItem = $runDetachedMenuItem
        }

        # Add event handler to update the favorite menu item text/icon and run/open visibility when context menu opens
        $contextMenu.Add_Opened({
            param($sender, $e)
            $currentGrid = $script:UI.TabControl.SelectedItem.Content
            $selectedItem = $currentGrid.SelectedItem
            if ($selectedItem -and $script:UI.Tabs["Favorites"]) {
                $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
                $existingFavorite = $favorites | Where-Object { $_.Id -eq $selectedItem.Id }

                # Get the favorite menu item from the context menu's tag
                $favMenuItem = $sender.Tag.FavoriteMenuItem
                $style = $sender.Tag.IconStyle

                # Create new icon each time to avoid reference issues
                $newFavIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
                $newFavIcon.Style = $style

                if ($existingFavorite) {
                    $favMenuItem.Header = "Remove from Favorites"
                    $newFavIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::StarOff
                } else {
                    $favMenuItem.Header = "Add to Favorites"
                    $newFavIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::Star
                }

                $favMenuItem.Icon = $newFavIcon

                # Update Run/Open menu item visibility based on SkipParameterSelect
                $openItem = $sender.Tag.OpenMenuItem
                $runAttachedItem = $sender.Tag.RunAttachedMenuItem
                $runDetachedItem = $sender.Tag.RunDetachedMenuItem

                if ($selectedItem.SkipParameterSelect) {
                    # Show Run (Attached) and Run (Detached), hide Open
                    $openItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Visible
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Visible
                } else {
                    # Show Open, hide Run (Attached) and Run (Detached)
                    $openItem.Visibility = [System.Windows.Visibility]::Visible
                    $runAttachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                    $runDetachedItem.Visibility = [System.Windows.Visibility]::Collapsed
                }
            }
        })

        [void]$contextMenu.Items.Add($favoriteMenuItem)

        [void]$contextMenu.Items.Add((New-Object System.Windows.Controls.Separator))

        $addMenuItem = New-Object System.Windows.Controls.MenuItem
        $addMenuItem.Header = "Add Command"
        $addMenuItem.Style = $menuItemStyle
        $addIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $addIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::AddBox
        $addIcon.Style = $iconStyle
        $addMenuItem.Icon = $addIcon
        $addMenuItem.Add_Click({ Add-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($addMenuItem)

        $duplicateMenuItem = New-Object System.Windows.Controls.MenuItem
        $duplicateMenuItem.Header = "Duplicate Command"
        $duplicateMenuItem.Style = $menuItemStyle
        $duplicateIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $duplicateIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::ContentCopy
        $duplicateIcon.Style = $iconStyle
        $duplicateMenuItem.Icon = $duplicateIcon
        $duplicateMenuItem.Add_Click({ Duplicate-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($duplicateMenuItem)

        $removeMenuItem = New-Object System.Windows.Controls.MenuItem
        $removeMenuItem.Header = "Remove Command"
        $removeMenuItem.Style = $menuItemStyle
        $removeIcon = New-Object MaterialDesignThemes.Wpf.PackIcon
        $removeIcon.Kind = [MaterialDesignThemes.Wpf.PackIconKind]::TrashCan
        $removeIcon.Style = $iconStyle
        $removeMenuItem.Icon = $removeIcon
        $removeMenuItem.Add_Click({ Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
        [void]$contextMenu.Items.Add($removeMenuItem)

        # Update Tag to include remove menu item
        $contextMenu.Tag.RemoveMenuItem = $removeMenuItem
    }

    # Update context menu when it opens to show/hide and update text for remove menu item
    $contextMenu.Add_Opened({
        param($sender, $e)

        # Only update remove menu item for regular tabs (not Favorites tab)
        if ($sender.Tag.RemoveMenuItem) {
            $currentGrid = $script:UI.TabControl.SelectedItem.Content
            $selectedCount = $currentGrid.SelectedItems.Count
            $removeItem = $sender.Tag.RemoveMenuItem

            if ($selectedCount -eq 0) {
                $removeItem.Visibility = [System.Windows.Visibility]::Collapsed
            } else {
                $removeItem.Visibility = [System.Windows.Visibility]::Visible
                if ($selectedCount -eq 1) {
                    $removeItem.Header = "Remove Command"
                } else {
                    $removeItem.Header = "Remove $selectedCount Commands"
                }
            }
        }
    })

    $grid.ContextMenu = $contextMenu
    $grid.AutoGenerateColumns = $false

    # Add selection changed event to update the Run button text
    $grid.Add_SelectionChanged({
        Update-MainRunButtonText
    })

    # Apply the favorite row style (skip for Favorites tab since it only contains favorites)
    if ($name -ne "*") {
        $rowStyle = $script:UI.Window.FindResource("FavoriteRowStyle")
        $grid.RowStyle = $rowStyle

        # Add event handler to set favorite highlighting when rows are loaded
        $grid.Add_LoadingRow({
            param($sender, $e)
            if ($script:UI.Tabs -and $script:UI.Tabs["Favorites"]) {
                $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
                $favoriteIds = @($favorites | ForEach-Object { $_.Id })

                $rowItem = $e.Row.Item
                if ($favoriteIds -contains $rowItem.Id) {
                    $e.Row.Tag = "IsFavorite"
                }
                else {
                    $e.Row.Tag = $null
                }
            }
        })
    }

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

    # Create a checkbox column for SkipParameterSelect and Log
    if ($propertyName -eq "SkipParameterSelect" -or $propertyName -eq "Log") {
        $column = New-Object System.Windows.Controls.DataGridCheckBoxColumn
        $column.Header = $propertyName
        $binding = New-Object System.Windows.Data.Binding $propertyName
        $binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
        $binding.UpdateSourceTrigger = [System.Windows.Data.UpdateSourceTrigger]::PropertyChanged
        $column.Binding = $binding
    }
    else {
        $column = New-Object System.Windows.Controls.DataGridTextColumn
        $column.Header = $propertyName
        $column.Binding = New-Object System.Windows.Data.Binding $propertyName
    }

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

    $script:UI.BtnMenuEdit.IsChecked = $script:State.TabsReadOnly
    $script:State.TabsReadOnly = (-not $script:State.TabsReadOnly)

    # Sync both toggle buttons (MaterialDesign will handle icon switching automatically)
    $script:UI.BtnToggleEditMode.IsChecked = (-not $script:State.TabsReadOnly)

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