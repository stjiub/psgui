function Save-Favorites {
    param (
        [System.Collections.ObjectModel.ObservableCollection[Object]]$favorites
    )

    try {
        $favoritesDir = Split-Path $script:Settings.FavoritesPath -Parent
        if (-not (Test-Path $favoritesDir)) {
            New-Item -ItemType Directory -Path $favoritesDir -Force | Out-Null
        }

        # Load existing favorites file or create new structure
        $allFavorites = @{}
        if (Test-Path $script:Settings.FavoritesPath) {
            try {
                $existingContent = Get-Content $script:Settings.FavoritesPath | ConvertFrom-Json
                # Handle both old format (array) and new format (object with command list IDs)
                if ($existingContent -is [Array]) {
                    # Convert old format - all favorites go under a default ID
                    $allFavorites["default"] = $existingContent
                } else {
                    # Convert PSCustomObject to hashtable for proper manipulation
                    $existingContent.PSObject.Properties | ForEach-Object {
                        $allFavorites[$_.Name] = $_.Value
                    }
                }
            }
            catch {
                Write-Log "Failed to parse existing favorites file, creating new one"
            }
        }

        # Save favorites for current command list
        $currentListId = $script:State.CurrentCommandListId
        if ($currentListId) {
            $favoriteData = $favorites | Select-Object Id, Order
            $allFavorites[$currentListId] = $favoriteData
        }

        # Convert hashtable to PSCustomObject for proper JSON serialization
        $outputObject = New-Object PSObject
        $allFavorites.GetEnumerator() | ForEach-Object {
            $outputObject | Add-Member -MemberType NoteProperty -Name $_.Key -Value $_.Value
        }

        $outputObject | ConvertTo-Json -Depth 3 | Set-Content $script:Settings.FavoritesPath
        Write-Status "Favorites saved"
    }
    catch {
        Write-Status "Failed to save favorites"
        Write-Log "Failed to save favorites: $_"
    }
}

function Load-Favorites {
    param (
        [System.Collections.ObjectModel.ObservableCollection[RowData]]$allData
    )

    try {
        if (Test-Path $script:Settings.FavoritesPath) {
            $allFavorites = Get-Content $script:Settings.FavoritesPath | ConvertFrom-Json
            $favorites = @()

            # Get favorites for current command list
            $currentListId = $script:State.CurrentCommandListId
            $favoriteData = $null

            # Handle both old format (array) and new format (object with command list IDs)
            if ($allFavorites -is [Array]) {
                # Old format - use as default
                $favoriteData = $allFavorites
            } else {
                # New format - get favorites for current command list
                if ($currentListId -and $allFavorites.PSObject.Properties[$currentListId]) {
                    $favoriteData = $allFavorites.$currentListId
                }
            }

            if ($favoriteData) {
                foreach ($fav in $favoriteData | Sort-Object Order) {
                    $rowData = $allData | Where-Object { $_.Id -eq $fav.Id }
                    if ($rowData) {
                        $favoriteRow = [FavoriteRowData]::new($rowData, $fav.Order)
                        $favorites += $favoriteRow
                        if ($fav.Order -gt $script:State.FavoritesHighestOrder) {
                            $script:State.FavoritesHighestOrder = $fav.Order
                        }
                    }
                }
            }
            return $favorites
        }
    }
    catch {
        Write-Log "Failed to load favorites: $_"
        return @()
    }
    return @()
}


# Update favorite highlighting across all tabs except the Favorites tab
function Update-FavoriteHighlighting {
    $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
    $favoriteIds = @($favorites | ForEach-Object { $_.Id })

    foreach ($tabEntry in $script:UI.Tabs.GetEnumerator()) {
        $tabName = $tabEntry.Key
        $tab = $tabEntry.Value

        # Skip the Favorites tab since it only contains favorites
        if ($tabName -eq "Favorites") { continue }

        $grid = $tab.Content
        if ($grid -and $grid.Items) {
            # Use Dispatcher to ensure UI updates happen on the UI thread
            $script:UI.Window.Dispatcher.Invoke([action]{
                foreach ($item in $grid.Items) {
                    $container = $grid.ItemContainerGenerator.ContainerFromItem($item)
                    if ($container -is [System.Windows.Controls.DataGridRow]) {
                        if ($favoriteIds -contains $item.Id) {
                            $container.Tag = "IsFavorite"
                        }
                        else {
                            $container.Tag = $null
                        }
                    }
                }
            }, "Normal")
        }
    }
}

# Initialize drag and drop functionality for the Favorites grid
function Initialize-FavoritesDragDrop {
    param (
        [System.Windows.Controls.DataGrid]$grid
    )

    # Enable drag/drop on the grid
    $grid.AllowDrop = $true

    # Handle mouse down to capture the item being dragged
    $grid.Add_PreviewMouseLeftButtonDown({
        param($sender, $e)

        $row = Get-DataGridRowFromPoint -Grid $sender -Point ($e.GetPosition($sender))
        if ($row -and $row.Item) {
            $script:State.DragDrop.DraggedItem = $row.Item
        }
    })

    # Handle mouse move to initiate drag operation
    $grid.Add_MouseMove({
        param($sender, $e)

        if ($e.LeftButton -eq [System.Windows.Input.MouseButtonState]::Pressed -and
            $script:State.DragDrop.DraggedItem -ne $null) {

            $dragData = New-Object System.Windows.DataObject([System.Windows.DataFormats]::Serializable, $script:State.DragDrop.DraggedItem)
            [System.Windows.DragDrop]::DoDragDrop($sender, $dragData, [System.Windows.DragDropEffects]::Move)
        }
    })

    # Handle drag over to show drop feedback
    $grid.Add_DragOver({
        param($sender, $e)

        $position = $e.GetPosition($sender)
        $row = Get-DataGridRowFromPoint -Grid $sender -Point $position

        if ($row) {
            # Normal drop - highlight top border of target row
            # Only update highlighting if the row object or border type changed
            if ($row -ne $script:State.DragDrop.LastHighlightedRow -or $script:State.DragDrop.IsBottomBorder -eq $true) {
                Clear-DropHighlight
                $script:State.DragDrop.IsBottomBorder = $false
                Set-DropHighlight -Row $row -IsBottomBorder $false
            }
            $e.Effects = [System.Windows.DragDropEffects]::Move
        }
        elseif (Test-IsPositionBelowLastRow -Grid $sender -Position $position) {
            # Drop after last item - highlight bottom border of last row
            $lastRow = $sender.ItemContainerGenerator.ContainerFromIndex($sender.Items.Count - 1)

            # Only update if we're not already highlighting this row's bottom border
            if ($lastRow -ne $script:State.DragDrop.LastHighlightedRow -or $script:State.DragDrop.IsBottomBorder -eq $false) {
                Clear-DropHighlight
                $script:State.DragDrop.IsBottomBorder = $true
                Set-DropHighlight -Row $lastRow -IsBottomBorder $true
            }
            $e.Effects = [System.Windows.DragDropEffects]::Move
        }
        else {
            $e.Effects = [System.Windows.DragDropEffects]::None
        }
        $e.Handled = $true
    })

    # Handle drag leave to clear feedback
    $grid.Add_DragLeave({
        param($sender, $e)
        Clear-DropHighlight
        $script:State.DragDrop.IsBottomBorder = $false
    })

    # Handle drop to reorder items
    $grid.Add_Drop({
        param($sender, $e)

        Clear-DropHighlight

        if ($script:State.DragDrop.DraggedItem -ne $null) {
            $itemsSource = $sender.ItemsSource
            $draggedItem = $script:State.DragDrop.DraggedItem
            $position = $e.GetPosition($sender)

            $targetRow = Get-DataGridRowFromPoint -Grid $sender -Point $position
            $targetItem = $null
            $isDropAfterLast = $false

            # Determine drop target and whether dropping after last item
            if (Test-IsPositionBelowLastRow -Grid $sender -Position $position) {
                # Dropping after last item
                $targetItem = $sender.Items[$sender.Items.Count - 1]
                $isDropAfterLast = $true
            }
            elseif ($targetRow) {
                # Normal drop on a row
                $targetItem = $targetRow.Item
            }

            if ($targetItem -and ($draggedItem -ne $targetItem -or $isDropAfterLast)) {
                $draggedOrder = $draggedItem.Order
                $targetOrder = $targetItem.Order

                if ($isDropAfterLast) {
                    # Move to end - set Order to current maximum
                    $maxOrder = ($itemsSource | Measure-Object -Property Order -Maximum).Maximum

                    # Shift all items after the dragged item up by one
                    foreach ($item in $itemsSource) {
                        if ($item.Order -gt $draggedOrder) {
                            $item.Order--
                        }
                    }
                    $draggedItem.Order = $maxOrder
                }
                elseif ($draggedOrder -lt $targetOrder) {
                    # Moving down - shift items between old and new position up
                    foreach ($item in $itemsSource) {
                        if ($item.Order -gt $draggedOrder -and $item.Order -le $targetOrder) {
                            $item.Order--
                        }
                    }
                    $draggedItem.Order = $targetOrder
                }
                elseif ($draggedOrder -gt $targetOrder) {
                    # Moving up - shift items between new and old position down
                    foreach ($item in $itemsSource) {
                        if ($item.Order -ge $targetOrder -and $item.Order -lt $draggedOrder) {
                            $item.Order++
                        }
                    }
                    $draggedItem.Order = $targetOrder
                }

                # Refresh the sort to reflect new order
                $sender.Items.SortDescriptions.Clear()
                $sortDescription = New-Object System.ComponentModel.SortDescription("Order", [System.ComponentModel.ListSortDirection]::Ascending)
                $sender.Items.SortDescriptions.Add($sortDescription)
                $sender.Items.Refresh()

                # Save favorites and keep selection
                Save-Favorites -Favorites $itemsSource
                $sender.SelectedItem = $draggedItem
                $sender.ScrollIntoView($draggedItem)
            }
        }

        # Reset drag state
        $script:State.DragDrop.DraggedItem = $null
        $script:State.DragDrop.IsBottomBorder = $false
        $e.Handled = $true
    })
}

# Get the DataGridRow at a specific point in the grid
function Get-DataGridRowFromPoint {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [System.Windows.Point]$point
    )

    # First try hit testing to find row from actual content
    $element = $grid.InputHitTest($point)
    while ($element -ne $null) {
        if ($element -is [System.Windows.Controls.DataGridRow]) {
            return $element
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }

    # If hit testing didn't find a row, iterate through rows to find the closest by Y position
    # This handles padding/margin areas between row content
    if ($grid.Items.Count -gt 0) {
        for ($i = 0; $i -lt $grid.Items.Count; $i++) {
            $row = $grid.ItemContainerGenerator.ContainerFromIndex($i)
            if ($row) {
                $rowPosition = $row.TranslatePoint([System.Windows.Point]::new(0, 0), $grid)
                $rowBottom = $rowPosition.Y + $row.ActualHeight

                # Check if point is within this row's vertical bounds
                if ($point.Y -ge $rowPosition.Y -and $point.Y -lt $rowBottom) {
                    return $row
                }
            }
        }
    }

    return $null
}

# Check if the mouse position is below the last row (for dropping at the end)
function Test-IsPositionBelowLastRow {
    param (
        [System.Windows.Controls.DataGrid]$grid,
        [System.Windows.Point]$position
    )

    if ($grid.Items.Count -eq 0) {
        return $false
    }

    $lastRow = $grid.ItemContainerGenerator.ContainerFromIndex($grid.Items.Count - 1)
    if ($lastRow) {
        $lastRowTop = $lastRow.TranslatePoint([System.Windows.Point]::new(0, 0), $grid).Y
        return $position.Y -gt $lastRowTop
    }

    return $false
}

# Set visual feedback for drop target
function Set-DropHighlight {
    param (
        [System.Windows.Controls.DataGridRow]$row,
        [bool]$isBottomBorder = $false
    )

    if ($row) {
        # Use app theme color from XAML resources
        $row.BorderBrush = $script:UI.Window.FindResource("AppPrimaryBrush")
        if ($isBottomBorder) {
            # Highlight bottom border for "drop after last item"
            $row.BorderThickness = New-Object System.Windows.Thickness(0, 0, 0, 2)
        } else {
            # Highlight top border for normal drops
            $row.BorderThickness = New-Object System.Windows.Thickness(0, 2, 0, 0)
        }
        $script:State.DragDrop.LastHighlightedRow = $row
    }
}

# Clear drop target visual feedback
function Clear-DropHighlight {
    if ($script:State.DragDrop.LastHighlightedRow) {
        $script:State.DragDrop.LastHighlightedRow.BorderThickness = New-Object System.Windows.Thickness(0)
        $script:State.DragDrop.LastHighlightedRow = $null
    }
}