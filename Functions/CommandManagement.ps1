# Handle the Main Window Add Button click event to add a new RowData object to the collection
function Add-CommandRow {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $newRow = New-Object RowData
    $newRow.Id = ++$script:State.HighestId
    $tab = $tabs["All"]
    $grid = $tab.Content
    $grid.ItemsSource.Add($newRow)
    Set-UnsavedChanges $true
    $tabControl.SelectedItem = $tab
    # We don't want to change the tabs read only status if they are already in edit mode
    if ($script:State.TabsReadOnly) {
        Set-TabsReadOnlyStatus -Tabs $tabs
        Set-TabsExtraColumnsVisibility -Tabs $tabs
    }
    # Select the new row and set it as the current item
    $grid.SelectedItem = $newRow
    $grid.CurrentItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.Focus()
    # Update the layout to ensure the selection is processed
    $grid.UpdateLayout()
    # Set the current cell to the Name column of the new row
    $nameColumn = $grid.Columns | Where-Object { $_.Header -eq "Name" } | Select-Object -First 1
    if ($nameColumn) {
        $grid.CurrentCell = New-Object System.Windows.Controls.DataGridCellInfo($newRow, $nameColumn)
    }
    $grid.BeginEdit()
}

# Handle the Duplicate Command to create a copy of the selected command row
function Duplicate-CommandRow {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $grid = $tabControl.SelectedItem.Content
    $selectedItem = $grid.SelectedItem

    if (-not $selectedItem) {
        Write-Status "No command selected to duplicate"
        return
    }

    # Create a new row with a new ID
    $newRow = New-Object RowData
    $newRow.Id = ++$script:State.HighestId

    # Copy all properties except Id
    $newRow.Name = $selectedItem.Name
    $newRow.Description = $selectedItem.Description
    $newRow.Category = $selectedItem.Category
    $newRow.Command = $selectedItem.Command
    $newRow.SkipParameterSelect = $selectedItem.SkipParameterSelect
    $newRow.PreCommand = $selectedItem.PreCommand

    # Add to All tab
    $allTab = $tabs["All"]
    $allGrid = $allTab.Content
    $allGrid.ItemsSource.Add($newRow)

    # If the item has a category, add to category tab as well
    if ($newRow.Category) {
        $categoryTab = $tabs[$newRow.Category]
        if ($categoryTab) {
            $categoryTab.Content.ItemsSource.Add($newRow)
        }
    }

    Set-UnsavedChanges $true

    # We don't want to change the tabs read only status if they are already in edit mode
    if ($script:State.TabsReadOnly) {
        Set-TabsReadOnlyStatus -Tabs $tabs
        Set-TabsExtraColumnsVisibility -Tabs $tabs
    }

    # Select the new row and set it as the current item
    $grid.SelectedItem = $newRow
    $grid.CurrentItem = $newRow
    $grid.ScrollIntoView($newRow)
    $grid.Focus()
    $grid.UpdateLayout()

    Write-Status "Command duplicated"
}

# Handle the Main Window Remove Button click event to remove one or multiple RowData objects from the collection
function Remove-CommandRow {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    $allGrid = $tabs["All"].Content
    $allData = $allGrid.ItemsSource
    $grid = $tabControl.SelectedItem.Content

    # We want to make a copy of the selected items to avoid issues
    # with the collection being modified while still enumerating
    $selectedItems = @()
    foreach ($item in $grid.SelectedItems) {
        $selectedItems += $item
    }

    # Show confirmation dialog if there are items to delete
    if ($selectedItems.Count -gt 0) {
        $itemText = if ($selectedItems.Count -eq 1) { "command" } else { "commands" }
        $message = "Are you sure you want to delete the selected $($selectedItems.Count) $($itemText)?"
        $result = [System.Windows.MessageBox]::Show($message, "Confirm Delete", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)

        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }
    }

    # Create a snapshot of the deleted items for the recycle bin
    $deletedBatch = @{
        Timestamp = Get-Date
        Items = @()
    }

    foreach ($item in $selectedItems) {
        $id = $item.Id

        # Create a deep copy of the item for the recycle bin
        $itemCopy = New-Object RowData
        $itemCopy.Id = $item.Id
        $itemCopy.Name = $item.Name
        $itemCopy.Description = $item.Description
        $itemCopy.Category = $item.Category
        $itemCopy.Command = $item.Command
        $itemCopy.SkipParameterSelect = $item.SkipParameterSelect
        $itemCopy.PreCommand = $item.PreCommand

        $deletedBatch.Items += $itemCopy

        # If item has a category then remove from category's tab and remove the tab
        # if it was the only item of that category
        $category = $item.Category
        if ($category) {
            $categoryGrid = $tabs[$category].Content
            $categoryData = $categoryGrid.ItemsSource
            $categoryIndex = Get-GridIndexOfId -Grid $categoryGrid -Id $id
            $categoryData.RemoveAt($categoryIndex)
            if ($categoryData.Count -eq 0) {
                $tabControl.Items.Remove($tabs[$category])
                $tabs.Remove($category)
            }
        }
        $allIndex = Get-GridIndexOfId -Grid $allGrid -Id $Id
        $allData.RemoveAt($allIndex)
    }

    if ($selectedItems.Count -gt 0) {
        # Add deleted items to recycle bin
        $script:State.RecycleBin.Enqueue($deletedBatch)

        # Maintain max size by removing oldest items
        while ($script:State.RecycleBin.Count -gt $script:State.RecycleBinMaxSize) {
            [void]$script:State.RecycleBin.Dequeue()
        }

        Set-UnsavedChanges $true
        $itemText = if ($selectedItems.Count -eq 1) { "command" } else { "commands" }
        Write-Status "Deleted $($selectedItems.Count) $($itemText) (can be restored with Undo Delete)"
    }
}

# Restore the last deleted command(s) from the recycle bin
function Restore-DeletedCommand {
    param (
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

    if ($script:State.RecycleBin.Count -eq 0) {
        Write-Status "No deleted commands to restore"
        [System.Windows.MessageBox]::Show("The recycle bin is empty. There are no deleted commands to restore.", "Recycle Bin Empty", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    # Get the most recent deletion batch
    $deletedBatch = $script:State.RecycleBin.Dequeue()
    $restoredCount = 0

    $allTab = $tabs["All"]
    $allGrid = $allTab.Content
    $allData = $allGrid.ItemsSource

    foreach ($item in $deletedBatch.Items) {
        # Check if an item with this ID already exists (to prevent duplicates)
        $existingIndex = Get-GridIndexOfId -Grid $allGrid -Id $item.Id
        if ($existingIndex -ge 0) {
            Write-Log "Skipping restore of command ID $($item.Id) - already exists"
            continue
        }

        # Add to All tab
        $allData.Add($item)

        # If the item has a category, add to category tab (create if needed)
        if ($item.Category) {
            $categoryTab = $tabs[$item.Category]
            if (-not $categoryTab) {
                # Create new category tab
                $itemsSource = New-Object System.Collections.ObjectModel.ObservableCollection[RowData]
                $categoryTab = New-DataTab -Name $item.Category -ItemsSource $itemsSource -TabControl $tabControl
                $tabs.Add($item.Category, $categoryTab)

                # Assign event handlers to the new tab
                $categoryTab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
                $categoryTab.Content.Add_PreviewKeyDown({
                    param($sender,$e)
                    if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                        Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
                    }
                    elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem -and -not $sender.IsInEditMode) {
                        $e.Handled = $true
                        Invoke-MainRunClick -TabControl $script:UI.TabControl
                    }
                })
                Sort-TabControl -TabControl $tabControl
            }
            $categoryTab.Content.ItemsSource.Add($item)
        }

        $restoredCount++
    }

    if ($restoredCount -gt 0) {
        Set-UnsavedChanges $true
        Update-FavoriteHighlighting
        $itemText = if ($restoredCount -eq 1) { "command" } else { "commands" }
        Write-Status "Restored $restoredCount $($itemText)"
    }
}

# Handle the Main Edit Button click event to enable or disable editing of the grids
function Toggle-EditMode {
    param (
        [hashtable]$tabs
    )

    Set-TabsReadOnlyStatus -Tabs $tabs
    Set-TabsExtraColumnsVisibility -Tabs $tabs
}

function Toggle-CommandFavorite {
    $selectedTab = $script:UI.TabControl.SelectedItem
    $grid = $selectedTab.Content
    $selectedItem = $grid.SelectedItem

    if ($selectedItem) {
        try {
            $favorites = $script:UI.Tabs["Favorites"].Content.ItemsSource
            $existingFavorite = $favorites | Where-Object { $_.Id -eq $selectedItem.Id }

            if ($existingFavorite) {
                [void]$favorites.Remove($existingFavorite)
                Save-Favorites -Favorites $favorites
                Update-FavoriteHighlighting
                Write-Status "Removed from favorites"
            }
            else {
                $script:State.FavoritesHighestOrder++
                $favoriteRow = [FavoriteRowData]::new($selectedItem, $script:State.FavoritesHighestOrder)
                [void]$favorites.Add($favoriteRow)
                Save-Favorites -Favorites $favorites
                Update-FavoriteHighlighting
                Write-Status "Added to favorites"
            }
        }
        catch {
            Write-Status "Failed to add/remove favorite"
            Write-Log "Failed to add/remove favorite: $_"
        }
    }
}

# Handle the Cell Edit ending event to make sure all tabs are updated properly for cell changes
function Invoke-CellEditEndingHandler {
    param (
        $sender,
        $e,
        [System.Windows.Controls.TabControl]$tabControl,
        [hashtable]$tabs
    )

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
    $allIndex = Get-GridIndexOfId -Grid $allGrid -Id $id

    # Sync only the changed property between tabs
    if (-not $newObject) {
        $categoryGrid = $tabs[$category].Content
        $categoryData = $categoryGrid.ItemsSource
        $categoryIndex = Get-GridIndexOfId -Grid $categoryGrid -Id $id
        
        # Only update the specific property that was edited, preserve other values
        if ($categoryIndex -ge 0) {
            $propertyName = $e.Column.Header
            $propertyValue = $editedObject.GetType().GetProperty($propertyName).GetValue($editedObject)
            $categoryData[$categoryIndex].GetType().GetProperty($propertyName).SetValue($categoryData[$categoryIndex], $propertyValue)
        }
    }

    # Update the specific edited property in the All tab
    $propertyName = $e.Column.Header
    $propertyValue = $editedObject.GetType().GetProperty($propertyName).GetValue($editedObject)
    $allData[$allIndex].GetType().GetProperty($propertyName).SetValue($allData[$allIndex], $propertyValue)

    # Mark as having unsaved changes
    Set-UnsavedChanges $true

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
            $newTab = New-DataTab -Name $newCategory -ItemsSource $itemsSource -TabControl $tabControl
            $tabs.Add($newCategory, $newTab)

            # Assign the CellEditEnding event to the new tab
            $newTab.Content.Add_CellEditEnding({ param($sender,$e) Invoke-CellEditEndingHandler -Sender $sender -E $e -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs })
            $newTab.Content.Add_PreviewKeyDown({
                param($sender,$e)
                if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
                    Remove-CommandRow -TabControl $script:UI.TabControl -Tabs $script:UI.Tabs
                }
                elseif ($e.Key -eq [System.Windows.Input.Key]::Enter -and $sender.SelectedItem -and -not $sender.IsInEditMode) {
                    $e.Handled = $true
                    Invoke-MainRunClick -TabControl $script:UI.TabControl
                }
            })
        }
        $newTab.Content.ItemsSource.Add($editedObject)
        Sort-TabControl -TabControl $tabControl
    }
}