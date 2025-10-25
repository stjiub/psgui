# Define the RowData object. This is the object that is used on all the Main window tabitem grids
class RowData {
    [string]$Id
    [string]$Name
    [string]$Description
    [string]$Category
    [string]$Command
    [bool]$SkipParameterSelect
    [string]$PreCommand
    [bool]$Log
}

class FavoriteRowData : RowData {
    [int]$Order

    FavoriteRowData([RowData]$rowData, [int]$order) {
        $this.Id = $rowData.Id
        $this.Name = $rowData.Name
        $this.Description = $rowData.Description
        $this.Category = $rowData.Category
        $this.Command = $rowData.Command
        $this.SkipParameterSelect = $rowData.SkipParameterSelect
        $this.PreCommand = $rowData.PreCommand
        $this.Log = $rowData.Log
        $this.Order = $order
    }
}

# Define the Command object. This is used by the CommandWindow to construct the grid and run the command
class Command {
    [string]$Root
    [string]$Full
    [string]$CleanCommand
    [string]$PreCommand
    [System.Object[]]$Parameters
    [bool]$SkipParameterSelect
    [bool]$Log
    [string]$LogPath
}