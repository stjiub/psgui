# Define the RowData object. This is the object that is used on all the Main window tabitem grids
class RowData {
    [string]$Id
    [string]$Name
    [string]$Description
    [string]$Category
    [string]$Command
    [bool]$SkipParameterSelect
    [string]$PreCommand
    [string]$PostCommand
    [bool]$Log
    [string]$ShellOverride
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
        $this.PostCommand = $rowData.PostCommand
        $this.Log = $rowData.Log
        $this.ShellOverride = $rowData.ShellOverride
        $this.Order = $order
    }
}

# Define the Command object. This is used by the CommandWindow to construct the grid and run the command
class Command {
    [string]$Root
    [string]$Full
    [string]$CleanCommand
    [string]$PreCommand
    [string]$PostCommand
    [System.Object[]]$Parameters
    [bool]$SkipParameterSelect
    [bool]$Log
    [string]$LogPath
    [string]$ShellOverride
}