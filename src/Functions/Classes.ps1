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
    [bool]$Transcript
    [bool]$PSTask
    [string]$PSTaskMode
    [string]$PSTaskVisibilityLevel
    [string]$ShellOverride
    [string]$LogParameterNames
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
        $this.Transcript = $rowData.Transcript
        $this.PSTask = $rowData.PSTask
        $this.PSTaskMode = $rowData.PSTaskMode
        $this.PSTaskVisibilityLevel = $rowData.PSTaskVisibilityLevel
        $this.ShellOverride = $rowData.ShellOverride
        $this.LogParameterNames = $rowData.LogParameterNames
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
    [bool]$Transcript
    [bool]$PSTask
    [string]$PSTaskMode
    [string]$PSTaskVisibilityLevel
    [string]$LogPath
    [string]$ShellOverride
    [string]$LogParameterNames
}