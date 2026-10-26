# Dot-source from BeforeAll to load PSGUI's functions and classes without starting the UI
Add-Type -AssemblyName PresentationCore, PresentationFramework

$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:SrcRoot = Join-Path $RepoRoot 'src'

. (Join-Path $SrcRoot 'Functions/Classes.ps1')
foreach ($name in 'Utilities', 'DataFileManagement', 'CommandWindow', 'CommandHistory', 'GridManagement', 'SettingsManagement', 'FavoritesManagement') {
    . (Join-Path $SrcRoot "Functions/$name.ps1")
}

function Reset-PSGUIState {
    $script:State = @{
        CurrentCommandListId = $null
        HasUnsavedChanges = $false
    }
    $script:Settings = @{
        UseProfile = $false
        ProfilePath = $null
    }
}
