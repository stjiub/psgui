BeforeAll {
    . $PSScriptRoot/TestSetup.ps1

    function New-TestRows {
        $rows = [System.Collections.ObjectModel.ObservableCollection[RowData]]::new()
        foreach ($id in 'a', 'b', 'c') { $rows.Add([RowData]@{ Id = $id; Name = "Cmd $id" }) }
        , $rows
    }
}

Describe 'Settings' {
    BeforeEach {
        Reset-PSGUIState
        Mock Write-Log { }
        Mock Write-Status { }
        $script:ApplicationPaths = @{ DefaultSettingsFile = Join-Path $RepoRoot 'assets/defaultsettings.json' }
        $script:Settings.SettingsPath = Join-Path $TestDrive "settings-$([guid]::NewGuid()).json"
    }

    It 'Load-DefaultSettings expands environment variables in paths' {
        $defaults = Load-DefaultSettings
        $defaults.DefaultDataFile | Should -Be (Join-Path $env:APPDATA 'PSGUI\data.json')
        $defaults.DefaultShell | Should -Not -BeNullOrEmpty
    }

    It 'Load-DefaultSettings returns null when the defaults file is missing' {
        $script:ApplicationPaths.DefaultSettingsFile = Join-Path $TestDrive 'nope.json'
        Load-DefaultSettings | Should -BeNullOrEmpty
    }

    It 'Ensure-SettingsFileExists writes a settings file from the current values' {
        Load-DefaultSettings | ForEach-Object { foreach ($k in $_.Keys) { $script:Settings[$k] = $_[$k] } }
        $script:Settings.SettingsPath = Join-Path $TestDrive 'sub/settings.json'
        Ensure-SettingsFileExists
        (Get-Content $script:Settings.SettingsPath -Raw | ConvertFrom-Json).DefaultShell | Should -Be $script:Settings.DefaultShell
    }

    It 'Load-Settings falls back to defaults for settings missing from the file' {
        Set-Content $script:Settings.SettingsPath '{ "DefaultShell": "pwsh" }'
        Load-Settings
        $script:Settings.DefaultShell | Should -Be 'pwsh'
        $script:Settings.CommandHistoryLimit | Should -Be 50
        $script:Settings.SaveHistory | Should -BeTrue
    }

    It 'Save-Settings and Load-Settings round-trip' {
        Load-DefaultSettings | ForEach-Object { foreach ($k in $_.Keys) { $script:Settings[$k] = $_[$k] } }
        $script:Settings.SettingsPath = Join-Path $TestDrive 'rt/settings.json'
        $script:Settings.DefaultShell = 'pwsh'
        $script:Settings.StatusTimeout = 12
        $script:Settings.UseProfile = $true
        Save-Settings
        $script:Settings.DefaultShell = 'changed'
        $script:Settings.StatusTimeout = 1
        Load-Settings
        $script:Settings.DefaultShell | Should -Be 'pwsh'
        $script:Settings.StatusTimeout | Should -Be 12
        $script:Settings.UseProfile | Should -BeTrue
    }

    It 'saves every setting that has a default' {
        Load-DefaultSettings | ForEach-Object { foreach ($k in $_.Keys) { $script:Settings[$k] = $_[$k] } }
        $script:Settings.SettingsPath = Join-Path $TestDrive 'all/settings.json'
        Save-Settings
        $saved = (Get-Content $script:Settings.SettingsPath -Raw | ConvertFrom-Json).PSObject.Properties.Name
        $defaults = (Get-Content $script:ApplicationPaths.DefaultSettingsFile -Raw | ConvertFrom-Json).PSObject.Properties.Name
        $defaults | Where-Object { $_ -notin $saved } | Should -BeNullOrEmpty
    }
}

Describe 'Favorites' {
    BeforeEach {
        Reset-PSGUIState
        Mock Write-Log { }
        Mock Write-Status { }
        $script:Settings.FavoritesPath = Join-Path $TestDrive "favs-$([guid]::NewGuid())/favorites.json"
        $script:State.CurrentCommandListId = 'list-1'
        $script:State.FavoritesHighestOrder = 0
        $script:rows = New-TestRows
    }

    It 'returns nothing when there is no favorites file' {
        @(Load-Favorites -AllData $rows) | Should -HaveCount 0
    }

    It 'round-trips favorites in order' {
        $favs = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()
        $favs.Add([FavoriteRowData]::new($rows[2], 1))
        $favs.Add([FavoriteRowData]::new($rows[0], 2))
        Save-Favorites -Favorites $favs
        $loaded = @(Load-Favorites -AllData $rows)
        $loaded.Id | Should -Be 'c', 'a'
        $loaded[0].Order | Should -Be 1
        $script:State.FavoritesHighestOrder | Should -Be 2
    }

    It 'keeps favorites for other command lists when saving' {
        $first = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()
        $first.Add([FavoriteRowData]::new($rows[0], 1))
        Save-Favorites -Favorites $first

        $script:State.CurrentCommandListId = 'list-2'
        $second = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()
        $second.Add([FavoriteRowData]::new($rows[1], 1))
        Save-Favorites -Favorites $second

        $script:State.CurrentCommandListId = 'list-1'
        @(Load-Favorites -AllData $rows).Id | Should -Be 'a'
        $script:State.CurrentCommandListId = 'list-2'
        @(Load-Favorites -AllData $rows).Id | Should -Be 'b'
    }

    It 'drops favorites whose command no longer exists' {
        $favs = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()
        $favs.Add([FavoriteRowData]::new([RowData]@{ Id = 'gone'; Name = 'Removed' }, 1))
        $favs.Add([FavoriteRowData]::new($rows[1], 2))
        Save-Favorites -Favorites $favs
        @(Load-Favorites -AllData $rows).Id | Should -Be 'b'
    }
}
