BeforeAll {
    . $PSScriptRoot/TestSetup.ps1
}

Describe 'Utilities' {
    BeforeEach { Reset-PSGUIState }

    Context 'Get-HighestId' {
        It 'returns 0 for an empty collection' {
            Get-HighestId -json @() | Should -Be 0
        }

        It 'returns the largest Id' {
            Get-HighestId -json @([pscustomobject]@{ Id = 3 }, [pscustomobject]@{ Id = 9 }, [pscustomobject]@{ Id = 4 }) | Should -Be 9
        }
    }

    Context 'ID generators' {
        It 'Get-UniqueCommandId returns distinct GUIDs' {
            $a = Get-UniqueCommandId
            $b = Get-UniqueCommandId
            [guid]::Parse($a) | Should -Not -BeNullOrEmpty
            $a | Should -Not -Be $b
        }

        It 'Get-CommandListId returns a GUID' {
            { [guid]::Parse((Get-CommandListId)) } | Should -Not -Throw
        }
    }

    Context 'Get-PowerShellArguments' {
        It 'wraps the command in a script block' {
            Get-PowerShellArguments -Command 'Get-Date' | Should -Be '-ExecutionPolicy Bypass -Command "& { Get-Date }"'
        }

        It 'adds -NoExit before the command' {
            Get-PowerShellArguments -Command 'Get-Date' -NoExit | Should -Be '-ExecutionPolicy Bypass -NoExit -Command "& { Get-Date }"'
        }

        It 'doubles embedded double quotes' {
            Get-PowerShellArguments -Command 'Write-Output "hi"' | Should -BeLike '*Write-Output ""hi""*'
        }

        It 'omits -Command when there is nothing to run' {
            Get-PowerShellArguments -NoExit | Should -Be '-ExecutionPolicy Bypass -NoExit'
        }

        It 'dot-sources the profile first when enabled and present' {
            $profilePath = Join-Path $TestDrive 'profile.ps1'
            Set-Content $profilePath '# profile'
            $script:Settings.UseProfile = $true
            $script:Settings.ProfilePath = $profilePath
            Get-PowerShellArguments -Command 'Get-Date' | Should -BeLike "*. '$profilePath'; Get-Date*"
        }

        It 'escapes single quotes in the profile path' {
            $dir = Join-Path $TestDrive "O'Brien"
            New-Item -ItemType Directory $dir | Out-Null
            $profilePath = Join-Path $dir 'profile.ps1'
            Set-Content $profilePath '# profile'
            $script:Settings.UseProfile = $true
            $script:Settings.ProfilePath = $profilePath
            Get-PowerShellArguments -Command 'x' | Should -BeLike "*O''Brien*"
        }

        It 'skips a profile that does not exist' {
            $script:Settings.UseProfile = $true
            $script:Settings.ProfilePath = Join-Path $TestDrive 'missing.ps1'
            Get-PowerShellArguments -Command 'Get-Date' | Should -Not -BeLike '*missing.ps1*'
        }

        It 'ignores the profile when disabled' {
            $profilePath = Join-Path $TestDrive 'profile2.ps1'
            Set-Content $profilePath '# profile'
            $script:Settings.UseProfile = $false
            $script:Settings.ProfilePath = $profilePath
            Get-PowerShellArguments -Command 'Get-Date' | Should -Not -BeLike '*profile2*'
        }
    }
}

Describe 'Classes' {
    It 'FavoriteRowData copies every RowData field and sets the order' {
        $row = [RowData]@{
            Id = 'a'; Name = 'n'; Description = 'd'; Category = 'c'; Command = 'cmd'; SkipParameterSelect = $true
            PreCommand = 'pre'; PostCommand = 'post'; Transcript = $true; PSTask = $true; PSTaskMode = 'Text'
            PSTaskVisibilityLevel = 'Normal'; ShellOverride = 'pwsh'; LogParameterNames = 'Name'
        }
        $fav = [FavoriteRowData]::new($row, 7)
        foreach ($prop in [RowData].GetProperties().Name) {
            $fav.$prop | Should -Be $row.$prop -Because "property $prop should be copied"
        }
        $fav.Order | Should -Be 7
    }

    It 'Get-GridPropertyType uses FavoriteRowData only for the favorites tab' {
        Get-GridPropertyType -name '*' | Should -Be ([FavoriteRowData])
        Get-GridPropertyType -name 'Network' | Should -Be ([RowData])
    }
}

Describe 'Data files' {
    BeforeEach {
        Reset-PSGUIState
        Mock Write-Log { }
        Mock Write-Status { }
        Mock Set-UnsavedChanges { }
        Mock Show-ErrorMessageBox { }
    }

    Context 'Initialize-DataFile' {
        It 'creates a file with a CommandListId and no commands' {
            $path = Join-Path $TestDrive 'new/data.json'
            Initialize-DataFile $path
            $json = Get-Content $path -Raw | ConvertFrom-Json
            [guid]::Parse($json.CommandListId) | Should -Not -BeNullOrEmpty
            @($json.Commands) | Should -BeNullOrEmpty
        }

        It 'does not overwrite an existing file' {
            $path = Join-Path $TestDrive 'existing.json'
            Set-Content $path 'keep me'
            Initialize-DataFile $path
            Get-Content $path -Raw | Should -Match 'keep me'
        }
    }

    Context 'Load-DataFile' {
        It 'loads the current format and keeps its CommandListId' {
            $path = Join-Path $TestDrive 'current.json'
            @{ CommandListId = 'list-1'; Commands = @(@{ Id = 'abc'; Name = 'One'; Category = 'X'; Command = 'Get-Date'; Transcript = $true }) } | ConvertTo-Json -Depth 3 | Set-Content $path
            $rows = Load-DataFile $path
            $rows.Count | Should -Be 1
            $rows[0].Name | Should -Be 'One'
            $rows[0].Id | Should -Be 'abc'
            $rows[0].Transcript | Should -BeTrue
            $script:State.CurrentCommandListId | Should -Be 'list-1'
        }

        It 'loads the legacy array format and assigns a CommandListId' {
            $path = Join-Path $TestDrive 'legacy.json'
            ConvertTo-Json -InputObject @(@{ Id = 'abc'; Name = 'Old'; Command = 'Get-Date' }) | Set-Content $path
            $rows = Load-DataFile $path
            $rows[0].Name | Should -Be 'Old'
            $script:State.CurrentCommandListId | Should -Not -BeNullOrEmpty
        }

        It 'generates a CommandListId when the file has none' {
            $path = Join-Path $TestDrive 'noid.json'
            @{ Commands = @(@{ Id = 'abc'; Name = 'X' }) } | ConvertTo-Json -Depth 3 | Set-Content $path
            Load-DataFile $path | Out-Null
            $script:State.CurrentCommandListId | Should -Not -BeNullOrEmpty
        }

        It 'migrates integer Ids to GUIDs' {
            $path = Join-Path $TestDrive 'intid.json'
            ConvertTo-Json -InputObject @(@{ Id = 5; Name = 'N' }) | Set-Content $path
            $rows = Load-DataFile $path
            [guid]::Parse($rows[0].Id) | Should -Not -BeNullOrEmpty
        }

        It 'migrates the old Log string to the Transcript and PSTask flags' {
            $path = Join-Path $TestDrive 'logstring.json'
            ConvertTo-Json -InputObject @(@{ Id = 'a'; Name = 'T'; Log = 'Transcript' }, @{ Id = 'b'; Name = 'P'; Log = 'PSTask' }) | Set-Content $path
            $rows = Load-DataFile $path
            $rows[0].Transcript | Should -BeTrue
            $rows[0].PSTask | Should -BeFalse
            $rows[1].PSTask | Should -BeTrue
            $rows[1].Transcript | Should -BeFalse
        }

        It 'migrates the oldest boolean Log to Transcript' {
            $path = Join-Path $TestDrive 'logbool.json'
            ConvertTo-Json -InputObject @(@{ Id = 'a'; Name = 'T'; Log = $true }) | Set-Content $path
            (Load-DataFile $path)[0].Transcript | Should -BeTrue
        }

        It 'returns an empty collection for an empty file' {
            $path = Join-Path $TestDrive 'empty.json'
            Set-Content $path ''
            (Load-DataFile $path).Count | Should -Be 0
            $script:State.CurrentCommandListId | Should -Not -BeNullOrEmpty
        }

        It 'returns an empty collection for invalid JSON instead of throwing' {
            $path = Join-Path $TestDrive 'bad.json'
            Set-Content $path '{ not json'
            $rows = Load-DataFile $path 2>$null
            $rows.Count | Should -Be 0
        }
    }

    Context 'Save-DataFile' {
        It 'round-trips rows, including a single row and non-ASCII text' {
            $path = Join-Path $TestDrive 'roundtrip.json'
            $script:State.CurrentCommandListId = 'list-rt'
            $row = [RowData]@{ Id = 'id1'; Name = 'Café résumé'; Category = 'Docs'; Command = 'Get-Date'; PSTask = $true; PSTaskMode = 'Text' }
            $data = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()
            $data.Add($row)
            Save-DataFile -filePath $path -data $data

            (Get-Content $path -Raw | ConvertFrom-Json).Commands | Should -HaveCount 1
            $loaded = Load-DataFile $path
            $loaded[0].Name | Should -Be 'Café résumé'
            $loaded[0].PSTask | Should -BeTrue
            $loaded[0].PSTaskMode | Should -Be 'Text'
            $script:State.CurrentCommandListId | Should -Be 'list-rt'
        }

        It 'skips unpopulated rows' {
            $path = Join-Path $TestDrive 'skip.json'
            $script:State.CurrentCommandListId = 'x'
            $data = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()
            $data.Add([RowData]@{ Id = '1'; Name = 'Keep' })
            $data.Add([RowData]::new())
            Save-DataFile -filePath $path -data $data
            @((Get-Content $path -Raw | ConvertFrom-Json).Commands) | Should -HaveCount 1
        }

        It 'marks the data as saved' {
            $script:State.CurrentCommandListId = 'x'
            $data = [System.Collections.ObjectModel.ObservableCollection[Object]]::new()
            Save-DataFile -filePath (Join-Path $TestDrive 'saved.json') -data $data
            Should -Invoke Set-UnsavedChanges -ParameterFilter { $args[0] -eq $false -or $PesterBoundParameters.Values -contains $false }
        }
    }
}

Describe 'Command history helpers' {
    It 'Get-ParameterSummaryFromCommand strips the root command, pre and post commands' {
        $cmd = [Command]@{ Root = 'Get-Service'; CleanCommand = 'Prep; Get-Service -Name "bits"; Cleanup'; PreCommand = 'Prep'; PostCommand = 'Cleanup' }
        Get-ParameterSummaryFromCommand -Command $cmd | Should -Be '-Name "bits"'
    }

    It 'reports no parameters when only the root command remains' {
        Get-ParameterSummaryFromCommand -Command ([Command]@{ Root = 'Get-Date'; CleanCommand = 'Get-Date' }) | Should -Be '(No parameters)'
    }

    It 'reports no parameters for an empty command' {
        Get-ParameterSummaryFromCommand -Command ([Command]::new()) | Should -Be '(No parameters)'
    }

    It 'truncates long parameter lists to 100 characters' {
        $long = '-Name "' + ('x' * 200) + '"'
        $summary = Get-ParameterSummaryFromCommand -Command ([Command]@{ Root = 'Get-Date'; CleanCommand = "Get-Date $long" })
        $summary.Length | Should -Be 100
        $summary | Should -BeLike '*...'
    }
}

Describe 'Log file naming' {
    It 'returns nothing when no parameter names are configured' {
        Get-LogParameterSuffix -Command ([Command]@{ LogParameterNames = '' }) | Should -Be ''
    }

    It 'returns nothing until the parameter grid exists' {
        Get-LogParameterSuffix -Command ([Command]@{ LogParameterNames = 'Name' }) | Should -Be ''
    }

    It 'builds a sanitized suffix from control values' {
        $grid = [System.Windows.Controls.Grid]::new()
        $name = [System.Windows.Controls.TextBox]@{ Name = 'Name'; Text = 'a/b:c' }
        $flag = [System.Windows.Controls.CheckBox]@{ Name = 'Force'; IsChecked = $true }
        [void]$grid.Children.Add($name)
        [void]$grid.Children.Add($flag)
        $params = foreach ($n in 'Name', 'Force') {
            $ast = [System.Management.Automation.Language.Parser]::ParseInput("param(`$$n)", [ref]$null, [ref]$null)
            $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true)
        }
        $cmd = [Command]@{ LogParameterNames = 'Name, Force'; Parameters = @($params) }
        Get-LogParameterSuffix -Command $cmd -CommandGrid $grid | Should -Be '-a_b_c_True'
    }
}

Describe 'Command introspection' {
    BeforeAll {
        function Test-PSGUISample {
            param(
                [Parameter(Mandatory)] [string]$Name,
                [ValidateSet('Red', 'Green')] [string]$Color,
                [switch]$Force
            )
        }
    }

    It 'Get-CommandType returns the command type' {
        Get-CommandType 'Test-PSGUISample' | Should -Be 'Function'
        Get-CommandType 'Get-Date' | Should -Be 'Cmdlet'
    }

    It 'Get-CommandType throws for unknown commands' {
        { Get-CommandType 'Definitely-NotACommand' } | Should -Throw '*not found*'
    }

    It 'Get-ScriptBlockParameters lists a function''s parameters' {
        (Get-ScriptBlockParameters 'Test-PSGUISample').Name.VariablePath.UserPath | Should -Be 'Name', 'Color', 'Force'
    }

    It 'Get-ScriptBlockParameters rejects compiled cmdlets' {
        { Get-ScriptBlockParameters 'Get-Date' } | Should -Throw
    }

    It 'Test-AttributeType detects switches and ValidateSet' {
        $params = Get-ScriptBlockParameters 'Test-PSGUISample'
        Test-AttributeType -parameter $params[2] -typeName 'switch' | Should -Not -BeNullOrEmpty
        Test-AttributeType -parameter $params[1] -typeName 'ValidateSet' | Should -Not -BeNullOrEmpty
        Test-AttributeType -parameter $params[0] -typeName 'switch' | Should -BeNullOrEmpty
    }

    It 'Get-ValidateSetValues returns a leading blank then the allowed values' {
        $params = Get-ScriptBlockParameters 'Test-PSGUISample'
        @(Get-ValidateSetValues -parameter $params[1]) | Should -Be @('', 'Red', 'Green')
    }
}
