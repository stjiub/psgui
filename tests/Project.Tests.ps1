BeforeDiscovery {
    $root = Split-Path -Parent $PSScriptRoot
    $script:sourceFiles = Get-ChildItem (Join-Path $root 'src') -Filter '*.ps1' -Recurse | ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } }
    $script:xamlFiles = Get-ChildItem (Join-Path $root 'src') -Filter '*.xaml' | ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } }
}

BeforeAll {
    $script:root = Split-Path -Parent $PSScriptRoot
    $script:allCode = (Get-ChildItem (Join-Path $root 'src') -Filter '*.ps1' -Recurse | Get-Content -Raw) -join "`n"
}

Describe 'Source files' {
    It '<Name> parses without syntax errors' -ForEach $sourceFiles {
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It '<Name> is well-formed XAML' -ForEach $xamlFiles {
        { [xml](Get-Content $Path -Raw) } | Should -Not -Throw
    }

    It 'embeds a compilable Win32 helper' {
        { Add-Type -TypeDefinition (Get-Content (Join-Path $root 'src/Win32API.cs') -Raw) -ErrorAction Stop } | Should -Not -Throw
    }
}

Describe 'Default settings' {
    BeforeAll {
        $script:defaults = Get-Content (Join-Path $root 'assets/defaultsettings.json') -Raw | ConvertFrom-Json
    }

    It 'is valid JSON' {
        $defaults | Should -Not -BeNullOrEmpty
    }

    It 'defines every setting the code reads' {
        $used = [regex]::Matches($allCode, '\$script:Settings\.(\w+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
        $missing = $used | Where-Object { $_ -notin $defaults.PSObject.Properties.Name }
        $missing | Should -BeNullOrEmpty
    }

    It 'does not define settings nothing reads' {
        $used = [regex]::Matches($allCode, '\$script:Settings\.(\w+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
        $unused = $defaults.PSObject.Properties.Name | Where-Object { $_ -notin $used }
        $unused | Should -BeNullOrEmpty
    }

    It 'keeps user data under %APPDATA%' {
        foreach ($key in 'SettingsPath', 'FavoritesPath', 'DefaultHistoryPath', 'DefaultDataFile', 'DefaultLogsPath') {
            $defaults.$key | Should -BeLike '%APPDATA%*'
        }
    }
}

Describe 'UI wiring' {
    It 'every $script:UI.<name> used in code exists in the XAML (or is set in code)' {
        $xamlNames = foreach ($file in Get-ChildItem (Join-Path $root 'src') -Filter '*.xaml') {
            [regex]::Matches((Get-Content $file.FullName -Raw), '(?<![\w:.])(?:x:)?Name="(\w+)"') | ForEach-Object { $_.Groups[1].Value }
        }
        $setInCode = 'Window', 'Tabs'
        $used = [regex]::Matches($allCode, '\$script:UI\.(\w+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
        $missing = $used | Where-Object { $_ -notin $xamlNames -and $_ -notin $setInCode }
        $missing | Should -BeNullOrEmpty
    }
}
