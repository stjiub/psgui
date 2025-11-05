properties {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
    if (Test-Path "$ProjectRoot\PSGUI.exe") {
        $script:CurrentVersion = (Get-Item "$ProjectRoot\PSGUI.exe").Versioninfo.FileVersionRaw
    }

    $RequiredFiles = @("PSGUI.exe", "icon.ico", "Assembly", "defaultsettings.json")
    $InstallLocations = @("C:\Program Files\PSGUI")
}

task default -depends Combine

    task Combine {
        # Start with Functions
        Get-Content "$($ProjectRoot)\Functions\*.ps1" | Set-Content -Path "$($ProjectRoot)\PSGUI.ps1"

        # Embed MainWindow.xaml as a here-string
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value "`n# Embedded MainWindow.xaml"
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value '$script:MainWindowXaml = @"'
        Get-Content "$($ProjectRoot)\MainWindow.xaml" -Raw | Add-Content -Path "$($ProjectRoot)\PSGUI.ps1"
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value '"@'

        # Embed CommandWindow.xaml as a here-string
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value "`n# Embedded CommandWindow.xaml"
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value '$script:CommandWindowXaml = @"'
        Get-Content "$($ProjectRoot)\CommandWindow.xaml" -Raw | Add-Content -Path "$($ProjectRoot)\PSGUI.ps1"
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value '"@'

        # Embed Win32API.cs as a here-string
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value "`n# Embedded Win32API.cs"
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value '$script:Win32API = @"'
        Get-Content "$($ProjectRoot)\Win32API.cs" -Raw | Add-Content -Path "$($ProjectRoot)\PSGUI.ps1"
        Add-Content -Path "$($ProjectRoot)\PSGUI.ps1" -Value '"@'

        # Add Main.ps1
        Get-Content "$($ProjectRoot)\Main.ps1" | Add-Content -Path "$($ProjectRoot)\PSGUI.ps1"
    }

    task Build -depends Combine {
        $module = Get-Module -ListAvailable PS2Exe

        if (-not $module) {
            Install-Module PS2Exe
        }

        $script:NewVersion = Read-Host "Current Version: $script:CurrentVersion`nNew Version"

        $iconFile = Get-ChildItem $projectRoot -Filter "*.ico"

        ps2exe $projectRoot\PSGUI.ps1 -requireAdmin -title "PSGUI" -noConsole -version $script:NewVersion -iconFile $iconFile.FullName
    }

    task Commit -depends Build {
        # Publish the new version back to remote git repo
        try {
            $env:Path += ";$env:ProgramFiles\Git\cmd"
            Import-Module posh-git -ErrorAction Stop
            git checkout main
            git add --all
            git status

            git commit -s -m "$($script:NewVersion.ToString())"
            git push origin main
            Write-Host "Project pushed to remote repo." -ForegroundColor Cyan
        }
        catch {
            Write-Warning "Pushing commit to remote repo failed."
            Write-Error $_
        }
    }

    task Publish -depends Build, Commit {
        foreach ($location in $InstallLocations) {
            if (-not (Test-Path $location)) {
                New-Item -ItemType "Directory" -Path $location
                Write-Host "Created new directory at: $location"
            }
            try {
                foreach ($file in $RequiredFiles) {
                    Copy-Item -Path "$projectRoot\$file" -Destination $location -Recurse -Force -ErrorAction Stop
                    Write-Host "Copied file $file to $location"
                }
            }
            catch {
                Write-Error $_
            }
        }
    }