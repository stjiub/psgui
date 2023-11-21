properties {
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
    $script:CurrentVersion = (Get-Item "$ProjectRoot\PSGUI.exe").Versioninfo.FileVersionRaw

    $RequiredFiles = @("PSGUI.exe", "MainWindow.xaml", "icon.ico", "Assembly")
    $InstallLocations = @("C:\Program Files\PSGUI")
}

task default -depends Build

    task Build {
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