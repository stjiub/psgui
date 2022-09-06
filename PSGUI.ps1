function InitializeConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$File
    )
    if (-not (Test-Path $File)) {
        try {
            New-Item -Path $File -ItemType "File" | Out-Null
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to create user_config file at path: $File", "Error", "Ok", "Error")
            exit(1)
        }
    }
}

function LoadConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$File
    )

    try {
        [string]$contentRaw = (Get-Content $File -Raw -ErrorAction Stop)
        if ($contentRaw) {
            [array]$contentJson = $contentRaw | ConvertFrom-Json
            return $contentJson
        }
        else {
            Write-Verbose "Config file $file is empty."
            return
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Failed to load configuration from: $File", "Error", "Ok", "Error")
        return
    }
}

function SaveConfig {
    param(
        [Parameter(Mandatory = $true)]
        [string]$File,

        [Parameter(Mandatory = $true)]
        [array]$Content
    )

    try {
        $json = ConvertTo-Json $Content
        Set-Content -Path $File -Value $json
    }
    catch {
        Write-Error "Failed to save configuration to: $File"
        [System.Windows.MessageBox]::Show("Failed to save configuration to: $File", "Error", "Ok", "Error")
        return
    }
}

function OutGridViewEx {
    [cmdletBinding(DefaultParameterSetName='PassThru')]
    param(
        [Parameter(ValueFromPipeline)]
        [PSObject]$InputObject,

        [ValidateNotNullOrEmpty()]
        [string]$Title,

        [Parameter(ParameterSetName='Wait')]
        [switch]$Wait,

        [Parameter(ParameterSetName='OutputMode')]
        [Microsoft.PowerShell.Commands.OutputModeOption]$OutputMode,

        [Parameter(ParameterSetName='PassThru')]
        [switch]$PassThru,

        [string[]]$VisibleProperty
    )

    begin {
        $customTypeName = 'outgridviewex'
        $userDefined = $PSBoundParameters.ContainsKey('VisibleProperty')

        if ($userDefined) {
            Update-TypeData -TypeName $customTypeName -DefaultDisplayPropertySet $VisibleProperty -Force
            $null = $PSBoundParameters.Remove('VisibleProperty')
        }
        $scriptCmd = {& 'Microsoft.PowerShell.Utility\Out-GridView' @PSBoundParameters }
        $steppablePipeline = $scriptCmd.GetSteppablePipeline($myInvocation.CommandOrigin)
        $steppablePipeline.Begin($PSCmdlet)
    }

    process {
        if ($userDefined) {
            [string[]]$oldType = $_.PSTypeNames
            $_.PSTypeNames.Clear()
            $_.PSTypeNames.Add($customTypeName)
            $steppablePipeline.Process($_)
            $_.PSTypeNames.Clear()
            foreach ($type in $oldType) { 
                $_.PSTypeNames.Add($type) 
            }
        }
        else {
            $steppablePipeline.Process($_)
        }
    }

    end {
        $steppablePipeline.End()
    }
}

function AddCommand {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [string]$Command,

        [Parameter(Mandatory = $false)]
        [string]$Category,

        [Parameter(Mandatory = $false)]
        [string]$PreCommand,

        [Parameter(Mandatory = $false)]
        [switch]$SkipParameterSelect
    )
    
    [array]$config = LoadConfig $script:defaultConfigFile

    $newCommand = [PSObject]@{
        Name = $Name
        Description = $Description
        Command = $Command
        Category = $Category
        PreCommand = $PreCommand
        SkipParameterSelect = $SkipParameterSelect.ToBool()
    }

    $config | foreach-object {
        if ($PSItem.Name -eq $Name) {
            [System.Windows.MessageBox]::Show("Command named $Name already exists.", "Error", "Ok", "Error")
            return
        }
    }

    $config += $newCommand
    SaveConfig -File $script:defaultConfigFile -Content $Config
}

function RemoveCommands {
    $config = LoadConfig $script:defaultConfigFile
    $commands = $config | OutGridViewEx -VisibleProperty Name,Description,Category -Title "SELECT COMMANDS TO REMOVE" -OutputMode Multiple
    $config = $config | Where-Object { $commands -notcontains $PSItem }
    SaveConfig -File $script:defaultConfigFile -Content $config
}

# Main script
Add-Type -AssemblyName PresentationFramework
$script:scriptDir = Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0]))
$script:configDir = "$scriptDir\config"
$script:defaultConfigFile = "$scriptDir\config\config.json"
if (-not (Test-Path $script:configDir)) {
    New-Item -Path $script:configDir -ItemType "directory"
}
InitializeConfig $script:defaultConfigFile

$metaConfig = @"
[
    {
        "Name":  "_ADD_",
        "Description":  "Add new command to launcher",
        "Command":  "AddCommand",
        "Category":  "LaunchPS"
    },
    {
        "Name":  "_REMOVE_",
        "Description":  "Removes commands from launcher",
        "Command":  "RemoveCommands",
        "Category":  "LaunchPS",
        "SkipParameterSelect": true
    },
    {
        "Name":  "_EDIT_",
        "Description":  "Opens primary config file for editing",
        "Command":  "Invoke-Item `$script:defaultConfigFile",
        "Category":  "LaunchPS",
        "SkipParameterSelect": true
    },
    {
        "Name":  "_REFRESH_",
        "Description":  "Refreshes command window",
        "Command":  "return",
        "Category":  "LaunchPS",
        "SkipParameterSelect": true
    }
]
"@ | ConvertFrom-Json

do {
    $stayOpen = $false
    $fullConfig = $null
    $userConfig = $null

    $configs = Get-ChildItem -Path "$scriptDir\config" -Filter "*.json"
    foreach ($config in $configs) {
        $userConfig += LoadConfig $config.FullName
    }
    $fullConfig = $metaConfig + $userConfig

    # Display command selection window
    $command = $fullConfig | OutGridViewEx -VisibleProperty Name,Description,Category -Title "Select command to run" -OutputMode Single

    if ($command.Command) {
        # As long as a command has been selected (ie. window hasn't been exited/canceled) we want to continue
        # running app so another command can be ran
        $stayOpen = $true

        if ($command.SkipParameterSelect) {
            $commandEx = $command.Command
        }
        else {
            $commandEx = Show-Command -Name $command.Command -NoCommonParameter -PassThru -ErrorAction Stop
        }

        # We do not want to run anything if Show-Command was canceled/exited
        if ($commandEx) {
            if ($command.PreCommand) {
                $commandEx = $command.PreCommand + " ; " + $commandEx
            }
            if ($command.Category -eq "LaunchPS") {
                Invoke-Expression $commandEx
            }
            else {
                # We must escape any quotation marks passed or it will cause problems being passed through Start-Process
                $commandEx = $commandEx -replace '"', '\"'
                Start-Process -FilePath powershell.exe -ArgumentList "-ExecutionPolicy Bypass -NoExit `" & { $commandEx } `""
            }
        }
    }
} while ($stayOpen)


