# Harmless sample commands for demoing PSGUI. They only write to the pipeline and sleep briefly.

function Test-DemoConnection {
    <#
    .SYNOPSIS
    Pretends to test connectivity to a host.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$ComputerName,

        [ValidateSet('ICMP', 'TCP')]
        [string]$Protocol = 'ICMP',

        [int]$Count = 4
    )

    foreach ($i in 1..$Count) {
        Start-Sleep -Milliseconds 400
        [pscustomobject]@{
            Attempt      = $i
            ComputerName = $ComputerName
            Protocol     = $Protocol
            LatencyMs    = Get-Random -Minimum 8 -Maximum 40
            Status       = 'Reachable'
        }
    }
}

function Get-DemoServiceReport {
    <#
    .SYNOPSIS
    Lists a sample of services and their state.
    #>
    param(
        [string]$Name = '*',

        [switch]$IncludeStopped
    )

    $services = @(
        [pscustomobject]@{ Name = 'Spooler';  DisplayName = 'Print Spooler';       Status = 'Running' }
        [pscustomobject]@{ Name = 'BITS';     DisplayName = 'Background Transfer'; Status = 'Running' }
        [pscustomobject]@{ Name = 'W32Time';  DisplayName = 'Windows Time';        Status = 'Running' }
        [pscustomobject]@{ Name = 'Fax';      DisplayName = 'Fax';                 Status = 'Stopped' }
        [pscustomobject]@{ Name = 'WSearch';  DisplayName = 'Windows Search';      Status = 'Stopped' }
    )
    $services | Where-Object { $_.Name -like $Name -and ($IncludeStopped -or $_.Status -eq 'Running') }
}

function Restart-DemoService {
    <#
    .SYNOPSIS
    Pretends to restart a service, with progress output.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [switch]$Force
    )

    Write-Output "Stopping $Name..."
    Start-Sleep -Seconds 1
    if ($Force) { Write-Output 'Force specified: dependent services stopped first' }
    Write-Output "Starting $Name..."
    Start-Sleep -Seconds 1
    Write-Output "$Name is Running"
}

function Clear-DemoCache {
    <#
    .SYNOPSIS
    Pretends to clear caches.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('User', 'System', 'All')]
        [string]$Scope,

        [switch]$RestartAfter
    )

    $targets = switch ($Scope) {
        'User'   { 'Browser cache', 'Temp files' }
        'System' { 'DNS resolver cache', 'Icon cache' }
        'All'    { 'Browser cache', 'Temp files', 'DNS resolver cache', 'Icon cache' }
    }
    foreach ($target in $targets) {
        Write-Output "Clearing $target"
        Start-Sleep -Milliseconds 600
    }
    Write-Output "Cleared $($targets.Count) caches"
    if ($RestartAfter) { Write-Output 'Restart requested' }
}

function Get-DemoDiskUsage {
    <#
    .SYNOPSIS
    Shows sample disk usage for a drive.
    #>
    param(
        [ValidateSet('C:', 'D:', 'E:')]
        [string]$Drive = 'C:',

        [switch]$Detailed
    )

    $usage = [pscustomobject]@{ Drive = $Drive; SizeGB = 512; UsedGB = 311; FreeGB = 201; PercentFree = 39 }
    if ($Detailed) {
        $usage | Add-Member -NotePropertyName LargestFolder -NotePropertyValue 'Users' -PassThru
    }
    else {
        $usage
    }
}

function New-DemoUserReport {
    <#
    .SYNOPSIS
    Builds a sample user report for a department.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Sales', 'IT', 'HR')]
        [string]$Department,

        [ValidateSet('Table', 'List')]
        [string]$Format = 'Table',

        [switch]$IncludeDisabled
    )

    $users = @(
        [pscustomobject]@{ Name = 'Avery Stone';  Department = 'Sales'; Enabled = $true }
        [pscustomobject]@{ Name = 'Jordan Reyes'; Department = 'IT';    Enabled = $true }
        [pscustomobject]@{ Name = 'Casey Morgan'; Department = 'IT';    Enabled = $false }
        [pscustomobject]@{ Name = 'Riley Chen';   Department = 'HR';    Enabled = $true }
        [pscustomobject]@{ Name = 'Taylor Brooks'; Department = 'Sales'; Enabled = $true }
    )
    $rows = $users | Where-Object { $_.Department -eq $Department -and ($IncludeDisabled -or $_.Enabled) }
    if ($Format -eq 'List') { $rows | Format-List } else { $rows | Format-Table -AutoSize }
}

function Invoke-DemoBackup {
    <#
    .SYNOPSIS
    Pretends to back up a folder, with step-by-step output.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Source,

        [Parameter(Mandatory)]
        [string]$Destination,

        [switch]$Compress
    )

    Write-Output "Scanning $Source"
    Start-Sleep -Seconds 1
    Write-Output 'Found 128 files (412 MB)'
    Start-Sleep -Milliseconds 800
    if ($Compress) {
        Write-Output 'Compressing...'
        Start-Sleep -Seconds 1
    }
    Write-Output "Copying to $Destination"
    Start-Sleep -Seconds 1
    Write-Output 'Backup complete'
}

function Sync-DemoInventory {
    <#
    .SYNOPSIS
    Pretends to sync an inventory with a server.
    #>
    param(
        [string]$Server = 'inventory.example.com'
    )

    Write-Output "Connecting to $Server"
    Start-Sleep -Milliseconds 700
    Write-Output 'Uploading 42 records'
    Start-Sleep -Seconds 1
    Write-Output 'Sync complete'
}

Export-ModuleMember -Function *-Demo*
