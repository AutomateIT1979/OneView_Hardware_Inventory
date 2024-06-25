[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Appliance
)
Begin {
    # Load necessary modules
    Import-Module ImportExcel
    # Initialize connection to OneView appliance
    try {
        $ovw = Connect-OVMgmt -Appliance $Appliance
        if ($null -eq $ovw) {
            throw "Failed to connect to OneView appliance."
        }
        Write-Host "Connected to appliance: $($ovw.Name)"
        # Extract appliance name from FQDN and convert to uppercase
        $global:applianceName = $Appliance.Split('.')[0].ToUpper()
    }
    catch {
        Write-Error "OneView appliance connection failed: $_"
        throw
    }
}
Process {
    Write-Host "Getting all server hardware from appliance: $($ovw.Name)"
    try {
        $serverHardware = Send-OVRequest -Uri "/rest/server-hardware" -Method GET -ApplianceConnection $ovw
        if ($null -eq $serverHardware -or $null -eq $serverHardware.members) {
            throw "Failed to retrieve server hardware information."
        }
        Write-Host "Retrieved server hardware: $($serverHardware.members.Count)"
        $serverHardwareResults = @()
        foreach ($server in $serverHardware.members) {
            $serverInfo = [PSCustomObject]@{
                ApplianceName       = $global:applianceName
                ServerName          = $server.serverName
                FormFactor          = $server.formFactor
                Model               = $server.model
                Generation          = $server.generation
                MemoryGB            = [math]::round($server.memoryMB / 1024, 2)
                OperatingSystem     = $server.operatingSystem
                Position            = $server.position
                ProcessorCoreCount  = $server.processorCoreCount
                ProcessorCount      = $server.processorCount
                ProcessorSpeedMHz   = $server.processorSpeedMhz
                ProcessorType       = $server.processorType
                SerialNumber        = $server.serialNumber
                LocationUri         = $server.locationUri
            }
            $serverHardwareResults += $serverInfo
        }
        # Export to Excel
        $serverHardwareOutputXlsx = "ServerHardware-$($global:applianceName)-$(Get-Date -format 'yyyy.MM.dd.HHmm').xlsx"
        $serverHardwareResults | Export-Excel -Path $serverHardwareOutputXlsx -AutoSize -BoldTopRow -WorkSheetname "ServerHardware"
        Write-Host "Server hardware results exported to $serverHardwareOutputXlsx"
    }
    catch {
        Write-Error "Failed to retrieve server hardware from appliance: $($ovw.Name). Error: $_"
    }
}
End {
    try {
        Disconnect-OVMgmt -ApplianceConnection $ovw
        Write-Host "Disconnected from appliance: $($ovw.Name)"
    }
    catch {
        Write-Error "Failed to disconnect from appliance: $($ovw.Name). Error: $_"
    }
}
