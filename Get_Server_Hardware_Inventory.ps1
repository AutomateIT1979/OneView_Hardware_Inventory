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
        Write-Host "Retrieved server hardware: $($serverHardware.members.Count)"
        $serverHardwareResults = @()
        foreach ($server in $serverHardware.members) {
            $serverInfo = [PSCustomObject]@{
                ApplianceName      = $global:applianceName
                ServerName         = $server.serverName
                FormFactor         = $server.formFactor
                Model              = $server.model
                Generation         = $server.generation
                MemoryGB           = [math]::round($server.memoryMB / 1024, 2)
                OperatingSystem    = $server.operatingSystem
                Position           = $server.position
                ProcessorCoreCount = $server.processorCoreCount
                ProcessorCount     = $server.processorCount
                ProcessorSpeedMHz  = $server.processorSpeedMhz
                ProcessorType      = $server.processorType
                SerialNumber       = $server.serialNumber
                LocationUri        = $server.locationUri
            }
            # Fetch additional details using LocationUri
            if ($server.locationUri) {
                try {
                    $locationDetails = Send-OVRequest -Uri $server.locationUri -Method GET -ApplianceConnection $ovw
                    # Update the serverInfo object with additional details
                    $serverInfo | Add-Member -MemberType NoteProperty -Name "LocationSerialNumber" -Value $locationDetails.serialNumber
                    if ($locationDetails.deviceBays) {
                        $deviceBaysInfo = $locationDetails.deviceBays | ForEach-Object {
                            [PSCustomObject]@{
                                DevicePresence       = $_.devicePresence
                                DeviceFormFactor     = $_.deviceFormFactor
                                PowerAllocationWatts = $_.powerAllocationWatts
                            }
                        }
                        $serverInfo | Add-Member -MemberType NoteProperty -Name "DeviceBays" -Value $deviceBaysInfo
                    }
                }
                catch {
                    Write-Error "Failed to retrieve location details for $($server.serverName). Error: $_"
                }
            }
            $serverHardwareResults += $serverInfo
        }
        # Export to CSV
        $serverHardwareOutputCsv = "ServerHardware-$($global:applianceName)-$(Get-Date -format 'yyyy.MM.dd.HHmm').csv"
        $serverHardwareResults | Export-Csv -Path $serverHardwareOutputCsv -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Host "Server hardware results exported to $serverHardwareOutputCsv"
        # Export to Excel
        $serverHardwareOutputXlsx = "ServerHardware-$($global:applianceName)-$(Get-Date -format 'yyyy.MM.dd.HHmm').xlsx"
        $serverHardwareResults | Export-Excel -Path $serverHardwareOutputXlsx -AutoSize -BoldTopRow -WorkSheetname "ServerHardware"
        $workbook = Open-ExcelPackage -Path $serverHardwareOutputXlsx
        $worksheet = $workbook.Workbook.Worksheets["ServerHardware"]
        # Apply design
        $worksheet.Cells.Style.HorizontalAlignment = 'Left'
        $worksheet.Cells.Style.VerticalAlignment = 'Top'
        $worksheet.Cells.AutoFitColumns()
        $worksheet.Cells["A1:Z1"].Style.Font.Bold = $true
        $worksheet.Cells["A1:Z1"].Style.Fill.PatternType = 'Solid'
        $worksheet.Cells["A1:Z1"].Style.Fill.BackgroundColor.SetColor('Yellow')
        $worksheet.Cells["A1:Z1"].Style.Font.Color.SetColor('Black')
        Close-ExcelPackage $workbook
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
