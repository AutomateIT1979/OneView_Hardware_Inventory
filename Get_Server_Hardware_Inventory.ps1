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
        $enclosureSlots = @{}
        foreach ($server in $serverHardware.members) {
            $serverInfo = [PSCustomObject]@{
                ApplianceName        = $global:applianceName
                ServerName          = $server.serverName
                FormFactor          = $server.formFactor
                Model            = $server.model
                Generation          = $server.generation
                MemoryGB            = [math]::round($server.memoryMB / 1024, 2)
                OperatingSystem      = $server.operatingSystem
                Position            = $server.position
                ProcessorCoreCount    = $server.processorCoreCount
                ProcessorCount        = $server.processorCount
                ProcessorSpeedMHz    = $server.processorSpeedMhz
                ProcessorType        = $server.processorType
                SerialNumber        = $server.serialNumber
                LocationUri        = $server.locationUri
                LocationSerialNumber = $null
                DevicePresence        = $null
                DeviceFormFactor      = $null
                PowerAllocationWatts = $null
            }
            # Fetch additional details using LocationUri
            if ($server.locationUri) {
                try {
                    $locationDetails = Send-OVRequest -Uri $server.locationUri -Method GET -ApplianceConnection $ovw
                    if ($null -eq $locationDetails) {
                        throw "Failed to retrieve location details."
                    }
                    # Update the serverInfo object with additional details
                    $serverInfo.LocationSerialNumber = $locationDetails.serialNumber
                    if ($locationDetails.deviceBays) {
                        # Assuming we are interested in the first device bay for simplicity
                        $firstDeviceBay = $locationDetails.deviceBays[0]
                        $serverInfo.DevicePresence = $firstDeviceBay.devicePresence
                        $serverInfo.DeviceFormFactor = $firstDeviceBay.deviceFormFactor
                        $serverInfo.PowerAllocationWatts = $firstDeviceBay.powerAllocationWatts
                    }
                    # Collect enclosure information for slot availability
                    if ($locationDetails.enclosureUri) {
                        $enclosureUri = $locationDetails.enclosureUri
                        if ($null -eq $enclosureUri) {
                            Write-Warning "No enclosureUri for server: $($server.serverName)"
                        } else {
                            $enclosureDetails = Send-OVRequest -Uri $enclosureUri -Method GET -ApplianceConnection $ovw
                            if ($null -eq $enclosureDetails) {
                                throw "Failed to retrieve enclosure details."
                            }
                            if (-not $enclosureSlots.ContainsKey($enclosureDetails.serialNumber)) {
                                $enclosureSlots[$enclosureDetails.serialNumber] = [PSCustomObject]@{
                                    ApplianceName          = $global:applianceName
                                    EnclosureSerialNumber = $enclosureDetails.serialNumber
                                    AvailableSlots        = 0
                                }
                            }
                        }
                        $availableSlots = ($enclosureDetails.deviceBays | Where-Object { $_.devicePresence -eq 'Absent' }).Count
                        $enclosureSlots[$enclosureDetails.serialNumber].AvailableSlots += $availableSlots
                        Write-Host "Enclosure $($enclosureDetails.serialNumber) has $availableSlots available slots."
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
        # Add server hardware details to a new worksheet
        $hardwareWorksheetName = "ServerHardwareDetails"
        $serverHardwareResults | Export-Excel -ExcelPackage $workbook -WorkSheetname $hardwareWorksheetName -AutoSize -BoldTopRow
        # Add location information to a new worksheet
        $locationWorksheetName = "ServerLocationDetails"
        $locationUriList = $serverHardwareResults | Select-Object -Unique LocationUri
        if ($locationUriList.Count -gt 0) {
            Write-Host "Exporting Server Location information to worksheet."
            $locationUriList | Export-Excel -ExcelPackage $workbook -WorkSheetname $locationWorksheetName -AutoSize -BoldTopRow
        } else {
            Write-Host "No Server Location information to export."
        }
        # Add enclosure slot availability to a new worksheet
        $enclosureWorksheetName = "EnclosureSlots"
        $enclosureSlotList = $enclosureSlots.Values
        if ($enclosureSlotList.Count -gt 0) {
            Write-Host "Exporting enclosure slot information to worksheet."
            $enclosureSlotList | Export-Excel -ExcelPackage $workbook -WorkSheetname $enclosureWorksheetName -AutoSize -BoldTopRow
        } else {
            Write-Host "No enclosure slot information to export."
        }
        # Apply design to all worksheets
        foreach ($worksheet in $workbook.Workbook.Worksheets) {
            $worksheet.Cells.Style.HorizontalAlignment = 'Left'
            $worksheet.Cells.Style.VerticalAlignment = 'Top'
            $worksheet.Cells.AutoFitColumns()
            $worksheet.Cells["A1"].EntireRow.Style.Font.Bold = $true  # Bold the first row
            $worksheet.Cells["A1"].EntireRow.Style.Fill.PatternType = 'Solid'
            $worksheet.Cells["A1"].EntireRow.Style.Fill.BackgroundColor.SetColor('Yellow')
            $worksheet.Cells["A1"].EntireRow.Style.Font.Color.SetColor('Black')
        }
        Close-ExcelPackage $workbook
        Write-Host "Server hardware, Server Location, and enclosure slot results exported to $serverHardwareOutputXlsx"
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
