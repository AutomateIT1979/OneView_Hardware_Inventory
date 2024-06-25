[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Appliance
)

function Save-Workbook {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [object]$Workbook
    )

    try {
        $Workbook.SaveAs($Path)
        Close-ExcelPackage $Workbook
        Write-Host "Workbook saved successfully to $Path"
    }
    catch {
        Write-Error "Failed to save workbook. Error: $_"
    }
}

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
                ApplianceName               = $global:applianceName
                ServerName                  = $server.serverName
                FormFactor                  = $server.formFactor
                Model                       = $server.model
                Generation                  = $server.generation
                MemoryGB                    = [math]::round($server.memoryMB / 1024, 2)
                OperatingSystem             = $server.operatingSystem
                'Position[Server]'          = $server.position
                'ProcessorCoreCount[Server]'= $server.processorCoreCount
                'ProcessorCount[Server]'    = $server.processorCount
                'ProcessorSpeedMHz[Server]' = $server.processorSpeedMhz
                'ProcessorType[Server]'     = $server.processorType
                'SerialNumber[Server]'      = $server.serialNumber
                LocationUri                 = $server.locationUri
            }

            $serverHardwareResults += $serverInfo
        }

        # Set the output path to the same directory where the script is executed
        $serverHardwareOutputXlsx = Join-Path -Path $PSScriptRoot -ChildPath "ServerHardware-$($global:applianceName)-$(Get-Date -format 'yyyy.MM.dd.HHmmss').xlsx"

        # Export to Excel
        $serverHardwareResults | Export-Excel -Path $serverHardwareOutputXlsx -AutoSize -BoldTopRow -WorkSheetname "ServerHardware"
        Write-Host "Server hardware results exported to $serverHardwareOutputXlsx"

        # Pause for a moment to ensure the workbook is saved
        Start-Sleep -Seconds 5

        # Reopen the workbook to verify it was created correctly
        $workbook = Open-ExcelPackage -Path $serverHardwareOutputXlsx
        if ($workbook.Workbook.Worksheets["ServerHardware"]) {
            Write-Host "Verified that the ServerHardware worksheet was created successfully."
        } else {
            throw "Failed to create the ServerHardware worksheet."
        }

        # Collect additional information using LocationUri
        $locationDetailsResults = @()
        foreach ($server in $serverHardwareResults) {
            if ($server.LocationUri) {
                try {
                    $locationDetails = Send-OVRequest -Uri $server.LocationUri -Method GET -ApplianceConnection $ovw
                    if ($locationDetails) {
                        # Collect necessary details from locationDetails
                        $enclosureUri = $locationDetails.enclosureUri
                        $deviceBays = $locationDetails.deviceBays

                        $usedSlots = ($deviceBays | Where-Object { $_.devicePresence -eq 'Present' }).Count
                        $availableSlots = $locationDetails.deviceBayCount - $usedSlots
                        $percentageAvailable = ($availableSlots / $locationDetails.deviceBayCount) * 100

                        $locationDetailsInfo = [PSCustomObject]@{
                            ApplianceName         = $global:applianceName
                            EnclosureUri          = $enclosureUri
                            DeviceBayCount        = $locationDetails.deviceBayCount
                            UsedSlots             = $usedSlots
                            AvailableSlots        = $availableSlots
                            PercentageAvailable   = [math]::round($percentageAvailable, 2)
                        }

                        $locationDetailsResults += $locationDetailsInfo
                    }
                }
                catch {
                    Write-Error "Failed to retrieve location details for $($server.ServerName). Error: $_"
                }
            }
        }

        # Add location details to a new worksheet
        $locationDetailsWorksheetName = "LocationDetails"
        if ($locationDetailsResults.Count -gt 0) {
            Write-Host "Exporting location details to worksheet."
            $locationDetailsResults | Export-Excel -ExcelPackage $workbook -WorkSheetname $locationDetailsWorksheetName -AutoSize -BoldTopRow
        } else {
            Write-Host "No location details information to export."
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

        # Save the workbook
        Save-Workbook -Path $serverHardwareOutputXlsx -Workbook $workbook
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
