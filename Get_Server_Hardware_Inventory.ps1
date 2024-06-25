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
        $serverHardwareResults = $serverHardware.members | ForEach-Object {
            [PSCustomObject]@{
                ApplianceName      = $global:applianceName
                ServerName         = $_.serverName
                FormFactor         = $_.formFactor
                Model              = $_.model
                Generation         = $_.generation
                MemoryGB           = [math]::round($_.memoryMB / 1024, 2)
                OperatingSystem    = $_.operatingSystem
                Position           = $_.position
                ProcessorCoreCount = $_.processorCoreCount
                ProcessorCount     = $_.processorCount
                ProcessorSpeedMHz  = $_.processorSpeedMhz
                ProcessorType      = $_.processorType
                SerialNumber       = $_.serialNumber
                LocationUri        = $_.locationUri
            }
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
