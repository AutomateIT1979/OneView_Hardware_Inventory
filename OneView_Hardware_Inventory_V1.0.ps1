# Clear the console window
Clear-Host
# Create a string of 4 spaces
$Spaces = [string]::new(' ', 4)
# Define the script version
$ScriptVersion = "1.0"
# Get the directory from which the script is being executed
$scriptDirectory = $PSScriptRoot
# Get the parent directory of the script's directory
$parentPath = Split-Path -Parent $scriptDirectory
# Define the logging function Directory
$loggingFunctionsDirectory = Join-Path -Path $parentPath -ChildPath "Logging_Function"
# Construct the path to the Logging_Functions.ps1 script
$loggingFunctionsPath = Join-Path -Path $loggingFunctionsDirectory -ChildPath "Logging_Functions.ps1"
# Script Header main script
$HeaderMainScript = @"
Author : Your Name
Description : This script does amazing things!
Created : $(Get-Date -Format "dd/MM/yyyy")
Last Modified : $((Get-Item $PSCommandPath).LastWriteTime.ToString("dd/MM/yyyy"))
"@
# Display the header information in the console with a design
$consoleWidth = $Host.UI.RawUI.WindowSize.Width
$line = "─" * ($consoleWidth - 2)
Write-Host "+$line+" -ForegroundColor DarkGray
# Split the header into lines and display each part in different colors
$HeaderMainScript -split "`n" | ForEach-Object {
    $parts = $_ -split ": ", 2
    Write-Host "`t" -NoNewline
    Write-Host $parts[0] -NoNewline -ForegroundColor DarkGray
    Write-Host ": " -NoNewline
    Write-Host $parts[1] -ForegroundColor Cyan
}
Write-Host "+$line+" -ForegroundColor DarkGray
# Check if the Logging_Functions.ps1 script exists
if (Test-Path -Path $loggingFunctionsPath) {
    # Dot-source the Logging_Functions.ps1 script
    . $loggingFunctionsPath
    # Write a message to the console indicating that the logging functions have been loaded
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Logging functions have been loaded." -ForegroundColor Green
}
else {
    # Write an error message to the console indicating that the logging functions script could not be found
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The logging functions script could not be found at: $loggingFunctionsPath" -ForegroundColor Red
    # Stop the script execution
    exit
}
# Initialize task counter
$script:taskNumber = 1
# Define the function to import required modules if they are not already imported
function Import-ModulesIfNotExists {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ModuleNames
    )
    # Start logging
    Start-Log -ScriptVersion $ScriptVersion -ScriptPath $PSCommandPath
    # Task 1: Checking required modules
    Write-Host "`n$Spaces$($taskNumber). Checking required modules:`n" -ForegroundColor Magenta
    # Log the task
    Write-Log -Message "Checking required modules." -Level "Info" -NoConsoleOutput
    # Increment $script:taskNumber after the function call
    $script:taskNumber++
    # Total number of modules to check
    $totalModules = $ModuleNames.Count
    # Initialize the current module counter
    $currentModuleNumber = 0
    foreach ($ModuleName in $ModuleNames) {
        $currentModuleNumber++
        # Simple text output for checking required modules
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Checking module " -NoNewline -ForegroundColor DarkGray
        Write-Host "$currentModuleNumber" -NoNewline -ForegroundColor White
        Write-Host " of " -NoNewline -ForegroundColor DarkGray
        Write-Host "${totalModules}" -NoNewline -ForegroundColor Cyan
        Write-Host ": $ModuleName" -ForegroundColor White
        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
                Write-Host "`t• " -NoNewline -ForegroundColor White
                Write-Host "Module " -NoNewline -ForegroundColor White
                Write-Host "$ModuleName" -NoNewline -ForegroundColor Red
                Write-Host " is not installed." -ForegroundColor White
                Write-Log -Message "Module '$ModuleName' is not installed." -Level "Error" -NoConsoleOutput
                continue
            }
            # Check if the module is already imported
            if (Get-Module -Name $ModuleName) {
                Write-Host "`t• " -NoNewline -ForegroundColor White
                Write-Host "Module " -NoNewline -ForegroundColor DarkGray
                Write-Host "$ModuleName" -NoNewline -ForegroundColor Yellow
                Write-Host " is already imported." -ForegroundColor DarkGray
                Write-Log -Message "Module '$ModuleName' is already imported." -Level "Info" -NoConsoleOutput
                continue
            }
            # Try to import the module
            Import-Module $ModuleName -ErrorAction Stop
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Module " -NoNewline -ForegroundColor DarkGray
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Green
            Write-Host " imported successfully." -ForegroundColor DarkGray
            Write-Log -Message "Module '[$ModuleName]' imported successfully." -Level "OK" -NoConsoleOutput
        }
        catch {
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Failed to import module " -NoNewline
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Red
            Write-Host ": $_" -ForegroundColor Red
            Write-Log -Message "Failed to import module '[$ModuleName]': $_" -Level "Error" -NoConsoleOutput
        }
        # Add a delay to slow down the progress bar
        Start-Sleep -Seconds 1
    }
}
# Import the required modules
Import-ModulesIfNotExists -ModuleNames 'HPEOneView.660', 'Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility', 'ImportExcel'
# Task 2: Checking if Excel is insalled on system
Write-Host "`n$Spaces$($taskNumber). Checking if Excel is installed:`n" -ForegroundColor Magenta
# Log Task
Write-Log -Message "Checking if Excel is installed." -Level "Info" -NoConsoleOutput
# Increment $script:taskNumber for Task 2
$script:taskNumber++
function Test-ExcelInstallation {
    # Attempt to create an Excel COM object
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        # Write a message to the console
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Excel is installed." -ForegroundColor Green
        # Write a message to the log file
        Write-Log -Message "Excel is installed." -Level "OK" -NoConsoleOutput
        # Retrieve and display additional information about the Excel installation
        $version = $excel.Version
        $build = $excel.Build
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Excel version: $version" -ForegroundColor Green
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Excel build: $build" -ForegroundColor Green
        # Write to the log file
        Write-Log -Message "Excel version: $version" -Level "Info" -NoConsoleOutput
        Write-Log -Message "Excel build: $build" -Level "Info" -NoConsoleOutput
    }
    catch {
        # Write a message to the console
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Excel is not installed." -ForegroundColor Red
        # Write a message to the log file
        Write-Log -Message "Excel is not installed." -Level "Error" -NoConsoleOutput
        return $false
    }
    finally {
        if ($null -ne $excel) {
            # Quit Excel
            $excel.Quit()
            # Release the COM object
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            Remove-Variable -Name excel
        }
    }
    return $true
}
# Check if Excel is installed at the beginning of the script
$excelInstalled = Test-ExcelInstallation
if (-not $excelInstalled) {
    Write-Host "Excel is not installed. Please install Excel and then run the script again." -ForegroundColor Red
    Write-LOg -Message "Excel is not installed. Please install Excel and then run the script again." -Level "Error" -NoConsoleOutput 
    return
} 
# Define the CSV file name
$csvFileName = ".\Appliances_List\Appliances_List.csv"
# Define the parent directory of the CSV file
$parentDirectory = Split-Path -Path $scriptDirectory -Parent
# Create the full path to the CSV file
$csvFilePath = Join-Path -Path $parentDirectory -ChildPath $csvFileName
# Define the path to the credential folder
$credentialFolder = Join-Path -Path $parentDirectory -ChildPath "Credential"
# Task 3: import Appliances list from the CSV file.
Write-Host "`n$Spaces$($taskNumber). Importing Appliances list from the CSV file:`n" -ForegroundColor Magenta
# Import Appliances list from CSV file
$Appliances = Import-Csv -Path $csvFilePath
# Increment $script:taskNumber for Task 3
$script:taskNumber++
# Confirm that the CSV file was imported successfully
if ($Appliances) {
    # Get the total number of appliances
    $totalAppliances = $Appliances.Count
    # Log the total number of appliances
    Write-Log -Message "There are $totalAppliances appliances in the CSV file." -Level "Info" -NoConsoleOutput
    # Display if the CSV file was imported successfully
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The CSV file was imported " -NoNewline -ForegroundColor DarkGray
    Write-Host "successfully" -NoNewline -ForegroundColor Green
    Write-Host "." -ForegroundColor DarkGray
    # Display the total number of appliances
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Total number of appliances:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $totalAppliances" -NoNewline -ForegroundColor Cyan
    Write-Host "" # This is to add a newline after the above output
    # Log the successful import of the CSV file
    Write-Log -Message "The CSV file was imported successfully." -Level "OK" -NoConsoleOutput
}
else {
    # Display an error message if the CSV file failed to import
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Failed to import the CSV file." -ForegroundColor Red
    # Log the failure to import the CSV file
    Write-Log -Message "Failed to import the CSV file." -Level "Error" -NoConsoleOutput
}
# Task 4: Check if credential folder exists
Write-Host "`n$Spaces$($taskNumber). Checking for credential folder:`n" -ForegroundColor Magenta
# Log the task
Write-Log -Message "Checking for credential folder." -Level "Info" -NoConsoleOutput
# Increment $script:taskNumber for Task 4
$script:taskNumber++
# Check if the credential folder exists, if not say it at console and create it, if already exist say it at console
if (Test-Path -Path $credentialFolder) {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Credential folder already exists at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $credentialFolder" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "Credential folder already exists at $credentialFolder" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Credential folder does not exist." -NoNewline -ForegroundColor Red
    Write-Host " Creating now..." -ForegroundColor DarkGray
    Write-Log -Message "Credential folder does not exist, creating now..." -Level "Info" -NoConsoleOutput
    # Create the credential folder if it does not exist already
    New-Item -ItemType Directory -Path $credentialFolder | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Credential folder created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $credentialFolder" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "Credential folder created at $credentialFolder" -Level "OK" -NoConsoleOutput
}
# Define the path to the credential file
$credentialFile = Join-Path -Path $credentialFolder -ChildPath "credential.txt"
# Log the task
Write-Log -Message "Checking for credential file." -Level "Info" -NoConsoleOutput
# Check if the credential file exists
if (-not (Test-Path -Path $credentialFile)) {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Credential file does not exist." -NoNewline -ForegroundColor Red
    Write-Host " Creating now..." -ForegroundColor DarkGray
    Write-Log -Message "Credential file does not exist, creating now..." -Level "Info" -NoConsoleOutput
    # Prompt the user to enter their login and password
    $credential = Get-Credential -Message "Please enter your login and password."
    # Save the credential to the credential file
    $credential | Export-Clixml -Path $credentialFile
}
else {
    # Load the credential from the credential file
    $credential = Import-Clixml -Path $credentialFile
}
# Define the directories for the CSV and Excel files
$csvDir = Join-Path -Path $script:ReportsDir -ChildPath 'CSV'
$excelDir = Join-Path -Path $script:ReportsDir -ChildPath 'Excel'
# Task 5: Check if Excel is running and close it if necessary
Write-Host "`n$Spaces$($taskNumber). Checking if Excel is running and closing it if necessary:`n" -ForegroundColor Magenta
# Log the task
Write-Log -Message "Checking if Excel is running and closing it if necessary." -Level "Info" -NoConsoleOutput
# Increment $script:taskNumber for Task 5
$script:taskNumber++
function Test-ExcelProcess {
    # Check if any Excel process is running
    $excelProcess = Get-Process excel -ErrorAction SilentlyContinue
    if ($excelProcess.Count -gt 0) {
        # Excel is running
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Excel is currently running. Attempting to close Excel..." -ForegroundColor Yellow
        # Close the Excel process
        Stop-Process -Name excel -Force -ErrorAction SilentlyContinue
        # Wait for a few seconds to allow the process to close
        Start-Sleep -Seconds 5
        # Recheck if any Excel process is running
        $excelProcess = Get-Process excel -ErrorAction SilentlyContinue
        if ($excelProcess.Count -gt 0) {
            # Excel is still running
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Unable to close Excel. Please close it manually and then run the script again." -ForegroundColor Red
        }
        else {
            # Excel has been closed
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Excel has been closed. You can proceed with the script." -ForegroundColor Green
        }
    }
    else {
        # Excel is not running
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Excel is not running. You can proceed with the script." -ForegroundColor Green
    }
}
# Check if Excel is running
Test-ExcelProcess
# Task 6: Check if the CSV and Excel directories exist
Write-Host "`n$Spaces$($taskNumber). Checking for CSV and Excel directories:`n" -ForegroundColor Magenta
# Log the task
Write-Log -Message "Checking for CSV and Excel directories." -Level "Info" -NoConsoleOutput
# Increment $script:taskNumber for Task 6
$script:taskNumber++
# Check if the CSV directory exists
if (Test-Path -Path $csvDir) {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory already exists at:" -NoNewline -ForegroundColor DarkGray
    write-host " $csvDir" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "CSV directory already exists at $csvDir" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory does not exist." -NoNewline -ForegroundColor Red
    Write-Host " Creating now..." -ForegroundColor DarkGray
    Write-Log -Message "CSV directory does not exist, creating now..." -Level "Info" -NoConsoleOutput
    # Create the CSV directory if it does not exist already
    New-Item -ItemType Directory -Path $csvDir | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $csvDir" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "CSV directory created at $csvDir" -Level "OK" -NoConsoleOutput
}
# Check if the Excel directory exists
if (Test-Path -Path $excelDir) {
    # Write a message to the console
    write-host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory already exists at:" -NoNewline -ForegroundColor DarkGray
    write-host " $excelDir" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "Excel directory already exists at $excelDir" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory does not exist at" -NoNewline -ForegroundColor Red
    Write-Host " $excelDir" -ForegroundColor DarkGray
    # Write a message to the log file
    Write-Log -Message "Excel directory does not exist at $excelDir, creating now..." -Level "Info" -NoConsoleOutput
    # Create the Excel directory if it does not exist already
    New-Item -ItemType Directory -Path $excelDir | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $excelDir" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "Excel directory created at $excelDir" -Level "OK" -NoConsoleOutput
}
# Task 7: Loop through the appliances list and get hardware inventory
# Increment $script:taskNumber for Task 7
$script:taskNumber++
# Task 7: Loop through the appliances list and get hardware inventory
Write-Host "`n$Spaces$($taskNumber). Loop through the appliances list and get hardware inventory:`n" -ForegroundColor Magenta
# Log the task
Write-Log -Message "Loop through the appliances list and get hardware inventory." -Level "Info" -NoConsoleOutput
# Loop through each appliance in the list
foreach ($appliance in $Appliances) {
    # Convert the FQDN to Upper Case
    $FQDN = $appliance.FQDN.ToUpper()
    # Check if there is a connection to the appliances
    $existingSessions = $ConnectedSessions
    if ($existingSessions) {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Existing sessions found: $($existingSessions.Count)" -ForegroundColor Yellow
        Write-Log -Message "Existing sessions found: $($existingSessions.Count)" -Level "Info" -NoConsoleOutput
        # Disconnect all existing sessions
        $existingSessions | ForEach-Object {
            Disconnect-OVMgmt -Hostname $_
        }
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "All existing sessions have been disconnected." -ForegroundColor Green
        Write-Log -Message "All existing sessions have been disconnected." -Level "OK" -NoConsoleOutput
        # Add a small delay to ensure the session is fully disconnected
        Start-Sleep -Seconds 5
    }
    else {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "No existing sessions found.`n" -ForegroundColor Gray
        Write-Log -Message "No existing sessions found." -Level "Info" -NoConsoleOutput
    }
    # Use the Connect-OVMgmt cmdlet to connect to the appliance
    Connect-OVMgmt -Hostname $FQDN -Credential $credential *> $null
    # Check if the connection was successful
    if ($?) {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Connected to appliance " -NoNewline -ForegroundColor DarkGray
        Write-Host "$FQDN" -NoNewline -ForegroundColor Cyan
        Write-Host " successfully" -ForegroundColor Green
        Write-Host "." -ForegroundColor DarkGray
        Write-Log -Message "Connected to appliance $FQDN successfully." -Level "OK" -NoConsoleOutput
    }
    else {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Failed to connect to appliance " -NoNewline -ForegroundColor DarkGray
        Write-Host "$FQDN" -NoNewline -ForegroundColor Red
        Write-Host ": $_" -ForegroundColor Red
        Write-Log -Message "Failed to connect to appliance "$FQDN": $_" -Level "Error" -NoConsoleOutput
        continue
    }
    # Get the hardware inventory
    $hardwareInventory = Get-OVServer
    # Check if the hardware inventory was retrieved successfully
    if ($hardwareInventory) {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Hardware inventory retrieved" -NoNewline -ForegroundColor DarkGray
        Write-Host " successfully" -NoNewline -ForegroundColor Green
        Write-Host " for appliance " -NoNewline -ForegroundColor DarkGray
        Write-Host "$FQDN" -NoNewline -ForegroundColor Cyan
        Write-Host "." -ForegroundColor DarkGray
        Write-Log -Message "Hardware inventory retrieved successfully for appliance $FQDN." -Level "OK" -NoConsoleOutput
    }
    else {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Failed to retrieve hardware inventory for appliance " -NoNewline -ForegroundColor DarkGray
        Write-Host "$FQDN" -NoNewline -ForegroundColor Red
        Write-Host "." -ForegroundColor DarkGray
        Write-Log -Message "Failed to retrieve hardware inventory for appliance $FQDN." -Level "Error" -NoConsoleOutput
        continue
    }
    # Export the hardware inventory to a CSV file
    $csvFileName = "$FQDN-HardwareInventory.csv"
    $csvFilePath = Join-Path -Path $csvDir -ChildPath $csvFileName
    $hardwareInventory | Export-Csv -Path $csvFilePath -NoTypeInformation
    # Check if the CSV file was exported successfully
    if (Test-Path -Path $csvFilePath) {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Hardware inventory exported to CSV file " -NoNewline -ForegroundColor DarkGray
        Write-Host "successfully." -ForegroundColor Green
        Write-Log -Message "Hardware inventory exported to CSV file $csvFilePath successfully." -Level "OK" -NoConsoleOutput
    }
    else {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Failed to export hardware inventory to CSV file " -NoNewline -ForegroundColor Red
        Write-Host "." -ForegroundColor DarkGray
        Write-Log -Message "Failed to export hardware inventory to CSV file $csvFilePath." -Level "Error" -NoConsoleOutput
    }
    # Export the hardware inventory to an Excel file
    $excelFileName = "$FQDN-HardwareInventory.xlsx"
    $excelFilePath = Join-Path -Path $excelDir -ChildPath $excelFileName
    # Now you can safely export the hardware inventory to the Excel file
    $hardwareInventory | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    # Check if the Excel file was exported successfully
    if (Test-Path -Path $excelFilePath) {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Hardware inventory exported to Excel file " -NoNewline -ForegroundColor DarkGray
        Write-Host "successfully" -NoNewline -ForegroundColor Green
        Write-Host "." -ForegroundColor DarkGray
        Write-Log -Message "Hardware inventory exported to Excel file $excelFilePath successfully." -Level "OK" -NoConsoleOutput
    }
    else {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Failed to export hardware inventory to Excel file " -NoNewline -ForegroundColor Red
        Write-Log -Message "Failed to export hardware inventory to Excel file $excelFilePath." -Level "Error" -NoConsoleOutput
    }
    # Disconnect from the appliance
    Disconnect-OVMgmt -Hostname $FQDN
    # Check if the disconnection was successful
    if ($?) {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Disconnected from appliance " -NoNewline -ForegroundColor DarkGray
        Write-Host "$FQDN" -NoNewline -ForegroundColor Cyan
        Write-Host " successfully" -ForegroundColor Green
        Write-Host "." -ForegroundColor DarkGray
        Write-Log -Message "Disconnected from appliance $FQDN successfully." -Level "OK" -NoConsoleOutput
    }
    else {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Failed to disconnect from appliance " -NoNewline -ForegroundColor DarkGray
        Write-Host "$FQDN" -NoNewline -ForegroundColor Red
        Write-Host "." -ForegroundColor DarkGray
        Write-Log -Message "Failed to disconnect from appliance $FQDN." -Level "Error" -NoConsoleOutput
    }
}
# Task 8: Script execution completed successfully 
Write-Host "`n$Spaces$($taskNumber). Script execution completed successfully.`n" -ForegroundColor Magenta
# Log the task
Write-Log -Message "Script execution completed successfully." -Level "Info" -NoConsoleOutput
# Increment $script:taskNumber for Task 8
$script:taskNumber++
# Just before calling Complete-Logging
$endTime = Get-Date
$totalRuntime = $endTime - $startTime
# Call Complete-Logging at the end of the script
Complete-Logging -LogPath $script:LogPath -ErrorCount $ErrorCount -WarningCount $WarningCount -TotalRuntime $totalRuntime