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
# Define the CSV file name
$csvFileName = ".\Appliances_List\Appliances_List.csv"
# Define the parent directory of the CSV file
$parentDirectory = Split-Path -Path $scriptDirectory -Parent
# Create the full path to the CSV file
$csvFilePath = Join-Path -Path $parentDirectory -ChildPath $csvFileName
# Define the path to the credential folder
$credentialFolder = Join-Path -Path $parentDirectory -ChildPath "Credential"
# Task 2: import Appliances list from the CSV file.
Write-Host "`n$Spaces$($taskNumber). Importing Appliances list from the CSV file:`n" -ForegroundColor Magenta
# Import Appliances list from CSV file
$Appliances = Import-Csv -Path $csvFilePath
# Confirm that the CSV file was imported successfully
if ($Appliances) {
    # Get the total number of appliances
    $totalAppliances = $Appliances.Count
    # Log the total number of appliances
    Write-Log -Message "There are $totalAppliances appliances in the CSV file." -Level "Info" -NoConsoleOutput
    # Display if the CSV file was imported successfully
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The CSV file was imported successfully." -ForegroundColor Green
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
# increment $script:taskNumber after the function call
$script:taskNumber++
# Task 3: Check if credential folder exists
Write-Host "`n$Spaces$($taskNumber). Checking for credential folder:`n" -ForegroundColor Magenta
# Log the task
Write-Log -Message "Checking for credential folder." -Level "Info" -NoConsoleOutput
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
# increment $script:taskNumber after the function call
$script:taskNumber++
# Task 4: Check CSV & Excel Folders exists.
Write-Host "`n$Spaces$($taskNumber). Check CSV & Excel Folders exists:`n" -ForegroundColor Magenta
# Check if the credential file exists
if (-not (Test-Path -Path $credentialFile)) {
    # Prompt the user to enter their login and password
    $credential = Get-Credential -Message "Please enter your login and password."
    # Save the credential to the credential file
    $credential | Export-Clixml -Path $credentialFile
}
else {
    # Load the credential from the credential file
    $credential = Import-Clixml -Path $credentialFile
}
# Initialize arrays
$allLocalUsers = @()
$allLdapGroups = @()
# Define the directories for the CSV and Excel files
$csvDir = Join-Path -Path $script:ReportsDir -ChildPath 'CSV'
$excelDir = Join-Path -Path $script:ReportsDir -ChildPath 'Excel'
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
# Define the path to the CSV and Excel files for local users and LDAP groups
$localUsersCsvPath = Join-Path -Path $csvDir -ChildPath 'LocalUsers.csv'
$ldapGroupsCsvPath = Join-Path -Path $csvDir -ChildPath 'LdapGroups.csv'
$localUsersExcelPath = Join-Path -Path $excelDir -ChildPath 'LocalUsers.xlsx'
$ldapGroupsExcelPath = Join-Path -Path $excelDir -ChildPath 'LdapGroups.xlsx'
$combinedUsersExcelPath = Join-Path -Path $excelDir -ChildPath 'CombinedUsers.xlsx'
# increment $script:taskNumber after the function call
$script:taskNumber++
# Task 5: Collecting user details from each appliance
Write-Host "`n$Spaces$($taskNumber). Collecting user details from each appliance:`n" -ForegroundColor Magenta
# Loop through each appliance
foreach ($appliance in $Appliances) {
    # Convert the FQDN to uppercase
    $fqdn = $appliance.FQDN.ToUpper()
    # Check for existing sessions and disconnect them
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
    Connect-OVMgmt -Hostname $fqdn -Credential $credential *> $null
    Write-Host "`t1- Successfully connected to:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn" -ForegroundColor Green
    Write-Log -Message "Successfully connected to: $fqdn" -Level "OK" -NoConsoleOutput
    # Collect user details
    Write-Host "`t2- Collecting user details from:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn"  -ForegroundColor Green
    $users = Get-OVUser | ForEach-Object {
        $_ | Add-Member -NotePropertyName 'Type of user' -NotePropertyValue ($_.permissions | ForEach-Object { $_.roleName }) -PassThru
    }
    $allLocalUsers += $users
    # Collect LDAP group details
    Write-Host "`t3- Collecting LDAP group details from:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn" -ForegroundColor Green
    $ldapGroups = Get-OVLdapGroup | ForEach-Object {
        $_ | Add-Member -NotePropertyName 'Type of user' -NotePropertyValue ($_.permissions | ForEach-Object { $_.roleName }) -PassThru
    }
    $allLdapGroups += $ldapGroups
    # Generate reports
    Write-Host "`t4- Generating report for:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn" -ForegroundColor Green
    # Disconnect from the appliance
    Disconnect-OVMgmt -Hostname $fqdn
    Write-Host "`t5- Successfully disconnected from:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn`n" -ForegroundColor Green
    Write-Log -Message "Successfully disconnected from $fqdn" -Level "OK" -NoConsoleOutput
}
# increment $script:taskNumber after the function call
$script:taskNumber++
# Task 6: Exporting user details to CSV and Excel files
Write-Host "`n$Spaces$($taskNumber). Exporting user details to CSV and Excel files:`n" -ForegroundColor Magenta
# Export the local users to an Excel file with timestamp
$localUsersExcelPath = Join-Path -Path $excelDir -ChildPath "LocalUsers_$((Get-Date).ToString('yyyyMMdd-HHmmss')).xlsx"
$allLocalUsers | Export-Excel -Path $localUsersExcelPath
# Export the LDAP groups to an Excel file with timestamp
$ldapGroupsExcelPath = Join-Path -Path $excelDir -ChildPath "LdapGroups_$((Get-Date).ToString('yyyyMMdd-HHmmss')).xlsx"
$allLdapGroups | Export-Excel -Path $ldapGroupsExcelPath
# Export the combined users to an Excel file with timestamp
$combinedUsersExcelPath = Join-Path -Path $excelDir -ChildPath "CombinedUsers_$((Get-Date).ToString('yyyyMMdd-HHmmss')).xlsx"
# Assign the local users and LDAP groups to variables
$allLocalUsersCsv = $allLocalUsers
$allLdapGroupsCsv = $allLdapGroups
# Export the local users to a CSV file, creating a new file named LocalUsers+Time.csv in the CSV Directory
$localUsersCsvPath = Join-Path -Path $csvDir -ChildPath "LocalUsers_$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"
$allLocalUsersCsv | Export-Csv -Path $localUsersCsvPath -NoTypeInformation
# Export the LDAP groups to a CSV file, creating a new file named LdapGroups+Time.csv in the CSV Directory
$ldapGroupsCsvPath = Join-Path -Path $csvDir -ChildPath "LdapGroups_$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv"
$allLdapGroupsCsv | Export-Csv -Path $ldapGroupsCsvPath -NoTypeInformation
# Select specific properties from local users and add LDAP group-specific properties with default values
$selectedLocalUsers = $allLocalUsersCsv | Select-Object ApplianceConnection, type, category, userName, fullName, 'Type of user', @{Name = 'loginDomain'; Expression = { 'Local' } }, @{Name = 'egroup'; Expression = { 'N/A' } }, @{Name = 'directoryType'; Expression = { 'User' } }, uri
# Select specific properties from LDAP groups and add local user-specific properties with default values
$selectedLdapGroups = $allLdapGroupsCsv | Select-Object ApplianceConnection, type, category, @{Name = 'userName'; Expression = { 'N/A' } }, @{Name = 'fullName'; Expression = { 'N/A' } }, 'Type of user', loginDomain, egroup, directoryType, uri
# Combine all local users and LDAP groups into a single array
$combinedUsers = $selectedLocalUsers + $selectedLdapGroups
# Define Close-ExcelFile function to close the Excel file if it is open, if not open says it at console
function Close-ExcelFile {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ExcelFilePath
    )
    try {
        # Try to open the file in ReadWrite mode
        $fileStream = [System.IO.File]::Open($ExcelFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($fileStream) {
            $fileStream.Close()
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "The Excel file was already closed.`n" -ForegroundColor Yellow
            Write-Log -Message "The Excel file was already closed." -Level "Info" -NoConsoleOutput
        }
    }
    catch {
        # If an exception is thrown, the file is open
        $excelFile = Get-Process | Where-Object { $_.MainWindowTitle -like "*Excel*" }
        if ($excelFile) {
            # Close the Excel file
            $excelFile | Stop-Process -Force
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "The Excel file was open and has been closed.`n" -ForegroundColor Green
            Write-Log -Message "The Excel file was open and has been closed." -Level "OK" -NoConsoleOutput
        }
    }
}
# Close the Excel file if it is open
Close-ExcelFile -ExcelFilePath $combinedUsersExcelPath
# Add a delay to ensure the Excel file is closed before exporting the data
Start-Sleep -Seconds 5
# Sort the combined users by ApplianceConnection and then by userName
$sortedCombinedUsers = $combinedUsers | Sort-Object ApplianceConnection, type
# Export the data to Excel
$excel = $sortedCombinedUsers | Export-Excel -Path $combinedUsersExcelPath `
    -ClearSheet `
    -AutoSize `
    -AutoFilter `
    -FreezeTopRow `
    -WorksheetName "CombinedUsers" `
    -TableStyle "Dark11" `
    -PassThru
# Check if the Excel file was created successfully
if ($excel) {
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The Excel file was created successfully.`n" -ForegroundColor Green
    Write-Log -Message "The Excel file was created successfully." -Level "OK" -NoConsoleOutput
}
else {
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Failed to create the Excel file.`n" -ForegroundColor Red
    Write-Log -Message "Failed to create the Excel file." -Level "Error" -NoConsoleOutput
}
# Save and close the Excel package
$excel.Save()
$excel.Dispose()
# Increment $script:taskNumber after the function call
$script:taskNumber++
# Task 7: Script execution completed successfully
# write a message to the console indicating a summary of the script execution
Write-Host "`n$Spaces$($taskNumber). Summary of script execution.`n" -ForegroundColor Magenta
# Just before calling Complete-Logging
$endTime = Get-Date
$totalRuntime = $endTime - $startTime
# Call Complete-Logging at the end of the script
Complete-Logging -LogPath $script:LogPath -ErrorCount $ErrorCount -WarningCount $WarningCount -TotalRuntime $totalRuntime
