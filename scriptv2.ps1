# Get the computer's hostname
try {
    $hostname = hostname
    $confirmation = Read-Host "Do you want to change hostname (current hostname: $hostname) (y/n)"
    if ($confirmation -eq 'y' -or $confirmation -eq 'Y') {
        $hostname = Read-Host 'Enter hostname for this computer'
        Rename-Computer -NewName "$hostname"
    }
}
catch {
    Write-Error "Failed to get hostname. $_.Exception.Message"
    exit 1
}

# Get the computer's IP address
$gateway = '172.17.2.254'
try {
    $ipaddress = Get-NetIPConfiguration | Where-Object {$_.IPv4DefaultGateway.NextHop -eq $gateway} | Select-Object -ExpandProperty IPv4Address | Select-Object -ExpandProperty IPAddress
}
catch {
    Write-Error "Failed to get IP address. $_.Exception.Message"
    exit 1
}

# Get the computer's MAC address
try {
    $macaddress = Get-WmiObject win32_networkadapterconfiguration | Where-Object {$_.IPAddress -eq $ipaddress} | Select-Object -ExpandProperty macaddress
 }
catch {
    Write-Error "Failed to get MAC address. $_.Exception.Message"
    exit 1
}

# Prompt the user to enter a description for the computer
try {
    $description = Read-Host 'Enter a description for the computer'
}
catch {
    Write-Error "Failed to get computer description. $_.Exception.Message"
    exit 1
}

# Create a custom object with the computer information
$computer = [PSCustomObject]@{
    Hostname    = $hostname
    IPAddress   = $ipaddress
    MACAddress  = $macaddress
    Description = $description
}

# Define the path and name of the Excel file to append the information to
$excelFilePath = Join-Path -Path $PSScriptRoot -ChildPath "ComputerInfo.xlsx"

# Check if the Excel file exists, and create it if it doesn't
if (-not (Test-Path $excelFilePath)) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "ComputerInfo"
    # Define the headers for the Excel file
    $headers = "Hostname", "IPAddress", "MACAddress", "Description"
    $worksheet.Range("A1:D1").Value2 = $headers # Add headers to first row
    $workbook.SaveAs($excelFilePath)
    $excel.Quit()
}

# Check if the Excel application is already running and get the running instance
$excel = Get-Process "Excel" -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowTitle -like "*ComputerInfo.xlsx*"} | Select-Object -First 1

# If $excel is null, create a new instance of Excel and open the file
if (!$excel) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($excelFilePath)
}
# If $excel is not null, assume the Excel file has already been opened, and use the existing instance
else {
    $workbook = $excel.Workbooks.Item($excelFilePath)
}

# Select the worksheet where you want to append the information
$worksheet = $workbook.Worksheets.Item("ComputerInfo")

# Get the last row in the worksheet
$lastRow = $worksheet.UsedRange.Rows.Count + 1

# Append the computer information to the next empty row in the worksheet
$worksheet.Cells.Item($lastRow, 1) = $computer.Hostname
$worksheet.Cells.Item($lastRow, 2) = $computer.IPAddress
$worksheet.Cells.Item($lastRow, 3) = $computer.MACAddress
$worksheet.Cells.Item($lastRow, 4) = $computer.Description

# Save and close the Excel file
$workbook.Save()
$workbook.Close()
$excel.Quit()

Exit