# Import the Excel PowerShell module
Import-Module -Name ImportExcel

# Get the computer's hostname
try {
    $hostname = hostname
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
    $macaddress = (Get-NetAdapter | Where-Object {$_.InterfaceAlias -eq 'Wi-Fi'}).MacAddress
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
$excelFilePath = 'ComputerInfo.xlsx'

# Append the computer information to the Excel file
try {
    $computer | Export-Excel -Path $excelFilePath -WorksheetName 'ComputerInfo' -AutoSize -Append -TableStyle 'Medium15' -ErrorAction Stop
}
catch {
    Write-Error "Failed to append computer information to Excel file. $_.Exception.Message"
    exit 1
}