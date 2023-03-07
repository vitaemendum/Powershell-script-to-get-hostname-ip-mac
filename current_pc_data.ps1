# Import the Excel PowerShell module
Import-Module -Name ImportExcel

# Get the computer's hostname
$hostname = hostname

# Get the computer's IP address
$gateway = '172.17.2.254'
$ipaddress = Get-NetIPConfiguration | Where-Object {$_.IPv4DefaultGateway.NextHop -eq $gateway} | Select-Object -ExpandProperty IPv4Address | Select-Object -ExpandProperty IPAddress

# Get the computer's MAC address
$macaddress = (Get-NetAdapter | Where-Object {$_.InterfaceAlias -eq 'Wi-Fi'}).MacAddress
# Prompt the user to enter a description for the computer
$description = Read-Host 'Enter a description for the computer'

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
$computer | Export-Excel -Path $excelFilePath -WorksheetName 'ComputerInfo' -AutoSize -Append -TableStyle 'Medium15'
