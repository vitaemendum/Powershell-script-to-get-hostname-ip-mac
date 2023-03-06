# Import the Excel module
Import-Module -Name ImportExcel
Import-Module ActiveDirectory

# Create a new Excel workbook and select the first worksheet
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Set up the column headers
$worksheet.Cells.Item(1,1) = "Computer"
$worksheet.Cells.Item(1,2) = "IP Address"
$worksheet.Cells.Item(1,3) = "MAC Address"

# Get the list of computers and network adapters
$computers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name
$networkAdapters = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.IPAddress -ne $null -and $_.IPEnabled -eq $true}

# Loop through each computer and retrieve the network adapter information
$row = 2
foreach ($computer in $computers) {
    $networkAdapter = $networkAdapters | Where-Object {$_.PSComputerName -eq $computer}
    if ($networkAdapter) {
        $ipAddress = $networkAdapter.IPAddress[0]
        $macAddress = $networkAdapter.MACAddress
    } else {
        $ipAddress = "Not found"
        $macAddress = "Not found"
    }
    # Write the data to the Excel worksheet
    $worksheet.Cells.Item($row,1) = $computer
    $worksheet.Cells.Item($row,2) = $ipAddress
    $worksheet.Cells.Item($row,3) = $macAddress
    $row++
}

# Autofit the columns
$range = $worksheet.UsedRange
$range.EntireColumn.AutoFit() | Out-Null

# Save the workbook and close Excel
$workbook.SaveAs("C:\Users\MATAS\Desktop\file.xlsx")
$excel.Quit()
