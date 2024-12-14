# Define the software names to ignore
$excludeList = @(
    "CData ODBC Driver for QuickBooks",
    "QODBC Driver for QuickBooks",
    "QuickBooks",
    "QuickBooks Advanced Reporting",
    "QuickBooks Desktop File Doctor",
    "QuickBooks Desktop migration tool",
    "QuickBooks Runtime Redistributable",
    "QuickBooks SDK 13.0",
    "QuickBooks Tool Hub",
    "QuickBooks_VC10_Debug",
    "Remote Connector for QuickBooks 2016"
)

# Get the installed software list
$installedSoftware = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | 
    Where-Object { $_.DisplayName -like "*QuickBooks*" } |
    Where-Object { $excludeList -notcontains $_.DisplayName } |
    Select-Object -ExpandProperty DisplayName

# Get the hostname of the system
$hostname = $env:COMPUTERNAME

# Format the output as hostname:qb1:qb2 (or other format as needed)
$formattedOutput = "${hostname}+" + ($installedSoftware -join "+")

# Output the result
Write-Host $formattedOutput
