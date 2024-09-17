<#
  .SYNOPSIS
    Installs QuickBooks Desktop.
  .DESCRIPTION
    Installs the specified version of QuickBooks Desktop based on user input in a categorized and user-friendly manner with detailed error handling.
#>

param(
  [String]$Cache
)

Function Confirm-SystemCheck {
  $CurrentUserSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
  if ($CurrentUserSID -eq 'S-1-5-18') {
    Write-Warning 'This script cannot run as SYSTEM. Please run as admin.'
    exit 1
  }
}

Function Install-XPSDocumentWriter {
  Write-Verbose "Checking Microsoft XPS Document Writer installation status..."
  $XPSFeature = Get-WindowsOptionalFeature -Online | Where-Object { $_.FeatureName -eq 'Printing-XPSServices-Features' }
  if ($XPSFeature.State -eq 'Disabled') {
    try {
      Write-Output "`nInstalling required PDF components (Microsoft XPS Document Writer)..."
      Enable-WindowsOptionalFeature -Online -FeatureName 'Printing-XPSServices-Features' -All -NoRestart | Out-Null
      Write-Output 'XPS Document Writer installation complete.'
    }
    catch {
      Write-Error "Unable to install Microsoft XPS Document Writer feature. Error: $_"
    }
  }
  else {
    Write-Output "XPS Document Writer is already installed."
  }
}

Function Install-QuickBooks {
  param(
    [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
    [PSObject]$QuickBooks,
    [String]$LicenseNumber,
    [String]$ProductNumber
  )

  Write-Verbose "Starting QuickBooks installation for $($QuickBooks.Name)..."
  
  $Exe = ($QuickBooks.URL -Split '/')[-1]
  $Installer = Join-Path -Path $env:TEMP -ChildPath $Exe
  if (Test-Path ("$Cache\$Exe")) { $CacheInstaller = Join-Path -Path $Cache -ChildPath $Exe }

  try {
    if ($CacheInstaller) { 
      Write-Output "`nCopying $($QuickBooks.Name) installer from cache..."
      Copy-Item -Path $CacheInstaller -Destination $Installer -Force
    }
    else { 
      Write-Output "`nDownloading $($QuickBooks.Name) installer..."
      Invoke-WebRequest -Uri $QuickBooks.URL -OutFile $Installer -ErrorAction Stop
    }
    Write-Output 'Starting installation...'
    Start-Process -Wait -NoNewWindow -FilePath $Installer -ArgumentList "-s -a QBMIGRATOR=1 MSICOMMAND=/s QB_PRODUCTNUM=$ProductNumber QB_LICENSENUM=$LicenseNumber" -ErrorAction Stop
    Write-Output 'QuickBooks installation complete.'
  }
  catch { 
    Write-Error "Error installing $($QuickBooks.Name): $_"
  }
  finally { 
    if (Test-Path $Installer) {
      Remove-Item $Installer -Force -ErrorAction Ignore
    }
  }
}

Function Install-ToolHub {
  Write-Verbose "Installing QuickBooks ToolHub..."
  $ToolHubURL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/QBFDT/QuickBooksToolHub.exe'
  $Exe = ($ToolHubURL -Split '/')[-1]
  $Installer = Join-Path -Path $env:TEMP -ChildPath $Exe
  if (Test-Path ("$Cache\$Exe")) { $CacheInstaller = Join-Path -Path $Cache -ChildPath $Exe }

  try {
    if ($CacheInstaller) {
      Write-Output "`nCopying ToolHub installer from cache..."
      Copy-Item -Path $CacheInstaller -Destination $Installer -Force
    }
    else {
      Write-Output "`nDownloading ToolHub installer..."
      Invoke-WebRequest -Uri $ToolHubURL -OutFile $Installer -ErrorAction Stop
    }
    Write-Output 'Starting ToolHub installation...'
    Start-Process -Wait -NoNewWindow -FilePath $Installer -ArgumentList '/S /v/qn' -ErrorAction Stop
    Write-Output 'ToolHub installation complete.'
  }
  catch {
    Write-Error "Error installing ToolHub: $_"
  }
  finally {
    if (Test-Path $Installer) {
      Remove-Item $Installer -Force -ErrorAction Ignore
    }
  }
}

# Enable Verbose output
$VerbosePreference = 'Continue'

# Abort if running as SYSTEM
Confirm-SystemCheck

# Adjust PowerShell settings
$ProgressPreference = 'SilentlyContinue'
if ([Net.ServicePointManager]::SecurityProtocol -notcontains 'Tls12' -and [Net.ServicePointManager]::SecurityProtocol -notcontains 'Tls13') {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

# Prompt user for License Number
$LicenseNumber = Read-Host "Enter your QuickBooks License Number (e.g., 0000-0000-0000-000)"

# Categorized list of available QuickBooks versions
$QuickBooksCategories = @{
  'Pro' = @(
    [PSCustomObject]@{Name = 'QuickBooks Pro 2023'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2023/Latest/QuickBooksProSub2023.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Pro 2022'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2022/Latest/QuickBooksProSub2022.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Pro 2021'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2021/Latest/QuickBooksPro2021.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Pro 2020'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2020/Latest/QuickBooksPro2020.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Pro 2019'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2019/Latest/QuickBooksPro2019.exe'; }
  )
  'Premier' = @(
    [PSCustomObject]@{Name = 'QuickBooks Premier 2023'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2023/Latest/QuickBooksPremierSub2023.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Premier 2022'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2022/Latest/QuickBooksPremier2022.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Premier 2021'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2021/Latest/QuickBooksPremier2021.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Premier 2020'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2020/Latest/QuickBooksPremier2020.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Premier 2019'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2019/Latest/QuickBooksPremier2019.exe'; }
  )
  'Enterprise' = @(
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 24'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2024/Latest/QuickBooksEnterprise24.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 23'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2023/Latest/QuickBooksEnterprise23.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 22'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2022/Latest/QuickBooksEnterprise22.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 21'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2021/Latest/QuickBooksEnterprise21.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 20'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2020/Latest/QuickBooksEnterprise20.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 19'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2019/Latest/QuickBooksEnterprise19.exe'; }
  )
  'Accountant' = @(
    [PSCustomObject]@{Name = 'QuickBooks Pro 2019 - Accountant'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2019/Latest/QuickBooksPro2019.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 23 - Accountant'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2023/Latest/QuickBooksEnterprise23.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 22 - Accountant'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2022/Latest/QuickBooksEnterprise22.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 20 - Accountant'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2020/Latest/QuickBooksEnterprise20.exe'; }
    [PSCustomObject]@{Name = 'QuickBooks Enterprise 19 - Accountant'; URL = 'https://dlm2.download.intuit.com/akdlm/SBD/QuickBooks/2019/Latest/QuickBooksEnterprise19.exe'; }
  )
}

# Display available categories
Write-Output "Select a QuickBooks category:"
Write-Output "1. Pro"
Write-Output "2. Premier"
Write-Output "3. Enterprise"
Write-Output "4. Accountant"

# Prompt user to select a category
$CategorySelection = [int](Read-Host "Enter the number corresponding to the QuickBooks category you wish to select")

# Map the user's selection to the appropriate category
$SelectedCategory = switch ($CategorySelection) {
  1 { 'Pro' }
  2 { 'Premier' }
  3 { 'Enterprise' }
  4 { 'Accountant' }
  Default { $null }
}

# Validate category selection
if ($SelectedCategory -and $QuickBooksCategories.ContainsKey($SelectedCategory)) {
  $SelectedVersions = $QuickBooksCategories[$SelectedCategory]

  # Display available versions in the selected category
  Write-Output "Select a version of QuickBooks ${SelectedCategory}:"
  $SelectedVersions | ForEach-Object { 
      [int]$index = [array]::IndexOf($SelectedVersions, $_) + 1
      Write-Output "$index. $($_.Name)"
  }

  # Prompt user to select a version
  $VersionSelection = [int](Read-Host "Enter the number corresponding to the QuickBooks version you wish to install")

  # Validate version selection
  if ($VersionSelection -gt 0 -and $VersionSelection -le $SelectedVersions.Count) {
    $SelectedVersion = $SelectedVersions[$VersionSelection - 1]

    # Prompt user for Product Number
    $ProductNumber = Read-Host "Enter your QuickBooks Product Number for $($SelectedVersion.Name)"

    # Install selected version
    Install-XPSDocumentWriter
    Install-QuickBooks -QuickBooks $SelectedVersion -LicenseNumber $LicenseNumber -ProductNumber $ProductNumber
  }
  else {
    Write-Warning "Invalid selection. Please run the script again and enter a valid number."
  }
}
else {
  Write-Warning "Invalid category selection. Please run the script again and enter a valid number."
}

# Ask if the user wants to install ToolHub
$InstallToolHub = Read-Host "Do you want to install QuickBooks ToolHub as well? (Y/N)"
if ($InstallToolHub -eq 'Y') { Install-ToolHub }
