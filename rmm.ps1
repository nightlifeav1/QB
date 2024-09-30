# Check and Start RMM.Agent Service
$service = Get-Service -Name 'RMM.Agent.exe' -ErrorAction SilentlyContinue
if ($service) {
    # Change the startup mode to Automatic
    $service | Set-Service -StartupType Automatic  
    # Check if the service is running
    $serviceStatus = Get-Service -Name 'RMM.Agent.exe'
    if ($serviceStatus.Status -eq "Running") {
        Write-Host "Service RMM.Agent is already running."
    } else {
        Write-Host "Service RMM.Agent is not running. Starting now..."
        Start-Service -Name 'RMM.Agent.exe'
    }
}

# Check and Start Connect Service
$service = Get-Service -Name 'Connect Service' -ErrorAction SilentlyContinue
if ($service) {
    # Change the startup mode to Automatic
    $service | Set-Service -StartupType Automatic  
    # Check if the service is running
    $serviceStatus = Get-Service -Name 'Connect Service'
    if ($serviceStatus.Status -eq "Running") {
        Write-Host "Service Connect Service is already running."
    } else {
        Write-Host "Service Connect Service is not running. Starting now..."
        Start-Service -Name 'Connect Service'
    }
}
