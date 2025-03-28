# Check for parameters
param (
    [switch]$SkipModuleCheck,
    [switch]$InstallMissing,
    [switch]$UpdateModules,
    [switch]$Force
)

# Main launcher script for Intune Explorer

# Define the root directory relative to this script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcPath = Join-Path -Path $scriptPath -ChildPath "src"

# Import module management functions
. "$srcPath\ModuleManager.ps1"

# Handle module updates if requested
if ($UpdateModules) {
    Update-RequiredModules -Force:$Force
    exit
}

# Skip module check if requested
if (-not $SkipModuleCheck) {
    # Test for required modules
    if (-not (Test-RequiredModules -InstallMissing:$InstallMissing -Force:$Force)) {
        Write-Host "`nRequired modules are missing. Please run the script with -InstallMissing parameter to install them." -ForegroundColor Red
        Write-Host "Example: .\IntuneExplorer.ps1 -InstallMissing" -ForegroundColor Yellow
        exit
    }
}
else {
    Write-Host "Running in test mode without checking for required modules." -ForegroundColor Yellow
}

# Load the main application
$mainScript = Join-Path -Path $srcPath -ChildPath "IntuneExplorer.ps1"
if (Test-Path $mainScript) {
    # Execute the main script
    if ($SkipModuleCheck) {
        & $mainScript -TestMode
    }
    else {
        & $mainScript
    }
}
else {
    Write-Host "Could not find the main application script at: $mainScript" -ForegroundColor Red
    Write-Host "Please make sure all files are in the correct location." -ForegroundColor Red
    exit
} 