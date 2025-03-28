# Module management utilities for Intune Explorer

function Test-RequiredModules {
    param (
        [Parameter(Mandatory = $false)]
        [switch]$InstallMissing = $false,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force = $false
    )
    
    # Define required modules
    $requiredModules = @(
        @{
            Name = "Microsoft.Graph.Intune"
            Description = "Microsoft Graph Intune module"
        },
        @{
            Name = "Microsoft.Graph.Beta.DeviceManagement"
            Description = "Microsoft Graph Beta Device Management module"
        },
        @{
            Name = "Microsoft.Graph.Beta.Users"
            Description = "Microsoft Graph Beta Users module"
        },
        @{
            Name = "Microsoft.Graph.Beta.DeviceManagement.Administration"
            Description = "Microsoft Graph Beta Device Management Administration module"
        },
        @{
            Name = "Microsoft.Graph.Beta.Applications"
            Description = "Microsoft Graph Beta Applications module"
        }
    )
    
    $missingModules = @()
    $installedModules = @()
    
    foreach ($module in $requiredModules) {
        $moduleInfo = Get-Module -ListAvailable -Name $module.Name
        if ($null -eq $moduleInfo) {
            $missingModules += $module
        }
        else {
            $installedModules += @{
                Name = $module.Name
                Version = $moduleInfo.Version
                Description = $module.Description
            }
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "`nMissing required modules:" -ForegroundColor Yellow
        foreach ($module in $missingModules) {
            Write-Host "- $($module.Name) ($($module.Description))" -ForegroundColor Yellow
        }
        
        if ($InstallMissing) {
            Write-Host "`nInstalling missing modules..." -ForegroundColor Cyan
            foreach ($module in $missingModules) {
                try {
                    Write-Host "Installing $($module.Name)..." -ForegroundColor Cyan
                    Install-Module -Name $module.Name -Force:$Force -Scope CurrentUser -ErrorAction Stop
                    Write-Host "Successfully installed $($module.Name)" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to install $($module.Name): $_" -ForegroundColor Red
                    return $false
                }
            }
        }
        else {
            Write-Host "`nTo install missing modules, run the script with -InstallMissing parameter." -ForegroundColor Yellow
            return $false
        }
    }
    
    if ($installedModules.Count -gt 0) {
        Write-Host "`nInstalled modules:" -ForegroundColor Green
        foreach ($module in $installedModules) {
            Write-Host "- $($module.Name) v$($module.Version)" -ForegroundColor Green
        }
    }
    
    return $true
}

function Update-RequiredModules {
    param (
        [Parameter(Mandatory = $false)]
        [switch]$Force = $false
    )
    
    Write-Host "Checking for module updates..." -ForegroundColor Cyan
    
    $requiredModules = @(
        "Microsoft.Graph.Intune",
        "Microsoft.Graph.Beta.DeviceManagement",
        "Microsoft.Graph.Beta.Users",
        "Microsoft.Graph.Beta.DeviceManagement.Administration",
        "Microsoft.Graph.Beta.Applications"
    )
    
    foreach ($module in $requiredModules) {
        try {
            $currentVersion = (Get-Module -ListAvailable -Name $module).Version
            Write-Host "`nChecking $module (Current: v$currentVersion)..." -ForegroundColor Cyan
            
            # Get the latest version from PSGallery
            $latestVersion = (Find-Module -Name $module).Version
            
            if ($latestVersion -gt $currentVersion) {
                Write-Host "Update available: v$latestVersion" -ForegroundColor Yellow
                $update = Read-Host "Do you want to update $module to v$latestVersion? (Y/N)"
                
                if ($update -eq "Y" -or $update -eq "y") {
                    Update-Module -Name $module -Force:$Force -ErrorAction Stop
                    Write-Host "Successfully updated $module to v$latestVersion" -ForegroundColor Green
                }
            }
            else {
                Write-Host "Module is up to date" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Error checking/updating $module : $_" -ForegroundColor Red
        }
    }
} 