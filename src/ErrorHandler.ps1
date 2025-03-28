# Error handling utilities for Intune Explorer

function Write-ErrorLog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null,
        
        [Parameter(Mandatory = $false)]
        [string]$LogPath = "$env:TEMP\IntuneExplorer_Error.log"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errorMessage = "[$timestamp] $Message"
    
    if ($null -ne $ErrorRecord) {
        $errorMessage += "`nException: $($ErrorRecord.Exception.Message)"
        $errorMessage += "`nCategory: $($ErrorRecord.CategoryInfo.Category)"
        $errorMessage += "`nTarget: $($ErrorRecord.TargetObject)"
        $errorMessage += "`nScript: $($ErrorRecord.InvocationInfo.ScriptName)"
        $errorMessage += "`nLine: $($ErrorRecord.InvocationInfo.ScriptLineNumber)"
    }
    
    $errorMessage += "`n----------------------------------------`n"
    
    # Write to log file
    try {
        Add-Content -Path $LogPath -Value $errorMessage -ErrorAction Stop
    }
    catch {
        # If writing to log fails, output to console
        Write-Host "Failed to write to error log: $_" -ForegroundColor Red
    }
    
    return $errorMessage
}

function Show-ErrorDialog {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
    )
    
    # Log the error
    $loggedMessage = Write-ErrorLog -Message $Message -ErrorRecord $ErrorRecord
    
    # Create error details for display
    $errorDetails = $Message
    if ($null -ne $ErrorRecord) {
        $errorDetails += "`n`nError: $($ErrorRecord.Exception.Message)"
    }
    
    # Show message box using WPF
    [System.Windows.MessageBox]::Show($errorDetails, $Title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}

function Invoke-SafeCommand {
    param (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorTitle = "Operation Failed",
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "An error occurred while performing the operation.",
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$ErrorHandler = $null,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$FinallyBlock = $null
    )
    
    try {
        & $ScriptBlock
    }
    catch {
        # Log and display error
        Write-ErrorLog -Message $ErrorMessage -ErrorRecord $_
        
        # Call custom error handler if provided
        if ($null -ne $ErrorHandler) {
            & $ErrorHandler $_
        }
        else {
            # Default error handling
            Show-ErrorDialog -Title $ErrorTitle -Message $ErrorMessage -ErrorRecord $_
        }
    }
    finally {
        if ($null -ne $FinallyBlock) {
            & $FinallyBlock
        }
    }
} 