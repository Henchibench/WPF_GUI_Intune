function Connect-ToGraph {
    try {
        # Update status
        $statusBar.Text = "Connecting to Microsoft Graph..."
        Show-Progress $true
        
        # Import necessary modules
        Import-Module Microsoft.Graph.Authentication
        
        # Connect to Microsoft Graph
        Connect-MgGraph
        
        # Update UI
        $connectionStatus.Text = "Connected: $((Get-MgContext).Account)"
        $global:connectedToGraph = $true
        $connectButton.Content = "Disconnect"
        $refreshButton.IsEnabled = $true
        $statusBar.Text = "Connected. Retrieving data..."
        
        # Enable UI elements for interaction
        $refreshButton.IsEnabled = $true
        $searchButton.IsEnabled = $true
        $searchTextBox.IsEnabled = $true
        $preloadButton.IsEnabled = $true
        
        # Initial data load
        FetchIntuneData
        
        # Automatically pre-load assignments after connection
        $statusBar.Text = "Starting automatic assignment pre-loading..."
        PreLoad-PolicyAssignments
    }
    catch {
        $connectionStatus.Text = "Error connecting: $($_.Exception.Message)"
        $statusBar.Text = "Error connecting"
        Write-Error "Error connecting to Microsoft Graph: $_"
    }
    finally {
        Show-Progress $false
    }
} 