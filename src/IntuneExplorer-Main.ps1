#Requires -Modules Microsoft.Graph.Intune, Microsoft.Graph.Beta.DeviceManagement, Microsoft.Graph.Beta.Users, Microsoft.Graph.Beta.DeviceManagement.Administration, Microsoft.Graph.Beta.Applications, Microsoft.Graph.Beta.Groups

param (
    [switch]$TestMode
)

# Only require modules if not in test mode
if (-not $TestMode) {
    #Requires -Modules Microsoft.Graph.Intune, Microsoft.Graph.Beta.DeviceManagement, Microsoft.Graph.Beta.Users, Microsoft.Graph.Beta.DeviceManagement.Administration, Microsoft.Graph.Beta.Applications, Microsoft.Graph.Beta.Groups
}

# Add necessary assemblies for WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Threading

# Import helper scripts
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath\ErrorHandler.ps1"
. "$scriptPath\TreeViewHelper.ps1"

# Runspace setup for background operations
$sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 10, $sessionState, $Host)
$runspacePool.Open()

# Hashtable to track active runspace jobs
$global:runspaceJobs = @{}

# Function to execute scriptblocks asynchronously
function Start-AsyncJob {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Parameters = @{},
        
        [Parameter(Mandatory = $true)]
        [string]$JobName,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$OnComplete = {}
    )
    
    # Create PowerShell instance and assign runspace
    $powershell = [powershell]::Create().AddScript($ScriptBlock).AddParameters($Parameters)
    $powershell.RunspacePool = $runspacePool
    
    # Create async result object
    $asyncObject = New-Object PSObject -Property @{
        Powershell = $powershell
        AsyncResult = $powershell.BeginInvoke()
        JobName = $JobName
        OnComplete = $OnComplete
        StartTime = Get-Date
    }
    
    # Store in global tracking hashtable
    $global:runspaceJobs[$JobName] = $asyncObject
    
    # Return the job ID
    return $JobName
}

# Function to check for completed jobs
function Update-AsyncJobs {
    # Create a safe copy of the job keys to avoid collection modification during enumeration
    $jobKeys = @($global:runspaceJobs.Keys)
    $jobsToRemove = @()
    
    foreach ($jobName in $jobKeys) {
        $job = $global:runspaceJobs[$jobName]
        
        if ($job.AsyncResult.IsCompleted) {
            try {
                $result = $job.Powershell.EndInvoke($job.AsyncResult)
                
                # Invoke the OnComplete scriptblock on the UI thread if provided
                if ($null -ne $job.OnComplete) {
                    $window.Dispatcher.Invoke([action]{
                        & $job.OnComplete -Result $result
                    })
                }
            }
            catch {
                $errorMsg = "Error in async job '$($job.JobName)': $_"
                Write-Host $errorMsg -ForegroundColor Red
                $window.Dispatcher.Invoke([action]{
                    Write-Terminal "ERROR: $errorMsg"
                })
            }
            finally {
                $job.Powershell.Dispose()
                $jobsToRemove += $job.JobName
            }
        }
        elseif ((Get-Date) - $job.StartTime -gt [TimeSpan]::FromMinutes(10)) {
            # Timeout for jobs running too long (10 minutes)
            try {
                $job.Powershell.Stop()
                $errorMsg = "Job '$($job.JobName)' timed out after 10 minutes"
                Write-Host $errorMsg -ForegroundColor Yellow
                $window.Dispatcher.Invoke([action]{
                    Write-Terminal "WARNING: $errorMsg"
                })
            }
            catch {}
            finally {
                $job.Powershell.Dispose()
                $jobsToRemove += $job.JobName
            }
        }
    }
    
    # Remove completed or timed out jobs
    foreach ($jobName in $jobsToRemove) {
        $global:runspaceJobs.Remove($jobName)
    }
}

# Timer to regularly check for completed jobs
$jobTimer = New-Object System.Windows.Threading.DispatcherTimer
$jobTimer.Interval = [TimeSpan]::FromMilliseconds(100)
$jobTimer.Add_Tick({
    Update-AsyncJobs
})

# XAML for the UI
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Explorer" Height="650" Width="1000" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#5B9BD5"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#4A7DAA"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}" 
                                CornerRadius="5">
                            <ScrollViewer x:Name="PART_ContentHost"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Connection Panel -->
        <Border Grid.Row="0" Background="#F5F5F5" CornerRadius="10" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Microsoft Graph Connection" FontWeight="Bold" Margin="0,0,0,5"/>
                    <Button x:Name="ConnectButton" Content="Connect to Microsoft Graph" Width="200" HorizontalAlignment="Left"/>
                    <TextBlock x:Name="ConnectionStatus" Text="Not Connected" Margin="5"/>
                </StackPanel>
                <StackPanel Grid.Column="1" VerticalAlignment="Top">
                    <Button x:Name="RefreshButton" Content="Refresh Data" IsEnabled="False"/>
                    <Button x:Name="PreloadButton" Content="Pre-load Assignments" IsEnabled="False" ToolTip="Pre-loads policy and app assignments for faster searches"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Terminal Output Panel -->
        <Border Grid.Row="1" Background="#1E1E1E" CornerRadius="10" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Connection Output" Foreground="#E0E0E0" FontWeight="Bold" Margin="0,0,0,5"/>
                <ScrollViewer x:Name="TerminalScrollViewer" Grid.Row="1" Height="100" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                    <TextBox x:Name="TerminalOutput" Background="#1E1E1E" Foreground="#00FF00" FontFamily="Consolas" 
                             IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" BorderThickness="0"/>
                </ScrollViewer>
            </Grid>
        </Border>
        
        <!-- Search Panel -->
        <Border Grid.Row="2" Background="#F5F5F5" CornerRadius="10" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <TextBlock Text="Search Type:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <ComboBox x:Name="SearchTypeComboBox" Grid.Column="1" Margin="0,0,10,0" SelectedIndex="0">
                    <ComboBoxItem Content="Devices"/>
                    <ComboBoxItem Content="Groups"/>
                </ComboBox>
                <TextBlock Text="Search:" VerticalAlignment="Center" Grid.Row="1" Margin="0,0,10,0"/>
                <TextBox x:Name="SearchTextBox" Grid.Row="1" Grid.Column="1"/>
                <Button x:Name="SearchButton" Content="Search" Grid.Row="1" Grid.Column="2"/>
            </Grid>
        </Border>
        
        <!-- Tree View and Details Panel -->
        <Grid Grid.Row="3">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto" MinWidth="350"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- Tree View -->
            <Border Background="#F5F5F5" CornerRadius="10" Padding="10" Margin="0,0,5,0">
                <TreeView x:Name="IntuneTreeView" BorderThickness="0" Background="Transparent" Width="Auto" MinWidth="330" HorizontalAlignment="Stretch"/>
            </Border>
            
            <!-- Details Panel -->
            <Border Grid.Column="1" Background="#F5F5F5" CornerRadius="10" Padding="15" Margin="5,0,0,0">
                <ScrollViewer HorizontalScrollBarVisibility="Disabled">
                    <StackPanel x:Name="DetailsPanel" Background="Transparent"/>
                </ScrollViewer>
            </Border>
        </Grid>
        
        <!-- Status Bar -->
        <Border Grid.Row="4" Background="#F5F5F5" CornerRadius="10" Padding="5" Margin="0,10,0,0">
            <StackPanel Orientation="Horizontal">
                <TextBlock x:Name="StatusBar" Text="Ready" Padding="5"/>
                <ProgressBar x:Name="ProgressIndicator" Width="100" Height="15" Margin="10,0,0,0" Visibility="Collapsed"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

# Create a form
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Connect UI elements to variables
$connectButton = $window.FindName('ConnectButton')
$connectionStatus = $window.FindName('ConnectionStatus')
$refreshButton = $window.FindName('RefreshButton')
$searchTypeComboBox = $window.FindName('SearchTypeComboBox')
$searchTextBox = $window.FindName('SearchTextBox')
$searchButton = $window.FindName('SearchButton')
$intuneTreeView = $window.FindName('IntuneTreeView')
$detailsPanel = $window.FindName('DetailsPanel')
$statusBar = $window.FindName('StatusBar')
$progressIndicator = $window.FindName('ProgressIndicator')
$preloadButton = $window.FindName('PreloadButton')
$terminalOutput = $window.FindName('TerminalOutput')
$terminalScrollViewer = $window.FindName('TerminalScrollViewer')

# Initialize the terminal with a welcome message
$terminalOutput.AppendText("Welcome to Intune Explorer - Terminal view is active`n")
$terminalOutput.AppendText("Use the controls above to connect to Microsoft Graph and search for devices/groups`n")
$terminalOutput.AppendText("----------------------------------------`n")

# Global variables
$global:connectedToGraph = $false
$global:cachedData = @{
    Devices = $null
    Users = $null
    Apps = $null
    ConfigPolicies = $null
    CompliancePolicies = $null
    HealthScripts = $null
    Scripts = $null
    Groups = $null
    DeviceConfigs = $null
    # Add caches for assignments
    ConfigPolicyAssignments = @{} # Key: policyId, Value: assignments array
    CompliancePolicyAssignments = @{} # Key: policyId, Value: assignments array
    DeviceConfigAssignments = @{} # Key: configId, Value: assignments array
    AppAssignments = @{} # Key: appId, Value: assignments array 
}

# Function to show/hide progress indicator
function Show-Progress {
    param([bool]$show)
    
    if ($show) {
        $progressIndicator.Visibility = [System.Windows.Visibility]::Visible
        $progressIndicator.IsIndeterminate = $true
    }
    else {
        $progressIndicator.Visibility = [System.Windows.Visibility]::Collapsed
        $progressIndicator.IsIndeterminate = $false
    }
}

# Function to add a message to the terminal output
function Write-Terminal {
    param([string]$message)
    
    if ($null -eq $terminalOutput) { return }
    
    # Get the current dispatcher to ensure UI updates happen on the UI thread
    $dispatcher = $terminalOutput.Dispatcher
    
    $dispatcher.Invoke([Action]{
        $timestamp = Get-Date -Format "HH:mm:ss"
        $terminalOutput.AppendText("[$timestamp] $message`n")
        
        # Force the terminal to scroll to the end
        $terminalOutput.ScrollToEnd()
        
        # Also scroll the containing ScrollViewer to ensure visibility
        if ($null -ne $terminalScrollViewer) {
            $terminalScrollViewer.ScrollToEnd()
        }
    }, [System.Windows.Threading.DispatcherPriority]::Background)
}

# Function to clear the terminal output
function Clear-Terminal {
    if ($null -eq $terminalOutput) { return }
    
    $terminalOutput.Clear()
}

# Function to connect to Microsoft Graph
function Connect-ToGraph {
    try {
        # Clear terminal and update status
        Clear-Terminal
        $statusBar.Text = "Connecting to Microsoft Graph..."
        Show-Progress $true
        Write-Terminal "Starting connection to Microsoft Graph..."
        
        # Disable the connect button while connecting to prevent multiple clicks
        $connectButton.IsEnabled = $false
        
        # Define the connection script block that will run in the background
        $connectionScriptBlock = {
            try {
                # Import necessary modules
                Import-Module Microsoft.Graph.Authentication
                
                # Connect to Microsoft Graph
                Connect-MgGraph
                
                # Get the connected account
                $account = (Get-MgContext).Account
                
                # Return the account info
                return @{
                    Success = $true
                    Account = $account
                }
            }
            catch {
                return @{
                    Success = $false
                    Error = $_.Exception.Message
                    StackTrace = $_.ScriptStackTrace
                }
            }
        }
        
        # Define what happens when the async job completes
        $onCompleteAction = {
            param($Result)
            
            try {
                if ($Result.Success) {
                    # Update UI with successful connection
                    $account = $Result.Account
                    $connectionStatus.Text = "Connected: $account"
                    $global:connectedToGraph = $true
                    $connectButton.Content = "Disconnect"
                    $refreshButton.IsEnabled = $true
                    $statusBar.Text = "Connected. Retrieving data..."
                    Write-Terminal "Successfully connected to Microsoft Graph as: $account"
                    Write-Terminal "Connection established. Starting to retrieve Intune data..."
                    
                    # Enable UI elements for interaction
                    $refreshButton.IsEnabled = $true
                    $searchButton.IsEnabled = $true
                    $searchTextBox.IsEnabled = $true
                    $preloadButton.IsEnabled = $true
                    
                    # Initial data load (also asynchronous)
                    Write-Terminal "Fetching initial Intune data..."
                    FetchIntuneData
                }
                else {
                    # Update UI with error
                    $errorMessage = $Result.Error
                    $connectionStatus.Text = "Error connecting: $errorMessage"
                    $statusBar.Text = "Error connecting"
                    Write-Terminal "ERROR: Failed to connect to Microsoft Graph: $errorMessage"
                    Write-Terminal "Stack Trace: $($Result.StackTrace)"
                    Write-Error "Error connecting to Microsoft Graph: $errorMessage"
                    
                    # Re-enable the connect button to allow retry
                    $connectButton.Content = "Connect to Microsoft Graph"
                    Show-Progress $false
                }
            }
            finally {
                # Always re-enable the connect button
                $connectButton.IsEnabled = $true
            }
        }
        
        # Start the async job
        Write-Terminal "Initiating Graph connection - you may see a sign-in prompt..."
        Start-AsyncJob -ScriptBlock $connectionScriptBlock -JobName "GraphConnection" -OnComplete $onCompleteAction
    }
    catch {
        $errorMessage = $_.Exception.Message
        $connectionStatus.Text = "Error preparing connection: $errorMessage"
        $statusBar.Text = "Error connecting"
        Write-Terminal "ERROR: Failed to prepare Graph connection: $errorMessage"
        Write-Terminal "Stack Trace: $($_.ScriptStackTrace)"
        Write-Error "Error preparing Graph connection: $_"
        
        # Re-enable the connect button
        $connectButton.IsEnabled = $true
        Show-Progress $false
    }
}

# Function to disconnect from Microsoft Graph
function Disconnect-FromGraph {
    try {
        Write-Terminal "Disconnecting from Microsoft Graph..."
        Show-Progress $true
        Disconnect-MgGraph
        $connectionStatus.Text = "Not Connected"
        $global:connectedToGraph = $false
        $connectButton.Content = "Connect to Microsoft Graph"
        $refreshButton.IsEnabled = $false
        $statusBar.Text = "Disconnected from Microsoft Graph"
        Write-Terminal "Successfully disconnected from Microsoft Graph."
        
        # Clear the TreeView
        $intuneTreeView.Items.Clear()
        $detailsPanel.Children.Clear()
        Write-Terminal "UI cleared - application is ready for a new connection."
        
        # Clear cached data
        Write-Terminal "Clearing cached data..."
        $global:cachedData.Devices = $null
        $global:cachedData.Users = $null
        $global:cachedData.Apps = $null
        $global:cachedData.ConfigPolicies = $null
        $global:cachedData.CompliancePolicies = $null
        $global:cachedData.HealthScripts = $null
        $global:cachedData.Scripts = $null
        $global:cachedData.Groups = $null
        $global:cachedData.DeviceConfigs = $null
        $global:cachedData.ConfigPolicyAssignments = @{}
        $global:cachedData.CompliancePolicyAssignments = @{}
        $global:cachedData.DeviceConfigAssignments = @{}
        $global:cachedData.AppAssignments = @{}
        Write-Terminal "All cached data cleared successfully."
    }
    catch {
        $errorMessage = $_.Exception.Message
        $statusBar.Text = "Disconnection failed: $_"
        Write-Terminal "ERROR: Failed to disconnect: $errorMessage"
    }
    finally {
        Show-Progress $false
    }
}

# Function to refresh data
function Refresh-Data {
    if (-not $global:connectedToGraph) {
        $statusBar.Text = "Please connect to Microsoft Graph first"
        return
    }
    
    try {
        Show-Progress $true
        $statusBar.Text = "Refreshing data..."
        
        # Clear cached data
        $global:cachedData = @{
            Devices = $null
            Users = $null
            Groups = $null
            ConfigPolicies = $null
            CompliancePolicies = $null
            DeviceConfigs = $null
            Apps = $null
            ConfigPolicyAssignments = @{}
            CompliancePolicyAssignments = @{}
            DeviceConfigAssignments = @{}
            AppAssignments = @{}
        }
        
        # Reset preload button state
        $preloadButton.Content = "Pre-load Assignments"
        $preloadButton.IsEnabled = $true
        
        # Reload data
        FetchIntuneData
        
        # If there was a search performed, re-run it with the fresh data
        if (-not [string]::IsNullOrEmpty($searchTextBox.Text.Trim())) {
            Search-Item
        }
        
        $statusBar.Text = "Data refreshed successfully. Consider clicking 'Pre-load Assignments' for faster group searches."
    }
    catch {
        $statusBar.Text = "Error refreshing data: $_"
    }
    finally {
        Show-Progress $false
    }
}

# Function to fetch Intune data
function FetchIntuneData {
    if (-not $global:connectedToGraph) {
        return
    }

    try {
        Show-Progress $true
        $statusBar.Text = "Refreshing data..."
        Write-Terminal "Starting to fetch Intune data..."
        
        # Clear existing data
        $global:cachedData = @{
            Devices = $null
            Users = $null
            Groups = $null
            ConfigPolicies = $null
            CompliancePolicies = $null
            DeviceConfigs = $null
            Apps = $null
            ConfigPolicyAssignments = @{}
            CompliancePolicyAssignments = @{}
            DeviceConfigAssignments = @{}
            AppAssignments = @{}
        }
        Write-Terminal "Cleared existing cached data."
        
        # Reset preload button state
        $preloadButton.Content = "Pre-load Assignments"
        $preloadButton.IsEnabled = $true
        
        # Clear the treeview
        $intuneTreeView.Items.Clear()
        
        # Define the data fetching script block
        $fetchDataScriptBlock = {
            # Initialize results container
            $results = @{
                Groups = $null
                Devices = $null
                Users = $null
                Errors = @{}
            }
            
            try {
                # Fetch groups
                try {
                    $results.Groups = Get-MgGroup -All -ErrorAction Stop
                }
                catch {
                    $results.Errors["Groups"] = @{
                        Message = $_.Exception.Message
                        StackTrace = $_.ScriptStackTrace
                    }
                }
                
                # Fetch devices
                try {
                    $results.Devices = Get-MgBetaDeviceManagementManagedDevice -All -ErrorAction Stop
                }
                catch {
                    $results.Errors["Devices"] = @{
                        Message = $_.Exception.Message
                        StackTrace = $_.ScriptStackTrace
                    }
                }
                
                # Fetch users
                try {
                    $results.Users = Get-MgUser -All -ErrorAction Stop
                }
                catch {
                    $results.Errors["Users"] = @{
                        Message = $_.Exception.Message
                        StackTrace = $_.ScriptStackTrace
                    }
                }
            }
            catch {
                $results.Errors["General"] = @{
                    Message = $_.Exception.Message
                    StackTrace = $_.ScriptStackTrace
                }
            }
            
            return $results
        }
        
        # Define what happens when the async job completes
        $onCompleteAction = {
            param($Result)
            
            # Process and store the fetched data
            if ($Result.Groups) {
                $global:cachedData.Groups = $Result.Groups
                $groupCount = $Result.Groups.Count
                Write-Terminal "Successfully loaded $groupCount groups."
            }
            elseif ($Result.Errors.ContainsKey("Groups")) {
                $errorMessage = $Result.Errors["Groups"].Message
                Write-Terminal "ERROR: Failed to fetch groups: $errorMessage"
            }
            
            if ($Result.Devices) {
                $global:cachedData.Devices = $Result.Devices
                $deviceCount = $Result.Devices.Count
                Write-Terminal "Successfully loaded $deviceCount devices."
            }
            elseif ($Result.Errors.ContainsKey("Devices")) {
                $errorMessage = $Result.Errors["Devices"].Message
                Write-Terminal "ERROR: Failed to fetch devices: $errorMessage"
            }
            
            if ($Result.Users) {
                $global:cachedData.Users = $Result.Users
                $userCount = $Result.Users.Count
                Write-Terminal "Successfully loaded $userCount users."
            }
            elseif ($Result.Errors.ContainsKey("Users")) {
                $errorMessage = $Result.Errors["Users"].Message
                Write-Terminal "ERROR: Failed to fetch users: $errorMessage"
            }
            
            # Data loaded successfully
            $statusBar.Text = "Basic data loaded. Use the search box or click Pre-load Assignments."
            Write-Terminal "Basic data load complete. UI is now ready for use."
            
            # Wait a brief moment to ensure the UI thread is responsive
            Start-Sleep -Milliseconds 500
            
            # Automatically start pre-loading assignments in a new async job
            $window.Dispatcher.InvokeAsync({
                $statusBar.Text = "Starting automatic assignment pre-loading..."
                Write-Terminal "Starting automatic pre-loading of assignments..."
                PreLoad-PolicyAssignments
            })
            
            Show-Progress $false
        }
        
        # Start the async job
        Start-AsyncJob -ScriptBlock $fetchDataScriptBlock -JobName "FetchIntuneData" -OnComplete $onCompleteAction
    }
    catch {
        $errorMessage = $_.Exception.Message
        $statusBar.Text = "Error starting data refresh: $errorMessage"
        Write-Error "Error starting data refresh: $_"
        Write-Terminal "ERROR: Failed to start data refresh: $errorMessage"
        Show-Progress $false
    }
}

# Function to search for a device or group
function Search-Item {
    $searchTerm = $searchTextBox.Text.Trim()
    
    # Ensure we have a valid search type selection
    if ($null -eq $searchTypeComboBox.SelectedItem) {
        $statusBar.Text = "Please select a search type"
        Write-Terminal "ERROR: No search type selected"
        return
    }
    
    $searchType = $searchTypeComboBox.SelectedItem.Content.ToString()
    
    if ([string]::IsNullOrEmpty($searchTerm)) {
        $statusBar.Text = "Please enter a search term"
        Write-Terminal "ERROR: No search term entered"
        return
    }
    
    if (-not $global:connectedToGraph) {
        $statusBar.Text = "Please connect to Microsoft Graph first"
        Write-Terminal "ERROR: Not connected to Microsoft Graph"
        return
    }
    
    Show-Progress $true
    $statusBar.Text = "Searching for $($searchType): $searchTerm..."
    Write-Terminal "Searching for $searchType matching: '$searchTerm'..."
    
    try {
        # Clear the TreeView
        $intuneTreeView.Items.Clear()
        $detailsPanel.Children.Clear()
        
        switch ($searchType) {
            "Devices" {
                try {
                    # Ensure devices are fetched
                    if ($null -eq $global:cachedData.Devices) {
                        Write-Terminal "No cached devices found. Fetching devices first..."
                        FetchIntuneData
                    }
                    
                    # Find devices matching the search term
                    Write-Terminal "Searching through $($global:cachedData.Devices.Count) devices..."
                    $matchingDevices = $global:cachedData.Devices | Where-Object { 
                        $_.DeviceName -like "*$searchTerm*" -or 
                        $_.SerialNumber -like "*$searchTerm*" -or 
                        $_.Id -like "*$searchTerm*" 
                    }
                    
                    if ($null -eq $matchingDevices -or ($matchingDevices -is [array] -and $matchingDevices.Count -eq 0)) {
                        $statusBar.Text = "No matching devices found"
                        Write-Terminal "No devices found matching: '$searchTerm'"
                        return
                    }
                    
                    # Convert to array if single object
                    if (-not ($matchingDevices -is [array])) {
                        $matchingDevices = @($matchingDevices)
                    }
                    
                    # Display the matching devices in the TreeView
                    Write-Terminal "Adding $($matchingDevices.Count) matching devices to the tree view..."
                    foreach ($device in $matchingDevices) {
                        Add-DeviceToTreeView -TreeView $intuneTreeView -Device $device
                    }
                    
                    $statusBar.Text = "Found $($matchingDevices.Count) devices"
                    Write-Terminal "✅ Successfully found and displayed $($matchingDevices.Count) matching devices"
                }
                catch {
                    $errorMsg = "Error searching for devices: $_"
                    Write-Host $errorMsg -ForegroundColor Red
                    $statusBar.Text = $errorMsg
                    Write-Terminal "ERROR: $errorMsg"
                }
            }
            
            "Groups" {
                try {
                    # Ensure groups are fetched
                    if ($null -eq $global:cachedData.Groups) {
                        $statusBar.Text = "Fetching groups..."
                        Write-Terminal "No cached groups found. Fetching groups first..."
                        try {
                            $global:cachedData.Groups = Get-MgBetaGroup -All -ErrorAction Stop
                            Write-Terminal "Successfully fetched $($global:cachedData.Groups.Count) groups"
                        }
                        catch {
                            $errorMsg = "Failed to fetch groups: $_. Please check your permissions."
                            Write-Host $errorMsg -ForegroundColor Red
                            $statusBar.Text = $errorMsg
                            Write-Terminal "ERROR: $errorMsg"
                            Show-Progress $false
                            return
                        }
                    }
                    
                    if ($null -eq $global:cachedData.Groups) {
                        $statusBar.Text = "Unable to retrieve groups. Check your connection and permissions."
                        Write-Terminal "ERROR: Unable to retrieve groups. Check connection and permissions."
                        Show-Progress $false
                        return
                    }
                    
                    # Find groups matching the search term
                    Write-Terminal "Searching through $($global:cachedData.Groups.Count) groups..."
                    $matchingGroups = $global:cachedData.Groups | Where-Object {
                        $_.DisplayName -like "*$searchTerm*" -or
                        $_.Id -like "*$searchTerm*"
                    }
                    
                    if ($null -eq $matchingGroups -or ($matchingGroups -is [array] -and $matchingGroups.Count -eq 0)) {
                        $statusBar.Text = "No matching groups found"
                        Write-Terminal "No groups found matching: '$searchTerm'"
                        return
                    }
                    
                    # Convert to array if single object
                    if (-not ($matchingGroups -is [array])) {
                        $matchingGroups = @($matchingGroups)
                    }
                    
                    # Display the matching groups in the TreeView
                    Write-Terminal "Found $($matchingGroups.Count) matching groups. Adding to tree view..."
                    $groupsAdded = 0
                    foreach ($group in $matchingGroups) {
                        try {
                            # Keep a cleaner reference to the group before we try to add it
                            $groupName = if ($group.DisplayName) { $group.DisplayName } else { "Group " + $group.Id }
                            
                            # Explicitly discard null returns from Add-GroupToTreeView
                            $node = Add-GroupToTreeView -TreeView $intuneTreeView -Group $group
                            if ($node -ne $null) {
                                $groupsAdded++
                            } else {
                                Write-Host "Failed to add group '$groupName' to tree view" -ForegroundColor Yellow
                                Write-Terminal "WARNING: Failed to add group '$groupName' to tree view"
                            }
                        }
                        catch {
                            Write-Host "Error adding group '$groupName' to tree view: $_" -ForegroundColor Red
                            Write-Terminal "ERROR: Failed to add group '$groupName': $($_.Exception.Message)"
                            # Continue with the next group
                        }
                    }
                    
                    if ($groupsAdded -gt 0) {
                        $statusBar.Text = "Found $groupsAdded groups"
                        Write-Terminal "✅ Successfully added $groupsAdded of $($matchingGroups.Count) groups to the tree view"
                    }
                    else {
                        $statusBar.Text = "Found $($matchingGroups.Count) groups but none could be displayed. Check the console for errors."
                        Write-Host "Error: All groups found, but none could be displayed. You may need additional permissions." -ForegroundColor Red
                        Write-Terminal "ERROR: Found $($matchingGroups.Count) groups but none could be displayed. Check permissions."
                    }
                }
                catch {
                    $errorMsg = "Error searching for groups: $_"
                    Write-Host $errorMsg -ForegroundColor Red 
                    $statusBar.Text = $errorMsg
                    Write-Terminal "ERROR: $errorMsg"
                }
            }
        }
    }
    catch {
        $errorMsg = "Error in Search-Item: $_"
        Write-Host $errorMsg -ForegroundColor Red
        $statusBar.Text = $errorMsg
        Write-Terminal "ERROR: $errorMsg"
    }
    finally {
        Show-Progress $false
    }
}

# Function to handle TreeView selection
function Handle-TreeViewSelection {
    param($selectedItem)
    
    if ($null -eq $selectedItem) {
        return
    }
    
    try {
        $detailsPanel.Children.Clear()
        
        # Check if the selected item is a device node by looking at its properties
        if ($selectedItem.Tag -and ($selectedItem.Tag.PSObject.Properties.Name -contains 'DeviceName')) {
            # Device node selected
            $device = $selectedItem.Tag
            $detailsText = "Device Information:`n`n"
            $detailsText += "Name: $($device.DeviceName)`n"
            $detailsText += "Model: $($device.Model)`n"
            $detailsText += "Manufacturer: $($device.Manufacturer)`n"
            $detailsText += "Serial Number: $($device.SerialNumber)`n"
            $detailsText += "OS: $($device.OperatingSystem)`n"
            $detailsText += "OS Version: $($device.OsVersion)`n"
            $detailsText += "Compliance State: $($device.ComplianceState)`n"
            $detailsText += "Management State: $($device.ManagementState)`n"
            $detailsText += "Enrolled Date: $($device.EnrolledDateTime)`n"
            $detailsText += "Last Sync Date: $($device.LastSyncDateTime)`n"
            
            $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
            $detailsTextBlock.Text = $detailsText
            $detailsTextBlock.TextWrapping = "Wrap"
            $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
            $detailsTextBlock.Foreground = "#000000"
            $detailsTextBlock.FontSize = 12
            
            $detailsPanel.Children.Clear()
            $detailsPanel.Children.Add($detailsTextBlock)
        }
        # Check if the selected item is a group node
        elseif ($selectedItem.Tag -and ($selectedItem.Tag.PSObject.Properties.Name -contains 'DisplayName')) {
            # Group node selected
            $group = $selectedItem.Tag
            $detailsText = "Group Information:`n`n"
            $detailsText += "Name: $($group.DisplayName)`n"
            $detailsText += "Description: $($group.Description)`n"
            $detailsText += "ID: $($group.Id)`n"
            $detailsText += "Group Type: $(if ($group.GroupTypes) { $group.GroupTypes -join ", " } else { "Security" })`n"
            $detailsText += "Mail Enabled: $($group.MailEnabled)`n"
            $detailsText += "Security Enabled: $($group.SecurityEnabled)`n"
            $detailsText += "Created: $($group.CreatedDateTime)`n"
            $detailsText += "Last Modified: $($group.LastModifiedDateTime)`n"
            
            $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
            $detailsTextBlock.Text = $detailsText
            $detailsTextBlock.TextWrapping = "Wrap"
            $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
            $detailsTextBlock.Foreground = "#000000"
            $detailsTextBlock.FontSize = 12
            
            $detailsPanel.Children.Clear()
            $detailsPanel.Children.Add($detailsTextBlock)
        }
        elseif ($selectedItem.Tag -is [Hashtable]) {
            # Sub-nodes selected
            $nodeType = $selectedItem.Tag.Type
            
            # Process device-related nodes
            if ($nodeType -like "Device*") {
                $device = $selectedItem.Tag.Device
                
                switch ($nodeType) {
                    "DeviceProperties" {
                        # Create a grid for two-column layout
                        $sectionsGrid = New-Object System.Windows.Controls.Grid
                        $sectionsGrid.Margin = New-Object System.Windows.Thickness(0)
                        
                        # Define two columns with equal width
                        $column1 = New-Object System.Windows.Controls.ColumnDefinition
                        $column1.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
                        $column2 = New-Object System.Windows.Controls.ColumnDefinition
                        $column2.Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)
                        $sectionsGrid.ColumnDefinitions.Add($column1)
                        $sectionsGrid.ColumnDefinitions.Add($column2)

                        # Group important properties first
                        $importantProps = [ordered]@{
                            "Basic Information" = @(
                                @{ Name = "Device Name"; Value = $device.DeviceName },
                                @{ Name = "Model"; Value = $device.Model },
                                @{ Name = "Manufacturer"; Value = $device.Manufacturer },
                                @{ Name = "Serial Number"; Value = $device.SerialNumber },
                                @{ Name = "Compliance"; Value = $device.ComplianceState }
                            )
                            "Management Status" = @(
                                @{ Name = "Management State"; Value = $device.ManagementState },
                                @{ Name = "Enrollment Type"; Value = $device.DeviceEnrollmentType },
                                @{ Name = "Ownership"; Value = $device.ManagedDeviceOwnerType }
                            )
                            "User Information" = @()  # This will be populated later
                            "Important Dates" = @(
                                @{ Name = "Enrolled Date"; Value = if ($device.EnrolledDateTime) { [DateTime]::Parse($device.EnrolledDateTime).ToString("MM/dd/yyyy hh:mm:ss tt") } else { "N/A" } },
                                @{ Name = "Last Sync"; Value = if ($device.LastSyncDateTime) { [DateTime]::Parse($device.LastSyncDateTime).ToString("MM/dd/yyyy hh:mm:ss tt") } else { "N/A" } }
                            )
                            "Operating System" = @(
                                @{ Name = "OS"; Value = $device.OperatingSystem },
                                @{ Name = "OS Version"; Value = $device.OSVersion },
                                @{ Name = "Device Type"; Value = $device.DeviceType }
                            )
                        }

                        # Add user information if available
                        if ($device.UserId) {
                            try {
                                $user = Get-MgBetaUser -UserId $device.UserId
                                $importantProps["User Information"] = @(
                                    @{ Name = "Display Name"; Value = $user.DisplayName },
                                    @{ Name = "Email"; Value = $user.Mail },
                                    @{ Name = "UPN"; Value = $user.UserPrincipalName },
                                    @{ Name = "Job Title"; Value = $user.JobTitle },
                                    @{ Name = "Department"; Value = $user.Department },
                                    @{ Name = "Office"; Value = $user.OfficeLocation }
                                )
                            }
                            catch {
                                $importantProps["User Information"] = @(
                                    @{ Name = "Error"; Value = "Failed to fetch user information" }
                                )
                            }
                        }
                        else {
                            $importantProps["User Information"] = @(
                                @{ Name = "Status"; Value = "No user associated with this device" }
                            )
                        }

                        # Calculate how many rows we'll need (sections ÷ 2, rounded up)
                        $numberOfRows = [Math]::Ceiling($importantProps.Count / 2)
                        
                        # Pre-create all row definitions
                        for ($i = 0; $i -lt $numberOfRows; $i++) {
                            $rowDef = New-Object System.Windows.Controls.RowDefinition
                            $rowDef.Height = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Auto)
                            $sectionsGrid.RowDefinitions.Add($rowDef)
                        }

                        # Create sections in a two-column layout
                        $currentRow = 0
                        $currentColumn = 0
                        
                        foreach ($section in $importantProps.Keys) {
                            # Skip empty sections
                            if ($importantProps[$section].Count -eq 0 -and $section -ne "User Information") {
                                continue
                            }
                            
                            # Create a container for each section
                            $sectionContainer = New-Object System.Windows.Controls.StackPanel
                            $sectionContainer.Margin = New-Object System.Windows.Thickness(8)
                            
                            # Add section header
                            $sectionHeader = New-Object System.Windows.Controls.TextBlock
                            $sectionHeader.Text = $section
                            $sectionHeader.FontWeight = "Bold"
                            $sectionHeader.FontSize = 16
                            $sectionHeader.Foreground = "#2B579A"
                            $sectionHeader.Margin = New-Object System.Windows.Thickness(0, 0, 0, 8)
                            $sectionContainer.Children.Add($sectionHeader)

                            # Create a border for the section content
                            $sectionBorder = New-Object System.Windows.Controls.Border
                            $sectionBorder.Background = "#FFFFFF"
                            $sectionBorder.CornerRadius = New-Object System.Windows.CornerRadius(5)
                            $sectionBorder.Padding = New-Object System.Windows.Thickness(15)
                            $sectionBorder.Height = 180  # Set fixed height for all sections
                            $sectionBorder.MinWidth = 300  # Set minimum width for sections

                            # Create ScrollViewer for section content
                            $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
                            $scrollViewer.VerticalScrollBarVisibility = "Auto"
                            $scrollViewer.HorizontalScrollBarVisibility = "Disabled"
                            $scrollViewer.Margin = New-Object System.Windows.Thickness(0)

                            # Add section properties
                            $sectionStack = New-Object System.Windows.Controls.StackPanel
                            $sectionStack.Margin = New-Object System.Windows.Thickness(0)
                            foreach ($prop in $importantProps[$section]) {
                                $propGrid = New-Object System.Windows.Controls.Grid
                                $propGrid.Margin = New-Object System.Windows.Thickness(0, 2, 0, 2)

                                # Define columns for label and value
                                $propGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
                                $propGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition))
                                $propGrid.ColumnDefinitions[0].Width = New-Object System.Windows.GridLength(150)
                                $propGrid.ColumnDefinitions[1].Width = New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star)

                                # Add label
                                $labelBlock = New-Object System.Windows.Controls.TextBlock
                                $labelBlock.Text = "$($prop.Name):"
                                $labelBlock.FontWeight = "SemiBold"
                                $labelBlock.Margin = New-Object System.Windows.Thickness(0, 0, 10, 0)
                                [System.Windows.Controls.Grid]::SetColumn($labelBlock, 0)
                                $propGrid.Children.Add($labelBlock)

                                # Add value with special formatting for Compliance
                                $valueBlock = New-Object System.Windows.Controls.TextBlock
                                $value = if ($null -eq $prop.Value -or $prop.Value -eq "") { "N/A" } else { $prop.Value.ToString() }
                                $valueBlock.Text = $value
                                
                                if ($prop.Name -eq "Compliance") {
                                    $complianceValue = $value.ToString().ToLower()
                                    switch ($complianceValue) {
                                        "compliant" { $valueBlock.Foreground = "#107C10" }  # Green
                                        "noncompliant" { $valueBlock.Foreground = "#D83B01" }  # Red
                                        default { $valueBlock.Foreground = "#666666" }  # Gray
                                    }
                                    $valueBlock.FontWeight = "Bold"
                                }
                                
                                $valueBlock.TextWrapping = "Wrap"
                                [System.Windows.Controls.Grid]::SetColumn($valueBlock, 1)
                                $propGrid.Children.Add($valueBlock)

                                $sectionStack.Children.Add($propGrid)
                            }

                            $sectionBorder.Child = $sectionStack
                            $scrollViewer.Content = $sectionBorder
                            $sectionContainer.Children.Add($scrollViewer)

                            # Set the grid position for the section
                            [System.Windows.Controls.Grid]::SetColumn($sectionContainer, $currentColumn)
                            [System.Windows.Controls.Grid]::SetRow($sectionContainer, $currentRow)
                            $sectionsGrid.Children.Add($sectionContainer)

                            # Update position for next section
                            if ($currentColumn -eq 0) {
                                $currentColumn = 1
                            } else {
                                $currentColumn = 0
                                $currentRow++
                            }
                        }

                        # Add the grid to the details panel
                        $detailsPanel.Children.Add($sectionsGrid)

                        # Add separator
                        $separator = New-Object System.Windows.Controls.Separator
                        $separator.Margin = New-Object System.Windows.Thickness(0, 10, 0, 10)
                        $detailsPanel.Children.Add($separator)

                        # Add additional properties section (full width)
                        $additionalContainer = New-Object System.Windows.Controls.StackPanel
                        $additionalContainer.Margin = New-Object System.Windows.Thickness(5)

                        $additionalHeader = New-Object System.Windows.Controls.TextBlock
                        $additionalHeader.Text = "Additional Properties"
                        $additionalHeader.FontWeight = "Bold"
                        $additionalHeader.FontSize = 14
                        $additionalHeader.Foreground = "#2B579A"
                        $additionalHeader.Margin = New-Object System.Windows.Thickness(0, 5, 0, 5)
                        $additionalContainer.Children.Add($additionalHeader)

                        $additionalBorder = New-Object System.Windows.Controls.Border
                        $additionalBorder.Background = "#FFFFFF"
                        $additionalBorder.CornerRadius = New-Object System.Windows.CornerRadius(5)
                        $additionalBorder.Padding = New-Object System.Windows.Thickness(15)

                        $additionalBlock = New-Object System.Windows.Controls.TextBlock
                        $additionalText = ""
                        $device | Get-Member -MemberType Properties | 
                            Where-Object { $_.Name -notin $importantProps.Values.Name } |
                            ForEach-Object {
                                $property = $_.Name
                                $value = $device.$property
                                if ($null -ne $value -and $value -ne "") {
                                    $additionalText += "$property : $value`n"
                                }
                            }
                        $additionalBlock.Text = $additionalText
                        $additionalBlock.TextWrapping = "Wrap"
                        $additionalBlock.FontSize = 12

                        $additionalBorder.Child = $additionalBlock
                        $additionalContainer.Children.Add($additionalBorder)
                        $detailsPanel.Children.Add($additionalContainer)
                    }
                    "DevicePolicies" {
                        $detailsText = "Fetching policies for device: $($device.DeviceName)...`n`n"
                        
                        # Lazy-load config policies if not already cached
                        if ($null -eq $global:cachedData.ConfigPolicies) {
                            $detailsText += "Loading configuration policies...`n"
                            $global:cachedData.ConfigPolicies = Get-MgBetaDeviceManagementConfigurationPolicy -All
                        }
                        
                        # Lazy-load compliance policies if not already cached
                        if ($null -eq $global:cachedData.CompliancePolicies) {
                            $detailsText += "Loading compliance policies...`n"
                            $global:cachedData.CompliancePolicies = Get-MgBetaDeviceManagementCompliancePolicy -All
                        }
                        
                        $detailsText += "Configuration Policies:`n"
                        foreach ($policy in $global:cachedData.ConfigPolicies) {
                            $detailsText += "- $($policy.Name)`n"
                        }
                        
                        $detailsText += "`nCompliance Policies:`n"
                        foreach ($policy in $global:cachedData.CompliancePolicies) {
                            $detailsText += "- $($policy.DisplayName)`n"
                        }
                        
                        $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
                        $detailsTextBlock.Text = $detailsText
                        $detailsTextBlock.TextWrapping = "Wrap"
                        $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                        $detailsTextBlock.Foreground = "#000000"
                        $detailsTextBlock.FontSize = 12
                        
                        $detailsPanel.Children.Clear()
                        $detailsPanel.Children.Add($detailsTextBlock)
                    }
                    "DeviceApps" {
                        $detailsText = "Fetching applications for device: $($device.DeviceName)...`n`n"
                        
                        # Lazy-load apps if not already cached
                        if ($null -eq $global:cachedData.Apps) {
                            $detailsText += "Loading applications...`n"
                            $global:cachedData.Apps = Get-MgBetaDeviceAppManagementMobileApp -All
                        }
                        
                        $detailsText += "Applications:`n"
                        foreach ($app in $global:cachedData.Apps) {
                            $detailsText += "- $($app.DisplayName)`n"
                        }
                        
                        $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
                        $detailsTextBlock.Text = $detailsText
                        $detailsTextBlock.TextWrapping = "Wrap"
                        $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                        $detailsTextBlock.Foreground = "#000000"
                        $detailsTextBlock.FontSize = 12
                        
                        $detailsPanel.Children.Clear()
                        $detailsPanel.Children.Add($detailsTextBlock)
                    }
                    "DeviceScripts" {
                        $detailsText = "Fetching scripts for device: $($device.DeviceName)...`n`n"
                        
                        # Lazy-load scripts if not already cached
                        if ($null -eq $global:cachedData.HealthScripts) {
                            $detailsText += "Loading remediation scripts...`n"
                            $global:cachedData.HealthScripts = Get-MgBetaDeviceManagementDeviceHealthScript -All
                        }
                        
                        if ($null -eq $global:cachedData.Scripts) {
                            $detailsText += "Loading platform scripts...`n"
                            $global:cachedData.Scripts = Get-MgBetaDeviceManagementScript -All
                        }
                        
                        $detailsText += "Remediation Scripts:`n"
                        foreach ($script in $global:cachedData.HealthScripts) {
                            $detailsText += "- $($script.DisplayName)`n"
                        }
                        
                        $detailsText += "`nPlatform Scripts:`n"
                        foreach ($script in $global:cachedData.Scripts) {
                            $detailsText += "- $($script.DisplayName)`n"
                        }
                        
                        $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
                        $detailsTextBlock.Text = $detailsText
                        $detailsTextBlock.TextWrapping = "Wrap"
                        $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                        $detailsTextBlock.Foreground = "#000000"
                        $detailsTextBlock.FontSize = 12
                        
                        $detailsPanel.Children.Clear()
                        $detailsPanel.Children.Add($detailsTextBlock)
                    }
                }
            }
            # Process group-related nodes
            elseif ($nodeType -like "Group*") {
                $group = $selectedItem.Tag.Group
                
                # Use the TreeViewHelper function to update details for the group nodes
                Update-PropertyDetails -SelectedItem $selectedItem -propertyPanel $detailsPanel
            }
            else {
                $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
                $detailsTextBlock.Text = "Select an item to view details"
                $detailsTextBlock.TextWrapping = "Wrap"
                $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                $detailsTextBlock.Foreground = "#000000"
                $detailsTextBlock.FontSize = 12
                
                $detailsPanel.Children.Clear()
                $detailsPanel.Children.Add($detailsTextBlock)
            }
        }
        else {
            $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
            $detailsTextBlock.Text = "Select an item to view details"
            $detailsTextBlock.TextWrapping = "Wrap"
            $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
            $detailsTextBlock.Foreground = "#000000"
            $detailsTextBlock.FontSize = 12
            
            $detailsPanel.Children.Clear()
            $detailsPanel.Children.Add($detailsTextBlock)
        }
    }
    catch {
        $detailsTextBlock = New-Object System.Windows.Controls.TextBlock
        $detailsTextBlock.Text = "Error displaying details: $_"
        $detailsTextBlock.TextWrapping = "Wrap"
        $detailsTextBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $detailsTextBlock.Foreground = "#000000"
        $detailsTextBlock.FontSize = 12
        
        $detailsPanel.Children.Clear()
        $detailsPanel.Children.Add($detailsTextBlock)
    }
}

# Event handlers
$connectButton.Add_Click({
    if ($global:connectedToGraph) {
        Disconnect-FromGraph
    }
    else {
        Connect-ToGraph
    }
})

$refreshButton.Add_Click({
    Refresh-Data
})

$searchButton.Add_Click({
    Search-Item
})

$preloadButton.Add_Click({
    PreLoad-PolicyAssignments
})

$searchTextBox.Add_KeyDown({
    param($sender, $e)
    
    if ($e.Key -eq "Return") {
        Search-Item
    }
})

$intuneTreeView.Add_SelectedItemChanged({
    $selectedNode = $intuneTreeView.SelectedItem
    if ($null -ne $selectedNode) {
        Handle-TreeViewSelection -SelectedItem $selectedNode
    }
})

# Function to pre-load policy assignments (optional, can be run on demand)
function PreLoad-PolicyAssignments {
    try {
        # Update status
        $statusBar.Text = "Pre-loading policy assignments..."
        Show-Progress $true
        Write-Terminal "Starting to pre-load all policy assignments..."
        
        # Disable preload button while loading
        $preloadButton.IsEnabled = $false
        
        # Clear any existing assignment data to start fresh
        $global:cachedData.ConfigPolicyAssignments = @{}
        $global:cachedData.CompliancePolicyAssignments = @{}
        $global:cachedData.DeviceConfigAssignments = @{}
        $global:cachedData.AppAssignments = @{}
        Write-Terminal "Cleared existing assignment data."
        
        # Define the pre-loading script block
        $preloadScriptBlock = {
            # Initialize results container
            $results = @{
                TotalAssignmentsLoaded = 0
                TotalPoliciesWithAssignments = 0
                ConfigPolicies = $null
                CompliancePolicies = $null
                DeviceConfigs = $null
                Apps = $null
                ConfigPolicyAssignments = @{}
                CompliancePolicyAssignments = @{}
                DeviceConfigAssignments = @{}
                AppAssignments = @{}
                Errors = @{}
                Progress = @{}
            }
            
            # Function to safely make Graph requests in the background job
            function Invoke-SafeGraphRequest {
                param($Uri, $PolicyId, $PolicyName, $PolicyType)
                
                try {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
                    return @{
                        Success = $true
                        Data = $response
                    }
                }
                catch {
                    return @{
                        Success = $false
                        PolicyId = $PolicyId
                        PolicyName = $PolicyName
                        PolicyType = $PolicyType
                        Error = $_.Exception.Message
                    }
                }
            }
            
            try {
                # 1. Fetch configuration policies if needed
                try {
                    $results.ConfigPolicies = Get-MgBetaDeviceManagementConfigurationPolicy -All -ErrorAction Stop
                    $results.Progress["ConfigPolicies"] = "Loaded $($results.ConfigPolicies.Count) configuration policies"
                }
                catch {
                    $results.Errors["ConfigPolicies"] = @{
                        Message = "Failed to load configuration policies: $($_.Exception.Message)"
                    }
                }
                
                # 2. Process configuration policy assignments
                if ($results.ConfigPolicies) {
                    $totalPolicies = $results.ConfigPolicies.Count
                    $counter = 0
                    
                    foreach ($policy in $results.ConfigPolicies) {
                        $counter++
                        $results.Progress["ConfigProgress"] = "Processing config policy $counter of $totalPolicies"
                        
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($policy.Id)')/assignments"
                        $response = Invoke-SafeGraphRequest -Uri $assignmentsUri -PolicyId $policy.Id -PolicyName $policy.Name -PolicyType "Configuration"
                        
                        if ($response.Success -and $response.Data -and $response.Data.Value) {
                            $results.ConfigPolicyAssignments[$policy.Id] = $response.Data.Value
                            $results.TotalAssignmentsLoaded += $response.Data.Value.Count
                            $results.TotalPoliciesWithAssignments++
                        }
                    }
                    $results.Progress["ConfigPoliciesDone"] = "Completed processing $totalPolicies configuration policies"
                }
                
                # 3. Fetch compliance policies if needed
                try {
                    $results.CompliancePolicies = Get-MgBetaDeviceManagementCompliancePolicy -All -ErrorAction Stop
                    $results.Progress["CompliancePolicies"] = "Loaded $($results.CompliancePolicies.Count) compliance policies"
                }
                catch {
                    $results.Errors["CompliancePolicies"] = @{
                        Message = "Failed to load compliance policies: $($_.Exception.Message)"
                    }
                }
                
                # 4. Process compliance policy assignments
                if ($results.CompliancePolicies) {
                    $totalPolicies = $results.CompliancePolicies.Count
                    $counter = 0
                    
                    foreach ($policy in $results.CompliancePolicies) {
                        $counter++
                        $results.Progress["ComplianceProgress"] = "Processing compliance policy $counter of $totalPolicies"
                        
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies('$($policy.Id)')/assignments"
                        $response = Invoke-SafeGraphRequest -Uri $assignmentsUri -PolicyId $policy.Id -PolicyName $policy.DisplayName -PolicyType "Compliance"
                        
                        if ($response.Success -and $response.Data -and $response.Data.Value) {
                            $results.CompliancePolicyAssignments[$policy.Id] = $response.Data.Value
                            $results.TotalAssignmentsLoaded += $response.Data.Value.Count
                            $results.TotalPoliciesWithAssignments++
                        }
                    }
                    $results.Progress["CompliancePoliciesDone"] = "Completed processing $totalPolicies compliance policies"
                }
                
                # 5. Fetch device configurations if needed
                try {
                    $results.DeviceConfigs = Get-MgBetaDeviceManagementDeviceConfiguration -All -ErrorAction Stop
                    $results.Progress["DeviceConfigs"] = "Loaded $($results.DeviceConfigs.Count) device configurations"
                }
                catch {
                    $results.Errors["DeviceConfigs"] = @{
                        Message = "Failed to load device configurations: $($_.Exception.Message)"
                    }
                }
                
                # 6. Process device configuration assignments
                if ($results.DeviceConfigs) {
                    $totalConfigs = $results.DeviceConfigs.Count
                    $counter = 0
                    
                    foreach ($config in $results.DeviceConfigs) {
                        $counter++
                        $results.Progress["DeviceConfigProgress"] = "Processing device config $counter of $totalConfigs"
                        
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('$($config.Id)')/assignments"
                        $response = Invoke-SafeGraphRequest -Uri $assignmentsUri -PolicyId $config.Id -PolicyName $config.DisplayName -PolicyType "DeviceConfig"
                        
                        if ($response.Success -and $response.Data -and $response.Data.Value) {
                            $results.DeviceConfigAssignments[$config.Id] = $response.Data.Value
                            $results.TotalAssignmentsLoaded += $response.Data.Value.Count
                            $results.TotalPoliciesWithAssignments++
                        }
                    }
                    $results.Progress["DeviceConfigsDone"] = "Completed processing $totalConfigs device configurations"
                }
                
                # 7. Fetch mobile apps if needed
                try {
                    $results.Apps = Get-MgBetaDeviceAppManagementMobileApp -All -ErrorAction Stop
                    $results.Progress["Apps"] = "Loaded $($results.Apps.Count) mobile apps"
                }
                catch {
                    $results.Errors["Apps"] = @{
                        Message = "Failed to load mobile apps: $($_.Exception.Message)"
                    }
                }
                
                # 8. Process app assignments
                if ($results.Apps) {
                    $totalApps = $results.Apps.Count
                    $counter = 0
                    
                    foreach ($app in $results.Apps) {
                        $counter++
                        $results.Progress["AppProgress"] = "Processing app $counter of $totalApps"
                        
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps('$($app.Id)')/assignments"
                        $response = Invoke-SafeGraphRequest -Uri $assignmentsUri -PolicyId $app.Id -PolicyName $app.DisplayName -PolicyType "App"
                        
                        if ($response.Success -and $response.Data -and $response.Data.Value) {
                            $results.AppAssignments[$app.Id] = $response.Data.Value
                            $results.TotalAssignmentsLoaded += $response.Data.Value.Count
                            $results.TotalPoliciesWithAssignments++
                        }
                    }
                    $results.Progress["AppsDone"] = "Completed processing $totalApps mobile apps"
                }
            }
            catch {
                $results.Errors["General"] = @{
                    Message = $_.Exception.Message
                    StackTrace = $_.ScriptStackTrace
                }
            }
            
            return $results
        }
        
        # Define what happens when the async job completes or provides progress
        $onCompleteAction = {
            param($Result)
            
            # Store the cached data
            if ($Result.ConfigPolicies) {
                $global:cachedData.ConfigPolicies = $Result.ConfigPolicies
                $configPoliciesCount = $Result.ConfigPolicies.Count
                Write-Terminal "Stored $configPoliciesCount configuration policies."
            }
            
            if ($Result.CompliancePolicies) {
                $global:cachedData.CompliancePolicies = $Result.CompliancePolicies
                $compliancePoliciesCount = $Result.CompliancePolicies.Count
                Write-Terminal "Stored $compliancePoliciesCount compliance policies."
            }
            
            if ($Result.DeviceConfigs) {
                $global:cachedData.DeviceConfigs = $Result.DeviceConfigs
                $deviceConfigsCount = $Result.DeviceConfigs.Count
                Write-Terminal "Stored $deviceConfigsCount device configurations."
            }
            
            if ($Result.Apps) {
                $global:cachedData.Apps = $Result.Apps
                $appsCount = $Result.Apps.Count
                Write-Terminal "Stored $appsCount applications."
            }
            
            # Store all the assignments
            $global:cachedData.ConfigPolicyAssignments = $Result.ConfigPolicyAssignments
            $global:cachedData.CompliancePolicyAssignments = $Result.CompliancePolicyAssignments
            $global:cachedData.DeviceConfigAssignments = $Result.DeviceConfigAssignments
            $global:cachedData.AppAssignments = $Result.AppAssignments
            
            # Report any errors
            if ($Result.Errors.Count -gt 0) {
                foreach ($errorType in $Result.Errors.Keys) {
                    Write-Terminal "ERROR in ${errorType}: $($Result.Errors[$errorType].Message)"
                }
            }
            
            # Update UI with completion status
            $totalAssignmentsLoaded = $Result.TotalAssignmentsLoaded
            $totalPoliciesWithAssignments = $Result.TotalPoliciesWithAssignments
            
            $statusBar.Text = "Assignment data pre-loaded: $totalAssignmentsLoaded assignments for $totalPoliciesWithAssignments resources."
            $preloadButton.Content = "Assignments Pre-loaded"
            
            Write-Terminal "✅ All assignments pre-loaded: $totalAssignmentsLoaded total assignments for $totalPoliciesWithAssignments resources."
            
            # Add tooltip with details
            $preloadButton.ToolTip = "Loaded $totalAssignmentsLoaded assignments for $totalPoliciesWithAssignments resources (policies, configurations, and apps)"
            
            # Re-enable the preload button
            $preloadButton.IsEnabled = $true
            Show-Progress $false
        }
        
        # Set up progress reporting
        $progressReportAction = {
            param($ProgressData)
            
            # Update UI with progress information
            if ($ProgressData -and $ProgressData.Progress) {
                foreach ($progressKey in $ProgressData.Progress.Keys) {
                    $progressMessage = $ProgressData.Progress[$progressKey]
                    $statusBar.Text = $progressMessage
                    Write-Terminal $progressMessage
                }
            }
        }
        
        # Start the async job
        Start-AsyncJob -ScriptBlock $preloadScriptBlock -JobName "PreloadAssignments" -OnComplete $onCompleteAction
    }
    catch {
        $errorMessage = $_.Exception.Message
        $statusBar.Text = "Error starting pre-load: $errorMessage"
        Write-Error "Error starting pre-load: $_"
        Write-Terminal "ERROR: Failed to start pre-loading assignments: $errorMessage"
        Write-Terminal "Stack Trace: $($_.ScriptStackTrace)"
        
        # Re-enable the preload button
        $preloadButton.IsEnabled = $true
        Show-Progress $false
    }
}

# Start the timer to check for completed async jobs
$jobTimer.Start()
Write-Terminal "Async job processor started successfully"

# Start the application
$window.ShowDialog() | Out-Null 