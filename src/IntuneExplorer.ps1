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

# Import helper scripts
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath\ErrorHandler.ps1"
. "$scriptPath\TreeViewHelper.ps1"

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
        
        <!-- Search Panel -->
        <Border Grid.Row="1" Background="#F5F5F5" CornerRadius="10" Padding="10" Margin="0,0,0,10">
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
        <Grid Grid.Row="2">
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
        <Border Grid.Row="3" Background="#F5F5F5" CornerRadius="10" Padding="5" Margin="0,10,0,0">
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

# Function to connect to Microsoft Graph
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
        
        # Automatically pre-load assignments after initial data load
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

# Function to disconnect from Microsoft Graph
function Disconnect-FromGraph {
    try {
        Show-Progress $true
        Disconnect-MgGraph
        $connectionStatus.Text = "Not Connected"
        $global:connectedToGraph = $false
        $connectButton.Content = "Connect to Microsoft Graph"
        $refreshButton.IsEnabled = $false
        $statusBar.Text = "Disconnected from Microsoft Graph"
        
        # Clear the TreeView
        $intuneTreeView.Items.Clear()
        $detailsPanel.Children.Clear()
        
        # Clear cached data
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
    }
    catch {
        $statusBar.Text = "Disconnection failed: $_"
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
        
        # Reset preload button state
        $preloadButton.Content = "Pre-load Assignments"
        $preloadButton.IsEnabled = $true
        
        # Clear the treeview
        $intuneTreeView.Items.Clear()
        
        # Load basic data (this makes the UI responsive without waiting for all data)
        try {
            $statusBar.Text = "Loading groups..."
            $global:cachedData.Groups = Get-MgGroup -All -ErrorAction Stop
        }
        catch {
            Write-Host "Error fetching groups: $_" -ForegroundColor Red
        }
        
        try {
            $statusBar.Text = "Loading devices..."
            $global:cachedData.Devices = Get-MgBetaDeviceManagementManagedDevice -All -ErrorAction Stop
        }
        catch {
            Write-Host "Error fetching devices: $_" -ForegroundColor Red
        }
        
        try {
            $statusBar.Text = "Loading users..."
            $global:cachedData.Users = Get-MgUser -All -ErrorAction Stop
        }
        catch {
            Write-Host "Error fetching users: $_" -ForegroundColor Red
        }
        
        # Data loaded successfully
        $statusBar.Text = "Basic data loaded. Starting assignment pre-loading automatically..."
    }
    catch {
        $statusBar.Text = "Error refreshing data: $($_.Exception.Message)"
        Write-Error "Error refreshing data: $_"
    }
    finally {
        Show-Progress $false
    }
}

# Function to search for a device or group
function Search-Item {
    $searchTerm = $searchTextBox.Text.Trim()
    
    # Ensure we have a valid search type selection
    if ($null -eq $searchTypeComboBox.SelectedItem) {
        $statusBar.Text = "Please select a search type"
        return
    }
    
    $searchType = $searchTypeComboBox.SelectedItem.Content.ToString()
    
    if ([string]::IsNullOrEmpty($searchTerm)) {
        $statusBar.Text = "Please enter a search term"
        return
    }
    
    if (-not $global:connectedToGraph) {
        $statusBar.Text = "Please connect to Microsoft Graph first"
        return
    }
    
    Show-Progress $true
    $statusBar.Text = "Searching for $($searchType): $searchTerm..."
    
    try {
        # Clear the TreeView
        $intuneTreeView.Items.Clear()
        $detailsPanel.Children.Clear()
        
        switch ($searchType) {
            "Devices" {
                try {
                    # Ensure devices are fetched
                    if ($null -eq $global:cachedData.Devices) {
                        FetchIntuneData
                    }
                    
                    # Find devices matching the search term
                    $matchingDevices = $global:cachedData.Devices | Where-Object { 
                        $_.DeviceName -like "*$searchTerm*" -or 
                        $_.SerialNumber -like "*$searchTerm*" -or 
                        $_.Id -like "*$searchTerm*" 
                    }
                    
                    if ($null -eq $matchingDevices -or ($matchingDevices -is [array] -and $matchingDevices.Count -eq 0)) {
                        $statusBar.Text = "No matching devices found"
                        return
                    }
                    
                    # Convert to array if single object
                    if (-not ($matchingDevices -is [array])) {
                        $matchingDevices = @($matchingDevices)
                    }
                    
                    # Display the matching devices in the TreeView
                    foreach ($device in $matchingDevices) {
                        Add-DeviceToTreeView -TreeView $intuneTreeView -Device $device
                    }
                    
                    $statusBar.Text = "Found $($matchingDevices.Count) devices"
                }
                catch {
                    $errorMsg = "Error searching for devices: $_"
                    Write-Host $errorMsg -ForegroundColor Red
                    $statusBar.Text = $errorMsg
                }
            }
            
            "Groups" {
                try {
                    # Ensure groups are fetched
                    if ($null -eq $global:cachedData.Groups) {
                        $statusBar.Text = "Fetching groups..."
                        try {
                            $global:cachedData.Groups = Get-MgBetaGroup -All -ErrorAction Stop
                        }
                        catch {
                            $errorMsg = "Failed to fetch groups: $_. Please check your permissions."
                            Write-Host $errorMsg -ForegroundColor Red
                            $statusBar.Text = $errorMsg
                            Show-Progress $false
                            return
                        }
                    }
                    
                    if ($null -eq $global:cachedData.Groups) {
                        $statusBar.Text = "Unable to retrieve groups. Check your connection and permissions."
                        Show-Progress $false
                        return
                    }
                    
                    # Find groups matching the search term
                    $matchingGroups = $global:cachedData.Groups | Where-Object {
                        $_.DisplayName -like "*$searchTerm*" -or
                        $_.Id -like "*$searchTerm*"
                    }
                    
                    if ($null -eq $matchingGroups -or ($matchingGroups -is [array] -and $matchingGroups.Count -eq 0)) {
                        $statusBar.Text = "No matching groups found"
                        return
                    }
                    
                    # Convert to array if single object
                    if (-not ($matchingGroups -is [array])) {
                        $matchingGroups = @($matchingGroups)
                    }
                    
                    # Display the matching groups in the TreeView
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
                            }
                        }
                        catch {
                            Write-Host "Error adding group '$groupName' to tree view: $_" -ForegroundColor Red
                            # Continue with the next group
                        }
                    }
                    
                    if ($groupsAdded -gt 0) {
                        $statusBar.Text = "Found $groupsAdded groups"
                    }
                    else {
                        $statusBar.Text = "Found $($matchingGroups.Count) groups but none could be displayed. Check the console for errors."
                        Write-Host "Error: All groups found, but none could be displayed. You may need additional permissions." -ForegroundColor Red
                    }
                }
                catch {
                    $errorMsg = "Error searching for groups: $_"
                    Write-Host $errorMsg -ForegroundColor Red 
                    $statusBar.Text = $errorMsg
                }
            }
        }
    }
    catch {
        $errorMsg = "Error in Search-Item: $_"
        Write-Host $errorMsg -ForegroundColor Red
        $statusBar.Text = $errorMsg
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
        
        # Clear any existing assignment data to start fresh
        $global:cachedData.ConfigPolicyAssignments = @{}
        $global:cachedData.CompliancePolicyAssignments = @{}
        $global:cachedData.DeviceConfigAssignments = @{}
        $global:cachedData.AppAssignments = @{}
        
        # Track counts for reporting
        $totalAssignmentsLoaded = 0
        $totalPoliciesWithAssignments = 0
        
        # Load configuration policies if they're not already loaded
        if ($null -eq $global:cachedData.ConfigPolicies) {
            $statusBar.Text = "Loading configuration policies..."
            $global:cachedData.ConfigPolicies = Get-MgBetaDeviceManagementConfigurationPolicy -All -ErrorAction SilentlyContinue
        }
        
        # For each policy, pre-fetch and cache its assignments
        $totalPolicies = $global:cachedData.ConfigPolicies.Count
        $counter = 0
        
        foreach ($policy in $global:cachedData.ConfigPolicies) {
            $counter++
            $statusBar.Text = "Loading config policy assignments ($counter of $totalPolicies)..."
            
            try {
                $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($policy.Id)')/assignments"
                $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                
                if ($assignments -and $assignments.Value) {
                    $global:cachedData.ConfigPolicyAssignments[$policy.Id] = $assignments.Value
                    $totalAssignmentsLoaded += $assignments.Value.Count
                    $totalPoliciesWithAssignments++
                }
            }
            catch {
                Write-Host "Error fetching assignments for policy $($policy.Name): $_" -ForegroundColor Yellow
            }
        }
        
        # Load compliance policies if they're not already loaded
        if ($null -eq $global:cachedData.CompliancePolicies) {
            $statusBar.Text = "Loading compliance policies..."
            $global:cachedData.CompliancePolicies = Get-MgBetaDeviceManagementCompliancePolicy -All -ErrorAction SilentlyContinue
        }
        
        # For each compliance policy, pre-fetch and cache its assignments
        $totalPolicies = $global:cachedData.CompliancePolicies.Count
        $counter = 0
        
        foreach ($policy in $global:cachedData.CompliancePolicies) {
            $counter++
            $statusBar.Text = "Loading compliance policy assignments ($counter of $totalPolicies)..."
            
            try {
                $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies('$($policy.Id)')/assignments"
                $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                
                if ($assignments -and $assignments.Value) {
                    $global:cachedData.CompliancePolicyAssignments[$policy.Id] = $assignments.Value
                    $totalAssignmentsLoaded += $assignments.Value.Count
                    $totalPoliciesWithAssignments++
                }
            }
            catch {
                Write-Host "Error fetching assignments for compliance policy $($policy.DisplayName): $_" -ForegroundColor Yellow
            }
        }
        
        # Load device configurations if they're not already loaded
        if ($null -eq $global:cachedData.DeviceConfigs) {
            $statusBar.Text = "Loading device configurations..."
            $global:cachedData.DeviceConfigs = Get-MgBetaDeviceManagementDeviceConfiguration -All -ErrorAction SilentlyContinue
        }
        
        # For each device config, pre-fetch and cache its assignments
        $totalConfigs = $global:cachedData.DeviceConfigs.Count
        $counter = 0
        
        foreach ($config in $global:cachedData.DeviceConfigs) {
            $counter++
            $statusBar.Text = "Loading device config assignments ($counter of $totalConfigs)..."
            
            try {
                $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('$($config.Id)')/assignments"
                $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                
                if ($assignments -and $assignments.Value) {
                    $global:cachedData.DeviceConfigAssignments[$config.Id] = $assignments.Value
                    $totalAssignmentsLoaded += $assignments.Value.Count
                    $totalPoliciesWithAssignments++
                }
            }
            catch {
                Write-Host "Error fetching assignments for device config $($config.DisplayName): $_" -ForegroundColor Yellow
            }
        }
        
        # Load mobile apps if they're not already loaded
        if ($null -eq $global:cachedData.Apps) {
            $statusBar.Text = "Loading mobile apps..."
            $global:cachedData.Apps = Get-MgBetaDeviceAppManagementMobileApp -All -ErrorAction SilentlyContinue
        }
        
        # For each app, pre-fetch and cache its assignments
        $totalApps = $global:cachedData.Apps.Count
        $counter = 0
        
        foreach ($app in $global:cachedData.Apps) {
            $counter++
            $statusBar.Text = "Loading app assignments ($counter of $totalApps)..."
            
            try {
                $assignmentsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps('$($app.Id)')/assignments"
                $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                
                if ($assignments -and $assignments.Value) {
                    $global:cachedData.AppAssignments[$app.Id] = $assignments.Value
                    $totalAssignmentsLoaded += $assignments.Value.Count
                    $totalPoliciesWithAssignments++
                }
            }
            catch {
                Write-Host "Error fetching assignments for app $($app.DisplayName): $_" -ForegroundColor Yellow
            }
        }
        
        # Update status with counts and disable pre-load button (since data is now loaded)
        $statusBar.Text = "Assignment data pre-loaded: $totalAssignmentsLoaded assignments for $totalPoliciesWithAssignments resources."
        $preloadButton.Content = "Assignments Pre-loaded"
        $preloadButton.IsEnabled = $false
        
        # Add tooltip with details
        $preloadButton.ToolTip = "Loaded $totalAssignmentsLoaded assignments for $totalPoliciesWithAssignments resources (policies, configurations, and apps)"
    }
    catch {
        $statusBar.Text = "Error pre-loading assignments: $($_.Exception.Message)"
        Write-Error "Error pre-loading assignments: $_"
    }
    finally {
        Show-Progress $false
    }
}

# Start the application
$window.ShowDialog() | Out-Null 