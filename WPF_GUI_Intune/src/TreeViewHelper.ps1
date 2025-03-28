# TreeView helper functions for Intune Explorer

# Required for UI update calls
Add-Type -AssemblyName System.Windows.Forms

function New-TreeViewNode {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Header,
        
        [Parameter(Mandatory = $false)]
        [object]$Tag = $null,
        
        [Parameter(Mandatory = $false)]
        [string]$Icon = $null,
        
        [Parameter(Mandatory = $false)]
        [System.Windows.Media.Brush]$HeaderColor = $null,
        
        [Parameter(Mandatory = $false)]
        [switch]$IsExpanded = $false,
        
        [Parameter(Mandatory = $false)]
        [switch]$LongGroupName = $false
    )
    
    try {
        # Safely create the TreeViewItem
        $node = New-Object System.Windows.Controls.TreeViewItem -ErrorAction Stop
        
        # Set a default header if null
        if ([string]::IsNullOrEmpty($Header)) {
            $Header = "Unknown"
        }
        
        if ([string]::IsNullOrEmpty($Icon)) {
            # Simple text header
            $node.Header = $Header
        }
        else {
            # Create header with icon
            $headerStack = New-Object System.Windows.Controls.StackPanel -ErrorAction Stop
            $headerStack.Orientation = "Horizontal"
            
            $iconText = New-Object System.Windows.Controls.TextBlock -ErrorAction Stop
            $iconText.Text = $Icon
            $iconText.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
            $null = $headerStack.Children.Add($iconText)
            
            $nameText = New-Object System.Windows.Controls.TextBlock -ErrorAction Stop
            $nameText.Text = $Header
            
            # For groups, handle long names better
            if ($LongGroupName -or $Icon -eq "[G]") {
                $nameText.TextWrapping = "NoWrap"
                $nameText.MaxWidth = 300
                $nameText.TextTrimming = "CharacterEllipsis"
                $nameText.ToolTip = $Header  # Show full name on hover
            }
            
            # Apply custom color if specified
            if ($null -ne $HeaderColor) {
                $nameText.Foreground = $HeaderColor
            }
            
            $null = $headerStack.Children.Add($nameText)
            
            $node.Header = $headerStack
        }
        
        $node.Tag = $Tag
        $node.IsExpanded = $IsExpanded
        
        return $node
    }
    catch {
        Write-Host "Error creating TreeViewNode: $_" -ForegroundColor Red
        return $null
    }
}

function Get-DeviceIconByType {
    param (
        [string]$OperatingSystem,
        [string]$Model
    )
    
    if ($OperatingSystem -like "*iOS*") {
        return "[M]"  # Mobile
    }
    elseif ($OperatingSystem -like "*Android*") {
        return "[M]"  # Mobile
    }
    elseif ($OperatingSystem -like "*Windows*" -and $Model -like "*Surface*") {
        return "[T]"  # Tablet
    }
    else {
        return "[P]"  # PC
    }
}

function Get-IconForNodeType {
    param (
        [Parameter(Mandatory = $true)]
        [string]$NodeType
    )
    
    switch ($NodeType) {
        "DeviceProperties" { return "[i]" }  # Info
        "DevicePolicies"   { return "[S]" }  # Security
        "DeviceApps"       { return "[A]" }  # Apps
        "DeviceScripts"    { return "[C]" }  # Code/Scripts
        "GroupProperties"  { return "[i]" }  # Info
        "GroupConfigs"     { return "[S]" }  # Security/Config
        "GroupApps"        { return "[A]" }  # Apps
        default            { return "*" }     # Default asterisk
    }
}

function Clear-TreeView {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TreeView]$TreeView
    )
    
    $TreeView.Items.Clear()
}

function Add-DeviceToTreeView {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TreeView]$TreeView,
        
        [Parameter(Mandatory = $true)]
        [object]$Device,
        
        [Parameter(Mandatory = $false)]
        [switch]$IsExpanded = $false
    )
    
    # Create the device node
    $deviceIcon = Get-DeviceIconByType -OperatingSystem $Device.OperatingSystem -Model $Device.Model
    $deviceNode = New-TreeViewNode -Header $Device.DeviceName -Tag $Device -Icon $deviceIcon -IsExpanded:$IsExpanded
    
    # Add child nodes
    $childNodes = @(
        @{ Header = "Properties"; Type = "DeviceProperties" },
        @{ Header = "Policies"; Type = "DevicePolicies" },
        @{ Header = "Applications"; Type = "DeviceApps" },
        @{ Header = "Scripts"; Type = "DeviceScripts" }
    )
    
    foreach ($childNode in $childNodes) {
        $icon = Get-IconForNodeType -NodeType $childNode.Type
        $node = New-TreeViewNode -Header $childNode.Header -Tag @{ Type = $childNode.Type; Device = $Device } -Icon $icon
        $deviceNode.Items.Add($node)
    }
    
    # Add to the tree view
    $TreeView.Items.Add($deviceNode)
    
    return $deviceNode
}

function Get-DeviceProperties {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Device
    )
    
    # Convert device object to hashtable of key properties
    $properties = @{
        "Device Name" = $Device.deviceName
        "Operating System" = $Device.operatingSystem
        "OS Version" = $Device.osVersion
        "Management State" = $Device.managementState
        "Enrollment Type" = $Device.enrollmentType
        "Last Sync" = $Device.lastSyncDateTime
        "Compliance State" = $Device.complianceState
        "Serial Number" = $Device.serialNumber
        "Model" = $Device.model
        "Manufacturer" = $Device.manufacturer
        "Storage" = if ($Device.totalStorageSpaceInBytes) { 
            "$([math]::Round($Device.totalStorageSpaceInBytes/1GB, 2)) GB" 
        } else { "Unknown" }
        "Free Storage" = if ($Device.freeStorageSpaceInBytes) { 
            "$([math]::Round($Device.freeStorageSpaceInBytes/1GB, 2)) GB" 
        } else { "Unknown" }
    }
    
    return $properties
}

function Update-PropertyDetails {
    param (
        [Parameter(Mandatory = $true)]
        [object]$SelectedItem,
        
        [Parameter(Mandatory = $true)]
        [object]$propertyPanel
    )
    
    try {
        Write-Host "Updating property details for $($SelectedItem.Tag.Type): $($SelectedItem.Header)"
        $propertyPanel.Children.Clear()
        
        # Add a title for the property panel
        $title = New-Object System.Windows.Controls.TextBlock
        $title.Text = "Properties for $($SelectedItem.Header)"
        $title.FontWeight = "Bold"
        $title.FontSize = 16
        $title.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
        $propertyPanel.Children.Add($title)
        
        # Add helper function to add property rows
        function Add-PropertyRow {
            param (
                [string]$Label,
                [object]$Value,
                [string]$ID = ""
            )
            
            $propStack = New-Object System.Windows.Controls.StackPanel
            $propStack.Orientation = "Horizontal"
            $propStack.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
            
            $labelBlock = New-Object System.Windows.Controls.TextBlock
            $labelBlock.Text = "$($Label): "
            $labelBlock.FontWeight = "Bold"
            $labelBlock.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
            
            $valueBlock = New-Object System.Windows.Controls.TextBlock
            $valueBlock.Text = if ($null -eq $Value) { "N/A" } else { $Value }
            $valueBlock.TextWrapping = "Wrap"
            
            if ($ID) {
                $valueBlock.Name = $ID
            }
            
            $propStack.Children.Add($labelBlock)
            $propStack.Children.Add($valueBlock)
            
            $propertyPanel.Children.Add($propStack)
        }
        
        # Add helper function to update property value
        function Update-PropertyValue {
            param (
                [string]$ID,
                [string]$Value
            )
            
            $controls = $propertyPanel.Children | Where-Object { $_ -is [System.Windows.Controls.StackPanel] }
            foreach ($control in $controls) {
                $valueBlock = $control.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.Name -eq $ID }
                if ($valueBlock) {
                    $valueBlock.Text = $Value
                    break
                }
            }
        }
        
        # Add basic properties based on the item type
        switch ($SelectedItem.Tag.Type) {
            "Device" {
                Add-PropertyRow "Device Name" $SelectedItem.Tag.DeviceName
                Add-PropertyRow "Management Name" $SelectedItem.Tag.DeviceObject.ManagedDeviceName
                Add-PropertyRow "OS" $SelectedItem.Tag.DeviceObject.OperatingSystem
                Add-PropertyRow "OS Version" $SelectedItem.Tag.DeviceObject.OSVersion
                Add-PropertyRow "Serial Number" $SelectedItem.Tag.DeviceObject.SerialNumber
                Add-PropertyRow "Last Sync" $SelectedItem.Tag.DeviceObject.LastSyncDateTime
                Add-PropertyRow "Compliance State" $SelectedItem.Tag.DeviceObject.ComplianceState
                Add-PropertyRow "Ownership" $SelectedItem.Tag.DeviceObject.Ownership
                Add-PropertyRow "Primary User" $SelectedItem.Tag.DeviceObject.UserPrincipalName
                Add-PropertyRow "Enrolled Date" $SelectedItem.Tag.DeviceObject.EnrolledDateTime
            }
            "User" {
                Add-PropertyRow "Display Name" $SelectedItem.Tag.UserObject.DisplayName
                Add-PropertyRow "UPN" $SelectedItem.Tag.UserObject.UserPrincipalName
                Add-PropertyRow "Email" $SelectedItem.Tag.UserObject.Mail
                Add-PropertyRow "Account Enabled" $SelectedItem.Tag.UserObject.AccountEnabled
                Add-PropertyRow "City" $SelectedItem.Tag.UserObject.City
                Add-PropertyRow "Department" $SelectedItem.Tag.UserObject.Department
                Add-PropertyRow "Job Title" $SelectedItem.Tag.UserObject.JobTitle
                Add-PropertyRow "License Status" (if ($SelectedItem.Tag.UserObject.AssignedLicenses) { "Licensed" } else { "Unlicensed" })
            }
            "Group" {
                Add-PropertyRow "Display Name" $SelectedItem.Tag.GroupObject.DisplayName
                Add-PropertyRow "Description" $SelectedItem.Tag.GroupObject.Description
                Add-PropertyRow "Group ID" $SelectedItem.Tag.GroupObject.Id
                Add-PropertyRow "Group Type" (if ($SelectedItem.Tag.GroupObject.SecurityEnabled) { "Security" } else { "Distribution" })
                Add-PropertyRow "Mail Enabled" $SelectedItem.Tag.GroupObject.MailEnabled
                Add-PropertyRow "Mail" $SelectedItem.Tag.GroupObject.Mail
                Add-PropertyRow "Membership Type" (if ($SelectedItem.Tag.GroupObject.GroupTypes -contains "DynamicMembership") { "Dynamic" } else { "Assigned" })
                Add-PropertyRow "Created" $SelectedItem.Tag.GroupObject.CreatedDateTime
            }
            "GroupConfigs" {
                try {
                    # Show loading message
                    $loadingBlock = New-Object System.Windows.Controls.TextBlock
                    $loadingBlock.Text = "Loading configurations for group: $($SelectedItem.Tag.Group.DisplayName)..."
                    $loadingBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                    $propertyPanel.Children.Add($loadingBlock)
                    
                    # Add more detail
                    $detailsText = "Retrieving assignments from Intune. This may take a moment...`n"
                    $detailsText += "Looking for configuration policies, compliance policies, and device configurations."
                    
                    # If using cached data, show that
                    if ($global:cachedData.ConfigPolicyAssignments.Count -gt 0 -or $global:cachedData.CompliancePolicyAssignments.Count -gt 0 -or $global:cachedData.DeviceConfigAssignments.Count -gt 0) {
                        $detailsText = "Using pre-loaded assignment data. This should be fast!`n"
                        $detailsText += "Checking cached assignments for this group..."
                    }
                    
                    $detailBlock = New-Object System.Windows.Controls.TextBlock
                    $detailBlock.Text = $detailsText
                    $detailBlock.TextWrapping = "Wrap"
                    $detailBlock.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
                    $propertyPanel.Children.Add($detailBlock)
                    
                    # Update the UI to show the loading message
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Get configurations
                    $configurations = Get-GroupConfigurations -Group $SelectedItem.Tag.Group
                    
                    # Clear loading message
                    $propertyPanel.Children.Remove($loadingBlock)
                    
                    if ($configurations.Count -eq 0) {
                        $noConfigsMsg = New-Object System.Windows.Controls.TextBlock
                        $noConfigsMsg.Text = "No configurations found for this group."
                        $noConfigsMsg.Margin = New-Object System.Windows.Thickness(0, 5, 0, 0)
                        $propertyPanel.Children.Add($noConfigsMsg)
                    }
                    else {
                        # Display each configuration
                        foreach ($config in $configurations) {
                            $configPanel = New-Object System.Windows.Controls.StackPanel
                            $configPanel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
                            
                            # Check if the panel supports the Background property before setting it
                            if ($configPanel.GetType().GetProperty("Background")) {
                                $configPanel.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::WhiteSmoke)
                            }
                            
                            # Check if the panel supports the Padding property before setting it
                            if ($configPanel.GetType().GetProperty("Padding")) {
                                $configPanel.Padding = New-Object System.Windows.Thickness(10)
                            }
                            
                            $nameBlock = New-Object System.Windows.Controls.TextBlock
                            $nameBlock.Text = $config.Name
                            $nameBlock.FontWeight = "Bold"
                            $nameBlock.FontSize = 14
                            $configPanel.Children.Add($nameBlock)
                            
                            $typeBlock = New-Object System.Windows.Controls.TextBlock
                            $typeBlock.Text = "Type: $($config.Type)"
                            $typeBlock.Margin = New-Object System.Windows.Thickness(0, 5, 0, 5)
                            $configPanel.Children.Add($typeBlock)
                            
                            if ($config.Description) {
                                $descBlock = New-Object System.Windows.Controls.TextBlock
                                $descBlock.Text = "Description: $($config.Description)"
                                $descBlock.TextWrapping = "Wrap"
                                $configPanel.Children.Add($descBlock)
                            }
                            
                            $propertyPanel.Children.Add($configPanel)
                        }
                    }
                }
                catch {
                    # Update the loading message to show the error
                    $loadingBlock.Text = "Error loading configurations: $($_.Exception.Message)"
                    $loadingBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Red)
                    Write-Error "Error getting configurations for group: $_"
                }
            }
            "GroupApps" {
                try {
                    # Show loading message
                    $loadingBlock = New-Object System.Windows.Controls.TextBlock
                    $loadingBlock.Text = "Loading applications for group: $($SelectedItem.Tag.Group.DisplayName)..."
                    $loadingBlock.Margin = New-Object System.Windows.Thickness(0, 0, 0, 10)
                    $propertyPanel.Children.Add($loadingBlock)
                    
                    # Add more detail
                    $detailsText = "Retrieving application assignments from Intune. This may take a moment...`n"
                    $detailsText += "Looking for mobile apps and Win32 apps assigned to this group."
                    
                    # If using cached data, show that
                    if ($global:cachedData.AppAssignments.Count -gt 0) {
                        $detailsText = "Using pre-loaded assignment data. This should be fast!`n"
                        $detailsText += "Checking cached assignments for this group..."
                    }
                    
                    $detailBlock = New-Object System.Windows.Controls.TextBlock
                    $detailBlock.Text = $detailsText
                    $detailBlock.TextWrapping = "Wrap"
                    $detailBlock.Margin = New-Object System.Windows.Thickness(0, 10, 0, 0)
                    $propertyPanel.Children.Add($detailBlock)
                    
                    # Update the UI to show the loading message
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Get applications
                    $applications = Get-GroupApplications -Group $SelectedItem.Tag.Group
                    
                    # Clear loading message
                    $propertyPanel.Children.Remove($loadingBlock)
                    
                    if ($applications.Count -eq 0) {
                        $noAppsMsg = New-Object System.Windows.Controls.TextBlock
                        $noAppsMsg.Text = "No applications found for this group."
                        $noAppsMsg.Margin = New-Object System.Windows.Thickness(0, 5, 0, 0)
                        $propertyPanel.Children.Add($noAppsMsg)
                    }
                    else {
                        # Display each application
                        foreach ($app in $applications) {
                            $appPanel = New-Object System.Windows.Controls.StackPanel
                            $appPanel.Margin = New-Object System.Windows.Thickness(0, 5, 0, 15)
                            
                            # Check if the panel supports the Background property before setting it
                            if ($appPanel.GetType().GetProperty("Background")) {
                                $appPanel.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::WhiteSmoke)
                            }
                            
                            # Check if the panel supports the Padding property before setting it
                            if ($appPanel.GetType().GetProperty("Padding")) {
                                $appPanel.Padding = New-Object System.Windows.Thickness(10)
                            }
                            
                            $nameBlock = New-Object System.Windows.Controls.TextBlock
                            $nameBlock.Text = $app.Name
                            $nameBlock.FontWeight = "Bold"
                            $nameBlock.FontSize = 14
                            $appPanel.Children.Add($nameBlock)
                            
                            $typeBlock = New-Object System.Windows.Controls.TextBlock
                            $typeBlock.Text = "Type: $($app.Type)"
                            $typeBlock.Margin = New-Object System.Windows.Thickness(0, 5, 0, 5)
                            $appPanel.Children.Add($typeBlock)
                            
                            if ($app.Description) {
                                $descBlock = New-Object System.Windows.Controls.TextBlock
                                $descBlock.Text = "Description: $($app.Description)"
                                $descBlock.TextWrapping = "Wrap"
                                $appPanel.Children.Add($descBlock)
                            }
                            
                            $propertyPanel.Children.Add($appPanel)
                        }
                    }
                }
                catch {
                    # Update the loading message to show the error
                    $loadingBlock.Text = "Error loading applications: $($_.Exception.Message)"
                    $loadingBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Red)
                    Write-Error "Error getting applications for group: $_"
                }
            }
            "DeviceApps" {
                # Show applications installed on the device
                $deviceObj = $SelectedItem.Parent.Tag.DeviceObject
                Add-PropertyRow "Device Name" $deviceObj.DeviceName
                Add-PropertyRow "Applications" "Loading..." "LoadingApps"
                
                # This would be a good place to use a background job to load apps
                # For now, we'll simulate with a simple list
                
                $apps = @(
                    "Microsoft Office 365",
                    "Microsoft Teams",
                    "Microsoft Edge",
                    "Adobe Acrobat Reader"
                )
                
                Update-PropertyValue "LoadingApps" ($apps -join ", ")
            }
            "DeviceConfigs" {
                # Show configurations assigned to the device
                $deviceObj = $SelectedItem.Parent.Tag.DeviceObject
                Add-PropertyRow "Device Name" $deviceObj.DeviceName
                Add-PropertyRow "Configurations" "Loading..." "LoadingConfigs"
                
                # This would be a good place to use a background job to load configs
                # For now, we'll simulate with a simple list
                
                $configs = @(
                    "Windows 10 Security Baseline",
                    "Corporate Compliance Settings",
                    "Device Encryption Policy"
                )
                
                Update-PropertyValue "LoadingConfigs" ($configs -join ", ")
            }
            default {
                Add-PropertyRow "Item Type" $SelectedItem.Tag.Type
                Add-PropertyRow "Name" $SelectedItem.Header
            }
        }
    }
    catch {
        Write-Error "Error updating property details: $_"
        
        # Show error in property panel
        $propertyPanel.Children.Clear()
        $errorBlock = New-Object System.Windows.Controls.TextBlock
        $errorBlock.Text = "Error loading properties: $($_.Exception.Message)"
        $errorBlock.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Colors]::Red)
        $errorBlock.TextWrapping = "Wrap"
        $propertyPanel.Children.Add($errorBlock)
    }
}

function Add-GroupToTreeView {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.TreeView]$TreeView,
        
        [Parameter(Mandatory = $true)]
        [object]$Group,
        
        [Parameter(Mandatory = $false)]
        [switch]$IsExpanded = $false
    )
    
    try {
        # Check for null references
        if ($null -eq $Group) {
            Write-Host "Error: Group is null" -ForegroundColor Red
            return $null
        }
        
        if ($null -eq $Group.DisplayName) {
            $displayName = "Group: " + ($Group.Id ?? "Unknown ID")
        } else {
            $displayName = $Group.DisplayName
        }
        
        # Create the group node with special handling for long names
        $groupNode = New-TreeViewNode -Header $displayName -Tag $Group -Icon "[G]" -IsExpanded:$IsExpanded -LongGroupName
        
        # Add child nodes
        # For configurations, first check if we have pre-loaded data
        $configCount = 0
        
        # If we have pre-loaded assignment data, we can get the count of configurations right away
        if ($global:cachedData.ConfigPolicyAssignments.Count -gt 0 -or 
            $global:cachedData.CompliancePolicyAssignments.Count -gt 0 -or 
            $global:cachedData.DeviceConfigAssignments.Count -gt 0) {
            
            # Check configuration policies
            foreach ($policyId in $global:cachedData.ConfigPolicyAssignments.Keys) {
                $assignments = $global:cachedData.ConfigPolicyAssignments[$policyId]
                if ($assignments | Where-Object { $_.target.groupId -eq $Group.Id }) {
                    $configCount++
                }
            }
            
            # Check compliance policies
            foreach ($policyId in $global:cachedData.CompliancePolicyAssignments.Keys) {
                $assignments = $global:cachedData.CompliancePolicyAssignments[$policyId]
                if ($assignments | Where-Object { $_.target.groupId -eq $Group.Id }) {
                    $configCount++
                }
            }
            
            # Check device configurations
            foreach ($configId in $global:cachedData.DeviceConfigAssignments.Keys) {
                $assignments = $global:cachedData.DeviceConfigAssignments[$configId]
                if ($assignments | Where-Object { $_.target.groupId -eq $Group.Id }) {
                    $configCount++
                }
            }
            
            $configNode = New-TreeViewNode -Header "Configurations ($configCount)" -Tag @{ Type = "GroupConfigs"; Group = $Group } -Icon "[C]"
        }
        else {
            $configNode = New-TreeViewNode -Header "Configurations" -Tag @{ Type = "GroupConfigs"; Group = $Group } -Icon "[C]"
        }
        
        $groupNode.Items.Add($configNode)
        
        # For applications, first check if we have pre-loaded data
        $appCount = 0
        
        # If we have pre-loaded assignment data, we can get the count of applications right away
        if ($global:cachedData.AppAssignments.Count -gt 0) {
            # Check applications
            foreach ($appId in $global:cachedData.AppAssignments.Keys) {
                $assignments = $global:cachedData.AppAssignments[$appId]
                if ($assignments | Where-Object { $_.target.groupId -eq $Group.Id }) {
                    $appCount++
                }
            }
            
            $appNode = New-TreeViewNode -Header "Applications ($appCount)" -Tag @{ Type = "GroupApps"; Group = $Group } -Icon "[A]"
        }
        else {
            $appNode = New-TreeViewNode -Header "Applications" -Tag @{ Type = "GroupApps"; Group = $Group } -Icon "[A]"
        }
        
        $groupNode.Items.Add($appNode)
        
        # Add Properties node for consistency
        $propNode = New-TreeViewNode -Header "Properties" -Tag @{ Type = "GroupProperties"; Group = $Group } -Icon "[P]"
        $groupNode.Items.Add($propNode)
        
        # Add to the tree view
        $null = $TreeView.Items.Add($groupNode)
        
        return $groupNode
    }
    catch {
        Write-Host "Error in Add-GroupToTreeView: $_" -ForegroundColor Red
        return $null
    }
}

function Get-GroupProperties {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Group
    )
    
    # Convert group object to hashtable of key properties
    $properties = @{
        "Group Name" = $Group.DisplayName
        "Description" = if ($Group.Description) { $Group.Description } else { "N/A" }
        "Group Type" = if ($Group.GroupTypes) { $Group.GroupTypes -join ", " } else { "Security" }
        "Mail Enabled" = $Group.MailEnabled
        "Security Enabled" = $Group.SecurityEnabled
        "Created" = $Group.CreatedDateTime
        "Last Modified" = $Group.LastModifiedDateTime
        "ID" = $Group.Id
    }
    
    return $properties
}

function Get-GroupConfigurations {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Group
    )
    
    $configurations = @()
    
    try {
        Write-Host "Searching configurations for group $($Group.DisplayName) with ID $($Group.Id)"
        $groupId = $Group.Id
        
        # Check if we have pre-loaded configuration policies and their assignments
        if ($null -ne $global:cachedData.ConfigPolicies -and $global:cachedData.ConfigPolicyAssignments.Count -gt 0) {
            Write-Host "Using pre-loaded config policy assignments data"
            
            # First check configuration policies
            foreach ($policy in $global:cachedData.ConfigPolicies) {
                # Check if we have cached assignments for this policy
                if ($global:cachedData.ConfigPolicyAssignments.ContainsKey($policy.Id)) {
                    $assignments = $global:cachedData.ConfigPolicyAssignments[$policy.Id]
                    
                    # Check if our group is in the assignments
                    $groupAssignment = $assignments | Where-Object { $_.target.groupId -eq $groupId }
                    
                    if ($groupAssignment) {
                        Write-Host "Group is assigned to policy $($policy.Name)" -ForegroundColor Green
                        
                        $configurations += @{
                            Name = $policy.Name
                            Description = $policy.Description
                            Type = "Configuration Policy"
                            PolicyId = $policy.Id
                        }
                    }
                }
                else {
                    # We don't have cached assignments for this policy, fetch them directly
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($policy.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.ConfigPolicyAssignments[$policy.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to policy $($policy.Name)" -ForegroundColor Green
                                
                                $configurations += @{
                                    Name = $policy.Name
                                    Description = $policy.Description
                                    Type = "Configuration Policy"
                                    PolicyId = $policy.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for policy $($policy.Name): $_" -ForegroundColor Red
                    }
                }
            }
        }
        else {
            # We don't have pre-loaded data, fetch config policies
            if ($null -eq $global:cachedData.ConfigPolicies) {
                try {
                    Write-Host "Fetching configuration policies..."
                    $global:cachedData.ConfigPolicies = Get-MgBetaDeviceManagementConfigurationPolicy -All -ErrorAction Stop
                }
                catch {
                    Write-Host "Error fetching configuration policies: $_" -ForegroundColor Red
                }
            }
            
            # Process config policies if fetched successfully
            if ($null -ne $global:cachedData.ConfigPolicies) {
                foreach ($policy in $global:cachedData.ConfigPolicies) {
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$($policy.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.ConfigPolicyAssignments[$policy.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to policy $($policy.Name)" -ForegroundColor Green
                                
                                $configurations += @{
                                    Name = $policy.Name
                                    Description = $policy.Description
                                    Type = "Configuration Policy"
                                    PolicyId = $policy.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for policy $($policy.Name): $_" -ForegroundColor Red
                    }
                }
            }
        }
        
        # Check if we have pre-loaded compliance policies and their assignments
        if ($null -ne $global:cachedData.CompliancePolicies -and $global:cachedData.CompliancePolicyAssignments.Count -gt 0) {
            Write-Host "Using pre-loaded compliance policy assignments data"
            
            # Check compliance policies
            foreach ($policy in $global:cachedData.CompliancePolicies) {
                # Check if we have cached assignments for this policy
                if ($global:cachedData.CompliancePolicyAssignments.ContainsKey($policy.Id)) {
                    $assignments = $global:cachedData.CompliancePolicyAssignments[$policy.Id]
                    
                    # Check if our group is in the assignments
                    $groupAssignment = $assignments | Where-Object { $_.target.groupId -eq $groupId }
                    
                    if ($groupAssignment) {
                        Write-Host "Group is assigned to compliance policy $($policy.DisplayName)" -ForegroundColor Green
                        
                        $configurations += @{
                            Name = $policy.DisplayName
                            Description = $policy.Description
                            Type = "Compliance Policy"
                            PolicyId = $policy.Id
                        }
                    }
                }
                else {
                    # We don't have cached assignments for this policy, fetch them directly
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies('$($policy.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.CompliancePolicyAssignments[$policy.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to compliance policy $($policy.DisplayName)" -ForegroundColor Green
                                
                                $configurations += @{
                                    Name = $policy.DisplayName
                                    Description = $policy.Description
                                    Type = "Compliance Policy"
                                    PolicyId = $policy.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for compliance policy $($policy.DisplayName): $_" -ForegroundColor Red
                    }
                }
            }
        }
        else {
            # We don't have pre-loaded data, fetch compliance policies
            if ($null -eq $global:cachedData.CompliancePolicies) {
                try {
                    Write-Host "Fetching compliance policies..."
                    $global:cachedData.CompliancePolicies = Get-MgBetaDeviceManagementCompliancePolicy -All -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "Error fetching compliance policies: $_" -ForegroundColor Red
                }
            }
            
            # Process compliance policies if fetched successfully
            if ($null -ne $global:cachedData.CompliancePolicies) {
                foreach ($policy in $global:cachedData.CompliancePolicies) {
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies('$($policy.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.CompliancePolicyAssignments[$policy.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to compliance policy $($policy.DisplayName)" -ForegroundColor Green
                                
                                $configurations += @{
                                    Name = $policy.DisplayName
                                    Description = $policy.Description
                                    Type = "Compliance Policy"
                                    PolicyId = $policy.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for compliance policy $($policy.DisplayName): $_" -ForegroundColor Red
                    }
                }
            }
        }
        
        # Check if we have pre-loaded device configs and their assignments
        if ($null -ne $global:cachedData.DeviceConfigs -and $global:cachedData.DeviceConfigAssignments.Count -gt 0) {
            Write-Host "Using pre-loaded device config assignments data"
            
            # Check device configurations
            foreach ($config in $global:cachedData.DeviceConfigs) {
                # Check if we have cached assignments for this config
                if ($global:cachedData.DeviceConfigAssignments.ContainsKey($config.Id)) {
                    $assignments = $global:cachedData.DeviceConfigAssignments[$config.Id]
                    
                    # Check if our group is in the assignments
                    $groupAssignment = $assignments | Where-Object { $_.target.groupId -eq $groupId }
                    
                    if ($groupAssignment) {
                        Write-Host "Group is assigned to device config $($config.DisplayName)" -ForegroundColor Green
                        
                        $configurations += @{
                            Name = $config.DisplayName
                            Description = $config.Description
                            Type = "Device Configuration"
                            PolicyId = $config.Id
                        }
                    }
                }
                else {
                    # We don't have cached assignments for this config, fetch them directly
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('$($config.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.DeviceConfigAssignments[$config.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to device config $($config.DisplayName)" -ForegroundColor Green
                                
                                $configurations += @{
                                    Name = $config.DisplayName
                                    Description = $config.Description
                                    Type = "Device Configuration"
                                    PolicyId = $config.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for device config $($config.DisplayName): $_" -ForegroundColor Red
                    }
                }
            }
        }
        else {
            # We don't have pre-loaded data, fetch device configurations
            if ($null -eq $global:cachedData.DeviceConfigs) {
                try {
                    Write-Host "Fetching device configurations..."
                    $global:cachedData.DeviceConfigs = Get-MgBetaDeviceManagementDeviceConfiguration -All -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "Error fetching device configurations: $_" -ForegroundColor Red
                }
            }
            
            # Process device configs if fetched successfully
            if ($null -ne $global:cachedData.DeviceConfigs) {
                foreach ($config in $global:cachedData.DeviceConfigs) {
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations('$($config.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.DeviceConfigAssignments[$config.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to device config $($config.DisplayName)" -ForegroundColor Green
                                
                                $configurations += @{
                                    Name = $config.DisplayName
                                    Description = $config.Description
                                    Type = "Device Configuration"
                                    PolicyId = $config.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for device config $($config.DisplayName): $_" -ForegroundColor Red
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Error in Get-GroupConfigurations: $_" -ForegroundColor Red
    }
    
    return $configurations
}

function Get-GroupApplications {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Group
    )
    
    $applications = @()
    
    try {
        Write-Host "Searching applications for group $($Group.DisplayName) with ID $($Group.Id)"
        $groupId = $Group.Id
        
        # Check if we have pre-loaded apps and their assignments
        if ($null -ne $global:cachedData.Apps -and $global:cachedData.AppAssignments.Count -gt 0) {
            Write-Host "Using pre-loaded app assignments data"
            
            # Check mobile apps
            foreach ($app in $global:cachedData.Apps) {
                # Check if we have cached assignments for this app
                if ($global:cachedData.AppAssignments.ContainsKey($app.Id)) {
                    $assignments = $global:cachedData.AppAssignments[$app.Id]
                    
                    # Check if our group is in the assignments
                    $groupAssignment = $assignments | Where-Object { $_.target.groupId -eq $groupId }
                    
                    if ($groupAssignment) {
                        Write-Host "Group is assigned to app $($app.DisplayName)" -ForegroundColor Green
                        
                        # Set a default value for AppType if it's null
                        $appType = "Application"
                        if ($app.AppType) {
                            $appType = $app.AppType
                        }
                        
                        $applications += @{
                            Name = $app.DisplayName
                            Description = $app.Description
                            Type = $appType
                            AppId = $app.Id
                        }
                    }
                }
                else {
                    # We don't have cached assignments for this app, fetch them directly
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps('$($app.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.AppAssignments[$app.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to app $($app.DisplayName)" -ForegroundColor Green
                                
                                # Set a default value for AppType if it's null
                                $appType = "Application"
                                if ($app.AppType) {
                                    $appType = $app.AppType
                                }
                                
                                $applications += @{
                                    Name = $app.DisplayName
                                    Description = $app.Description
                                    Type = $appType
                                    AppId = $app.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for app $($app.DisplayName): $_" -ForegroundColor Red
                    }
                }
            }
        }
        else {
            # We don't have pre-loaded data, fetch mobile apps
            if ($null -eq $global:cachedData.Apps) {
                try {
                    Write-Host "Fetching mobile apps..."
                    $global:cachedData.Apps = Get-MgBetaDeviceAppManagementMobileApp -All -ErrorAction Stop
                }
                catch {
                    Write-Host "Error fetching mobile apps: $_" -ForegroundColor Red
                }
            }
            
            # Process apps if fetched successfully
            if ($null -ne $global:cachedData.Apps) {
                foreach ($app in $global:cachedData.Apps) {
                    try {
                        $assignmentsUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps('$($app.Id)')/assignments"
                        $assignments = Invoke-MgGraphRequest -Method GET -Uri $assignmentsUri -ErrorAction SilentlyContinue
                        
                        if ($assignments -and $assignments.Value) {
                            # Cache these assignments for future use
                            $global:cachedData.AppAssignments[$app.Id] = $assignments.Value
                            
                            # Check if our group is in the assignments
                            $groupAssignment = $assignments.Value | Where-Object { $_.target.groupId -eq $groupId }
                            
                            if ($groupAssignment) {
                                Write-Host "Group is assigned to app $($app.DisplayName)" -ForegroundColor Green
                                
                                # Set a default value for AppType if it's null
                                $appType = "Application"
                                if ($app.AppType) {
                                    $appType = $app.AppType
                                }
                                
                                $applications += @{
                                    Name = $app.DisplayName
                                    Description = $app.Description
                                    Type = $appType
                                    AppId = $app.Id
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Error checking assignments for app $($app.DisplayName): $_" -ForegroundColor Red
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Error in Get-GroupApplications: $_" -ForegroundColor Red
    }
    
    return $applications
} 