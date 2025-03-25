# TreeView helper functions for Intune Explorer

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
        [switch]$IsExpanded = $false
    )
    
    $node = New-Object System.Windows.Controls.TreeViewItem
    
    if ([string]::IsNullOrEmpty($Icon)) {
        # Simple text header
        $node.Header = $Header
    }
    else {
        # Create header with icon
        $headerStack = New-Object System.Windows.Controls.StackPanel
        $headerStack.Orientation = "Horizontal"
        
        $iconText = New-Object System.Windows.Controls.TextBlock
        $iconText.Text = $Icon
        $iconText.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
        $headerStack.Children.Add($iconText)
        
        $nameText = New-Object System.Windows.Controls.TextBlock
        $nameText.Text = $Header
        
        # Apply custom color if specified
        if ($null -ne $HeaderColor) {
            $nameText.Foreground = $HeaderColor
        }
        
        $headerStack.Children.Add($nameText)
        
        $node.Header = $headerStack
    }
    
    $node.Tag = $Tag
    $node.IsExpanded = $IsExpanded
    
    return $node
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
        [System.Windows.Controls.TreeViewItem]$Node,
        
        [Parameter(Mandatory = $true)]
        [System.Windows.Controls.StackPanel]$DetailsPanel
    )
    
    try {
        $DetailsPanel.Children.Clear()
        
        $nodeType = $Node.Tag.Type
        $device = $Node.Tag.Device
        
        switch ($nodeType) {
            "DeviceProperties" {
                $properties = Get-DeviceProperties -Device $device
                
                foreach ($prop in $properties.GetEnumerator()) {
                    $propStack = New-Object System.Windows.Controls.StackPanel
                    $propStack.Orientation = "Horizontal"
                    $propStack.Margin = New-Object System.Windows.Thickness(0, 0, 0, 5)
                    
                    $labelBlock = New-Object System.Windows.Controls.TextBlock
                    $labelBlock.Text = "$($prop.Key): "
                    $labelBlock.FontWeight = "Bold"
                    $labelBlock.Margin = New-Object System.Windows.Thickness(0, 0, 5, 0)
                    
                    $valueBlock = New-Object System.Windows.Controls.TextBlock
                    $valueBlock.Text = $prop.Value
                    $valueBlock.TextWrapping = "Wrap"
                    
                    $propStack.Children.Add($labelBlock)
                    $propStack.Children.Add($valueBlock)
                    
                    $DetailsPanel.Children.Add($propStack)
                }
            }
            "DevicePolicies" {
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.Text = "Loading policies..."
                $DetailsPanel.Children.Add($textBlock)
                
                # Here you would add code to fetch and display policies
            }
            "DeviceApps" {
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.Text = "Loading applications..."
                $DetailsPanel.Children.Add($textBlock)
                
                # Here you would add code to fetch and display applications
            }
            "DeviceScripts" {
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.Text = "Loading scripts..."
                $DetailsPanel.Children.Add($textBlock)
                
                # Here you would add code to fetch and display scripts
            }
            default {
                $textBlock = New-Object System.Windows.Controls.TextBlock
                $textBlock.Text = "Select a category to view details."
                $DetailsPanel.Children.Add($textBlock)
            }
        }
    }
    catch {
        $errorBlock = New-Object System.Windows.Controls.TextBlock
        $errorBlock.Text = "Error displaying details: $_"
        $errorBlock.Foreground = "Red"
        $errorBlock.TextWrapping = "Wrap"
        $DetailsPanel.Children.Clear()
        $DetailsPanel.Children.Add($errorBlock)
    }
} 