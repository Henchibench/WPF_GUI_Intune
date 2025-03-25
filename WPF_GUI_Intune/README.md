# Intune Explorer

A modern WPF-based GUI for exploring Microsoft Intune data. This application connects to Microsoft Graph API and allows you to search for devices and view related information in a tree-structured format with a clean, user-friendly interface.

## Features

- Modern UI with rounded corners, consistent styling, and clean design
- Connect to Microsoft Graph with appropriate permissions
- Search for devices by name, serial number, or ID
- View device properties in a hierarchical tree structure with intuitive icons
- Explore device information in an organized card-based layout:
  - Basic Information (device details and compliance status)
  - Management Status (management state, enrollment type, ownership)
  - User Information (display name, email, UPN, job title, etc.)
  - Important Dates (enrollment and last sync times)
  - Operating System (OS type, version, and device type)
- Additional detailed sections:
  - Configuration and compliance policies
  - Applications
  - Remediation and platform scripts
- Responsive layout that adapts to window size
- Color-coded compliance status (green for compliant, red for non-compliant)
- Automatically loads data when needed (lazy loading)
- Automatic installation of required modules using the `-InstallMissing` parameter

## Screenshots

*[Screenshot images would be placed here]*

## Prerequisites

- PowerShell 5.1 or higher
- Microsoft Graph PowerShell modules:
  - Microsoft.Graph.Intune
  - Microsoft.Graph.Beta.DeviceManagement
  - Microsoft.Graph.Beta.Users
  - Microsoft.Graph.Beta.DeviceManagement.Administration
  - Microsoft.Graph.Beta.Applications

## Installation

### Option 1: Manually install the required modules
Install the required Microsoft Graph PowerShell modules:

```powershell
Install-Module Microsoft.Graph.Intune -Force
Install-Module Microsoft.Graph.Beta.DeviceManagement -Force
Install-Module Microsoft.Graph.Beta.Users -Force
Install-Module Microsoft.Graph.Beta.DeviceManagement.Administration -Force
Install-Module Microsoft.Graph.Beta.Applications -Force
```

### Option 2: Let the application install missing modules
Run the application with the `-InstallMissing` parameter, which will check for and install any missing modules:

```powershell
.\IntuneExplorer.ps1 -InstallMissing
```

## Usage

1. Run the PowerShell script (with or without the installation parameter):

```powershell
# Standard launch
.\IntuneExplorer.ps1

# Launch with automatic module installation
.\IntuneExplorer.ps1 -InstallMissing
```

2. Click "Connect to Microsoft Graph" to authenticate.
3. Enter a device name, serial number, or ID in the search box and click "Search" or press Enter.
4. Navigate the tree view to explore device information:
   - Click on a device to see basic information
   - Click on "Properties" to see detailed device information in a card-based layout
   - Click on "Policies" to view assigned configuration and compliance policies
   - Click on "Applications" to view installed applications
   - Click on "Scripts" to view assigned remediation and platform scripts
5. Click "Refresh Data" to update the cached information from Intune

## Recent Improvements

- Enhanced visual layout with card-based sections for better organization
- Improved device property display with categorized information
- Added user information directly in the Properties view
- Fixed grid layout for better space utilization
- Implemented consistent section heights for better visual alignment
- Added color-coded compliance status
- Improved error handling for device type detection
- Added scrollable sections for viewing extensive information

## Development Notes

The application architecture includes:
- IntuneExplorer.ps1: Main application script with UI and core functionality
- TreeViewHelper.ps1: Helper functions for managing the tree view component
- ErrorHandler.ps1: Functions for error handling and reporting
- ModuleManager.ps1: Functions for managing required PowerShell modules

## Data Retrieval

The application uses the following Microsoft Graph Beta endpoints:

- `Get-MgBetaDeviceManagementManagedDevice` - Fetch device information
- `Get-MgBetaUser` - Fetch user information
- `Get-MgBetaDeviceAppManagementMobileApp` - Fetch application information
- `Get-MgBetaDeviceManagementConfigurationPolicy` - Fetch configuration policies
- `Get-MgBetaDeviceManagementCompliancePolicy` - Fetch compliance policies
- `Get-MgBetaDeviceManagementDeviceHealthScript` - Fetch remediation scripts
- `Get-MgBetaDeviceManagementScript` - Fetch platform scripts

## Note

This application uses the Microsoft Graph Beta API, which may change without notice. Always check for updates to ensure compatibility with the latest API changes. 