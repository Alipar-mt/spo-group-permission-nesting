# spo-group-permission-nesting

A PowerShell automation script for creating SharePoint groups with permission levels and integrating Entra ID security groups.

## Overview

This script automates the process of:
- Creating SharePoint groups in specified sites
- Assigning permission levels to these groups
- Adding Entra ID (Azure AD) security groups as members to the SharePoint groups

Perfect for bulk operations when setting up permissions across multiple SharePoint sites with consistent security group mappings.

## Features

- ✅ **Bulk Processing**: Process multiple sites and groups from a single CSV file
- ✅ **Error Handling**: Comprehensive error handling with detailed logging
- ✅ **Duplicate Prevention**: Checks for existing groups before creating new ones
- ✅ **Flexible Input**: Automatically detects CSV column structure
- ✅ **Progress Tracking**: Real-time feedback with colored console output
- ✅ **Connection Management**: Proper handling of SharePoint and Microsoft Graph connections

## Prerequisites

### PowerShell Modules
Install the required PowerShell modules:

```powershell
# Install PnP PowerShell for SharePoint operations
Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser

# Install Microsoft Graph PowerShell for Entra ID operations
Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser
```

### Permissions Required
Ensure you have the following permissions:
- **SharePoint**: Site Collection Administrator or Site Owner rights on target sites
- **Entra ID**: Directory Reader permissions to read security groups
- **Microsoft Graph**: `Group.Read.All` scope (granted during script execution)

## CSV File Format

Create a CSV file with the following columns:

| Column A | Column B | Column C | Column D |
|----------|----------|----------|----------|
| SharePoint Site URL | SharePoint Group Name | Permission Level | Entra Security Group Name |

### Example CSV:
```csv
SiteUrl,GroupName,PermissionLevel,EntraGroupName
https://contoso.sharepoint.com/sites/marketing,Marketing Contributors,Contribute,Marketing-Team-Security-Group
https://contoso.sharepoint.com/sites/finance,Finance Readers,Read,Finance-Team-Security-Group
https://contoso.sharepoint.com/sites/hr,HR Full Access,Full Control,HR-Admin-Security-Group
https://contoso.sharepoint.com/sites/projects,Project Members,Contribute,Project-Team-Members
```

### Column Details:
- **Column A (SiteUrl)**: Full URL to the SharePoint site
- **Column B (GroupName)**: Name for the new SharePoint group to be created
- **Column C (PermissionLevel)**: SharePoint permission level (Read, Contribute, Full Control, or custom levels)
- **Column D (EntraGroupName)**: Display name of the Entra ID security group to add as members

## Usage

### Basic Usage
```powershell
.\New-SharePointGroupsWithPermissions.ps1 -CsvPath "C:\path\to\your\groups.csv" -TenantUrl "https://yourtenant.sharepoint.com"
```

### Parameters
- **`-CsvPath`** (Required): Path to the CSV file containing group configuration
- **`-TenantUrl`** (Required): Your SharePoint tenant URL

### Example Execution
```powershell
# Navigate to the script directory
cd "C:\Scripts\spo-group-permission-nesting"

# Run the script
.\New-SharePointGroupsWithPermissions.ps1 -CsvPath ".\groups-config.csv" -TenantUrl "https://contoso.sharepoint.com"
```

## Script Workflow

1. **Validation**: Checks if CSV file exists and is readable
2. **Authentication**: Connects to Microsoft Graph and SharePoint Online
3. **Processing**: For each row in the CSV:
   - Connects to the specified SharePoint site
   - Creates the SharePoint group (if it doesn't exist)
   - Assigns the specified permission level
   - Looks up the Entra security group
   - Adds the Entra group as a member of the SharePoint group
4. **Cleanup**: Disconnects all connections and provides summary

## Output and Logging

The script provides real-time feedback with color-coded messages:
- 🟢 **Green**: Success messages
- 🟡 **Yellow**: Warning messages  
- 🔴 **Red**: Error messages
- ⚪ **White**: General information

### Example Output:
```
[2024-01-15 10:30:15] [INFO] Starting SharePoint Groups Creation Script
[2024-01-15 10:30:16] [INFO] Reading CSV file: C:\temp\groups.csv
[2024-01-15 10:30:17] [INFO] Connecting to Microsoft Graph...
[2024-01-15 10:30:18] [SUCCESS] SharePoint group 'Marketing Contributors' created successfully
[2024-01-15 10:30:19] [SUCCESS] Permission level 'Contribute' assigned successfully
[2024-01-15 10:30:20] [SUCCESS] Entra group 'Marketing-Team-Security-Group' added successfully
```

## Common Permission Levels

Standard SharePoint permission levels you can use:
- `Read` - View only access
- `Contribute` - Add, edit, and delete items
- `Edit` - Add, edit, delete, and manage lists
- `Design` - Create lists and document libraries, edit pages
- `Full Control` - Complete control over the site

## Troubleshooting

### Common Issues

**Issue**: "SharePoint group already exists"
- **Solution**: The script skips existing groups. This is expected behavior to prevent duplicates.

**Issue**: "Entra group not found"
- **Solution**: Verify the exact display name of the Entra security group. Names are case-sensitive.

**Issue**: "Permission level not found"
- **Solution**: Ensure the permission level exists in the target SharePoint site. Custom permission levels must be created first.

**Issue**: "Access denied" errors
- **Solution**: Verify you have the required permissions on both SharePoint sites and Entra ID.

### Debug Tips
1. Test with a small CSV file first
2. Verify SharePoint site URLs are accessible
3. Confirm Entra group names are exact matches
4. Check that permission levels exist in target sites

## Security Considerations

- The script uses interactive authentication for enhanced security
- Connections are properly closed after each operation
- No credentials are stored or logged
- Minimum required permissions are requested

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/improvement`)
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If you encounter issues or have questions:
1. Check the [Issues](https://github.com/yourusername/spo-group-permission-nesting/issues) section
2. Create a new issue with detailed information about your problem
3. Include relevant error messages and your CSV file structure (with sensitive data removed)

## Changelog

### Version 1.0.0
- Initial release
- CSV-based bulk processing
- SharePoint group creation
- Permission level assignment
- Entra ID security group integration
- Comprehensive error handling and logging
