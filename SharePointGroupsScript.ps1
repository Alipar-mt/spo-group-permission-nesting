# SharePoint Groups Creation and Entra Security Group Assignment Script
# Requires: PnP.PowerShell and Microsoft.Graph modules

#Requires -Modules PnP.PowerShell, Microsoft.Graph.Authentication, Microsoft.Graph.Groups

<#
.SYNOPSIS
Creates SharePoint groups with specified permission levels and adds Entra security groups as members.

.DESCRIPTION
This script reads from a CSV file with the following columns:
- Column A: SharePoint Site URL
- Column B: SharePoint Group Name to create
- Column C: Permission Level to assign to the group
- Column D: Entra Security Group Name/ID to add to the SharePoint group

.PARAMETER CsvPath
Path to the CSV file containing the configuration data

.PARAMETER TenantUrl
SharePoint tenant URL (e.g., https://contoso.sharepoint.com)

.EXAMPLE
.\SharePointGroupsScript.ps1 -CsvPath "C:\temp\groups.csv" -TenantUrl "https://contoso.sharepoint.com"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$true)]
    [string]$TenantUrl
)

# Function to write colored output
function Write-ColorOutput($ForegroundColor) {
    $fc = $host.UI.RawUI.ForegroundColor
    $host.UI.RawUI.ForegroundColor = $ForegroundColor
    if ($args) {
        Write-Output $args
    } else {
        $input | Write-Output
    }
    $host.UI.RawUI.ForegroundColor = $fc
}

# Function to log messages
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "ERROR" { Write-ColorOutput Red $logMessage }
        "WARNING" { Write-ColorOutput Yellow $logMessage }
        "SUCCESS" { Write-ColorOutput Green $logMessage }
        default { Write-Output $logMessage }
    }
}

# Function to get Entra group ID by name
function Get-EntraGroupId {
    param([string]$GroupName)
    
    try {
        $group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction Stop
        if ($group) {
            return $group.Id
        } else {
            $group = Get-MgGroup -Filter "mailNickname eq '$GroupName'" -ErrorAction Stop
            if ($group) {
                return $group.Id
            }
        }
        return $null
    }
    catch {
        Write-Log "Error finding Entra group '$GroupName': $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

# Main script execution
try {
    Write-Log "Starting SharePoint Groups Creation Script"
    
    # Check if CSV file exists
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found at path: $CsvPath"
    }
    
    # Read CSV file
    Write-Log "Reading CSV file: $CsvPath"
    $csvData = Import-Csv $CsvPath
    
    if ($csvData.Count -eq 0) {
        throw "CSV file is empty or could not be read properly"
    }
    
    # Get column names (assuming first 4 columns)
    $columns = ($csvData | Get-Member -MemberType NoteProperty | Select-Object -First 4).Name
    Write-Log "Detected columns: $($columns -join ', ')"
    
    # Connect to Microsoft Graph
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Group.Read.All" -NoWelcome
    
    # Process each row
    $successCount = 0
    $errorCount = 0
    
    foreach ($row in $csvData) {
        $siteUrl = $row.$($columns[0])
        $groupName = $row.$($columns[1])
        $permissionLevel = $row.$($columns[2])
        $entraGroupName = $row.$($columns[3])
        
        Write-Log "Processing: Site='$siteUrl', Group='$groupName', Permission='$permissionLevel', EntraGroup='$entraGroupName'"
        
        # Skip empty rows
        if ([string]::IsNullOrWhiteSpace($siteUrl) -or [string]::IsNullOrWhiteSpace($groupName)) {
            Write-Log "Skipping row with empty Site URL or Group Name" -Level "WARNING"
            continue
        }
        
        try {
            # Connect to SharePoint site
            Write-Log "Connecting to SharePoint site: $siteUrl"
            Connect-PnPOnline -Url $siteUrl -Interactive
            
            # Check if SharePoint group already exists
            $existingGroup = Get-PnPGroup -Identity $groupName -ErrorAction SilentlyContinue
            if ($existingGroup) {
                Write-Log "SharePoint group '$groupName' already exists, skipping creation" -Level "WARNING"
            } else {
                # Create SharePoint group
                Write-Log "Creating SharePoint group: $groupName"
                $spGroup = New-PnPGroup -Title $groupName -Description "Created by automation script"
                Write-Log "SharePoint group '$groupName' created successfully" -Level "SUCCESS"
            }
            
            # Assign permission level to the group
            if (-not [string]::IsNullOrWhiteSpace($permissionLevel)) {
                Write-Log "Assigning permission level '$permissionLevel' to group '$groupName'"
                Set-PnPGroupPermissions -Identity $groupName -AddRole $permissionLevel
                Write-Log "Permission level '$permissionLevel' assigned successfully" -Level "SUCCESS"
            }
            
            # Add Entra security group as member
            if (-not [string]::IsNullOrWhiteSpace($entraGroupName)) {
                Write-Log "Looking up Entra group: $entraGroupName"
                $entraGroupId = Get-EntraGroupId -GroupName $entraGroupName
                
                if ($entraGroupId) {
                    Write-Log "Adding Entra group '$entraGroupName' to SharePoint group '$groupName'"
                    
                    # Get the Entra group details for the login name
                    $entraGroup = Get-MgGroup -GroupId $entraGroupId
                    $loginName = "c:0t.c|tenant|$entraGroupId"
                    
                    # Add the Entra group to SharePoint group
                    Add-PnPGroupMember -LoginName $loginName -Identity $groupName
                    Write-Log "Entra group '$entraGroupName' added to SharePoint group '$groupName' successfully" -Level "SUCCESS"
                } else {
                    Write-Log "Entra group '$entraGroupName' not found" -Level "ERROR"
                    $errorCount++
                    continue
                }
            }
            
            $successCount++
            Write-Log "Row processed successfully" -Level "SUCCESS"
            
        }
        catch {
            Write-Log "Error processing row: $($_.Exception.Message)" -Level "ERROR"
            $errorCount++
        }
        finally {
            # Disconnect from current site
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        
        Write-Log "---"
    }
    
    # Summary
    Write-Log "Script execution completed" -Level "SUCCESS"
    Write-Log "Successful operations: $successCount" -Level "SUCCESS"
    Write-Log "Failed operations: $errorCount" -Level $(if($errorCount -gt 0) { "ERROR" } else { "SUCCESS" })
    
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level "ERROR"
    exit 1
}
finally {
    # Cleanup connections
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}

Write-Log "Script finished"
