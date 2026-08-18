#Requires -Version 7.0
<#
.SYNOPSIS
Graph App Role Manager - GUI helper for assigning Microsoft Graph application
permissions (app roles) to a Managed Identity or another service principal.

.DESCRIPTION
Windows Forms GUI that:
- Connects to Microsoft Graph with delegated administrative scopes.
- Searches a target service principal by display name.
- Lists Microsoft Graph application permissions.
- Shows permissions already assigned to the target.
- Assigns one or more selected application permissions.
- Removes one or more selected application permissions after confirmation.
- Avoids duplicate assignments and writes an execution log.

REQUIREMENTS
- Windows
- PowerShell 7+
- Microsoft.Graph.Authentication
- Microsoft.Graph.Applications

The signed-in administrator must be allowed to create app role assignments.
Typical delegated scopes requested by this tool:
- Application.Read.All
- AppRoleAssignment.ReadWrite.All

SECURITY
Grant only the permissions needed by the target workload.

.PARAMETER NoGui
Builds the Windows Forms interface without displaying it. Intended for smoke tests.

.PARAMETER LogLevel
Minimum log level written to the activity log and console output. Defaults to INFO.
#>

param(
    [switch]$NoGui,

    [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
    [string]$LogLevel = 'INFO'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$modulePath = Join-Path $PSScriptRoot 'Graph-App-Role-Manager.psm1'
if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "Graph App Role Manager module not found: $modulePath"
}

Import-Module $modulePath -Force
Start-GraphAppRoleManager -NoGui:$NoGui.IsPresent -LogLevel $LogLevel
