# Windows setup

The cross-platform GUI is the recommended interface, including on Windows. This native
PowerShell/Windows Forms edition remains available for administrators who prefer a
PowerShell-only Windows experience. It follows the same three-step workflow and visual
structure as the cross-platform edition.

## Requirements

- Windows 10/11 or Windows Server with a desktop session.
- PowerShell 7 or later.
- `Microsoft.Graph.Authentication`.
- `Microsoft.Graph.Applications`.
- An administrator account authorized to manage app-role assignments.

## Start

From the repository root:

```powershell
pwsh -File .\windows\Graph-App-Role-Manager.ps1
```

The launcher imports the local PowerShell module next to it and starts the Windows
Forms application. This command remains the normal quick-start entry point.

Optional module usage:

```powershell
Import-Module .\windows\Graph-App-Role-Manager.psm1
Start-GraphAppRoleManager
```

Optional log filtering:

```powershell
pwsh -File .\windows\Graph-App-Role-Manager.ps1 -LogLevel WARNING
```

```powershell
Start-GraphAppRoleManager -LogLevel WARNING
```

`-LogLevel` accepts `INFO`, `SUCCESS`, `WARNING`, or `ERROR`. The default is `INFO`,
which preserves the current informational activity log.

If a required Microsoft Graph module is missing, the interface displays the exact module
names and asks whether it may install them for the current user. Declining leaves the
machine unchanged and stops the connection.

The module check safely handles zero, one, or several missing modules. This is important
with PowerShell strict mode, which otherwise treats a single result differently from a
collection.

Modules can also be installed beforehand:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Applications -Scope CurrentUser
```

## Authentication

The tool calls `Connect-MgGraph` with:

- `Application.Read.All`
- `AppRoleAssignment.ReadWrite.All`

Authentication is process-scoped. The Graph session is disconnected when the form closes.

## Usage

1. Select **Connect to Microsoft Graph**.
2. Wait while the interface automatically loads the tenant's managed identities and
   service principals.
3. Select the correct identity from the list, or type a name, application ID, or object ID
   to filter the list locally.
4. Filter the Graph permission catalog.
5. Check permissions and select **Assign selected permissions**.
6. To remove assignments, select rows in the current-assignment table and choose
   **Remove selected**.

Every change displays the target and permissions in a confirmation dialog.

The permissions catalog and current-assignment table share the lower window equally, and
remain balanced when the window is resized.
