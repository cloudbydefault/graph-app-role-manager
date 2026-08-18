#Requires -Version 7.0
<#
.SYNOPSIS
Graph App Role Manager module for the native Windows GUI.

.DESCRIPTION
Provides Start-GraphAppRoleManager for launching the Windows Forms interface that
manages Microsoft Graph application permissions for Managed Identities and service
principals.
#>

Set-StrictMode -Version Latest

$script:LogLevel = 'INFO'
$script:LogLevelRank = @{
    INFO    = 0
    SUCCESS = 1
    WARNING = 2
    ERROR   = 3
}

function Initialize-GraphAppRoleManagerModuleState {
    $script:GraphConnected = $false
    $script:GraphServicePrincipal = $null
    $script:TargetServicePrincipal = $null
    $script:AllTargetServicePrincipals = @()
    $script:SuppressIdentitySelectionEvent = $false
    $script:AllApplicationRoles = @()
    $script:VisibleApplicationRoles = @()

    $script:form = $null
    $script:rootLayout = $null
    $script:headerPanel = $null
    $script:titlePanel = $null
    $script:lblTitle = $null
    $script:lblSubtitle = $null
    $script:btnConnect = $null
    $script:tabs = $null
    $script:tabAssign = $null
    $script:mainLayout = $null
    $script:grpIdentity = $null
    $script:identityLayout = $null
    $script:lblIdentitySearch = $null
    $script:txtIdentityName = $null
    $script:btnSearchIdentity = $null
    $script:resultsLayout = $null
    $script:lblSearchResults = $null
    $script:cmbIdentityResults = $null
    $script:summaryPanel = $null
    $script:lblIdentityName = $null
    $script:lblIdentityNameValue = $null
    $script:lblObjectId = $null
    $script:lblObjectIdValue = $null
    $script:lblAppId = $null
    $script:lblAppIdValue = $null
    $script:splitPermissions = $null
    $script:grpPermissions = $null
    $script:permissionLayout = $null
    $script:txtPermissionFilter = $null
    $script:btnLoadPermissions = $null
    $script:checkedPermissions = $null
    $script:lblPermissionCount = $null
    $script:btnAssign = $null
    $script:grpCurrent = $null
    $script:assignmentLayout = $null
    $script:lblAssignedCount = $null
    $script:btnRefreshAssignments = $null
    $script:gridAssignments = $null
    $script:btnRemove = $null
    $script:tabLog = $null
    $script:txtLog = $null
    $script:colorCanvas = $null
    $script:colorHeader = $null
    $script:colorBlue = $null
    $script:colorGreen = $null
    $script:colorRed = $null
    $script:colorMuted = $null
}

function Reset-GraphAppRoleManagerState {
    if ($null -ne $script:form) {
        try {
            $script:form.Dispose()
        }
        catch {}
    }

    Initialize-GraphAppRoleManagerModuleState
}

Initialize-GraphAppRoleManagerModuleState

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

function Write-UiLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if ($script:LogLevelRank[$Level] -lt $script:LogLevelRank[$script:LogLevel]) {
        return
    }

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $line = "[$timestamp][$Level] $Message"

    if ($null -ne $script:txtLog) {
        $script:txtLog.AppendText($line + [Environment]::NewLine)
        $script:txtLog.SelectionStart = $script:txtLog.TextLength
        $script:txtLog.ScrollToCaret()
    }

    Write-Host $line
}


function Show-UiError {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-UiLog -Message $Message -Level 'ERROR'
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        'Graph App Role Manager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Set-BusyState {
    param([bool]$Busy)

    $script:form.UseWaitCursor = $Busy

    foreach ($control in @(
        $script:btnConnect,
        $script:btnSearchIdentity,
        $script:btnLoadPermissions,
        $script:btnAssign,
        $script:btnRemove,
        $script:btnRefreshAssignments
    )) {
        if ($null -ne $control) {
            $control.Enabled = -not $Busy
        }
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Ensure-GraphModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Applications'
    )

    # Always preserve collection semantics. Without @(...), PowerShell unwraps a
    # single missing module to a scalar string and StrictMode rejects .Count.
    $missingModules = @(
        foreach ($moduleName in $requiredModules) {
            if (-not (Get-Module -ListAvailable -Name $moduleName)) {
                $moduleName
            }
        }
    )

    if ($missingModules.Count -gt 0) {
        $message = @"
The following PowerShell modules are missing:

$($missingModules -join [Environment]::NewLine)

Install them for the current user now?
"@

        $answer = [System.Windows.Forms.MessageBox]::Show(
            $message,
            'Missing Microsoft Graph modules',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            throw 'Required Microsoft Graph modules are not installed.'
        }

        foreach ($moduleName in $missingModules) {
            Write-UiLog -Message "Installing module: $moduleName"
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
        }
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
}

function Get-SelectedTargetServicePrincipal {
    if ($null -eq $script:cmbIdentityResults -or $null -eq $script:cmbIdentityResults.SelectedItem) {
        return $null
    }

    return $script:cmbIdentityResults.SelectedItem.Tag
}

function Update-IdentitySummary {
    param($ServicePrincipal)

    if ($null -eq $ServicePrincipal) {
        $script:lblIdentityNameValue.Text = '-'
        $script:lblObjectIdValue.Text = '-'
        $script:lblAppIdValue.Text = '-'
        return
    }

    $script:lblIdentityNameValue.Text = [string]$ServicePrincipal.DisplayName
    $script:lblObjectIdValue.Text = [string]$ServicePrincipal.Id
    $script:lblAppIdValue.Text = [string]$ServicePrincipal.AppId
}

function Load-GraphApplicationRoles {
    if (-not $script:GraphConnected) {
        throw 'Connect to Microsoft Graph first.'
    }

    Write-UiLog -Message 'Loading Microsoft Graph application permissions...'

    $script:GraphServicePrincipal = Get-MgServicePrincipal `
        -Filter "appId eq '00000003-0000-0000-c000-000000000000'" `
        -Property 'id,appId,displayName,appRoles' `
        -ConsistencyLevel eventual `
        -ErrorAction Stop |
        Select-Object -First 1

    if ($null -eq $script:GraphServicePrincipal) {
        throw 'The Microsoft Graph service principal could not be found.'
    }

    $script:AllApplicationRoles = @(
        $script:GraphServicePrincipal.AppRoles |
        Where-Object {
            $_.IsEnabled -eq $true -and
            $_.AllowedMemberTypes -contains 'Application' -and
            -not [string]::IsNullOrWhiteSpace($_.Value)
        } |
        Sort-Object Value
    )

    Apply-PermissionFilter

    Write-UiLog -Message (
        "Loaded {0} Microsoft Graph application permissions." -f
        $script:AllApplicationRoles.Count
    ) -Level 'SUCCESS'
}

function Apply-PermissionFilter {
    $filter = $script:txtPermissionFilter.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($filter)) {
        $script:VisibleApplicationRoles = @($script:AllApplicationRoles)
    }
    else {
        $script:VisibleApplicationRoles = @(
            $script:AllApplicationRoles |
            Where-Object {
                $_.Value -like "*$filter*" -or
                $_.DisplayName -like "*$filter*" -or
                $_.Description -like "*$filter*"
            }
        )
    }

    $checkedValues = @{}
    foreach ($index in $script:checkedPermissions.CheckedIndices) {
        $role = $script:checkedPermissions.Items[$index]
        $checkedValues[[string]$role.Value] = $true
    }

    $script:checkedPermissions.BeginUpdate()
    try {
        $script:checkedPermissions.Items.Clear()

        foreach ($role in $script:VisibleApplicationRoles) {
            $item = [pscustomobject]@{
                Value       = [string]$role.Value
                DisplayName = [string]$role.DisplayName
                Description = [string]$role.Description
                Id          = [guid]$role.Id
                Text        = "{0} — {1}" -f $role.Value, $role.DisplayName
            }

            $newIndex = $script:checkedPermissions.Items.Add($item)
            if ($checkedValues.ContainsKey($item.Value)) {
                $script:checkedPermissions.SetItemChecked($newIndex, $true)
            }
        }
    }
    finally {
        $script:checkedPermissions.EndUpdate()
    }

    $script:lblPermissionCount.Text = "{0} permission(s) displayed" -f $script:VisibleApplicationRoles.Count
}

function Refresh-CurrentAssignments {
    if ($null -eq $script:TargetServicePrincipal) {
        throw 'Select a target Managed Identity or service principal first.'
    }

    if ($null -eq $script:GraphServicePrincipal) {
        Load-GraphApplicationRoles
    }

    Write-UiLog -Message "Reading existing assignments for '$($script:TargetServicePrincipal.DisplayName)'..."

    $assignments = @(
        Get-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $script:TargetServicePrincipal.Id `
            -All `
            -ErrorAction Stop |
        Where-Object {
            $_.ResourceId -eq $script:GraphServicePrincipal.Id
        }
    )

    $roleById = @{}
    foreach ($role in $script:AllApplicationRoles) {
        $roleById[[string]$role.Id] = $role
    }

    $script:gridAssignments.Rows.Clear()

    foreach ($assignment in $assignments | Sort-Object AppRoleId) {
        $role = $roleById[[string]$assignment.AppRoleId]

        $permissionValue = if ($null -ne $role) {
            [string]$role.Value
        }
        else {
            "Unknown role: $($assignment.AppRoleId)"
        }

        $displayName = if ($null -ne $role) {
            [string]$role.DisplayName
        }
        else {
            ''
        }

        [void]$script:gridAssignments.Rows.Add(
            $permissionValue,
            $displayName,
            [string]$assignment.AppRoleId,
            [string]$assignment.Id
        )
    }

    $script:lblAssignedCount.Text = "{0} Microsoft Graph permission(s) assigned" -f $assignments.Count
    Write-UiLog -Message "Current assignments loaded: $($assignments.Count)." -Level 'SUCCESS'
}

function Set-TargetServicePrincipalResults {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Results
    )

    $selectedId = if ($null -ne $script:TargetServicePrincipal) {
        [string]$script:TargetServicePrincipal.Id
    }
    else {
        ''
    }

    $script:SuppressIdentitySelectionEvent = $true
    $script:cmbIdentityResults.BeginUpdate()
    try {
        $script:cmbIdentityResults.Items.Clear()

        foreach ($sp in $Results) {
            $label = "{0} | {1} | {2}" -f (
                $sp.DisplayName,
                $sp.ServicePrincipalType,
                $sp.Id
            )

            # ComboBox cannot directly host controls; use a PSObject with ToString().
            $item = [pscustomobject]@{
                Text = $label
                Tag  = $sp
            }
            $item | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Text } -Force
            [void]$script:cmbIdentityResults.Items.Add($item)
        }
    }
    finally {
        $script:cmbIdentityResults.EndUpdate()
    }

    $script:lblSearchResults.Text = "{0} matching identities" -f $Results.Count

    if ($Results.Count -eq 0) {
        $script:SuppressIdentitySelectionEvent = $false
        $script:TargetServicePrincipal = $null
        Update-IdentitySummary -ServicePrincipal $null
        $script:gridAssignments.Rows.Clear()
        $script:lblAssignedCount.Text = '0 permission(s) assigned'
        return
    }

    $selectedIndex = 0
    if (-not [string]::IsNullOrWhiteSpace($selectedId)) {
        for ($index = 0; $index -lt $Results.Count; $index++) {
            if ([string]$Results[$index].Id -eq $selectedId) {
                $selectedIndex = $index
                break
            }
        }
    }

    $script:cmbIdentityResults.SelectedIndex = $selectedIndex
    $script:SuppressIdentitySelectionEvent = $false

    $selected = $Results[$selectedIndex]
    $selectionChanged = [string]$selected.Id -ne $selectedId
    $script:TargetServicePrincipal = $selected
    Update-IdentitySummary -ServicePrincipal $selected

    if ($selectionChanged -and $script:GraphConnected) {
        Write-UiLog -Message "Selected target: $($selected.DisplayName)"
        Refresh-CurrentAssignments
    }
}

function Load-TargetServicePrincipals {
    if (-not $script:GraphConnected) {
        throw 'Connect to Microsoft Graph first.'
    }

    Write-UiLog -Message 'Loading Managed Identities and service principals...'
    $script:AllTargetServicePrincipals = @(
        Get-MgServicePrincipal `
            -Property 'id,appId,displayName,servicePrincipalType,accountEnabled' `
            -All `
            -ErrorAction Stop |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_.DisplayName)
        } |
        Sort-Object DisplayName
    )

    Filter-TargetServicePrincipals

    Write-UiLog -Message (
        "Loaded {0} Managed Identities and service principals." -f
        $script:AllTargetServicePrincipals.Count
    ) -Level 'SUCCESS'
}

function Filter-TargetServicePrincipals {
    $term = $script:txtIdentityName.Text.Trim()
    $results = if ([string]::IsNullOrWhiteSpace($term)) {
        @($script:AllTargetServicePrincipals)
    }
    else {
        @(
            $script:AllTargetServicePrincipals |
            Where-Object {
                ([string]$_.DisplayName).IndexOf(
                    $term,
                    [System.StringComparison]::OrdinalIgnoreCase
                ) -ge 0 -or
                ([string]$_.AppId).IndexOf(
                    $term,
                    [System.StringComparison]::OrdinalIgnoreCase
                ) -ge 0 -or
                ([string]$_.Id).IndexOf(
                    $term,
                    [System.StringComparison]::OrdinalIgnoreCase
                ) -ge 0
            }
        )
    }

    Set-TargetServicePrincipalResults -Results $results
}

function Assign-SelectedPermissions {
    if ($null -eq $script:TargetServicePrincipal) {
        throw 'Select a target Managed Identity or service principal first.'
    }

    if ($null -eq $script:GraphServicePrincipal) {
        Load-GraphApplicationRoles
    }

    $selectedItems = @()
    foreach ($item in $script:checkedPermissions.CheckedItems) {
        $selectedItems += $item
    }

    if ($selectedItems.Count -eq 0) {
        throw 'Select at least one Microsoft Graph application permission.'
    }

    $confirmationMessage = @"
Target:
$($script:TargetServicePrincipal.DisplayName)

Permissions to process:
$($selectedItems.Value -join [Environment]::NewLine)

Proceed with the app role assignments?
"@

    $answer = [System.Windows.Forms.MessageBox]::Show(
        $confirmationMessage,
        'Confirm permission assignment',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-UiLog -Message 'Permission assignment cancelled by the user.' -Level 'WARNING'
        return
    }

    $existingAssignments = @(
        Get-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $script:TargetServicePrincipal.Id `
            -All `
            -ErrorAction Stop |
        Where-Object {
            $_.ResourceId -eq $script:GraphServicePrincipal.Id
        }
    )

    $existingRoleIds = @{}
    foreach ($assignment in $existingAssignments) {
        $existingRoleIds[[string]$assignment.AppRoleId] = $true
    }

    $assigned = 0
    $skipped  = 0
    $failed   = 0

    foreach ($item in $selectedItems) {
        $roleId = [string]$item.Id

        if ($existingRoleIds.ContainsKey($roleId)) {
            Write-UiLog -Message "Already assigned: $($item.Value)" -Level 'WARNING'
            $skipped++
            continue
        }

        try {
            $body = @{
                PrincipalId = [guid]$script:TargetServicePrincipal.Id
                ResourceId  = [guid]$script:GraphServicePrincipal.Id
                AppRoleId   = [guid]$item.Id
            }

            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $script:TargetServicePrincipal.Id `
                -BodyParameter $body `
                -ErrorAction Stop | Out-Null

            Write-UiLog -Message "Assigned successfully: $($item.Value)" -Level 'SUCCESS'
            $existingRoleIds[$roleId] = $true
            $assigned++
        }
        catch {
            Write-UiLog -Message (
                "Failed to assign {0}: {1}" -f
                $item.Value,
                $_.Exception.Message
            ) -Level 'ERROR'
            $failed++
        }
    }

    Refresh-CurrentAssignments

    [System.Windows.Forms.MessageBox]::Show(
        "Completed.`n`nAssigned: $assigned`nAlready present: $skipped`nFailed: $failed",
        'Graph App Role Manager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Remove-SelectedPermissions {
    if ($null -eq $script:TargetServicePrincipal) {
        throw 'Select a target Managed Identity or service principal first.'
    }

    $selectedRows = @($script:gridAssignments.SelectedRows)
    if ($selectedRows.Count -eq 0) {
        throw 'Select at least one existing Microsoft Graph permission to remove.'
    }

    $selectedAssignments = @(
        foreach ($row in $selectedRows) {
            [pscustomobject]@{
                Permission   = [string]$row.Cells['Permission'].Value
                AssignmentId = [string]$row.Cells['AssignmentId'].Value
            }
        }
    )

    $confirmationMessage = @"
Target:
$($script:TargetServicePrincipal.DisplayName)

Permissions to remove:
$($selectedAssignments.Permission -join [Environment]::NewLine)

Remove these app role assignments?
"@

    $answer = [System.Windows.Forms.MessageBox]::Show(
        $confirmationMessage,
        'Confirm permission removal',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-UiLog -Message 'Permission removal cancelled by the user.' -Level 'WARNING'
        return
    }

    $removed = 0
    $failed  = 0

    foreach ($assignment in $selectedAssignments) {
        try {
            Remove-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $script:TargetServicePrincipal.Id `
                -AppRoleAssignmentId $assignment.AssignmentId `
                -ErrorAction Stop

            Write-UiLog -Message "Removed successfully: $($assignment.Permission)" -Level 'SUCCESS'
            $removed++
        }
        catch {
            Write-UiLog -Message (
                "Failed to remove {0}: {1}" -f
                $assignment.Permission,
                $_.Exception.Message
            ) -Level 'ERROR'
            $failed++
        }
    }

    Refresh-CurrentAssignments

    [System.Windows.Forms.MessageBox]::Show(
        "Completed.`n`nRemoved: $removed`nFailed: $failed",
        'Graph App Role Manager',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Start-GraphAppRoleManager {
    [CmdletBinding()]
    param(
        [switch]$NoGui,

        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$LogLevel = 'INFO'
    )

    if (-not $IsWindows) {
        throw 'This GUI uses Windows Forms and must be run on Windows.'
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    Reset-GraphAppRoleManagerState
    $script:LogLevel = $LogLevel

    # ---------------------------------------------------------------------------
    # FORM
    # ---------------------------------------------------------------------------

    $script:colorCanvas = [System.Drawing.Color]::FromArgb(243, 246, 250)
    $script:colorHeader = [System.Drawing.Color]::FromArgb(28, 38, 53)
    $script:colorBlue = [System.Drawing.Color]::FromArgb(47, 128, 237)
    $script:colorGreen = [System.Drawing.Color]::FromArgb(22, 163, 74)
    $script:colorRed = [System.Drawing.Color]::FromArgb(190, 45, 45)
    $script:colorMuted = [System.Drawing.Color]::FromArgb(96, 108, 126)

    function Set-RoundedButtonRegion {
        param([Parameter(Mandatory)]$Button)

        if ($Button.Width -le 0 -or $Button.Height -le 0) {
            return
        }

        $radius = 7
        $diameter = $radius * 2
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        try {
            $path.AddArc(0, 0, $diameter, $diameter, 180, 90)
            $path.AddArc($Button.Width - $diameter - 1, 0, $diameter, $diameter, 270, 90)
            $path.AddArc(
                $Button.Width - $diameter - 1,
                $Button.Height - $diameter - 1,
                $diameter,
                $diameter,
                0,
                90
            )
            $path.AddArc(0, $Button.Height - $diameter - 1, $diameter, $diameter, 90, 90)
            $path.CloseFigure()

            $previousRegion = $Button.Region
            $Button.Region = New-Object System.Drawing.Region($path)
            if ($null -ne $previousRegion) {
                $previousRegion.Dispose()
            }
        }
        finally {
            $path.Dispose()
        }
    }

    function Set-PrimaryButtonStyle {
        param(
            [Parameter(Mandatory)]$Button,
            [System.Drawing.Color]$BackColor = $script:colorBlue
        )

        $Button.BackColor = $BackColor
        $Button.ForeColor = [System.Drawing.Color]::White
        $Button.FlatStyle = 'Flat'
        $Button.FlatAppearance.BorderSize = 0
        $Button.FlatAppearance.MouseOverBackColor = [System.Windows.Forms.ControlPaint]::Light($BackColor)
        $Button.FlatAppearance.MouseDownBackColor = [System.Windows.Forms.ControlPaint]::Dark($BackColor)
        $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
        $Button.Margin = New-Object System.Windows.Forms.Padding(6)
        $Button.Add_Resize({
            Set-RoundedButtonRegion -Button $this
        })
        Set-RoundedButtonRegion -Button $Button
    }

    function New-FieldLabel {
        param([Parameter(Mandatory)][string]$Text)

        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Text
        $label.AutoSize = $true
        $label.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
        $label.ForeColor = [System.Drawing.Color]::FromArgb(55, 65, 81)
        $label.Anchor = 'Left'
        return $label
    }

    $script:form = New-Object System.Windows.Forms.Form
    $script:form.Text = 'Graph App Role Manager — Native Windows'
    $script:form.StartPosition = 'CenterScreen'
    $script:form.Size = New-Object System.Drawing.Size(1280, 860)
    $script:form.MinimumSize = New-Object System.Drawing.Size(1080, 740)
    $script:form.BackColor = $script:colorCanvas
    $script:form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $script:form.AutoScaleMode = 'Dpi'

    $script:rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:rootLayout.Dock = 'Fill'
    $script:rootLayout.ColumnCount = 1
    $script:rootLayout.RowCount = 2
    $script:rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 96)))
    $script:rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $script:form.Controls.Add($script:rootLayout)

    $script:headerPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $script:headerPanel.Dock = 'Fill'
    $script:headerPanel.BackColor = $script:colorHeader
    $script:headerPanel.ColumnCount = 2
    $script:headerPanel.RowCount = 1
    $script:headerPanel.Padding = New-Object System.Windows.Forms.Padding(24, 14, 24, 14)
    $script:headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    $script:headerPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 250)))
    $script:rootLayout.Controls.Add($script:headerPanel, 0, 0)

    $script:titlePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $script:titlePanel.Dock = 'Fill'
    $script:titlePanel.ColumnCount = 1
    $script:titlePanel.RowCount = 2
    $script:titlePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 62)))
    $script:titlePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 38)))
    $script:headerPanel.Controls.Add($script:titlePanel, 0, 0)

    $script:lblTitle = New-Object System.Windows.Forms.Label
    $script:lblTitle.Text = 'Graph App Role Manager'
    $script:lblTitle.ForeColor = [System.Drawing.Color]::White
    $script:lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 20)
    $script:lblTitle.AutoSize = $true
    $script:lblTitle.Anchor = 'Left'
    $script:titlePanel.Controls.Add($script:lblTitle, 0, 0)

    $script:lblSubtitle = New-Object System.Windows.Forms.Label
    $script:lblSubtitle.Text = 'Native Windows alternative · Assign Microsoft Graph application permissions'
    $script:lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(195, 205, 219)
    $script:lblSubtitle.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $script:lblSubtitle.AutoSize = $true
    $script:lblSubtitle.Anchor = 'Left'
    $script:titlePanel.Controls.Add($script:lblSubtitle, 0, 1)

    $script:btnConnect = New-Object System.Windows.Forms.Button
    $script:btnConnect.Text = 'Connect to Microsoft Graph'
    $script:btnConnect.Size = New-Object System.Drawing.Size(230, 40)
    $script:btnConnect.Anchor = 'Right'
    $script:btnConnect.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    Set-PrimaryButtonStyle -Button $script:btnConnect
    $script:headerPanel.Controls.Add($script:btnConnect, 1, 0)

    $script:tabs = New-Object System.Windows.Forms.TabControl
    $script:tabs.Dock = 'Fill'
    $script:tabs.Padding = New-Object System.Drawing.Point(18, 7)
    $script:tabs.Margin = New-Object System.Windows.Forms.Padding(12)
    $script:rootLayout.Controls.Add($script:tabs, 0, 1)

    # ---------------------------------------------------------------------------
    # TAB 1 - ASSIGN
    # ---------------------------------------------------------------------------

    $script:tabAssign = New-Object System.Windows.Forms.TabPage
    $script:tabAssign.Text = 'Manage permissions'
    $script:tabAssign.BackColor = $script:colorCanvas
    $script:tabAssign.Padding = New-Object System.Windows.Forms.Padding(14)
    $script:tabs.TabPages.Add($script:tabAssign)

    $script:mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:mainLayout.Dock = 'Fill'
    $script:mainLayout.ColumnCount = 1
    $script:mainLayout.RowCount = 2
    $script:mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 235)))
    $script:mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $script:tabAssign.Controls.Add($script:mainLayout)

    $script:grpIdentity = New-Object System.Windows.Forms.GroupBox
    $script:grpIdentity.Text = '1. Choose the target identity'
    $script:grpIdentity.Dock = 'Fill'
    $script:grpIdentity.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 14)
    $script:grpIdentity.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
    $script:mainLayout.Controls.Add($script:grpIdentity, 0, 0)

    $script:identityLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:identityLayout.Dock = 'Fill'
    $script:identityLayout.ColumnCount = 2
    $script:identityLayout.RowCount = 4
    $script:identityLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    $script:identityLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 110)))
    $script:identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 27)))
    $script:identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 38)))
    $script:identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 57)))
    $script:identityLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $script:grpIdentity.Controls.Add($script:identityLayout)

    $script:lblIdentitySearch = New-FieldLabel -Text 'Filter the loaded identities by name, application ID, or object ID'
    $script:identityLayout.Controls.Add($script:lblIdentitySearch, 0, 0)

    $script:txtIdentityName = New-Object System.Windows.Forms.TextBox
    $script:txtIdentityName.PlaceholderText = 'All identities are shown after connection — type here to filter'
    $script:txtIdentityName.Dock = 'Fill'
    $script:txtIdentityName.Margin = New-Object System.Windows.Forms.Padding(0, 3, 8, 6)
    $script:identityLayout.Controls.Add($script:txtIdentityName, 0, 1)

    $script:btnSearchIdentity = New-Object System.Windows.Forms.Button
    $script:btnSearchIdentity.Text = 'Clear'
    $script:btnSearchIdentity.Dock = 'Fill'
    $script:btnSearchIdentity.Margin = New-Object System.Windows.Forms.Padding(4, 2, 0, 5)
    Set-PrimaryButtonStyle `
        -Button $script:btnSearchIdentity `
        -BackColor ([System.Drawing.Color]::FromArgb(226, 232, 240))
    $script:btnSearchIdentity.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
    $script:identityLayout.Controls.Add($script:btnSearchIdentity, 1, 1)

    $script:resultsLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:resultsLayout.Dock = 'Fill'
    $script:resultsLayout.ColumnCount = 1
    $script:resultsLayout.RowCount = 2
    $script:resultsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 23)))
    $script:resultsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $script:identityLayout.SetColumnSpan($script:resultsLayout, 2)
    $script:identityLayout.Controls.Add($script:resultsLayout, 0, 2)

    $script:lblSearchResults = New-FieldLabel -Text 'Matching identities'
    $script:resultsLayout.Controls.Add($script:lblSearchResults, 0, 0)

    $script:cmbIdentityResults = New-Object System.Windows.Forms.ComboBox
    $script:cmbIdentityResults.DropDownStyle = 'DropDownList'
    $script:cmbIdentityResults.Dock = 'Fill'
    $script:cmbIdentityResults.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
    $script:resultsLayout.Controls.Add($script:cmbIdentityResults, 0, 1)

    $script:summaryPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $script:summaryPanel.Dock = 'Fill'
    $script:summaryPanel.BackColor = [System.Drawing.Color]::White
    $script:summaryPanel.CellBorderStyle = 'Single'
    $script:summaryPanel.ColumnCount = 4
    $script:summaryPanel.RowCount = 2
    $script:summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 85)))
    $script:summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
    $script:summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 85)))
    $script:summaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
    $script:summaryPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))
    $script:summaryPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))
    $script:identityLayout.SetColumnSpan($script:summaryPanel, 2)
    $script:identityLayout.Controls.Add($script:summaryPanel, 0, 3)

    $script:lblIdentityName = New-FieldLabel -Text 'Name'
    $script:lblIdentityNameValue = New-FieldLabel -Text '-'
    $script:lblIdentityNameValue.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $script:lblObjectId = New-FieldLabel -Text 'Object ID'
    $script:lblObjectIdValue = New-FieldLabel -Text '-'
    $script:lblObjectIdValue.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $script:lblAppId = New-FieldLabel -Text 'App ID'
    $script:lblAppIdValue = New-FieldLabel -Text '-'
    $script:lblAppIdValue.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $script:summaryPanel.Controls.Add($script:lblIdentityName, 0, 0)
    $script:summaryPanel.Controls.Add($script:lblIdentityNameValue, 1, 0)
    $script:summaryPanel.Controls.Add($script:lblObjectId, 0, 1)
    $script:summaryPanel.Controls.Add($script:lblObjectIdValue, 1, 1)
    $script:summaryPanel.Controls.Add($script:lblAppId, 2, 1)
    $script:summaryPanel.Controls.Add($script:lblAppIdValue, 3, 1)
    $script:summaryPanel.SetColumnSpan($script:lblIdentityNameValue, 3)

    $script:splitPermissions = New-Object System.Windows.Forms.SplitContainer
    $script:splitPermissions.Dock = 'Fill'
    $script:splitPermissions.SplitterWidth = 8
    $script:splitPermissions.BackColor = $script:colorCanvas
    $script:mainLayout.Controls.Add($script:splitPermissions, 0, 1)

    $script:grpPermissions = New-Object System.Windows.Forms.GroupBox
    $script:grpPermissions.Text = '2. Select Microsoft Graph application permissions'
    $script:grpPermissions.Dock = 'Fill'
    $script:grpPermissions.Padding = New-Object System.Windows.Forms.Padding(14)
    $script:splitPermissions.Panel1.Controls.Add($script:grpPermissions)

    $script:permissionLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:permissionLayout.Dock = 'Fill'
    $script:permissionLayout.ColumnCount = 2
    $script:permissionLayout.RowCount = 3
    $script:permissionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    $script:permissionLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 170)))
    $script:permissionLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 42)))
    $script:permissionLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $script:permissionLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))
    $script:grpPermissions.Controls.Add($script:permissionLayout)

    $script:txtPermissionFilter = New-Object System.Windows.Forms.TextBox
    $script:txtPermissionFilter.PlaceholderText = 'Filter permissions, e.g. DeviceManagement or Sites'
    $script:txtPermissionFilter.Dock = 'Fill'
    $script:txtPermissionFilter.Margin = New-Object System.Windows.Forms.Padding(0, 5, 8, 7)
    $script:permissionLayout.Controls.Add($script:txtPermissionFilter, 0, 0)

    $script:btnLoadPermissions = New-Object System.Windows.Forms.Button
    $script:btnLoadPermissions.Text = 'Reload permissions'
    $script:btnLoadPermissions.Dock = 'Fill'
    $script:btnLoadPermissions.Margin = New-Object System.Windows.Forms.Padding(4, 3, 0, 6)
    Set-PrimaryButtonStyle -Button $script:btnLoadPermissions
    $script:permissionLayout.Controls.Add($script:btnLoadPermissions, 1, 0)

    $script:checkedPermissions = New-Object System.Windows.Forms.CheckedListBox
    $script:checkedPermissions.CheckOnClick = $true
    $script:checkedPermissions.Dock = 'Fill'
    $script:checkedPermissions.HorizontalScrollbar = $true
    $script:checkedPermissions.DisplayMember = 'Text'
    $script:checkedPermissions.BackColor = [System.Drawing.Color]::White
    $script:permissionLayout.SetColumnSpan($script:checkedPermissions, 2)
    $script:permissionLayout.Controls.Add($script:checkedPermissions, 0, 1)

    $script:lblPermissionCount = New-Object System.Windows.Forms.Label
    $script:lblPermissionCount.Text = '0 permission(s) displayed'
    $script:lblPermissionCount.AutoSize = $true
    $script:lblPermissionCount.ForeColor = $script:colorMuted
    $script:lblPermissionCount.Anchor = 'Left'
    $script:permissionLayout.Controls.Add($script:lblPermissionCount, 0, 2)

    $script:btnAssign = New-Object System.Windows.Forms.Button
    $script:btnAssign.Text = 'Assign selected'
    $script:btnAssign.Dock = 'Fill'
    $script:btnAssign.Margin = New-Object System.Windows.Forms.Padding(4, 6, 0, 0)
    Set-PrimaryButtonStyle -Button $script:btnAssign -BackColor $script:colorGreen
    $script:permissionLayout.Controls.Add($script:btnAssign, 1, 2)

    $script:grpCurrent = New-Object System.Windows.Forms.GroupBox
    $script:grpCurrent.Text = '3. Current Microsoft Graph assignments'
    $script:grpCurrent.Dock = 'Fill'
    $script:grpCurrent.Padding = New-Object System.Windows.Forms.Padding(14)
    $script:splitPermissions.Panel2.Controls.Add($script:grpCurrent)

    $script:assignmentLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:assignmentLayout.Dock = 'Fill'
    $script:assignmentLayout.ColumnCount = 2
    $script:assignmentLayout.RowCount = 3
    $script:assignmentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    $script:assignmentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 120)))
    $script:assignmentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 42)))
    $script:assignmentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    $script:assignmentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))
    $script:grpCurrent.Controls.Add($script:assignmentLayout)

    $script:lblAssignedCount = New-Object System.Windows.Forms.Label
    $script:lblAssignedCount.Text = '0 permission(s) assigned'
    $script:lblAssignedCount.ForeColor = $script:colorMuted
    $script:lblAssignedCount.AutoSize = $true
    $script:lblAssignedCount.Anchor = 'Left'
    $script:assignmentLayout.Controls.Add($script:lblAssignedCount, 0, 0)

    $script:btnRefreshAssignments = New-Object System.Windows.Forms.Button
    $script:btnRefreshAssignments.Text = 'Refresh'
    $script:btnRefreshAssignments.Dock = 'Fill'
    $script:btnRefreshAssignments.Margin = New-Object System.Windows.Forms.Padding(4, 3, 0, 6)
    Set-PrimaryButtonStyle -Button $script:btnRefreshAssignments
    $script:assignmentLayout.Controls.Add($script:btnRefreshAssignments, 1, 0)

    $script:gridAssignments = New-Object System.Windows.Forms.DataGridView
    $script:gridAssignments.Dock = 'Fill'
    $script:gridAssignments.AllowUserToAddRows = $false
    $script:gridAssignments.AllowUserToDeleteRows = $false
    $script:gridAssignments.AllowUserToResizeRows = $false
    $script:gridAssignments.ReadOnly = $true
    $script:gridAssignments.RowHeadersVisible = $false
    $script:gridAssignments.SelectionMode = 'FullRowSelect'
    $script:gridAssignments.MultiSelect = $true
    $script:gridAssignments.AutoSizeColumnsMode = 'Fill'
    $script:gridAssignments.BackgroundColor = [System.Drawing.Color]::White
    $script:gridAssignments.BorderStyle = 'Fixed3D'
    $script:gridAssignments.EnableHeadersVisualStyles = $false
    $script:gridAssignments.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(229, 235, 244)
    $script:gridAssignments.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(31, 41, 55)
    $script:gridAssignments.DefaultCellStyle.SelectionBackColor = $script:colorBlue
    $script:assignmentLayout.SetColumnSpan($script:gridAssignments, 2)
    $script:assignmentLayout.Controls.Add($script:gridAssignments, 0, 1)
    [void]$script:gridAssignments.Columns.Add('Permission', 'Permission')
    [void]$script:gridAssignments.Columns.Add('DisplayName', 'Display name')
    [void]$script:gridAssignments.Columns.Add('RoleId', 'App role ID')
    [void]$script:gridAssignments.Columns.Add('AssignmentId', 'Assignment ID')
    $script:gridAssignments.Columns['RoleId'].Visible = $false
    $script:gridAssignments.Columns['AssignmentId'].Visible = $false

    $script:btnRemove = New-Object System.Windows.Forms.Button
    $script:btnRemove.Text = 'Remove selected'
    $script:btnRemove.Dock = 'Fill'
    $script:btnRemove.Margin = New-Object System.Windows.Forms.Padding(4, 6, 0, 0)
    Set-PrimaryButtonStyle -Button $script:btnRemove -BackColor $script:colorRed
    $script:assignmentLayout.Controls.Add($script:btnRemove, 1, 2)

    $script:tabLog = New-Object System.Windows.Forms.TabPage
    $script:tabLog.Text = 'Activity log'
    $script:tabLog.Padding = New-Object System.Windows.Forms.Padding(12)
    $script:tabs.TabPages.Add($script:tabLog)

    $script:txtLog = New-Object System.Windows.Forms.TextBox
    $script:txtLog.Dock = 'Fill'
    $script:txtLog.Multiline = $true
    $script:txtLog.ReadOnly = $true
    $script:txtLog.ScrollBars = 'Both'
    $script:txtLog.WordWrap = $false
    $script:txtLog.BackColor = [System.Drawing.Color]::FromArgb(17, 24, 39)
    $script:txtLog.ForeColor = [System.Drawing.Color]::FromArgb(229, 231, 235)
    $script:txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
    $script:tabLog.Controls.Add($script:txtLog)

    # ---------------------------------------------------------------------------
    # EVENTS
    # ---------------------------------------------------------------------------

    $script:btnConnect.Add_Click({
        try {
            Set-BusyState -Busy $true
            Ensure-GraphModules

            Write-UiLog -Message 'Opening Microsoft Graph authentication...'

            Connect-MgGraph `
                -Scopes @(
                    'Application.Read.All',
                    'AppRoleAssignment.ReadWrite.All'
                ) `
                -NoWelcome `
                -ContextScope Process `
                -ErrorAction Stop

            $context = Get-MgContext
            $script:GraphConnected = $true

            $script:btnConnect.Text = "Connected: $($context.Account)"
            $script:btnConnect.BackColor = [System.Drawing.Color]::FromArgb(22, 163, 74)
            $script:btnConnect.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(34, 180, 88)
            $script:btnConnect.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(18, 130, 59)

            Write-UiLog -Message (
                "Connected to tenant {0} as {1}." -f
                $context.TenantId,
                $context.Account
            ) -Level 'SUCCESS'

            Load-GraphApplicationRoles
            Load-TargetServicePrincipals
        }
        catch {
            Show-UiError -Message $_.Exception.Message
        }
        finally {
            Set-BusyState -Busy $false
        }
    })

    $script:btnSearchIdentity.Add_Click({
        $script:txtIdentityName.Clear()
        $script:txtIdentityName.Focus()
    })

    $script:txtIdentityName.Add_TextChanged({
        try {
            Filter-TargetServicePrincipals
        }
        catch {
            Write-UiLog -Message $_.Exception.Message -Level 'ERROR'
        }
    })

    $script:cmbIdentityResults.Add_SelectedIndexChanged({
        if ($script:SuppressIdentitySelectionEvent) {
            return
        }

        try {
            $selected = Get-SelectedTargetServicePrincipal
            $script:TargetServicePrincipal = $selected
            Update-IdentitySummary -ServicePrincipal $selected

            if ($null -ne $selected) {
                Write-UiLog -Message "Selected target: $($selected.DisplayName)"
                Refresh-CurrentAssignments
            }
        }
        catch {
            Show-UiError -Message $_.Exception.Message
        }
    })

    $script:btnLoadPermissions.Add_Click({
        try {
            Set-BusyState -Busy $true
            Load-GraphApplicationRoles
        }
        catch {
            Show-UiError -Message $_.Exception.Message
        }
        finally {
            Set-BusyState -Busy $false
        }
    })

    $script:txtPermissionFilter.Add_TextChanged({
        try {
            Apply-PermissionFilter
        }
        catch {
            Write-UiLog -Message $_.Exception.Message -Level 'ERROR'
        }
    })

    $script:btnRefreshAssignments.Add_Click({
        try {
            Set-BusyState -Busy $true
            Refresh-CurrentAssignments
        }
        catch {
            Show-UiError -Message $_.Exception.Message
        }
        finally {
            Set-BusyState -Busy $false
        }
    })

    $script:btnAssign.Add_Click({
        try {
            Set-BusyState -Busy $true
            Assign-SelectedPermissions
        }
        catch {
            Show-UiError -Message $_.Exception.Message
        }
        finally {
            Set-BusyState -Busy $false
        }
    })

    $script:btnRemove.Add_Click({
        try {
            Set-BusyState -Busy $true
            Remove-SelectedPermissions
        }
        catch {
            Show-UiError -Message $_.Exception.Message
        }
        finally {
            Set-BusyState -Busy $false
        }
    })

    $script:form.Add_FormClosing({
        try {
            if ($script:GraphConnected) {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }
        }
        catch {}
    })

    $script:form.Add_Shown({
        if ($script:splitPermissions.ClientSize.Width -gt 0) {
            $script:splitPermissions.SplitterDistance = [Math]::Floor(
                ($script:splitPermissions.ClientSize.Width - $script:splitPermissions.SplitterWidth) / 2
            )
        }

        Write-UiLog -Message 'Graph App Role Manager started.'
        Write-UiLog -Message 'Connect to Microsoft Graph, choose a loaded identity, then select application permissions.'
    })

    $script:splitPermissions.Add_SizeChanged({
        if ($script:splitPermissions.ClientSize.Width -le 0) {
            return
        }

        $balancedDistance = [Math]::Floor(
            ($script:splitPermissions.ClientSize.Width - $script:splitPermissions.SplitterWidth) / 2
        )

        if ($balancedDistance -gt 0) {
            $script:splitPermissions.SplitterDistance = $balancedDistance
        }
    })


    if (-not $NoGui) {
        [void]$script:form.ShowDialog()
    }
}

Export-ModuleMember -Function Start-GraphAppRoleManager
