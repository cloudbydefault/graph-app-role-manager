$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Assert-True {
    param(
        [Parameter(Mandatory)][bool]$Condition,
        [Parameter(Mandatory)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Get-ModuleScriptVariable {
    param(
        [Parameter(Mandatory)]$Module,
        [Parameter(Mandatory)][string]$Name
    )

    return $Module.SessionState.PSVariable.GetValue($Name)
}

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$launcherPath = Join-Path $repositoryRoot 'windows/Graph-App-Role-Manager.ps1'
$modulePath = Join-Path $repositoryRoot 'windows/Graph-App-Role-Manager.psm1'
$launcherSource = Get-Content -LiteralPath $launcherPath -Raw
$moduleSource = Get-Content -LiteralPath $modulePath -Raw

Assert-True ($launcherSource -match 'Join-Path \$PSScriptRoot ''Graph-App-Role-Manager\.psm1''') `
    'The launcher must resolve the module from $PSScriptRoot.'
Assert-True ($launcherSource -match 'Start-GraphAppRoleManager') `
    'The launcher must invoke the public module entry point.'
Assert-True ($moduleSource -match 'Export-ModuleMember -Function Start-GraphAppRoleManager') `
    'The module must export Start-GraphAppRoleManager.'
Assert-True ($moduleSource -notmatch '\$Global:Globalloglevel') `
    'Log level must not use a global session variable.'

# Load and construct the complete form without showing it.
$module = Import-Module $modulePath -PassThru -Force
Start-GraphAppRoleManager -NoGui

try {
    $form = Get-ModuleScriptVariable -Module $module -Name 'form'
    $rootLayout = Get-ModuleScriptVariable -Module $module -Name 'rootLayout'
    $btnConnect = Get-ModuleScriptVariable -Module $module -Name 'btnConnect'
    $btnSearchIdentity = Get-ModuleScriptVariable -Module $module -Name 'btnSearchIdentity'
    $tabs = Get-ModuleScriptVariable -Module $module -Name 'tabs'
    $splitPermissions = Get-ModuleScriptVariable -Module $module -Name 'splitPermissions'

    Assert-True ($form.Text -like '*Native Windows*') 'The native Windows form was not constructed.'
    Assert-True ($rootLayout -is [System.Windows.Forms.TableLayoutPanel]) 'The responsive root layout is missing.'
    Assert-True ($btnConnect.Text -eq 'Connect to Microsoft Graph') 'The Graph connection action is missing.'
    Assert-True ($btnConnect.Height -eq 40) 'The compact Graph connection button is missing.'
    Assert-True ($btnConnect.FlatStyle -eq 'Flat') 'The modern button styling is missing.'
    Assert-True ($null -ne $btnConnect.Region) 'The rounded button styling is missing.'
    Assert-True ($tabs.TabPages.Count -eq 2) 'The permissions and activity-log tabs were not created.'
    Assert-True ($btnSearchIdentity.Text -eq 'Clear') 'The identity filter clear action is missing.'

    # Regression check for the StrictMode failure reported when exactly one
    # Microsoft Graph module is missing.
    Assert-True (
        $moduleSource -match '(?s)\$missingModules\s*=\s*@\(\s*foreach'
    ) 'The missing-module result must remain array-wrapped.'
    Assert-True (
        $moduleSource -match '(?s)Load-GraphApplicationRoles\s+Load-TargetServicePrincipals'
    ) 'Identities must load automatically after connection.'
    Assert-True (
        $moduleSource -match 'ClientSize\.Width\s*-\s*\$script:splitPermissions\.SplitterWidth\)\s*/\s*2'
    ) 'The permission and assignment panels must remain balanced.'
}
finally {
    if ($null -ne $form) {
        $form.Dispose()
    }
}

Write-Host 'Native Windows GUI smoke test passed.'
