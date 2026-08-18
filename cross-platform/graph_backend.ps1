#Requires -Version 7.0

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'

$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8NoBom
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom

$GraphAppId = '00000003-0000-0000-c000-000000000000'
$GraphServicePrincipal = $null
$GraphRoles = @()

function Send-ProtocolResponse {
    param(
        [Parameter(Mandatory)] [hashtable] $Payload
    )

    $json = $Payload | ConvertTo-Json -Depth 20 -Compress
    [Console]::Out.WriteLine("__GARM__$json")
    [Console]::Out.Flush()
}

function Invoke-GraphGetAll {
    param([Parameter(Mandatory)] [string] $Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    while ($next) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        foreach ($item in @($response.value)) {
            $items.Add($item) | Out-Null
        }
        $next = $response.'@odata.nextLink'
    }

    return $items.ToArray()
}

function Ensure-GraphServicePrincipal {
    if ($null -ne $script:GraphServicePrincipal) {
        return
    }

    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$GraphAppId'&`$select=id,appId,displayName,appRoles"
    $result = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    $script:GraphServicePrincipal = @($result.value) | Select-Object -First 1

    if (-not $script:GraphServicePrincipal) {
        throw 'Microsoft Graph service principal was not found in this tenant.'
    }
}

function Handle-Command {
    param([Parameter(Mandatory)] [pscustomobject] $Request)

    switch ($Request.command) {
        'ping' {
            return @{ ok = $true; data = @{ backend = 'PowerShell'; version = $PSVersionTable.PSVersion.ToString() } }
        }

        'prerequisites' {
            $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication |
                Sort-Object Version -Descending |
                Select-Object -First 1

            return @{
                ok = $true
                data = @{
                    powershellVersion = $PSVersionTable.PSVersion.ToString()
                    graphModuleInstalled = [bool]$module
                    graphModuleVersion = if ($module) { $module.Version.ToString() } else { $null }
                }
            }
        }

        'connect' {
            $tenantId = [string]$Request.tenantId
            $params = @{
                Scopes = @('Application.Read.All', 'AppRoleAssignment.ReadWrite.All')
                ContextScope = 'Process'
                NoWelcome = $true
            }
            if (-not [string]::IsNullOrWhiteSpace($tenantId)) {
                $params.TenantId = $tenantId
            }

            if ($IsWindows) {
                # Microsoft Graph PowerShell uses WAM on Windows. This backend is
                # intentionally headless, so it cannot provide WAM with a parent
                # window handle. Device code flow is the supported alternative.
                $params.UseDeviceCode = $true
                Connect-MgGraph @params | ForEach-Object {
                    $message = [string]$_
                    if (-not [string]::IsNullOrWhiteSpace($message)) {
                        $uriMatch = [regex]::Match($message, 'https://[^\s]+')
                        $codeMatch = [regex]::Match(
                            $message,
                            '(?i)\bcode\s+([A-Z0-9]{8,12})\b'
                        )
                        if ($uriMatch.Success -and $codeMatch.Success) {
                            Send-ProtocolResponse @{
                                ok = $true
                                event = 'deviceCode'
                                requestId = [string]$Request.requestId
                                data = @{
                                    verificationUri = $uriMatch.Value.TrimEnd('.', ',', ';')
                                    userCode = $codeMatch.Groups[1].Value.ToUpperInvariant()
                                    message = $message
                                }
                            }
                        }
                        else {
                            Send-ProtocolResponse @{
                                ok = $true
                                event = 'authMessage'
                                requestId = [string]$Request.requestId
                                data = @{ message = $message }
                            }
                        }
                    }
                }
            }
            else {
                Connect-MgGraph @params | Out-Null
            }
            $context = Get-MgContext
            $script:GraphServicePrincipal = $null
            $script:GraphRoles = @()

            return @{
                ok = $true
                data = @{
                    account = $context.Account
                    tenantId = $context.TenantId
                    authType = $context.AuthType
                    scopes = @($context.Scopes)
                }
            }
        }

        'disconnect' {
            Disconnect-MgGraph | Out-Null
            $script:GraphServicePrincipal = $null
            $script:GraphRoles = @()
            return @{ ok = $true; data = @{ disconnected = $true } }
        }

        'listServicePrincipals' {
            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=id,appId,displayName,servicePrincipalType,accountEnabled&`$top=999"
            $values = @(
                Invoke-GraphGetAll -Uri $uri |
                    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.displayName) } |
                    Sort-Object displayName
            )

            return @{ ok = $true; data = $values }
        }

        'searchServicePrincipals' {
            $term = ([string]$Request.prefix).Trim()
            if ([string]::IsNullOrWhiteSpace($term)) {
                throw 'A display-name search term is required.'
            }

            # Microsoft Graph service-principal searches are advanced queries.
            # Use $search + ConsistencyLevel:eventual so the text may appear
            # anywhere in the display name, not only at its beginning.
            $searchText = $term.Replace('"', '\"')
            $searchExpression = [uri]::EscapeDataString('"displayName:' + $searchText + '"')
            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$search=$searchExpression&`$select=id,appId,displayName,servicePrincipalType,accountEnabled&`$count=true&`$top=100"
            $headers = @{ ConsistencyLevel = 'eventual' }
            $result = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers -OutputType PSObject

            $values = @(
                $result.value |
                    Where-Object {
                        -not [string]::IsNullOrWhiteSpace([string]$_.displayName) -and
                        ([string]$_.displayName).IndexOf($term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                    } |
                    Sort-Object displayName
            )

            return @{ ok = $true; data = $values }
        }

        'loadGraphRoles' {
            Ensure-GraphServicePrincipal
            $roles = @(
                $script:GraphServicePrincipal.appRoles |
                    Where-Object {
                        $_.isEnabled -eq $true -and
                        @($_.allowedMemberTypes) -contains 'Application' -and
                        -not [string]::IsNullOrWhiteSpace([string]$_.value)
                    } |
                    Sort-Object value |
                    ForEach-Object {
                        [pscustomobject]@{
                            id = [string]$_.id
                            value = [string]$_.value
                            displayName = [string]$_.displayName
                            description = [string]$_.description
                        }
                    }
            )
            $script:GraphRoles = $roles
            return @{ ok = $true; data = $roles }
        }

        'getAssignments' {
            Ensure-GraphServicePrincipal
            $targetId = [string]$Request.targetId
            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$targetId/appRoleAssignments?`$top=999"
            $all = @(Invoke-GraphGetAll -Uri $uri)
            $filtered = @($all | Where-Object { [string]$_.resourceId -eq [string]$script:GraphServicePrincipal.id })
            return @{ ok = $true; data = $filtered }
        }

        'assignRole' {
            Ensure-GraphServicePrincipal
            $targetId = [string]$Request.targetId
            $appRoleId = [string]$Request.appRoleId
            $body = @{
                principalId = $targetId
                resourceId = [string]$script:GraphServicePrincipal.id
                appRoleId = $appRoleId
            } | ConvertTo-Json -Compress
            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$targetId/appRoleAssignments"
            $result = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json' -OutputType PSObject
            return @{ ok = $true; data = $result }
        }

        'removeAssignment' {
            $targetId = [string]$Request.targetId
            $assignmentId = [string]$Request.assignmentId
            $uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$targetId/appRoleAssignments/$assignmentId"
            Invoke-MgGraphRequest -Method DELETE -Uri $uri | Out-Null
            return @{ ok = $true; data = @{ removed = $true } }
        }

        default {
            throw "Unknown command: $($Request.command)"
        }
    }
}

Send-ProtocolResponse @{ ok = $true; event = 'ready'; data = @{ pid = $PID } }

while (($line = [Console]::In.ReadLine()) -ne $null) {
    if ([string]::IsNullOrWhiteSpace($line)) {
        continue
    }

    try {
        $request = $line | ConvertFrom-Json
        $response = Handle-Command -Request $request
        $response.requestId = $request.requestId
        Send-ProtocolResponse $response
    }
    catch {
        $failedRequestId = $null
        if ($null -ne $request -and $null -ne $request.requestId) {
            $failedRequestId = [string]$request.requestId
        }

        Send-ProtocolResponse @{
            ok = $false
            requestId = $failedRequestId
            error = $_.Exception.Message
            details = $_.ScriptStackTrace
        }
    }
}
