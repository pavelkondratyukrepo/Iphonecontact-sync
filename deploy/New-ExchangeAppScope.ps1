#Requires -Version 7.4
#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Grants the Automation account's managed identity scoped Contacts.ReadWrite access through Exchange RBAC for Applications.

.DESCRIPTION
    Run once as an Exchange Administrator (Organization Management) after Deploy-Automation.ps1 printed the identity's
    application id and object id. Idempotent: each of the three Exchange objects is created only if it does not exist.

      1. New-ServicePrincipal          registers the managed identity in Exchange Online by app id and object id.
      2. New-ManagementScope           CorpiPhoneMailboxes = mailboxes where CustomAttribute1 -eq 'CorpiPhone'.
                                       Exchange's CustomAttribute1 is Entra's extensionAttribute1, so the stamp written by
                                       Runbook 1 both fills the dynamic group and opens the mailbox.
      3. New-ManagementRoleAssignment  Application Contacts.ReadWrite, restricted to that scope.

    Then it runs Test-ServicePrincipalAuthorization for one stamped and one unstamped user and prints both results,
    and checks through Microsoft Graph that the identity holds NO Contacts.* application permission in Entra.
    That last check matters: Microsoft documents that Entra grants and Exchange RBAC grants are a union, so an unscoped
    Contacts.ReadWrite consent in Entra would silently defeat the mailbox scope created here.

    Timing: live role assignment and scope changes take 30 minutes to 2 hours to propagate to the Exchange permission
    cache. Test-ServicePrincipalAuthorization bypasses that cache, so its result is immediate, but a runbook started
    right after this script may still see 403 on fresh mailboxes. Right after first deployment both test users show
    InScope = False because nobody is stamped yet; re-run with -TestOnly after Runbook 1 has run live.

.PARAMETER ApplicationId
    Application (client) id of the managed identity, printed by Deploy-Automation.ps1.
.PARAMETER ObjectId
    Object id of the managed identity's service principal, printed by Deploy-Automation.ps1.
.PARAMETER StampedUserPrincipalName
    UPN of a user you expect to be in scope (a corporate iPhone primary user).
.PARAMETER UnstampedUserPrincipalName
    UPN of a user you expect NOT to be in scope.
.PARAMETER ServicePrincipalDisplayName
    Display name for the Exchange service principal object. Default 'CorpiPhone Directory Sync'.
.PARAMETER ScopeName
    Name of the management scope. Default CorpiPhoneMailboxes.
.PARAMETER ScopeAttributeName
    Exchange custom attribute that mirrors the Entra stamp. Default CustomAttribute1 (= extensionAttribute1).
.PARAMETER ScopeAttributeValue
    Stamp value. Default CorpiPhone.
.PARAMETER Organization
    Optional tenant domain for Connect-ExchangeOnline (for example contoso.onmicrosoft.com).
.PARAMETER TestOnly
    Skip creation and only run the authorization tests and the Entra permission check.
.PARAMETER SkipGraphCheck
    Skip the Microsoft Graph check for Contacts.* app role assignments.
.PARAMETER UseDeviceCode
    Sign in to Exchange Online and Graph with a device code (Windows Sandbox, no cached account). Without it the
    script tries the interactive flow first and falls back to device code when it fails.

.EXAMPLE
    ./New-ExchangeAppScope.ps1 -ApplicationId <app id> -ObjectId <object id> `
        -StampedUserPrincipalName alice@contoso.com -UnstampedUserPrincipalName bob@contoso.com
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [ValidatePattern('^[0-9a-fA-F-]{36}$')] [string]$ApplicationId,
    [Parameter(Mandatory)] [ValidatePattern('^[0-9a-fA-F-]{36}$')] [string]$ObjectId,
    [Parameter(Mandatory)] [string]$StampedUserPrincipalName,
    [Parameter(Mandatory)] [string]$UnstampedUserPrincipalName,
    [string]$ServicePrincipalDisplayName = 'CorpiPhone Directory Sync',
    [string]$ScopeName = 'CorpiPhoneMailboxes',
    [string]$ScopeAttributeName = 'CustomAttribute1',
    [string]$ScopeAttributeValue = 'CorpiPhone',
    [string]$Organization,
    [switch]$TestOnly,
    [switch]$SkipGraphCheck,
    [switch]$UseDeviceCode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RoleName = 'Application Contacts.ReadWrite'
$GraphAppId = '00000003-0000-0000-c000-000000000000'

function Write-Step {
    param([string]$Message)
    Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message) -ForegroundColor Cyan
}

function Test-IsThisIdentity {
    # Exchange reports a service principal assignee by its Exchange identity, which can be the object id, the app id,
    # the display name or the service principal's Name, depending on module version. Match any of them.
    param([Parameter(Mandatory)] [object]$Assignment, [string[]]$Names)
    foreach ($prop in 'RoleAssignee', 'RoleAssigneeName', 'App') {
        $v = [string](Get-PropertyOrNull $Assignment $prop)
        if ($v -and ($Names -contains $v)) { return $true }
    }
    return $false
}

function Get-AssignmentScope {
    param([Parameter(Mandatory)] [object]$Assignment)
    foreach ($prop in 'CustomResourceScope', 'CustomRecipientWriteScope') {
        $v = [string](Get-PropertyOrNull $Assignment $prop)
        if ($v) { return $v }
    }
    return ''
}

function Get-PropertyOrNull {
    param([Parameter(Mandatory)] [object]$InputObject, [Parameter(Mandatory)] [string]$Name)
    $p = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

# ----------------------------------------------------------------------------------------------------------------
# Connect
# ----------------------------------------------------------------------------------------------------------------
Write-Step "Connecting to Exchange Online$(if ($UseDeviceCode) { ' (device code)' })"
$connectArgs = @{ ShowBanner = $false; ErrorAction = 'Stop' }
if ($Organization) { $connectArgs.Organization = $Organization }
if ($UseDeviceCode) { Connect-ExchangeOnline @connectArgs -Device }
else {
    try { Connect-ExchangeOnline @connectArgs }
    catch {
        Write-Warning "Interactive Exchange sign-in failed ($($_.Exception.Message)). Falling back to device code, which is expected inside Windows Sandbox."
        Connect-ExchangeOnline @connectArgs -Device
    }
}

try {
if (-not $TestOnly) {
    # ------------------------------------------------------------------------------------------------------------
    # 1. Service principal
    # ------------------------------------------------------------------------------------------------------------
    Write-Step "Ensuring Exchange service principal for object id $ObjectId"
    $sp = Get-ServicePrincipal -Identity $ObjectId -ErrorAction SilentlyContinue
    if (-not $sp) {
        $sp = New-ServicePrincipal -AppId $ApplicationId -ObjectId $ObjectId -DisplayName $ServicePrincipalDisplayName
        Write-Host "  created '$ServicePrincipalDisplayName'"
    }
    else {
        Write-Host "  exists '$($sp.DisplayName)'"
        if ((Get-PropertyOrNull $sp 'AppId') -and $sp.AppId -ne $ApplicationId) {
            throw "Exchange service principal $ObjectId exists with AppId $($sp.AppId), which does not match -ApplicationId $ApplicationId."
        }
    }

    # ------------------------------------------------------------------------------------------------------------
    # 2. Management scope
    # ------------------------------------------------------------------------------------------------------------
    $filter = "$ScopeAttributeName -eq '$ScopeAttributeValue'"
    Write-Step "Ensuring management scope '$ScopeName' with filter [$filter]"
    $scope = Get-ManagementScope -Identity $ScopeName -ErrorAction SilentlyContinue
    if (-not $scope) {
        $scope = New-ManagementScope -Name $ScopeName -RecipientRestrictionFilter $filter
        Write-Host '  created'
    }
    else {
        $existingFilter = [string](Get-PropertyOrNull $scope 'RecipientFilter')
        if ($existingFilter -and ($existingFilter -replace '\s', '') -ne ($filter -replace '\s', '')) {
            Write-Warning "Scope '$ScopeName' exists with filter [$existingFilter], expected [$filter]. Fix with Set-ManagementScope if that is unintended."
        }
        Write-Host '  exists'
    }

    # ------------------------------------------------------------------------------------------------------------
    # 3. Role assignment
    # ------------------------------------------------------------------------------------------------------------
    Write-Step "Ensuring role assignment '$RoleName' for the identity, restricted to scope '$ScopeName'"
    # A fixed name makes the assignment easy to find and makes an accidental duplicate fail loudly instead of stacking.
    $assignmentName = "CorpiPhone-ContactsReadWrite-$ScopeName"
    $identityNames = @($ObjectId, $ApplicationId, $ServicePrincipalDisplayName, [string](Get-PropertyOrNull $sp 'Identity'), [string](Get-PropertyOrNull $sp 'Name')) |
        Where-Object { $_ } | Select-Object -Unique
    $assignments = @(Get-ManagementRoleAssignment -Role $RoleName -ErrorAction SilentlyContinue | Where-Object { Test-IsThisIdentity -Assignment $_ -Names $identityNames })
    $existing = $assignments | Where-Object { $_.Name -eq $assignmentName -or (Get-AssignmentScope $_) -eq $ScopeName } | Select-Object -First 1
    if (-not $existing) {
        $existing = New-ManagementRoleAssignment -Name $assignmentName -App $ObjectId -Role $RoleName -CustomResourceScope $ScopeName
        Write-Host "  created '$($existing.Name)'"
    }
    else {
        Write-Host "  exists '$($existing.Name)' (scope [$(Get-AssignmentScope $existing)])"
    }

    # Any other assignment of the same role to this identity (for example an unscoped one made by hand) widens access.
    foreach ($o in @($assignments | Where-Object { $_.Name -ne $existing.Name })) {
        $otherScope = Get-AssignmentScope $o
        if (-not $otherScope) { $otherScope = 'NONE - tenant-wide' }
        Write-Warning "Additional '$RoleName' assignment '$($o.Name)' exists for this identity with scope [$otherScope]. Remove it unless intended: Remove-ManagementRoleAssignment -Identity '$($o.Name)'"
    }

    Write-Host ''
    Write-Host '  NOTE: live assignment and scope changes take 30 minutes to 2 hours to reach the Exchange permission cache.' -ForegroundColor Yellow
    Write-Host '        Test-ServicePrincipalAuthorization below bypasses that cache and is accurate immediately.' -ForegroundColor Yellow
}

# ----------------------------------------------------------------------------------------------------------------
# 4. Authorization tests
# ----------------------------------------------------------------------------------------------------------------
Write-Step 'Test-ServicePrincipalAuthorization'
$results = foreach ($upn in @($StampedUserPrincipalName, $UnstampedUserPrincipalName)) {
    $expected = if ($upn -eq $StampedUserPrincipalName) { 'expected InScope=True once stamped' } else { 'expected InScope=False' }
    try {
        $rows = @(Test-ServicePrincipalAuthorization -Identity $ObjectId -Resource $upn)
        $row = $rows | Where-Object { [string](Get-PropertyOrNull $_ 'RoleName') -eq $RoleName } | Select-Object -First 1
        if (-not $row -and $rows.Count -gt 0) { $row = $rows[0] }
        [pscustomobject]@{
            User        = $upn
            RoleName    = if ($row) { Get-PropertyOrNull $row 'RoleName' } else { '(no roles returned)' }
            Permissions = if ($row) { Get-PropertyOrNull $row 'GrantedPermissions' } else { $null }
            Scope       = if ($row) { Get-PropertyOrNull $row 'AllowedResourceScope' } else { $null }
            ScopeType   = if ($row) { Get-PropertyOrNull $row 'ScopeType' } else { $null }
            InScope     = if ($row) { Get-PropertyOrNull $row 'InScope' } else { $false }
            Expectation = $expected
        }
    }
    catch {
        [pscustomobject]@{ User = $upn; RoleName = $null; Permissions = $null; Scope = $null; ScopeType = $null; InScope = $null; Expectation = "$expected (test failed: $($_.Exception.Message))" }
    }
}
$results | Format-Table -AutoSize | Out-String | Write-Host
Write-Host '  Right after first deployment both show InScope=False because nobody carries the stamp yet.' -ForegroundColor Yellow
Write-Host '  After Runbook 1 has run live, re-run this script with -TestOnly: the stamped user must show True, the other False.' -ForegroundColor Yellow
}
finally {
    # Disconnect before Graph: loading the Exchange and Graph authentication libraries into one session can clash.
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

# ----------------------------------------------------------------------------------------------------------------
# 5. Verify the identity has no Contacts.* permission consented in Entra
# ----------------------------------------------------------------------------------------------------------------
if ($SkipGraphCheck) {
    Write-Step 'Skipping Entra permission check (-SkipGraphCheck)'
}
else {
    Write-Step 'Checking Entra for Contacts.* application permissions on the identity (there must be none)'
    # Runs in a separate pwsh process: ExchangeOnlineManagement and Microsoft.Graph ship different MSAL versions and
    # loading both into one session can fail with an assembly conflict.
    $checkScript = Join-Path ([System.IO.Path]::GetTempPath()) "CorpiPhone-EntraCheck-$([guid]::NewGuid().ToString('N')).ps1"
    Set-Content -Path $checkScript -Encoding UTF8 -Value @'
param([string]$ObjectId, [string]$GraphAppId, [string]$DeviceCode)
$UseDeviceCode = $DeviceCode -eq 'yes'
$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications
if ($UseDeviceCode) { Connect-MgGraph -Scopes 'Application.Read.All' -UseDeviceCode -NoWelcome }
else {
    try { Connect-MgGraph -Scopes 'Application.Read.All' -NoWelcome }
    catch {
        Write-Warning "Interactive Graph sign-in failed ($($_.Exception.Message)). Falling back to device code."
        Connect-MgGraph -Scopes 'Application.Read.All' -UseDeviceCode -NoWelcome
    }
}
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'" -Property Id, AppRoles | Select-Object -First 1
$assigned = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ObjectId -All | Where-Object { $_.ResourceId -eq $graphSp.Id })
$roleById = @{}
foreach ($r in $graphSp.AppRoles) { $roleById[[string]$r.Id] = $r.Value }
$granted = @($assigned | ForEach-Object { $roleById[[string]$_.AppRoleId] } | Where-Object { $_ })
$contactsGrants = @($granted | Where-Object { $_ -like 'Contacts.*' })
$null = Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "  Graph application permissions on the identity: $($granted -join ', ')"
if ($contactsGrants.Count -gt 0) {
    Write-Host ''
    Write-Host '  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' -ForegroundColor Red
    Write-Host "  !!! The identity holds $($contactsGrants -join ', ') in ENTRA. That grant is tenant-wide and UNSCOPED." -ForegroundColor Red
    Write-Host '  !!! Entra grants and Exchange RBAC grants are a union: the Exchange scope created above does NOTHING' -ForegroundColor Red
    Write-Host '  !!! until that Entra grant is removed. Remove it in Entra > Enterprise applications > the identity >' -ForegroundColor Red
    Write-Host '  !!! Permissions, or with Remove-MgServicePrincipalAppRoleAssignment, then re-run with -TestOnly.' -ForegroundColor Red
    Write-Host '  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!' -ForegroundColor Red
    exit 2
}
Write-Host '  OK: no Contacts.* permission in Entra. Contacts access comes only from the scoped Exchange role.' -ForegroundColor Green
exit 0
'@
    try {
        & ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -NoProfile -File $checkScript -ObjectId $ObjectId -GraphAppId $GraphAppId -DeviceCode $(if ($UseDeviceCode) { 'yes' } else { 'no' })
        $checkExit = $LASTEXITCODE
    }
    finally { Remove-Item -Path $checkScript -Force -ErrorAction SilentlyContinue }
    if ($checkExit -eq 2) { Write-Warning 'Unscoped Contacts permission present in Entra. See the red banner above.' }
    elseif ($checkExit -ne 0) { Write-Warning "The Entra permission check did not complete (exit code $checkExit). Make sure Microsoft.Graph.Authentication and Microsoft.Graph.Applications are installed, or run Deploy-Automation.ps1 again, which performs the same check." }
}
