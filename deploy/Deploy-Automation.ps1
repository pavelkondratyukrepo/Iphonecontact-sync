#Requires -Version 7.4
#Requires -Modules Az.Accounts, Az.Resources, Az.Automation, Az.Monitor, Microsoft.Graph.Authentication, Microsoft.Graph.Groups, Microsoft.Graph.Applications
<#
.SYNOPSIS
    One-time deployment of the CorpiPhone directory sync: groups, Automation account, managed identity,
    Graph permissions, PowerShell 7.6 runtime environment, runbooks, schedules and alerting.

.DESCRIPTION
    Run once from an admin workstation. Idempotent: every step checks for the resource before creating it,
    so the script can be re-run after a partial failure or to pick up a new runbook version.

    What it creates
      1. Dynamic security group (default SG-CorpiPhone-Users) with rule (user.extensionAttribute1 -eq "CorpiPhone")
         and the static exclusion group (default SG-Directory-Exclude).
      2. Resource group and Automation account (Basic plan) with a system-assigned managed identity. Basic includes the
         same 500 free job minutes a month as Free (expected use is 60 to 80), but keeps running past them at a tiny
         per-minute rate instead of refusing to run jobs for the rest of the month.
      3. Four Microsoft Graph application permissions on the identity:
             DeviceManagementManagedDevices.Read.All, User.ReadUpdate.All, User-Phone.ReadWrite.All, GroupMember.Read.All,
             RoleManagement.Read.Directory. These are the least privileged roles Microsoft documents for updating a user
             (User.ReadUpdate.All, which cannot create or delete users) and its phone numbers (User-Phone.ReadWrite.All).
         Deliberately NOT Contacts.ReadWrite. Microsoft documents that Entra grants and Exchange RBAC for Applications
         grants are a union, so an unscoped Entra grant would defeat the mailbox scope created by New-ExchangeAppScope.ps1.
      4. A PowerShell runtime environment (7.6 by default; 7.4 reaches end of support on 10 November 2026) with the
         Microsoft.Graph submodules the runbooks need, pinned to -GraphModuleVersion. Re-running with a different
         -PowerShellVersion creates the new environment and moves both runbooks into it; the old one is left in place.
      5. Both runbooks imported and published into that runtime environment.
      6. Schedules (daily 05:00 and weekly Monday 08:00 in -ScheduleTimeZone) linked to the runbooks with
         WhatIf = $false. New schedules are created disabled unless -EnableSchedules is passed; a re-run without the
         switch leaves existing schedules exactly as they are. -DeferSchedules skips them.
      7. An action group and an Azure Monitor metric alert that emails -AlertEmail when a job ends Failed, Suspended or Stopped.

    What it does NOT do
      - It does not touch Exchange Online. Run deploy/New-ExchangeAppScope.ps1 next with the two ids this script prints.
      - It does not create Intune policies. Run deploy/New-OutlookContactSyncPolicies.ps1 for those.

.PARAMETER SubscriptionId
    Azure subscription that hosts the Automation account.
.PARAMETER ResourceGroupName
    Resource group to create or reuse.
.PARAMETER Location
    Azure region for the resource group and Automation account, for example eastus2.
.PARAMETER AutomationAccountName
    Name of the Automation account to create or reuse.
.PARAMETER AlertEmail
    Address that receives the Failed/Suspended/Stopped job alert.
.PARAMETER ControlGroupName
    Display name of the dynamic security group that Runbook 1 fills through the stamp. Default SG-CorpiPhone-Users.
.PARAMETER ExclusionGroupName
    Display name of the static group whose members receive contacts but are not listed. Default SG-Directory-Exclude.
.PARAMETER ScopeAttributeName
    Extension attribute used as the stamp. Default extensionAttribute1. If you change it, also change the constant
    at the top of runbooks/Sync-IntuneToEntra.ps1 and the -ScopeAttributeName of New-ExchangeAppScope.ps1.
.PARAMETER ScopeAttributeValue
    Value written to the stamp attribute. Default CorpiPhone.
.PARAMETER GraphModuleVersion
    Microsoft.Graph SDK version pinned into the runtime environment. 2.40.0 is known to load on the 7.4 runtime;
    2.41.0 and 2.41.1 are known to fail there. Before changing it, check the PowerShell Gallery and the Azure Automation
    release notes for the newest version confirmed to load, and test in a second runtime environment.
.PARAMETER PowerShellVersion
    Runtime version, 7.6 (default) or 7.4. If the first WhatIf job on 7.6 fails to load a Microsoft.Graph module, re-run
    with -PowerShellVersion 7.4 to move the runbooks back, and plan the move to 7.6 with a newer pinned Graph version.
.PARAMETER RuntimeEnvironmentName
    Name of the runtime environment. Default PowerShell76-GraphSync (PowerShell74-GraphSync for -PowerShellVersion 7.4).
.PARAMETER ScheduleTimeZone
    IANA time zone for the schedules. Default America/New_York.
.PARAMETER RunbookDirectory
    Folder that holds Sync-IntuneToEntra.ps1 and Sync-DirectoryToMailboxes.ps1. Default ../runbooks relative to this script.
.PARAMETER DeferSchedules
    Skip creating and linking schedules entirely.
.PARAMETER EnableSchedules
    Enable the schedules (new or existing). Without this, new schedules are created (and linked) but disabled, so
    nothing runs until step 6 of the deployment guide, and existing schedules are left exactly as they are.
.PARAMETER UseDeviceCode
    Sign in to Azure and Graph with a device code instead of the interactive browser/broker flow. Use this inside
    Windows Sandbox or any session without a cached Microsoft account. Without the switch the script tries the
    interactive flow first and falls back to device code when it fails.

.EXAMPLE
    ./Deploy-Automation.ps1 -SubscriptionId 00000000-0000-0000-0000-000000000000 -ResourceGroupName rg-corpiphone-sync `
        -Location eastus2 -AutomationAccountName aa-corpiphone-sync -AlertEmail it-alerts@contoso.com

.NOTES
    Requires: Owner or Contributor on the subscription, Global Administrator or Privileged Role Administrator
    (for the Graph app role assignments) and Groups Administrator (or Global Administrator) for the groups.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$SubscriptionId,
    [Parameter(Mandatory)] [string]$ResourceGroupName,
    [Parameter(Mandatory)] [string]$Location,
    [Parameter(Mandatory)] [string]$AutomationAccountName,
    [Parameter(Mandatory)] [string]$AlertEmail,
    [string]$ControlGroupName = 'SG-CorpiPhone-Users',
    [string]$ExclusionGroupName = 'SG-Directory-Exclude',
    [ValidatePattern('^extensionAttribute([1-9]|1[0-5])$')] [string]$ScopeAttributeName = 'extensionAttribute1',
    [ValidatePattern('^[A-Za-z0-9_-]{1,64}$')] [string]$ScopeAttributeValue = 'CorpiPhone',
    [string]$GraphModuleVersion = '2.40.0',
    [ValidateSet('7.6', '7.4')] [string]$PowerShellVersion = '7.6',
    [string]$RuntimeEnvironmentName = '',
    [string]$ScheduleTimeZone = 'America/New_York',
    [string]$RunbookDirectory = (Join-Path $PSScriptRoot '..' 'runbooks'),
    [switch]$DeferSchedules,
    [switch]$EnableSchedules,
    [switch]$UseDeviceCode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------------------------------------------
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$RequiredGraphPermissions = @(
    'DeviceManagementManagedDevices.Read.All',
    'User.ReadUpdate.All',
    'User-Phone.ReadWrite.All',
    'GroupMember.Read.All',
    'RoleManagement.Read.Directory'
)
# Modules imported into the runtime environment. Authentication must be first; the others depend on it.
$GraphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.PersonalContacts'
)
$Runbooks = @(
    @{ Name = 'Sync-IntuneToEntra'; Description = 'Stamps corporate iPhone primary users and writes their numbers to mobilePhone.' },
    @{ Name = 'Sync-DirectoryToMailboxes'; Description = 'Writes one Company Directory contact per coworker into each stamped mailbox.' }
)
$AutomationApiVersion = '2023-05-15-preview'
$AlertActionGroupName = "$AutomationAccountName-failed-jobs"
$AlertRuleName = "$AutomationAccountName-job-failed-or-suspended"

# ----------------------------------------------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------------------------------------------
function Write-Step {
    param([string]$Message)
    Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message) -ForegroundColor Cyan
}

function Connect-AzInteractive {
    <# Connect-AzAccount with device-code fallback. Windows Sandbox has no cached account, so the WAM/browser flow fails there. #>
    param([Parameter(Mandatory)] [string]$Subscription, [switch]$DeviceCode)
    if ($DeviceCode) { return (Connect-AzAccount -Subscription $Subscription -UseDeviceAuthentication -ErrorAction Stop) }
    try { return (Connect-AzAccount -Subscription $Subscription -ErrorAction Stop) }
    catch {
        Write-Warning "Interactive Azure sign-in failed ($($_.Exception.Message)). Falling back to device code, which is expected inside Windows Sandbox."
        return (Connect-AzAccount -Subscription $Subscription -UseDeviceAuthentication -ErrorAction Stop)
    }
}

function Connect-GraphInteractive {
    <# Connect-MgGraph with device-code fallback, same reason as above. #>
    param([Parameter(Mandatory)] [string[]]$Scopes, [switch]$DeviceCode)
    if ($DeviceCode) { Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome -ErrorAction Stop; return }
    try { Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop }
    catch {
        Write-Warning "Interactive Graph sign-in failed ($($_.Exception.Message)). Falling back to device code, which is expected inside Windows Sandbox."
        Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome -ErrorAction Stop
    }
}

function Get-ArmToken {
    # Az.Accounts 2.x returns a plain string token; 4.x/5.x return a SecureString. Handle both.
    $t = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/'
    if ($t.Token -is [securestring]) {
        return [System.Net.NetworkCredential]::new('', $t.Token).Password
    }
    return [string]$t.Token
}

function Invoke-Arm {
    <#
        Thin wrapper over Invoke-AzRestMethod that throws on non-2xx and returns the parsed JSON body (or $null).
    #>
    param(
        [Parameter(Mandatory)] [ValidateSet('GET', 'PUT', 'POST', 'PATCH', 'DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Path,
        [object]$Body
    )
    $restArgs = @{ Method = $Method; Path = $Path }
    if ($null -ne $Body) { $restArgs.Payload = ($Body | ConvertTo-Json -Depth 20) }
    $resp = Invoke-AzRestMethod @restArgs
    if ($resp.StatusCode -ge 400) {
        throw "ARM $Method $Path failed with $($resp.StatusCode): $($resp.Content)"
    }
    if ([string]::IsNullOrWhiteSpace($resp.Content)) { return $null }
    return ($resp.Content | ConvertFrom-Json -Depth 20)
}

function Get-ArmOrNull {
    param([Parameter(Mandatory)] [string]$Path)
    $resp = Invoke-AzRestMethod -Method GET -Path $Path
    if ($resp.StatusCode -eq 404) { return $null }
    if ($resp.StatusCode -ge 400) { throw "ARM GET $Path failed with $($resp.StatusCode): $($resp.Content)" }
    return ($resp.Content | ConvertFrom-Json -Depth 20)
}

function Get-SingleGroupByName {
    <# Same rule as the runbooks: escape quotes, refuse duplicates, so a deploy never binds an arbitrary same-named group. #>
    param([Parameter(Mandatory)] [string]$DisplayName, [string[]]$Property = @('Id', 'DisplayName'))
    $escaped = $DisplayName.Replace("'", "''")
    $found = @(Get-MgGroup -Filter "displayName eq '$escaped'" -Property $Property -Top 2)
    if ($found.Count -gt 1) { throw "More than one group is named '$DisplayName'. Group names used by this project must be unique." }
    if ($found.Count -eq 1) { return $found[0] }
    return $null
}

function Get-NextOccurrence {
    <#
        Returns the next wall-clock occurrence of $Hour:00 in $TimeZoneId (optionally on $DayOfWeek) as a UTC DateTimeOffset
        at least 15 minutes in the future, which is what New-AzAutomationSchedule requires.
    #>
    param(
        [Parameter(Mandatory)] [string]$TimeZoneId,
        [Parameter(Mandatory)] [int]$Hour,
        [System.DayOfWeek]$DayOfWeek
    )
    $tz = [TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneId)
    $nowLocal = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
    $candidate = [DateTime]::new($nowLocal.Year, $nowLocal.Month, $nowLocal.Day, $Hour, 0, 0, [DateTimeKind]::Unspecified)
    while (($candidate -le $nowLocal.AddMinutes(15)) -or ($PSBoundParameters.ContainsKey('DayOfWeek') -and $candidate.DayOfWeek -ne $DayOfWeek)) {
        $candidate = $candidate.AddDays(1)
    }
    # Wall-clock time with the zone's own offset: unambiguous whether the service reads StartTime as an instant or adds
    # the -TimeZone offset to it. The created schedule's NextRun is printed so the operator can confirm.
    return [DateTimeOffset]::new($candidate, $tz.GetUtcOffset($candidate))
}

function Wait-ForCondition {
    param(
        [Parameter(Mandatory)] [scriptblock]$Condition,
        [Parameter(Mandatory)] [string]$Description,
        [int]$TimeoutSeconds = 600,
        [int]$IntervalSeconds = 10
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $result = & $Condition
        if ($result) { return $result }
        Start-Sleep -Seconds $IntervalSeconds
    }
    throw "Timed out after $TimeoutSeconds seconds waiting for: $Description"
}

# ----------------------------------------------------------------------------------------------------------------
# 0. Preflight
# ----------------------------------------------------------------------------------------------------------------
foreach ($rb in $Runbooks) {
    $path = Join-Path $RunbookDirectory "$($rb.Name).ps1"
    if (-not (Test-Path $path)) { throw "Runbook file not found: $path" }
    $rb.Path = (Resolve-Path $path).Path
}

Write-Step "Connecting to Azure subscription $SubscriptionId$(if ($UseDeviceCode) { ' (device code)' })"
$null = Connect-AzInteractive -Subscription $SubscriptionId -DeviceCode:$UseDeviceCode
$null = Set-AzContext -Subscription $SubscriptionId

Write-Step 'Connecting to Microsoft Graph (Group.ReadWrite.All, AppRoleAssignment.ReadWrite.All, Application.Read.All)'
Connect-GraphInteractive -Scopes 'Group.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All' -DeviceCode:$UseDeviceCode

# ----------------------------------------------------------------------------------------------------------------
# 1. Groups
# ----------------------------------------------------------------------------------------------------------------
Write-Step "Ensuring dynamic group '$ControlGroupName'"
$membershipRule = "(user.$ScopeAttributeName -eq `"$ScopeAttributeValue`")"
$controlGroup = Get-SingleGroupByName -DisplayName $ControlGroupName -Property Id, DisplayName, GroupTypes, MembershipRule
if (-not $controlGroup) {
    $controlGroup = New-MgGroup -DisplayName $ControlGroupName `
        -Description "Corporate iPhone holders. Filled automatically by the Sync-IntuneToEntra runbook through $ScopeAttributeName. Do not edit by hand." `
        -MailEnabled:$false -MailNickname ($ControlGroupName -replace '[^A-Za-z0-9]', '') -SecurityEnabled `
        -GroupTypes @('DynamicMembership') -MembershipRule $membershipRule -MembershipRuleProcessingState 'On'
    Write-Host "  created $($controlGroup.Id)"
}
else {
    if ($controlGroup.GroupTypes -notcontains 'DynamicMembership') {
        throw "Group '$ControlGroupName' exists but is not a dynamic group. Rename or delete it, or pass a different -ControlGroupName."
    }
    if ($controlGroup.MembershipRule -ne $membershipRule) {
        Write-Warning "Group '$ControlGroupName' has rule [$($controlGroup.MembershipRule)] but the expected rule is [$membershipRule]. Fix it in Entra if that is unintended."
    }
    Write-Host "  exists $($controlGroup.Id)"
}

Write-Step "Ensuring static exclusion group '$ExclusionGroupName'"
$exclusionGroup = Get-SingleGroupByName -DisplayName $ExclusionGroupName
if (-not $exclusionGroup) {
    $exclusionGroup = New-MgGroup -DisplayName $ExclusionGroupName `
        -Description 'Members receive Company Directory contacts but are not listed in the directory, and their phone number is not written to mobilePhone.' `
        -MailEnabled:$false -MailNickname ($ExclusionGroupName -replace '[^A-Za-z0-9]', '') -SecurityEnabled
    Write-Host "  created $($exclusionGroup.Id)"
}
else {
    Write-Host "  exists $($exclusionGroup.Id)"
}

# ----------------------------------------------------------------------------------------------------------------
# 2. Resource group, Automation account, managed identity
# ----------------------------------------------------------------------------------------------------------------
Write-Step "Ensuring resource group '$ResourceGroupName' in $Location"
if (-not (Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue)) {
    $null = New-AzResourceGroup -Name $ResourceGroupName -Location $Location
    Write-Host '  created'
}
else { Write-Host '  exists' }

Write-Step "Ensuring Automation account '$AutomationAccountName' (Basic plan, system-assigned identity)"
$aa = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction SilentlyContinue
if (-not $aa) {
    $aa = New-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -Location $Location -Plan Basic -AssignSystemIdentity
    Write-Host '  created'
}
elseif (-not $aa.Identity -or -not $aa.Identity.PrincipalId) {
    $aa = Set-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -AssignSystemIdentity
    Write-Host '  exists, enabled system-assigned identity'
}
else { Write-Host '  exists' }
$planProperty = $aa.PSObject.Properties['Plan']
if ($planProperty -and [string]$planProperty.Value -eq 'Free') {
    Write-Warning "Automation account '$AutomationAccountName' is on the Free plan: once 500 job minutes are used in a month, jobs stop running until the next month. Switch it to Basic (portal: the account > Pricing) to avoid a silent outage; normal use stays inside the free minutes."
}

$identityObjectId = $aa.Identity.PrincipalId
if (-not $identityObjectId) {
    # Set-AzAutomationAccount sometimes returns before the identity is populated; re-read.
    $aa = Wait-ForCondition -Description 'managed identity to appear on the Automation account' -TimeoutSeconds 120 -Condition {
        $a = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName
        if ($a.Identity -and $a.Identity.PrincipalId) { $a } else { $null }
    }
    $identityObjectId = $aa.Identity.PrincipalId
}

Write-Step "Waiting for the identity's service principal $identityObjectId to appear in Entra"
$identitySp = Wait-ForCondition -Description 'managed identity service principal in Microsoft Graph' -TimeoutSeconds 300 -Condition {
    Get-MgServicePrincipal -ServicePrincipalId $identityObjectId -ErrorAction SilentlyContinue
}
$identityAppId = $identitySp.AppId
Write-Host "  application id $identityAppId, object id $identityObjectId"

# ----------------------------------------------------------------------------------------------------------------
# 3. Graph application permissions (app role assignments on the Microsoft Graph service principal)
# ----------------------------------------------------------------------------------------------------------------
Write-Step 'Granting Graph application permissions to the identity'
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'" -Property Id, AppRoles | Select-Object -First 1
if (-not $graphSp) { throw 'Microsoft Graph service principal not found in the tenant.' }
$existingAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $identityObjectId -All)

foreach ($permission in $RequiredGraphPermissions) {
    $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $permission -and $_.AllowedMemberTypes -contains 'Application' } | Select-Object -First 1
    if (-not $role) { throw "Graph app role '$permission' not found on the Microsoft Graph service principal." }
    if ($existingAssignments | Where-Object { $_.AppRoleId -eq $role.Id -and $_.ResourceId -eq $graphSp.Id }) {
        Write-Host "  exists  $permission"
        continue
    }
    $null = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $identityObjectId -PrincipalId $identityObjectId -ResourceId $graphSp.Id -AppRoleId $role.Id
    Write-Host "  granted $permission"
}

# An earlier version granted User.ReadWrite.All, which can also create and delete users. It is no longer needed.
$broadRole = $graphSp.AppRoles | Where-Object { $_.Value -eq 'User.ReadWrite.All' } | Select-Object -First 1
$broad = @($existingAssignments | Where-Object { $broadRole -and $_.ResourceId -eq $graphSp.Id -and $_.AppRoleId -eq $broadRole.Id })
if ($broad.Count -gt 0) {
    Write-Warning "The identity still holds User.ReadWrite.All from an earlier deployment. After one clean live run of Sync-IntuneToEntra, remove it: Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $identityObjectId -AppRoleAssignmentId $($broad[0].Id)"
}

# Guard: the identity must not hold any Contacts.* Graph permission, or the Exchange RBAC scope is defeated.
$contactsRoles = $graphSp.AppRoles | Where-Object { $_.Value -like 'Contacts.*' }
$badAssignments = $existingAssignments | Where-Object { $_.ResourceId -eq $graphSp.Id -and $_.AppRoleId -in $contactsRoles.Id }
if ($badAssignments) {
    foreach ($bad in $badAssignments) {
        $value = ($contactsRoles | Where-Object { $_.Id -eq $bad.AppRoleId }).Value
        Write-Warning "The identity holds Graph permission '$value' in Entra. This is an UNSCOPED grant and makes the Exchange mailbox scope meaningless."
        Write-Warning "  Remove it with: Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $identityObjectId -AppRoleAssignmentId $($bad.Id)"
    }
    # Runbook 2 refuses to run with such a grant anyway; stop here so the problem is fixed before anything is deployed on top.
    throw 'Unscoped Contacts.* permission found on the identity. Remove it (commands above) and re-run.'
}
else { Write-Host '  verified: no Contacts.* permission on the identity (contacts access comes only from Exchange RBAC)' }

# ----------------------------------------------------------------------------------------------------------------
# 4. PowerShell runtime environment with pinned Microsoft.Graph modules
# ----------------------------------------------------------------------------------------------------------------
$aaPath = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName"
if ([string]::IsNullOrWhiteSpace($RuntimeEnvironmentName)) { $RuntimeEnvironmentName = "PowerShell$($PowerShellVersion -replace '\.', '')-GraphSync" }
$rtPath = "$aaPath/runtimeEnvironments/$RuntimeEnvironmentName"

Write-Step "Ensuring PowerShell $PowerShellVersion runtime environment '$RuntimeEnvironmentName'"
$rt = Get-ArmOrNull -Path "$rtPath`?api-version=$AutomationApiVersion"
if (-not $rt) {
    $null = Invoke-Arm -Method PUT -Path "$rtPath`?api-version=$AutomationApiVersion" -Body @{
        location   = $Location
        properties = @{
            runtime         = @{ language = 'PowerShell'; version = $PowerShellVersion }
            defaultPackages = @{}
            description     = "PowerShell $PowerShellVersion with Microsoft.Graph $GraphModuleVersion for the CorpiPhone directory sync runbooks."
        }
    }
    Write-Host '  created'
}
else {
    $rtVersion = $rt.properties.runtime.version
    if ($rtVersion -ne $PowerShellVersion) { throw "Runtime environment '$RuntimeEnvironmentName' exists but is PowerShell $rtVersion, not $PowerShellVersion. Pass a different -RuntimeEnvironmentName." }
    Write-Host '  exists'
}

Write-Step "Adding Microsoft.Graph $GraphModuleVersion modules to the runtime environment (each waits until imported)"
foreach ($module in $GraphModules) {
    $pkgPath = "$rtPath/packages/$module`?api-version=$AutomationApiVersion"
    $pkg = Get-ArmOrNull -Path $pkgPath
    $currentVersion = if ($pkg) { $pkg.properties.version } else { $null }
    $state = if ($pkg) { $pkg.properties.provisioningState } else { $null }
    if ($pkg -and $currentVersion -eq $GraphModuleVersion -and $state -eq 'Succeeded') {
        Write-Host "  exists  $module $currentVersion"
        continue
    }
    $null = Invoke-Arm -Method PUT -Path $pkgPath -Body @{
        properties = @{
            contentLink = @{ uri = "https://www.powershellgallery.com/api/v2/package/$module/$GraphModuleVersion" }
        }
    }
    $final = Wait-ForCondition -Description "$module $GraphModuleVersion import" -TimeoutSeconds 900 -IntervalSeconds 15 -Condition {
        $p = Get-ArmOrNull -Path $pkgPath
        if ($p -and $p.properties.provisioningState -in @('Succeeded', 'Failed', 'Canceled')) { $p } else { $null }
    }
    if ($final.properties.provisioningState -ne 'Succeeded') {
        throw "Module $module $GraphModuleVersion failed to import into the runtime environment: $($final.properties | ConvertTo-Json -Depth 5 -Compress)"
    }
    Write-Host "  imported $module $GraphModuleVersion"
}

# ----------------------------------------------------------------------------------------------------------------
# 5. Runbooks: create (in the runtime environment), upload draft content, publish
# ----------------------------------------------------------------------------------------------------------------
Write-Step 'Importing and publishing runbooks'
$armToken = Get-ArmToken
foreach ($rb in $Runbooks) {
    $rbPath = "$aaPath/runbooks/$($rb.Name)"
    $existing = Get-ArmOrNull -Path "$rbPath`?api-version=$AutomationApiVersion"
    if (-not $existing) {
        $null = Invoke-Arm -Method PUT -Path "$rbPath`?api-version=$AutomationApiVersion" -Body @{
            location   = $Location
            properties = @{
                runbookType        = 'PowerShell'
                runtimeEnvironment = $RuntimeEnvironmentName
                logProgress        = $false
                logVerbose         = $false
                description        = $rb.Description
                draft              = @{}
            }
        }
        Write-Host "  created   $($rb.Name)"
    }
    elseif ($existing.properties.runtimeEnvironment -ne $RuntimeEnvironmentName) {
        # PATCH does not change runtimeEnvironment; a PUT with the full body does. Content is re-uploaded and
        # published just below, so replacing the runbook definition loses nothing.
        $null = Invoke-Arm -Method PUT -Path "$rbPath`?api-version=$AutomationApiVersion" -Body @{
            location   = $Location
            properties = @{
                runbookType        = 'PowerShell'
                runtimeEnvironment = $RuntimeEnvironmentName
                logProgress        = $false
                logVerbose         = $false
                description        = $rb.Description
                draft              = @{}
            }
        }
        $moved = Get-ArmOrNull -Path "$rbPath`?api-version=$AutomationApiVersion"
        if (-not $moved -or $moved.properties.runtimeEnvironment -ne $RuntimeEnvironmentName) {
            throw "Runbook $($rb.Name) is still not in runtime environment $RuntimeEnvironmentName. Move it in the portal (Runbook > Edit > Runtime environment) and re-run."
        }
        Write-Host "  moved     $($rb.Name) to runtime environment $RuntimeEnvironmentName"
    }

    $content = Get-Content -Path $rb.Path -Raw
    $contentUri = "https://management.azure.com$rbPath/draft/content?api-version=$AutomationApiVersion"
    $null = Invoke-RestMethod -Method Put -Uri $contentUri -ContentType 'text/powershell' -Body $content -Headers @{ Authorization = "Bearer $armToken" }

    # The content upload is asynchronous; publish only once the draft reflects the uploaded script.
    $null = Wait-ForCondition -Description "$($rb.Name) draft content upload" -TimeoutSeconds 180 -IntervalSeconds 5 -Condition {
        try {
            $draft = Invoke-RestMethod -Method Get -Uri $contentUri -Headers @{ Authorization = "Bearer $armToken" }
            # Compare ignoring line endings and a BOM: the service may store CRLF as LF.
            $norm = { param($t) ([string]$t).TrimStart([char]0xFEFF).Replace("`r`n", "`n").Trim() }
            if (($draft -is [string]) -and ((& $norm $draft) -eq (& $norm $content))) { $true } else { $null }
        }
        catch { $null }
    }
    $null = Invoke-Arm -Method POST -Path "$rbPath/publish?api-version=$AutomationApiVersion"
    $null = Wait-ForCondition -Description "$($rb.Name) publish" -TimeoutSeconds 180 -IntervalSeconds 5 -Condition {
        $r = Get-ArmOrNull -Path "$rbPath`?api-version=$AutomationApiVersion"
        if ($r -and $r.properties.state -eq 'Published') { $r } else { $null }
    }
    Write-Host "  published $($rb.Name)"
}

# ----------------------------------------------------------------------------------------------------------------
# 6. Schedules
# ----------------------------------------------------------------------------------------------------------------
if ($DeferSchedules) {
    Write-Step 'Skipping schedules (-DeferSchedules)'
}
else {
    Write-Step "Ensuring schedules in $ScheduleTimeZone (daily 05:00 for Sync-IntuneToEntra, Monday 08:00 for Sync-DirectoryToMailboxes)"
    $scheduleDefinitions = @(
        @{
            Name       = 'Sync-IntuneToEntra-Daily-0500'
            Runbook    = 'Sync-IntuneToEntra'
            StartTime  = (Get-NextOccurrence -TimeZoneId $ScheduleTimeZone -Hour 5)
            Weekly     = $false
            Parameters = @{
                WhatIf              = $false
                ExclusionGroupName  = $ExclusionGroupName
                DefaultCountryCode  = '+1'
                ScopeAttributeValue = $ScopeAttributeValue
            }
        },
        @{
            Name       = 'Sync-DirectoryToMailboxes-Weekly-Mon-0800'
            Runbook    = 'Sync-DirectoryToMailboxes'
            StartTime  = (Get-NextOccurrence -TimeZoneId $ScheduleTimeZone -Hour 8 -DayOfWeek Monday)
            Weekly     = $true
            Parameters = @{
                WhatIf             = $false
                GroupName          = $ControlGroupName
                ExclusionGroupName = $ExclusionGroupName
                UserPrincipalName  = ''
                FullRebuild        = $false
            }
        }
    )

    foreach ($def in $scheduleDefinitions) {
        $schedule = Get-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $def.Name -ErrorAction SilentlyContinue
        $isNew = $false
        if ($schedule -and $EnableSchedules -and -not $schedule.IsEnabled -and $schedule.StartTime -lt [DateTimeOffset]::UtcNow) {
            # A disabled schedule whose start time has passed is rejected when re-enabled; recreate it with the next start time.
            $null = Unregister-AzAutomationScheduledRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
                -RunbookName $def.Runbook -ScheduleName $def.Name -Force -ErrorAction SilentlyContinue
            Remove-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $def.Name -Force
            Write-Host "  removed stale schedule $($def.Name) so it is recreated with the next start time"
            $schedule = $null
        }
        if (-not $schedule) {
            $isNew = $true
            if ($def.Weekly) {
                $schedule = New-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $def.Name `
                    -StartTime $def.StartTime -WeekInterval 1 -DaysOfWeek Monday -TimeZone $ScheduleTimeZone
            }
            else {
                $schedule = New-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $def.Name `
                    -StartTime $def.StartTime -DayInterval 1 -TimeZone $ScheduleTimeZone
            }
            Write-Host "  created schedule $($def.Name) (first run $($def.StartTime.ToString('yyyy-MM-dd HH:mm zzz')) = $($schedule.NextRun) as Azure reports it; must be $($def.StartTime.Hour):00 $ScheduleTimeZone)"
        }
        else { Write-Host "  exists  schedule $($def.Name)" }

        $link = Get-AzAutomationScheduledRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
            -RunbookName $def.Runbook -ScheduleName $def.Name -ErrorAction SilentlyContinue
        if (-not $link) {
            $null = Register-AzAutomationScheduledRunbook -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName `
                -RunbookName $def.Runbook -ScheduleName $def.Name -Parameters $def.Parameters
            Write-Host "  linked  $($def.Name) -> $($def.Runbook) with WhatIf=false"
        }
        else { Write-Host "  linked  $($def.Name) -> $($def.Runbook) (already)" }

        # Only change enablement when asked: a plain re-run must never switch off schedules that are already live.
        if ($EnableSchedules) {
            $null = Set-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $def.Name -IsEnabled $true
            Write-Host "  ENABLED $($def.Name)"
        }
        elseif ($isNew) {
            $null = Set-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $def.Name -IsEnabled $false
            Write-Host "  disabled $($def.Name) (new schedules start disabled until the pilot passes)"
        }
        else {
            $current = Get-AzAutomationSchedule -ResourceGroupName $ResourceGroupName -AutomationAccountName $AutomationAccountName -Name $def.Name
            Write-Host "  left    $($def.Name) as it was (Enabled: $($current.IsEnabled))"
        }
    }
    if (-not $EnableSchedules) { Write-Host '  To switch the schedules on after the pilot, re-run with -EnableSchedules (deployment guide step 6).' }
}

# ----------------------------------------------------------------------------------------------------------------
# 7. Alerting: action group + metric alert on job status Failed or Suspended
# ----------------------------------------------------------------------------------------------------------------
Write-Step "Ensuring action group '$AlertActionGroupName' emailing $AlertEmail"
$actionGroup = Get-AzActionGroup -ResourceGroupName $ResourceGroupName -Name $AlertActionGroupName -ErrorAction SilentlyContinue
if (-not $actionGroup) {
    $shortName = ($AutomationAccountName -replace '[^A-Za-z0-9]', '')
    if ($shortName.Length -gt 12) { $shortName = $shortName.Substring(0, 12) }
    if (Get-Command -Name New-AzActionGroupEmailReceiverObject -ErrorAction SilentlyContinue) {
        # Az.Monitor 5.x and later. The short-name parameter was renamed to GroupShortName, and -Enabled is passed
        # explicitly so the group is never created switched off (which would silently drop every alert email).
        $receiver = New-AzActionGroupEmailReceiverObject -Name 'ops-email' -EmailAddress $AlertEmail
        $agParams = @{ ResourceGroupName = $ResourceGroupName; Name = $AlertActionGroupName; Location = 'Global'; EmailReceiver = $receiver }
        $agCommand = Get-Command -Name New-AzActionGroup
        if ($agCommand.Parameters.ContainsKey('GroupShortName')) { $agParams.GroupShortName = $shortName } else { $agParams.ShortName = $shortName }
        if ($agCommand.Parameters.ContainsKey('Enabled')) { $agParams.Enabled = $true }
        $actionGroup = New-AzActionGroup @agParams
    }
    else {
        # Az.Monitor 4.x and earlier
        $receiver = New-AzActionGroupReceiver -Name 'ops-email' -EmailReceiver -EmailAddress $AlertEmail
        $actionGroup = Set-AzActionGroup -ResourceGroupName $ResourceGroupName -Name $AlertActionGroupName -ShortName $shortName -Receiver $receiver
    }
    Write-Host '  created'
}
else { Write-Host '  exists' }
if ($actionGroup.PSObject.Properties['Enabled'] -and -not $actionGroup.Enabled) {
    if (Get-Command -Name Update-AzActionGroup -ErrorAction SilentlyContinue) {
        $actionGroup = Update-AzActionGroup -ResourceGroupName $ResourceGroupName -Name $AlertActionGroupName -Enabled
        Write-Host '  was disabled; enabled it'
    }
    else { Write-Warning "Action group $AlertActionGroupName is DISABLED, so failure emails will not be sent. Enable it in the portal (Monitor > Alerts > Action groups)." }
}

Write-Step "Ensuring metric alert '$AlertRuleName' (TotalJob where Status in Failed, Suspended, Stopped)"
# Stopped is included because a job that hits the Automation 3-hour fair-share limit ends Stopped, not Failed.
$dimension = New-AzMetricAlertRuleV2DimensionSelection -DimensionName 'Status' -ValuesToInclude 'Failed', 'Suspended', 'Stopped'
$criteria = New-AzMetricAlertRuleV2Criteria -MetricName 'TotalJob' -MetricNamespace 'Microsoft.Automation/automationAccounts' `
    -DimensionSelection $dimension -TimeAggregation Total -Operator GreaterThan -Threshold 0
$null = Add-AzMetricAlertRuleV2 -Name $AlertRuleName -ResourceGroupName $ResourceGroupName -Severity 2 `
    -WindowSize ([TimeSpan]::FromHours(1)) -Frequency ([TimeSpan]::FromMinutes(15)) `
    -TargetResourceId $aaPath -Condition $criteria -ActionGroupId $actionGroup.Id `
    -Description 'A CorpiPhone directory sync runbook job ended Failed, Suspended or Stopped.'
Write-Host '  upserted'

# ----------------------------------------------------------------------------------------------------------------
# Done
# ----------------------------------------------------------------------------------------------------------------
Write-Host ''
Write-Host '================================================================================' -ForegroundColor Green
Write-Host ' Deployment complete. Values for deploy/New-ExchangeAppScope.ps1:' -ForegroundColor Green
Write-Host "   -ApplicationId $identityAppId"
Write-Host "   -ObjectId      $identityObjectId"
Write-Host ''
Write-Host ' Next steps'
Write-Host "   1. Run New-ExchangeAppScope.ps1 with the two ids above (Exchange Administrator)."
Write-Host "   2. Start the Sync-IntuneToEntra runbook with WhatIf = true and review the stamp list."
Write-Host "   3. Confirm in Entra > Enterprise applications that the identity shows exactly: $($RequiredGraphPermissions -join ', ')"
Write-Host '================================================================================' -ForegroundColor Green

[pscustomobject]@{
    IdentityApplicationId = $identityAppId
    IdentityObjectId      = $identityObjectId
    ControlGroupId        = $controlGroup.Id
    ExclusionGroupId      = $exclusionGroup.Id
    AutomationAccountId   = $aaPath
    RuntimeEnvironment    = $RuntimeEnvironmentName
    GraphModuleVersion    = $GraphModuleVersion
}
