#Requires -Version 7.4
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Groups
<#
.SYNOPSIS
    Creates or updates the two Intune policies that make Outlook for iOS sync Company Directory contacts into the
    native iOS Contacts app, and assigns both to a group.

.DESCRIPTION
    Run from an admin workstation as an Intune Administrator. Idempotent: policies are matched by display name.

    (a) iOS App Protection Policy (APP) targeting Microsoft Outlook, for apps on Intune-managed devices.
        The one setting this project needs is Data protection > Data transfer >
        "Sync policy managed app data with native apps or add-ins" = Allow, which is contactSyncBlocked = false in Graph.
        If a policy with -AppProtectionPolicyName already exists, ONLY that setting is changed, Outlook is added to its
        target apps if missing, and the group is added to its assignments. All other settings keep your existing values.
        If it does not exist, a new policy is created with contactSyncBlocked = false and a conservative data-transfer
        baseline (managed apps only, Save As to OneDrive/SharePoint only, backup blocked, compliance required). Review
        it in the portal before widening the assignment.

    (b) App Configuration Policy (ACP, managed apps) for Outlook on iOS and Android:
        Save Contacts = Yes, user cannot change, and the contact field sync list from the design:
        company, department, job title, mobile phone, work phone and work email on; display name, first and last name
        are always synced; home address, birthday, IM, nickname, notes, home/other/pager/fax phones, prefix and suffix off.
        Office location has no Outlook field-sync toggle (iOS contacts have no such field); it remains on the Exchange contact.
        Existing custom settings on a policy of the same name are preserved; only the keys listed here are set.

    Both policies are assigned to -GroupName (added to any existing assignments).

    Finally the script prints the two iOS device restriction settings to verify by hand; it does not create or edit
    device restriction profiles.

.PARAMETER GroupName
    Display name of the Entra group to assign both policies to (pilot group first, then SG-CorpiPhone-Users).
.PARAMETER AppProtectionPolicyName
    Display name of the iOS APP to create or update. Use the name of your existing Outlook APP for enrolled devices
    if you have one, so this script edits it instead of stacking a second policy.
.PARAMETER AppConfigurationPolicyName
    Display name of the managed-apps ACP to create or update.
.PARAMETER UseDeviceCode
    Sign in to Graph with a device code (Windows Sandbox, no cached account). Without it the script tries the
    interactive flow first and falls back to device code when it fails.

.EXAMPLE
    ./New-OutlookContactSyncPolicies.ps1 -GroupName 'SG-CorpiPhone-Pilot'
.EXAMPLE
    ./New-OutlookContactSyncPolicies.ps1 -GroupName 'SG-CorpiPhone-Users' -AppProtectionPolicyName 'iOS APP - Enrolled devices'

.NOTES
    Field sync keys are documented by Microsoft under "Deploying Outlook for iOS and Android app configuration settings",
    section "Contact field sync to native contacts". Verify the list against that page if Outlook ignores a setting.
    Uses the Graph beta endpoint for deviceAppManagement, which is where Intune app policy management is complete.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$GroupName,
    [string]$AppProtectionPolicyName = 'iOS APP - Outlook (Intune managed devices)',
    [string]$AppConfigurationPolicyName = 'ACP - Outlook contact sync (iOS and Android)',
    [switch]$UseDeviceCode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Beta = 'https://graph.microsoft.com/beta'
$OutlookIosBundleId = 'com.microsoft.Office.Outlook'
$OutlookAndroidPackageId = 'com.microsoft.office.outlook'

# Managed app configuration keys for Outlook. 'true'/'false' are strings in Intune custom settings.
$OutlookContactSyncSettings = [ordered]@{
    # Save Contacts: on, and the user cannot change it
    'com.microsoft.outlook.Contacts.LocalSyncEnabled'                   = 'true'
    'com.microsoft.outlook.Contacts.LocalSyncEnabled.UserChangeAllowed' = 'false'
    # Contact fields synced to the native Contacts app (display name, first and last name are always synced)
    'com.microsoft.outlook.ContactSync.CompanyAllowed'                  = 'true'
    'com.microsoft.outlook.ContactSync.DepartmentAllowed'               = 'true'
    'com.microsoft.outlook.ContactSync.JobTitleAllowed'                 = 'true'
    'com.microsoft.outlook.ContactSync.PhoneMobileAllowed'              = 'true'
    'com.microsoft.outlook.ContactSync.PhoneWorkAllowed'                = 'true'
    'com.microsoft.outlook.ContactSync.EmailAllowed'                    = 'true'
    # Personal and unused fields stay off
    'com.microsoft.outlook.ContactSync.AddressAllowed'                  = 'false'
    'com.microsoft.outlook.ContactSync.BirthdayAllowed'                 = 'false'
    'com.microsoft.outlook.ContactSync.InstantMessageAllowed'           = 'false'
    'com.microsoft.outlook.ContactSync.NicknameAllowed'                 = 'false'
    'com.microsoft.outlook.ContactSync.NotesAllowed'                    = 'false'
    'com.microsoft.outlook.ContactSync.PhoneHomeAllowed'                = 'false'
    'com.microsoft.outlook.ContactSync.PhoneHomeFaxAllowed'             = 'false'
    'com.microsoft.outlook.ContactSync.PhoneOtherAllowed'               = 'false'
    'com.microsoft.outlook.ContactSync.PhonePagerAllowed'               = 'false'
    'com.microsoft.outlook.ContactSync.PhoneWorkFaxAllowed'             = 'false'
    'com.microsoft.outlook.ContactSync.PrefixAllowed'                   = 'false'
    'com.microsoft.outlook.ContactSync.SuffixAllowed'                   = 'false'
}

function Write-Step {
    param([string]$Message)
    Write-Host ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $Message) -ForegroundColor Cyan
}

function Get-Field {
    param([object]$Object, [Parameter(Mandatory)] [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } else { return $null } }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Invoke-Graph {
    param(
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Uri,
        [object]$Body
    )
    if ($null -ne $Body) {
        $json = $Body | ConvertTo-Json -Depth 20
        return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $json -ContentType 'application/json'
    }
    return Invoke-MgGraphRequest -Method $Method -Uri $Uri
}

function Get-GraphCollection {
    param([Parameter(Mandatory)] [string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-Graph -Method GET -Uri $next
        $value = Get-Field $page 'value'
        if ($null -ne $value) { foreach ($v in $value) { $items.Add($v) } }
        $next = Get-Field $page '@odata.nextLink'
    }
    return $items.ToArray()
}

function Find-PolicyByName {
    param([Parameter(Mandatory)] [string]$Collection, [Parameter(Mandatory)] [string]$DisplayName)
    # Intune does not reliably support $filter on displayName for these collections; filter client-side.
    $all = Get-GraphCollection -Uri "$Beta/deviceAppManagement/$Collection"
    $hits = @($all | Where-Object { [string](Get-Field $_ 'displayName') -eq $DisplayName })
    if ($hits.Count -gt 1) { throw "More than one $Collection policy is named '$DisplayName'. Rename one and re-run." }
    if ($hits.Count -eq 0) { return $null }
    return Invoke-Graph -Method GET -Uri "$Beta/deviceAppManagement/$Collection/$(Get-Field $hits[0] 'id')?`$expand=apps,assignments"
}

function Get-TargetAppIdentifiers {
    param([object]$Policy)
    $ids = [System.Collections.Generic.List[string]]::new()
    $apps = Get-Field $Policy 'apps'
    if ($null -eq $apps) { return $ids }
    foreach ($a in $apps) {
        $ident = Get-Field $a 'mobileAppIdentifier'
        $bundle = Get-Field $ident 'bundleId'
        $package = Get-Field $ident 'packageId'
        if ($bundle) { $ids.Add("ios:$bundle") }
        if ($package) { $ids.Add("android:$package") }
    }
    return $ids
}

function Set-PolicyTargetApps {
    <# Replaces the policy's target app list with the existing apps plus any missing ones from $Required (list of 'ios:<bundle>' / 'android:<package>'). #>
    param([Parameter(Mandatory)] [string]$Collection, [Parameter(Mandatory)] [object]$Policy, [Parameter(Mandatory)] [string[]]$Required)
    # A policy that targets a built-in app group (allMicrosoftApps, allCoreMicrosoftApps, allApps) already covers Outlook.
    # Posting selectedPublicApps to it would narrow the policy to the listed apps, so leave it alone.
    $groupType = [string](Get-Field $Policy 'appGroupType')
    if ($groupType -and $groupType -ne 'selectedPublicApps') {
        Write-Host "  target apps left as is: the policy targets the '$groupType' group, which includes Outlook"
        return
    }
    $current = Get-TargetAppIdentifiers -Policy $Policy
    $missing = @($Required | Where-Object { $current -notcontains $_ })
    if ($missing.Count -eq 0) { Write-Host '  target apps already include Outlook'; return }
    $apps = [System.Collections.Generic.List[object]]::new()
    foreach ($id in @($current) + $missing) {
        $kind, $value = $id.Split(':', 2)
        if ($kind -eq 'ios') {
            $apps.Add(@{ '@odata.type' = '#microsoft.graph.managedMobileApp'; mobileAppIdentifier = @{ '@odata.type' = '#microsoft.graph.iosMobileAppIdentifier'; bundleId = $value } })
        }
        else {
            $apps.Add(@{ '@odata.type' = '#microsoft.graph.managedMobileApp'; mobileAppIdentifier = @{ '@odata.type' = '#microsoft.graph.androidMobileAppIdentifier'; packageId = $value } })
        }
    }
    $null = Invoke-Graph -Method POST -Uri "$Beta/deviceAppManagement/$Collection/$(Get-Field $Policy 'id')/targetApps" -Body @{ apps = $apps.ToArray(); appGroupType = 'selectedPublicApps' }
    Write-Host "  target apps updated: added $($missing -join ', ')"
}

function Add-PolicyGroupAssignment {
    <# The assign action replaces the whole assignment list, so existing assignments are re-sent together with the new group. #>
    param([Parameter(Mandatory)] [string]$Collection, [Parameter(Mandatory)] [object]$Policy, [Parameter(Mandatory)] [string]$GroupId)
    $existing = @()
    $assignments = Get-Field $Policy 'assignments'
    if ($null -ne $assignments) { $existing = @($assignments) }
    foreach ($a in $existing) {
        $target = Get-Field $a 'target'
        if ([string](Get-Field $target 'groupId') -eq $GroupId -and [string](Get-Field $target '@odata.type') -eq '#microsoft.graph.groupAssignmentTarget') {
            Write-Host '  already assigned to the group'
            return
        }
    }
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($a in $existing) {
        $list.Add(@{ '@odata.type' = '#microsoft.graph.targetedManagedAppPolicyAssignment'; target = (Get-Field $a 'target') })
    }
    $list.Add(@{
            '@odata.type' = '#microsoft.graph.targetedManagedAppPolicyAssignment'
            target        = @{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = $GroupId }
        })
    $null = Invoke-Graph -Method POST -Uri "$Beta/deviceAppManagement/$Collection/$(Get-Field $Policy 'id')/assign" -Body @{ assignments = $list.ToArray() }
    Write-Host "  assigned to group $GroupId"
}

# ----------------------------------------------------------------------------------------------------------------
Write-Step "Connecting to Microsoft Graph (DeviceManagementApps.ReadWrite.All, Group.Read.All)$(if ($UseDeviceCode) { ' (device code)' })"
$scopes = @('DeviceManagementApps.ReadWrite.All', 'Group.Read.All')
if ($UseDeviceCode) { Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome -ErrorAction Stop }
else {
    try { Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop }
    catch {
        Write-Warning "Interactive Graph sign-in failed ($($_.Exception.Message)). Falling back to device code, which is expected inside Windows Sandbox."
        Connect-MgGraph -Scopes $scopes -UseDeviceCode -NoWelcome -ErrorAction Stop
    }
}

Write-Step "Resolving group '$GroupName'"
$group = @(Get-MgGroup -Filter "displayName eq '$($GroupName.Replace("'", "''"))'" -Property Id, DisplayName -Top 2)
if ($group.Count -ne 1) { throw "Expected exactly one group named '$GroupName', found $($group.Count)." }
$groupId = [string]$group[0].Id
Write-Host "  $groupId"

# ----------------------------------------------------------------------------------------------------------------
# (a) App Protection Policy
# ----------------------------------------------------------------------------------------------------------------
Write-Step "Ensuring iOS App Protection Policy '$AppProtectionPolicyName'"
$app = Find-PolicyByName -Collection 'iosManagedAppProtections' -DisplayName $AppProtectionPolicyName
if (-not $app) {
    $created = Invoke-Graph -Method POST -Uri "$Beta/deviceAppManagement/iosManagedAppProtections" -Body @{
        '@odata.type'                             = '#microsoft.graph.iosManagedAppProtection'
        displayName                               = $AppProtectionPolicyName
        description                               = 'Outlook on Intune-managed iPhones. Allows Outlook to sync contacts to the native Contacts app (Company Directory caller ID). Created by New-OutlookContactSyncPolicies.ps1.'
        # The gate for native contact sync
        contactSyncBlocked                        = $false
        # Apps on Intune-managed devices only
        targetedAppManagementLevels               = 'mdm'
        appGroupType                              = 'selectedPublicApps'
        # Conservative data-transfer baseline; review in the portal
        allowedInboundDataTransferSources         = 'managedApps'
        allowedOutboundDataTransferDestinations   = 'managedApps'
        allowedOutboundClipboardSharingLevel      = 'managedAppsWithPasteIn'
        saveAsBlocked                             = $true
        allowedDataStorageLocations               = @('oneDriveForBusiness', 'sharePoint')
        dataBackupBlocked                         = $true
        organizationalCredentialsRequired         = $false
        managedBrowserToOpenLinksRequired         = $false
        printBlocked                              = $false
        deviceComplianceRequired                  = $true
        pinRequired                               = $false
        faceIdBlocked                             = $false
        periodOfflineBeforeAccessCheck            = 'PT12H'
        periodOnlineBeforeAccessCheck             = 'PT30M'
        periodOfflineBeforeWipeIsEnforced         = 'P90D'
    }
    $app = Invoke-Graph -Method GET -Uri "$Beta/deviceAppManagement/iosManagedAppProtections/$(Get-Field $created 'id')?`$expand=apps,assignments"
    Write-Host "  created $(Get-Field $app 'id') (new policy: review data protection settings in the portal)"
}
else {
    $blocked = Get-Field $app 'contactSyncBlocked'
    if ($blocked -eq $false) {
        Write-Host "  exists $(Get-Field $app 'id'); 'Sync policy managed app data with native apps or add-ins' already Allow"
    }
    else {
        $null = Invoke-Graph -Method PATCH -Uri "$Beta/deviceAppManagement/iosManagedAppProtections/$(Get-Field $app 'id')" -Body @{
            '@odata.type'      = '#microsoft.graph.iosManagedAppProtection'
            contactSyncBlocked = $false
        }
        Write-Host "  exists $(Get-Field $app 'id'); set 'Sync policy managed app data with native apps or add-ins' = Allow (other settings untouched)"
    }
    $level = [string](Get-Field $app 'targetedAppManagementLevels')
    if ($level -and $level -notmatch 'mdm') {
        Write-Warning "APP '$AppProtectionPolicyName' targets '$level', not apps on Intune-managed devices (mdm). Corporate iPhones are enrolled, so check this is the policy they receive."
    }
}
Set-PolicyTargetApps -Collection 'iosManagedAppProtections' -Policy $app -Required @("ios:$OutlookIosBundleId")
Add-PolicyGroupAssignment -Collection 'iosManagedAppProtections' -Policy $app -GroupId $groupId

# ----------------------------------------------------------------------------------------------------------------
# (b) App Configuration Policy (managed apps)
# ----------------------------------------------------------------------------------------------------------------
Write-Step "Ensuring App Configuration Policy '$AppConfigurationPolicyName'"
$desired = [System.Collections.Generic.List[object]]::new()
foreach ($k in $OutlookContactSyncSettings.Keys) { $desired.Add(@{ name = $k; value = [string]$OutlookContactSyncSettings[$k] }) }

$acp = Find-PolicyByName -Collection 'targetedManagedAppConfigurations' -DisplayName $AppConfigurationPolicyName
if (-not $acp) {
    $created = Invoke-Graph -Method POST -Uri "$Beta/deviceAppManagement/targetedManagedAppConfigurations" -Body @{
        displayName    = $AppConfigurationPolicyName
        description    = 'Outlook for iOS and Android: Save Contacts locked on, Company Directory field list. Created by New-OutlookContactSyncPolicies.ps1.'
        appGroupType   = 'selectedPublicApps'
        customSettings = $desired.ToArray()
    }
    $acp = Invoke-Graph -Method GET -Uri "$Beta/deviceAppManagement/targetedManagedAppConfigurations/$(Get-Field $created 'id')?`$expand=apps,assignments"
    Write-Host "  created $(Get-Field $acp 'id')"
}
else {
    # Merge: keep every existing key not managed here, overwrite the managed keys.
    $merged = [ordered]@{}
    $existingSettings = Get-Field $acp 'customSettings'
    if ($null -ne $existingSettings) {
        foreach ($s in $existingSettings) { $merged[[string](Get-Field $s 'name')] = [string](Get-Field $s 'value') }
    }
    $changed = $false
    foreach ($k in $OutlookContactSyncSettings.Keys) {
        if (-not $merged.Contains($k) -or $merged[$k] -ne [string]$OutlookContactSyncSettings[$k]) { $changed = $true }
        $merged[$k] = [string]$OutlookContactSyncSettings[$k]
    }
    if ($changed) {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($k in $merged.Keys) { $list.Add(@{ name = $k; value = $merged[$k] }) }
        $null = Invoke-Graph -Method PATCH -Uri "$Beta/deviceAppManagement/targetedManagedAppConfigurations/$(Get-Field $acp 'id')" -Body @{ customSettings = $list.ToArray() }
        Write-Host "  exists $(Get-Field $acp 'id'); custom settings updated ($($merged.Count) keys, other keys preserved)"
    }
    else { Write-Host "  exists $(Get-Field $acp 'id'); custom settings already correct" }
}
Set-PolicyTargetApps -Collection 'targetedManagedAppConfigurations' -Policy $acp -Required @("ios:$OutlookIosBundleId", "android:$OutlookAndroidPackageId")
Add-PolicyGroupAssignment -Collection 'targetedManagedAppConfigurations' -Policy $acp -GroupId $groupId

# ----------------------------------------------------------------------------------------------------------------
# Manual verification
# ----------------------------------------------------------------------------------------------------------------
Write-Host ''
Write-Host '================================================================================' -ForegroundColor Green
Write-Host ' Done. Verify by hand (this script does not edit device restriction profiles):' -ForegroundColor Green
Write-Host ''
Write-Host ' Intune > Devices > iOS/iPadOS > Configuration > any Device restrictions profile assigned to corporate iPhones'
Write-Host ' > App Store, Doc Viewing, Gaming:'
Write-Host "   1. 'Block viewing non-corporate documents in corporate apps'               must NOT be Yes"
Write-Host "   2. If 'Block viewing corporate documents in unmanaged apps' is Yes, then"
Write-Host "      'Allow managed apps to write contacts to unmanaged contacts accounts'  must be Yes (Not configured is not enough)"
Write-Host "   3. 'Allow unmanaged apps to read from managed contacts accounts'         optional, your choice"
Write-Host ' Get item 2 wrong and Outlook contact sync fails silently on the phone.'
Write-Host ''
Write-Host ' User experience: Outlook shows a one-time notice that the administrator enabled contact sync and iOS asks for'
Write-Host ' Contacts access. The user taps Allow. If they tap Don''t Allow, Outlook shows an Open Settings link under the account.'
Write-Host '================================================================================' -ForegroundColor Green

[pscustomobject]@{
    AppProtectionPolicyId      = Get-Field $app 'id'
    AppConfigurationPolicyId   = Get-Field $acp 'id'
    AssignedGroupId            = $groupId
    DeviceRestrictionsToVerify = @(
        "'Block viewing non-corporate documents in corporate apps' is not Yes",
        "If 'Block viewing corporate documents in unmanaged apps' is Yes: 'Allow managed apps to write contacts to unmanaged contacts accounts' is Yes"
    )
}
