<#
.SYNOPSIS
    Azure Automation runbook (PowerShell 7.4). Reads corporate-owned iPhones and their primary users from Intune,
    stamps those users as in scope (extensionAttribute1 = CorpiPhone) and writes each phone's number to the
    user's mobilePhone attribute in Entra ID.

.DESCRIPTION
    Runs under the Automation account's system-assigned managed identity (Connect-MgGraph -Identity).

    Required Microsoft.Graph submodules in the runtime environment (pinned, see deploy/Deploy-Automation.ps1):
        Microsoft.Graph.Authentication
        Microsoft.Graph.Users
        Microsoft.Graph.Groups
        Microsoft.Graph.DeviceManagement

    Required Graph application permissions on the managed identity:
        DeviceManagementManagedDevices.Read.All   read managed devices
        User.ReadWrite.All                        read users, write extensionAttribute1 and mobilePhone
        GroupMember.Read.All                      read the exclusion group
        RoleManagement.Read.Directory             read active directory role assignments

    Logic
      1. Read managed devices where operatingSystem eq 'iOS' and managedDeviceOwnerType eq 'company'. Drop devices without a primary user.
      2. In-scope set = distinct primary users that are enabled member users.
      3. Stamp: set extensionAttribute1 to -ScopeAttributeValue on in-scope users where it differs; clear it on users that carry
         the value but are no longer in scope (including disabled users). Only extensionAttribute1 is ever written.
         If any user carries a different, foreign value in extensionAttribute1 the runbook stops before writing anything.
      4. Phone: per user keep the device with the newest lastSyncDateTime, normalize its number to E.164 and write mobilePhone
         when it differs. The phone write (never the stamp) is skipped for exclusion-group members and for any principal
         holding an active directory role assignment.
      5. Emit one log line per change and a summary object. Throw at the end if any error occurred so the job shows Failed.

    This runbook never calls the Entra authentication methods API. mobilePhone is a profile attribute only.

.PARAMETER WhatIf
    $true (default) logs every change it would make without writing. Set $false to apply. Bool, not a switch: Automation
    passes bools and strings.
.PARAMETER ExclusionGroupName
    Display name of the static group whose members do not get their phone number written. Empty string disables the check.
.PARAMETER DefaultCountryCode
    Prefix added to 10-digit national numbers. Default '+1'.
.PARAMETER ScopeAttributeValue
    Value written to extensionAttribute1 for in-scope users. Default 'CorpiPhone'. Must match the dynamic group rule
    and the Exchange management scope filter.

.OUTPUTS
    Log lines (strings) followed by one summary [pscustomobject].

.NOTES
    Known limits: Intune refreshes hardware inventory about every 7 days, so a SIM swap can take up to a week to appear.
    A brand-new enrollment may report its primary user before its number: the stamp lands first, the number follows.
    Some carriers and eSIM configurations never report a number; those users stay stamped without a mobile number.
#>
param(
    [bool]$WhatIf = $true,
    [string]$ExclusionGroupName = 'SG-Directory-Exclude',
    [string]$DefaultCountryCode = '+1',
    [string]$ScopeAttributeValue = 'CorpiPhone'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# The only extension attribute this runbook reads or writes. Change here, in the dynamic group rule and in the Exchange scope together.
$ScopeAttributeName = 'extensionAttribute1'
$RequiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Groups', 'Microsoft.Graph.DeviceManagement')
$GraphBase = 'https://graph.microsoft.com/v1.0'

#region ---------------------------------------------------------------- helpers ------------------------------------------------------------
function Write-SyncLog {
    <# Writes a timestamped line to the job output stream. Never Write-Host: Automation keeps the output stream in job history. #>
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ValidateSet('INFO', 'CHANGE', 'SKIP', 'WARN', 'ERROR')] [string]$Level = 'INFO'
    )
    $line = '{0:yyyy-MM-ddTHH:mm:ssZ} [{1}] {2}' -f [DateTime]::UtcNow, $Level, $Message
    if ($Level -eq 'WARN') {
        # Warnings go to the warning stream, which Automation records separately. Helper functions that return values
        # (retry, batch) log only at this level, so their output stream carries nothing but the return value.
        Write-Warning -Message $line
    }
    else {
        Write-Output $line
    }
}

function Get-Field {
    <# Strict-mode-safe property read that works for hashtables (Invoke-MgGraphRequest output) and objects (SDK output). Returns $null when absent. #>
    param([object]$Object, [Parameter(Mandatory)] [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Get-HttpStatusCode {
    <# Best-effort extraction of the HTTP status code from an Invoke-MgGraphRequest error record. #>
    param([Parameter(Mandatory)] [System.Management.Automation.ErrorRecord]$ErrorRecord)
    $ex = $ErrorRecord.Exception
    while ($null -ne $ex) {
        foreach ($candidate in @((Get-Field -Object $ex -Name 'StatusCode'), (Get-Field -Object (Get-Field -Object $ex -Name 'Response') -Name 'StatusCode'))) {
            if ($null -eq $candidate) { continue }
            if ($candidate -is [System.Enum]) { return [int]$candidate }
            $code = 0
            if ([int]::TryParse([string]$candidate, [ref]$code)) { return $code }
        }
        $ex = $ex.InnerException
    }
    $text = [string]$ErrorRecord.Exception.Message
    if ($text -match '\b(429|503|504|403|404|400|401)\b') { return [int]$Matches[1] }
    return 0
}

function Invoke-GraphRequestWithRetry {
    <#
        Invoke-MgGraphRequest with retry on 429/503/504. The Graph SDK already honors Retry-After inside its own retry handler;
        this outer loop covers the cases where the SDK gives up, with exponential backoff capped at 60 seconds.
    #>
    param(
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Uri,
        [string]$Body,
        [int]$MaxAttempts = 6
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($Method -in @('POST', 'PATCH') -and $null -ne $Body) {
                return Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Body -ContentType 'application/json'
            }
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri
        }
        catch {
            $status = Get-HttpStatusCode -ErrorRecord $_
            if ($status -in @(429, 503, 504) -and $attempt -lt $MaxAttempts) {
                $delay = [Math]::Min(60, [Math]::Pow(2, $attempt)) + (Get-Random -Minimum 0 -Maximum 3)
                Write-SyncLog -Level WARN -Message "Graph $Method $Uri returned $status, retrying in $delay s (attempt $attempt of $MaxAttempts)"
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Get-GraphCollection {
    <# GET with @odata.nextLink paging. Returns an array of items (hashtables). #>
    param([Parameter(Mandatory)] [string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-GraphRequestWithRetry -Method GET -Uri $next
        $value = Get-Field -Object $page -Name 'value'
        if ($null -ne $value) { foreach ($v in $value) { $items.Add($v) } }
        $next = Get-Field -Object $page -Name '@odata.nextLink'
    }
    return $items.ToArray()
}

function ConvertTo-E164 {
    <#
        Normalizes a phone number as reported by Intune to E.164.
          - strips spaces, dashes, dots and parentheses
          - 10 digits                      -> DefaultCountryCode + digits        (8045550123      -> +18045550123)
          - 11 digits starting with 1      -> + digits                          (1 804 555 0123  -> +18045550123)
          - explicit + with 8 to 15 digits -> + digits                          (+44 20 7946 0958 -> +442079460958)
          - anything else                  -> $null (caller logs and skips)
    #>
    param(
        [AllowNull()] [AllowEmptyString()] [string]$Number,
        [string]$DefaultCountryCode = '+1'
    )
    if ([string]::IsNullOrWhiteSpace($Number)) { return $null }
    $trimmed = $Number.Trim()
    $hasPlus = $trimmed.StartsWith('+')
    $digits = ($trimmed -replace '[^0-9]', '')
    if ($digits.Length -eq 0) { return $null }
    if ($hasPlus) {
        if ($digits.Length -ge 8 -and $digits.Length -le 15) { return "+$digits" }
        return $null
    }
    if ($digits.Length -eq 10) {
        $cc = ($DefaultCountryCode -replace '[^0-9]', '')
        if ($cc.Length -eq 0) { return $null }
        return "+$cc$digits"
    }
    if ($digits.Length -eq 11 -and $digits.StartsWith('1')) { return "+$digits" }
    return $null
}

function ConvertTo-DeviceRecord {
    <# Normalizes a managedDevice (SDK object or hashtable) into a flat record with fixed properties. #>
    param([Parameter(Mandatory)] [object]$Device)
    $sync = Get-Field -Object $Device -Name 'lastSyncDateTime'
    if ($null -eq $sync) { $sync = Get-Field -Object $Device -Name 'LastSyncDateTime' }
    $syncDate = [DateTime]::MinValue
    if ($null -ne $sync -and -not [string]::IsNullOrWhiteSpace([string]$sync)) {
        try { $syncDate = [DateTime]::Parse([string]$sync, [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal) } catch { $syncDate = [DateTime]::MinValue }
    }
    $pick = {
        param($camel, $pascal)
        $v = Get-Field -Object $Device -Name $camel
        if ($null -eq $v) { $v = Get-Field -Object $Device -Name $pascal }
        if ($null -eq $v) { return '' }
        return [string]$v
    }
    [pscustomobject]@{
        Id                = & $pick 'id' 'Id'
        UserId            = & $pick 'userId' 'UserId'
        UserPrincipalName = & $pick 'userPrincipalName' 'UserPrincipalName'
        PhoneNumber       = & $pick 'phoneNumber' 'PhoneNumber'
        LastSyncDateTime  = $syncDate
        DeviceName        = & $pick 'deviceName' 'DeviceName'
        SerialNumber      = & $pick 'serialNumber' 'SerialNumber'
    }
}

function Select-NewestDevicePerUser {
    <# Groups device records by UserId and keeps the one with the newest LastSyncDateTime. Returns hashtable UserId -> record. #>
    param([AllowEmptyCollection()] [object[]]$Devices)
    $byUser = @{}
    foreach ($d in $Devices) {
        if ([string]::IsNullOrWhiteSpace($d.UserId)) { continue }
        $key = $d.UserId.ToLowerInvariant()
        if (-not $byUser.ContainsKey($key) -or $d.LastSyncDateTime -gt $byUser[$key].LastSyncDateTime) {
            $byUser[$key] = $d
        }
    }
    return $byUser
}

function ConvertTo-UserRecord {
    <# Normalizes a user (SDK object or hashtable) into a flat record. #>
    param([Parameter(Mandatory)] [object]$User, [string]$AttributeName = 'extensionAttribute1')
    $ext = Get-Field -Object $User -Name 'onPremisesExtensionAttributes'
    if ($null -eq $ext) { $ext = Get-Field -Object $User -Name 'OnPremisesExtensionAttributes' }
    $attrValue = $null
    if ($null -ne $ext) {
        $attrValue = Get-Field -Object $ext -Name $AttributeName
        if ($null -eq $attrValue) {
            $pascal = $AttributeName.Substring(0, 1).ToUpperInvariant() + $AttributeName.Substring(1)
            $attrValue = Get-Field -Object $ext -Name $pascal
        }
    }
    $enabled = Get-Field -Object $User -Name 'accountEnabled'
    if ($null -eq $enabled) { $enabled = Get-Field -Object $User -Name 'AccountEnabled' }
    $id = Get-Field -Object $User -Name 'id'
    if ($null -eq $id) { $id = Get-Field -Object $User -Name 'Id' }
    $upn = Get-Field -Object $User -Name 'userPrincipalName'
    if ($null -eq $upn) { $upn = Get-Field -Object $User -Name 'UserPrincipalName' }
    $mobile = Get-Field -Object $User -Name 'mobilePhone'
    if ($null -eq $mobile) { $mobile = Get-Field -Object $User -Name 'MobilePhone' }
    [pscustomobject]@{
        Id                = [string]$id
        UserPrincipalName = [string]$upn
        AccountEnabled    = [bool]$enabled
        ScopeAttribute    = if ([string]::IsNullOrEmpty([string]$attrValue)) { $null } else { [string]$attrValue }
        MobilePhone       = if ([string]::IsNullOrEmpty([string]$mobile)) { $null } else { [string]$mobile }
    }
}

function Resolve-StampChanges {
    <#
        Pure decision function. Given all member users and the set of in-scope user ids, returns
          ToStamp   : enabled in-scope users whose attribute is not already the scope value
          ToUnstamp : users carrying the scope value that are not in scope (or disabled)
          Conflicts : users whose attribute holds a foreign value (anything other than empty or the scope value)
    #>
    param(
        [AllowEmptyCollection()] [object[]]$Users,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]]$InScopeUserIds,
        [Parameter(Mandatory)] [string]$ScopeAttributeValue
    )
    $toStamp = [System.Collections.Generic.List[object]]::new()
    $toUnstamp = [System.Collections.Generic.List[object]]::new()
    $conflicts = [System.Collections.Generic.List[object]]::new()
    foreach ($u in $Users) {
        $inScope = $u.AccountEnabled -and $InScopeUserIds.Contains($u.Id)
        $current = if ([string]::IsNullOrEmpty([string]$u.ScopeAttribute)) { $null } else { [string]$u.ScopeAttribute }
        if ($null -ne $current -and $current -ne $ScopeAttributeValue) {
            $conflicts.Add($u)
            continue
        }
        if ($inScope -and $current -ne $ScopeAttributeValue) { $toStamp.Add($u) }
        elseif (-not $inScope -and $current -eq $ScopeAttributeValue) { $toUnstamp.Add($u) }
    }
    [pscustomobject]@{
        ToStamp   = $toStamp.ToArray()
        ToUnstamp = $toUnstamp.ToArray()
        Conflicts = $conflicts.ToArray()
    }
}

function Set-UserScopeAttribute {
    <# PATCHes exactly one extension attribute. $Value $null clears it. Serialized by hand so null survives. #>
    param(
        [Parameter(Mandatory)] [string]$UserId,
        [Parameter(Mandatory)] [string]$AttributeName,
        [AllowNull()] [AllowEmptyString()] [string]$Value
    )
    # A [string] parameter turns $null into '', so map empty back to a JSON null, which is what clears the attribute.
    $jsonValue = if ([string]::IsNullOrEmpty($Value)) { $null } else { $Value }
    $body = @{ onPremisesExtensionAttributes = @{ $AttributeName = $jsonValue } } | ConvertTo-Json -Depth 5 -Compress
    $null = Invoke-GraphRequestWithRetry -Method PATCH -Uri "$GraphBase/users/$UserId" -Body $body
}

function Set-UserMobilePhone {
    param([Parameter(Mandatory)] [string]$UserId, [Parameter(Mandatory)] [string]$MobilePhone)
    $body = @{ mobilePhone = $MobilePhone } | ConvertTo-Json -Compress
    $null = Invoke-GraphRequestWithRetry -Method PATCH -Uri "$GraphBase/users/$UserId" -Body $body
}

function Get-GroupIdByName {
    param([Parameter(Mandatory)] [string]$DisplayName)
    $escaped = $DisplayName.Replace("'", "''")
    $g = @(Get-MgGroup -Filter "displayName eq '$escaped'" -Property Id, DisplayName -Top 2)
    if ($g.Count -eq 0) { return $null }
    if ($g.Count -gt 1) { throw "More than one group is named '$DisplayName'. Group names used by the runbooks must be unique." }
    return [string]$g[0].Id
}

function Get-GroupTransitiveUserIds {
    param([Parameter(Mandatory)] [string]$GroupId)
    $members = @(Get-GraphCollection -Uri "$GraphBase/groups/$GroupId/transitiveMembers/microsoft.graph.user?`$select=id&`$top=999")
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $members) {
        $id = Get-Field -Object $m -Name 'id'
        if ($id) { $null = $set.Add([string]$id) }
    }
    return , $set   # comma: return the HashSet itself, not its enumerated members
}

function Get-PrivilegedUserIds {
    <#
        Every principal in an active directory role assignment. Group principals are expanded to their transitive user members
        (role-assignable groups). Service principals are ignored. PIM-eligible assignments are not active and are not included.
    #>
    param([Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]]$KnownUserIds)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $assignments = @(Get-GraphCollection -Uri "$GraphBase/roleManagement/directory/roleAssignments?`$select=principalId")
    $principalIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in $assignments) {
        $principalId = Get-Field -Object $a -Name 'principalId'
        if ($principalId) { $null = $principalIds.Add([string]$principalId) }
    }
    foreach ($principalId in $principalIds) {
        if ($KnownUserIds.Contains($principalId)) { $null = $set.Add($principalId); continue }
        # Not a known user: may be a group (expand) or a service principal (ignore).
        try {
            $members = Get-GroupTransitiveUserIds -GroupId $principalId
            foreach ($m in $members) { $null = $set.Add($m) }
        }
        catch {
            $status = Get-HttpStatusCode -ErrorRecord $_
            if ($status -ne 404 -and $status -ne 400) { throw }
        }
    }
    return , $set   # comma: return the HashSet itself, not its enumerated members
}
#endregion

#region ------------------------------------------------------------------ main -------------------------------------------------------------
$summary = [ordered]@{
    WhatIf           = $WhatIf
    DevicesRead      = 0
    DevicesWithoutPrimaryUser = 0
    UsersInScope     = 0
    UsersStamped     = 0
    UsersUnstamped   = 0
    NumbersUpdated   = 0
    NumbersUnchanged = 0
    UsersSkipped     = [ordered]@{}
    Errors           = 0
}
$errorMessages = [System.Collections.Generic.List[string]]::new()

function Add-Skip {
    param([Parameter(Mandatory)] [string]$Reason, [Parameter(Mandatory)] [string]$Detail)
    if (-not $summary.UsersSkipped.Contains($Reason)) { $summary.UsersSkipped[$Reason] = 0 }
    $summary.UsersSkipped[$Reason]++
    Write-SyncLog -Level SKIP -Message "$Reason : $Detail"
}

function Add-Error {
    param([Parameter(Mandatory)] [string]$Message)
    $summary.Errors++
    $errorMessages.Add($Message)
    Write-SyncLog -Level ERROR -Message $Message
}

$mode = if ($WhatIf) { 'WHATIF' } else { 'APPLY' }
Write-SyncLog "Sync-IntuneToEntra starting. Mode=$mode ScopeAttribute=$ScopeAttributeName=$ScopeAttributeValue ExclusionGroup='$ExclusionGroupName' DefaultCountryCode=$DefaultCountryCode"

# --- connect -------------------------------------------------------------------------------------------------------
foreach ($m in $RequiredModules) { Import-Module -Name $m -ErrorAction Stop }
Connect-MgGraph -Identity -NoWelcome
$ctx = Get-MgContext
if (-not $ctx) { throw 'Connect-MgGraph -Identity did not produce a Graph context. Is the system-assigned managed identity enabled?' }
Write-SyncLog "Connected to Graph as managed identity (client $($ctx.ClientId)) in tenant $($ctx.TenantId)"

# --- 1. devices ------------------------------------------------------------------------------------------------------
$rawDevices = @(Get-MgDeviceManagementManagedDevice -All -PageSize 999 `
        -Filter "operatingSystem eq 'iOS' and managedDeviceOwnerType eq 'company'" `
        -Property Id, UserId, UserPrincipalName, PhoneNumber, LastSyncDateTime, DeviceName, SerialNumber)
$devices = @($rawDevices | ForEach-Object { ConvertTo-DeviceRecord -Device $_ })
$summary.DevicesRead = $devices.Count
$withUser = @($devices | Where-Object { -not [string]::IsNullOrWhiteSpace($_.UserId) })
$summary.DevicesWithoutPrimaryUser = $devices.Count - $withUser.Count
foreach ($d in ($devices | Where-Object { [string]::IsNullOrWhiteSpace($_.UserId) })) {
    Write-SyncLog -Level WARN -Message "Device '$($d.DeviceName)' serial $($d.SerialNumber) has no primary user and is ignored. Set the primary user in Intune."
}
Write-SyncLog "Read $($devices.Count) corporate iOS devices, $($withUser.Count) with a primary user"

# --- 2. users --------------------------------------------------------------------------------------------------------
$rawUsers = @(Get-MgUser -All -PageSize 999 -ConsistencyLevel eventual -CountVariable memberCount `
        -Filter "userType eq 'Member'" `
        -Property Id, UserPrincipalName, AccountEnabled, MobilePhone, OnPremisesExtensionAttributes)
$users = @($rawUsers | ForEach-Object { ConvertTo-UserRecord -User $_ -AttributeName $ScopeAttributeName })
$userById = @{}
foreach ($u in $users) { $userById[$u.Id.ToLowerInvariant()] = $u }
$knownUserIds = [System.Collections.Generic.HashSet[string]]::new([string[]]$userById.Keys, [StringComparer]::OrdinalIgnoreCase)
Write-SyncLog "Read $($users.Count) member users"

$inScopeIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($d in $withUser) {
    $key = $d.UserId.ToLowerInvariant()
    if (-not $userById.ContainsKey($key)) {
        Add-Skip -Reason 'DeviceUserNotMember' -Detail "device '$($d.DeviceName)' serial $($d.SerialNumber) primary user $($d.UserPrincipalName) ($($d.UserId)) is not a member user"
        continue
    }
    if (-not $userById[$key].AccountEnabled) {
        Add-Skip -Reason 'DeviceUserDisabled' -Detail "device '$($d.DeviceName)' serial $($d.SerialNumber) primary user $($d.UserPrincipalName) is disabled"
        continue
    }
    $null = $inScopeIds.Add($key)
}
$summary.UsersInScope = $inScopeIds.Count
Write-SyncLog "In-scope users (enabled primary users of corporate iPhones): $($inScopeIds.Count)"

# --- 3. exclusions for the phone write (resolved before any write so a missing group fails the job cleanly) ---------------------------------------------------------------------------------
$excludedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
if (-not [string]::IsNullOrWhiteSpace($ExclusionGroupName)) {
    $exclusionGroupId = Get-GroupIdByName -DisplayName $ExclusionGroupName
    if (-not $exclusionGroupId) {
        throw "Exclusion group '$ExclusionGroupName' was not found. Refusing to write phone numbers without it. Create the group or pass an empty -ExclusionGroupName to disable the check."
    }
    $excludedIds = Get-GroupTransitiveUserIds -GroupId $exclusionGroupId
    Write-SyncLog "Exclusion group '$ExclusionGroupName' has $($excludedIds.Count) user members"
}
$privilegedIds = Get-PrivilegedUserIds -KnownUserIds $knownUserIds
Write-SyncLog "Users holding an active directory role assignment: $($privilegedIds.Count)"

# --- 4. stamp --------------------------------------------------------------------------------------------------------
$plan = Resolve-StampChanges -Users $users -InScopeUserIds $inScopeIds -ScopeAttributeValue $ScopeAttributeValue
if ($plan.Conflicts.Count -gt 0) {
    foreach ($c in $plan.Conflicts) {
        Write-SyncLog -Level ERROR -Message "CONFLICT $($c.UserPrincipalName) has $ScopeAttributeName='$($c.ScopeAttribute)', which is neither empty nor '$ScopeAttributeValue'."
    }
    throw "$($plan.Conflicts.Count) user(s) carry a foreign value in $ScopeAttributeName. This runbook owns that attribute and will not overwrite it. Clear the attribute or choose another one (see README). No changes were made."
}

foreach ($u in $plan.ToStamp) {
    try {
        if (-not $WhatIf) { Set-UserScopeAttribute -UserId $u.Id -AttributeName $ScopeAttributeName -Value $ScopeAttributeValue }
        $summary.UsersStamped++
        Write-SyncLog -Level CHANGE -Message "$mode STAMP   $($u.UserPrincipalName) $ScopeAttributeName '' -> '$ScopeAttributeValue'"
    }
    catch { Add-Error "Stamp failed for $($u.UserPrincipalName): $($_.Exception.Message)" }
}
foreach ($u in $plan.ToUnstamp) {
    try {
        if (-not $WhatIf) { Set-UserScopeAttribute -UserId $u.Id -AttributeName $ScopeAttributeName -Value $null }
        $summary.UsersUnstamped++
        $why = if (-not $u.AccountEnabled) { 'account disabled' } else { 'no longer primary user of a corporate iPhone' }
        Write-SyncLog -Level CHANGE -Message "$mode UNSTAMP $($u.UserPrincipalName) $ScopeAttributeName '$ScopeAttributeValue' -> '' ($why)"
    }
    catch { Add-Error "Unstamp failed for $($u.UserPrincipalName): $($_.Exception.Message)" }
}
Write-SyncLog "Stamp pass done: $($summary.UsersStamped) stamped, $($summary.UsersUnstamped) unstamped"

# --- 5. phone numbers ---------------------------------------------------------------------------------------------------
$newestByUser = Select-NewestDevicePerUser -Devices $withUser
foreach ($key in ($inScopeIds | Sort-Object)) {
    $u = $userById[$key]
    $d = $newestByUser[$key]
    if ($null -eq $d) { Add-Skip -Reason 'NoDevice' -Detail "$($u.UserPrincipalName) has no device record (unexpected)"; continue }
    if ([string]::IsNullOrWhiteSpace($d.PhoneNumber)) {
        Add-Skip -Reason 'NoPhoneNumber' -Detail "$($u.UserPrincipalName) newest device '$($d.DeviceName)' serial $($d.SerialNumber) reports no phone number (eSIM/carrier or inventory not yet collected)"
        continue
    }
    $normalized = ConvertTo-E164 -Number $d.PhoneNumber -DefaultCountryCode $DefaultCountryCode
    if (-not $normalized) {
        Add-Skip -Reason 'InvalidPhoneNumber' -Detail "$($u.UserPrincipalName) device serial $($d.SerialNumber) reports '$($d.PhoneNumber)', which does not normalize to E.164"
        continue
    }
    if ($excludedIds.Contains($key)) {
        Add-Skip -Reason 'ExcludedGroupMember' -Detail "$($u.UserPrincipalName) is a member of '$ExclusionGroupName'; number $normalized not written (stamp kept)"
        continue
    }
    if ($privilegedIds.Contains($key)) {
        Add-Skip -Reason 'PrivilegedRole' -Detail "$($u.UserPrincipalName) holds an active directory role; number $normalized not written (stamp kept). Set it by hand if wanted."
        continue
    }
    if ($u.MobilePhone -eq $normalized) { $summary.NumbersUnchanged++; continue }
    try {
        if (-not $WhatIf) { Set-UserMobilePhone -UserId $u.Id -MobilePhone $normalized }
        $summary.NumbersUpdated++
        $old = if ($u.MobilePhone) { $u.MobilePhone } else { '(none)' }
        Write-SyncLog -Level CHANGE -Message "$mode PHONE   $($u.UserPrincipalName) mobilePhone $old -> $normalized (device '$($d.DeviceName)' serial $($d.SerialNumber), raw '$($d.PhoneNumber)')"
    }
    catch { Add-Error "mobilePhone write failed for $($u.UserPrincipalName): $($_.Exception.Message)" }
}

# --- 6. summary ----------------------------------------------------------------------------------------------------------
Write-SyncLog "Sync-IntuneToEntra finished. Devices=$($summary.DevicesRead) InScope=$($summary.UsersInScope) Stamped=$($summary.UsersStamped) Unstamped=$($summary.UsersUnstamped) NumbersUpdated=$($summary.NumbersUpdated) Unchanged=$($summary.NumbersUnchanged) Skipped=$(($summary.UsersSkipped.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ',') Errors=$($summary.Errors)"
[pscustomobject]$summary

if ($summary.Errors -gt 0) {
    throw "Sync-IntuneToEntra completed with $($summary.Errors) error(s): $($errorMessages -join ' | ')"
}
#endregion
