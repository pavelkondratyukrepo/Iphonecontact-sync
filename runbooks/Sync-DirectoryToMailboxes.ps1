<#
.SYNOPSIS
    Azure Automation runbook (PowerShell 7.6 runtime; also runs on 7.4). Writes one tagged "Company Directory" contact per coworker into the
    default Contacts folder of every member of the control group, through Microsoft Graph, and keeps them current.

.DESCRIPTION
    Runs under the Automation account's system-assigned managed identity (Connect-MgGraph -Identity).

    Required Microsoft.Graph submodules in the runtime environment (pinned, see deploy/Deploy-Automation.ps1):
        Microsoft.Graph.Authentication
        Microsoft.Graph.Groups

    Required Graph application permissions on the managed identity:
        GroupMember.Read.All   read the control and exclusion groups
        User.ReadUpdate.All    read member profiles (granted for Runbook 1; User.Read.All would be enough for this runbook)

    Contacts access: NO Graph permission in Entra. Do not consent Contacts.ReadWrite to the identity. Mailbox access comes
    only from Exchange RBAC for Applications (deploy/New-ExchangeAppScope.ps1): role 'Application Contacts.ReadWrite' with a
    management scope of CustomAttribute1 -eq 'CorpiPhone'. Entra and Exchange grants are a union, so an unscoped Entra grant
    would remove the scoping.

    Logic
      1. Read the control group's transitive user members. This set is both the directory (who appears as a contact) and the
         target list (whose mailboxes receive contacts). Directory = enabled members minus exclusion-group members.
         Target = all enabled members, or the one user in -UserPrincipalName.
      2. Per target mailbox, list the contacts this runbook owns: those carrying the DirectorySourceId extended property.
         To find them the listing returns every contact in the default folder with the fields the runbook writes; contacts
         without DirectorySourceId are discarded immediately, never logged or kept, and never modified or deleted.
      3. Reconcile, excluding the target's own entry: create when no owned contact has that coworker's id, PATCH when the stored
         hash differs from the hash of the current directory fields, when the fields on the contact itself no longer match
         (edited on the phone; -DetectLocalEdits), or always with -FullRebuild. DELETE owned contacts whose coworker left the
         directory, and DELETE duplicates that share a source id.
      4. Send writes through POST /$batch in groups of 20, honoring Retry-After on 429. Mailboxes are processed sequentially.
         A 403 on a mailbox means the Exchange scope has not caught up with a fresh stamp: logged as skipped, not as an error.
      5. One log line per change, a summary object, and a throw at the end if any error occurred so the job shows Failed.
         The job also fails when mailboxes were targeted but none could be processed (for example every one returned 403,
         which means the Exchange RBAC assignment is broken, not just lagging), and when the run nears the Automation
         3-hour limit: it stops cleanly after $TimeBudgetMinutes minutes and reports how many mailboxes were not reached.

    Concurrency: do not start a manual full run while the scheduled run is active (Monday 08:00). Two overlapping runs can
    both create the same contact; the next run deletes the duplicate, but the user sees it twice until then.

    Ownership tagging: every contact written carries categories ['Company Directory'] and two single-value extended properties
    under one fixed GUID namespace: DirectorySourceId (the coworker's Entra object id) and DirectorySourceHash (SHA256 of the
    written fields). Never change the namespace GUID after the first run or the runbook stops recognizing its own contacts.

.PARAMETER WhatIf
    $true (default) logs every change it would make without writing. Set $false to apply.
.PARAMETER GroupName
    Display name of the control group. Default SG-CorpiPhone-Users.
.PARAMETER ExclusionGroupName
    Display name of the group whose members receive contacts but are not listed. Empty string disables the exclusion.
.PARAMETER UserPrincipalName
    When set, only this member's mailbox is processed (onboarding). The directory is still the whole group.
.PARAMETER FullRebuild
    Rewrite every owned contact regardless of the stored hash. Does not touch contacts the runbook does not own.
.PARAMETER DetectLocalEdits
    $true (default) also compares the fields currently on each owned contact with the directory and rewrites the contact
    when someone edited it on the phone or in Outlook. Set $false if the pilot shows Exchange normalizing a field so that
    every contact is rewritten on every run (look for 'edited in mailbox' on contacts nobody touched).

.OUTPUTS
    Log lines (strings) followed by one summary [pscustomobject].

.NOTES
    Why the default Contacts folder: Outlook for iOS syncs the default folder to the native Contacts app. A subfolder would stay
    inside Outlook. Volume: 80 members is about 6,400 contacts on the first run, roughly 320 batched requests, a few minutes.
#>
param(
    [bool]$WhatIf = $true,
    [string]$GroupName = 'SG-CorpiPhone-Users',
    [string]$ExclusionGroupName = 'SG-Directory-Exclude',
    [string]$UserPrincipalName = '',
    [bool]$FullRebuild = $false,
    [bool]$DetectLocalEdits = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RequiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups')
$GraphBase = 'https://graph.microsoft.com/v1.0'
$BatchSize = 20
$TimeBudgetMinutes = 150   # Azure Automation stops cloud jobs at 3 hours (fair share); stop cleanly well before that.

#region ---------------------------------------------------------------- helpers ------------------------------------------------------------
function Get-DirectoryContactConstants {
    <# The ownership constants. Exposed as a function so the Pester tests can read the same values the runbook uses. #>
    [ordered]@{
        # Fixed namespace GUID for the two extended properties. NEVER change after the first live run.
        NamespaceGuid        = '{b4b9c0a1-5f0e-4c3d-9a7e-2d6f1c8e0a11}'
        SourceIdPropertyId   = 'String {b4b9c0a1-5f0e-4c3d-9a7e-2d6f1c8e0a11} Name DirectorySourceId'
        SourceHashPropertyId = 'String {b4b9c0a1-5f0e-4c3d-9a7e-2d6f1c8e0a11} Name DirectorySourceHash'
        Category             = 'Company Directory'
    }
}

function Write-SyncLog {
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
    <# Strict-mode-safe property read for hashtables and objects. Returns $null when absent. #>
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
    if ($text -match '\b(429|500|502|503|504|403|404|400|401)\b') { return [int]$Matches[1] }
    return 0
}

function Get-RetryAfterFromError {
    <# Best-effort read of a Retry-After header (seconds) off a Graph error record; 0 when absent. #>
    param([Parameter(Mandatory)] [System.Management.Automation.ErrorRecord]$ErrorRecord)
    $resp = Get-Field -Object $ErrorRecord.Exception -Name 'Response'
    $headers = Get-Field -Object $resp -Name 'Headers'
    $ra = Get-Field -Object $headers -Name 'RetryAfter'
    if ($null -eq $ra) { return 0 }
    $delta = Get-Field -Object $ra -Name 'Delta'
    if ($null -ne $delta) { return [int][Math]::Min(120, [Math]::Max(0, ([TimeSpan]$delta).TotalSeconds)) }
    $date = Get-Field -Object $ra -Name 'Date'
    if ($null -ne $date) { return [int][Math]::Min(120, [Math]::Max(0, (([DateTimeOffset]$date) - [DateTimeOffset]::UtcNow).TotalSeconds)) }
    return 0
}

function Invoke-GraphRequestWithRetry {
    param(
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Uri,
        [string]$Body,
        [hashtable]$Headers,
        [int]$MaxAttempts = 6
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $call = @{ Method = $Method; Uri = $Uri }
            if ($null -ne $Headers -and $Headers.Count -gt 0) { $call.Headers = $Headers }
            if ($Method -in @('POST', 'PATCH') -and $null -ne $Body) { $call.Body = $Body; $call.ContentType = 'application/json' }
            return Invoke-MgGraphRequest @call
        }
        catch {
            $status = Get-HttpStatusCode -ErrorRecord $_
            if ($status -in @(429, 500, 502, 503, 504) -and $attempt -lt $MaxAttempts) {
                $delay = [Math]::Min(60, [Math]::Pow(2, $attempt)) + (Get-Random -Minimum 0 -Maximum 3)
                $retryAfter = Get-RetryAfterFromError -ErrorRecord $_
                if ($retryAfter -gt $delay) { $delay = $retryAfter }
                Write-SyncLog -Level WARN -Message "Graph $Method $Uri returned $status, retrying in $delay s (attempt $attempt of $MaxAttempts)"
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Get-GraphCollection {
    param([Parameter(Mandatory)] [string]$Uri, [hashtable]$Headers)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-GraphRequestWithRetry -Method GET -Uri $next -Headers $Headers
        $value = Get-Field -Object $page -Name 'value'
        if ($null -ne $value) { foreach ($v in $value) { $items.Add($v) } }
        $next = Get-Field -Object $page -Name '@odata.nextLink'
    }
    return $items.ToArray()
}

function ConvertTo-NullIfEmpty {
    param([AllowNull()] [object]$Value)
    if ($null -eq $Value) { return $null }
    $s = ([string]$Value).Trim()
    if ($s.Length -eq 0) { return $null }
    return $s
}

function ConvertTo-DirectoryUserRecord {
    <# Normalizes a Graph user (hashtable or object) into a flat record with fixed properties. #>
    param([Parameter(Mandatory)] [object]$User)
    $phones = @()
    $bp = Get-Field -Object $User -Name 'businessPhones'
    if ($null -ne $bp) { $phones = @(@($bp) | ForEach-Object { ConvertTo-NullIfEmpty $_ } | Where-Object { $null -ne $_ }) }
    $enabled = Get-Field -Object $User -Name 'accountEnabled'
    [pscustomobject]@{
        Id                = [string](Get-Field -Object $User -Name 'id')
        UserPrincipalName = [string](Get-Field -Object $User -Name 'userPrincipalName')
        AccountEnabled    = if ($null -eq $enabled) { $true } else { [bool]$enabled }
        DisplayName       = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'displayName')
        GivenName         = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'givenName')
        Surname           = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'surname')
        JobTitle          = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'jobTitle')
        Department        = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'department')
        CompanyName       = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'companyName')
        MobilePhone       = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'mobilePhone')
        BusinessPhones    = [string[]]$phones
        Mail              = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'mail')
        OfficeLocation    = ConvertTo-NullIfEmpty (Get-Field -Object $User -Name 'officeLocation')
    }
}

function ConvertTo-DirectoryContactFields {
    <#
        Builds the contact fields written for one coworker, as an ordered hashtable in the exact shape Graph expects.
        fileAs is 'Last, First'. Missing name parts fall back sensibly so every contact has a displayName and fileAs.
    #>
    param([Parameter(Mandatory)] [object]$User)
    $given = ConvertTo-NullIfEmpty $User.GivenName
    $surname = ConvertTo-NullIfEmpty $User.Surname
    $display = ConvertTo-NullIfEmpty $User.DisplayName
    if (-not $display) {
        $display = ConvertTo-NullIfEmpty (("$given $surname").Trim())
        if (-not $display) { $display = $User.UserPrincipalName }
    }
    $fileAs = if ($surname -and $given) { "$surname, $given" } elseif ($surname) { $surname } elseif ($given) { $given } else { $display }

    $emails = @()
    if ($User.Mail) { $emails = @(@{ address = $User.Mail; name = $display }) }
    $phones = @()
    if ($null -ne $User.BusinessPhones) { $phones = @(@($User.BusinessPhones) | ForEach-Object { ConvertTo-NullIfEmpty $_ } | Where-Object { $null -ne $_ }) }

    [ordered]@{
        givenName      = $given
        surname        = $surname
        displayName    = $display
        fileAs         = $fileAs
        jobTitle       = ConvertTo-NullIfEmpty $User.JobTitle
        department     = ConvertTo-NullIfEmpty $User.Department
        companyName    = ConvertTo-NullIfEmpty $User.CompanyName
        mobilePhone    = ConvertTo-NullIfEmpty $User.MobilePhone
        businessPhones = [string[]]$phones
        emailAddresses = $emails
        officeLocation = ConvertTo-NullIfEmpty $User.OfficeLocation
    }
}

function Get-DirectoryContactHash {
    <#
        SHA256 (lowercase hex) over a canonical, order-fixed serialization of the contact fields. Field values are separated by
        U+001F, list items by U+001E, nulls become empty strings. Stable across runs and machines; changes if any written field changes.
    #>
    param([Parameter(Mandatory)] [System.Collections.IDictionary]$Fields)
    $fs = [char]0x1F
    $rs = [char]0x1E
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($name in @('givenName', 'surname', 'displayName', 'fileAs', 'jobTitle', 'department', 'companyName', 'mobilePhone', 'officeLocation')) {
        $v = if ($Fields.Contains($name)) { $Fields[$name] } else { $null }
        $parts.Add([string]$(if ($null -eq $v) { '' } else { $v }))
    }
    $phones = if ($Fields.Contains('businessPhones') -and $null -ne $Fields['businessPhones']) { @($Fields['businessPhones']) } else { @() }
    $parts.Add(($phones -join $rs))
    $emails = if ($Fields.Contains('emailAddresses') -and $null -ne $Fields['emailAddresses']) { @($Fields['emailAddresses']) } else { @() }
    $parts.Add((@($emails | ForEach-Object { [string](Get-Field -Object $_ -Name 'address') }) -join $rs))
    $canonical = $parts -join $fs
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($canonical))
    }
    finally { $sha.Dispose() }
    return ([System.BitConverter]::ToString($bytes) -replace '-', '').ToLowerInvariant()
}

function New-DirectoryContactBody {
    <# Full request body for POST or PATCH: the fields plus the ownership category and the two extended properties. #>
    param(
        [Parameter(Mandatory)] [System.Collections.IDictionary]$Fields,
        [Parameter(Mandatory)] [string]$SourceId,
        [Parameter(Mandatory)] [string]$Hash
    )
    $c = Get-DirectoryContactConstants
    $body = [ordered]@{}
    foreach ($k in $Fields.Keys) { $body[$k] = $Fields[$k] }
    $body['categories'] = @($c.Category)
    $body['singleValueExtendedProperties'] = @(
        @{ id = $c.SourceIdPropertyId; value = $SourceId },
        @{ id = $c.SourceHashPropertyId; value = $Hash }
    )
    return $body
}

function Get-ObservedContactHash {
    <#
        Hash of the fields as they currently sit on the contact in the mailbox, computed exactly like the hash of the
        fields the runbook writes. Differs from the stored DirectorySourceHash when someone edited the contact on the phone.
        Returns $null when the listing did not include the fields (then local edits cannot be detected for that contact).
    #>
    param([Parameter(Mandatory)] [object]$Contact)
    $hasFields = if ($Contact -is [System.Collections.IDictionary]) { $Contact.Contains('fileAs') } else { $null -ne $Contact.PSObject.Properties['fileAs'] }
    if (-not $hasFields) { return $null }
    $fields = [ordered]@{}
    foreach ($name in @('givenName', 'surname', 'displayName', 'fileAs', 'jobTitle', 'department', 'companyName', 'mobilePhone', 'officeLocation')) {
        $fields[$name] = ConvertTo-NullIfEmpty (Get-Field -Object $Contact -Name $name)
    }
    $phones = Get-Field -Object $Contact -Name 'businessPhones'
    $fields['businessPhones'] = @(@($phones) | ForEach-Object { ConvertTo-NullIfEmpty $_ } | Where-Object { $null -ne $_ })
    $mails = Get-Field -Object $Contact -Name 'emailAddresses'
    $fields['emailAddresses'] = @(@($mails) | Where-Object { $null -ne $_ -and (ConvertTo-NullIfEmpty (Get-Field -Object $_ -Name 'address')) })
    return Get-DirectoryContactHash -Fields $fields
}

function ConvertTo-OwnedContactRecord {
    <# Reads the two extended properties off a contact. Returns $null when the contact does not carry DirectorySourceId (not ours). #>
    param([Parameter(Mandatory)] [object]$Contact)
    $c = Get-DirectoryContactConstants
    $sourceId = $null
    $hash = $null
    $props = Get-Field -Object $Contact -Name 'singleValueExtendedProperties'
    if ($null -ne $props) {
        foreach ($p in @($props)) {
            $propId = [string](Get-Field -Object $p -Name 'id')
            if ($propId -eq $c.SourceIdPropertyId) { $sourceId = [string](Get-Field -Object $p -Name 'value') }
            elseif ($propId -eq $c.SourceHashPropertyId) { $hash = [string](Get-Field -Object $p -Name 'value') }
        }
    }
    if ([string]::IsNullOrWhiteSpace($sourceId)) { return $null }
    [pscustomobject]@{
        ContactId    = [string](Get-Field -Object $Contact -Name 'id')
        DisplayName  = [string](Get-Field -Object $Contact -Name 'displayName')
        SourceId     = $sourceId
        Hash         = $hash
        ObservedHash = Get-ObservedContactHash -Contact $Contact
    }
}

function Get-OwnedContacts {
    <# Lists the runbook-owned contacts in a mailbox's default Contacts folder. Contacts without DirectorySourceId are dropped and never touched. #>
    param([Parameter(Mandatory)] [string]$UserId)
    $c = Get-DirectoryContactConstants
    $filter = "id eq '$($c.SourceIdPropertyId)' or id eq '$($c.SourceHashPropertyId)'"
    # The written fields are selected so phone-side edits can be detected. Contacts without DirectorySourceId are dropped
    # right below, before anything is logged or kept: the runbook never acts on, stores or logs another contact's data.
    $fieldSelect = 'givenName,surname,displayName,fileAs,jobTitle,department,companyName,mobilePhone,businessPhones,emailAddresses,officeLocation'
    $uri = "$GraphBase/users/$UserId/contacts?`$select=id,$fieldSelect&`$expand=singleValueExtendedProperties(`$filter=$filter)&`$top=999"
    $raw = @(Get-GraphCollection -Uri $uri)
    $owned = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $raw) {
        $rec = ConvertTo-OwnedContactRecord -Contact $r
        if ($null -ne $rec) { $owned.Add($rec) }
    }
    return $owned.ToArray()
}

function Resolve-MailboxChanges {
    <#
        Pure reconciliation. Given the directory records, the target's own id and the owned contacts in the mailbox, returns
          Creates   : directory records with no owned contact
          Updates   : @{ ContactId; User; Fields; Hash; Reason } where the stored hash differs from the directory, the fields
                      on the contact were edited (DetectLocalEdits), or always when FullRebuild
          Deletes   : owned contacts whose SourceId left the directory, plus duplicate owned contacts for one SourceId
          Unchanged : count
    #>
    param(
        [AllowEmptyCollection()] [object[]]$Directory,
        [Parameter(Mandatory)] [string]$TargetUserId,
        [AllowEmptyCollection()] [object[]]$OwnedContacts,
        [bool]$FullRebuild = $false,
        [bool]$DetectLocalEdits = $true
    )
    $creates = [System.Collections.Generic.List[object]]::new()
    $updates = [System.Collections.Generic.List[object]]::new()
    $deletes = [System.Collections.Generic.List[object]]::new()
    $unchanged = 0

    $bySource = @{}
    foreach ($o in $OwnedContacts) {
        $key = $o.SourceId.ToLowerInvariant()
        if ($bySource.ContainsKey($key)) { $deletes.Add([pscustomobject]@{ Contact = $o; Reason = 'duplicate' }) }
        else { $bySource[$key] = $o }
    }

    $wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($u in $Directory) {
        if ($u.Id -eq $TargetUserId) { continue }   # nobody gets a contact card for themselves
        $null = $wanted.Add($u.Id)
        $fields = ConvertTo-DirectoryContactFields -User $u
        $hash = Get-DirectoryContactHash -Fields $fields
        $key = $u.Id.ToLowerInvariant()
        if (-not $bySource.ContainsKey($key)) {
            $creates.Add([pscustomobject]@{ User = $u; Fields = $fields; Hash = $hash })
        }
        else {
            $existing = $bySource[$key]
            $observed = Get-Field -Object $existing -Name 'ObservedHash'
            $reason = $null
            if ($FullRebuild) { $reason = 'full rebuild' }
            elseif ($existing.Hash -ne $hash) { $reason = 'directory changed' }
            elseif ($DetectLocalEdits -and $observed -and $observed -ne $hash) { $reason = 'edited in mailbox' }
            if ($reason) { $updates.Add([pscustomobject]@{ ContactId = $existing.ContactId; User = $u; Fields = $fields; Hash = $hash; Reason = $reason }) }
            else { $unchanged++ }
        }
    }
    foreach ($key in $bySource.Keys) {
        if (-not $wanted.Contains($key)) { $deletes.Add([pscustomobject]@{ Contact = $bySource[$key]; Reason = 'left directory' }) }
    }
    [pscustomobject]@{
        Creates   = $creates.ToArray()
        Updates   = $updates.ToArray()
        Deletes   = $deletes.ToArray()
        Unchanged = $unchanged
    }
}

function Get-RetryAfterSeconds {
    param([object]$Headers, [int]$Default = 5)
    if ($null -eq $Headers) { return $Default }
    foreach ($name in @('Retry-After', 'retry-after')) {
        $v = Get-Field -Object $Headers -Name $name
        if ($null -ne $v) {
            $n = 0
            if ([int]::TryParse([string]$v, [ref]$n) -and $n -ge 0) { return [Math]::Min($n, 120) }
        }
    }
    return $Default
}

function Invoke-GraphBatch {
    <#
        Sends $Requests (each @{ id; method; url; body }) through POST /$batch in groups of $BatchSize. Sub-requests that come back
        429/503/504 are retried after Retry-After, up to $MaxRounds times. Returns one record per request: @{ Id; Status; Body }.
    #>
    param(
        [AllowEmptyCollection()] [object[]]$Requests,
        [int]$MaxRounds = 6
    )
    $results = @{}
    $pending = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $Requests) { $pending.Add($r) }
    $round = 0
    while ($pending.Count -gt 0 -and $round -lt $MaxRounds) {
        $round++
        $retry = [System.Collections.Generic.List[object]]::new()
        $maxRetryAfter = 0
        for ($i = 0; $i -lt $pending.Count; $i += $BatchSize) {
            $end = [Math]::Min($i + $BatchSize, $pending.Count) - 1
            $chunk = @($pending[$i..$end])
            $batchRequests = foreach ($r in $chunk) {
                $item = [ordered]@{ id = [string]$r.id; method = $r.method; url = $r.url }
                if ($r.method -in @('POST', 'PATCH')) {
                    $item.headers = @{ 'Content-Type' = 'application/json' }
                    $item.body = $r.body
                }
                $item
            }
            $payload = @{ requests = @($batchRequests) } | ConvertTo-Json -Depth 15
            $response = Invoke-GraphRequestWithRetry -Method POST -Uri "$GraphBase/`$batch" -Body $payload
            $responses = Get-Field -Object $response -Name 'responses'
            $seen = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($resp in @($responses)) {
                $id = [string](Get-Field -Object $resp -Name 'id')
                $null = $seen.Add($id)
                $status = [int](Get-Field -Object $resp -Name 'status')
                if ($status -in @(429, 500, 502, 503, 504)) {
                    $retry.Add(($chunk | Where-Object { [string]$_.id -eq $id } | Select-Object -First 1))
                    $maxRetryAfter = [Math]::Max($maxRetryAfter, (Get-RetryAfterSeconds -Headers (Get-Field -Object $resp -Name 'headers')))
                    continue
                }
                $results[$id] = [pscustomobject]@{ Id = $id; Status = $status; Body = (Get-Field -Object $resp -Name 'body') }
            }
            foreach ($r in $chunk) {
                if (-not $seen.Contains([string]$r.id)) {
                    $results[[string]$r.id] = [pscustomobject]@{ Id = [string]$r.id; Status = 0; Body = @{ error = @{ message = 'no response in batch' } } }
                }
            }
        }
        $pending = $retry
        if ($pending.Count -gt 0) {
            $wait = if ($maxRetryAfter -gt 0) { $maxRetryAfter } else { [Math]::Min(60, [Math]::Pow(2, $round)) }
            Write-SyncLog -Level WARN -Message "$($pending.Count) batched request(s) throttled, retrying in $wait s (round $round of $MaxRounds)"
            Start-Sleep -Seconds $wait
        }
    }
    foreach ($r in $pending) {
        $results[[string]$r.id] = [pscustomobject]@{ Id = [string]$r.id; Status = 429; Body = @{ error = @{ message = "still throttled after $MaxRounds rounds" } } }
    }
    return $results
}

function Get-GraphErrorText {
    param([object]$Body)
    $err = Get-Field -Object $Body -Name 'error'
    if ($null -eq $err) { return [string]$Body }
    return "$(Get-Field -Object $err -Name 'code'): $(Get-Field -Object $err -Name 'message')"
}

function Get-GroupIdByName {
    param([Parameter(Mandatory)] [string]$DisplayName)
    $escaped = $DisplayName.Replace("'", "''")
    $g = @(Get-MgGroup -Filter "displayName eq '$escaped'" -Property Id, DisplayName -Top 2)
    if ($g.Count -eq 0) { return $null }
    if ($g.Count -gt 1) { throw "More than one group is named '$DisplayName'. Group names used by the runbooks must be unique." }
    return [string]$g[0].Id
}
#endregion

#region ------------------------------------------------------------------ main -------------------------------------------------------------
$summary = [ordered]@{
    WhatIf             = $WhatIf
    FullRebuild        = $FullRebuild
    DirectorySize      = 0
    MailboxesTargeted  = 0
    MailboxesProcessed = 0
    ContactsAdded      = 0
    ContactsUpdated    = 0
    ContactsDeleted    = 0
    ContactsUnchanged  = 0
    LocalEditsReverted = 0
    DuplicatesRemoved  = 0
    MailboxesNotReached = 0
    MailboxesSkipped   = [ordered]@{}
    Errors             = 0
}
$errorMessages = [System.Collections.Generic.List[string]]::new()

function Add-Skip {
    param([Parameter(Mandatory)] [string]$Reason, [Parameter(Mandatory)] [string]$Detail)
    if (-not $summary.MailboxesSkipped.Contains($Reason)) { $summary.MailboxesSkipped[$Reason] = 0 }
    $summary.MailboxesSkipped[$Reason]++
    Write-SyncLog -Level SKIP -Message "$Reason : $Detail"
}

function Add-Error {
    param([Parameter(Mandatory)] [string]$Message)
    $summary.Errors++
    $errorMessages.Add($Message)
    Write-SyncLog -Level ERROR -Message $Message
}

$mode = if ($WhatIf) { 'WHATIF' } else { 'APPLY' }
$runStarted = [DateTime]::UtcNow
Write-SyncLog "Sync-DirectoryToMailboxes starting. Mode=$mode Group='$GroupName' ExclusionGroup='$ExclusionGroupName' UserPrincipalName='$UserPrincipalName' FullRebuild=$FullRebuild DetectLocalEdits=$DetectLocalEdits"

# --- connect -------------------------------------------------------------------------------------------------------
foreach ($m in $RequiredModules) { Import-Module -Name $m -ErrorAction Stop }
Connect-MgGraph -Identity -NoWelcome
$ctx = Get-MgContext
if (-not $ctx) { throw 'Connect-MgGraph -Identity did not produce a Graph context. Is the system-assigned managed identity enabled?' }
Write-SyncLog "Connected to Graph as managed identity (client $($ctx.ClientId)) in tenant $($ctx.TenantId)"
# Guard: a Contacts.* permission granted in Entra is tenant-wide and silently overrides the Exchange RBAC mailbox scope
# (Entra and Exchange grants are a union). Refuse to run rather than write with unscoped access.
$appRoles = @(Get-Field -Object $ctx -Name 'Scopes') | Where-Object { $_ }
if (@($appRoles).Count -eq 0) {
    Write-SyncLog -Level WARN -Message 'Could not read the identity''s application permissions from the Graph context; the unscoped-Contacts guard did not run.'
}
else {
    Write-SyncLog "Identity application permissions: $(@($appRoles | Sort-Object) -join ', ')"
    $unscoped = @($appRoles | Where-Object { $_ -match '^Contacts\.' })
    if ($unscoped.Count -gt 0) {
        throw "The managed identity holds $($unscoped -join ', ') in Entra. That grant is tenant-wide and defeats the Exchange mailbox scope, so this runbook refuses to run. Remove it in Entra (Enterprise applications > the identity > Permissions) and re-run."
    }
}

# --- 1. members ------------------------------------------------------------------------------------------------------
$groupId = Get-GroupIdByName -DisplayName $GroupName
if (-not $groupId) { throw "Control group '$GroupName' was not found." }
$select = 'id,userPrincipalName,accountEnabled,displayName,givenName,surname,jobTitle,department,companyName,mobilePhone,businessPhones,mail,officeLocation'
# The OData cast is an advanced query: ConsistencyLevel eventual and $count=true, as Microsoft Learn documents.
$rawMembers = @(Get-GraphCollection -Uri "$GraphBase/groups/$groupId/transitiveMembers/microsoft.graph.user?`$select=$select&`$count=true&`$top=999" -Headers @{ ConsistencyLevel = 'eventual' })
$members = @($rawMembers | ForEach-Object { ConvertTo-DirectoryUserRecord -User $_ })
Write-SyncLog "Control group '$GroupName' has $($members.Count) user members"

$excludedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
if (-not [string]::IsNullOrWhiteSpace($ExclusionGroupName)) {
    $exclusionGroupId = Get-GroupIdByName -DisplayName $ExclusionGroupName
    if (-not $exclusionGroupId) { throw "Exclusion group '$ExclusionGroupName' was not found. Create it or pass an empty -ExclusionGroupName." }
    $exMembers = @(Get-GraphCollection -Uri "$GraphBase/groups/$exclusionGroupId/transitiveMembers/microsoft.graph.user?`$select=id&`$count=true&`$top=999" -Headers @{ ConsistencyLevel = 'eventual' })
    foreach ($m in $exMembers) { $id = Get-Field -Object $m -Name 'id'; if ($id) { $null = $excludedIds.Add([string]$id) } }
    Write-SyncLog "Exclusion group '$ExclusionGroupName' has $($excludedIds.Count) user members (they receive contacts but are not listed)"
}

$directory = @($members | Where-Object { $_.AccountEnabled -and -not $excludedIds.Contains($_.Id) } | Sort-Object -Property UserPrincipalName)
$summary.DirectorySize = $directory.Count
Write-SyncLog "Directory (enabled members minus exclusions): $($directory.Count) people"

$targets = @($members | Where-Object { $_.AccountEnabled } | Sort-Object -Property UserPrincipalName)
foreach ($m in ($members | Where-Object { -not $_.AccountEnabled })) { Add-Skip -Reason 'AccountDisabled' -Detail $m.UserPrincipalName }
if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
    $targets = @($targets | Where-Object { $_.UserPrincipalName -ieq $UserPrincipalName.Trim() })
    if ($targets.Count -ne 1) { throw "UserPrincipalName '$UserPrincipalName' is not an enabled member of '$GroupName'. The user must carry the stamp before their mailbox can be filled." }
}
$summary.MailboxesTargeted = $targets.Count
Write-SyncLog "Target mailboxes: $($targets.Count)"

# --- 2..4. per mailbox ------------------------------------------------------------------------------------------------
$targetIndex = 0
foreach ($target in $targets) {
    $targetIndex++
    $upn = $target.UserPrincipalName
    if (([DateTime]::UtcNow - $runStarted).TotalMinutes -ge $TimeBudgetMinutes) {
        $summary.MailboxesNotReached = $targets.Count - $targetIndex + 1
        Add-Error "Time budget of $TimeBudgetMinutes minutes used up before the Automation 3-hour limit: $($summary.MailboxesNotReached) mailbox(es) not processed, starting with $upn. Re-run; finished mailboxes are unchanged and go fast."
        break
    }
    try {
        $owned = @(Get-OwnedContacts -UserId $target.Id)
    }
    catch {
        $status = Get-HttpStatusCode -ErrorRecord $_
        if ($status -eq 403) { Add-Skip -Reason 'AccessDenied' -Detail "$upn : 403 from Exchange. The RBAC scope has not caught up with a fresh stamp (30 min to 2 h) or the user is not stamped. Re-run later."; continue }
        if ($status -eq 404) { Add-Skip -Reason 'NoMailbox' -Detail "$upn : 404, no Exchange Online mailbox or REST not enabled for it."; continue }
        Add-Error "Listing contacts failed for $upn : $($_.Exception.Message)"
        continue
    }

    $plan = Resolve-MailboxChanges -Directory $directory -TargetUserId $target.Id -OwnedContacts $owned -FullRebuild $FullRebuild -DetectLocalEdits $DetectLocalEdits
    $summary.ContactsUnchanged += $plan.Unchanged
    Write-SyncLog "$upn : owned=$($owned.Count) create=$($plan.Creates.Count) update=$($plan.Updates.Count) delete=$($plan.Deletes.Count) unchanged=$($plan.Unchanged)"

    $requests = [System.Collections.Generic.List[object]]::new()
    $actionById = @{}
    $n = 0
    foreach ($c in $plan.Creates) {
        $n++; $id = "c$n"
        $requests.Add(@{ id = $id; method = 'POST'; url = "/users/$($target.Id)/contacts"; body = (New-DirectoryContactBody -Fields $c.Fields -SourceId $c.User.Id -Hash $c.Hash) })
        $actionById[$id] = @{ Action = 'CREATE'; Label = "$($c.Fields.fileAs) <$($c.User.UserPrincipalName)>" }
    }
    foreach ($u in $plan.Updates) {
        $n++; $id = "u$n"
        $requests.Add(@{ id = $id; method = 'PATCH'; url = "/users/$($target.Id)/contacts/$($u.ContactId)"; body = (New-DirectoryContactBody -Fields $u.Fields -SourceId $u.User.Id -Hash $u.Hash) })
        $actionById[$id] = @{ Action = 'UPDATE'; Label = "$($u.Fields.fileAs) <$($u.User.UserPrincipalName)> ($($u.Reason))"; LocalEdit = ($u.Reason -eq 'edited in mailbox') }
    }
    foreach ($d in $plan.Deletes) {
        $n++; $id = "d$n"
        $requests.Add(@{ id = $id; method = 'DELETE'; url = "/users/$($target.Id)/contacts/$($d.Contact.ContactId)"; body = $null })
        $actionById[$id] = @{ Action = 'DELETE'; Label = "'$($d.Contact.DisplayName)' source=$($d.Contact.SourceId) ($($d.Reason))" }
    }

    if ($WhatIf) {
        foreach ($r in $requests) {
            Write-SyncLog -Level CHANGE -Message "WHATIF $($actionById[$r.id].Action) mailbox=$upn $($actionById[$r.id].Label)"
        }
        $summary.ContactsAdded += $plan.Creates.Count
        $summary.ContactsUpdated += $plan.Updates.Count
        $summary.LocalEditsReverted += @($plan.Updates | Where-Object { $_.Reason -eq 'edited in mailbox' }).Count
        $summary.ContactsDeleted += $plan.Deletes.Count
        $summary.MailboxesProcessed++
        continue
    }

    if ($requests.Count -gt 0) {
        try {
            $results = Invoke-GraphBatch -Requests $requests.ToArray()
        }
        catch {
            # A $batch call that fails as a whole (network, 5xx after retries) costs this mailbox only; the rest still run.
            $status = Get-HttpStatusCode -ErrorRecord $_
            if ($status -eq 403) { Add-Skip -Reason 'AccessDenied' -Detail "$upn : 403 on write. The Exchange RBAC scope has not caught up with a fresh stamp (30 min to 2 h). Re-run later."; continue }
            Add-Error "Writing contacts failed for $upn (HTTP $status): $($_.Exception.Message)"
            continue
        }
        $mailboxForbidden = $false
        $createsDone = 0
        foreach ($r in $requests) {
            $id = $r.id
            $action = $actionById[$id]
            $res = $results[$id]
            $status = if ($null -ne $res) { $res.Status } else { 0 }
            if ($status -ge 200 -and $status -lt 300) {
                switch ($action.Action) {
                    'CREATE' { $summary.ContactsAdded++; $createsDone++ }
                    'UPDATE' { $summary.ContactsUpdated++; if ($action.Contains('LocalEdit') -and $action.LocalEdit) { $summary.LocalEditsReverted++ } }
                    'DELETE' { $summary.ContactsDeleted++ }
                }
                Write-SyncLog -Level CHANGE -Message "APPLY $($action.Action) mailbox=$upn $($action.Label)"
            }
            elseif ($status -eq 403) { $mailboxForbidden = $true }
            elseif ($status -eq 404 -and $action.Action -eq 'UPDATE') {
                # The user deleted the card between listing and writing. Not an error: the next run re-creates it.
                Write-SyncLog -Level WARN -Message "UPDATE skipped in mailbox $upn for $($action.Label): the card was deleted in the meantime and is re-created on the next run"
            }
            elseif ($status -eq 404 -and $action.Action -eq 'DELETE') {
                # Already gone (user deleted it by hand): the desired state holds.
                $summary.ContactsDeleted++
                Write-SyncLog -Level CHANGE -Message "APPLY DELETE mailbox=$upn $($action.Label) (was already gone)"
            }
            else {
                Add-Error "$($action.Action) failed in mailbox $upn for $($action.Label): HTTP $status $(Get-GraphErrorText -Body $(if ($null -ne $res) { $res.Body } else { $null }))"
            }
        }
        if ($mailboxForbidden) {
            Add-Skip -Reason 'AccessDenied' -Detail "$upn : 403 on write. The Exchange RBAC scope has not caught up with a fresh stamp (30 min to 2 h). Re-run later."
            continue
        }
        if ($createsDone -gt 0) {
            # Duplicate sweep. A retried $batch call (timeout or 5xx after Exchange had already committed some creates), or a
            # second run overlapping this one, can leave two cards for one coworker. Re-list and remove the extras now
            # instead of leaving them on the phone until next week.
            try {
                $after = @(Get-OwnedContacts -UserId $target.Id)
                $dupes = @((Resolve-MailboxChanges -Directory @() -TargetUserId $target.Id -OwnedContacts $after).Deletes | Where-Object { $_.Reason -eq 'duplicate' })
                if ($dupes.Count -gt 0) {
                    $n = 0
                    $sweep = foreach ($d in $dupes) { $n++; @{ id = "x$n"; method = 'DELETE'; url = "/users/$($target.Id)/contacts/$($d.Contact.ContactId)"; body = $null } }
                    $sweepResults = Invoke-GraphBatch -Requests @($sweep)
                    $removed = @($sweepResults.Values | Where-Object { ($_.Status -ge 200 -and $_.Status -lt 300) -or $_.Status -eq 404 }).Count
                    $summary.DuplicatesRemoved += $removed
                    Write-SyncLog -Level CHANGE -Message "APPLY DELETE mailbox=$upn removed $removed duplicate card(s) left by a retried or overlapping write"
                }
            }
            catch {
                Write-SyncLog -Level WARN -Message "Duplicate check after writing failed for $upn ($($_.Exception.Message)); the next run removes any duplicates."
            }
        }
    }
    $summary.MailboxesProcessed++
}

# --- 5. summary ----------------------------------------------------------------------------------------------------------
if ($summary.MailboxesTargeted -gt 0 -and $summary.MailboxesProcessed -eq 0 -and $summary.MailboxesNotReached -eq 0) {
    $accessDenied = if ($summary.MailboxesSkipped.Contains('AccessDenied')) { $summary.MailboxesSkipped['AccessDenied'] } else { 0 }
    if ($accessDenied -gt 0) {
        Add-Error "None of the $($summary.MailboxesTargeted) targeted mailboxes could be processed ($accessDenied returned 403). Scope lag affects fresh stamps only, so this means the Exchange RBAC assignment is missing or broken. Check it with deploy/New-ExchangeAppScope.ps1 -TestOnly."
    }
    elseif ($summary.Errors -eq 0) {
        Add-Error "None of the $($summary.MailboxesTargeted) targeted mailboxes could be processed. See the SKIP lines above."
    }
}
Write-SyncLog "Sync-DirectoryToMailboxes finished. Directory=$($summary.DirectorySize) Targeted=$($summary.MailboxesTargeted) Processed=$($summary.MailboxesProcessed) Added=$($summary.ContactsAdded) Updated=$($summary.ContactsUpdated) Deleted=$($summary.ContactsDeleted) Unchanged=$($summary.ContactsUnchanged) LocalEditsReverted=$($summary.LocalEditsReverted) DuplicatesRemoved=$($summary.DuplicatesRemoved) Skipped=$(($summary.MailboxesSkipped.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ',') Errors=$($summary.Errors)"
[pscustomobject]$summary

if ($summary.Errors -gt 0) {
    throw "Sync-DirectoryToMailboxes completed with $($summary.Errors) error(s): $($errorMessages -join ' | ')"
}
#endregion
