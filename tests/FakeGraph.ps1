<#
.SYNOPSIS
    In-memory stand-in for the Microsoft Graph SDK cmdlets and Invoke-MgGraphRequest used by the runbooks.

.DESCRIPTION
    Dot-source from a Pester BeforeAll. It defines functions with the same names as the SDK cmdlets the runbooks call, so when a
    runbook is invoked with & from the test scope, command resolution finds these functions first. Nothing here talks to a network.

    Supported surface (exactly what the runbooks use):
      Import-Module, Connect-MgGraph -Identity, Get-MgContext, Get-MgGroup -Filter, Get-MgUser -Filter -Property,
      Get-MgDeviceManagementManagedDevice -Filter -Property,
      Invoke-MgGraphRequest GET  /groups/{id}/transitiveMembers/microsoft.graph.user
                            GET  /roleManagement/directory/roleAssignments
                            GET  /users/{id}/contacts (paged with $skip, 403/404 simulation)
                            PATCH /users/{id}
                            POST /$batch (contacts POST/PATCH/DELETE, optional one-time 429)
#>

function New-FakeTenant {
    $global:FakeTenant = @{
        Users              = [System.Collections.Generic.List[object]]::new()   # camelCase hashtables like Graph JSON
        Devices            = [System.Collections.Generic.List[object]]::new()
        Groups             = @{}                                               # displayName -> @{ id; members = List[string] }
        RoleAssignments    = [System.Collections.Generic.List[object]]::new()   # @{ principalId }
        Mailboxes          = @{}                                               # userId -> List[hashtable contact]
        ForbiddenMailboxes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        NoMailbox          = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        Patches            = [System.Collections.Generic.List[object]]::new()   # @{ UserId; Body }
        BatchCalls         = 0
        ContactPageSize    = 2
        ThrottleOnce       = [System.Collections.Generic.HashSet[string]]::new()  # "METHOD url" to answer 429 exactly once
        FailPatchUsers     = @{}                                               # userId -> HTTP status to throw on PATCH /users/{id}
        FailBatchUrls      = @{}                                               # "METHOD url" -> status returned inside $batch (every time)
        BatchThrow         = 0                                                 # when > 0, the next $batch call throws this HTTP status once
        LastHeaders        = @{}                                               # uri -> headers seen (for assertions)
        NextContactId      = 1
    }
    return $global:FakeTenant
}

function Add-FakeUser {
    param([string]$Id, [string]$Upn, [bool]$Enabled = $true, [string]$Ext1 = $null, [string]$Mobile = $null, [hashtable]$Extra = @{})
    $u = @{
        id                            = $Id
        userPrincipalName             = $Upn
        accountEnabled                = $Enabled
        mobilePhone                   = $Mobile
        displayName                   = $Upn.Split('@')[0]
        givenName                     = ($Upn.Split('@')[0].Split('.')[0])
        surname                       = $(if ($Upn.Split('@')[0].Contains('.')) { $Upn.Split('@')[0].Split('.')[1] } else { $null })
        jobTitle                      = 'Staff'
        department                    = 'Dept'
        companyName                   = 'Contoso'
        businessPhones                = @()
        mail                          = $Upn
        officeLocation                = 'HQ'
        onPremisesExtensionAttributes = @{ extensionAttribute1 = $Ext1; extensionAttribute2 = $null }
    }
    foreach ($k in $Extra.Keys) { $u[$k] = $Extra[$k] }
    $global:FakeTenant.Users.Add($u)
    if (-not $global:FakeTenant.Mailboxes.ContainsKey($Id)) { $global:FakeTenant.Mailboxes[$Id] = [System.Collections.Generic.List[object]]::new() }
    return $u
}

function Add-FakeDevice {
    param([string]$UserId, [string]$Upn, [string]$Phone, [string]$Sync = '2026-10-30T12:00:00Z', [string]$Serial = 'SER', [string]$Name = 'iPhone', [string]$Os = 'iOS', [string]$Owner = 'company', [string]$Model = 'iPhone 15')
    $global:FakeTenant.Devices.Add(@{ id = [guid]::NewGuid().ToString(); userId = $UserId; userPrincipalName = $Upn; phoneNumber = $Phone; lastSyncDateTime = $Sync; serialNumber = $Serial; deviceName = $Name; operatingSystem = $Os; managedDeviceOwnerType = $Owner; model = $Model })
}

function Add-FakeGroup {
    param([string]$Name, [string]$Id, [string[]]$Members = @())
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($m in $Members) { $list.Add($m) }
    $global:FakeTenant.Groups[$Name] = @{ id = $Id; members = $list }
}

function Add-FakeContact {
    <# Adds a contact to a mailbox. With -SourceId it is tagged as runbook-owned. #>
    param([string]$MailboxUserId, [string]$DisplayName, [string]$SourceId, [string]$Hash = 'stale', [hashtable]$Fields = @{})
    $c = @{ id = "c$($global:FakeTenant.NextContactId)"; displayName = $DisplayName; singleValueExtendedProperties = @() }
    $global:FakeTenant.NextContactId++
    foreach ($k in $Fields.Keys) { $c[$k] = $Fields[$k] }
    if ($SourceId) {
        $c.singleValueExtendedProperties = @(
            @{ id = 'String {b4b9c0a1-5f0e-4c3d-9a7e-2d6f1c8e0a11} Name DirectorySourceId'; value = $SourceId },
            @{ id = 'String {b4b9c0a1-5f0e-4c3d-9a7e-2d6f1c8e0a11} Name DirectorySourceHash'; value = $Hash }
        )
    }
    $global:FakeTenant.Mailboxes[$MailboxUserId].Add($c)
    return $c
}

function New-FakeHttpError {
    <#
        Alternates between the two exception shapes the runbooks must understand: HttpRequestException with .StatusCode
        (what the SDK surfaces on some paths) and HttpResponseException with .Response.StatusCode (what Invoke-MgGraphRequest
        throws on non-success responses).
    #>
    param([int]$Status, [string]$Message = 'simulated')
    if ($Status % 2 -eq 0) {
        return [System.Net.Http.HttpRequestException]::new("Response status code does not indicate success: $Status ($Message).", $null, [System.Net.HttpStatusCode]$Status)
    }
    $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Status)
    return [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Status ($Message).", $response)
}

function ConvertTo-PascalObject {
    <# Graph SDK cmdlets return typed objects with PascalCase properties; fake that from a camelCase hashtable. #>
    param([hashtable]$Table)
    $o = [ordered]@{}
    foreach ($k in $Table.Keys) {
        $pascal = $k.Substring(0, 1).ToUpperInvariant() + $k.Substring(1)
        $v = $Table[$k]
        if ($v -is [hashtable]) { $v = ConvertTo-PascalObject -Table $v }
        $o[$pascal] = $v
    }
    return [pscustomobject]$o
}

# ------------------------------------------------------------------ SDK stubs -------------------------------------------------------------
function Import-Module { param($Name) }
function Connect-MgGraph { param([switch]$Identity, [switch]$NoWelcome, [string[]]$Scopes) }
function Get-MgContext { [pscustomobject]@{ ClientId = 'fake-client'; TenantId = 'fake-tenant' } }

function Get-MgGroup {
    param([string]$Filter, [string[]]$Property, [int]$Top)
    if ($Filter -notmatch "displayName eq '(.+)'") { throw "FakeGraph: unsupported group filter '$Filter'" }
    $name = $Matches[1].Replace("''", "'")
    if ($global:FakeTenant.Groups.ContainsKey($name)) {
        return [pscustomobject]@{ Id = $global:FakeTenant.Groups[$name].id; DisplayName = $name }
    }
}

function Get-MgUser {
    param([switch]$All, [int]$PageSize, [string]$ConsistencyLevel, [string]$CountVariable, [string]$Filter, [string[]]$Property)
    if ($Filter) { throw "FakeGraph: the runbook must not filter users server-side (got '$Filter'); it reads all users and drops guests itself" }
    foreach ($u in $global:FakeTenant.Users) { ConvertTo-PascalObject -Table $u }
}

function Get-MgDeviceManagementManagedDevice {
    param([switch]$All, [int]$PageSize, [string]$Filter, [string[]]$Property)
    foreach ($d in $global:FakeTenant.Devices) {
        if ($d.operatingSystem -eq 'iOS' -and $d.managedDeviceOwnerType -eq 'company') { ConvertTo-PascalObject -Table $d }
    }
}

function Invoke-MgGraphRequest {
    param([string]$Method, [string]$Uri, $Body, [string]$ContentType, [string]$OutputType, [hashtable]$Headers)
    $t = $global:FakeTenant
    $path = ([uri]$Uri).AbsolutePath -replace '^/v1\.0', ''
    $query = [System.Web.HttpUtility]::ParseQueryString(([uri]$Uri).Query)
    if ($null -ne $Headers) { $t.LastHeaders[$path] = $Headers }

    # GET /groups/{id}/transitiveMembers/microsoft.graph.user  (advanced query: requires ConsistencyLevel eventual + $count)
    if ($Method -eq 'GET' -and $path -match '^/groups/([^/]+)/transitiveMembers/microsoft\.graph\.user$') {
        $gid = $Matches[1]
        if ($null -eq $Headers -or [string]$Headers['ConsistencyLevel'] -ne 'eventual' -or $query['$count'] -ne 'true') {
            throw (New-FakeHttpError -Status 400 -Message 'Request_UnsupportedQuery: OData cast requires ConsistencyLevel eventual and $count=true')
        }
        $group = $t.Groups.Values | Where-Object { $_.id -eq $gid } | Select-Object -First 1
        if (-not $group) { throw (New-FakeHttpError -Status 404 -Message 'group not found') }
        $members = @($t.Users | Where-Object { $group.members -contains $_.id } | ForEach-Object { $_.Clone() })
        return @{ value = $members }
    }
    # GET /roleManagement/directory/roleAssignments
    if ($Method -eq 'GET' -and $path -eq '/roleManagement/directory/roleAssignments') {
        return @{ value = @($t.RoleAssignments) }
    }
    # GET /users/{id}/contacts  (paged)
    if ($Method -eq 'GET' -and $path -match '^/users/([^/]+)/contacts$') {
        $uid = $Matches[1]
        if ($t.ForbiddenMailboxes.Contains($uid)) { throw (New-FakeHttpError -Status 403 -Message 'Forbidden') }
        if ($t.NoMailbox.Contains($uid) -or -not $t.Mailboxes.ContainsKey($uid)) { throw (New-FakeHttpError -Status 404 -Message 'MailboxNotEnabledForRESTAPI') }
        if ($query['$expand'] -notmatch 'singleValueExtendedProperties') { throw 'FakeGraph: contacts listed without $expand of extended properties' }
        $skip = 0
        if ($query['$skip']) { $skip = [int]$query['$skip'] }
        $all = @($t.Mailboxes[$uid])
        $page = @()
        if ($skip -lt $all.Count) { $page = @($all[$skip..([Math]::Min($skip + $t.ContactPageSize, $all.Count) - 1)]) }
        $projected = @($page | ForEach-Object {
                $c = $_
                $out = @{ id = $c.id; displayName = $c.displayName }
                foreach ($f in @(([string]$query['$select']).Split(',') | Where-Object { $_ -and $_ -ne 'id' })) {
                    # Like Graph: a selected property is always present, null when the contact has no value.
                    $out[$f] = if ($c.ContainsKey($f)) { $c[$f] } else { $null }
                }
                $props = @($c.singleValueExtendedProperties | Where-Object { $_.id -match 'DirectorySourceId|DirectorySourceHash' })
                if ($props.Count -gt 0) { $out.singleValueExtendedProperties = $props }
                $out
            })
        $result = @{ value = $projected }
        if ($skip + $t.ContactPageSize -lt $all.Count) {
            $result['@odata.nextLink'] = "https://graph.microsoft.com/v1.0$path`?`$select=$($query['$select'])&`$expand=singleValueExtendedProperties(x)&`$skip=$($skip + $t.ContactPageSize)"
        }
        return $result
    }
    # PATCH /users/{id}
    if ($Method -eq 'PATCH' -and $path -match '^/users/([^/]+)$') {
        $uid = $Matches[1]
        $u = $t.Users | Where-Object { $_.id -eq $uid } | Select-Object -First 1
        if (-not $u) { throw (New-FakeHttpError -Status 404 -Message 'user not found') }
        if ($t.FailPatchUsers.ContainsKey($uid)) { throw (New-FakeHttpError -Status ([int]$t.FailPatchUsers[$uid]) -Message 'Authorization_RequestDenied') }
        $b = $Body | ConvertFrom-Json -AsHashtable
        $t.Patches.Add(@{ UserId = $uid; Body = $Body })
        if ($b.ContainsKey('onPremisesExtensionAttributes')) {
            foreach ($k in $b.onPremisesExtensionAttributes.Keys) { $u.onPremisesExtensionAttributes[$k] = $b.onPremisesExtensionAttributes[$k] }
        }
        if ($b.ContainsKey('mobilePhone')) { $u.mobilePhone = $b.mobilePhone }
        return $null
    }
    # POST /$batch
    if ($Method -eq 'POST' -and $path -eq '/$batch') {
        $t.BatchCalls++
        if ($t.BatchThrow -gt 0) { $st = $t.BatchThrow; $t.BatchThrow = 0; throw (New-FakeHttpError -Status $st -Message 'batch endpoint failure') }
        $b = $Body | ConvertFrom-Json -AsHashtable
        if ($b.requests.Count -gt 20) { throw 'FakeGraph: batch larger than 20 requests' }
        $responses = foreach ($r in $b.requests) {
            $key = "$($r.method) $($r.url)"
            if ($t.FailBatchUrls.ContainsKey($key)) { @{ id = $r.id; status = [int]$t.FailBatchUrls[$key]; body = @{ error = @{ code = 'Simulated'; message = "forced $($t.FailBatchUrls[$key])" } } }; continue }
            if ($t.ThrottleOnce.Contains($key)) {
                $null = $t.ThrottleOnce.Remove($key)
                @{ id = $r.id; status = 429; headers = @{ 'Retry-After' = '0' } }
                continue
            }
            if ($r.url -notmatch '^/users/([^/]+)/contacts(?:/([^/?]+))?$') { @{ id = $r.id; status = 400; body = @{ error = @{ code = 'BadRequest'; message = "unsupported url $($r.url)" } } }; continue }
            $uid = $Matches[1]
            $cid = if ($Matches.Count -gt 2) { $Matches[2] } else { $null }
            if ($t.ForbiddenMailboxes.Contains($uid)) { @{ id = $r.id; status = 403; body = @{ error = @{ code = 'ErrorAccessDenied'; message = 'Access is denied' } } }; continue }
            $box = $t.Mailboxes[$uid]
            switch ($r.method) {
                'POST' {
                    $new = @{ id = "c$($t.NextContactId)" }
                    $t.NextContactId++
                    foreach ($k in $r.body.Keys) { $new[$k] = $r.body[$k] }
                    $box.Add($new)
                    @{ id = $r.id; status = 201; body = @{ id = $new.id } }
                }
                'PATCH' {
                    $existing = $box | Where-Object { $_.id -eq $cid } | Select-Object -First 1
                    if (-not $existing) { @{ id = $r.id; status = 404; body = @{ error = @{ code = 'ErrorItemNotFound'; message = 'not found' } } } }
                    else {
                        foreach ($k in $r.body.Keys) { $existing[$k] = $r.body[$k] }
                        @{ id = $r.id; status = 200; body = @{ id = $cid } }
                    }
                }
                'DELETE' {
                    $existing = $box | Where-Object { $_.id -eq $cid } | Select-Object -First 1
                    if (-not $existing) { @{ id = $r.id; status = 404; body = @{ error = @{ code = 'ErrorItemNotFound'; message = 'not found' } } } }
                    else { $null = $box.Remove($existing); @{ id = $r.id; status = 204 } }
                }
            }
        }
        return @{ responses = @($responses) }
    }
    throw "FakeGraph: unsupported request $Method $Uri"
}

function Start-Sleep { param($Seconds, $Milliseconds) }

function Get-OwnedFakeContacts {
    param([string]$MailboxUserId)
    @($global:FakeTenant.Mailboxes[$MailboxUserId] | Where-Object { @($_.singleValueExtendedProperties | Where-Object { $_.id -match 'DirectorySourceId' }).Count -gt 0 })
}
