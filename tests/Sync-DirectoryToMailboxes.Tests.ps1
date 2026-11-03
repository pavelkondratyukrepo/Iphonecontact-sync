#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Local Pester tests for runbooks/Sync-DirectoryToMailboxes.ps1. They never touch Graph or a tenant: the runbook's functions are
    extracted with the AST and exercised with sample data, and the batch sender is tested against a mocked Graph.
    Run from the repo root:  Invoke-Pester ./tests
#>

BeforeAll {
    Set-StrictMode -Version Latest
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $runbook = Join-Path (Get-RepoRoot) 'runbooks' 'Sync-DirectoryToMailboxes.ps1'
    . ([scriptblock]::Create((Get-RunbookFunctionText -Path $runbook)))

    # Script-scope constants the functions read
    $GraphBase = 'https://graph.microsoft.com/v1.0'
    $BatchSize = 20

    function New-SampleUser {
        param([hashtable]$Overrides = @{})
        $base = @{
            id                = '11111111-1111-1111-1111-111111111111'
            userPrincipalName = 'jane.doe@contoso.com'
            accountEnabled    = $true
            displayName       = 'Jane Doe'
            givenName         = 'Jane'
            surname           = 'Doe'
            jobTitle          = 'Field Engineer'
            department        = 'Operations'
            companyName       = 'Contoso'
            mobilePhone       = '+18045550123'
            businessPhones    = @('+18045550100')
            mail              = 'jane.doe@contoso.com'
            officeLocation    = 'Richmond 3F'
        }
        foreach ($k in $Overrides.Keys) { $base[$k] = $Overrides[$k] }
        return ConvertTo-DirectoryUserRecord -User $base
    }

    # The three sample users required by the design
    $script:UserAllFields = New-SampleUser
    $script:UserNoMobile = New-SampleUser -Overrides @{
        id = '22222222-2222-2222-2222-222222222222'; userPrincipalName = 'john.smith@contoso.com'; displayName = 'John Smith'
        givenName = 'John'; surname = 'Smith'; jobTitle = 'Controller'; department = 'Finance'; mobilePhone = $null
        businessPhones = @('+18045550101', '+18045550102'); mail = 'john.smith@contoso.com'; officeLocation = $null
    }
    $script:UserCommaApostrophe = New-SampleUser -Overrides @{
        id = '33333333-3333-3333-3333-333333333333'; userPrincipalName = 'siobhan.obrien@contoso.com'
        displayName = "O'Brien, Siobhán (Contractor)"; givenName = 'Siobhán'; surname = "O'Brien, Jr."
        jobTitle = "VP, Sales & Marketing"; department = 'Sales'; mobilePhone = '+18045550199'; businessPhones = @()
        mail = "siobhan.o'brien@contoso.com"; officeLocation = 'Remote'
    }
}

Describe 'Ownership constants' {
    It 'defines the two extended property ids under one fixed GUID namespace' {
        $c = Get-DirectoryContactConstants
        $c.NamespaceGuid | Should -Match '^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$'
        $c.SourceIdPropertyId | Should -Be "String $($c.NamespaceGuid) Name DirectorySourceId"
        $c.SourceHashPropertyId | Should -Be "String $($c.NamespaceGuid) Name DirectorySourceHash"
        $c.Category | Should -Be 'Company Directory'
    }
}

Describe 'ConvertTo-DirectoryContactFields' {
    It 'maps every field for a user with all fields' {
        $f = ConvertTo-DirectoryContactFields -User $UserAllFields
        $f.givenName | Should -Be 'Jane'
        $f.surname | Should -Be 'Doe'
        $f.displayName | Should -Be 'Jane Doe'
        $f.fileAs | Should -Be 'Doe, Jane'
        $f.jobTitle | Should -Be 'Field Engineer'
        $f.department | Should -Be 'Operations'
        $f.companyName | Should -Be 'Contoso'
        $f.mobilePhone | Should -Be '+18045550123'
        @($f.businessPhones) | Should -Be @('+18045550100')
        @($f.emailAddresses).Count | Should -Be 1
        $f.emailAddresses[0].address | Should -Be 'jane.doe@contoso.com'
        $f.emailAddresses[0].name | Should -Be 'Jane Doe'
        $f.officeLocation | Should -Be 'Richmond 3F'
    }

    It 'leaves mobilePhone and officeLocation null for a user without them and keeps both work phones' {
        $f = ConvertTo-DirectoryContactFields -User $UserNoMobile
        $f.mobilePhone | Should -BeNullOrEmpty
        $f.officeLocation | Should -BeNullOrEmpty
        @($f.businessPhones).Count | Should -Be 2
        $f.fileAs | Should -Be 'Smith, John'
    }

    It 'handles a name with a comma and an apostrophe without altering it' {
        $f = ConvertTo-DirectoryContactFields -User $UserCommaApostrophe
        $f.displayName | Should -Be "O'Brien, Siobhán (Contractor)"
        $f.surname | Should -Be "O'Brien, Jr."
        $f.fileAs | Should -Be "O'Brien, Jr., Siobhán"
        $f.jobTitle | Should -Be 'VP, Sales & Marketing'
        @($f.businessPhones).Count | Should -Be 0
        $f.emailAddresses[0].address | Should -Be "siobhan.o'brien@contoso.com"
    }

    It 'falls back for fileAs and displayName when name parts are missing' {
        $u = New-SampleUser -Overrides @{ displayName = $null; givenName = $null; surname = 'Solo' }
        $f = ConvertTo-DirectoryContactFields -User $u
        $f.displayName | Should -Be 'Solo'
        $f.fileAs | Should -Be 'Solo'
        $u2 = New-SampleUser -Overrides @{ displayName = $null; givenName = $null; surname = $null }
        $f2 = ConvertTo-DirectoryContactFields -User $u2
        $f2.displayName | Should -Be 'jane.doe@contoso.com'
        $f2.fileAs | Should -Be 'jane.doe@contoso.com'
    }

    It 'produces the exact set of Graph contact properties in a fixed order' {
        $f = ConvertTo-DirectoryContactFields -User $UserAllFields
        @($f.Keys) | Should -Be @('givenName', 'surname', 'displayName', 'fileAs', 'jobTitle', 'department', 'companyName', 'mobilePhone', 'businessPhones', 'emailAddresses', 'officeLocation')
    }
}

Describe 'Get-DirectoryContactHash' {
    It 'is a lowercase 64-character SHA256 hex string' {
        $h = Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $UserAllFields)
        $h | Should -Match '^[0-9a-f]{64}$'
    }

    It 'is stable across runs and across freshly built records for all three sample users' {
        foreach ($u in @($UserAllFields, $UserNoMobile, $UserCommaApostrophe)) {
            $h1 = Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $u)
            $h2 = Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $u)
            $h1 | Should -Be $h2
        }
        # Rebuilt from scratch, same input, same hash
        $again = New-SampleUser
        (Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $again)) |
            Should -Be (Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $UserAllFields))
    }

    It 'matches the pinned reference value so a canonicalization change is caught' {
        # If this fails on purpose (you changed the canonical form), every contact will be rewritten once on the next run.
        $h = Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $UserAllFields)
        $h | Should -Be '1dfbac46786a5a247176cbbe80dad720d5c12591b2930be80cd4b03e6d1428cb'
    }

    It 'differs for the three sample users' {
        $hashes = @($UserAllFields, $UserNoMobile, $UserCommaApostrophe) | ForEach-Object { Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $_) }
        ($hashes | Select-Object -Unique).Count | Should -Be 3
    }

    It 'changes when any written field changes' {
        $base = Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $UserAllFields)
        foreach ($change in @(
                @{ jobTitle = 'Senior Field Engineer' },
                @{ mobilePhone = '+18045550124' },
                @{ businessPhones = @() },
                @{ mail = 'jdoe@contoso.com' },
                @{ officeLocation = 'Richmond 4F' },
                @{ department = 'Ops' },
                @{ companyName = 'Fabrikam' },
                @{ surname = 'Doe-Ray' }
            )) {
            $u = New-SampleUser -Overrides $change
            (Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $u)) | Should -Not -Be $base
        }
    }

    It 'treats null and empty the same so whitespace-only Entra values do not churn contacts' {
        $a = New-SampleUser -Overrides @{ officeLocation = $null }
        $b = New-SampleUser -Overrides @{ officeLocation = '   ' }
        (Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $a)) |
            Should -Be (Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $b))
    }
}

Describe 'New-DirectoryContactBody' {
    BeforeAll {
        $fields = ConvertTo-DirectoryContactFields -User $UserNoMobile
        $hash = Get-DirectoryContactHash -Fields $fields
        $body = New-DirectoryContactBody -Fields $fields -SourceId $UserNoMobile.Id -Hash $hash
        $c = Get-DirectoryContactConstants
    }

    It 'carries the Company Directory category and both extended properties' {
        @($body.categories) | Should -Be @('Company Directory')
        @($body.singleValueExtendedProperties).Count | Should -Be 2
        ($body.singleValueExtendedProperties | Where-Object { $_.id -eq $c.SourceIdPropertyId }).value | Should -Be $UserNoMobile.Id
        ($body.singleValueExtendedProperties | Where-Object { $_.id -eq $c.SourceHashPropertyId }).value | Should -Be $hash
    }

    It 'serializes to JSON with explicit nulls and real arrays (so PATCH clears removed values)' {
        $json = $body | ConvertTo-Json -Depth 10 -Compress
        $json | Should -Match '"mobilePhone":null'
        $json | Should -Match '"officeLocation":null'
        $json | Should -Match '"businessPhones":\["\+18045550101","\+18045550102"\]'
        $json | Should -Match '"categories":\["Company Directory"\]'
        $json | Should -Match '"emailAddresses":\[\{'
        # A single business phone must still be an array, not a bare string
        $one = New-DirectoryContactBody -Fields (ConvertTo-DirectoryContactFields -User $UserAllFields) -SourceId 'x' -Hash 'y' | ConvertTo-Json -Depth 10 -Compress
        $one | Should -Match '"businessPhones":\["\+18045550100"\]'
        # And an empty list is an empty array
        $none = New-DirectoryContactBody -Fields (ConvertTo-DirectoryContactFields -User $UserCommaApostrophe) -SourceId 'x' -Hash 'y' | ConvertTo-Json -Depth 10 -Compress
        $none | Should -Match '"businessPhones":\[\]'
    }
}

Describe 'ConvertTo-OwnedContactRecord' {
    It 'returns null for a contact without DirectorySourceId so it is never touched' {
        $personal = @{ id = 'AAMk1'; displayName = 'Mom'; singleValueExtendedProperties = @() }
        ConvertTo-OwnedContactRecord -Contact $personal | Should -BeNullOrEmpty
        $noProps = @{ id = 'AAMk2'; displayName = 'Dentist' }
        ConvertTo-OwnedContactRecord -Contact $noProps | Should -BeNullOrEmpty
    }

    It 'reads source id and hash from an owned contact' {
        $c = Get-DirectoryContactConstants
        $owned = @{
            id = 'AAMk3'; displayName = 'Jane Doe'
            singleValueExtendedProperties = @(
                @{ id = $c.SourceIdPropertyId; value = $UserAllFields.Id },
                @{ id = $c.SourceHashPropertyId; value = 'abc' }
            )
        }
        $r = ConvertTo-OwnedContactRecord -Contact $owned
        $r.ContactId | Should -Be 'AAMk3'
        $r.SourceId | Should -Be $UserAllFields.Id
        $r.Hash | Should -Be 'abc'
    }
}

Describe 'Resolve-MailboxChanges' {
    BeforeAll {
        $directory = @($UserAllFields, $UserNoMobile, $UserCommaApostrophe)
        function Owned {
            param([string]$ContactId, [object]$User, [string]$Hash)
            if (-not $Hash) { $Hash = Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User $User) }
            [pscustomobject]@{ ContactId = $ContactId; DisplayName = $User.DisplayName; SourceId = $User.Id; Hash = $Hash }
        }
    }

    It 'creates everyone except the target on an empty mailbox' {
        $plan = Resolve-MailboxChanges -Directory $directory -TargetUserId $UserAllFields.Id -OwnedContacts @()
        $plan.Creates.Count | Should -Be 2
        ($plan.Creates.User.Id) | Should -Not -Contain $UserAllFields.Id
        $plan.Updates.Count | Should -Be 0
        $plan.Deletes.Count | Should -Be 0
        $plan.Unchanged | Should -Be 0
    }

    It 'leaves current contacts unchanged, updates a stale hash, deletes a leaver and a duplicate' {
        $leaver = New-SampleUser -Overrides @{ id = '99999999-9999-9999-9999-999999999999'; userPrincipalName = 'leaver@contoso.com'; displayName = 'Lee Ver' }
        $owned = @(
            (Owned -ContactId 'c-john' -User $UserNoMobile),                       # current
            (Owned -ContactId 'c-siobhan' -User $UserCommaApostrophe -Hash 'old'), # stale
            (Owned -ContactId 'c-siobhan-dupe' -User $UserCommaApostrophe),        # duplicate for same source id
            (Owned -ContactId 'c-leaver' -User $leaver)                            # no longer in the directory
        )
        $plan = Resolve-MailboxChanges -Directory $directory -TargetUserId $UserAllFields.Id -OwnedContacts $owned
        $plan.Creates.Count | Should -Be 0
        $plan.Unchanged | Should -Be 1
        $plan.Updates.Count | Should -Be 1
        $plan.Updates[0].ContactId | Should -Be 'c-siobhan'
        @($plan.Deletes.Contact.ContactId) | Should -Be @('c-siobhan-dupe', 'c-leaver')
        @($plan.Deletes.Reason) | Should -Be @('duplicate', 'left directory')
    }

    It 'deletes the target''s own card if one was ever written' {
        $owned = @((Owned -ContactId 'c-self' -User $UserAllFields))
        $plan = Resolve-MailboxChanges -Directory $directory -TargetUserId $UserAllFields.Id -OwnedContacts $owned
        $plan.Deletes.Count | Should -Be 1
        $plan.Deletes[0].Contact.ContactId | Should -Be 'c-self'
    }

    It 'rewrites every owned contact with FullRebuild even when hashes match' {
        $owned = @((Owned -ContactId 'c-john' -User $UserNoMobile), (Owned -ContactId 'c-siobhan' -User $UserCommaApostrophe))
        $plan = Resolve-MailboxChanges -Directory $directory -TargetUserId $UserAllFields.Id -OwnedContacts $owned -FullRebuild $true
        $plan.Updates.Count | Should -Be 2
        $plan.Unchanged | Should -Be 0
    }

    It 'deletes every owned contact when the directory is empty (purge path used for rollback)' {
        $owned = @((Owned -ContactId 'c-john' -User $UserNoMobile), (Owned -ContactId 'c-siobhan' -User $UserCommaApostrophe))
        $plan = Resolve-MailboxChanges -Directory @() -TargetUserId $UserAllFields.Id -OwnedContacts $owned
        $plan.Deletes.Count | Should -Be 2
        $plan.Creates.Count | Should -Be 0
    }
}

Describe 'Invoke-GraphBatch' {
    BeforeAll {
        Mock Start-Sleep { }
        Mock Write-SyncLog { }
    }

    It 'splits requests into batches of 20 and returns one result per request' {
        $script:batchCalls = @()
        Mock Invoke-GraphRequestWithRetry {
            param($Method, $Uri, $Body)
            $parsed = $Body | ConvertFrom-Json
            $script:batchCalls += , @($parsed.requests)
            @{ responses = @($parsed.requests | ForEach-Object { @{ id = $_.id; status = 201; body = @{ id = "new-$($_.id)" } } }) }
        }
        $requests = 1..45 | ForEach-Object { @{ id = "r$_"; method = 'POST'; url = '/users/u/contacts'; body = @{ displayName = "p$_" } } }
        $results = Invoke-GraphBatch -Requests $requests
        $script:batchCalls.Count | Should -Be 3
        $script:batchCalls[0].Count | Should -Be 20
        $script:batchCalls[2].Count | Should -Be 5
        $results.Count | Should -Be 45
        $results['r45'].Status | Should -Be 201
        Should -Invoke Invoke-GraphRequestWithRetry -Times 3 -Exactly
    }

    It 'includes a JSON content-type header and body only for POST and PATCH' {
        $script:lastBatch = $null
        Mock Invoke-GraphRequestWithRetry {
            param($Method, $Uri, $Body)
            $script:lastBatch = $Body | ConvertFrom-Json
            @{ responses = @($script:lastBatch.requests | ForEach-Object { @{ id = $_.id; status = 204 } }) }
        }
        $requests = @(
            @{ id = 'a'; method = 'POST'; url = '/u'; body = @{ x = 1 } },
            @{ id = 'b'; method = 'DELETE'; url = '/u/1'; body = $null }
        )
        $null = Invoke-GraphBatch -Requests $requests
        $post = $script:lastBatch.requests | Where-Object id -EQ 'a'
        $del = $script:lastBatch.requests | Where-Object id -EQ 'b'
        $post.headers.'Content-Type' | Should -Be 'application/json'
        $post.body.x | Should -Be 1
        $del.PSObject.Properties.Name | Should -Not -Contain 'body'
        $del.PSObject.Properties.Name | Should -Not -Contain 'headers'
    }

    It 'retries only the sub-requests that came back 429 and honors Retry-After' {
        $script:round = 0
        Mock Invoke-GraphRequestWithRetry {
            param($Method, $Uri, $Body)
            $script:round++
            $parsed = $Body | ConvertFrom-Json
            @{
                responses = @($parsed.requests | ForEach-Object {
                        if ($script:round -eq 1 -and $_.id -eq 'r2') { @{ id = $_.id; status = 429; headers = @{ 'Retry-After' = '7' } } }
                        else { @{ id = $_.id; status = 200 } }
                    })
            }
        }
        $requests = 1..3 | ForEach-Object { @{ id = "r$_"; method = 'PATCH'; url = "/u/$_"; body = @{ } } }
        $results = Invoke-GraphBatch -Requests $requests
        $results.Count | Should -Be 3
        $results['r2'].Status | Should -Be 200
        Should -Invoke Invoke-GraphRequestWithRetry -Times 2 -Exactly
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 7 }
    }

    It 'gives up with a 429 result after the retry rounds are exhausted' {
        Mock Invoke-GraphRequestWithRetry {
            param($Method, $Uri, $Body)
            $parsed = $Body | ConvertFrom-Json
            @{ responses = @($parsed.requests | ForEach-Object { @{ id = $_.id; status = 429; headers = @{ 'Retry-After' = '1' } } }) }
        }
        $results = Invoke-GraphBatch -Requests @(@{ id = 'x'; method = 'DELETE'; url = '/u/1'; body = $null }) -MaxRounds 2
        $results['x'].Status | Should -Be 429
        Should -Invoke Invoke-GraphRequestWithRetry -Times 2 -Exactly
    }
}

Describe 'Get-OwnedContacts request shape' {
    It 'expands only the two ownership properties, selects minimal fields and pages at 999' {
        $script:uri = $null
        Mock Get-GraphCollection { param($Uri) $script:uri = $Uri; @() }
        $null = Get-OwnedContacts -UserId 'user-1'
        $c = Get-DirectoryContactConstants
        $script:uri | Should -Match '^https://graph\.microsoft\.com/v1\.0/users/user-1/contacts\?'
        $script:uri | Should -Match '\$top=999'
        $script:uri | Should -Match '\$select=id,displayName'
        $script:uri | Should -Match ([regex]::Escape("`$expand=singleValueExtendedProperties(`$filter=id eq '$($c.SourceIdPropertyId)' or id eq '$($c.SourceHashPropertyId)')"))
    }
}
