#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    End-to-end run of runbooks/Sync-DirectoryToMailboxes.ps1 against the in-memory fake Graph in FakeGraph.ps1.
    Exercises the main block: group reads, exclusion, per-mailbox listing with paging, reconcile, batching with a 429 retry,
    403/404 handling, -UserPrincipalName, idempotency and the purge path. No network, no tenant.
#>

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    . (Join-Path $PSScriptRoot 'FakeGraph.ps1')
    $script:Runbook = Join-Path (Get-RepoRoot) 'runbooks' 'Sync-DirectoryToMailboxes.ps1'
    # The runbook's own pure functions, used here only to compute a "current" hash for pre-seeded contacts.
    . ([scriptblock]::Create((Get-RunbookFunctionText -Path $script:Runbook)))
    $GraphBase = 'https://graph.microsoft.com/v1.0'
    $BatchSize = 20

    function Get-CurrentHash {
        param([string]$UserId)
        $u = $global:FakeTenant.Users | Where-Object { $_.id -eq $UserId }
        Get-DirectoryContactHash -Fields (ConvertTo-DirectoryContactFields -User (ConvertTo-DirectoryUserRecord -User $u))
    }

    function Get-CurrentFields {
        <# The fields exactly as the runbook writes them for this user, as a plain hashtable for Add-FakeContact. #>
        param([string]$UserId)
        $u = $global:FakeTenant.Users | Where-Object { $_.id -eq $UserId }
        $f = ConvertTo-DirectoryContactFields -User (ConvertTo-DirectoryUserRecord -User $u)
        $h = @{}
        foreach ($k in $f.Keys) { $h[$k] = $f[$k] }
        return $h
    }

    function New-Scenario {
        $t = New-FakeTenant
        $null = Add-FakeUser -Id 'u1' -Upn 'ann.lee@contoso.com' -Ext1 'CorpiPhone' -Mobile '+18045550001' -Extra @{ businessPhones = @('+18045550100'); jobTitle = 'CFO' }
        $null = Add-FakeUser -Id 'u2' -Upn 'bob.ray@contoso.com' -Ext1 'CorpiPhone' -Mobile '+18045550002'
        $null = Add-FakeUser -Id 'u3' -Upn 'cat.exec@contoso.com' -Ext1 'CorpiPhone' -Mobile '+18045550003'   # excluded: receives, not listed
        $null = Add-FakeUser -Id 'u4' -Upn 'dan.gone@contoso.com' -Enabled $false -Ext1 'CorpiPhone'          # disabled
        $null = Add-FakeUser -Id 'u5' -Upn 'eve.fresh@contoso.com' -Ext1 'CorpiPhone' -Mobile '+18045550005'  # mailbox 403 (scope lag)
        $null = Add-FakeUser -Id 'u6' -Upn 'fay.nomail@contoso.com' -Ext1 'CorpiPhone'                       # no mailbox
        $null = Add-FakeUser -Id 'u7' -Upn 'gus.ok@contoso.com' -Ext1 'CorpiPhone' -Mobile '+18045550007'
        $null = Add-FakeUser -Id 'u8' -Upn "hal.o'neil@contoso.com" -Ext1 'CorpiPhone' -Extra @{ displayName = "O'Neil, Hal"; surname = "O'Neil"; givenName = 'Hal' }
        $null = Add-FakeUser -Id 'u9' -Upn 'ivy.outside@contoso.com'                                          # not in the group at all
        Add-FakeGroup -Name 'SG-CorpiPhone-Users' -Id 'g-control' -Members @('u1', 'u2', 'u3', 'u4', 'u5', 'u6', 'u7', 'u8')
        Add-FakeGroup -Name 'SG-Directory-Exclude' -Id 'g-excl' -Members @('u3')
        $null = $t.ForbiddenMailboxes.Add('u5')
        $null = $t.NoMailbox.Add('u6')

        # Mailbox u1 pre-state
        $null = Add-FakeContact -MailboxUserId 'u1' -DisplayName 'Mom'                                           # personal, untouchable
        $null = Add-FakeContact -MailboxUserId 'u1' -DisplayName 'Dentist' -Fields @{ singleValueExtendedProperties = @(@{ id = 'String {00000000-0000-0000-0000-000000000001} Name SomethingElse'; value = 'x' }) }
        $null = Add-FakeContact -MailboxUserId 'u1' -DisplayName 'bob.ray (old title)' -SourceId 'u2' -Hash 'stale'   # stale -> update
        $null = Add-FakeContact -MailboxUserId 'u1' -DisplayName 'bob.ray (dupe)' -SourceId 'u2' -Hash 'stale'        # duplicate -> delete
        $null = Add-FakeContact -MailboxUserId 'u1' -DisplayName 'Lee Ver' -SourceId 'u-gone'                           # leaver -> delete
        $null = Add-FakeContact -MailboxUserId 'u1' -DisplayName 'gus.ok' -SourceId 'u7' -Hash (Get-CurrentHash 'u7') -Fields (Get-CurrentFields 'u7')   # current -> unchanged
        # Mailbox u2 pre-state: a card for themselves, written by an older version -> deleted
        $null = Add-FakeContact -MailboxUserId 'u2' -DisplayName 'bob.ray' -SourceId 'u2' -Hash (Get-CurrentHash 'u2')
        # Throttle the first create into u2 once
        $null = $t.ThrottleOnce.Add('POST /users/u2/contacts')
        return $t
    }

    function Invoke-Runbook {
        param([hashtable]$Parameters)
        # 3>&1 merges the warning stream (retry/throttle messages) into the captured output.
        $output = @(& $script:Runbook @Parameters 3>&1)
        $lines = @($output | ForEach-Object { if ($_ -is [System.Management.Automation.WarningRecord]) { "WARNING: $($_.Message)" } elseif ($_ -is [string]) { $_ } })
        [pscustomobject]@{
            Lines   = $lines
            Summary = ($output | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.WarningRecord] } | Select-Object -Last 1)
        }
    }

    function Get-SourceIds { param([string]$MailboxUserId) @(Get-OwnedFakeContacts -MailboxUserId $MailboxUserId | ForEach-Object { ($_.singleValueExtendedProperties | Where-Object { $_.id -match 'DirectorySourceId' }).value } | Sort-Object) }
}

Describe 'Sync-DirectoryToMailboxes main flow' {
    Context 'WhatIf (default)' {
        BeforeAll {
            $null = New-Scenario
            $script:run = Invoke-Runbook -Parameters @{ GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
        }

        It 'sends no batch and leaves mailboxes untouched' {
            $global:FakeTenant.BatchCalls | Should -Be 0
            $global:FakeTenant.Mailboxes['u2'].Count | Should -Be 1
            $global:FakeTenant.Mailboxes['u3'].Count | Should -Be 0
        }

        It 'reports the planned changes' {
            $s = $script:run.Summary
            $s.WhatIf | Should -BeTrue
            $s.DirectorySize | Should -Be 6
            $s.MailboxesTargeted | Should -Be 7
            $s.ContactsAdded | Should -Be 24
            $s.ContactsUpdated | Should -Be 1
            $s.ContactsDeleted | Should -Be 3
            $s.ContactsUnchanged | Should -Be 1
            $s.Errors | Should -Be 0
            ($script:run.Lines -join "`n") | Should -Match 'WHATIF UPDATE mailbox=ann\.lee@contoso\.com ray, bob <bob\.ray@contoso\.com>'
        }
    }

    Context 'Live run' {
        BeforeAll {
            $null = New-Scenario
            $script:run = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $script:summary = $script:run.Summary
            $script:log = $script:run.Lines -join "`n"
        }

        It 'builds the directory from enabled members minus the exclusion group' {
            $script:summary.DirectorySize | Should -Be 6
            $script:log | Should -Match "Exclusion group 'SG-Directory-Exclude' has 1 user members"
        }

        It 'gives every reachable member one card per coworker, never one for themselves' {
            (Get-SourceIds 'u1') | Should -Be @('u2', 'u5', 'u6', 'u7', 'u8')
            (Get-SourceIds 'u2') | Should -Be @('u1', 'u5', 'u6', 'u7', 'u8')
            (Get-SourceIds 'u7') | Should -Be @('u1', 'u2', 'u5', 'u6', 'u8')
            (Get-SourceIds 'u8') | Should -Be @('u1', 'u2', 'u5', 'u6', 'u7')
        }

        It 'excluded members receive the full directory but appear in nobody''s contacts' {
            (Get-SourceIds 'u3') | Should -Be @('u1', 'u2', 'u5', 'u6', 'u7', 'u8')
            foreach ($box in 'u1', 'u2', 'u7', 'u8') { (Get-SourceIds $box) | Should -Not -Contain 'u3' }
        }

        It 'leaves personal contacts and contacts with other extended properties untouched' {
            $names = @($global:FakeTenant.Mailboxes['u1'] | ForEach-Object displayName)
            $names | Should -Contain 'Mom'
            $names | Should -Contain 'Dentist'
            $global:FakeTenant.Mailboxes['u1'].Count | Should -Be 7   # 2 personal + 5 owned
        }

        It 'updates the stale card in place, deletes the duplicate and the leaver, keeps the current one' {
            $bob = @($global:FakeTenant.Mailboxes['u1'] | Where-Object { ($_.singleValueExtendedProperties | Where-Object { $_.id -match 'SourceId' }).value -eq 'u2' })
            $bob.Count | Should -Be 1
            $bob[0].id | Should -Be 'c3'                                   # same contact id as the pre-seeded stale card
            $bob[0].displayName | Should -Be 'bob.ray'
            $bob[0].fileAs | Should -Be 'ray, bob'
            ($bob[0].singleValueExtendedProperties | Where-Object { $_.id -match 'SourceHash' }).value | Should -Be (Get-CurrentHash 'u2')
            $script:summary.ContactsUpdated | Should -Be 1
            $script:summary.ContactsDeleted | Should -Be 3                 # dupe + leaver in u1, self-card in u2
            $script:summary.ContactsUnchanged | Should -Be 1
            $script:log | Should -Match "APPLY DELETE mailbox=ann\.lee@contoso\.com 'Lee Ver' source=u-gone \(left directory\)"
            $script:log | Should -Match "APPLY DELETE mailbox=ann\.lee@contoso\.com 'bob\.ray \(dupe\)' source=u2 \(duplicate\)"
        }

        It 'writes complete, tagged contact bodies' {
            $hal = @($global:FakeTenant.Mailboxes['u1'] | Where-Object { ($_.singleValueExtendedProperties | Where-Object { $_.id -match 'SourceId' }).value -eq 'u8' })[0]
            $hal.displayName | Should -Be "O'Neil, Hal"
            $hal.fileAs | Should -Be "O'Neil, Hal"
            $hal.givenName | Should -Be 'Hal'
            $hal.companyName | Should -Be 'Contoso'
            @($hal.categories) | Should -Be @('Company Directory')
            @($hal.emailAddresses)[0].address | Should -Be "hal.o'neil@contoso.com"
            $hal.ContainsKey('mobilePhone') | Should -BeTrue
            $hal.mobilePhone | Should -BeNullOrEmpty
        }

        It 'retries a throttled sub-request and still creates everything' {
            $global:FakeTenant.ThrottleOnce.Count | Should -Be 0          # the 429 was consumed
            (Get-SourceIds 'u2').Count | Should -Be 5
            $script:log | Should -Match 'throttled, retrying'
        }

        It 'skips unreachable mailboxes without failing the job' {
            $script:summary.MailboxesSkipped.AccountDisabled | Should -Be 1
            $script:summary.MailboxesSkipped.AccessDenied | Should -Be 1
            $script:summary.MailboxesSkipped.NoMailbox | Should -Be 1
            $script:summary.MailboxesProcessed | Should -Be 5
            $script:summary.Errors | Should -Be 0
            $script:log | Should -Match 'AccessDenied : eve\.fresh@contoso\.com : 403'
        }

        It 'totals add up' {
            $script:summary.ContactsAdded | Should -Be 24
            $script:summary.MailboxesTargeted | Should -Be 7
        }

        It 'is idempotent: a second live run sends no batches and reports everything unchanged' {
            $batchesBefore = $global:FakeTenant.BatchCalls
            $again = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $global:FakeTenant.BatchCalls | Should -Be $batchesBefore
            $again.Summary.ContactsAdded | Should -Be 0
            $again.Summary.ContactsUpdated | Should -Be 0
            $again.Summary.ContactsDeleted | Should -Be 0
            $again.Summary.ContactsUnchanged | Should -Be 26
        }

        It 'leaves a phone-side edit alone when DetectLocalEdits is off' {
            # An edit on the phone changes the fields but not the stored hash (the phone never touches extended properties).
            $gus = @($global:FakeTenant.Mailboxes['u1'] | Where-Object { ($_.singleValueExtendedProperties | Where-Object { $_.id -match 'SourceId' }).value -eq 'u7' })[0]
            $gus.jobTitle = 'Edited on phone'
            $again = Invoke-Runbook -Parameters @{ WhatIf = $false; DetectLocalEdits = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $again.Summary.ContactsUpdated | Should -Be 0
            $gus.jobTitle | Should -Be 'Edited on phone'
        }

        It 'restores a contact a user edited on their phone and changes nothing else' {
            $gus = @($global:FakeTenant.Mailboxes['u1'] | Where-Object { ($_.singleValueExtendedProperties | Where-Object { $_.id -match 'SourceId' }).value -eq 'u7' })[0]
            $gus.jobTitle | Should -Be 'Edited on phone'
            $hashBefore = ($gus.singleValueExtendedProperties | Where-Object { $_.id -match 'SourceHash' }).value
            $hashBefore | Should -Be (Get-CurrentHash 'u7')                 # stored hash still says "current"
            $again = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $again.Summary.ContactsUpdated | Should -Be 1
            $again.Summary.LocalEditsReverted | Should -Be 1
            $again.Summary.ContactsAdded | Should -Be 0
            $again.Summary.ContactsDeleted | Should -Be 0
            ($again.Lines -join "`n") | Should -Match 'APPLY UPDATE mailbox=ann\.lee@contoso\.com .*gus\.ok@contoso\.com.*\(edited in mailbox\)'
            $gus.jobTitle | Should -Be 'Staff'
        }

        It 'does not see an edit after restoring it (no rewrite loop)' {
            $again = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $again.Summary.ContactsUpdated | Should -Be 0
            $again.Summary.LocalEditsReverted | Should -Be 0
        }

        It 'picks up a title change and a mobile number change on the next run' {
            $u7 = $global:FakeTenant.Users | Where-Object { $_.id -eq 'u7' }
            $u7.jobTitle = 'Head of Field Ops'
            $u7.mobilePhone = '+18045559999'
            $again = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $again.Summary.ContactsUpdated | Should -Be 4                 # one card for u7 in each other reachable mailbox (u1, u2, u3, u8)
            $gus = @($global:FakeTenant.Mailboxes['u1'] | Where-Object { ($_.singleValueExtendedProperties | Where-Object { $_.id -match 'SourceId' }).value -eq 'u7' })[0]
            $gus.mobilePhone | Should -Be '+18045559999'
        }
    }

    Context 'Onboarding with -UserPrincipalName' {
        BeforeAll { $null = New-Scenario }

        It 'fills only that mailbox, with the whole directory' {
            $run = Invoke-Runbook -Parameters @{ WhatIf = $false; UserPrincipalName = 'BOB.RAY@contoso.com'; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Summary.MailboxesTargeted | Should -Be 1
            $run.Summary.MailboxesProcessed | Should -Be 1
            (Get-SourceIds 'u2') | Should -Be @('u1', 'u5', 'u6', 'u7', 'u8')
            (Get-SourceIds 'u3').Count | Should -Be 0
            $global:FakeTenant.Mailboxes['u1'].Count | Should -Be 6       # untouched pre-state
        }

        It 'refuses a user who is not an enabled member of the group' {
            { & $script:Runbook -WhatIf $false -UserPrincipalName 'ivy.outside@contoso.com' -GroupName 'SG-CorpiPhone-Users' -ExclusionGroupName 'SG-Directory-Exclude' } | Should -Throw '*not an enabled member*'
            { & $script:Runbook -WhatIf $false -UserPrincipalName 'dan.gone@contoso.com' -GroupName 'SG-CorpiPhone-Users' -ExclusionGroupName 'SG-Directory-Exclude' } | Should -Throw '*not an enabled member*'
        }
    }

    Context 'FullRebuild and purge' {
        It 'FullRebuild rewrites every owned contact even when unchanged' {
            $null = New-Scenario
            $null = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run = Invoke-Runbook -Parameters @{ WhatIf = $false; FullRebuild = $true; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Summary.ContactsUpdated | Should -Be 26
            $run.Summary.ContactsUnchanged | Should -Be 0
            $run.Summary.ContactsAdded | Should -Be 0
        }

        It 'removes every owned contact and nothing else when everyone is excluded (rollback path)' {
            $null = New-Scenario
            $null = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            Add-FakeGroup -Name 'SG-Directory-Exclude' -Id 'g-excl' -Members @('u1', 'u2', 'u3', 'u4', 'u5', 'u6', 'u7', 'u8')
            $run = Invoke-Runbook -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Summary.DirectorySize | Should -Be 0
            $run.Summary.ContactsDeleted | Should -Be 26
            foreach ($box in 'u1', 'u2', 'u3', 'u7', 'u8') { @(Get-OwnedFakeContacts -MailboxUserId $box).Count | Should -Be 0 }
            @($global:FakeTenant.Mailboxes['u1'] | ForEach-Object displayName) | Should -Be @('Mom', 'Dentist')
        }
    }

    Context 'Failure handling' {
        BeforeAll {
            function Invoke-RunbookCatching {
                param([hashtable]$Parameters)
                $collected = [System.Collections.Generic.List[object]]::new()
                $thrown = $null
                try { & $script:Runbook @Parameters 3>&1 | ForEach-Object { $collected.Add($_) } }
                catch { $thrown = $_ }
                [pscustomobject]@{
                    Lines   = @($collected | Where-Object { $_ -is [string] })
                    Summary = ($collected | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.WarningRecord] } | Select-Object -Last 1)
                    Error   = $thrown
                }
            }
        }

        It 'keeps going when a whole $batch call fails for one mailbox, then fails the job' {
            $null = New-Scenario
            $global:FakeTenant.BatchThrow = 400
            $run = Invoke-RunbookCatching -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Error | Should -Not -BeNullOrEmpty
            "$($run.Error)" | Should -Match 'completed with 1 error'
            $run.Summary | Should -Not -BeNullOrEmpty
            $run.Summary.Errors | Should -Be 1
            $run.Summary.MailboxesProcessed | Should -Be 4                  # the other four reachable mailboxes were still filled
            ($run.Lines -join "`n") | Should -Match 'Writing contacts failed for ann\.lee@contoso\.com'
            (Get-SourceIds 'u8') | Should -Be @('u1', 'u2', 'u5', 'u6', 'u7')
        }

        It 'retries a $batch call that fails with a transient 503' {
            $null = New-Scenario
            $global:FakeTenant.BatchThrow = 503
            $run = Invoke-RunbookCatching -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Error | Should -BeNullOrEmpty
            $run.Summary.Errors | Should -Be 0
            $run.Summary.MailboxesProcessed | Should -Be 5
        }

        It 'removes duplicates when a $batch call is retried after Exchange already committed the creates' {
            $null = New-Scenario
            $global:FakeTenant.BatchThrowAfterCommit = 504
            $run = Invoke-RunbookCatching -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Error | Should -BeNullOrEmpty
            $run.Summary.DuplicatesRemoved | Should -BeGreaterThan 0
            foreach ($box in 'u1', 'u2', 'u3', 'u7', 'u8') {
                $ids = @(Get-SourceIds $box)
                $ids.Count | Should -Be @($ids | Select-Object -Unique).Count -Because "mailbox $box must hold one card per coworker"
            }
            ($run.Lines -join "`n") | Should -Match 'duplicate card\(s\) left by a retried or overlapping write'
        }

        It 'treats a card deleted by the user between listing and update as a warning, not an error' {
            $null = New-Scenario
            $global:FakeTenant.FailBatchUrls['PATCH /users/u1/contacts/c3'] = 404
            $run = Invoke-RunbookCatching -Parameters @{ WhatIf = $false; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Error | Should -BeNullOrEmpty
            $run.Summary.Errors | Should -Be 0
        }

        It 'refuses to run when the identity holds a Contacts permission in Entra (unscoped grant)' {
            $null = New-Scenario
            $global:FakeTenant.ContextScopes = @('GroupMember.Read.All', 'User.ReadWrite.All', 'Contacts.ReadWrite')
            $run = Invoke-RunbookCatching -Parameters @{ WhatIf = $true; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            "$($run.Error)" | Should -Match 'holds Contacts\.ReadWrite in Entra'
            $global:FakeTenant.BatchCalls | Should -Be 0
        }

        It 'fails the job when every targeted mailbox returns 403 (broken RBAC assignment, not scope lag)' {
            $null = New-Scenario
            foreach ($id in 'u1', 'u2', 'u3', 'u5', 'u6', 'u7', 'u8') { $null = $global:FakeTenant.ForbiddenMailboxes.Add($id) }
            $null = $global:FakeTenant.NoMailbox.Remove('u6')
            $global:FakeTenant.Mailboxes['u6'] = [System.Collections.Generic.List[object]]::new()
            $run = Invoke-RunbookCatching -Parameters @{ WhatIf = $true; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Summary.MailboxesProcessed | Should -Be 0
            $run.Summary.MailboxesSkipped.AccessDenied | Should -Be 7
            "$($run.Error)" | Should -Match 'New-ExchangeAppScope\.ps1 -TestOnly'
        }

        It 'does not fail the job when only some mailboxes return 403' {
            $null = New-Scenario
            $run = Invoke-RunbookCatching -Parameters @{ WhatIf = $true; GroupName = 'SG-CorpiPhone-Users'; ExclusionGroupName = 'SG-Directory-Exclude' }
            $run.Error | Should -BeNullOrEmpty
            $run.Summary.MailboxesSkipped.AccessDenied | Should -Be 1
        }
    }

    Context 'Configuration errors' {
        It 'fails when the control group does not exist' {
            $null = New-Scenario
            { & $script:Runbook -WhatIf $false -GroupName 'SG-Nope' -ExclusionGroupName 'SG-Directory-Exclude' } | Should -Throw "*Control group 'SG-Nope' was not found*"
        }

        It 'fails before writing when the exclusion group does not exist' {
            $null = New-Scenario
            { & $script:Runbook -WhatIf $false -GroupName 'SG-CorpiPhone-Users' -ExclusionGroupName 'SG-Nope' } | Should -Throw '*Exclusion group*not found*'
            $global:FakeTenant.BatchCalls | Should -Be 0
        }
    }
}
