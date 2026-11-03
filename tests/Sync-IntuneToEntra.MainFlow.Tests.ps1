#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    End-to-end run of runbooks/Sync-IntuneToEntra.ps1 against the in-memory fake Graph in FakeGraph.ps1.
    Exercises the main block: device filtering, stamping and unstamping, conflicts, exclusions, privileged users, phone
    normalization and the summary object. No network, no tenant.
#>

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    . (Join-Path $PSScriptRoot 'FakeGraph.ps1')
    $script:Runbook = Join-Path (Get-RepoRoot) 'runbooks' 'Sync-IntuneToEntra.ps1'

    function New-Scenario {
        $t = New-FakeTenant
        # enabled, no stamp, one phone                      -> stamp + number
        $null = Add-FakeUser -Id 'u-new' -Upn 'new@contoso.com'
        Add-FakeDevice -UserId 'u-new' -Upn 'new@contoso.com' -Phone '(804) 555-0101' -Serial 'S-NEW'
        # already stamped, number already correct           -> nothing
        $null = Add-FakeUser -Id 'u-already' -Upn 'already@contoso.com' -Ext1 'CorpiPhone' -Mobile '+18045550102'
        Add-FakeDevice -UserId 'u-already' -Upn 'already@contoso.com' -Phone '804-555-0102'
        # stamped but no device any more                    -> unstamp
        $null = Add-FakeUser -Id 'u-left' -Upn 'left@contoso.com' -Ext1 'CorpiPhone' -Mobile '+18045550103'
        # disabled but still stamped and still has a device -> unstamp, device skipped
        $null = Add-FakeUser -Id 'u-disabled' -Upn 'disabled@contoso.com' -Enabled $false -Ext1 'CorpiPhone'
        Add-FakeDevice -UserId 'u-disabled' -Upn 'disabled@contoso.com' -Phone '8045550104'
        # exclusion group member                            -> stamp, number skipped
        $null = Add-FakeUser -Id 'u-excl' -Upn 'excl@contoso.com'
        Add-FakeDevice -UserId 'u-excl' -Upn 'excl@contoso.com' -Phone '8045550105'
        # direct role assignment                            -> stamp, number skipped
        $null = Add-FakeUser -Id 'u-admin' -Upn 'admin@contoso.com'
        Add-FakeDevice -UserId 'u-admin' -Upn 'admin@contoso.com' -Phone '8045550106'
        # role assigned through a group                     -> stamp, number skipped
        $null = Add-FakeUser -Id 'u-grpadmin' -Upn 'grpadmin@contoso.com'
        Add-FakeDevice -UserId 'u-grpadmin' -Upn 'grpadmin@contoso.com' -Phone '8045550107'
        # device reports no number                          -> stamp, NoPhoneNumber
        $null = Add-FakeUser -Id 'u-nonum' -Upn 'nonum@contoso.com'
        Add-FakeDevice -UserId 'u-nonum' -Upn 'nonum@contoso.com' -Phone ''
        # device reports garbage                            -> stamp, InvalidPhoneNumber
        $null = Add-FakeUser -Id 'u-bad' -Upn 'bad@contoso.com'
        Add-FakeDevice -UserId 'u-bad' -Upn 'bad@contoso.com' -Phone 'Not Available'
        # two devices, newest wins                          -> +18045550111
        $null = Add-FakeUser -Id 'u-two' -Upn 'two@contoso.com' -Mobile '+18045550110'
        Add-FakeDevice -UserId 'u-two' -Upn 'two@contoso.com' -Phone '8045550110' -Sync '2026-09-01T00:00:00Z' -Serial 'S-OLD'
        Add-FakeDevice -UserId 'u-two' -Upn 'two@contoso.com' -Phone '8045550111' -Sync '2026-10-30T00:00:00Z' -Serial 'S-NEWER'
        # noise: no primary user, unknown user, Android, personal
        Add-FakeDevice -UserId '' -Upn '' -Phone '8045550120' -Serial 'S-NOUSER'
        Add-FakeDevice -UserId 'ghost' -Upn 'ghost@contoso.com' -Phone '8045550121' -Serial 'S-GHOST'
        Add-FakeDevice -UserId 'u-new' -Upn 'new@contoso.com' -Phone '8045550122' -Os 'Android'
        Add-FakeDevice -UserId 'u-new' -Upn 'new@contoso.com' -Phone '8045550123' -Owner 'personal'

        Add-FakeGroup -Name 'SG-Directory-Exclude' -Id 'g-excl' -Members @('u-excl')
        Add-FakeGroup -Name 'Tenant Admins' -Id 'g-admins' -Members @('u-grpadmin')
        $t.RoleAssignments.Add(@{ principalId = 'u-admin' })
        $t.RoleAssignments.Add(@{ principalId = 'g-admins' })
        $t.RoleAssignments.Add(@{ principalId = 'sp-some-service-principal' })
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

    function Get-Ext1 { param([string]$Id) ($global:FakeTenant.Users | Where-Object { $_.id -eq $Id }).onPremisesExtensionAttributes.extensionAttribute1 }
    function Get-Mobile { param([string]$Id) ($global:FakeTenant.Users | Where-Object { $_.id -eq $Id }).mobilePhone }
}

Describe 'Sync-IntuneToEntra main flow' {
    Context 'WhatIf (default)' {
        BeforeAll {
            $null = New-Scenario
            $script:run = Invoke-Runbook -Parameters @{ ExclusionGroupName = 'SG-Directory-Exclude' }
        }

        It 'writes nothing' {
            $global:FakeTenant.Patches.Count | Should -Be 0
            (Get-Ext1 'u-new') | Should -BeNullOrEmpty
            (Get-Ext1 'u-left') | Should -Be 'CorpiPhone'
        }

        It 'reports what it would do' {
            $s = $script:run.Summary
            $s.WhatIf | Should -BeTrue
            $s.DevicesRead | Should -Be 12
            $s.DevicesWithoutPrimaryUser | Should -Be 1
            $s.UsersInScope | Should -Be 8
            $s.UsersStamped | Should -Be 7
            $s.UsersUnstamped | Should -Be 2
            $s.NumbersUpdated | Should -Be 2
            $s.Errors | Should -Be 0
            ($script:run.Lines -join "`n") | Should -Match 'WHATIF STAMP\s+new@contoso\.com'
            ($script:run.Lines -join "`n") | Should -Match 'WHATIF PHONE\s+new@contoso\.com mobilePhone \(none\) -> \+18045550101'
        }
    }

    Context 'Live run' {
        BeforeAll {
            $null = New-Scenario
            $script:run = Invoke-Runbook -Parameters @{ WhatIf = $false; ExclusionGroupName = 'SG-Directory-Exclude' }
            $script:summary = $script:run.Summary
            $script:log = $script:run.Lines -join "`n"
        }

        It 'stamps every enabled primary user of a corporate iPhone' {
            foreach ($id in 'u-new', 'u-excl', 'u-admin', 'u-grpadmin', 'u-nonum', 'u-bad', 'u-two', 'u-already') { (Get-Ext1 $id) | Should -Be 'CorpiPhone' -Because "$id should be stamped" }
            $script:summary.UsersStamped | Should -Be 7
        }

        It 'clears the stamp for users without a corporate iPhone and for disabled users, with an explicit JSON null' {
            (Get-Ext1 'u-left') | Should -BeNullOrEmpty
            (Get-Ext1 'u-disabled') | Should -BeNullOrEmpty
            $script:summary.UsersUnstamped | Should -Be 2
            $unstamp = @($global:FakeTenant.Patches | Where-Object { $_.UserId -eq 'u-left' })
            $unstamp.Count | Should -Be 1
            $unstamp[0].Body | Should -Be '{"onPremisesExtensionAttributes":{"extensionAttribute1":null}}'
        }

        It 'never touches any other extension attribute' {
            foreach ($p in $global:FakeTenant.Patches) {
                $p.Body | Should -Not -Match 'extensionAttribute(2|3|4|5|6|7|8|9|1[0-5])'
            }
        }

        It 'writes normalized numbers from the newest device and leaves correct numbers alone' {
            (Get-Mobile 'u-new') | Should -Be '+18045550101'
            (Get-Mobile 'u-two') | Should -Be '+18045550111'
            (Get-Mobile 'u-already') | Should -Be '+18045550102'
            @($global:FakeTenant.Patches | Where-Object { $_.UserId -eq 'u-already' }).Count | Should -Be 0
            $script:summary.NumbersUpdated | Should -Be 2
            $script:summary.NumbersUnchanged | Should -Be 1
            $script:log | Should -Match 'APPLY PHONE\s+two@contoso\.com mobilePhone \+18045550110 -> \+18045550111 \(device .* serial S-NEWER'
        }

        It 'skips the phone write but keeps the stamp for excluded and privileged users' {
            (Get-Mobile 'u-excl') | Should -BeNullOrEmpty
            (Get-Mobile 'u-admin') | Should -BeNullOrEmpty
            (Get-Mobile 'u-grpadmin') | Should -BeNullOrEmpty
            $script:summary.UsersSkipped.ExcludedGroupMember | Should -Be 1
            $script:summary.UsersSkipped.PrivilegedRole | Should -Be 2
        }

        It 'skips and explains missing or invalid numbers and orphaned devices' {
            $script:summary.UsersSkipped.NoPhoneNumber | Should -Be 1
            $script:summary.UsersSkipped.InvalidPhoneNumber | Should -Be 1
            $script:summary.UsersSkipped.DeviceUserNotMember | Should -Be 1
            $script:summary.UsersSkipped.DeviceUserDisabled | Should -Be 1
            $script:summary.DevicesWithoutPrimaryUser | Should -Be 1
            $script:log | Should -Match "serial S-NOUSER has no primary user"
            $script:log | Should -Match "InvalidPhoneNumber : bad@contoso\.com .* 'Not Available'"
        }

        It 'ends clean' {
            $script:summary.Errors | Should -Be 0
            $script:summary.WhatIf | Should -BeFalse
        }

        It 'is idempotent: a second live run changes nothing' {
            $before = $global:FakeTenant.Patches.Count
            $again = Invoke-Runbook -Parameters @{ WhatIf = $false; ExclusionGroupName = 'SG-Directory-Exclude' }
            $global:FakeTenant.Patches.Count | Should -Be $before
            $again.Summary.UsersStamped | Should -Be 0
            $again.Summary.UsersUnstamped | Should -Be 0
            $again.Summary.NumbersUpdated | Should -Be 0
        }
    }

    Context 'Safety rails' {
        It 'stops before writing anything when a user carries a foreign value in extensionAttribute1' {
            $null = New-Scenario
            $null = Add-FakeUser -Id 'u-foreign' -Upn 'foreign@contoso.com' -Ext1 'CostCenter-42'
            { & $script:Runbook -WhatIf $false -ExclusionGroupName 'SG-Directory-Exclude' } | Should -Throw '*foreign value*'
            $global:FakeTenant.Patches.Count | Should -Be 0
        }

        It 'fails before writing anything when the exclusion group does not exist' {
            $null = New-Scenario
            { & $script:Runbook -WhatIf $false -ExclusionGroupName 'SG-Does-Not-Exist' } | Should -Throw '*Exclusion group*not found*'
            $global:FakeTenant.Patches.Count | Should -Be 0
        }

        It 'runs without an exclusion group when the name is empty' {
            $null = New-Scenario
            $run = Invoke-Runbook -Parameters @{ WhatIf = $false; ExclusionGroupName = '' }
            (Get-Mobile 'u-excl') | Should -Be '+18045550105'
            $run.Summary.Errors | Should -Be 0
        }

        It 'never calls the authentication methods API' {
            $text = Get-Content -Path $script:Runbook -Raw
            $text | Should -Not -Match 'authentication/'
            $text | Should -Not -Match 'phoneMethods'
        }
    }
}
