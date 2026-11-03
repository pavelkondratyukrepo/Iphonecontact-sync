#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Local Pester tests for runbooks/Sync-IntuneToEntra.ps1: phone normalization, newest-device selection and the stamp decision
    logic. Nothing here touches Graph or a tenant.
#>

BeforeAll {
    Set-StrictMode -Version Latest
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $runbook = Join-Path (Get-RepoRoot) 'runbooks' 'Sync-IntuneToEntra.ps1'
    . ([scriptblock]::Create((Get-RunbookFunctionText -Path $runbook)))
    $GraphBase = 'https://graph.microsoft.com/v1.0'

    function New-User {
        param([string]$Id, [string]$Upn, [bool]$Enabled = $true, [string]$Attr = '', [string]$Mobile = '')
        [pscustomobject]@{
            Id = $Id; UserPrincipalName = $Upn; AccountEnabled = $Enabled
            ScopeAttribute = $(if ($Attr) { $Attr } else { $null })
            MobilePhone    = $(if ($Mobile) { $Mobile } else { $null })
        }
    }
    function New-Device {
        param([string]$UserId, [string]$Phone, [string]$Sync, [string]$Serial = 'SER', [string]$Name = 'iPhone')
        ConvertTo-DeviceRecord -Device @{ id = [guid]::NewGuid().ToString(); userId = $UserId; userPrincipalName = 'x@contoso.com'; phoneNumber = $Phone; lastSyncDateTime = $Sync; serialNumber = $Serial; deviceName = $Name }
    }
}

Describe 'ConvertTo-E164' {
    It 'normalizes <Number> to <Expected>' -ForEach @(
        @{ Number = '8045550123'; Expected = '+18045550123' }
        @{ Number = '(804) 555-0123'; Expected = '+18045550123' }
        @{ Number = '804.555.0123'; Expected = '+18045550123' }
        @{ Number = '1 804 555 0123'; Expected = '+18045550123' }
        @{ Number = '18045550123'; Expected = '+18045550123' }
        @{ Number = '+1 (804) 555-0123'; Expected = '+18045550123' }
        @{ Number = '+18045550123'; Expected = '+18045550123' }
        @{ Number = '+44 20 7946 0958'; Expected = '+442079460958' }
        @{ Number = ' +49 30 901820 '; Expected = '+4930901820' }
    ) {
        ConvertTo-E164 -Number $Number -DefaultCountryCode '+1' | Should -Be $Expected
    }

    It 'rejects <Number>' -ForEach @(
        @{ Number = '' }
        @{ Number = '   ' }
        @{ Number = $null }
        @{ Number = '555-0123' }              # 7 digits
        @{ Number = '28045550123' }           # 11 digits not starting with 1
        @{ Number = '123456789012' }          # 12 digits, no +
        @{ Number = '+1234' }                 # too short even with +
        @{ Number = '+1234567890123456' }     # 16 digits
        @{ Number = 'Not Available' }
        @{ Number = 'N/A' }
    ) {
        ConvertTo-E164 -Number $Number -DefaultCountryCode '+1' | Should -BeNullOrEmpty
    }

    It 'uses the configured default country code for 10-digit numbers' {
        ConvertTo-E164 -Number '2079460958' -DefaultCountryCode '+44' | Should -Be '+442079460958'
        ConvertTo-E164 -Number '2079460958' -DefaultCountryCode '44' | Should -Be '+442079460958'
    }
}

Describe 'ConvertTo-DeviceRecord' {
    It 'reads camelCase hashtable fields (Invoke-MgGraphRequest) and parses the sync time as UTC' {
        $d = New-Device -UserId 'u1' -Phone '8045550123' -Sync '2026-10-30T12:00:00Z' -Serial 'F2LX' -Name 'Jane iPhone'
        $d.UserId | Should -Be 'u1'
        $d.PhoneNumber | Should -Be '8045550123'
        $d.SerialNumber | Should -Be 'F2LX'
        $d.LastSyncDateTime | Should -Be ([DateTime]::Parse('2026-10-30T12:00:00Z', [cultureinfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal))
    }

    It 'reads PascalCase object properties (Graph SDK cmdlets) and tolerates missing values' {
        $sdk = [pscustomobject]@{ Id = 'd1'; UserId = 'u1'; UserPrincipalName = 'a@b'; PhoneNumber = $null; LastSyncDateTime = [DateTime]'2026-10-30T12:00:00Z'; DeviceName = 'x'; SerialNumber = 's' }
        $d = ConvertTo-DeviceRecord -Device $sdk
        $d.UserId | Should -Be 'u1'
        $d.PhoneNumber | Should -Be ''
        $d.LastSyncDateTime.Year | Should -Be 2026
        $bare = ConvertTo-DeviceRecord -Device @{ id = 'd2' }
        $bare.UserId | Should -Be ''
        $bare.LastSyncDateTime | Should -Be ([DateTime]::MinValue)
    }
}

Describe 'Select-NewestDevicePerUser' {
    It 'keeps the device with the newest lastSyncDateTime per user, case-insensitive on user id' {
        $devices = @(
            (New-Device -UserId 'U1' -Phone '8045550001' -Sync '2026-10-01T00:00:00Z' -Serial 'OLD'),
            (New-Device -UserId 'u1' -Phone '8045550002' -Sync '2026-10-20T00:00:00Z' -Serial 'NEW'),
            (New-Device -UserId 'u1' -Phone '8045550003' -Sync '2026-10-10T00:00:00Z' -Serial 'MID'),
            (New-Device -UserId 'u2' -Phone '8045550004' -Sync '2026-10-05T00:00:00Z' -Serial 'ONLY'),
            (New-Device -UserId '' -Phone '8045550005' -Sync '2026-10-05T00:00:00Z' -Serial 'NOUSER')
        )
        $byUser = Select-NewestDevicePerUser -Devices $devices
        $byUser.Count | Should -Be 2
        $byUser['u1'].SerialNumber | Should -Be 'NEW'
        $byUser['u2'].SerialNumber | Should -Be 'ONLY'
    }

    It 'keeps the newest device even when it has no number (the number follows on a later inventory)' {
        $devices = @(
            (New-Device -UserId 'u1' -Phone '8045550001' -Sync '2026-10-01T00:00:00Z' -Serial 'OLD'),
            (New-Device -UserId 'u1' -Phone '' -Sync '2026-10-20T00:00:00Z' -Serial 'NEW')
        )
        (Select-NewestDevicePerUser -Devices $devices)['u1'].SerialNumber | Should -Be 'NEW'
    }

    It 'returns an empty table for no devices' {
        (Select-NewestDevicePerUser -Devices @()).Count | Should -Be 0
    }
}

Describe 'ConvertTo-UserRecord' {
    It 'reads extensionAttribute1 from a camelCase hashtable' {
        $u = ConvertTo-UserRecord -User @{ id = 'u1'; userPrincipalName = 'a@contoso.com'; accountEnabled = $true; mobilePhone = '+18045550123'; onPremisesExtensionAttributes = @{ extensionAttribute1 = 'CorpiPhone'; extensionAttribute2 = 'other' } }
        $u.ScopeAttribute | Should -Be 'CorpiPhone'
        $u.MobilePhone | Should -Be '+18045550123'
        $u.AccountEnabled | Should -BeTrue
    }

    It 'reads ExtensionAttribute1 from a PascalCase SDK object and maps empty to null' {
        $sdk = [pscustomobject]@{ Id = 'u1'; UserPrincipalName = 'a@contoso.com'; AccountEnabled = $false; MobilePhone = ''; OnPremisesExtensionAttributes = [pscustomobject]@{ ExtensionAttribute1 = '' } }
        $u = ConvertTo-UserRecord -User $sdk
        $u.ScopeAttribute | Should -BeNullOrEmpty
        $u.MobilePhone | Should -BeNullOrEmpty
        $u.AccountEnabled | Should -BeFalse
    }

    It 'tolerates a missing onPremisesExtensionAttributes block' {
        $u = ConvertTo-UserRecord -User @{ id = 'u1'; userPrincipalName = 'a@contoso.com'; accountEnabled = $true }
        $u.ScopeAttribute | Should -BeNullOrEmpty
    }
}

Describe 'Resolve-StampChanges' {
    BeforeAll {
        $inScope = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($id in 'u-new', 'u-already', 'u-disabled-in-scope') { $null = $inScope.Add($id) }
        $users = @(
            (New-User -Id 'u-new' -Upn 'new@contoso.com'),                                           # in scope, no stamp     -> stamp
            (New-User -Id 'u-already' -Upn 'already@contoso.com' -Attr 'CorpiPhone'),                # in scope, stamped      -> nothing
            (New-User -Id 'u-left' -Upn 'left@contoso.com' -Attr 'CorpiPhone'),                      # not in scope, stamped  -> unstamp
            (New-User -Id 'u-disabled-in-scope' -Upn 'disabled@contoso.com' -Enabled $false -Attr 'CorpiPhone'), # disabled -> unstamp
            (New-User -Id 'u-nothing' -Upn 'nothing@contoso.com')                                     # not in scope, empty    -> nothing
        )
    }

    It 'stamps new in-scope users and unstamps users that left scope or were disabled' {
        $plan = Resolve-StampChanges -Users $users -InScopeUserIds $inScope -ScopeAttributeValue 'CorpiPhone'
        @($plan.ToStamp | ForEach-Object Id) | Should -Be @('u-new')
        @($plan.ToUnstamp | ForEach-Object Id) | Should -Be @('u-left', 'u-disabled-in-scope')
        $plan.Conflicts.Count | Should -Be 0
    }

    It 'reports a foreign value as a conflict instead of overwriting it' {
        $withForeign = $users + @((New-User -Id 'u-foreign' -Upn 'foreign@contoso.com' -Attr 'CostCenter-42'))
        $null = $inScope.Add('u-foreign')
        $plan = Resolve-StampChanges -Users $withForeign -InScopeUserIds $inScope -ScopeAttributeValue 'CorpiPhone'
        @($plan.Conflicts | ForEach-Object Id) | Should -Be @('u-foreign')
        @($plan.ToStamp | ForEach-Object Id) | Should -Not -Contain 'u-foreign'
        $null = $inScope.Remove('u-foreign')
    }

    It 'is a no-op when everything already matches' {
        $steady = @((New-User -Id 'u-already' -Upn 'a@contoso.com' -Attr 'CorpiPhone'), (New-User -Id 'u-nothing' -Upn 'n@contoso.com'))
        $plan = Resolve-StampChanges -Users $steady -InScopeUserIds $inScope -ScopeAttributeValue 'CorpiPhone'
        $plan.ToStamp.Count | Should -Be 0
        $plan.ToUnstamp.Count | Should -Be 0
    }
}

Describe 'Set-UserScopeAttribute body' {
    It 'PATCHes only extensionAttribute1 and serializes null explicitly when clearing' {
        $script:bodies = @()
        Mock Invoke-GraphRequestWithRetry { param($Method, $Uri, $Body) $script:bodies += $Body; $null }
        Set-UserScopeAttribute -UserId 'u1' -AttributeName 'extensionAttribute1' -Value 'CorpiPhone'
        Set-UserScopeAttribute -UserId 'u1' -AttributeName 'extensionAttribute1' -Value $null
        $script:bodies[0] | Should -Be '{"onPremisesExtensionAttributes":{"extensionAttribute1":"CorpiPhone"}}'
        $script:bodies[1] | Should -Be '{"onPremisesExtensionAttributes":{"extensionAttribute1":null}}'
        Should -Invoke Invoke-GraphRequestWithRetry -Times 2 -Exactly -ParameterFilter { $Method -eq 'PATCH' -and $Uri -eq 'https://graph.microsoft.com/v1.0/users/u1' }
    }
}
