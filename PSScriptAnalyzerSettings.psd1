@{
    Severity     = @('Error', 'Warning')
    ExcludeRules = @(
        # Runbook parameters are plain bools and strings (Azure Automation does not pass switches),
        # and the design doc mandates a [bool]$WhatIf parameter on both runbooks.
        'PSReservedParams',
        'PSShouldProcess',
        'PSUseShouldProcessForStateChangingFunctions',
        # Write-Host is used only for operator-facing progress text in the deploy scripts, never for data.
        'PSAvoidUsingWriteHost',
        # Style only. Functions that return collections (Get-OwnedContacts, Resolve-MailboxChanges) read better plural.
        'PSUseSingularNouns'
    )
    Rules        = @{
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @('7.0')
        }
    }
}
