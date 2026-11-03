#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }, PSScriptAnalyzer
<#
    Lints every script in runbooks/ and deploy/ with PSScriptAnalyzer using the repo's settings file, and checks the
    runbook-specific constraints from the design: bool/string parameters only, strict mode, no Write-Host, no interactive prompts.
#>

BeforeDiscovery {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $scripts = @(Get-ChildItem -Path (Join-Path $root 'runbooks'), (Join-Path $root 'deploy') -Filter '*.ps1' | ForEach-Object { @{ Path = $_.FullName; Name = $_.Name } })
    $runbooks = @(Get-ChildItem -Path (Join-Path $root 'runbooks') -Filter '*.ps1' | ForEach-Object { @{ Path = $_.FullName; Name = $_.Name } })
}

BeforeAll {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $settings = Join-Path $root 'PSScriptAnalyzerSettings.psd1'
}

Describe 'PSScriptAnalyzer' {
    It '<Name> has no errors or warnings' -ForEach $scripts {
        $findings = @(Invoke-ScriptAnalyzer -Path $Path -Settings $settings -Severity Error, Warning)
        $report = ($findings | ForEach-Object { "  [$($_.Severity)] $($_.RuleName) line $($_.Line): $($_.Message)" }) -join "`n"
        $findings.Count | Should -Be 0 -Because "the analyzer reported:`n$report"
    }

    It '<Name> parses without errors' -ForEach $scripts {
        $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)
        @($errors).Count | Should -Be 0
    }
}

Describe 'Runbook constraints' {
    It '<Name> declares only bool and string parameters (Automation passes no switches)' -ForEach $runbooks {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $params = $ast.ParamBlock.Parameters
        $params.Count | Should -BeGreaterThan 0
        foreach ($p in $params) {
            $typeName = ($p.Attributes | Where-Object { $_ -is [System.Management.Automation.Language.TypeConstraintAst] } | Select-Object -First 1).TypeName.FullName
            $typeName | Should -BeIn @('bool', 'string', 'System.Boolean', 'System.String') -Because "parameter $($p.Name) must be a bool or string"
        }
    }

    It '<Name> defaults WhatIf to $true' -ForEach $runbooks {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $whatIf = $ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'WhatIf' }
        $whatIf | Should -Not -BeNullOrEmpty
        $whatIf.DefaultValue.Extent.Text | Should -Be '$true'
    }

    It '<Name> sets strict mode, uses no Write-Host, no Read-Host and no -Confirm' -ForEach $runbooks {
        $text = Get-Content -Path $Path -Raw
        $text | Should -Match 'Set-StrictMode -Version Latest'
        $text | Should -Not -Match '(?m)^\s*Write-Host'
        $text | Should -Not -Match 'Read-Host'
        $text | Should -Not -Match '-Confirm'
        $text | Should -Match 'Connect-MgGraph -Identity'
    }

    It '<Name> has comment-based help with a synopsis and description' -ForEach $runbooks {
        $help = (Get-Content -Path $Path -Raw)
        $help | Should -Match '\.SYNOPSIS'
        $help | Should -Match '\.DESCRIPTION'
    }

    It '<Name> never consents or requests Contacts permissions in Entra' -ForEach $runbooks {
        $text = Get-Content -Path $Path -Raw
        # The runbook may mention Contacts.ReadWrite in comments explaining the Exchange scope, but must never request it.
        $text | Should -Not -Match 'Connect-MgGraph\s+-Scopes'
    }
}

Describe 'Repository hygiene' {
    It 'contains no GUID that looks like a tenant, subscription or object id outside the fixed extended-property namespace' {
        $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
        $allowed = @(
            'b4b9c0a1-5f0e-4c3d-9a7e-2d6f1c8e0a11',   # extended property namespace (fixed by design)
            '00000003-0000-0000-c000-000000000000',   # Microsoft Graph well-known app id
            '00000000-0000-0000-0000-000000000000'    # placeholder in examples
        )
        $files = Get-ChildItem -Path $root -Recurse -Include '*.ps1', '*.psd1', '*.md' | Where-Object { $_.FullName -notmatch '[\\/]tests[\\/]' -and $_.FullName -notmatch '[\\/]\.git[\\/]' }
        $offenders = foreach ($f in $files) {
            $text = Get-Content -Path $f.FullName -Raw
            foreach ($m in [regex]::Matches($text, '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')) {
                if ($allowed -notcontains $m.Value.ToLowerInvariant()) { "$($f.Name): $($m.Value)" }
            }
        }
        @($offenders) | Should -BeNullOrEmpty
    }
}
