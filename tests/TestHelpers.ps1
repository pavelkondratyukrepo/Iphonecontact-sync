<#
.SYNOPSIS
    Shared helpers for the Pester tests. Dot-source from BeforeAll.

.DESCRIPTION
    The runbooks are single self-contained scripts whose main block connects to Graph as soon as it runs, so the tests cannot
    dot-source them. Instead the test parses the runbook with the PowerShell AST, extracts every top-level function definition
    and defines those functions in the test scope. The main block never executes.
#>

function Get-RunbookFunctionText {
    <# Returns the source text of every top-level function in the script at $Path, joined, ready for [scriptblock]::Create(). #>
    param([Parameter(Mandatory)] [string]$Path)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        throw "Parse errors in $Path`n$($errors | ForEach-Object { "  line $($_.Extent.StartLineNumber): $($_.Message)" } | Out-String)"
    }
    $functions = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            -not ($node.Parent -is [System.Management.Automation.Language.FunctionMemberAst])
        }, $false)
    # $false = do not descend into nested script blocks, so only top-level functions are returned.
    return (($functions | ForEach-Object { $_.Extent.Text }) -join "`n`n")
}

function Get-RepoRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}
