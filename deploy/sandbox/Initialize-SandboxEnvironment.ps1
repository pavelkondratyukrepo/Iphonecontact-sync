<#
.SYNOPSIS
    Turns a fresh Windows Sandbox (or any clean Windows machine) into a workstation that can run this repo's
    deploy scripts and tests. Paste the whole file into the Windows PowerShell 5.1 window that Sandbox opens.

.DESCRIPTION
    Windows Sandbox starts clean every time: PowerShell 5.1 only, no Git, no modules, no cached account.
    This script, in order:
      1. Installs PowerShell 7.4 (the deploy scripts and runbooks target 7.4) from the official MSI.
      2. Installs Git for Windows (silent) and clones the repo, or downloads a zip if you set a token.
      3. Inside PowerShell 7, installs the exact modules the repo needs into the PowerShell 7 module path:
            Az.Accounts, Az.Resources, Az.Automation, Az.Monitor
            Microsoft.Graph.Authentication / Users / Groups / Applications / DeviceManagement / PersonalContacts
               pinned to the same version the Automation runtime environment uses (2.40.0)
            ExchangeOnlineManagement 3.x, Pester 5+, PSScriptAnalyzer
         and turns off WAM (broker) sign-in for Az and Graph, which cannot work in Sandbox.
      4. Runs the repo's Pester tests so you know the checkout is good before touching the tenant.
      5. Opens a PowerShell 7 window in the repo folder with the next commands printed.

    Everything is lost when the Sandbox closes. Re-paste next time, or use deploy/sandbox/CorpiPhone-Sync.wsb to
    map a host folder and run this automatically at logon.

    Paste-safe: no Break/Return at script scope, no #Requires, no PowerShell 7 syntax, no interactive prompts
    except the sign-ins the deploy scripts themselves do later.

.NOTES
    Author:  Generated for the CorpiPhone directory sync project
    Date:    2026-11-03
    Version: 1.0
    Time:    about 10 to 15 minutes, most of it Az module downloads.
#>

# ============================================================
# CONFIGURATION  (edit these before pasting if needed)
# ============================================================
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$RepoUrl            = 'https://github.com/pavelkondratyukrepo/Iphonecontact-sync'
$RepoBranch         = 'claude/pensive-lamport-u50q36'   # change when the branch is merged or renamed
$WorkDir            = Join-Path $env:USERPROFILE 'Desktop\CorpiPhone'
$RepoDir            = Join-Path $WorkDir 'Iphonecontact-sync'
$GraphModuleVersion = '2.40.0'                          # must match -GraphModuleVersion in Deploy-Automation.ps1
$RunTests           = $true                             # run Invoke-Pester ./tests after setup
$OpenWindow         = $true                             # open a PowerShell 7 window in the repo when done

# Repo access. Leave $GitHubToken empty to clone with Git (Git Credential Manager opens a browser sign-in).
# Set a fine-grained PAT with read access to the repo to download a zip without Git (useful for a private repo
# when the browser flow is inconvenient). The token is used once, for one download, and never written to disk.
$GitHubToken        = ''

# Fallback versions if the GitHub API cannot be reached to resolve the newest 7.4.x / Git release.
$PwshFallbackVersion = '7.4.6'
$GitFallbackTag      = 'v2.47.1.windows.1'

$timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$downloadDir = Join-Path $env:TEMP "corpiphone-setup-$timestamp"
$pwshExe = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
$gitExe  = Join-Path $env:ProgramFiles 'Git\cmd\git.exe'

# ============================================================
# FUNCTIONS
# ============================================================
function Write-Step {
    param([string]$Message)
    Write-Host ''
    Write-Host ('==> ' + $Message) -ForegroundColor Cyan
}

function Get-LatestReleaseAsset {
    <#
        Resolves a download URL from the GitHub releases API. $TagPattern filters release tags (regex),
        $AssetPattern filters asset names (regex). Returns $null when the API is unreachable so the caller can fall back.
    #>
    param(
        [string]$Repo,
        [string]$TagPattern,
        [string]$AssetPattern
    )
    try {
        $releases = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases?per_page=30" -Headers @{ 'User-Agent' = 'corpiphone-sandbox-setup' } -TimeoutSec 30
        foreach ($rel in $releases) {
            if ($rel.prerelease -or $rel.draft) { continue }
            if ($rel.tag_name -notmatch $TagPattern) { continue }
            $asset = $rel.assets | Where-Object { $_.name -match $AssetPattern } | Select-Object -First 1
            if ($asset) {
                return [PSCustomObject]@{ Tag = $rel.tag_name; Name = $asset.name; Url = $asset.browser_download_url }
            }
        }
    }
    catch {
        Write-Host ("    GitHub API lookup failed for $Repo ($($_.Exception.Message)); using fallback version.") -ForegroundColor Yellow
    }
    return $null
}

function Save-Download {
    param([string]$Url, [string]$Destination)
    Write-Host "    downloading $Url"
    $previous = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'   # Invoke-WebRequest is many times faster without the progress bar
    try {
        Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -Headers @{ 'User-Agent' = 'corpiphone-sandbox-setup' }
    }
    finally {
        $ProgressPreference = $previous
    }
    return $Destination
}

function Install-PowerShell7 {
    if (Test-Path $pwshExe) {
        $existing = (& $pwshExe -NoProfile -Command '$PSVersionTable.PSVersion.ToString()')
        Write-Host "    PowerShell $existing already installed"
        return
    }
    $asset = Get-LatestReleaseAsset -Repo 'PowerShell/PowerShell' -TagPattern '^v7\.4\.\d+$' -AssetPattern '^PowerShell-7\.4\.\d+-win-x64\.msi$'
    if (-not $asset) {
        $asset = [PSCustomObject]@{
            Tag  = "v$PwshFallbackVersion"
            Name = "PowerShell-$PwshFallbackVersion-win-x64.msi"
            Url  = "https://github.com/PowerShell/PowerShell/releases/download/v$PwshFallbackVersion/PowerShell-$PwshFallbackVersion-win-x64.msi"
        }
    }
    Write-Host "    installing PowerShell $($asset.Tag)"
    $msi = Save-Download -Url $asset.Url -Destination (Join-Path $downloadDir $asset.Name)
    $proc = Start-Process -FilePath 'msiexec.exe' -Wait -PassThru -ArgumentList @(
        '/i', "`"$msi`"", '/quiet', '/norestart',
        'ADD_PATH=1', 'REGISTER_MANIFEST=1', 'ENABLE_PSREMOTING=0', 'ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=0'
    )
    if ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 3010) {
        throw "PowerShell 7 MSI install failed with exit code $($proc.ExitCode)."
    }
    if (-not (Test-Path $pwshExe)) { throw "PowerShell 7 was installed but $pwshExe was not found." }
    Write-Host "    PowerShell 7 installed at $pwshExe" -ForegroundColor Green
}

function Install-GitForWindows {
    if (Test-Path $gitExe) {
        Write-Host "    Git already installed: $(& $gitExe --version)"
        return
    }
    $asset = Get-LatestReleaseAsset -Repo 'git-for-windows/git' -TagPattern '^v2\.\d+\.\d+\.windows\.\d+$' -AssetPattern '^Git-2\.\d+\.\d+-64-bit\.exe$'
    if (-not $asset) {
        $ver = ($GitFallbackTag -replace '^v', '') -replace '\.windows\.\d+$', ''
        $asset = [PSCustomObject]@{
            Tag  = $GitFallbackTag
            Name = "Git-$ver-64-bit.exe"
            Url  = "https://github.com/git-for-windows/git/releases/download/$GitFallbackTag/Git-$ver-64-bit.exe"
        }
    }
    Write-Host "    installing Git $($asset.Tag)"
    $exe = Save-Download -Url $asset.Url -Destination (Join-Path $downloadDir $asset.Name)
    # Inno Setup silent flags. Default components include Git Credential Manager, which handles the GitHub sign-in.
    $proc = Start-Process -FilePath $exe -Wait -PassThru -ArgumentList @('/VERYSILENT', '/NORESTART', '/NOCANCEL', '/SP-', '/CLOSEAPPLICATIONS', '/RESTARTAPPLICATIONS')
    if ($proc.ExitCode -ne 0) { throw "Git for Windows install failed with exit code $($proc.ExitCode)." }
    if (-not (Test-Path $gitExe)) { throw "Git was installed but $gitExe was not found." }
    Write-Host "    $(& $gitExe --version) installed" -ForegroundColor Green
}

function Get-RepoCheckout {
    if (Test-Path (Join-Path $RepoDir 'README.md')) {
        Write-Host "    repo already present at $RepoDir"
        return
    }
    New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

    if ($GitHubToken) {
        # Zip download with a token: no Git needed, works for private repos, nothing is written to disk but the zip.
        $owner, $name = ($RepoUrl -replace '^https://github\.com/', '' -replace '\.git$', '').Split('/')
        $zip = Join-Path $downloadDir "$name.zip"
        Write-Host "    downloading $owner/$name@$RepoBranch as zip"
        $previous = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Invoke-WebRequest -Uri "https://api.github.com/repos/$owner/$name/zipball/$RepoBranch" -OutFile $zip -UseBasicParsing `
                -Headers @{ Authorization = "Bearer $GitHubToken"; 'User-Agent' = 'corpiphone-sandbox-setup'; Accept = 'application/vnd.github+json' }
        }
        finally { $ProgressPreference = $previous }
        $extract = Join-Path $downloadDir 'repo'
        Expand-Archive -Path $zip -DestinationPath $extract -Force
        $inner = Get-ChildItem -Path $extract -Directory | Select-Object -First 1
        Move-Item -Path $inner.FullName -Destination $RepoDir
        Remove-Item -Path $zip -Force
    }
    else {
        Install-GitForWindows
        Write-Host "    cloning $RepoUrl ($RepoBranch). If the repo is private, Git Credential Manager opens a browser sign-in."
        & $gitExe clone --branch $RepoBranch --single-branch $RepoUrl $RepoDir
        if ($LASTEXITCODE -ne 0) { throw "git clone failed with exit code $LASTEXITCODE. Check the branch name and your access to the repo." }
    }
    # Files that came from the internet are marked and PowerShell 7's RemoteSigned policy would refuse the deploy scripts.
    Get-ChildItem -Path $RepoDir -Recurse -File | Unblock-File
    Write-Host "    repo ready at $RepoDir" -ForegroundColor Green
}

function Install-ModulesInPowerShell7 {
    # Modules must be installed BY PowerShell 7 so they land in Documents\PowerShell\Modules, which 5.1 does not share.
    $inner = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
}
if ((Get-PSRepository -Name PSGallery).InstallationPolicy -ne 'Trusted') {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
}

$graphVersion = '__GRAPH_VERSION__'
$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Applications',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.PersonalContacts'
)
foreach ($m in $graphModules) {
    $have = Get-Module -ListAvailable -Name $m | Where-Object { $_.Version.ToString() -eq $graphVersion }
    if ($have) { Write-Host "    exists  $m $graphVersion"; continue }
    Write-Host "    install $m $graphVersion"
    Install-Module -Name $m -RequiredVersion $graphVersion -Scope CurrentUser -Force -AllowClobber
}

$others = @(
    @{ Name = 'Az.Accounts' },
    @{ Name = 'Az.Resources' },
    @{ Name = 'Az.Automation' },
    @{ Name = 'Az.Monitor' },
    @{ Name = 'ExchangeOnlineManagement'; MinimumVersion = '3.0.0' },
    @{ Name = 'Pester';                   MinimumVersion = '5.5.0' },
    @{ Name = 'PSScriptAnalyzer' }
)
foreach ($spec in $others) {
    $have = Get-Module -ListAvailable -Name $spec.Name
    if ($spec.ContainsKey('MinimumVersion')) { $have = $have | Where-Object { $_.Version -ge [version]$spec.MinimumVersion } }
    if ($have) { Write-Host "    exists  $($spec.Name) $(($have | Sort-Object Version -Descending | Select-Object -First 1).Version)"; continue }
    Write-Host "    install $($spec.Name)"
    $installArgs = @{ Name = $spec.Name; Scope = 'CurrentUser'; Force = $true; AllowClobber = $true }
    if ($spec.ContainsKey('MinimumVersion')) { $installArgs.MinimumVersion = $spec.MinimumVersion }
    if ($spec.Name -eq 'Pester') { $installArgs.SkipPublisherCheck = $true }
    Install-Module @installArgs
}

# Windows Sandbox has no cached Microsoft account, so broker (WAM) sign-in cannot work. Turn it off where the
# modules allow it; the deploy scripts also fall back to device code on their own and accept -UseDeviceCode.
Import-Module Az.Accounts
try { Update-AzConfig -EnableLoginByWam $false -Scope CurrentUser | Out-Null; Write-Host '    Az: WAM sign-in disabled' }
catch { Write-Host "    Az: could not disable WAM ($($_.Exception.Message)); use -UseDeviceCode on the deploy scripts" -ForegroundColor Yellow }
Import-Module Microsoft.Graph.Authentication
try { Set-MgGraphOption -EnableLoginByWAM $false; Write-Host '    Graph: WAM sign-in disabled' }
catch { Write-Host "    Graph: could not disable WAM ($($_.Exception.Message)); use -UseDeviceCode on the deploy scripts" -ForegroundColor Yellow }

Write-Host ''
Write-Host '    Installed module versions:'
foreach ($m in $graphModules + @('Az.Accounts', 'Az.Resources', 'Az.Automation', 'Az.Monitor', 'ExchangeOnlineManagement', 'Pester', 'PSScriptAnalyzer')) {
    $v = (Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending | Select-Object -First 1).Version
    Write-Host ("      {0,-36} {1}" -f $m, $v)
}
'@
    $inner = $inner.Replace('__GRAPH_VERSION__', $GraphModuleVersion)
    $innerPath = Join-Path $downloadDir 'install-modules.ps1'
    Set-Content -Path $innerPath -Value $inner -Encoding UTF8
    & $pwshExe -NoProfile -ExecutionPolicy Bypass -File $innerPath
    if ($LASTEXITCODE -ne 0) { throw "Module installation inside PowerShell 7 failed with exit code $LASTEXITCODE." }
}

function Invoke-RepoTests {
    $inner = @'
Set-Location -LiteralPath '__REPO_DIR__'
$cfg = New-PesterConfiguration
$cfg.Run.Path = './tests'
$cfg.Run.Exit = $true
$cfg.Output.Verbosity = 'Normal'
Invoke-Pester -Configuration $cfg
'@
    $innerPath = Join-Path $downloadDir 'run-tests.ps1'
    Set-Content -Path $innerPath -Value $inner.Replace('__REPO_DIR__', $RepoDir) -Encoding UTF8
    & $pwshExe -NoProfile -ExecutionPolicy Bypass -File $innerPath
    if ($LASTEXITCODE -ne 0) {
        Write-Host "    Pester reported failures (exit code $LASTEXITCODE). Do not run the deploy scripts until the suite is green." -ForegroundColor Red
    }
    else {
        Write-Host '    all tests passed' -ForegroundColor Green
    }
}

function Show-NextSteps {
    $lines = @(
        '',
        '================================================================================',
        ' Sandbox is ready. Work in the PowerShell 7 window that just opened (prompt starts with PS).',
        " Repo: $RepoDir",
        '',
        ' Every deploy script accepts -UseDeviceCode. Use it in Sandbox: it prints a code you enter at',
        ' https://microsoft.com/devicelogin in a browser on your HOST machine, signed in as the admin.',
        '',
        ' Step 1  Groups, Automation account, identity, Graph permissions, runtime, runbooks, schedules, alert:',
        '   ./deploy/Deploy-Automation.ps1 -UseDeviceCode -SubscriptionId <sub id> -ResourceGroupName rg-corpiphone-sync `',
        '       -Location eastus2 -AutomationAccountName aa-corpiphone-sync -AlertEmail you@contoso.com',
        '',
        ' Step 2  Exchange scope (use the two ids step 1 printed):',
        '   ./deploy/New-ExchangeAppScope.ps1 -UseDeviceCode -ApplicationId <app id> -ObjectId <object id> `',
        '       -StampedUserPrincipalName someone@contoso.com -UnstampedUserPrincipalName other@contoso.com',
        '',
        ' Step 4  Intune policies (pilot group first):',
        '   ./deploy/New-OutlookContactSyncPolicies.ps1 -UseDeviceCode -GroupName SG-CorpiPhone-Pilot',
        '',
        ' Steps 3, 5, 6 run in the Azure portal (start the runbooks, enable the schedules); see README.md.',
        ' Everything in this Sandbox disappears when you close it. Nothing here needs to survive.',
        '================================================================================'
    )
    $lines | ForEach-Object { Write-Host $_ -ForegroundColor Green }
}

# ============================================================
# MAIN LOGIC
# ============================================================
$setupOk = $false
try {
    New-Item -ItemType Directory -Path $downloadDir -Force | Out-Null
    Write-Host "CorpiPhone sandbox setup starting $(Get-Date -Format 'u')" -ForegroundColor Cyan
    Write-Host "Windows PowerShell $($PSVersionTable.PSVersion) on $([Environment]::OSVersion.VersionString)"

    Write-Step '1/4 PowerShell 7.4'
    Install-PowerShell7

    Write-Step '2/4 Repository'
    Get-RepoCheckout

    Write-Step "3/4 Modules (inside PowerShell 7, Microsoft.Graph pinned to $GraphModuleVersion). This is the slow part."
    Install-ModulesInPowerShell7

    if ($RunTests) {
        Write-Step '4/4 Pester tests'
        Invoke-RepoTests
    }
    else {
        Write-Step '4/4 Pester tests skipped ($RunTests = $false)'
    }
    $setupOk = $true
}
catch {
    Write-Host ''
    Write-Host "SETUP FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "At line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" -ForegroundColor Red
    Write-Host 'Fix the cause and paste the script again; every step skips what already exists.' -ForegroundColor Yellow
}

# ============================================================
# OUTPUT
# ============================================================
if ($setupOk) {
    Show-NextSteps
    if ($OpenWindow) {
        Start-Process -FilePath $pwshExe -ArgumentList @('-NoLogo', '-NoExit', '-ExecutionPolicy', 'RemoteSigned', '-WorkingDirectory', "`"$RepoDir`"")
    }
}

# ============================================================
# CLEANUP
# ============================================================
if (Test-Path $downloadDir) {
    Remove-Item -Path $downloadDir -Recurse -Force -ErrorAction SilentlyContinue
}
