# Company Directory in iPhone Contacts

Every corporate iPhone shows an up-to-date company directory, including each coworker's corporate mobile number, in the native iOS Contacts app, so caller ID resolves. Built only on Microsoft 365, Intune, Exchange Online, Azure Automation and PowerShell. No server, no SaaS, no certificates, no secrets.

## Architecture

Entra ID is the single source of truth and Intune's primary-user assignment is the single source of who is in scope. Runbook 1 (`Sync-IntuneToEntra`) reads corporate iOS devices from Intune, stamps their primary users with `extensionAttribute1 = CorpiPhone`, and writes each phone's number to the user's `mobilePhone`. A dynamic security group fills itself from that stamp, Intune assigns the Outlook policies to it, and an Exchange RBAC for Applications scope on the same attribute (`CustomAttribute1`) is the only thing that lets the managed identity touch mailboxes. Runbook 2 (`Sync-DirectoryToMailboxes`) writes one tagged "Company Directory" contact per coworker into every group member's default Contacts folder through Graph and keeps them current. Outlook for iOS, configured by an App Protection Policy and an App Configuration Policy, syncs that folder into native Contacts.

```
Intune managedDevice (primary user, phoneNumber)
   -> Runbook 1 (daily 05:00) -> Entra user.extensionAttribute1 = CorpiPhone, user.mobilePhone
   -> dynamic group SG-CorpiPhone-Users  -+-> Intune APP + ACP (Save Contacts locked on) -> Outlook for iOS
                                          +-> Exchange RBAC scope CustomAttribute1 -eq 'CorpiPhone'
   -> Runbook 2 (Monday 08:00) -> each member's default Contacts folder -> Outlook for iOS -> native iOS Contacts
```

## Repository layout

| Path | Purpose |
| --- | --- |
| `runbooks/Sync-IntuneToEntra.ps1` | Runbook 1: scope stamp and phone number. Azure Automation, PowerShell 7.4, managed identity. |
| `runbooks/Sync-DirectoryToMailboxes.ps1` | Runbook 2: directory to mailbox contacts. Same runtime. |
| `deploy/Deploy-Automation.ps1` | One-time: groups, Automation account, identity, Graph permissions, 7.4 runtime environment with pinned modules, runbooks, schedules, alert. |
| `deploy/New-ExchangeAppScope.ps1` | One-time, Exchange Administrator: service principal, management scope, scoped `Application Contacts.ReadWrite`, authorization test, Entra permission check. |
| `deploy/New-OutlookContactSyncPolicies.ps1` | One-time, Intune Administrator: iOS APP and Outlook ACP, assigned to a group. |
| `deploy/sandbox/` | Windows Sandbox bootstrap (paste-and-go PowerShell 5.1 script) and a `.wsb` that runs it at logon. |
| `docs/Deployment-Guide-for-First-Timers.md` | Step-by-step deployment guide for someone who has never used Azure, Intune or PowerShell. Every click and command spelled out. |
| `tests/` | Pester tests (pure functions, full runbook flows against an in-memory fake Graph, PSScriptAnalyzer). |
| `PSScriptAnalyzerSettings.psd1` | Lint settings used by the tests. |

Everything tenant-specific (ids, names, addresses) is a script parameter. Nothing in the repo identifies a tenant.

## Build order

Follow the Deployment guide for the hands-on steps. The order is dictated by the data flow:

1. `deploy/Deploy-Automation.ps1` creates the groups, the Automation account with its system-assigned managed identity, the four Graph application permissions, the PowerShell 7.4 runtime environment with `Microsoft.Graph` pinned to 2.40.0, imports and publishes both runbooks, creates the schedules (disabled) and the alert rule. It prints the identity's application id and object id.
2. `deploy/New-ExchangeAppScope.ps1` registers the identity in Exchange Online, creates the `CorpiPhoneMailboxes` scope on `CustomAttribute1 -eq 'CorpiPhone'`, assigns `Application Contacts.ReadWrite` restricted to it, runs `Test-ServicePrincipalAuthorization`, and checks Entra for a `Contacts.*` grant that must not exist.
3. Runbook 1 `Sync-IntuneToEntra`, first with `WhatIf = true`, then live. The stamp and the dynamic group now exist.
4. `deploy/New-OutlookContactSyncPolicies.ps1` creates the APP and ACP and assigns them to a pilot group, later to `SG-CorpiPhone-Users`.
5. Runbook 2 `Sync-DirectoryToMailboxes`, first with `WhatIf = true` and `UserPrincipalName` for one pilot user, then live.
6. Enable the schedules: Runbook 1 daily 05:00, Runbook 2 weekly Monday 08:00 (America/New_York). The three-hour gap lets Exchange's permission cache catch up for users stamped that morning.

### Running from Windows Sandbox

Windows Sandbox starts clean every time (PowerShell 5.1 only, no modules, no cached account). `deploy/sandbox/Initialize-SandboxEnvironment.ps1` turns it into a ready admin workstation: paste the whole file into the Windows PowerShell window the Sandbox opens. It installs PowerShell 7.4 and Git, clones this repo, installs the exact module set into PowerShell 7 (Graph pinned to the runtime environment's version), disables broker sign-in, runs the Pester suite and opens a PowerShell 7 window in the repo. About 10 to 15 minutes. Edit the configuration block at the top first if the branch name or repo URL changed.

To automate it, copy the script into a host folder, point `deploy/sandbox/CorpiPhone-Sync.wsb` at that folder and double-click the `.wsb`; the Sandbox runs the setup at logon.

Inside Sandbox, run every deploy script with `-UseDeviceCode`. There is no cached Microsoft account for the broker flow, so the scripts show a code you enter at microsoft.com/devicelogin in a browser on the host. Without the switch they try the interactive flow first and fall back to device code when it fails.

### Test locally before touching the tenant

```powershell
Install-Module Pester, PSScriptAnalyzer -Scope CurrentUser
Invoke-Pester ./tests
```

The suite needs no modules from the Graph SDK and no network. It runs the runbooks end to end against a fake Graph (`tests/FakeGraph.ps1`), checks the contact body builder and hash against three sample users, and lints every script.

## Permissions

| Item | Setting |
| --- | --- |
| Runbook identity | System-assigned managed identity on the Automation account. `Connect-MgGraph -Identity`. No app registration, certificate or secret. |
| Graph application permissions | `DeviceManagementManagedDevices.Read.All`, `User.ReadWrite.All`, `GroupMember.Read.All`, `RoleManagement.Read.Directory`. |
| Contacts access | Exchange RBAC for Applications only: role `Application Contacts.ReadWrite`, scope `CustomAttribute1 -eq 'CorpiPhone'`. **Never consent `Contacts.ReadWrite` in Entra.** Entra and Exchange grants are a union; an unscoped Entra grant would silently remove the mailbox scoping. Both deploy scripts check for this and warn. |
| Control group | `SG-CorpiPhone-Users`, dynamic, rule `(user.extensionAttribute1 -eq "CorpiPhone")`. Filled by Runbook 1, never edited by hand. Needs Entra ID P1. |
| Exclusion group | `SG-Directory-Exclude`, static. Members receive contacts but are not listed, and Runbook 1 does not write their number. |
| Schedules | Automation schedules linked with `WhatIf = false` and the group names as parameters. |
| Logging | Job output stream (one line per change plus a summary object) and warning stream (throttling retries). Kept in job history. |
| Alerting | Azure Monitor metric alert on `TotalJob` with `Status` in Failed, Suspended, emailing the address passed to the deploy script. |

Things this deliberately does not touch: MFA phone registration (the authentication methods API is never called), users' personal contacts (only contacts carrying the `DirectorySourceId` extended property are read, modified or deleted), and Outlook's own GAL behavior.

## Runbook parameters

Azure Automation passes bools and strings, not switches. Both runbooks default to `WhatIf = true`.

**Sync-IntuneToEntra**

| Parameter | Default | Meaning |
| --- | --- | --- |
| `WhatIf` | `true` | Log without writing. |
| `ExclusionGroupName` | `SG-Directory-Exclude` | Members keep the stamp but their number is not written. Empty disables. |
| `DefaultCountryCode` | `+1` | Prefix for 10-digit numbers. |
| `ScopeAttributeValue` | `CorpiPhone` | Value written to `extensionAttribute1`. |

**Sync-DirectoryToMailboxes**

| Parameter | Default | Meaning |
| --- | --- | --- |
| `WhatIf` | `true` | Log without writing. |
| `GroupName` | `SG-CorpiPhone-Users` | Directory and target list. |
| `ExclusionGroupName` | `SG-Directory-Exclude` | Removed from the directory, still receive contacts. |
| `UserPrincipalName` | empty | Process only this member's mailbox (onboarding). |
| `FullRebuild` | `false` | Rewrite every owned contact regardless of the stored hash. |

Both end with a summary object and throw if any error occurred, so the job shows Failed and the alert fires.

## Pilot checklist

1. Groups, Automation account, identity, the four Graph permissions and the Exchange scope created. Confirm `Contacts.ReadWrite` is absent from the identity's Entra permissions (Enterprise applications > the identity > Permissions).
2. Runbook 1 with `WhatIf`: the stamp list matches the corporate iPhone primary users in Intune exactly. Fix primary users in Intune where wrong. Run live. Dynamic group fills within minutes.
3. `Test-ServicePrincipalAuthorization` (re-run `New-ExchangeAppScope.ps1 -TestOnly`) shows the Contacts role in scope for a stamped user and not for an unstamped one.
4. APP and ACP assigned to the pilot group. Users tap Allow; an existing Outlook contact appears in native Contacts.
5. Runbook 2 with `WhatIf` for one user, review, run live. Company Directory contacts appear in Outlook and native Contacts; a test call from a coworker resolves caller ID. Then run for the full group.
6. Leaver test: change a test phone's primary user or retire it, run Runbook 1, stamp clears and user leaves the group; run Runbook 2, their contact is gone from other phones.
7. Edit test: a user edits a directory contact; next run restores it, nothing else changes.
8. Schedules enabled, alert rule on, three clean weeks.

Exit criteria: Runbook 1's stamped set matches Intune's corporate iPhone primary users exactly, zero unexplained skips, each member's mailbox holds group size minus one runbook-owned contacts, users confirm caller ID, both schedules have run clean for three weeks.

## Onboarding

Enroll the phone in Intune with the user as primary user. That is the whole action. Runbook 1 stamps them at the next 05:00 run, the dynamic group and the Outlook policies follow within minutes, and Runbook 2 fills their mailbox on Monday.

To skip the wait: after Runbook 1 has run (and at least 30 minutes have passed for the Exchange scope to catch up), start `Sync-DirectoryToMailboxes` with `WhatIf = false` and `UserPrincipalName = <the user>`. About 80 creates, under a minute.

## Phone handover

Change the device's primary user in Intune. Runbook 1 moves the stamp on its next run: the old holder is unstamped (and drops out of the group, the policies and the Exchange scope), the new holder is stamped. Runbook 2 removes the old holder's card from everyone else and adds the new one on Monday, or sooner if you start it by hand.

## User notice (send when widening to the full group)

> Your corporate iPhone will soon show the company directory in the Contacts app, so you see a coworker's name when they call. Outlook does this for you. The next time you open Outlook it will tell you that your administrator turned on contact sync, and iOS will ask whether Outlook may access your Contacts. Tap **Allow**. Nothing is uploaded from your phone; the directory flows one way, from the company directory into Outlook and then into Contacts. If you tapped **Don't Allow** by mistake, open Outlook, go to Settings, your account, and use the **Open Settings** link under Save Contacts to turn it on.

## Rollback

- Stop contacts arriving: disable the two schedules.
- Remove the directory from phones: unassign the ACP, or set Save Contacts to off in it. Outlook stops syncing and removes the native entries on the next policy refresh.
- Remove the contacts from mailboxes while keeping scope: nest `SG-CorpiPhone-Users` inside `SG-Directory-Exclude` (exclusions are read transitively), start `Sync-DirectoryToMailboxes` live once. With everyone excluded the directory is empty, so the runbook deletes every tagged contact and touches nothing else. Then remove the nesting.
- Remove the stamp: clear `extensionAttribute1` on the users. The dynamic group empties and the Exchange scope matches nobody.
- Full teardown: remove the Exchange role assignment and scope, delete the Automation account, delete the two groups.

## Troubleshooting

**Phone has no primary user.** Runbook 1 logs `Device '<name>' serial <serial> has no primary user and is ignored` and counts it in `DevicesWithoutPrimaryUser`. Set the primary user in Intune (Devices > the device > Properties > Primary user). The user is stamped on the next run.

**Device reports no phone number.** Skip reason `NoPhoneNumber`. Intune collects hardware inventory about every 7 days, so a fresh enrollment may show its primary user before its number; the stamp lands first and the number follows. Some carriers and eSIM profiles never report a number. The user stays in the directory without a mobile number. Set `mobilePhone` by hand if you need it sooner.

**Number looks wrong or is rejected.** Skip reason `InvalidPhoneNumber` with the raw value. Accepted: 10 digits (gets `DefaultCountryCode`), 11 digits starting with 1, or an explicit `+` with 8 to 15 digits. Anything else is logged and skipped.

**User skipped as privileged.** Skip reason `PrivilegedRole`. Writing `mobilePhone` on a user who holds an active directory role requires a privileged caller, so the runbook never tries. The stamp is still set and the user is still in the directory. Set their number by hand, or exclude them from the directory via `SG-Directory-Exclude`. Role-assignable groups are expanded, so members of a group that holds a role are also skipped.

**Runbook 1 fails with "foreign value in extensionAttribute1".** Someone else uses that attribute. The runbook owns `extensionAttribute1` and refuses to overwrite other values; nothing was changed. Either clear the attribute on the listed users or move the whole design to a free attribute: change the constant at the top of `Sync-IntuneToEntra.ps1`, the dynamic group rule, `-ScopeAttributeName` in both deploy scripts and the Exchange scope filter.

**Dynamic group not updating.** Check the rule is exactly `(user.extensionAttribute1 -eq "CorpiPhone")` and processing state is On. Dynamic membership needs Entra ID P1. Processing normally finishes in minutes at this size but can take up to 24 hours in large tenants. Check Entra > Groups > the group > Dynamic membership rules > Validate Rules against a stamped user.

**Exchange 403 after a fresh stamp (scope lag).** Runbook 2 logs `AccessDenied : <upn> : 403` and counts the mailbox under `MailboxesSkipped`, not as an error. Exchange caches RBAC results for 30 minutes to 2 hours. Wait and re-run for that user with `UserPrincipalName`. If it persists past 2 hours: run `New-ExchangeAppScope.ps1 -TestOnly`; if `InScope` is False for a stamped user, confirm the user's `CustomAttribute1` in Exchange (`Get-Mailbox <upn> | Select CustomAttribute1`) shows `CorpiPhone` and that the role assignment carries the `CorpiPhoneMailboxes` scope.

**Every mailbox is 403, or the identity can read mailboxes it should not.** Check for a `Contacts.*` permission on the identity in Entra. If present, the Exchange scope is being bypassed (union of grants). Remove it. If absent and everything is 403, the Exchange service principal, scope or role assignment is missing; re-run `New-ExchangeAppScope.ps1`.

**No mailbox.** Skip reason `NoMailbox` (404). The member has no Exchange Online mailbox or REST is not enabled for it. Nothing to fix unless the user should have a mailbox.

**Outlook contacts not appearing in the native app.** In order:
1. APP: Data protection > Data transfer > *Sync policy managed app data with native apps or add-ins* must be Allow on the policy the device receives. Check Intune > Apps > App protection policies > the policy > Properties, and the device's App protection status report.
2. ACP: *Save Contacts* = Yes on the managed-apps ACP assigned to the user. Check the Outlook settings page on the phone under the account: Save Contacts should show on and locked.
3. iOS Contacts permission denied: iOS Settings > Outlook > Contacts must be on. Outlook also shows an *Open Settings* link under the account when denied.
4. Device restriction profile: *Allow managed apps to write contacts to unmanaged contacts accounts* must be Allow (or not configured) in any iOS device restrictions profile assigned to the phone. With managed open-in blocked and this not allowed, Outlook fails silently.
5. Policies apply on the next Outlook check-in; force one by closing and reopening Outlook or by syncing the device from Company Portal.

**Duplicate contacts.** Two kinds. Duplicates among runbook-owned contacts (same `DirectorySourceId`) are deleted automatically on the next run; the log shows `DELETE ... (duplicate)`. Duplicates between a runbook contact and a personal contact the user created themselves are expected and untouched; the user can delete their own copy. If the native app shows a coworker twice with the same data, check that Outlook is not also syncing contacts from a second account.

**Module import failures in Automation.** The runtime environment must be PowerShell 7.4 (7.2 is .NET 6 and current Graph SDK modules fail to load there) and every `Microsoft.Graph.*` package must show provisioning state Succeeded at the same version. 2.40.0 is known good; 2.41.0 and 2.41.1 fail. Never update the pinned version in place: create a second runtime environment with the candidate version, point a test copy of the runbook at it, and switch only when it runs clean. Import order matters: `Microsoft.Graph.Authentication` first.

**429 throttling.** Both runbooks honor `Retry-After` with exponential backoff, and Runbook 2 retries only the throttled sub-requests of a batch. Retries appear in the job's warning stream. A job that ends with `still throttled after 6 rounds` on some requests reports them as errors and shows Failed; re-run it, the reconcile is idempotent. Sustained throttling at this scale (about 320 requests on a first run) is unusual and usually means another workload is hammering the same mailboxes.

**Job shows Failed with a summary object present.** The summary's `Errors` count is non-zero and the thrown message lists each error. Every other change in the run was applied. Fix the cause and re-run; both runbooks are safe to repeat.

## Privacy note for change control

Every corporate iPhone number becomes visible to every other corporate iPhone user in their Contacts app. The number is already in the GAL once Runbook 1 runs, so the mailbox step adds reach, not new exposure. Any app on the phone with Contacts permission can read synced contacts, which is acceptable on supervised corporate devices and should be stated in the change record. Selective wipe and retire remove the Outlook-synced contacts along with Outlook's data.
