# Company Directory on iPhones: Deployment Guide for First-Timers

Nov 3, 2026. Living copy with comments: https://claude.ai/artifact/QUJdMdk61nuxqawLNvQZ2x

## What you are about to do

When you finish, every company iPhone shows the whole company directory in its Contacts app, so when a coworker calls, their name appears instead of a number. You will not install a server or buy anything. You will connect pieces of Microsoft 365 you already own and let two small scheduled scripts keep everything up to date.

In plain terms, this is what gets built:

1. A script runs every morning. It asks Intune "which iPhones are company phones, and who uses each one?" It then marks those people with an invisible tag in the company directory and writes each phone's number into that person's profile.
2. A security group fills itself automatically with everyone who carries the tag. Nobody ever adds people to it by hand.
3. A second script runs every Monday. For every tagged person, it writes one contact card per coworker into their mailbox, the same place Outlook keeps their personal contacts. It updates cards when a title or number changes and removes cards when someone leaves.
4. Two Intune policies tell the Outlook app on the iPhone to copy those contacts into the phone's own Contacts app. Each user taps **Allow** once when the phone asks.

You only type commands in three places, and this guide gives you every command ready to paste. Everything else is clicking in web portals, and every click is spelled out.

**Time needed.** About 2 to 3 hours of hands-on work for the one-time setup, spread over a day because some steps need waiting. Then a quiet pilot week with a handful of phones, then one short step to roll out to everyone.

**What you need to be comfortable with.** Copying and pasting, reading text on a screen, and following instructions in order. You do not need to understand PowerShell or Azure. Where a word is unfamiliar, the next section explains it.

**One rule.** Do the steps in order and do not skip the "check" lines. Each step builds on the one before it, and the checks tell you early if something went wrong, when it is still easy to fix.

## Words you will see

You do not need to memorize these. Come back here when a word stops you.

| Word | What it means here |
| --- | --- |
| Microsoft 365 | The bundle your company pays for: email, Office apps, and the admin tools below. |
| Entra ID | Microsoft's directory of your users and groups. Formerly called Azure AD. Think of it as the company phone book that everything else trusts. Admin site: entra.microsoft.com |
| Intune | The tool that manages company phones and laptops. It knows which iPhones are company-owned and who uses each. Admin site: intune.microsoft.com |
| Exchange Online | The mail service behind Outlook. Each person's contacts live in their mailbox there. |
| Azure | Microsoft's cloud for running things. You will create one small item in it. Site: portal.azure.com |
| Subscription | A billing container in Azure. Everything you create in Azure lives inside one. You will need its ID, a long code. |
| Resource group | A folder inside a subscription. You will make one to hold the automation account. |
| Automation account | The Azure service that runs scripts on a schedule for you, like a very reliable alarm clock that also does the work. Free for this amount of use. |
| Runbook | One script stored inside the automation account. This project has two. |
| Managed identity | The automation account's own login. It lets the runbooks sign in without a password you would have to store. You grant it permissions once. |
| Dynamic group | A security group that fills itself from a rule, here "everyone tagged as a company iPhone user". |
| extensionAttribute1 | A spare text field on every user in Entra. The morning runbook writes the word CorpiPhone there as the tag. The group rule and the mailbox permission both look at it. |
| Graph | Microsoft's programming interface to all of the above. The scripts talk to it. You will see the word in permission names. |
| PowerShell | The command-line tool the scripts are written in. Windows has an old version (5.1) built in. The scripts need the newer version 7 (7.4 or later), which the setup script installs for you. |
| Windows Sandbox | A throwaway copy of Windows that starts clean and is wiped when you close it. Safe place to run setup scripts without changing your real PC. |
| UPN | User principal name. A person's sign-in address, usually the same as their email, like jane.doe@contoso.com. |
| App Protection Policy (APP) | An Intune rule for what an app may do with company data. Here it allows Outlook to write contacts to the phone. |
| App Configuration Policy (ACP) | An Intune rule that sets an app's settings for the user. Here it turns on Outlook's Save Contacts and locks it on. |
| WhatIf | A setting on each runbook. **true** means "only show me what you would do". **false** means "really do it". You always run with true first. |
| Device code | A way to sign in when a script cannot open a browser itself. It shows a short code; you type it into a web page on your own PC and sign in there. |

## Before you start

Get these four things sorted before you type anything. Each one takes a few minutes.

### 1. The right admin account

You need one admin account that has all of these roles. If you are the only IT person, you probably have them. If not, ask whoever set up your Microsoft 365.

| Role | Needed for |
| --- | --- |
| Global Administrator | Granting the automation account its permissions, creating groups |
| Exchange Administrator | The mailbox permission step |
| Intune Administrator | The Outlook policies |
| Owner or Contributor on an Azure subscription | Creating the automation account |

**How to check your roles.** Go to entra.microsoft.com, sign in, click **Users**, find your own account, open it, and click **Assigned roles** in the left menu. Global Administrator covers the first three rows above. For the Azure one, go to portal.azure.com, search **Subscriptions** at the top, open your subscription, click **Access control (IAM)**, then **View my access**.

If your company uses PIM (Privileged Identity Management), activate the roles before you start. They expire after a few hours, so activate for 8 hours.

### 2. Things to write down

Open a notepad and collect these. You will paste them into commands later.

- [ ] **Subscription ID.** portal.azure.com, search **Subscriptions**, click your subscription, copy the **Subscription ID** (looks like `1a2b3c4d-....`).
- [ ] **An Azure region**, for example `eastus2` or `westeurope`. Pick the one closest to you. Type it exactly, lowercase, no spaces.
- [ ] **An email address** that should receive an alert when a nightly run fails. Your own is fine to start.
- [ ] **The sign-in address (UPN) of one person who has a company iPhone**, for testing.
- [ ] **The UPN of one person who does not have a company iPhone**, for testing.
- [ ] **Names of 5 to 10 iPhone users** willing to be the pilot group.

### 3. Two quick checks in your tenant

**Check A: nobody already uses extensionAttribute1.** The scripts use this spare field as the tag and will refuse to run if it already holds other values. The Entra portal does not show this field, so check it with one command. Do this right after Step 0 (the Sandbox setup), before Step 1. Paste this into the PowerShell 7 window as one line and sign in with the code it shows:

```powershell
Connect-MgGraph -Scopes User.Read.All -UseDeviceCode -NoWelcome; Get-MgUser -All -Property UserPrincipalName,OnPremisesExtensionAttributes | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } | Select-Object UserPrincipalName, @{ n = 'Attribute1'; e = { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 } }
```

No output means nobody uses it, which is what you want. If names appear, your company uses the field for something else (cost centers, for example): stop and read the "If something goes wrong" section about choosing another attribute before continuing. (You can also see the same value per person in the Exchange admin center: **Recipients**, **Mailboxes**, the person, **Others**, **Custom attributes**, attribute 1.)

**Check B: every company iPhone has a primary user.** In intune.microsoft.com, go to **Devices**, then **All devices**. Click **Filter**, choose OS = iOS/iPadOS and Ownership = Corporate. Look at the **Primary user** column. Any blank is a phone the scripts will ignore. Intune can only change the primary user of Windows PCs. For an iPhone the primary user is whoever enrolled it, so a phone with the wrong or no primary user must be retired (or wiped) and enrolled again with the right person signing in. For Automated Device Enrollment phones, use an enrollment profile **with user affinity**. Do this for every blank (and every phone showing the wrong person) before Step 3.

### 4. A Windows PC with Windows Sandbox

You need a Windows 10 or 11 PC running the Pro, Enterprise or Education edition (Home edition cannot run Sandbox). To turn Sandbox on: press the Windows key, type **Turn Windows features on or off**, open it, tick **Windows Sandbox**, click OK, and restart when asked. If the box is greyed out, virtualization is turned off in your PC's BIOS; ask whoever manages your hardware, or use the "Option B" note in Step 0 to work on your own PC instead.

Also have the repository link ready. Your copy is at github.com/pavelkondratyukrepo/Iphonecontact-sync. Make sure you can open it in your browser while signed in to GitHub.

## How to read the commands in this guide

Every command in this guide is one line, even when it wraps on your screen. Copy the whole line, paste it into the PowerShell window, replace the parts in angle brackets, and press Enter.

```powershell
./deploy/Deploy-Automation.ps1 -UseDeviceCode -SubscriptionId <subscription id> -Location <region>
```

- `./deploy/Deploy-Automation.ps1` is the script to run. The `./` means "in the current folder". The PowerShell window the setup opens is already in the right folder.
- Each `-Word` is a setting, and the value for it comes right after. `-Location eastus2` means "set Location to eastus2".
- `<subscription id>` in angle brackets means **replace this, brackets included, with your own value**. So `-SubscriptionId <subscription id>` becomes `-SubscriptionId 1a2b3c4d-5e6f-7a8b-9c0d-1e2f3a4b5c6d`.
- A value with spaces goes in single quotes: `-AlertEmail 'it-alerts@contoso.com'` works with or without quotes, but `-GroupName 'SG-CorpiPhone-Pilot'` is safe either way.
- `-UseDeviceCode` has no value after it. It is an on/off switch, and you leave it on.

**Pasting.** Right-click inside the PowerShell window to paste, or press Ctrl+V. If the window seems to wait after you paste, press Enter once.

**The prompt.** Before you type, the window shows something like `PS C:\Users\WDAGUtilityAccount\Desktop\CorpiPhone\Iphonecontact-sync>`. That is PowerShell waiting for you. If it shows `>>` instead, it thinks your command is unfinished; press Ctrl+C, and paste again as one line.

**Red text** is an error. Do not keep going. Read the message, then look in "If something goes wrong" below. **Yellow text** is a warning you should read but that did not stop the script. **Cyan lines starting with a time** are the script telling you which step it is on.

### How device-code sign-in works

Inside Windows Sandbox there is no Microsoft account signed in, so the scripts cannot pop up the normal login window. Instead they print something like this:

```
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD1234 to authenticate.
```

When you see that:

1. On your **own PC** (not inside the Sandbox), open a browser and go to microsoft.com/devicelogin.
2. Type the code shown in the Sandbox window and click Next.
3. Sign in with your **admin account**, the one with the roles from "Before you start". Approve the MFA prompt.
4. The web page says you are signed in and can close the window. Go back to the Sandbox; the script continues on its own within a few seconds.

You will do this a few times per script, because each script talks to more than one Microsoft service and each needs its own sign-in. The code is different every time. It expires after about 15 minutes, so do it right away.

## Step 0: Set up your workstation (Windows Sandbox)

At the end of this step you have a PowerShell 7 window open in the project folder, with all tools installed and the project's self-tests passed. Allow 15 minutes, most of it waiting for downloads.

**Remember:** everything inside the Sandbox is erased when you close it. That is fine. Nothing you create here needs to survive; the real results live in Microsoft 365 and Azure. If you close the Sandbox halfway through the deployment, just repeat Step 0 and continue where you stopped. Every deploy script skips what already exists.

### 0.1 Get the setup script onto your clipboard

1. On your PC, open your browser and go to the repository: github.com/pavelkondratyukrepo/Iphonecontact-sync.
2. Click the folder **deploy**, then **sandbox**, then the file **Initialize-SandboxEnvironment.ps1**.
3. Near the top right of the file view there is a small **copy** icon (two overlapping squares) labelled "Copy raw file". Click it. The whole script is now on your clipboard. Do not open it in Word or Notepad first; copying straight from GitHub keeps it intact.

### 0.2 Open the Sandbox and paste

1. Press the Windows key, type **Windows Sandbox**, and open it. A new window appears that looks like a second, empty Windows desktop. Wait until the desktop fully loads (about 30 seconds).
2. Inside the Sandbox, click its Start button, type **PowerShell**, and open **Windows PowerShell** (the plain one, not "ISE" and not "x86"). A blue window opens.
3. Right-click anywhere inside that blue window. The script pastes and starts running immediately. If nothing happens, press Enter once.

The clipboard is shared between your PC and the Sandbox, so the paste works. If it does not, in the Sandbox window click the small **...** or the clipboard setting and make sure clipboard sharing is on, then copy again on your PC and paste again.

### 0.3 What you will see, and how long it takes

| Line on screen | What is happening | Typical wait |
| --- | --- | --- |
| `==> 1/4 PowerShell 7.6 LTS` | Downloading and installing PowerShell 7 | 1 to 2 minutes |
| `==> 2/4 Repository` | Installing Git, then downloading the project. A browser window may open inside the Sandbox asking you to sign in to GitHub. Sign in with your GitHub account and click Authorize. | 1 to 3 minutes |
| `==> 3/4 Modules` | Downloading about a dozen PowerShell modules. Lines like `install Az.Accounts` appear one by one. This is the slow part. | 5 to 10 minutes |
| `==> 4/4 Pester tests` | Running the project's self-tests. You want to see `Tests Passed: 117, Failed: 0` (the first number may be higher in a newer version; Failed must be 0). | 30 seconds |
| Green box `Sandbox is ready` | Done. A **second** PowerShell window opens, with a prompt ending in `Iphonecontact-sync>`. |  |

Use that second window for everything that follows. You can minimize the blue one; do not close it until you are done with the whole deployment.

**If it stops with red text** saying `SETUP FAILED`, read the line under it. The two common causes are no internet inside the Sandbox (check your PC is online, then close the Sandbox and start again) and a GitHub sign-in that was cancelled (paste the script again; it continues from the download).

### Option B: using your own PC instead of Sandbox

If you cannot use Sandbox, the same script works on your own Windows PC. Open **Windows PowerShell as administrator** (right-click, Run as administrator) and paste it. It installs PowerShell 7 and Git for real on that PC, and puts the project on your Desktop in a folder called CorpiPhone. Everything else in this guide is identical, except that you may be able to drop `-UseDeviceCode` from the commands, because a normal sign-in window can open on a real PC. Keeping it does no harm.

## Step 1: Create the groups, the automation account and its permissions

At the end of this step the two groups exist, the automation account exists with both runbooks loaded, its schedules exist but are switched off, and you have written down two IDs. Allow 20 to 30 minutes, mostly waiting while Azure imports modules.

### 1.1 Build your command

In your notepad, fill in this line with your values from "Before you start":

```powershell
./deploy/Deploy-Automation.ps1 -UseDeviceCode -SubscriptionId <subscription id> -ResourceGroupName rg-corpiphone-sync -Location <region> -AutomationAccountName aa-corpiphone-sync -AlertEmail <your email>
```

| Blank | Put here | Example |
| --- | --- | --- |
| `<subscription id>` | The Subscription ID you copied | `1a2b3c4d-5e6f-7a8b-9c0d-1e2f3a4b5c6d` |
| `<region>` | Your Azure region, lowercase | `eastus2` |
| `<your email>` | Where failure alerts should go | `it-alerts@contoso.com` |

The other two names, `rg-corpiphone-sync` and `aa-corpiphone-sync`, are suggestions. Keep them unless your company has naming rules. The automation account name must be unique only within your subscription.

### 1.2 Run it

1. Click into the PowerShell 7 window from Step 0 (the one whose prompt ends in `Iphonecontact-sync>`).
2. Paste your filled-in line and press Enter.
3. **First sign-in (Azure).** Within a few seconds a device code appears. Follow "How device-code sign-in works" above with your admin account.
4. **Second sign-in (Graph).** Shortly after, a second code appears. Same steps. This time, after signing in, a page lists permissions the script is asking for (reading and writing groups, assigning app roles). Click **Accept**.
5. Now wait. The script prints a cyan line for each stage. Nothing else is needed from you until it finishes.

### 1.3 What you will see

| Cyan line | Meaning | Wait |
| --- | --- | --- |
| `Ensuring dynamic group 'SG-CorpiPhone-Users'` | Creates the self-filling group | seconds |
| `Ensuring static exclusion group` | Creates the "receive but do not list" group | seconds |
| `Ensuring resource group` / `Ensuring Automation account` | Creates the Azure pieces | 1 minute |
| `Waiting for the identity's service principal` | Azure is publishing the account's login to Entra | up to 2 minutes |
| `Granting Graph application permissions` | Five lines starting with `granted` (or `exists` on a re-run) | seconds |
| `verified: no Contacts.* permission` | Important safety check passed. If the script stops here with red text about an unscoped Contacts permission, remove it as the text says and run again |  |
| `Ensuring PowerShell 7.6 runtime environment` | Creates the engine the runbooks run on | seconds |
| `Adding Microsoft.Graph 2.40.0 modules` | Five lines, each ending `imported`. **This is the slow part.** | 2 to 4 minutes each |
| `Importing and publishing runbooks` | Two lines ending `published` | 1 minute |
| `Ensuring schedules` | Creates both schedules and ends with `schedules created DISABLED` | seconds |
| `Ensuring action group` / `Ensuring metric alert` | Sets up the failure email | seconds |
| Green box `Deployment complete` | Done |  |

### 1.4 Write these down

The green box shows two lines:

```
-ApplicationId  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
-ObjectId       yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy
```

Copy both into your notepad, labelled. Step 2 needs them. To copy from the window, select the text with the mouse and press Enter, or right-click.

If you lose them: portal.azure.com, search for your automation account name, open it, click **Identity** in the left menu. The Object ID is shown there. For the Application ID, go to entra.microsoft.com, **Enterprise applications**, change the filter **Application type** to **Managed Identities**, find the account by name, and read **Application ID**.

### 1.5 Check it worked

- [ ] In entra.microsoft.com, click **Groups**. You see `SG-CorpiPhone-Users` with membership type **Dynamic** and `SG-Directory-Exclude` with **Assigned**. Both are empty for now. That is correct.
- [ ] Still in Entra, go to **Enterprise applications**, set the **Application type** filter to **Managed Identities**, and open `aa-corpiphone-sync`. Click **Permissions**. You see exactly five: `DeviceManagementManagedDevices.Read.All`, `User.ReadUpdate.All`, `User-Phone.ReadWrite.All`, `GroupMember.Read.All`, `RoleManagement.Read.Directory`. You must **not** see anything starting with `Contacts`. If you do, stop and read "If something goes wrong".
- [ ] In portal.azure.com, open the automation account, click **Runbooks**. Two runbooks show status **Published**.

**If people should receive the directory but not appear in it** (executives, security staff), add them now: in Entra, open `SG-Directory-Exclude`, click **Members**, **Add members**, pick them, **Select**.

## Step 2: Give the automation account permission to write contacts, but only in tagged mailboxes

At the end of this step Exchange knows the automation account and lets it edit contacts only in mailboxes carrying the CorpiPhone tag. Allow 10 minutes.

This is the one step where the *way* the permission is granted matters for privacy. The permission comes from Exchange with a scope attached, so the account can never touch a mailbox that is not tagged. That is why Step 1 deliberately did **not** grant a Contacts permission in Entra: an Entra permission has no scope and would quietly override the Exchange one.

### 2.1 Build your command

```powershell
./deploy/New-ExchangeAppScope.ps1 -UseDeviceCode -ApplicationId <application id> -ObjectId <object id> -StampedUserPrincipalName <iPhone user upn> -UnstampedUserPrincipalName <non-iPhone user upn>
```

| Blank | Put here |
| --- | --- |
| `<application id>` | The ApplicationId you wrote down in Step 1.4 |
| `<object id>` | The ObjectId from Step 1.4 |
| `<iPhone user upn>` | Sign-in address of someone who has a company iPhone, e.g. `jane.doe@contoso.com` |
| `<non-iPhone user upn>` | Sign-in address of someone who does not |

### 2.2 Run it

1. Paste the filled-in line into the PowerShell 7 window and press Enter.
2. **Sign-in (Exchange).** A device code appears. Use your admin account. This one can take a little longer to come back; give it a minute.
3. The script prints `created` or `exists` for three things: the service principal, the management scope `CorpiPhoneMailboxes`, and the role assignment.
4. A yellow note reminds you that Exchange needs 30 minutes to 2 hours before the permission is active for real runs. That is normal and expected.
5. The InScope table described in 2.3 prints next. Then, **Sign-in (Graph):** a second device code appears for the safety check. Same steps. The green or red result of that check prints after you sign in.

### 2.3 Read the results

You see a small table with two rows, one per test user, and a column **InScope**.

**Right now both rows must say InScope = False.** This is correct. Nobody carries the tag yet, so nobody is in scope. You will re-run this test after Step 3 and the iPhone user should then say True.

Below the table the script prints the automation account's Graph permissions. You want the green line `OK: no Contacts.* permission in Entra`. If instead you see a block of red exclamation marks, someone has granted a Contacts permission in Entra. Do not continue until it is removed (see "If something goes wrong").

### 2.4 Check it worked

- [ ] Three `created` or `exists` lines, no red text.
- [ ] Both test users InScope = False.
- [ ] Green `OK: no Contacts.* permission` line.

## Step 3: Run the morning runbook for the first time

At the end of this step every company iPhone user carries the tag, the dynamic group is full, and Exchange recognizes those mailboxes. Allow 20 minutes plus a wait of up to 2 hours before Step 5.

This step happens in the Azure portal, not in PowerShell. You will run the runbook twice: once in preview mode to read what it would do, then for real.

### 3.1 Open the runbook

1. Go to portal.azure.com and sign in with your admin account.
2. In the search box at the top, type `aa-corpiphone-sync` (or whatever you named the automation account) and click it.
3. In the left menu, under **Process Automation**, click **Runbooks**.
4. Click **Sync-IntuneToEntra**.

### 3.2 Preview run (WhatIf = true)

1. Click **Start** at the top. A panel opens on the right asking for parameters.
2. Leave every box as it is. **WHATIF** should show `True`. The group name box should show `SG-Directory-Exclude`. Click **OK**.
3. The page changes to a **Job** page. Its status goes from Queued to Starting to Running to Completed. This takes 1 to 3 minutes. If the page does not refresh, click **Refresh**.
4. When it says **Completed**, click the **Output** tab.

### 3.3 Read the preview

The output is a list of time-stamped lines. Read them like this:

| Line contains | Meaning |
| --- | --- |
| `Read 42 corporate iOS devices, 40 with a primary user` | How many company iPhones Intune reported, and how many have a person assigned |
| `WARN ... has no primary user` | A phone nobody is assigned to. It is ignored. These lines are on the job's **Warnings** tab, not the Output tab, so open that tab too. Fix them in Intune (see Before you start, Check B) |
| `WHATIF STAMP jane.doe@contoso.com` | This person **would** get the tag |
| `WHATIF PHONE jane.doe@contoso.com mobilePhone (none) -> +18045550123` | Her number **would** be written to her profile |
| `SKIP NoPhoneNumber : ...` | The phone has not reported its number to Intune yet. Normal for new phones; it arrives within a week |
| `SKIP PrivilegedRole : ...` | An admin. Their tag is set but their number is not written, by design |
| `SKIP ExcludedGroupMember : ...` | Someone in SG-Directory-Exclude. Same, by design |
| Last line, a summary with `UsersStamped=`, `Errors=0` | The totals |

**Your job here:** read the `WHATIF STAMP` list and ask "is this exactly the list of people who have a company iPhone?" If someone is missing, their phone has no primary user in Intune. If someone extra appears, a phone is marked Corporate in Intune that should not be. Fix those in Intune first, then run the preview again. Do not go live until the list looks right.

**If you see red text about a "foreign value"**, someone already uses the extensionAttribute1 field. The runbook stopped without changing anything. See "If something goes wrong".

### 3.4 Real run (WhatIf = false)

1. Go back to the runbook (click **Sync-IntuneToEntra** in the breadcrumb at the top, or use the search box again).
2. Click **Start**. In the panel, change **WHATIF** to `False`. Leave the rest. Click **OK**.
3. Wait for **Completed**, open **Output**. The same lines now say `APPLY STAMP` and `APPLY PHONE` instead of `WHATIF`. The summary should end with `Errors=0`.

If the status is **Failed** instead of Completed, click the **Errors** tab and read the message, then see "If something goes wrong".

### 3.5 Check the group filled up

1. Go to entra.microsoft.com, click **Groups**, open **SG-CorpiPhone-Users**, click **Members**.
2. Within about 10 minutes, the member list matches the people from the STAMP lines. If it is still empty after 15 minutes, click **Dynamic membership rules** in the group's menu and check that the rule reads `(user.extensionAttribute1 -eq "CorpiPhone")` and that **Processing status** is not showing an error. Very large tenants can take longer, up to 24 hours, but 80 users is a few minutes.

### 3.6 Re-test the Exchange permission

Back in the PowerShell 7 window in the Sandbox, run the Step 2 command again with `-TestOnly` added at the end:

```powershell
./deploy/New-ExchangeAppScope.ps1 -UseDeviceCode -TestOnly -ApplicationId <application id> -ObjectId <object id> -StampedUserPrincipalName <iPhone user upn> -UnstampedUserPrincipalName <non-iPhone user upn>
```

Now the iPhone user's row should say **InScope = True** and the other user **False**. If both still say False, wait 30 minutes and try again; Exchange is still catching up. If the iPhone user still says False after 2 hours, see "If something goes wrong".

### 3.7 Start the clock

Note the time you finished 3.4. Exchange needs **30 minutes to 2 hours** from that moment before Step 5 will work. Step 4 does not need to wait, so do it now.

## Step 4: Create the pilot group and the Outlook policies

At the end of this step 5 to 10 pilot phones have the Outlook policies, and Outlook on those phones has asked for Contacts access. Allow 20 minutes plus whatever time the pilot users need to open Outlook.

### 4.1 Make a pilot group

1. In entra.microsoft.com, click **Groups**, then **New group**.
2. Group type **Security**. Group name `SG-CorpiPhone-Pilot`. Membership type **Assigned**. Click **No members selected** under Members, search and pick your 5 to 10 pilot users, click **Select**, then **Create**.

Pick pilot users who are patient and reachable, and make sure at least two of them call each other sometimes, so you can test caller ID.

### 4.2 Run the policy script

In the PowerShell 7 window:

```powershell
./deploy/New-OutlookContactSyncPolicies.ps1 -UseDeviceCode -GroupName SG-CorpiPhone-Pilot
```

1. One device-code sign-in. Accept the permissions page if it appears (managing Intune apps, reading groups).
2. The script prints `created` for the App Protection Policy and the App Configuration Policy, `target apps updated` for each, and `assigned to group`.
3. It ends with a green box listing two device-restriction settings to verify by hand. Read 4.4 below.

**If your company already has an App Protection Policy for Outlook on iPhones**, do not create a second one. Find its exact name in intune.microsoft.com under **Apps**, **Protection** (labelled **App protection policies** in older portal versions), and run the script with that name added: `-AppProtectionPolicyName '<exact existing name>'`. The script then changes only the one setting it needs and leaves the rest of your policy alone.

### 4.3 What pilot users will see

Tell your pilot users this before they open Outlook:

> Open Outlook on your work iPhone. You will see a message that your administrator turned on contact sync, and then the phone will ask whether Outlook may access your Contacts. Tap **Allow**. Nothing leaves your phone; this only lets Outlook put the company directory into your Contacts app.

Policies reach a phone the next time Outlook checks in, usually within an hour of the user opening the app. If a user tapped **Don't Allow** by mistake: in Outlook, tap the profile icon, the gear icon, their account, and under **Save Contacts** there is an **Open Settings** link that takes them to the iOS permission.

### 4.4 Verify a device restriction setting

If your company has an iOS **Device restrictions** profile in Intune, one setting in it can silently block the whole thing.

1. In intune.microsoft.com, go to **Devices**, **iOS/iPadOS**, **Configuration**. Look for profiles of type **Device restrictions** assigned to your iPhones. If there are none, you are done with this step.
2. Open each one, click **Properties**, then **Edit** next to Configuration settings, and open the section **App Store, Doc Viewing, Gaming**.
3. Check these settings on that page:
   - **Block viewing non-corporate documents in corporate apps** must **not** be **Yes**. If it is, set it to **Not configured** and save.
   - Look at **Block viewing corporate documents in unmanaged apps**. If it says **Yes**, then **Allow managed apps to write contacts to unmanaged contacts accounts** must also say **Yes**. **Not configured** is not good enough here; if you get this wrong, contacts silently never reach the phone. Change it to **Yes** and save.
   - **Allow unmanaged apps to read from managed contacts accounts** is your choice and does not affect this project.

### 4.5 Check it worked

- [ ] In intune.microsoft.com, **Apps**, **Protection** (labelled **App protection policies** in older portal versions): a policy named `iOS APP - Outlook (Intune managed devices)` (or your existing one) assigned to `SG-CorpiPhone-Pilot`.
- [ ] **Apps**, **Configuration** (labelled **App configuration policies** in older portal versions): `ACP - Outlook contact sync (iOS and Android)` assigned to the same group.
- [ ] At least one pilot user has opened Outlook, tapped Allow, and an existing Outlook contact of theirs now appears in the iPhone's Contacts app. This proves the plumbing works before the directory arrives.

## Step 5: Put the directory into mailboxes, pilot first

At the end of this step the pilot phones show the company directory in Contacts and caller ID works. Then every stamped mailbox gets the directory. Allow 30 minutes of your time, spread out, after the Exchange wait from Step 3.7 has passed.

Everything here is in the Azure portal, same runbook page as Step 3, but the other runbook: **Sync-DirectoryToMailboxes**.

### 5.1 Wait check

Has it been at least 30 minutes since the real run in Step 3.4, and did the `-TestOnly` check in 3.6 show **True** for the iPhone user? If not, wait. Running early does no harm, but every mailbox will show as skipped with `AccessDenied` and you will just have to run again.

### 5.2 Preview for one pilot user

1. portal.azure.com, search your automation account, **Runbooks**, click **Sync-DirectoryToMailboxes**, click **Start**.
2. In the panel: leave **WHATIF** at `True`. In **USERPRINCIPALNAME**, type the sign-in address of one pilot user, for example `jane.doe@contoso.com`. Leave the group names as they are. Click **OK**.
3. Wait for **Completed**, open **Output**.

You should see a line like `jane.doe@contoso.com : owned=0 create=79 update=0 delete=0 unchanged=0`, then one `WHATIF CREATE` line per coworker, then a summary ending `Errors=0`. The create count should be the group size minus one (her own card is never written).

If instead you see `SKIP AccessDenied : jane.doe@contoso.com : 403`, Exchange has not caught up yet. Wait 30 minutes and start again. Not an error.

### 5.3 Real run for that one user

1. **Start** again. **WHATIF** = `False`, **USERPRINCIPALNAME** = the same pilot user. **OK**.
2. Completed, Output: the lines now say `APPLY CREATE`. Under a minute.

### 5.4 Look at the phone

On that pilot user's iPhone:

1. Open **Outlook**, tap the profile icon top left, then the small contacts/people icon. Scroll: the coworkers are there, each with a **Company Directory** category. This can take a few minutes to sync down.
2. Now open the iPhone's own **Contacts** app. The same coworkers appear. If Outlook has them but Contacts does not, go back to Step 4.3 and 4.4.
3. Have one coworker who is in the directory call this phone from their company iPhone. The screen shows their name, not just the number. **This is the whole point; when you see it, celebrate.**

### 5.5 Real run for everyone

1. **Start** once more. **WHATIF** = `False`. Leave **USERPRINCIPALNAME** empty this time. **OK**.
2. This run takes a few minutes (about 80 mailboxes, roughly 6,400 contacts on the first pass). Output shows one `owned=... create=...` line per mailbox and `APPLY CREATE` lines under each.
3. Read the summary at the end: `Added=` should be about group size × (group size − 1). `Errors=0`. If `MailboxesSkipped` shows some `AccessDenied`, those people were stamped very recently; run again in an hour and they fill in. `NoMailbox` means a person without an Exchange mailbox, usually a service account, which is fine.

From now on, each pilot user's phone picks up the directory the next time Outlook syncs.

### 5.6 Two tests that prove the upkeep works

**Leaver test.** Use a spare test iPhone enrolled to a test account that is in the pilot (not a real person's phone). In intune.microsoft.com, open that iPhone and click **Retire**. Run **Sync-IntuneToEntra** live (Step 3.4): the output shows `APPLY UNSTAMP` for the test account and it drops out of `SG-CorpiPhone-Users` within minutes. Then run **Sync-DirectoryToMailboxes** live with no user name: the output shows `APPLY DELETE` for its card in every other mailbox, and the contact disappears from the other pilot phones. Finally enroll the test phone again with the test account and run both runbooks live again; it is restored. Note: the leaver's own mailbox keeps the cards it already had, because once untagged the automation is no longer allowed into that mailbox. Their phone loses them when Outlook is retired from it.

**Edit test.** Sign in to Outlook on the web (outlook.office.com) as a pilot user, open **People**, open one contact in the **Company Directory** category and change the job title to something silly. Save. (Use Outlook on the web, not the iPhone Contacts app: depending on the Outlook version, an edit made in the Contacts app is either synced back or simply overwritten by Outlook, so it is not a reliable test.) Run **Sync-DirectoryToMailboxes** live. The output shows exactly one `APPLY UPDATE` for that card ending in `(edited in mailbox)`, the summary shows `LocalEditsReverted=1`, and Outlook on the web and the phone show the real title again within minutes. Nothing else changed.

**No-churn check.** Run **Sync-DirectoryToMailboxes** live once more right away. The summary must show `Updated=0`. If instead many cards show `(edited in mailbox)` that nobody touched, Exchange is reformatting a field: run it with **DETECTLOCALEDITS** = `False` from then on (also edit the Monday schedule's parameters) and tell the repo maintainer which field differs.

### 5.7 Check it worked

- [ ] A pilot phone shows coworkers in the Contacts app and resolves caller ID.
- [ ] The full run ended `Errors=0`.
- [ ] Leaver test, edit test and no-churn check behaved as described.

## Step 6: Turn on the schedules and check the alert

At the end of this step the two runbooks run by themselves, every morning and every Monday, and you get an email if one fails. Allow 10 minutes, then a week of leaving it alone.

Step 1 already created both schedules but left them switched off so nothing could run before you had tested. Now you switch them on.

### 6.1 Enable the two schedules

1. In the Sandbox PowerShell 7 window, run the Step 1 command again with `-EnableSchedules` added at the end. It changes nothing else. If a schedule's original start time has already passed (it will have, after a pilot), the script recreates that schedule with the next correct start time before switching it on.
2. Check: portal.azure.com, your automation account, **Shared Resources**, **Schedules**. Both `Sync-IntuneToEntra-Daily-0500` and `Sync-DirectoryToMailboxes-Weekly-Mon-0800` show **Enabled: Yes** and a **Next run** time in the future.

The times are 05:00 every day and 08:00 every Monday, in the time zone chosen at deployment (America/New\_York unless you changed it). The three-hour gap on Monday is deliberate: it gives Exchange time to recognize anyone tagged that morning.

Portal alternative: open each schedule, switch **Enabled** to **Yes**, **Save**. If the portal complains the start time is in the past, set it to the next 05:00 (daily schedule) or the next Monday 08:00 (weekly schedule) in America/New\_York, then save.

Re-running the Step 1 command later without `-EnableSchedules` is safe: it leaves the schedules exactly as they are.

### 6.2 Check the schedules are linked with the right settings

1. Still in the automation account, click **Runbooks**, open **Sync-IntuneToEntra**, then click **Schedules** in its left menu (under Resources). The daily schedule is listed. Click it and look at **Parameters**: **WHATIF** must be `False`. If it said True, the nightly run would only pretend.
2. Do the same for **Sync-DirectoryToMailboxes** and its Monday schedule.

### 6.3 Test the failure email

1. Open **Sync-DirectoryToMailboxes**, click **Start**, set **WHATIF** to `True` and **GROUPNAME** to `this-group-does-not-exist`, click **OK**. The job fails on purpose within a minute.
2. Within about 15 minutes an email titled something like "Fired: Sev2 Azure Monitor Alert aa-corpiphone-sync-job-failed-or-suspended" arrives at the address you gave in Step 1. If it does not arrive after 30 minutes, check the spam folder, then in portal.azure.com search **Alerts**, open **Alert rules**, and confirm the rule exists and is enabled.

The failed job did nothing to your tenant; it stopped before reading any mailbox.

### 6.4 Leave it for a week

Let both schedules run on their own. On the following Monday afternoon, open the automation account, click **Jobs**, and confirm the week's jobs show **Completed**. Open Monday's `Sync-DirectoryToMailboxes` job, click **Output**, and read the summary line at the end: a handful of `Added`, `Updated` or `Deleted` if people changed, `Errors=0`.

### 6.5 Check it worked

- [ ] Both schedules show **Enabled: Yes**.
- [ ] Both linked schedules have **WHATIF = False**.
- [ ] The test failure email arrived.
- [ ] The first scheduled job of each runbook: its **Output** starts with a line containing `Mode=APPLY` (not `Mode=WHATIF`). If it says WHATIF, the schedule's WHATIF setting did not arrive as False; fix it on the schedule's parameters.
- [ ] After a week, all scheduled jobs are Completed.

## Step 7: Roll out to everyone

At the end of this step every company iPhone user has the directory. Allow 15 minutes of your time, then three weeks of letting it run before you call it finished.

### 7.1 Move the Outlook policies from the pilot group to the real group

The mailboxes already have the directory (Step 5.5 did everyone). What is still pilot-only is the pair of Outlook policies that copy contacts to the phone.

Run the Step 4 command again with the real group name:

```powershell
./deploy/New-OutlookContactSyncPolicies.ps1 -UseDeviceCode -GroupName SG-CorpiPhone-Users
```

The script finds the two existing policies by name, leaves their settings alone, and **adds** `SG-CorpiPhone-Users` to their assignments. The pilot group stays assigned too, which does no harm; you can remove it later in intune.microsoft.com under **Apps**, **Protection** (labelled **App protection policies** in older portal versions), the policy, **Properties**, **Edit** next to Assignments.

If you closed the Sandbox since Step 5, repeat Step 0 first; it takes 15 minutes and nothing else is lost.

### 7.2 Tell people what to expect

Over the next day, every iPhone outside the pilot shows the one-time Contacts permission prompt. Send this paragraph to all company iPhone users before that happens (email, Teams, whatever you normally use):

> Your corporate iPhone will soon show the company directory in the Contacts app, so you see a coworker's name when they call. Outlook does this for you. The next time you open Outlook it will tell you that your administrator turned on contact sync, and iOS will ask whether Outlook may access your Contacts. Tap **Allow**. Nothing is uploaded from your phone; the directory flows one way, from the company directory into Outlook and then into Contacts. If you tapped **Don't Allow** by mistake, open Outlook, go to Settings, your account, and use the **Open Settings** link under Save Contacts to turn it on.

### 7.3 Run the directory sync once so nobody waits for Monday

In the portal, start **Sync-DirectoryToMailboxes** with **WHATIF** = `False` and no user name (exactly like Step 5.5). It mostly reports `unchanged`, and fills in any mailbox that was skipped before.

### 7.4 Three clean weeks, then done

Check the **Jobs** list each Monday for three weeks. All Completed, summaries sensible, nobody complaining that a contact is wrong. Then close your change record. The exit criteria from the project plan are:

- [ ] The tagged users match Intune's corporate iPhone primary users exactly (compare `SG-CorpiPhone-Users` members with Intune's device list).
- [ ] No unexplained `SKIP` lines in the morning job.
- [ ] Each member's mailbox holds group size minus one directory contacts (the Monday job's per-mailbox `owned=` number).
- [ ] Users confirm caller ID works.
- [ ] Three consecutive weeks of Completed scheduled jobs.

## Day to day

Once it is running, almost nothing needs you. These are the only things that ever come up.

| Situation | What you do | What happens on its own |
| --- | --- | --- |
| **A new company iPhone** | Enroll it in Intune with the user set as **primary user**. That is all. | Next morning at 05:00 the user is tagged. Within minutes they join the group and the Outlook policies reach the phone. Monday at 08:00 their mailbox gets the directory and they appear in everyone else's. |
| **New phone, but the user cannot wait for Monday** | After the next morning run, start **Sync-DirectoryToMailboxes** in the portal with **WHATIF** = `False` and **USERPRINCIPALNAME** = their address. Wait at least 30 minutes after the morning run before doing this. | Their mailbox fills in under a minute. |
| **A phone is handed to someone else** | Retire (or wipe) the phone in intune.microsoft.com, then enroll it again with the new person signing in. Intune cannot change an iPhone's primary user in place. | The old holder is untagged and the new one tagged next morning. Monday swaps the contact cards everywhere. The old holder's mailbox keeps the cards it had; they can delete the **Company Directory** category in Outlook if they want them gone. |
| **Someone leaves** | Retire or wipe the phone in Intune, or disable the account as you normally do. | Untagged next morning, removed from everyone's contacts on Monday. |
| **Someone should receive the directory but not be in it** (an executive, security staff) | In entra.microsoft.com, add them to **SG-Directory-Exclude**. | From Monday their card is removed from other phones. They keep receiving everyone else's cards, and their phone number stops being written to their profile. |
| **A person's title, department or number changed in the directory** | Nothing. | Monday's run updates their card in every mailbox. |
| **You get the failure email** | Open the automation account, **Jobs**, click the failed job, read the **Errors** tab, then the **Output** tab's last lines. Match the message against "If something goes wrong". Fix the cause and start the runbook again; running twice is always safe. |  |
| **You want to see what ran last night** | Automation account, **Jobs**, click the job, **Output**. One line per change, a summary at the end. Then open the **Warnings** tab: it holds retry notices (harmless) and the `has no primary user` warnings (phones to fix in Intune). Read it after every morning run during the first weeks. |  |

**Two things never to do.** Do not add or remove members of `SG-CorpiPhone-Users` by hand; the rule refills it and your change is undone within minutes. Do not update the Microsoft.Graph modules inside the automation account's runtime environment; a newer version may not load and both runbooks would fail. If a newer version is ever needed, create a second runtime environment to test it first (the README explains how).

## If something goes wrong

Find the message or symptom you are seeing in the left column. Everything here has been seen before and has a fix. The project's README has a longer troubleshooting section if you need more.

| You see | What it means | What to do |
| --- | --- | --- |
| Sandbox setup: `SETUP FAILED` with a download error | No internet in the Sandbox, or GitHub was unreachable for a moment | Check your PC is online. Close the Sandbox, open a new one, paste again. |
| Sandbox setup: a browser asks you to sign in to GitHub and you cancelled | The repo is private and Git needs your GitHub login | Paste the script again and complete the GitHub sign-in in the browser window that opens. |
| Sandbox setup: `Tests Passed: ..., Failed: 1` or more | The project files on your branch have a problem | Do not continue. Send the red lines to whoever maintains the repo. |
| Any script: the device code page says "code expired" | You took more than about 15 minutes | Press Ctrl+C in PowerShell, paste the command again, and enter the new code promptly. |
| Any script: `insufficient privileges` or `Authorization_RequestDenied` | Your admin account lacks a role, or a PIM role was not activated | Check roles (Before you start, 1). Activate PIM roles. Run again. |
| Step 1: red `The identity holds Graph permission 'Contacts.ReadWrite'` | Someone granted an unscoped Contacts permission in Entra. It would let the account read **every** mailbox | In entra.microsoft.com, **Enterprise applications**, filter **Managed Identities**, open the account, **Permissions**, remove the Contacts entry (or run the exact `Remove-...` command the script printed). Re-run Step 2 with `-TestOnly`. |
| Step 1: a module line ends `Failed` instead of `imported` | Azure could not import that Graph module version | Run the Step 1 command again; transient failures clear. If it fails twice on the same module, the pinned version has a problem; ask the repo maintainer before changing `-GraphModuleVersion`. |
| Any morning: job **Failed** with `Mass unstamp refused` | A safety brake. Intune returned no company iPhones, or more than a quarter of tagged people were about to lose the tag at once. Nothing was changed | Read the `WOULD UNSTAMP` lines in the Output. If it looks like a glitch, do nothing; tomorrow's run sorts itself out. If it is real (many phones really were returned), start **Sync-IntuneToEntra** once by hand with **WHATIF** = `False` and **ALLOWMASSUNSTAMP** = `True`. |
| Step 3: red `user(s) carry a foreign value in extensionAttribute1` | Your company already uses that spare field for something else. The runbook refused to overwrite it and changed nothing | Either clear the field on the users listed, or ask the repo maintainer to switch the project to a different attribute. This must be decided before continuing. To clear it for a cloud-only user: `Update-MgUser -UserId <upn> -OnPremisesExtensionAttributes @{ extensionAttribute1 = $null }` (after `Connect-MgGraph -Scopes User.ReadWrite.All -UseDeviceCode`). For users synced from on-premises Active Directory, clear it in AD instead. |
| Step 3: a person with an iPhone is missing from the STAMP list | Their phone has no primary user in Intune, or is not marked Corporate | Intune, **Devices**, open the phone. If **Ownership** is not Corporate, change it under **Properties**. If the primary user is blank or wrong, retire the phone and enroll it again with that person signing in (Intune cannot change an iPhone's primary user in place). If the output shows `SKIP SyncedFromOnPremises` for them, their account comes from on-premises Active Directory, which this project does not write to: set the tag and number in AD, or ask whoever runs AD. Run the preview again. |
| Step 3: `SKIP NoPhoneNumber` for someone | The phone has not reported its number yet, or the carrier/eSIM never does | Wait up to a week for Intune's inventory. If it never appears, type the number into the user's **Mobile phone** field in Entra by hand. Everything else still works. |
| Step 3: `SKIP PrivilegedRole` | The person is an admin. The runbook never writes admins' numbers, on purpose | If you want their number in the directory, type it into their profile in Entra by hand. |
| Step 3.5: the dynamic group stays empty | The rule is wrong, or dynamic groups are not licensed | Entra, the group, **Dynamic membership rules**: rule must be exactly `(user.extensionAttribute1 -eq "CorpiPhone")`. Dynamic groups need an Entra ID P1 license (included in E3/E5). |
| Step 3.6 or 5: the iPhone user still shows `InScope = False` after 2 hours, or every mailbox is `AccessDenied` | Exchange does not see the tag, or the permission was not created | Run Step 2 again without `-TestOnly`; it repairs anything missing. In Exchange admin, open the user and check their **Custom attribute 1** shows `CorpiPhone`; if blank, run the morning runbook live again and wait 2 hours. |
| Step 5: one or two mailboxes `AccessDenied`, the rest fine | Those people were tagged very recently | Wait an hour and run again. |
| Step 5: `NoMailbox` | That account has no Exchange mailbox (a service account, a guest) | Nothing to do, unless they should have a mailbox. |
| Phone: Outlook shows the contacts but the iPhone Contacts app does not | One of four switches is off | 1. Outlook on the phone: profile icon, gear, the account, **Save Contacts** must be on. 2. iPhone **Settings**, **Outlook**, **Contacts** must be on. 3. Intune: the App Protection Policy must have **Sync policy managed app data with native apps** = Allow and be assigned to the user. 4. Step 4.4's device restriction must not say Block. Then close and reopen Outlook. |
| Phone: the user tapped Don't Allow | iOS permission denied | Outlook, profile icon, gear, the account, **Open Settings** under Save Contacts, switch Contacts on. |
| Phone: a coworker appears twice | Either the user also saved that person as a personal contact (fine, they can delete their copy), or two runbook cards exist for one person | The Monday run deletes duplicate runbook cards by itself. Personal duplicates are the user's to clean up. |
| Job **Failed**, Errors tab mentions `could not load` a Microsoft.Graph module | On the very first run: the pinned Graph version does not load on PowerShell 7.6. Later: the runtime environment lost a module or someone changed its version | First run: run the Step 1 command again with `-PowerShellVersion 7.4` added; it moves the runbooks to a 7.4 environment. Later: run the Step 1 command again unchanged; it re-imports the pinned version. Never update modules in place. Tell the repo maintainer either way. |
| Job Output has yellow `throttled, retrying` lines | Microsoft slowed the script down briefly. It waited and continued | Nothing. Only a job that ends **Failed** with `still throttled` needs a re-run. |
| Job **Failed** but the output shows a summary with `Errors=` bigger than 0 | Some changes failed, the rest were applied | Read the error lines above the summary, fix the cause, start the job again. Re-running is always safe; it only touches what is still wrong. |

## Undo everything

You can stop or remove this at any depth, from "pause it" to "it never existed". Do the levels in order; each one is a superset of the one before.

### Level 1: Pause (nothing changes any more, phones keep what they have)

portal.azure.com, automation account, **Schedules**, open each schedule, **Enabled** = **No**, **Save**.

### Level 2: Remove the directory from the phones

In intune.microsoft.com, **Apps**, **Configuration** (labelled **App configuration policies** in older portal versions), open `ACP - Outlook contact sync (iOS and Android)`, **Properties**, **Edit** next to Assignments, remove the groups, **Review + save**. Outlook stops syncing the next time it checks in. Microsoft does not document whether Outlook then removes the entries it already put in the Contacts app, so look at one phone. If you are going on to Level 3, do Level 3 first, while Outlook is still syncing, so the deletions reach the phones. Personal contacts are untouched.

### Level 3: Remove the contact cards from everyone's mailbox

1. In entra.microsoft.com, open **SG-Directory-Exclude**, **Members**, **Add members**, and add the group **SG-CorpiPhone-Users** itself as a member (groups can contain groups). Now everyone is "excluded", so the directory is empty.
2. In the portal, start **Sync-DirectoryToMailboxes** with **WHATIF** = `True` first and read the output: it should show one `WHATIF DELETE` per card and no creates. Then run it with **WHATIF** = `False`.
3. Every runbook-written card is deleted. Nothing else in any mailbox is touched; the runbook only ever deletes cards carrying its own hidden tag.
4. Remove `SG-CorpiPhone-Users` from the exclusion group again if you plan to come back.

### Level 4: Remove the tag

Start **Sync-IntuneToEntra** live after retiring the phones, or clear the attribute by hand for every tagged user with one line in the PowerShell 7 window: `Connect-MgGraph -Scopes User.ReadWrite.All -UseDeviceCode -NoWelcome; Get-MgUser -All -Property Id,OnPremisesExtensionAttributes | Where-Object { $_.OnPremisesExtensionAttributes.ExtensionAttribute1 -eq 'CorpiPhone' } | ForEach-Object { Update-MgUser -UserId $_.Id -OnPremisesExtensionAttributes @{ extensionAttribute1 = $null } }`. The dynamic group empties and the Exchange permission matches nobody. Do Level 3 first: once the tag is gone the automation can no longer reach the mailboxes to remove the cards.

### Level 5: Delete it all

1. In the Sandbox PowerShell window, connect to Exchange and remove the permission: `Connect-ExchangeOnline -Device` then `Remove-ManagementRoleAssignment -Identity CorpiPhone-ContactsReadWrite-CorpiPhoneMailboxes` (if that name is not found, list the assignments with `Get-ManagementRoleAssignment -Role 'Application Contacts.ReadWrite' | Format-List Name,RoleAssignee*,CustomResourceScope` and remove the one for this identity) then `Remove-ManagementScope -Identity CorpiPhoneMailboxes` then `Remove-ServicePrincipal -Identity <object id>`.
2. portal.azure.com, open the resource group `rg-corpiphone-sync`, click **Delete resource group**, type its name to confirm. This removes the automation account, its identity, schedules and the alert in one go.
3. entra.microsoft.com, **Groups**, delete `SG-CorpiPhone-Users` and `SG-Directory-Exclude`.
4. intune.microsoft.com, delete the two policies under **Apps**, or just remove their assignments if you want to keep them.

Nothing else was created anywhere.

## Cheat sheet

Every command in one place. Fill in the angle brackets, paste each as one line into the PowerShell 7 window the Sandbox setup opened. Portal steps are listed where there is no command.

| Step | Where | Command or action |
| --- | --- | --- |
| 0 | Sandbox, Windows PowerShell (blue) | Paste the contents of `deploy/sandbox/Initialize-SandboxEnvironment.ps1` |
| 1 | PowerShell 7 | `./deploy/Deploy-Automation.ps1 -UseDeviceCode -SubscriptionId <subscription id> -ResourceGroupName rg-corpiphone-sync -Location <region> -AutomationAccountName aa-corpiphone-sync -AlertEmail <your email>` |
| 2 | PowerShell 7 | `./deploy/New-ExchangeAppScope.ps1 -UseDeviceCode -ApplicationId <application id> -ObjectId <object id> -StampedUserPrincipalName <iPhone user upn> -UnstampedUserPrincipalName <non-iPhone user upn>` |
| 3 | Azure portal | Automation account, Runbooks, **Sync-IntuneToEntra**, Start with WHATIF = True; read; Start again with WHATIF = False |
| 3.6 | PowerShell 7 | Same as step 2 with `-TestOnly` added |
| 4 | Entra | Groups, New group, `SG-CorpiPhone-Pilot`, add 5 to 10 users |
| 4 | PowerShell 7 | `./deploy/New-OutlookContactSyncPolicies.ps1 -UseDeviceCode -GroupName SG-CorpiPhone-Pilot` |
| 5 | Azure portal | **Sync-DirectoryToMailboxes**, Start: WHATIF = True + USERPRINCIPALNAME = pilot user; then WHATIF = False + same user; then WHATIF = False, no user |
| 6 | PowerShell 7 | Step 1 command with `-EnableSchedules` added. Then check both schedules show Enabled = Yes and a future Next run |
| 7 | PowerShell 7 | `./deploy/New-OutlookContactSyncPolicies.ps1 -UseDeviceCode -GroupName SG-CorpiPhone-Users` |
| Onboarding one user early | Azure portal | **Sync-DirectoryToMailboxes**, Start: WHATIF = False, USERPRINCIPALNAME = the user |

**Values to keep in your notepad**

| Name | Value | Where it came from |
| --- | --- | --- |
| Subscription ID |  | portal.azure.com, Subscriptions |
| Region |  | your choice, e.g. eastus2 |
| Alert email |  | your choice |
| Application ID |  | Step 1 green box |
| Object ID |  | Step 1 green box |
| Test user with iPhone |  | your choice |
| Test user without iPhone |  | your choice |

**Names this guide assumes.** Resource group `rg-corpiphone-sync`, automation account `aa-corpiphone-sync`, groups `SG-CorpiPhone-Users`, `SG-Directory-Exclude`, `SG-CorpiPhone-Pilot`, tag value `CorpiPhone`. If you changed any of them in Step 1, use your names everywhere after that.
