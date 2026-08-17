# Employee offboarding script

Keep these files in the same folder:

- `Invoke-EmployeeOffboarding.ps1`
- `Run-EmployeeOffboarding.bat`

## What the script does

The script performs a complete read-only preflight before its first change. It resolves the employee, Exchange Online and Microsoft Entra identities, group memberships, mailbox state, and license assignments. It resolves the employee's AD `manager` and both hybrid mailbox objects when shared-mailbox conversion or manager group-ownership transfer is selected.

After preflight and confirmation, it:

1. Disables the on-premises AD account, sets the AD Notes (`info`) termination marker to `TERM: MM/DD/YYYY`, blocks Microsoft Entra sign-in, and revokes cloud sign-in sessions. Existing unrelated Notes text is preserved; an existing `TERM:` line is replaced.
2. Removes direct on-premises AD memberships except the account's primary group and the names in `-KeepADGroups`. The defaults are `Domain Users` and `SG_Intranet Users`.
3. Removes cloud-managed Exchange distribution/mail-enabled-security group memberships, Microsoft 365 group membership and ownership, and cloud-only static Entra security-group memberships.
4. Prompts `Need to create a shared mailbox (YES or NO)`. `YES` converts the hybrid mailbox on both sides and grants the AD manager Full Access. `NO` skips conversion and manager mailbox access. `-GrantSendAs` is optional with `YES`.
5. Removes eligible directly assigned Microsoft 365 licenses. With `YES`, removal requires verified mailbox conversion and manager Full Access. With `NO`, conversion/delegation is skipped and Exchange-license removal can deprovision the mailbox according to the organization's retention policy.
6. Writes a time-stamped CSV audit log beside the scripts, including the shared-mailbox answer and termination-note value.

The script is repeatable: already-completed actions are reported as skipped where they can be verified.

## Prerequisites

- Run the BAT normally from the employee-administration workstation. Local **Run as administrator** is not required for delegated AD/Exchange work.
- Windows PowerShell 5.1 and the Active Directory RSAT module.
- `ExchangeOnlineManagement`, `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions`, and `Microsoft.Graph.Groups` installed for the account running PowerShell.
- The delegated AD account needs rights to disable users and remove the applicable AD group memberships.
- The signed-in Exchange account needs rights to manage recipients, group memberships, mailbox types, and mailbox delegation.
- The signed-in Microsoft Graph account needs the tenant roles that allow blocking sign-in, revoking sessions, changing group memberships, and removing licenses. The script requests `User.Read.All`, `User.EnableDisableAccount.All`, `User.RevokeSessions.All`, `GroupMember.ReadWrite.All`, and `LicenseAssignment.ReadWrite.All`; Graph consent scopes alone do not grant the administrative role.
- When answering `YES` in a hybrid environment, supply the on-premises Exchange server FQDN so the delegated account can open an Exchange remote PowerShell session. The endpoint is expected at `http://SERVER/PowerShell/` unless a full URL is supplied. It isn't required when answering `NO`.

## Preview first

Replace the examples with the real SAM account name, Exchange admin UPN, and Exchange server FQDN:

```bat
Run-EmployeeOffboarding.bat jsmith -OnPremExchangeServer exch01.contoso.local -ExchangeAdminUPN admin@contoso.com -GraphAdminUPN admin@contoso.com -WhatIf
```

Review every preview line. In particular, confirm that the protected intranet group name exactly matches your environment.

If `-CreateSharedMailbox` is omitted, the script asks the required `YES` or `NO` question. For a non-interactive choice, add either `-CreateSharedMailbox YES` or `-CreateSharedMailbox NO`.

## Live run

Run the same command without `-WhatIf`:

```bat
Run-EmployeeOffboarding.bat jsmith -OnPremExchangeServer exch01.contoso.local -ExchangeAdminUPN admin@contoso.com -GraphAdminUPN admin@contoso.com
```

The script displays the resolved offboarding summary and, when required, the manager, then requires this exact confirmation:

```text
OFFBOARD jsmith
```

Exit code `2` means the script completed but the CSV contains follow-up items. Any other nonzero exit code means an error stopped the script.

## Important switches

- `-CreateSharedMailbox YES|NO`: supplies the mailbox answer on the command line. If omitted, the script prompts for it.
- `-AdminUser DOMAIN\username`: supplies the delegated AD/on-premises Exchange username. If omitted, the script prompts for it before requesting the password.
- `-TerminationDate YYYY-MM-DD`: overrides today's date in the AD Notes marker. Without it, the script writes the date on which it runs, such as `TERM: 08/17/2026`.
- `-GrantSendAs`: also grants the manager Send As permission. Full Access alone does not let the manager send as the former employee.
- `-GraphAdminUPN`: verifies that Microsoft Graph authenticated as the intended administrator instead of silently using a cached everyday-account context.
- `-TransferSoleOwnedMicrosoft365GroupsToManager`: if the employee is the only Microsoft 365 group owner, adds the manager as member/owner before removing the employee. Without it, the script leaves that group unchanged and flags it for attention.
- `-SkipCloudGroupRemoval`: skips Exchange Online, Microsoft 365, and cloud security-group cleanup.
- `-SkipLicenseRemoval`: keeps every license.
- `-OverrideSharedMailboxLicenseSafety`: allows direct-license removal even when the mailbox is over 50 GB, has an active archive, is on hold, or its size cannot be verified. Use only after the licensing/retention requirement is confirmed.
- `-CloudOnlyMailbox`: with `-CreateSharedMailbox YES`, skips `Set-RemoteMailbox`. The script rejects this switch when Microsoft Graph reports that the account is directory-synchronized, because a later sync can revert or disconnect a mismatched unlicensed mailbox.
- `-Force`: skips the typed `OFFBOARD username` prompt. It does not bypass permissions or safety gates.

## Follow-up items outside this script

The termination ticket should separately address anything that applies:

- Reassign the former employee's direct reports and any application, security-group, SharePoint, Power Platform, Azure, or administrative ownership.
- Review Microsoft Entra directory roles, PIM eligibility, enterprise-app assignments, and privileged/role-assignable groups.
- Decide whether policy requires resetting the password to a random value. The script disables both identities and revokes cloud sessions, but it does not reset the password.
- Review dynamic AD, Entra, and dynamic distribution-group rules. Dynamic memberships cannot be removed directly and may continue assigning licenses.
- Reassign Teams private/shared-channel roles and direct SharePoint/OneDrive permissions; standard Team membership follows its Microsoft 365 group.
- Configure mailbox forwarding or automatic replies if policy requires them. Full Access does not forward new mail.
- Apply OneDrive retention/delegation, device wipe/retirement, phone/Teams Calling cleanup, and line-of-business application offboarding.
- Follow legal-hold, retention, archive, and deletion policy before overriding the mailbox license gate or deleting the account.
- After directory synchronization and license processing, verify that no group-inherited license remains.

## Validation status

The PowerShell parser check passed. The script has not been run against live Active Directory, on-premises Exchange, Exchange Online, or Microsoft Graph in this environment. Always use `-WhatIf` first.
