# Employee offboarding script - Graph-free branch

This branch never imports, connects to, or calls Microsoft Graph. It uses on-premises Active Directory, on-premises Exchange PowerShell when hybrid conversion is selected, and Exchange Online PowerShell. Entra sign-in/session controls, cloud-only Entra security groups, and Microsoft 365 licenses are mandatory manual follow-up items.

Keep these files in the same folder:

- `Invoke-EmployeeOffboarding.ps1`
- `Run-EmployeeOffboarding.bat`

## What the script does

The script performs a read-only AD and Exchange preflight before its first change. It resolves the employee, Exchange recipients, mailbox state, direct AD groups, Exchange distribution groups, and Microsoft 365 group membership/ownership. It resolves the employee's AD `manager` and both hybrid mailbox objects when shared-mailbox conversion or manager group-ownership transfer is selected. It does not inspect Entra-only groups, cloud session state, or license assignments.

After preflight and confirmation, it:

1. Disables the on-premises AD account and sets the AD Notes (`info`) termination marker to `TERM: MM/DD/YYYY`. Existing unrelated Notes text is preserved; an existing `TERM:` line is replaced. It records manual follow-ups to block Entra sign-in and revoke cloud sessions.
2. Removes direct on-premises AD memberships except the account's primary group and the names in `-KeepADGroups`. The defaults are `Domain Users` and `SG_Intranet Users`.
3. Uses Exchange Online to remove cloud-managed distribution/mail-enabled-security group memberships and Microsoft 365 group membership/ownership. It records a manual follow-up for non-mail-enabled cloud-only Entra security groups.
4. Prompts `Need to create a shared mailbox (YES or NO)`. `YES` converts the hybrid mailbox on both sides and grants the AD manager Full Access. `NO` skips conversion and manager mailbox access. `-GrantSendAs` is optional with `YES`.
5. Does not inventory licenses or call a license-assignment API. It records a manual Admin Center follow-up. Removing AD or Microsoft 365 group memberships can still revoke group-assigned licenses.
6. Writes a time-stamped CSV audit log beside the scripts, including the shared-mailbox answer and termination-note value.

The script is repeatable: already-completed actions are reported as skipped where they can be verified.

## Prerequisites

- Run the BAT normally from the employee-administration workstation. Local **Run as administrator** is not required for delegated AD/Exchange work.
- Windows PowerShell 5.1 and the Active Directory RSAT module.
- The `ExchangeOnlineManagement` module. No Microsoft Graph module, Graph consent, or Graph administrator identity is used.
- The delegated AD account needs rights to disable users and remove the applicable AD group memberships.
- The signed-in Exchange account needs rights to manage recipients, group memberships, mailbox types, and mailbox delegation.
- When answering `YES` in a hybrid environment, supply the on-premises Exchange server FQDN so the delegated account can open an Exchange remote PowerShell session. The endpoint is expected at `http://SERVER/PowerShell/` unless a full URL is supplied. It isn't required when answering `NO`.

## Preview first

Replace the examples with the real SAM account name, Exchange admin UPN, and Exchange server FQDN:

```bat
Run-EmployeeOffboarding.bat jsmith -OnPremExchangeServer exch01.contoso.local -ExchangeAdminUPN admin@contoso.com -WhatIf
```

Review every preview line. In particular, confirm that the protected intranet group name exactly matches your environment.

If `-CreateSharedMailbox` is omitted, the script asks the required `YES` or `NO` question. For a non-interactive choice, add either `-CreateSharedMailbox YES` or `-CreateSharedMailbox NO`.

## Live run

Run the same command without `-WhatIf`:

```bat
Run-EmployeeOffboarding.bat jsmith -OnPremExchangeServer exch01.contoso.local -ExchangeAdminUPN admin@contoso.com
```

The script displays the resolved offboarding summary and, when required, the manager, then requires this exact confirmation:

```text
OFFBOARD jsmith
```

Exit code `2` means the script completed but the CSV contains follow-up items. Exit code `2` is expected on every run of this branch because Entra and license tasks remain manual. Any other nonzero exit code means an error stopped the script.

## Important switches

- `-CreateSharedMailbox YES|NO`: supplies the mailbox answer on the command line. If omitted, the script prompts for it.
- `-AdminUser DOMAIN\username`: supplies the delegated AD/on-premises Exchange username. If omitted, the script prompts for it before requesting the password.
- `-TerminationDate YYYY-MM-DD`: overrides today's date in the AD Notes marker. Without it, the script writes the date on which it runs, such as `TERM: 08/17/2026`.
- `-GrantSendAs`: also grants the manager Send As permission. Full Access alone does not let the manager send as the former employee.
- `-TransferSoleOwnedMicrosoft365GroupsToManager`: if the employee is the only Microsoft 365 group owner, adds the manager as member/owner before removing the employee. Without it, the script leaves that group unchanged and flags it for attention.
- `-SkipCloudGroupRemoval`: skips Exchange Online distribution-group and Microsoft 365 group cleanup. Entra-only security groups are always manual on this branch.
- `-CloudOnlyMailbox`: with `-CreateSharedMailbox YES`, skips `Set-RemoteMailbox`. The script uses Exchange Online's `IsDirSynced` value and fails closed when the synchronization state is true or unavailable.
- `-Force`: skips the typed `OFFBOARD username` prompt. It does not bypass permissions or safety gates.

## Follow-up items outside this script

The termination ticket should separately address anything that applies:

- Immediately block Entra sign-in and revoke cloud sessions in the Microsoft 365 or Entra admin center. Disabling on-premises AD does not immediately revoke existing cloud sessions.
- Review cloud-only Entra security groups, directory roles, PIM eligibility, enterprise-app assignments, and privileged/role-assignable groups.
- Reassign the former employee's direct reports and any application, security-group, SharePoint, Power Platform, Azure, or administrative ownership.
- Decide whether policy requires resetting the password to a random value. The script disables the on-premises AD account but does not reset its password.
- Review dynamic AD, Entra, and dynamic distribution-group rules. Dynamic memberships cannot be removed directly and may continue assigning licenses.
- Reassign Teams private/shared-channel roles and direct SharePoint/OneDrive permissions; standard Team membership follows its Microsoft 365 group.
- Configure mailbox forwarding or automatic replies if policy requires them. Full Access does not forward new mail.
- Apply OneDrive retention/delegation, device wipe/retirement, phone/Teams Calling cleanup, and line-of-business application offboarding.
- Follow legal-hold, retention, archive, and deletion policy before manually removing an Exchange-bearing license or deleting the account.
- Review all direct and group-based license assignments in the Microsoft 365 Admin Center. After successful shared-mailbox conversion and manager delegation, remove the intended direct license manually. Do not remove an Exchange-bearing license when the audit reports mailbox conversion, delegation, size, archive, or hold warnings.
- AD and Microsoft 365 group cleanup can indirectly revoke group-assigned licenses. If no license may change before manual review, protect known license-bearing AD groups with `-KeepADGroups` and use `-SkipCloudGroupRemoval`.
- After directory synchronization and license processing, verify the final license state.

## Validation status

The PowerShell parser check passed, and the executable script contains no Microsoft Graph module import, connection, or cmdlet. The script has not been run against live Active Directory, on-premises Exchange, or Exchange Online in this environment. Always use `-WhatIf` first.
