# Employee offboarding script

Put these two files in the same folder:

- `Invoke-EmployeeOffboarding.ps1`
- `Run-EmployeeOffboarding.bat`

## What this script does

Before changing anything, the script runs a read-only check. It finds the employee's on-premises AD account, Exchange Online group memberships, mailbox status, and the AD `manager` when mailbox access or group ownership transfer is needed. It does not make separate Graph requests to read the employee's profile, sessions, groups, or current license inventory.

After you review the results and confirm the offboarding, the script does the following:

1. Disables the employee's on-premises AD account and adds `TERM: MM/DD/YYYY` to AD Notes (`info`). It keeps any other text already in Notes. If Notes already has a `TERM:` line, the script replaces that line.
2. Removes the employee from directly assigned on-premises AD groups, except for the account's primary group and any groups listed in `-KeepADGroups`. By default, it keeps `Domain Users` and `SG_Intranet Users`.
3. Removes the employee from cloud-managed Exchange distribution groups and mail-enabled security groups. It also removes Microsoft 365 group membership and ownership through Exchange Online. It does not ask Graph to inspect or remove Entra-only security groups.
4. Asks `Need to create a shared mailbox (YES or NO)`. If you answer `YES`, it converts both sides of the hybrid mailbox and gives the AD manager Full Access. If you answer `NO`, it skips the mailbox conversion and does not give the manager mailbox access. You can also use `-GrantSendAs` with `YES`.
5. Uses Graph only to remove the one E3 product you select from a simple menu. You do not need to know or type its internal SKU GUID. Graph does not separately discover the employee's license or read whether it is direct versus group-assigned. Other licenses and add-ons remain for manual review. If you answer `YES`, the script must verify the mailbox conversion and the manager's Full Access before sending the removal. If you answer `NO`, it skips conversion and delegation. Removing an Exchange license in that case can deprovision the mailbox according to your organization's retention policy.
6. Saves a time-stamped CSV audit log in the same folder as the scripts. The log includes the shared-mailbox answer, termination-note value, selected license name, SKU part number, and SKU GUID.

You can run the script again if needed. When it can verify that an action is already complete, it reports that action as skipped.

## Prerequisites

- Run the BAT normally on the workstation you use to manage employees. You do not need to use local **Run as administrator** for delegated AD/Exchange work.
- The workstation needs Windows PowerShell 5.1 and the Active Directory RSAT module.
- The account running PowerShell needs `ExchangeOnlineManagement`. License-removal runs also need `Microsoft.Graph.Authentication` and `Microsoft.Graph.Users.Actions`.
- Your delegated AD account needs permission to disable users and remove the AD group memberships that apply.
- The Exchange account you sign in with needs permission to manage recipients, group memberships, mailbox types, and mailbox delegation.
- The Microsoft Graph account needs a role that can change licenses, such as **License Administrator**, plus admin approval for `LicenseAssignment.ReadWrite.All`. That is the only Graph data permission requested. Role assignment and Graph admin consent are separate requirements, so adding License Administrator alone does not approve the Graph prompt.
- The login page can also show **Maintain access to data you have given it access to** (`offline_access`). Microsoft authentication adds that sign-in scope automatically so the login can refresh its token; it does not grant access to any additional user, group, or license data.
- If you answer `YES` in a hybrid environment, enter the on-premises Exchange server FQDN. The delegated account uses it to open a remote Exchange PowerShell session. Unless you enter a full URL, the script expects the endpoint to be `http://SERVER/PowerShell/`. You do not need the server FQDN when you answer `NO`.

## Choosing the E3 license

When you run the BAT without either `-LicenseSkuId` or `-SkipLicenseRemoval`, it asks you to choose:

```text
1. Microsoft 365 E3 [SPE_E3]
2. Office 365 E3 [ENTERPRISEPACK]
3. Microsoft 365 E3 (no Teams) [Microsoft_365_E3_(no_Teams)]
4. Office 365 E3 (no Teams) [Office_365_E3_(no_Teams)]
S. Skip license removal for this run
```

Open the employee in the Microsoft 365 admin center and look under **Licenses and apps**. Match the full product name whose checkbox is checked for that employee: if it literally says **Microsoft 365 E3**, choose `1`; if it says **Office 365 E3**, choose `2`. If the checked product name does not exactly match one of the choices, choose `S` instead of guessing.

The script keeps Microsoft's published GUID behind each friendly name. `-LicenseSkuId` is still available as an advanced override for one externally verified SKU, but the normal BAT workflow does not need it. Microsoft publishes the mappings in its [license service-plan reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference).

## Preview first

Start with a preview. Replace the example values with the employee's real SAM account name, administrator UPNs, and Exchange server FQDN. The script will ask which E3 product to remove:

```bat
Run-EmployeeOffboarding.bat jsmith -OnPremExchangeServer exch01.contoso.local -ExchangeAdminUPN admin@contoso.com -GraphAdminUPN admin@contoso.com -WhatIf
```

Read every line of the preview before doing a live run. Pay special attention to the protected intranet group name and make sure it exactly matches the name in your environment.

`Inspecting Exchange Online group memberships...` can take several minutes. Without Graph group-read permission, Exchange has to check the groups individually. Use `-SkipCloudGroupRemoval` only when that cleanup is being handled somewhere else.

To run the rest of the preview with no Graph module or Graph login at all, use `-SkipLicenseRemoval` and leave out `-GraphAdminUPN`:

```bat
Run-EmployeeOffboarding.bat jsmith -OnPremExchangeServer exch01.contoso.local -ExchangeAdminUPN admin@contoso.com -SkipLicenseRemoval -WhatIf
```

If you leave out `-CreateSharedMailbox`, the script asks you to answer `YES` or `NO`. If you do not want the prompt, add either `-CreateSharedMailbox YES` or `-CreateSharedMailbox NO` to the command.

## Live run

When the preview looks correct, run the same command without `-WhatIf`:

```bat
Run-EmployeeOffboarding.bat jsmith -OnPremExchangeServer exch01.contoso.local -ExchangeAdminUPN admin@contoso.com -GraphAdminUPN admin@contoso.com
```

The script shows you the offboarding summary and, when needed, the manager it found. It then asks you to type this exact confirmation:

```text
OFFBOARD jsmith
```

Exit code `2` means the script finished, but the CSV log has items that still need follow-up. This version always flags Entra sign-in, cloud-session revocation, and Entra-only security groups for confirmation by the separate process. Any other nonzero exit code means an error stopped the script.

## Important switches

- `-CreateSharedMailbox YES|NO`: puts your mailbox answer directly in the command. If you leave it out, the script asks you.
- `-AdminUser DOMAIN\username`: gives the script your delegated AD/on-premises Exchange username. If you leave it out, the script asks for the username and then the password.
- `-TerminationDate YYYY-MM-DD`: uses the date you enter instead of today's date in the AD Notes marker. If you leave it out, the script uses the date on which it runs, such as `TERM: 08/17/2026`.
- `-GrantSendAs`: also gives the manager Send As permission. Full Access by itself does not let the manager send mail as the former employee.
- `-GraphAdminUPN`: checks that Microsoft Graph signed in with the administrator you intended to use, instead of quietly reusing a cached everyday-account sign-in.
- `-LicenseSkuId GUID`: optional advanced override for one externally verified SKU. Leave it out for the normal friendly-name menu. The script prints the selected product and GUID before confirmation but does not discover the employee's current assignment through Graph.
- `-TransferSoleOwnedMicrosoft365GroupsToManager`: if the employee is the only owner of a Microsoft 365 group, the script adds the manager as a member and owner before removing the employee. Without this switch, it leaves that group alone and flags it for attention.
- `-SkipCloudGroupRemoval`: skips cleanup of Exchange Online distribution and Microsoft 365 groups. The script can still remove on-premises AD groups, and those changes can sync to directory-synchronized cloud groups.
- `-SkipLicenseRemoval`: skips the selected license-removal request and completely skips Graph module loading and Graph login. Group cleanup can still remove a license that was assigned through a group.
- `-OverrideSharedMailboxLicenseSafety`: allows the script to submit the selected license even when the mailbox is over 50 GB, has an active archive, is on hold, or has a size the script cannot verify. Use this only after confirming the licensing and retention requirements.
- `-CloudOnlyMailbox`: when used with `-CreateSharedMailbox YES`, skips `Set-RemoteMailbox`. The script allows this only when Exchange Online explicitly reports that the mailbox is not directory-synchronized. A later sync could otherwise revert or disconnect a mismatched, unlicensed mailbox.
- `-Force`: skips the typed `OFFBOARD username` confirmation. It does not bypass permissions or any safety checks.

## Things you still need to check outside the script

The script cannot handle every system or policy. Make sure the termination ticket also covers any of these that apply:

- Reassign the former employee's direct reports. Also reassign any application, security-group, SharePoint, Power Platform, Azure, or administrative ownership.
- Confirm that the separate offboarding process blocked Microsoft Entra sign-in and revoked existing cloud sessions. This script no longer requests those Graph permissions.
- Check Microsoft Entra directory roles, PIM eligibility, enterprise-app assignments, and privileged or role-assignable groups.
- Check whether your policy requires changing the password to a random value. This script disables the AD account, but it does not reset the password.
- Check dynamic AD groups, dynamic Entra groups, and dynamic distribution-group rules. You cannot directly remove dynamic memberships, and those memberships may continue assigning licenses.
- Reassign Teams private/shared-channel roles and direct SharePoint/OneDrive permissions. Standard Team membership follows the Microsoft 365 group.
- Set up mailbox forwarding or automatic replies if your policy requires them. Full Access does not forward new mail.
- Handle OneDrive retention/delegation, device wipe or retirement, phone/Teams Calling cleanup, and offboarding from line-of-business applications.
- Follow your legal-hold, retention, archive, and deletion policies before you override the mailbox license safety check or delete the account.
- After directory synchronization and license processing finish, confirm that the intended SKU was removed and that no unwanted group-inherited license remains.

## Validation status

The PowerShell parser and static Graph-scope checks passed. The script has not been tested here against live Active Directory, on-premises Exchange, Exchange Online, or Microsoft Graph. Always run it with `-WhatIf` first.
