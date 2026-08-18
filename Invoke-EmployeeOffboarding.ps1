#Requires -Version 5.1
#Requires -Modules ActiveDirectory

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$User,

    [string]$AdminUser,
    [System.Management.Automation.PSCredential]$ADCredential,
    [string]$ExchangeAdminUPN,

    # Required for a directory-synchronized remote mailbox unless the on-premises
    # Exchange cmdlets are already loaded in this PowerShell session.
    [string]$OnPremExchangeServer,
    [switch]$CloudOnlyMailbox,

    [ValidateSet('YES', 'NO')]
    [string]$CreateSharedMailbox,
    [datetime]$TerminationDate = (Get-Date),

    [string[]]$KeepADGroups = @('Domain Users', 'SG_Intranet Users'),
    [switch]$GrantSendAs,
    [switch]$TransferSoleOwnedMicrosoft365GroupsToManager,
    [switch]$SkipCloudGroupRemoval,
    [switch]$Force,

    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        throw 'The script folder could not be determined. Run the saved PS1 file or use Run-EmployeeOffboarding.bat; do not run pasted or selected code.'
    }

    $LogPath = Join-Path -Path $PSScriptRoot -ChildPath ("EmployeeOffboarding-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
}

if (-not $User) {
    $User = Read-Host 'Employee to offboard (SAMAccountName)'
}

while ($CreateSharedMailbox -notin @('YES', 'NO')) {
    $CreateSharedMailbox = (Read-Host 'Need to create a shared mailbox (YES or NO)').Trim().ToUpperInvariant()
    if ($CreateSharedMailbox -notin @('YES', 'NO')) {
        Write-Warning 'Enter exactly YES or NO.'
    }
}

$CreateSharedMailbox = $CreateSharedMailbox.ToUpperInvariant()
$CreateSharedMailboxRequested = $CreateSharedMailbox -eq 'YES'
$TerminationNote = 'TERM: {0:MM/dd/yyyy}' -f $TerminationDate

if (-not $CreateSharedMailboxRequested -and $CloudOnlyMailbox) {
    throw '-CloudOnlyMailbox is only meaningful when -CreateSharedMailbox YES is selected.'
}
if (-not $CreateSharedMailboxRequested -and $GrantSendAs) {
    throw '-GrantSendAs requires -CreateSharedMailbox YES.'
}

if (-not $ADCredential) {
    while ([string]::IsNullOrWhiteSpace($AdminUser)) {
        $AdminUser = Read-Host 'Delegated AD / on-premises Exchange username (DOMAIN\username)'
    }
    Write-Host "`nActive Directory / on-premises Exchange credential: $AdminUser" -ForegroundColor Cyan
    $SecurePassword = Read-Host 'Enter the delegated account password' -AsSecureString
    $ADCredential = [System.Management.Automation.PSCredential]::new($AdminUser, $SecurePassword)
}

$Audit = [System.Collections.Generic.List[object]]::new()
$Counts = @{
    Completed = 0
    Failed    = 0
    Skipped   = 0
    Preview   = 0
    Attention = 0
}
$ManualLicenseRemovalReady = $true
$OnPremExchangeSession = $null
$ImportedOnPremExchangeModule = $null

function Add-AuditRecord {
    param(
        [string]$System,
        [string]$Item,
        [string]$Action,
        [string]$Status,
        [string]$Message = ''
    )

    $script:Audit.Add([pscustomobject]@{
        Timestamp              = Get-Date
        User                   = $script:User
        SharedMailboxRequested = $script:CreateSharedMailbox
        TerminationNote        = $script:TerminationNote
        System                 = $System
        Item                   = $Item
        Action                 = $Action
        Status                 = $Status
        Message                = $Message
    })
}

function Add-Result {
    param(
        [string]$System,
        [string]$Item,
        [string]$Action,
        [ValidateSet('Completed', 'Failed', 'Skipped', 'Preview', 'Attention')]
        [string]$Status,
        [string]$Message = ''
    )

    $script:Counts[$Status]++
    Add-AuditRecord -System $System -Item $Item -Action $Action -Status $Status -Message $Message

    $Color = switch ($Status) {
        'Completed' { 'Green' }
        'Failed'    { 'Red' }
        'Attention' { 'Yellow' }
        'Preview'   { 'Yellow' }
        default     { 'Gray' }
    }
    $Suffix = if ($Message) { ": $Message" } else { '' }
    Write-Host "  [$Status] $Action - $Item$Suffix" -ForegroundColor $Color
}

function Write-Section {
    param([string]$Title)

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor Cyan
}

function Get-NormalizedAddress {
    param($Value)

    if ($null -eq $Value) { return $null }
    return $Value.ToString().Trim().ToLowerInvariant()
}

function Resolve-EXORecipient {
    param(
        [Microsoft.ActiveDirectory.Management.ADUser]$ADUser,
        [string]$Role
    )

    $Candidates = @($ADUser.EmailAddress, $ADUser.UserPrincipalName) |
        Where-Object { $_ } |
        Select-Object -Unique

    foreach ($Candidate in $Candidates) {
        try {
            return Get-EXORecipient -Identity $Candidate -Properties PrimarySmtpAddress, RecipientTypeDetails -ErrorAction Stop
        }
        catch {
            # Try the next unambiguous identity.
        }
    }

    throw "Could not resolve the $Role '$($ADUser.SamAccountName)' in Exchange Online."
}

function ConvertTo-ByteCount {
    param($TotalItemSize)

    if ($null -eq $TotalItemSize) { return $null }

    try {
        return [int64]$TotalItemSize.Value.ToBytes()
    }
    catch {
        # REST-backed Exchange output can be serialized as text instead.
    }

    $Text = $TotalItemSize.ToString()
    if ($Text -match '\(([\d,]+) bytes\)') {
        return [int64]($Matches[1] -replace ',', '')
    }

    if ($Text -match '^\s*([\d.]+)\s*(KB|MB|GB|TB)') {
        $Multiplier = switch ($Matches[2].ToUpperInvariant()) {
            'KB' { 1KB }
            'MB' { 1MB }
            'GB' { 1GB }
            'TB' { 1TB }
        }
        return [int64]([double]$Matches[1] * $Multiplier)
    }

    return $null
}

function Test-PrimaryADGroup {
    param(
        [Microsoft.ActiveDirectory.Management.ADGroup]$Group,
        [int]$PrimaryGroupID
    )

    if (-not $Group.SID) { return $false }
    $GroupRid = [int]($Group.SID.Value.Split('-')[-1])
    return $GroupRid -eq $PrimaryGroupID
}

Write-Section 'Preflight - resolve identities and required tools'

try {
    $EmployeeAD = Get-ADUser -Identity $User -Properties DisplayName, EmailAddress, UserPrincipalName, Manager, Enabled, PrimaryGroupID, info -Credential $ADCredential
}
catch {
    throw "Unable to resolve the employee in Active Directory: $($_.Exception.Message)"
}

$User = $EmployeeAD.SamAccountName
$ExistingADNotes = [string]$EmployeeAD.info
$TerminationLinePattern = [regex]::new('(?im)^TERM:\s*.*$')
if ($TerminationLinePattern.IsMatch($ExistingADNotes)) {
    $UpdatedADNotes = $TerminationLinePattern.Replace($ExistingADNotes, $TerminationNote, 1)
}
elseif ([string]::IsNullOrWhiteSpace($ExistingADNotes)) {
    $UpdatedADNotes = $TerminationNote
}
else {
    $UpdatedADNotes = $ExistingADNotes.TrimEnd() + "`r`n" + $TerminationNote
}

if ($UpdatedADNotes.Length -gt 1024) {
    throw "The AD Notes value would exceed 1,024 characters after adding '$TerminationNote'. Shorten the existing Notes value first."
}

$ManagerRequired = $CreateSharedMailboxRequested -or $TransferSoleOwnedMicrosoft365GroupsToManager
$ManagerAD = $null

if ($ManagerRequired) {
    if (-not $EmployeeAD.Manager) {
        throw "The AD manager attribute is empty for '$User'. A manager is required for the selected mailbox or group-ownership action."
    }

    try {
        $ManagerAD = Get-ADUser -Identity $EmployeeAD.Manager -Properties DisplayName, EmailAddress, UserPrincipalName, Enabled -Credential $ADCredential
    }
    catch {
        throw "Unable to resolve the manager stored in AD: $($_.Exception.Message)"
    }

    if (-not $ManagerAD.Enabled) {
        throw "The manager '$($ManagerAD.SamAccountName)' is disabled. Assign an active manager before offboarding."
    }

    if ($EmployeeAD.DistinguishedName -eq $ManagerAD.DistinguishedName) {
        throw 'The employee cannot be their own manager.'
    }
}

$RequiredModules = @(
    'ExchangeOnlineManagement'
)

foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        throw "Required module '$Module' is not installed. Install it before running this offboarding script."
    }
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

if ($CreateSharedMailboxRequested -and -not $CloudOnlyMailbox) {
    if ($OnPremExchangeServer) {
        $ConnectionUri = if ($OnPremExchangeServer -match '^https?://.+/PowerShell/?$') {
            $OnPremExchangeServer.TrimEnd('/') + '/'
        }
        elseif ($OnPremExchangeServer -match '^https?://') {
            $OnPremExchangeServer.TrimEnd('/') + '/PowerShell/'
        }
        else {
            "http://$OnPremExchangeServer/PowerShell/"
        }

        try {
            $OnPremExchangeSession = New-PSSession `
                -ConfigurationName Microsoft.Exchange `
                -ConnectionUri $ConnectionUri `
                -Authentication Kerberos `
                -Credential $ADCredential `
                -ErrorAction Stop

            $ImportedOnPremExchangeModule = Import-PSSession `
                -Session $OnPremExchangeSession `
                -CommandName Get-RemoteMailbox, Set-RemoteMailbox `
                -AllowClobber `
                -DisableNameChecking `
                -ErrorAction Stop
        }
        catch {
            if ($OnPremExchangeSession) { Remove-PSSession $OnPremExchangeSession -ErrorAction SilentlyContinue }
            throw "Could not connect to on-premises Exchange at '$ConnectionUri': $($_.Exception.Message)"
        }
    }
    elseif (-not (Get-Command Get-RemoteMailbox -ErrorAction SilentlyContinue) -or
            -not (Get-Command Set-RemoteMailbox -ErrorAction SilentlyContinue)) {
        throw 'Shared-mailbox conversion in hybrid requires the on-premises Exchange cmdlets. Supply -OnPremExchangeServer <server FQDN>, run from an authenticated Exchange Management Shell, or use -CloudOnlyMailbox only when the mailbox is not directory-synchronized.'
    }
}

try {
    Get-OrganizationConfig -ErrorAction Stop | Out-Null
    Write-Host '  Using the existing Exchange Online connection.' -ForegroundColor Green
}
catch {
    if ($ExchangeAdminUPN) {
        Connect-ExchangeOnline -UserPrincipalName $ExchangeAdminUPN -ShowBanner:$false -ErrorAction Stop
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
    Write-Host '  Connected to Exchange Online.' -ForegroundColor Green
}

$EmployeeEXO = Resolve-EXORecipient -ADUser $EmployeeAD -Role 'employee'
$EmployeeAddress = Get-NormalizedAddress $EmployeeEXO.PrimarySmtpAddress
$ManagerEXO = $null
$ManagerAddress = $null
if ($ManagerRequired) {
    $ManagerEXO = Resolve-EXORecipient -ADUser $ManagerAD -Role 'manager'
    $ManagerAddress = Get-NormalizedAddress $ManagerEXO.PrimarySmtpAddress
}

try {
    $Mailbox = Get-EXOMailbox `
        -Identity $EmployeeAddress `
        -Properties RecipientTypeDetails, LitigationHoldEnabled, InPlaceHolds, ArchiveStatus `
        -ErrorAction Stop
    $MailboxStatistics = Get-EXOMailboxStatistics -Identity $EmployeeAddress -Properties TotalItemSize -ErrorAction Stop
}
catch {
    throw "Could not resolve and inspect the employee mailbox: $($_.Exception.Message)"
}

if ($CreateSharedMailboxRequested -and $CloudOnlyMailbox) {
    try {
        $MailboxDirectoryState = Get-Mailbox -Identity $EmployeeAddress -ErrorAction Stop
    }
    catch {
        throw "Could not verify whether '$EmployeeAddress' is directory-synchronized: $($_.Exception.Message)"
    }

    $IsDirSyncedProperty = $MailboxDirectoryState.PSObject.Properties['IsDirSynced']
    if (-not $IsDirSyncedProperty -or $IsDirSyncedProperty.Value -isnot [bool]) {
        throw "-CloudOnlyMailbox cannot be used because Exchange Online did not return the mailbox directory-synchronization state. Supply -OnPremExchangeServer so both mailbox objects are converted safely."
    }
    if ($IsDirSyncedProperty.Value) {
        throw "-CloudOnlyMailbox cannot be used because '$EmployeeAddress' is directory-synchronized. Supply -OnPremExchangeServer so both mailbox objects are converted safely."
    }
}

$RemoteMailbox = $null
if ($CreateSharedMailboxRequested -and -not $CloudOnlyMailbox) {
    try {
        $RemoteMailbox = Get-RemoteMailbox -Identity $EmployeeAD.DistinguishedName -ErrorAction Stop
    }
    catch {
        throw "Could not resolve the on-premises remote-mailbox object for '$User': $($_.Exception.Message)"
    }
}

try {
    $ADGroups = @(Get-ADPrincipalGroupMembership -Identity $EmployeeAD -Credential $ADCredential)
}
catch {
    throw "Could not enumerate the employee's Active Directory groups: $($_.Exception.Message)"
}

$MailboxSizeBytes = ConvertTo-ByteCount $MailboxStatistics.TotalItemSize
$LicenseBlockers = [System.Collections.Generic.List[string]]::new()
if ($Mailbox.LitigationHoldEnabled -or @($Mailbox.InPlaceHolds | Where-Object { $_ }).Count -gt 0) {
    $LicenseBlockers.Add('mailbox is on hold')
}
if ($Mailbox.ArchiveStatus -eq 'Active') {
    $LicenseBlockers.Add('archive mailbox is active')
}
if ($CreateSharedMailboxRequested) {
    if ($null -eq $MailboxSizeBytes) {
        $LicenseBlockers.Add('mailbox size could not be verified')
    }
    elseif ($MailboxSizeBytes -gt 50GB) {
        $LicenseBlockers.Add('mailbox is larger than 50 GB')
    }
}

$CloudDistributionGroups = [System.Collections.Generic.List[object]]::new()
$DistributionGroupInspectionFailures = [System.Collections.Generic.List[object]]::new()
$Microsoft365GroupPlans = [System.Collections.Generic.List[object]]::new()
$Microsoft365GroupInspectionFailures = [System.Collections.Generic.List[object]]::new()

if (-not $SkipCloudGroupRemoval) {
    Write-Host '  Inspecting Exchange Online group memberships...' -ForegroundColor Gray

    try {
        foreach ($Group in @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)) {
            if ($Group.IsDirSynced) { continue }

            try {
                $IsMember = @(Get-DistributionGroupMember -Identity $Group.Identity -ResultSize Unlimited -ErrorAction Stop |
                    Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $EmployeeAddress }).Count -gt 0

                if ($IsMember) { $CloudDistributionGroups.Add($Group) }
            }
            catch {
                $DistributionGroupInspectionFailures.Add([pscustomobject]@{
                    Group   = $Group.DisplayName
                    Message = $_.Exception.Message
                })
            }
        }
    }
    catch {
        throw "Could not enumerate Exchange Online distribution groups: $($_.Exception.Message)"
    }

    try {
        $Microsoft365Groups = @(Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop)
    }
    catch {
        throw "Could not enumerate Exchange Online Microsoft 365 groups: $($_.Exception.Message)"
    }

    foreach ($Group in $Microsoft365Groups) {
        try {
            $GroupIdentity = Get-NormalizedAddress $Group.PrimarySmtpAddress
            if (-not $GroupIdentity) {
                throw 'Exchange Online did not return a primary SMTP address for the group.'
            }

            $Owners = @(Get-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
            $Members = @(Get-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
            $IsOwner = @($Owners | Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $EmployeeAddress }).Count -gt 0
            $IsMember = @($Members | Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $EmployeeAddress }).Count -gt 0
            if (-not $IsOwner -and -not $IsMember) { continue }

            $ManagerIsMember = $false
            $ManagerIsOwner = $false
            if ($ManagerAddress) {
                $ManagerIsMember = @($Members | Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $ManagerAddress }).Count -gt 0
                $ManagerIsOwner = @($Owners | Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $ManagerAddress }).Count -gt 0
            }

            $Microsoft365GroupPlans.Add([pscustomobject]@{
                Group           = $Group
                Identity        = $GroupIdentity
                Owners          = $Owners
                IsOwner         = $IsOwner
                IsMember        = $IsMember
                ManagerIsMember = $ManagerIsMember
                ManagerIsOwner  = $ManagerIsOwner
            })
        }
        catch {
            $Microsoft365GroupInspectionFailures.Add([pscustomobject]@{
                Group   = $Group.DisplayName
                Message = $_.Exception.Message
            })
        }
    }
}

Write-Host "`nEmployee : $($EmployeeAD.DisplayName) [$User]" -ForegroundColor White
if ($ManagerAD) {
    Write-Host "Manager  : $($ManagerAD.DisplayName) [$($ManagerAD.SamAccountName)]" -ForegroundColor White
}
else {
    Write-Host 'Manager  : Not required for the selected actions' -ForegroundColor Gray
}
Write-Host "Mailbox  : $EmployeeAddress" -ForegroundColor White
Write-Host "Shared   : $CreateSharedMailbox" -ForegroundColor White
Write-Host "AD Notes : $TerminationNote" -ForegroundColor White
Write-Host "Keep AD  : $($KeepADGroups -join ', ')" -ForegroundColor White
Write-Host "AD groups: $($ADGroups.Count) direct membership(s) found" -ForegroundColor White
if (-not $SkipCloudGroupRemoval) {
    Write-Host "Exchange : $($CloudDistributionGroups.Count) distribution and $($Microsoft365GroupPlans.Count) Microsoft 365 group membership/ownership record(s) found" -ForegroundColor White
    Write-Warning 'Cloud-only, non-mail-enabled Microsoft Entra security groups are not inspected by this Graph-free branch and require manual review.'
}
if ($DistributionGroupInspectionFailures.Count -gt 0) {
    Write-Warning "Preflight could not inspect $($DistributionGroupInspectionFailures.Count) Exchange distribution group(s): $(@($DistributionGroupInspectionFailures.Group) -join ', '). These will be flagged for manual follow-up."
}
if ($Microsoft365GroupInspectionFailures.Count -gt 0) {
    Write-Warning "Preflight could not inspect $($Microsoft365GroupInspectionFailures.Count) Microsoft 365 group(s): $(@($Microsoft365GroupInspectionFailures.Group) -join ', '). These will be flagged for manual follow-up."
}
Write-Warning 'Microsoft 365 licenses are not inventoried and no license-assignment API is called. AD or Microsoft 365 group removal can still revoke group-assigned licenses.'
if ($CreateSharedMailboxRequested) {
    Write-Host 'Mailbox action: convert to shared and grant the AD manager Full Access. Direct license removal remains manual.' -ForegroundColor Yellow
}
else {
    Write-Warning 'Shared mailbox is NO. The script will not convert, delegate, or directly change license assignments. Group cleanup can revoke inherited licenses; manually removing an Exchange-bearing license can deprovision the mailbox according to your retention policy.'
}

if ($LicenseBlockers.Count -gt 0) {
    Write-Warning "Manual license-removal warning: $($LicenseBlockers -join '; '). Review these conditions and confirm the intended product before removing it in the Microsoft 365 Admin Center."
}

if (-not $Force -and -not $WhatIfPreference) {
    $Answer = Read-Host "Continue? Type OFFBOARD $User to make these changes"
    if ($Answer -cne "OFFBOARD $User") {
        Write-Host 'Cancelled. No changes were made.' -ForegroundColor Yellow
        return
    }
}

Write-Section 'Step 1 - disable Active Directory account'

if (-not $EmployeeAD.Enabled) {
    Add-Result -System 'Active Directory' -Item $User -Action 'Disable account' -Status 'Skipped' -Message 'Account is already disabled.'
}
elseif ($PSCmdlet.ShouldProcess($User, 'Disable the Active Directory account')) {
    try {
        Disable-ADAccount -Identity $EmployeeAD -Credential $ADCredential -ErrorAction Stop
        Add-Result -System 'Active Directory' -Item $User -Action 'Disable account' -Status 'Completed'
    }
    catch {
        Add-Result -System 'Active Directory' -Item $User -Action 'Disable account' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
    Add-Result -System 'Active Directory' -Item $User -Action 'Disable account' -Status 'Preview'
}

if ($ExistingADNotes -ceq $UpdatedADNotes) {
    Add-Result -System 'Active Directory' -Item $User -Action 'Set AD Notes termination marker' -Status 'Skipped' -Message "$TerminationNote is already present."
}
elseif ($PSCmdlet.ShouldProcess($User, "Set the AD Notes termination marker to '$TerminationNote'")) {
    try {
        Set-ADUser -Identity $EmployeeAD -Replace @{ info = $UpdatedADNotes } -Credential $ADCredential -ErrorAction Stop
        $VerifiedADNotes = [string](Get-ADUser -Identity $EmployeeAD -Properties info -Credential $ADCredential -ErrorAction Stop).info
        if ($VerifiedADNotes -cne $UpdatedADNotes) {
            throw 'The AD Notes value did not match after the update.'
        }
        Add-Result -System 'Active Directory' -Item $User -Action 'Set AD Notes termination marker' -Status 'Completed' -Message $TerminationNote
    }
    catch {
        Add-Result -System 'Active Directory' -Item $User -Action 'Set AD Notes termination marker' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
    Add-Result -System 'Active Directory' -Item $User -Action 'Set AD Notes termination marker' -Status 'Preview' -Message $TerminationNote
}

Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Block cloud sign-in manually' -Status 'Attention' -Message 'This Graph-free branch cannot verify or change Entra sign-in status. Block sign-in in the Microsoft 365 or Entra admin center; disabling AD may not take effect in the cloud until directory synchronization finishes.'
Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Revoke cloud sign-in sessions manually' -Status 'Attention' -Message 'This Graph-free branch cannot revoke existing cloud sessions. Revoke sessions in the Microsoft 365 or Entra admin center.'

Write-Section 'Step 2 - remove Active Directory groups'

foreach ($Group in $ADGroups) {
    $IsPrimaryGroup = Test-PrimaryADGroup -Group $Group -PrimaryGroupID $EmployeeAD.PrimaryGroupID
    if ($IsPrimaryGroup -or $Group.Name -in $KeepADGroups) {
        $Reason = if ($IsPrimaryGroup) { 'Primary/protected group.' } else { 'Protected by KeepADGroups.' }
        Add-Result -System 'Active Directory' -Item $Group.Name -Action 'Remove group membership' -Status 'Skipped' -Message $Reason
        continue
    }

    if ($PSCmdlet.ShouldProcess($Group.Name, "Remove $User from the Active Directory group")) {
        try {
            Remove-ADGroupMember -Identity $Group.DistinguishedName -Members $EmployeeAD -Credential $ADCredential -Confirm:$false -ErrorAction Stop
            Add-Result -System 'Active Directory' -Item $Group.Name -Action 'Remove group membership' -Status 'Completed'
        }
        catch {
            Add-Result -System 'Active Directory' -Item $Group.Name -Action 'Remove group membership' -Status 'Failed' -Message $_.Exception.Message
        }
    }
    else {
        Add-Result -System 'Active Directory' -Item $Group.Name -Action 'Remove group membership' -Status 'Preview'
    }
}

Write-Section 'Step 3 - remove Exchange Online group memberships'

if ($SkipCloudGroupRemoval) {
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Remove Exchange-manageable group memberships' -Status 'Skipped' -Message 'SkipCloudGroupRemoval was specified. Entra-only security groups still require manual review.'
}
else {
    foreach ($Failure in $DistributionGroupInspectionFailures) {
        Add-Result -System 'Exchange Online DL' -Item $Failure.Group -Action 'Inspect membership' -Status 'Attention' -Message $Failure.Message
    }

    foreach ($Failure in $Microsoft365GroupInspectionFailures) {
        Add-Result -System 'Microsoft 365 Group' -Item $Failure.Group -Action 'Inspect membership and ownership' -Status 'Attention' -Message $Failure.Message
    }

    foreach ($Group in $CloudDistributionGroups) {
        if ($PSCmdlet.ShouldProcess($Group.DisplayName, "Remove $EmployeeAddress from the distribution group")) {
            try {
                Remove-DistributionGroupMember `
                    -Identity $Group.Identity `
                    -Member $EmployeeAddress `
                    -BypassSecurityGroupManagerCheck `
                    -Confirm:$false `
                    -ErrorAction Stop
                Add-Result -System 'Exchange Online DL' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Completed'
            }
            catch {
                Add-Result -System 'Exchange Online DL' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Failed' -Message $_.Exception.Message
            }
        }
        else {
            Add-Result -System 'Exchange Online DL' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Preview'
        }
    }

    foreach ($Plan in $Microsoft365GroupPlans) {
        $Group = $Plan.Group
        $OnlyOwner = $Plan.IsOwner -and $Plan.Owners.Count -le 1

        if ($OnlyOwner -and -not $TransferSoleOwnedMicrosoft365GroupsToManager) {
            Add-Result `
                -System 'Microsoft 365 Group' `
                -Item $Group.DisplayName `
                -Action 'Remove membership and ownership' `
                -Status 'Attention' `
                -Message 'Employee is the only owner. Re-run with -TransferSoleOwnedMicrosoft365GroupsToManager or assign another owner manually.'
            continue
        }

        if ($PSCmdlet.ShouldProcess($Group.DisplayName, "Remove $EmployeeAddress from the Microsoft 365 group")) {
            try {
                if ($OnlyOwner) {
                    if (-not $Plan.ManagerIsMember) {
                        Add-UnifiedGroupLinks -Identity $Plan.Identity -LinkType Members -Links $ManagerAddress -Confirm:$false -ErrorAction Stop
                    }
                    if (-not $Plan.ManagerIsOwner) {
                        Add-UnifiedGroupLinks -Identity $Plan.Identity -LinkType Owners -Links $ManagerAddress -Confirm:$false -ErrorAction Stop
                    }
                }

                if ($Plan.IsOwner) {
                    Remove-UnifiedGroupLinks -Identity $Plan.Identity -LinkType Owners -Links $EmployeeAddress -Confirm:$false -ErrorAction Stop
                }
                if ($Plan.IsMember) {
                    Remove-UnifiedGroupLinks -Identity $Plan.Identity -LinkType Members -Links $EmployeeAddress -Confirm:$false -ErrorAction Stop
                }

                $Message = if ($OnlyOwner) { "Ownership transferred to $ManagerAddress." } else { '' }
                Add-Result -System 'Microsoft 365 Group' -Item $Group.DisplayName -Action 'Remove membership and ownership' -Status 'Completed' -Message $Message
            }
            catch {
                Add-Result -System 'Microsoft 365 Group' -Item $Group.DisplayName -Action 'Remove membership and ownership' -Status 'Failed' -Message $_.Exception.Message
            }
        }
        else {
            $Message = if ($OnlyOwner) { "Would transfer ownership to $ManagerAddress first." } else { '' }
            Add-Result -System 'Microsoft 365 Group' -Item $Group.DisplayName -Action 'Remove membership and ownership' -Status 'Preview' -Message $Message
        }
    }

    Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Review cloud-only security-group memberships manually' -Status 'Attention' -Message 'This Graph-free branch cannot enumerate or remove non-mail-enabled Entra security groups, including dynamic and role-assignable groups. Review them in the Entra admin center.'
}

Write-Section 'Step 4 - convert mailbox and delegate the manager'

if (-not $CreateSharedMailboxRequested) {
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Create shared mailbox and grant manager access' -Status 'Skipped' -Message 'Operator selected NO.'
}
else {
$OnPremMailboxReady = $true
if ($CloudOnlyMailbox) {
    Add-Result -System 'On-premises Exchange' -Item $User -Action 'Convert remote mailbox to shared' -Status 'Skipped' -Message 'CloudOnlyMailbox was explicitly specified.'
}
elseif ($RemoteMailbox.RecipientTypeDetails -eq 'RemoteSharedMailbox') {
    Add-Result -System 'On-premises Exchange' -Item $User -Action 'Convert remote mailbox to shared' -Status 'Skipped' -Message 'Remote mailbox is already shared.'
}
elseif ($PSCmdlet.ShouldProcess($User, 'Convert the on-premises remote mailbox to Shared')) {
    try {
        Set-RemoteMailbox -Identity $EmployeeAD.DistinguishedName -Type Shared -Confirm:$false -ErrorAction Stop
        $VerifiedRemoteMailbox = Get-RemoteMailbox -Identity $EmployeeAD.DistinguishedName -ErrorAction Stop
        if ($VerifiedRemoteMailbox.RecipientTypeDetails -ne 'RemoteSharedMailbox') {
            throw "Conversion did not verify; recipient type is '$($VerifiedRemoteMailbox.RecipientTypeDetails)'."
        }
        Add-Result -System 'On-premises Exchange' -Item $User -Action 'Convert remote mailbox to shared' -Status 'Completed'
    }
    catch {
        $OnPremMailboxReady = $false
        $ManualLicenseRemovalReady = $false
        Add-Result -System 'On-premises Exchange' -Item $User -Action 'Convert remote mailbox to shared' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
    $ManualLicenseRemovalReady = $false
    Add-Result -System 'On-premises Exchange' -Item $User -Action 'Convert remote mailbox to shared' -Status 'Preview'
}

if (-not $OnPremMailboxReady) {
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Convert to shared mailbox' -Status 'Attention' -Message 'Skipped because the on-premises remote-mailbox conversion failed.'
}
elseif ($Mailbox.RecipientTypeDetails -eq 'SharedMailbox') {
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Convert to shared mailbox' -Status 'Skipped' -Message 'Mailbox is already shared.'
}
elseif ($PSCmdlet.ShouldProcess($EmployeeAddress, 'Convert the Exchange Online mailbox to Shared')) {
    try {
        Set-Mailbox -Identity $EmployeeAddress -Type Shared -Confirm:$false -ErrorAction Stop
        $VerifiedMailbox = Get-EXOMailbox -Identity $EmployeeAddress -Properties RecipientTypeDetails -ErrorAction Stop
        if ($VerifiedMailbox.RecipientTypeDetails -ne 'SharedMailbox') {
            throw "Conversion did not verify; recipient type is '$($VerifiedMailbox.RecipientTypeDetails)'."
        }
        Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Convert to shared mailbox' -Status 'Completed'
    }
    catch {
        $ManualLicenseRemovalReady = $false
        Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Convert to shared mailbox' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
    $ManualLicenseRemovalReady = $false
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Convert to shared mailbox' -Status 'Preview'
}

try {
    $ExistingFullAccess = @(Get-MailboxPermission -Identity $EmployeeAddress -User $ManagerAddress -ErrorAction SilentlyContinue |
        Where-Object { -not $_.Deny -and 'FullAccess' -in $_.AccessRights }).Count -gt 0

    if ($ExistingFullAccess) {
        Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant mailbox Full Access' -Status 'Skipped' -Message 'Permission already exists.'
    }
    elseif ($PSCmdlet.ShouldProcess($EmployeeAddress, "Grant Full Access to $ManagerAddress")) {
        Add-MailboxPermission `
            -Identity $EmployeeAddress `
            -User $ManagerAddress `
            -AccessRights FullAccess `
            -InheritanceType All `
            -AutoMapping:$true `
            -Confirm:$false `
            -ErrorAction Stop | Out-Null

        $VerifiedPermission = @(Get-MailboxPermission -Identity $EmployeeAddress -User $ManagerAddress -ErrorAction Stop |
            Where-Object { -not $_.Deny -and 'FullAccess' -in $_.AccessRights }).Count -gt 0
        if (-not $VerifiedPermission) { throw 'Full Access permission could not be verified.' }
        Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant mailbox Full Access' -Status 'Completed'
    }
    else {
        $ManualLicenseRemovalReady = $false
        Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant mailbox Full Access' -Status 'Preview'
    }
}
catch {
    $ManualLicenseRemovalReady = $false
    Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant mailbox Full Access' -Status 'Failed' -Message $_.Exception.Message
}

if ($GrantSendAs) {
    try {
        $ExistingSendAs = @(Get-RecipientPermission -Identity $EmployeeAddress -Trustee $ManagerAddress -ErrorAction SilentlyContinue |
            Where-Object { -not $_.Deny -and 'SendAs' -in $_.AccessRights }).Count -gt 0

        if ($ExistingSendAs) {
            Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant Send As' -Status 'Skipped' -Message 'Permission already exists.'
        }
        elseif ($PSCmdlet.ShouldProcess($EmployeeAddress, "Grant Send As to $ManagerAddress")) {
            Add-RecipientPermission -Identity $EmployeeAddress -Trustee $ManagerAddress -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant Send As' -Status 'Completed'
        }
        else {
            Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant Send As' -Status 'Preview'
        }
    }
    catch {
        Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant Send As' -Status 'Failed' -Message $_.Exception.Message
    }
}
}

if ($LicenseBlockers.Count -gt 0) {
    $ManualLicenseRemovalReady = $false
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Review manual license removal' -Status 'Attention' -Message ($LicenseBlockers -join '; ')
}

Write-Section 'Step 5 - review Microsoft 365 licenses manually'

if (-not $CreateSharedMailboxRequested) {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Review and remove licenses manually' -Status 'Attention' -Message "License assignments were not inventoried and no license-assignment API was called. Shared mailbox was NO; removing an Exchange-bearing license can deprovision the mailbox, so follow the organization's retention decision. Group cleanup may also revoke group-assigned licenses."
}
elseif (-not $ManualLicenseRemovalReady) {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Review and remove licenses manually' -Status 'Attention' -Message 'License assignments were not inventoried and no license-assignment API was called. Resolve the mailbox conversion, delegation, or licensing warnings before removing an Exchange-bearing license in the Microsoft 365 Admin Center. Group cleanup may also revoke group-assigned licenses.'
}
else {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Review and remove licenses manually' -Status 'Attention' -Message 'License assignments were not inventoried and no license-assignment API was called. Confirm the intended product and remove it manually in the Microsoft 365 Admin Center after reviewing the audit log. Group cleanup may also revoke group-assigned licenses.'
}

$LogDirectory = Split-Path -Parent $LogPath
if ($LogDirectory -and -not (Test-Path -LiteralPath $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}
$Audit | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding UTF8

if ($ImportedOnPremExchangeModule) {
    Remove-Module $ImportedOnPremExchangeModule.Name -Force -ErrorAction SilentlyContinue
}
if ($OnPremExchangeSession) {
    Remove-PSSession $OnPremExchangeSession -ErrorAction SilentlyContinue
}

Write-Section 'Summary'
Write-Host "  Completed      : $($Counts.Completed)"
Write-Host "  Preview only   : $($Counts.Preview)"
Write-Host "  Skipped        : $($Counts.Skipped)"
Write-Host "  Needs attention: $($Counts.Attention)"
Write-Host "  Failed         : $($Counts.Failed)"
Write-Host "  Audit log      : $LogPath" -ForegroundColor Cyan

if ($Counts.Failed -gt 0 -or $Counts.Attention -gt 0) {
    Write-Warning 'Offboarding completed with failures or follow-up items. Review the audit log before closing the ticket.'
    exit 2
}
