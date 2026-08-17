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
    [string]$GraphAdminUPN,

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
    [switch]$SkipLicenseRemoval,
    [switch]$OverrideSharedMailboxLicenseSafety,
    [switch]$Force,

    [string]$LogPath = (Join-Path $PSScriptRoot ("EmployeeOffboarding-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date)))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
$CanRemoveLicenses = $true
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
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Users.Actions'
    'Microsoft.Graph.Groups'
)

foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        throw "Required module '$Module' is not installed. Install it before running this offboarding script."
    }
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
Import-Module Microsoft.Graph.Groups -ErrorAction Stop

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

$RequiredScopes = @(
    'User.Read.All'
    'User.EnableDisableAccount.All'
    'User.RevokeSessions.All'
    'LicenseAssignment.ReadWrite.All'
    'GroupMember.ReadWrite.All'
)
$GraphContext = Get-MgContext
$MissingScopes = if ($GraphContext) {
    @($RequiredScopes | Where-Object { $_ -notin $GraphContext.Scopes })
}
else {
    @($RequiredScopes)
}

if (-not $GraphContext -or $GraphContext.ContextScope -ne 'Process' -or $MissingScopes.Count -gt 0 -or
        ($GraphAdminUPN -and $GraphContext.Account -ine $GraphAdminUPN)) {
    Connect-MgGraph -Scopes $RequiredScopes -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
    $GraphContext = Get-MgContext
}

if ($GraphAdminUPN -and $GraphContext.Account -ine $GraphAdminUPN) {
    throw "Microsoft Graph connected as '$($GraphContext.Account)', but -GraphAdminUPN requires '$GraphAdminUPN'. Re-run and choose the required Graph administrator account during sign-in."
}

Write-Host "  Microsoft Graph account: $($GraphContext.Account)" -ForegroundColor Green

$EmployeeEXO = Resolve-EXORecipient -ADUser $EmployeeAD -Role 'employee'
$EmployeeAddress = Get-NormalizedAddress $EmployeeEXO.PrimarySmtpAddress
$ManagerEXO = $null
$ManagerAddress = $null
if ($ManagerRequired) {
    $ManagerEXO = Resolve-EXORecipient -ADUser $ManagerAD -Role 'manager'
    $ManagerAddress = Get-NormalizedAddress $ManagerEXO.PrimarySmtpAddress
}

try {
    $GraphUser = Get-MgUser `
        -UserId $EmployeeAD.UserPrincipalName `
        -Property Id, DisplayName, UserPrincipalName, AccountEnabled, OnPremisesSyncEnabled, AssignedLicenses, LicenseAssignmentStates `
        -ErrorAction Stop
}
catch {
    throw "Could not resolve '$($EmployeeAD.UserPrincipalName)' in Microsoft Graph: $($_.Exception.Message)"
}

if ($CreateSharedMailboxRequested -and $CloudOnlyMailbox -and $GraphUser.OnPremisesSyncEnabled) {
    throw "-CloudOnlyMailbox cannot be used because '$($GraphUser.UserPrincipalName)' is directory-synchronized. Supply -OnPremExchangeServer so both mailbox objects are converted safely."
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

$DirectSkuIds = @($GraphUser.LicenseAssignmentStates |
    Where-Object { -not $_.AssignedByGroup } |
    ForEach-Object { $_.SkuId } |
    Sort-Object -Unique)
$InheritedLicenseStates = @($GraphUser.LicenseAssignmentStates | Where-Object { $_.AssignedByGroup })

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
$CloudSecurityGroups = @()

if (-not $SkipCloudGroupRemoval) {
    Write-Host '  Inspecting Exchange Online and Microsoft Entra group memberships...' -ForegroundColor Gray

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
        $GraphGroups = @(Get-MgUserMemberOfAsGroup `
            -UserId $GraphUser.Id `
            -All `
            -Property Id, DisplayName, Mail, GroupTypes, MailEnabled, SecurityEnabled, OnPremisesSyncEnabled, MembershipRule, IsAssignableToRole, AssignedLicenses `
            -ErrorAction Stop)
    }
    catch {
        throw "Could not enumerate the employee's Microsoft Entra group memberships: $($_.Exception.Message)"
    }

    $CloudSecurityGroups = @($GraphGroups | Where-Object {
        $_.SecurityEnabled -and -not $_.MailEnabled -and @($_.GroupTypes) -notcontains 'Unified'
    })

    foreach ($Group in @($GraphGroups | Where-Object { @($_.GroupTypes) -contains 'Unified' })) {
        try {
            $GroupIdentity = Get-NormalizedAddress $Group.Mail
            if (-not $GroupIdentity) {
                throw 'Microsoft Graph did not return the group mail address required by Exchange Online.'
            }

            $Owners = @(Get-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Owners -ResultSize Unlimited -ErrorAction Stop)
            $Members = @(Get-UnifiedGroupLinks -Identity $GroupIdentity -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
            $IsOwner = @($Owners | Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $EmployeeAddress }).Count -gt 0
            $ManagerIsMember = @($Members | Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $ManagerAddress }).Count -gt 0
            $ManagerIsOwner = @($Owners | Where-Object { (Get-NormalizedAddress $_.PrimarySmtpAddress) -eq $ManagerAddress }).Count -gt 0

            $Microsoft365GroupPlans.Add([pscustomobject]@{
                Group           = $Group
                Identity        = $GroupIdentity
                Owners          = $Owners
                IsOwner         = $IsOwner
                ManagerIsMember = $ManagerIsMember
                ManagerIsOwner  = $ManagerIsOwner
            })
        }
        catch {
            throw "Could not inspect Microsoft 365 group '$($Group.DisplayName)': $($_.Exception.Message)"
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
    Write-Host "Cloud    : $($CloudDistributionGroups.Count) distribution, $($Microsoft365GroupPlans.Count) Microsoft 365, $($CloudSecurityGroups.Count) security group membership(s) found" -ForegroundColor White
}
Write-Host "Licenses : $($DirectSkuIds.Count) direct; $($InheritedLicenseStates.Count) group-inherited assignment state(s)" -ForegroundColor White
if ($CreateSharedMailboxRequested) {
    Write-Host 'Mailbox action: convert to shared, grant the AD manager Full Access, then remove eligible licenses.' -ForegroundColor Yellow
}
else {
    Write-Warning 'Shared mailbox is NO. The script will not convert or delegate the mailbox. Removing its Exchange license can deprovision the mailbox according to your retention policy.'
}

if ($LicenseBlockers.Count -gt 0) {
    Write-Warning "License safety check: $($LicenseBlockers -join '; '). License removal will be blocked unless -OverrideSharedMailboxLicenseSafety is explicitly supplied."
}

if (-not $Force -and -not $WhatIfPreference) {
    $Answer = Read-Host "Continue? Type OFFBOARD $User to make these changes"
    if ($Answer -cne "OFFBOARD $User") {
        Write-Host 'Cancelled. No changes were made.' -ForegroundColor Yellow
        return
    }
}

Write-Section 'Step 1 - disable sign-in and revoke sessions'

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

if ($GraphUser.AccountEnabled -eq $false) {
    Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Block cloud sign-in' -Status 'Skipped' -Message 'Cloud sign-in is already blocked.'
}
elseif ($PSCmdlet.ShouldProcess($EmployeeAddress, 'Block Microsoft Entra ID sign-in')) {
    try {
        Update-MgUser -UserId $GraphUser.Id -AccountEnabled:$false -ErrorAction Stop
        Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Block cloud sign-in' -Status 'Completed'
    }
    catch {
        Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Block cloud sign-in' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
    Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Block cloud sign-in' -Status 'Preview'
}

if ($PSCmdlet.ShouldProcess($EmployeeAddress, 'Revoke Microsoft 365 sign-in sessions')) {
    try {
        Revoke-MgUserSignInSession -UserId $GraphUser.Id -ErrorAction Stop | Out-Null
        Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Revoke sign-in sessions' -Status 'Completed'
    }
    catch {
        Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Revoke sign-in sessions' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
    Add-Result -System 'Microsoft Entra ID' -Item $EmployeeAddress -Action 'Revoke sign-in sessions' -Status 'Preview'
}

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

Write-Section 'Step 3 - remove Microsoft 365 group memberships'

if ($SkipCloudGroupRemoval) {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Remove cloud group memberships' -Status 'Skipped' -Message 'SkipCloudGroupRemoval was specified.'
}
else {
    foreach ($Failure in $DistributionGroupInspectionFailures) {
        Add-Result -System 'Exchange Online DL' -Item $Failure.Group -Action 'Inspect membership' -Status 'Attention' -Message $Failure.Message
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
                Remove-UnifiedGroupLinks -Identity $Plan.Identity -LinkType Members -Links $EmployeeAddress -Confirm:$false -ErrorAction Stop

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

    foreach ($Group in $CloudSecurityGroups) {
        if ($Group.OnPremisesSyncEnabled) {
            Add-Result -System 'Microsoft Entra security group' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Skipped' -Message 'Directory-synchronized; the on-premises AD removal must synchronize to Microsoft Entra ID.'
            continue
        }
        if ($Group.MembershipRule) {
            Add-Result -System 'Microsoft Entra security group' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Attention' -Message 'Dynamic group; membership cannot be removed manually. Review its rule, especially if it assigns licenses.'
            continue
        }
        if ($Group.IsAssignableToRole) {
            Add-Result -System 'Microsoft Entra security group' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Attention' -Message 'Role-assignable group; remove with the appropriate privileged-role process.'
            continue
        }

        if ($PSCmdlet.ShouldProcess($Group.DisplayName, "Remove $EmployeeAddress from the Microsoft Entra security group")) {
            try {
                Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $GraphUser.Id -Confirm:$false -ErrorAction Stop
                Add-Result -System 'Microsoft Entra security group' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Completed'
            }
            catch {
                Add-Result -System 'Microsoft Entra security group' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Failed' -Message $_.Exception.Message
            }
        }
        else {
            Add-Result -System 'Microsoft Entra security group' -Item $Group.DisplayName -Action 'Remove membership' -Status 'Preview'
        }
    }
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
        $CanRemoveLicenses = $false
        Add-Result -System 'On-premises Exchange' -Item $User -Action 'Convert remote mailbox to shared' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
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
        $CanRemoveLicenses = $false
        Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'Convert to shared mailbox' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
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
        Add-Result -System 'Exchange Online' -Item $ManagerAddress -Action 'Grant mailbox Full Access' -Status 'Preview'
    }
}
catch {
    $CanRemoveLicenses = $false
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

if ($LicenseBlockers.Count -gt 0 -and -not $OverrideSharedMailboxLicenseSafety) {
    $CanRemoveLicenses = $false
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'License safety check' -Status 'Attention' -Message ($LicenseBlockers -join '; ')
}
elseif ($LicenseBlockers.Count -gt 0) {
    Add-Result -System 'Exchange Online' -Item $EmployeeAddress -Action 'License safety check' -Status 'Attention' -Message ("Override supplied for: " + ($LicenseBlockers -join '; '))
}

Write-Section 'Step 5 - remove Microsoft 365 licenses'

if ($SkipLicenseRemoval) {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Remove licenses' -Status 'Skipped' -Message 'SkipLicenseRemoval was specified.'
}
elseif (-not $CanRemoveLicenses) {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Remove licenses' -Status 'Attention' -Message 'Safety gate blocked removal; resolve mailbox conversion, delegation, or licensing warnings first.'
}
elseif ($DirectSkuIds.Count -eq 0) {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Remove direct licenses' -Status 'Skipped' -Message 'No directly assigned licenses found.'
}
elseif ($PSCmdlet.ShouldProcess($EmployeeAddress, "Remove $($DirectSkuIds.Count) directly assigned Microsoft 365 license(s)")) {
    try {
        Set-MgUserLicense -UserId $GraphUser.Id -AddLicenses @() -RemoveLicenses $DirectSkuIds -ErrorAction Stop | Out-Null

        $VerifiedGraphUser = Get-MgUser -UserId $GraphUser.Id -Property LicenseAssignmentStates -ErrorAction Stop
        $RemainingDirectSkuIds = @($VerifiedGraphUser.LicenseAssignmentStates |
            Where-Object { -not $_.AssignedByGroup } |
            ForEach-Object { $_.SkuId } |
            Sort-Object -Unique)
        if ($RemainingDirectSkuIds.Count -gt 0) {
            throw "$($RemainingDirectSkuIds.Count) directly assigned license(s) remain after the removal request."
        }

        Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Remove direct licenses' -Status 'Completed' -Message "$($DirectSkuIds.Count) license(s) removed and verified."
    }
    catch {
        Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Remove direct licenses' -Status 'Failed' -Message $_.Exception.Message
    }
}
else {
    Add-Result -System 'Microsoft 365' -Item $EmployeeAddress -Action 'Remove direct licenses' -Status 'Preview' -Message "$($DirectSkuIds.Count) license(s)."
}

if ($InheritedLicenseStates.Count -gt 0) {
    Add-Result `
        -System 'Microsoft 365' `
        -Item $EmployeeAddress `
        -Action 'Verify group-inherited licenses' `
        -Status 'Attention' `
        -Message "$($InheritedLicenseStates.Count) group-inherited assignment state(s) existed at preflight. They remain until group removal and directory synchronization/licensing processing finish."
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
