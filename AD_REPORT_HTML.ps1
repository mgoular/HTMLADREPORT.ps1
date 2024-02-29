<#
.SYNOPSIS
    Generate graphed report for all Active Directory objects.
.DESCRIPTION
    Generate graphed report for all Active Directory objects.
.PARAMETER CompanyLogo
    Enter URL or UNC path to your desired Company Logo for generated report.
    Example: "\\Server01\Admin\Files\CompanyLogo.png"
.PARAMETER RightLogo
    Enter URL or UNC path to your desired right-side logo for generated report.
    Example: "https://www.yoursite/yourimage.png"
.PARAMETER ReportTitle
    Enter desired title for generated report.
    Default: "Active Directory Report"
.PARAMETER Days
    Users that have not logged in [X] amount of days or more.
    Default: 30
.PARAMETER UserCreatedDays
    Users that have been created within [X] amount of days.
    Default: 7
.PARAMETER DaysUntilPWExpireINT
    Users password expires within [X] amount of days
    Default: 7
.PARAMETER ADModNumber
    Active Directory Objects that have been modified within [X] amount of days.
    Default: 3
.NOTES
    Version: 1.0.0
    Author: Michael Goulart
#>

param (
    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path to Company Logo")]
    [String]$CompanyLogo = "",

    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter URL or UNC path for Side Logo")]
    [String]$RightLogo = "https://www.psmpartners.com/wp-content/uploads/2017/10/porcaro-stolarek-mete.png",

    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired title for report")]
    [String]$ReportTitle = "Active Directory Report",

    [Parameter(ValueFromPipeline = $true, HelpMessage = "Enter desired directory path to save; Default: C:\Automation\")]
    [String]$ReportSavePath = "C:\Automation\",

    [Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have not logged on in more than [X] days. amount of days; Default: 30")]
    $Days = 30,

    [Parameter(ValueFromPipeline = $true, HelpMessage = "Users that have been created within [X] amount of days; Default: 7")]
    $UserCreatedDays = 7,

    [Parameter(ValueFromPipeline = $true, HelpMessage = "Users password expires within [X] amount of days; Default: 7")]
    $DaysUntilPWExpireINT = 7,

    [Parameter(ValueFromPipeline = $true, HelpMessage = "AD Objects that have been modified within [X] amount of days; Default: 3")]
    $ADModNumber = 3
)

# Function to convert last logon time
function Convert-LastLogon {
    param ($fileTime)
    if ($fileTime -le 0 -or $fileTime -eq $null) {
        return "Never"
    } else {
        return [DateTime]::FromFileTime($fileTime)
    }
}

Write-Host "Gathering Report Customization..." -ForegroundColor White
Write-Host "__________________________________" -ForegroundColor White
Write-Host "Company Logo (left): " -NoNewline -ForegroundColor Yellow; Write-Host $CompanyLogo -ForegroundColor White
Write-Host "Company Logo (right): " -NoNewline -ForegroundColor Yellow; Write-Host $RightLogo -ForegroundColor White
Write-Host "Report Title: " -NoNewline -ForegroundColor Yellow; Write-Host $ReportTitle -ForegroundColor White
Write-Host "Report Save Path: " -NoNewline -ForegroundColor Yellow; Write-Host $ReportSavePath -ForegroundColor White
Write-Host "Amount of Days from Last User Logon Report: " -NoNewline -ForegroundColor Yellow; Write-Host $Days -ForegroundColor White
Write-Host "Amount of Days for New User Creation Report: " -NoNewline -ForegroundColor Yellow; Write-Host $UserCreatedDays -ForegroundColor White
Write-Host "Amount of Days for User Password Expiration Report: " -NoNewline -ForegroundColor Yellow; Write-Host $DaysUntilPWExpireINT -ForegroundColor White
Write-Host "Amount of Days for Newly Modified AD Objects Report: " -NoNewline -ForegroundColor Yellow; Write-Host $ADModNumber -ForegroundColor White
Write-Host "__________________________________" -ForegroundColor White

# Check for and install ReportHTML module if not present
$ReportHTMLModule = Get-Module -ListAvailable -Name "ReportHTML"
if ($null -eq $ReportHTMLModule) {
    Write-Host "ReportHTML Module is not present, attempting to install it"
    Install-Module -Name ReportHTML -Force
    Import-Module ReportHTML -ErrorAction SilentlyContinue
}

# Array of default Security Groups
$DefaultSecurityGroups = @(
    "Access Control Assistance Operators",
    "Account Operators",
    # Add other default security groups here...
)

# Get all users right away
$AllUsers = Get-ADUser -Filter * -Properties *

# Initialize tables
$ADObjectTable = @()
$CompanyInfoTable = @()

# Working on Dashboard Report
Write-Host "Working on Dashboard Report..." -ForegroundColor Green

# Get recently modified AD objects
$modifiedDate = (Get-Date).AddDays(- $ADModNumber)
$ADObjs = Get-ADObject -Filter { whenchanged -gt $modifiedDate -and ObjectClass -notin @("domainDNS", "rIDManager", "rIDSet") } -Properties *

foreach ($ADObj in $ADObjs) {
    $name = if ($ADObj.ObjectClass -eq "GroupPolicyContainer") { $ADObj.DisplayName } else { $ADObj.Name }
    $obj = [PSCustomObject]@{
        'Name' = $name
        'Object Type' = $ADObj.ObjectClass
        'When Changed' = $ADObj.WhenChanged
    }
    $ADObjectTable += $obj
}

# Check if AD Recycle Bin is enabled
$ADRecycleBinStatus = (Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"').EnabledScopes
$ADRecycleBin = if ($ADRecycleBinStatus.Count -lt 1) { "Disabled" } else { "Enabled" }

# Get domain information
$ADInfo = Get-ADDomain
$ForestObj = Get-ADForest
$DomainControllerObj = Get-ADDomain
$Domain = $ADInfo.Forest
$InfrastructureMaster = $DomainControllerObj.InfrastructureMaster
$RIDMaster = $DomainControllerObj.RIDMaster
$PDCEmulator = $DomainControllerObj.PDCEmulator
$DomainNamingMaster = $ForestObj.DomainNamingMaster
$SchemaMaster = $ForestObj.SchemaMaster

$companyInfo = [PSCustomObject]@{
    'Domain' = $Domain
    'AD Recycle Bin' = $ADRecycleBin
    'Infrastructure Master' = $InfrastructureMaster
    'RID Master' = $RIDMaster
    'PDC Emulator' = $PDCEmulator
    'Domain Naming Master' = $DomainNamingMaster
    'Schema Master' = $SchemaMaster
}
$CompanyInfoTable += $companyInfo

# Display information if no data is available
if ($ADObjectTable.Count -eq 0) {
    $infoObj = [PSCustomObject]@{
        Information = 'Information: No AD Objects have been modified recently'
    }
    $ADObjectTable += $infoObj
}

if ($CompanyInfoTable.Count -eq 0) {
    $infoObj = [PSCustomObject]@{
        Information = 'Information: Could not get items for table'
    }
    $CompanyInfoTable += $infoObj
}

# Get newly created users
$When = (Get-Date).AddDays(-$UserCreatedDays).Date
$NewUsers = $AllUsers | Where-Object { $_.whenCreated -ge $When }

foreach ($NewUser in $NewUsers) {
    $obj = [PSCustomObject]@{
        'Name' = $NewUser.Name
        'Enabled' = $NewUser.Enabled
        'Creation Date' = $NewUser.whenCreated
    }
    $NewCreatedUsersTable.Add($obj)
}

if ($NewCreatedUsersTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'Information: No new users have been recently created'
    }
    $NewCreatedUsersTable.Add($obj)
}

# Get Domain Admins
$DomainAdminMembers = Get-ADGroupMember "Domain Admins"

foreach ($DomainAdminMember in $DomainAdminMembers) {
    $Name = $DomainAdminMember.Name
    $Type = $DomainAdminMember.ObjectClass
    $Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled

    $obj = [PSCustomObject]@{
        'Name' = $Name
        'Enabled' = $Enabled
        'Type' = $Type
    }
    $DomainAdminTable.Add($obj)
}

if ($DomainAdminTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'Information: No Domain Admin Members were found'
    }
    $DomainAdminTable.Add($obj)
}

# Get Enterprise Admins
$EnterpriseAdminsMembers = Get-ADGroupMember "Enterprise Admins" -Server $SchemaMaster

foreach ($EnterpriseAdminsMember in $EnterpriseAdminsMembers) {
    $Name = $EnterpriseAdminsMember.Name
    $Type = $EnterpriseAdminsMember.ObjectClass
    $Enabled = ($AllUsers | Where-Object { $_.Name -eq $Name }).Enabled

    $obj = [PSCustomObject]@{
        'Name' = $Name
        'Enabled' = $Enabled
        'Type' = $Type
    }
    $EnterpriseAdminTable.Add($obj)
}

if ($EnterpriseAdminTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'Information: Enterprise Admin members were found'
    }
    $EnterpriseAdminTable.Add($obj)
}

# Get computers in the default OU
$DefaultComputersOU = (Get-ADDomain).computerscontainer
$DefaultComputers = Get-ADComputer -Filter * -Properties * -SearchBase "$DefaultComputersOU"

foreach ($DefaultComputer in $DefaultComputers) {
    $obj = [PSCustomObject]@{
        'Name' = $DefaultComputer.Name
        'Enabled' = $DefaultComputer.Enabled
        'Operating System' = $DefaultComputer.OperatingSystem
        'Modified Date' = $DefaultComputer.Modified
        'Password Last Set' = $DefaultComputer.PasswordLastSet
        'Protect from Deletion' = $DefaultComputer.ProtectedFromAccidentalDeletion
    }
    $DefaultComputersinDefaultOUTable.Add($obj)
}

if ($DefaultComputersinDefaultOUTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'Information: No computers were found in the Default OU'
    }
    $DefaultComputersinDefaultOUTable.Add($obj)
}

# Get users in the default OU
$DefaultUsersOU = (Get-ADDomain).UsersContainer
$DefaultUsers = $AllUsers | Where-Object { $_.DistinguishedName -like "*$DefaultUsersOU" } | Select-Object Name, UserPrincipalName, Enabled, ProtectedFromAccidentalDeletion, EmailAddress, @{Name='lastlogon'; Expression={LastLogonConvert $_.lastlogon}}, DistinguishedName

foreach ($DefaultUser in $DefaultUsers) {
    $obj = [PSCustomObject]@{
        'Name' = $DefaultUser.Name
        'UserPrincipalName' = $DefaultUser.UserPrincipalName
        'Enabled' = $DefaultUser.Enabled
        'Protected from Deletion' = $DefaultUser.ProtectedFromAccidentalDeletion
        'Last Logon' = $DefaultUser.LastLogon
        'Email Address' = $DefaultUser.EmailAddress
    }
    $DefaultUsersinDefaultOUTable.Add($obj)
}

if ($DefaultUsersinDefaultOUTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'Information: No Users were found in the default OU'
    }
    $DefaultUsersinDefaultOUTable.Add($obj)
}

# Expiring Accounts
$LooseUsers = Search-ADAccount -AccountExpiring -UsersOnly

foreach ($LooseUser in $LooseUsers) {
    $NameLoose = $LooseUser.Name
    $UPNLoose = $LooseUser.UserPrincipalName
    $ExpirationDate = $LooseUser.AccountExpirationDate
    $Enabled = $LooseUser.Enabled

    $obj = [PSCustomObject]@{
        'Name' = $NameLoose
        'UserPrincipalName' = $UPNLoose
        'Expiration Date' = $ExpirationDate
        'Enabled' = $Enabled
    }
    $ExpiringAccountsTable.Add($obj)
}

if ($ExpiringAccountsTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'Information: No Users were found to expire soon'
    }
    $ExpiringAccountsTable.Add($obj)
}
# Security Logs
$SecurityLogs = Get-WinEvent -LogName "Security" -MaxEvents 7 | Where-Object { $_.Message -like "*An account*" }

$SecurityEventTable = @()

foreach ($SecurityLog in $SecurityLogs) {
    $TimeGenerated = $SecurityLog.TimeCreated
    $EntryType = $SecurityLog.LevelDisplayName
    $Recipient = $SecurityLog.Message

    $obj = [PSCustomObject]@{
        'Time'    = $TimeGenerated
        'Type'    = $EntryType
        'Message' = $Recipient
    }

    $SecurityEventTable += $obj
}

if ($SecurityEventTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'No logon security events were found'
    }
    $SecurityEventTable += $obj
}

# Tenant Domain
$DomainTable = @()

$Domains = Get-ADForest | Select-Object -ExpandProperty UPNSuffixes

foreach ($Domain in $Domains) {
    $obj = [PSCustomObject]@{
        'UPN Suffixes' = $Domain
        Valid          = $true
    }

    $DomainTable += $obj
}

if ($DomainTable.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'No UPN Suffixes were found'
    }
    $DomainTable += $obj
}

Write-Host "Security Logs and Tenant Domain processed." -ForegroundColor Green

# Groups
$Groups = Get-ADGroup -Filter * -Properties * | Sort-Object Name

$GroupData = @()

foreach ($Group in $Groups) {
    $Type = switch ($Group.GroupCategory) {
        'Distribution' { 'Distribution Group' }
        'Security' {
            if ($Group.mail -ne $null) { 'Mail-Enabled Security Group' }
            else { 'Security Group' }
        }
    }

    $Members = @()
    if ($Group.Name -ne 'Domain Users') {
        $Members = (Get-ADGroupMember -Identity $Group | Sort-Object DisplayName).Name -join ', '
    }
    else {
        $Members = "Skipped Domain Users Membership"
    }

    $Manager = try {
        $Owner = Get-ADUser -Identity $Group.ManagedBy
        $Owner.Name
    } catch {
        "Cannot resolve manager for $($Group.Name)"
    }

    $obj = [PSCustomObject]@{
        'Name'                  = $Group.Name
        'Type'                  = $Type
        'Members'               = $Members
        'Managed By'            = $Manager
        'Email Address'         = $Group.Mail
        'Protected from Deletion' = $Group.ProtectedFromAccidentalDeletion
    }

    $GroupData += $obj
}

if ($GroupData.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'No groups were found'
    }
    $GroupData += $obj
}

Write-Host "Groups processed." -ForegroundColor Green

# Organizational Units
$OUs = Get-ADOrganizationalUnit -Filter * -Properties * | Sort-Object Name

$OuData = @()

foreach ($OU in $OUs) {
    $LinkedGPOs = $OU.linkedGroupPolicyObjects | ForEach-Object { (Get-GPO -Guid $_.Guid).DisplayName }
    $LinkedGPOs = if ($LinkedGPOs) { $LinkedGPOs -join ', ' } else { 'None' }

    $obj = [PSCustomObject]@{
        'Name'                  = $OU.Name
        'Linked GPOs'           = $LinkedGPOs
        'Modified Date'         = $OU.WhenChanged
        'Protected from Deletion' = $OU.ProtectedFromAccidentalDeletion
    }

    $OuData += $obj
}

if ($OuData.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'No OUs were found'
    }
    $OuData += $obj
}

Write-Host "Organizational Units processed." -ForegroundColor Green

# Users
$Users = Get-ADUser -Filter * -Properties * | Sort-Object Name

$UserData = @()

foreach ($User in $Users) {
    $DaysUntilPasswordExpires = if ($User.PasswordNeverExpires) { 'Never Expires' }
                                else { [math]::Ceiling(($User.PasswordLastSet.AddDays((Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.TotalDays) - (Get-Date)).TotalDays) }

    $LastLogon = if ($User.LastLogon -ne $null) { $User.LastLogon } else { "Never logged in" }

    $obj = [PSCustomObject]@{
        'Name'                  = $User.Name
        'UserPrincipalName'     = $User.UserPrincipalName
        'Enabled'               = $User.Enabled
        'Protected from Deletion' = $User.ProtectedFromAccidentalDeletion
        'Last Logon'            = $LastLogon
        'Email Address'         = $User.EmailAddress
        'Account Expiration'    = $User.AccountExpirationDate
        'Change Password Next Logon' = $User.PasswordExpired
        'Password Last Set'     = $User.PasswordLastSet
        'Password Never Expires' = $User.PasswordNeverExpires
        'Days Until Password Expires' = $DaysUntilPasswordExpires
    }

    $UserData += $obj
}

if ($UserData.Count -eq 0) {
    $obj = [PSCustomObject]@{
        Information = 'No users were found'
    }
    $UserData += $obj
}

Write-Host "Users processed." -ForegroundColor Green

# Output
$SecurityEventTable
$DomainTable
$GroupData
$OuData
$UserData
<#
    Group Policy Report
#>
Write-Host "Working on Group Policy Report..." -ForegroundColor Green

$GPOTable = @()

foreach ($GPO in $GPOs) {
    $obj = [PSCustomObject]@{
        'Name' = $GPO.DisplayName
        'Status' = $GPO.GpoStatus
        'Modified Date' = $GPO.ModificationTime
        'User Version' = $GPO.UserVersion
        'Computer Version' = $GPO.ComputerVersion
    }
    $GPOTable += $obj
}

if ($GPOTable.Count -eq 0) {
    $Obj = [PSCustomObject]@{
        Information = 'No Group Policy Objects were found'
    }
    $GPOTable += $obj
}

Write-Host "Group Policy Report done!" -ForegroundColor White

<#
    Computers Report
#>
Write-Host "Working on Computers Report..." -ForegroundColor Green

$Computers = Get-ADComputer -Filter *
$ComputersTable = @()
$OsVersions = @{}
$WindowsRegex = "(Windows (Server )?(\d+|XP)?( R2)?).*"

foreach ($Computer in $Computers) {
    $obj = [PSCustomObject]@{
        'Name' = $Computer.Name
        'Enabled' = $Computer.Enabled
        'Operating System' = $Computer.OperatingSystem
        'Modified Date' = $Computer.Modified
        'Password Last Set' = $Computer.PasswordLastSet
        'Protect from Deletion' = $Computer.ProtectedFromAccidentalDeletion
    }
    $ComputersTable += $obj

    if ($Computer.OperatingSystem -match $WindowsRegex) {
        $OsVersions[$matches[1]]++
    }
}

$ComputerStats = @{
    'Protected' = ($Computers | Where-Object { $_.ProtectedFromAccidentalDeletion }).Count
    'Not Protected' = ($Computers | Where-Object { -not $_.ProtectedFromAccidentalDeletion }).Count
    'Enabled' = ($Computers | Where-Object { $_.Enabled }).Count
    'Disabled' = ($Computers | Where-Object { -not $_.Enabled }).Count
}

Write-Host "Computers Report done!" -ForegroundColor White

<#
    Generating Reports
#>
$ReportTitle = "Active Directory Report"
$ReportDate = Get-Date -Format MM-dd-yyyy
$ReportName = "$ReportDate - AD Report"

$FinalReport = @()

# Add Group Policy Report
$FinalReport += @{
    'TabName' = 'Group Policy'
    'Content' = $GPOTable
}

# Add Computers Report
$FinalReport += @{
    'TabName' = 'Computers'
    'Content' = $ComputersTable
    'Stats' = $ComputerStats
}

# Add other reports...

# Save and display the HTML report
Save-HTMLReport -ReportContent $FinalReport -ShowReport -ReportName $ReportName -ReportPath $ReportSavePath



