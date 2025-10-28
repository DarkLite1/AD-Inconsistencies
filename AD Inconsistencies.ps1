#Requires -Version 7
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, ImportExcel
#Requires -Modules ActiveDirectory, Toolbox.ActiveDirectory

<#
    .SYNOPSIS
        Script to check the AD for inconsistencies.

    .DESCRIPTION
        Script to check the active directory for inconsistencies and violations against the BNL Naming convention.

    .PARAMETER ImportFile
        A .json file containing the script arguments.

        When the key 'Tickets' is used, a ticket will be created for that
        specific topic.

        Ex.
        {
            "MailTo": ["notification@contoso.net"],
            "InactiveDays": 40,
            "Prefix": {
                "QuotaGroup": "Quota limit"
            },
            "RolGroup": {
                "Prefix": "ROL-",
                "PlaceHolderAccount": "placeholder"
            },
            "AllowedEmployeeType": [ "Vendor", "Employee" ],
            "Group": [
                {
                    "Name": "Leavers",
                    "Type": "Exclude",
                    "ListMembers": true
                },
                {
                    "Name": "Deprovisioned users",
                    "Type": null,
                    "ListMembers": true
                }
            ],
            "OU": ["OU=BEL,OU=EU,DC=contoso,DC=net"],
            "Git": {
                "OU": "OU=GIT,DC=contoso,DC=net",
                "CountryCode": ["BE", "LU", "NL"]
            },
            "Tickets": {
                "User - VendorsNotExpiring": {
                    "ShortDescription": "Vendor account not expiring",
                    "Description": "Please set the expiration date"
                }
            }
        }

    .PARAMETER ScriptCreateTicketsFile
        File name or path of the script that creates the tickets and saves
        the details in the SQL database.

    .PARAMETER NoEmail
        When 'NoEmail' is used then the report will not be mailed to the
        addresses in MailTo.

        This can be useful in case the script needs to run every day to create
        the required tickets, but there's no need to send a report by mail every day.
#>

param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Switch]$NoEmail,
    [String]$ScriptCreateTicketsFile = 'Create tickets.ps1',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Inconsistencies\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

begin {
    function Get-PathItemHC {
        <#
        .SYNOPSIS
            Get the path item from a relative or absolute path

        .DESCRIPTION
            Perform Get-Item on a file located in a relative or absolute path

        .EXAMPLE
            Get-PathItemHC -Parent $PSScriptRoot -Leaf 'copy.ps1'

            Perform Get-Item on the script 'copy.ps1' in the current directory
        #>
        param (
            [Parameter(Mandatory)]
            [string]$Leaf,
            $Parent = $PSScriptRoot
        )
        if (Test-Path -LiteralPath (Join-Path -Path $Parent -ChildPath $Leaf) -PathType Leaf) {
            Get-Item -LiteralPath (Join-Path -Path $Parent -ChildPath $Leaf) -EA Stop
        }
        else {
            Get-Item -LiteralPath $Leaf -EA Stop
        }
    }

    try {
        $Error.Clear()
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $Now = Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import input file
        $File = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ([Int]$InactiveDays = $File.InactiveDays)) {
            throw "Input file '$ImportFile': No 'InactiveDays' path found."
        }

        if (-not ($MailTo = $File.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($OU = $File.OU)) {
            throw "Input file '$ImportFile': No 'OU' found."
        }

        $AllowedEmployeeType = $File.AllowedEmployeeType

        if ($GitOU = $File.Git.OU) {
            if (-not (Test-ADOuExistsHC $GitOU)) {
                throw "Input file '$ImportFile': GIT OU '$GitOU' does not exist."
            }

            if (-not ($GitCountryCode = $File.Git.CountryCode)) {
                throw "Input file '$ImportFile': GitCountryCode not found."
            }

            $GitSearchCountries = ($GitCountryCode | ForEach-Object { "(Country -EQ '$_')" }) -join ' -or '
        }
        #endregion

        #region Get tickets file
        try {
            $scriptCreateTicketsItem = Get-PathItemHC -Leaf $ScriptCreateTicketsFile
        }
        catch {
            throw "Create tickets script file '$ScriptCreateTicketsFile' not found"
        }
        #endregion

        $CompareDate = $Now.AddDays(-$InactiveDays)
        $YearAheadDate = $Now.AddYears(1)

        $Computers = $Groups = $Users = $allAdUsers = $OuCountry = @()
        $allAdGroups = @{ }
        $i = 0

        foreach ($O in $OU) {
            #region Match OU with country name
            try {
                Write-Verbose 'Match OU with country name'
                $ADou = Get-ADOrganizationalUnit $O -Properties Description, Country, CanonicalName
                $OuCountry += [PSCustomObject]@{
                    OU      = ($ADou.CanonicalName -replace '/', '\').ToUpper()
                    Country = $ADou.Description
                }
            }
            catch {
                throw "Input file '$ImportFile': OU '$O' does not exist: $_"
            }

            if (
                (-not $OuCountry[-1].Country) -or
                ($null -eq $OuCountry[-1].Country) -or
                ($null -eq $OuCountry[-1].OU)
            ) {
                throw "The AD Organizational Unit '$O' is incomplete. OU '$($OuCountry[-1].OU)' has country '$($OuCountry[-1].Country)'"
            }
            #endregion

            #region Get all AD computers
            $M = "Get all computers in OU '$O'"
            Write-Verbose $M
            Write-EventLog @EventVerboseParams -Message $M

            $Computers += Get-ADComputer -Filter * -SearchBase $O -Properties CanonicalName,
            Created, LastLogonDate, Location, ManagedBy, OperatingSystem,
            mS-DS-CreatorSID | Select-Object *,
            @{N = 'OU'; E = { ConvertTo-OuNameHC $_.CanonicalName } },
            @{N = 'ManagedByDisplayName'; E = { if ($_.ManagedBy) { Get-ADDisplayNameHC $_.ManagedBy } } },
            @{N = 'Creator'; E = { Get-ADDisplayNameFromSID -SID $_.'mS-DS-CreatorSID'.Value } }
            #endregion

            #region Get all AD groups
            $M = "Get all groups in OU '$O'"
            Write-Verbose $M
            Write-EventLog @EventVerboseParams -Message $M

            $groupsWithOrphans = @()
            $groupNonTraversable = @()

            $adGroupParams = @{
                Filter     = '*'
                SearchBase = $O
                Properties = @('CanonicalName', 'CN', 'Description',
                    'DisplayName', 'Mail', 'ManagedBy')
            }
            foreach ($group in (Get-ADGroup @adGroupParams)) {
                try {
                    $i++
                    Write-Verbose "$i Get group members '$($group.name)'"

                    $key = $group | Select-Object *,
                    @{N = 'ManagedByDisplayName'; E = { if ($_.ManagedBy) { Get-ADDisplayNameHC $_.ManagedBy } } }, @{N = 'OU'; E = { ConvertTo-OuNameHC $_.CanonicalName } }

                    try {
                        $groupMembers, $noDistinguishedName = @(
                            Get-ADGroupMember -Identity $group -Recursive -EA Stop).Where( {
                                $_.DistinguishedName
                            }, 'Split')
                    }
                    catch {
                        $groupNonTraversable += $key
                        $Error.RemoveAt(0)
                        continue
                    }

                    if ($noDistinguishedName) {
                        $groupsWithOrphans += $key
                    }

                    $allAdGroups[$key] = @($groupMembers | Select-Object *,
                        @{N = 'OU'; E = { ConvertTo-OuNameHC $_.DistinguishedName -EA Stop } }
                    )
                }
                catch {
                    Write-Error "Failed creating a group object for group '$($group.Name)': $_"
                    $Error.RemoveAt(1)
                }
            }
            #endregion

            #region Get all AD users
            $M = "Get all users in OU '$O'"
            Write-Verbose $M
            Write-EventLog @EventVerboseParams -Message $M

            $allAdUsers += @(
                Get-ADUser -Filter * -SearchBase $O -Properties whenCreated, displayName,
                sn, Title, Department, Company, manager, EmployeeID, extensionAttribute8,
                employeeType, CanonicalName, Description, co, physicalDeliveryOfficeName,
                OfficePhone, HomePhone, MobilePhone, ipPhone, Fax, pager, info,
                EmailAddress, scriptPath, homeDirectory, AccountExpirationDate,
                LastLogonDate, PasswordExpired, PasswordNeverExpires, LockedOut |
                Select-Object -Property *,
                @{N = 'LastName'; E = { $_.sn } },
                @{N = 'FirstName'; E = { $_.givenName } },
                @{N = 'ManagerDisplayName'; E = { if ($_.manager) { Get-ADDisplayNameHC $_.manager } } },
                @{N = 'HeidelbergCement Billing ID'; E = { $_.extensionAttribute8 } },
                @{N = 'OU'; E = { ConvertTo-OuNameHC $_.CanonicalName } },
                @{N = 'Office'; E = { $_.physicalDeliveryOfficeName } },
                @{N = 'Notes'; E = { $_.info -replace '`n', ' ' } },
                @{N = 'LogonScript'; E = { $_.scriptPath } },
                @{N = 'TSUserProfile'; E = { Get-ADTsProfileHC $_.DistinguishedName 'UserProfile' } },
                @{N = 'TSHomeDirectory'; E = { Get-ADTsProfileHC $_.DistinguishedName 'HomeDirectory' } },
                @{N = 'TSHomeDrive'; E = { Get-ADTsProfileHC $_.DistinguishedName 'HomeDrive' } }
            )
            #endregion
        }

        #region Get group members for excluded groups and listing groups, these can be outside the OU
        $M = "Get group members for excluded groups and listing groups in OU '$O'"
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M

        $tmpGroups = foreach (
            $E in
            @($File.Group).where(
                {
                    (
                        ($_.Type -eq 'Exclude') -or
                        ($_.ListMembers)
                    ) -and
                    ($groupNonTraversable.SamAccountName -notcontains $_.Name)
                }
            )
        ) {
            Write-Verbose "Get group members '$($E.Name)'"

            [PSCustomObject]@{
                Name        = $E.Name
                Type        = $E.Type
                ListMembers = $E.ListMembers
                Members     = Get-ADGroupMember $E.Name -Recursive | Get-ADUser -Properties whenCreated, displayName, sn,
                Title, Department, Company, manager, EmployeeID, extensionAttribute8, employeeType,
                CanonicalName, Description, co, physicalDeliveryOfficeName, OfficePhone, HomePhone,
                MobilePhone, ipPhone, Fax, pager, info, EmailAddress, scriptPath, homeDirectory,
                AccountExpirationDate, LastLogonDate, PasswordExpired, PasswordNeverExpires, LockedOut |
                Select-Object *,
                @{N = 'LastName'; E = { $_.sn } },
                @{N = 'FirstName'; E = { $_.givenName } },
                @{N = 'ManagerDisplayName'; E = { if ($_.manager) { Get-ADDisplayNameHC $_.manager } } },
                @{N = 'HeidelbergCement Billing ID'; E = { $_.extensionAttribute8 } },
                @{N = 'OU'; E = { ConvertTo-OuNameHC $_.CanonicalName } },
                @{N = 'Office'; E = { $_.physicalDeliveryOfficeName } },
                @{N = 'Notes'; E = { $_.info -replace '`n', ' ' } },
                @{N = 'LogonScript'; E = { $_.scriptPath } },
                @{N = 'TSUserProfile'; E = { Get-ADTsProfileHC $_.DistinguishedName 'UserProfile' } },
                @{N = 'TSHomeDirectory'; E = { Get-ADTsProfileHC $_.DistinguishedName 'HomeDirectory' } },
                @{N = 'TSHomeDrive'; E = { Get-ADTsProfileHC $_.DistinguishedName 'HomeDrive' } }
            }
        }

        $ExcludedGroups = $tmpGroups.where( { $_.Type -eq 'Exclude' })
        $GroupMembers = $tmpGroups.where( { $_.ListMembers })
        #endregion

        $Users = $allAdUsers.Where( {
                ($_.CanonicalName -notmatch '/Terminated users/|/Disabled/') -and
                ($ExcludedGroups.Members.SamAccountName -notcontains $_.SamAccountName)
            })

        $Groups = $allAdGroups.Keys
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; exit 1
    }
}

process {
    try {
        $AllObjects = @{ }

        #region Computers
        $M = 'Get all computers with issues'
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M

        $ComputerProperties = @(
            'Name', 'Description', 'Enabled', 'OperatingSystem',
            'LastLogonDate', 'Created', 'Creator', 'Location', 'ManagedByDisplayName', 'OU'
        )

        Write-Verbose 'Get computer Inactive'
        $AllObjects['Computer - Inactive'] = @{
            Description      = "'LastLogonDate' over $InactiveDays days<br>(Excluding OU 'Disabled')"
            WorksheetName    = 'Inactive'
            PropertyToExport = $ComputerProperties
            Type             = 'Computer'
            Data             = $Computers.where( {
                    ($_.OU -notmatch 'Terminated|Disabled') -and ($_.Enabled -eq $true) -and
                    (($_.LastLogonDate -eq $null) -or ($_.LastLogonDate -le $CompareDate))
                })
        }

        Write-Verbose 'Get computer EnabledInDisabledOU'
        $AllObjects['Computer - EnabledInDisabledOU'] = @{
            Description      = "'Enabled' in the OU 'Disabled'"
            WorksheetName    = 'EnabledInDisabledOU'
            PropertyToExport = $ComputerProperties
            Type             = 'Computer'
            Data             = $Computers.where( { ($_.OU -match 'Disabled') -and ($_.Enabled -eq $true) })
        }
        #endregion

        #region Groups
        $M = 'Get all groups with issues'
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M

        #region GroupsWithOrphans
        Write-Verbose 'Get group GroupsWithOrphans'
        $AllObjects['Group - GroupsWithOrphans'] = @{
            Description      = 'Groups with members that are no longer valid AD accounts because they are missing a DistinguishedName'
            WorksheetName    = 'GroupsWithOrphans'
            PropertyToExport = 'Name', 'DisplayName', 'Description', 'GroupCategory', 'GroupScope', 'OU'
            Type             = 'Group'
            Data             = $groupsWithOrphans
        }
        #endregion

        #region NonTraversableGroups
        Write-Verbose 'Get group NonTraversableGroups'
        $AllObjects['Group - NonTraversableGroups'] = @{
            Description      = "Groups where 'Get-ADGroupMember -Recursive' fails. Most likely these groups contain members from another domain."
            WorksheetName    = 'NonTraversableGroups'
            PropertyToExport = 'Name', 'DisplayName', 'Description', 'GroupCategory', 'GroupScope', 'OU'
            Type             = 'Group'
            Data             = $groupNonTraversable
        }
        #endregion

        #region CircularGroups
        Write-Verbose 'Get group CircularGroups'
        $AllObjects['Group - CircularGroups'] = @{
            Description      = 'Circular group membership'
            WorksheetName    = 'CircularGroups'
            PropertyToExport = 'Name', 'DisplayName', 'Description', 'GroupCategory', 'GroupScope', 'OU'
            Type             = 'Group'
            Data             = Get-ADCircularGroupsHC -OU $OU |
            Get-ADGroup -Properties Description, DisplayName, CanonicalName |
            Select-Object *, @{N = 'OU'; E = { ConvertTo-OuNameHC $_.CanonicalName } }
        }
        #endregion

        #region Distribution list without manager
        Write-Verbose 'Get group DisListNoManager'
        $AllObjects['Group - DisListNoManager'] = @{
            Description      = "GroupCategory 'Distribution' and 'ManagedBy' blank"
            WorksheetName    = 'DisListNoManager'
            PropertyToExport = 'Name', 'DisplayName', 'Description',
            'GroupCategory', 'ManagedBy', 'GroupScope', 'OU'
            Type             = 'Group'
            Data             = $Groups.where( { ($_.GroupCategory -eq 'Distribution') -and (-not $_.ManagedBy) })
        }
        #endregion

        #region Members that are not in our OU
        Write-Verbose 'Get group MembersNotInOU'
        [Regex]$ouFilter = $OU.ForEach( { "$_" }) -join '|'

        $i = 0
        $MembersNotInOU = foreach ($G in $allAdGroups.GetEnumerator()) {
            $foreignUsers = $G.Value.Where( {
                    ($_.ObjectClass -eq 'user') -and
                    ($_.distinguishedName -notmatch $ouFilter )
                })

            if ($foreignUsers) {
                $i++
                $foreignUsers |
                Select-Object @{Name = 'GroupName'; Expression = { $G.Key.SamAccountName } },
                @{Name = 'UserSamAccountName'; Expression = { $_.SamAccountName } },
                @{Name = 'UserName'; Expression = { $_.Name } },
                OU
            }
        }

        $AllObjects['Group - MembersNotInOU'] = @{
            Description      = 'Groups with members not in OU'
            WorksheetName    = 'MembersNotInOU'
            PropertyToExport = 'GroupName', 'UserName', 'UserSamAccountName', 'OU'
            Type             = 'Group'
            Data             = $MembersNotInOU
            Count            = $i
        }
        #endregion

        #region Group members
        foreach ($G in $GroupMembers) {
            Write-Verbose "Get group member list '$G'"
            $AllObjects["GroupMembers - $($G.Name)"] = @{
                Description      = 'List of group members'
                WorksheetName    = $G.Name
                PropertyToExport = 'SamAccountName', 'Name', 'Enabled', 'Description', 'LastLogonDate',
                'AccountExpirationDate', 'EmployeeType', 'homeDirectory', 'manager', 'OU'
                Type             = 'ListGroupMembers'
                Data             = $G.Members
            }
        }
        #endregion

        #region ROL Groups
        Write-Verbose 'Get ROL groups'
        if ($RolGroupPrefix = $File.RolGroup.Prefix) {
            $RolGroups = $Groups.where( { $_.SamAccountName -like "$RolGroupPrefix*" })


            [Array]$RolGroupsIncorrect = foreach ($G in $RolGroups) {
                Write-Verbose "ROL Group '$($G.SamAccountName)'"

                $Problem = @()

                #region Place holder account
                if ($RolPlaceholderAccount = $File.RolGroup.PlaceHolderAccount) {
                    $PlaceHolder = Get-ADGroupMember -Identity $G.SamAccountName -Recursive |
                    Where-Object SamAccountName -EQ $RolPlaceholderAccount

                    if (-not $PlaceHolder) {
                        $Problem += 'PlaceHolder'
                    }
                }
                #endregion

                #region Mail
                if ((-not $G.Mail) -or ($G.Mail -match '^\s+$')) {
                    $Problem += 'Mail'
                }
                #endregion

                #region GroupScope
                if (
                    ($G.GroupScope -ne 'Universal') -and
                    ($file.Tickets.'RolGroup - GroupScope'.Exclude -notcontains $G.SamAccountName)
                ) {
                    $Problem += 'GroupScope'
                }
                #endregion

                #region GroupCategory
                if ($G.GroupCategory -ne 'Security') {
                    $Problem += 'GroupCategory'
                }
                #endregion

                #region CN
                if ($G.CN -ne $G.Name) {
                    $Problem += 'CN'
                }
                #endregion

                #region DisplayName
                if ($G.DisplayName -ne (([Regex]'ROL').replace($G.Name, 'DIS', 1))) {
                    $Problem += 'DisplayName'
                }
                #endregion

                #region ManagedBy
                if (-not $G.ManagedBy) {
                    $Problem += 'ManagedBy'
                }
                #endregion

                if ($Problem) {
                    Write-Verbose "Problem '$Problem'"

                    $G | Select-Object *,
                    @{N = 'PlaceHolder'; E = { $PlaceHolder } },
                    @{N = 'Problem'; E = { $Problem } }
                }
            }

            $RolGroupType = 'RolGroup'

            Write-Verbose 'Get ROL group incorrect'
            $AllObjects['RolGroup - Incorrect'] = @{
                Description      = 'Incorrect ROL groups'
                WorksheetName    = 'ROL_Groups_incorrect'
                PropertyToExport = 'Name', 'CN', 'DisplayName', 'Description',
                'GroupCategory', 'GroupScope',
                'PlaceHolder', 'Mail', 'ManagedByDisplayName', 'OU',
                @{N = 'Problem'; E = { $_.Problem -join ',' } }
                Type             = $RolGroupType
                Data             = $RolGroupsIncorrect
            }

            if ($RolPlaceholderAccount) {
                Write-Verbose 'Get ROL group PlaceHolder'
                $AllObjects['RolGroup - PlaceHolder'] = @{
                    Description      = "Missing place holder account '$RolPlaceholderAccount' as member"
                    WorksheetName    = 'MissingPlaceHolder'
                    PropertyToExport = 'Name', 'Description', 'OU'
                    Type             = $RolGroupType
                    Data             = $RolGroupsIncorrect.where( { $_.Problem -contains 'PlaceHolder' })
                }
            }

            Write-Verbose 'Get ROL group Mail'
            $AllObjects['RolGroup - Mail'] = @{
                Description      = "'Mail' blank"
                WorksheetName    = 'MailBlank'
                PropertyToExport = 'Name', 'Description', 'OU', 'Mail'
                Type             = $RolGroupType
                Data             = $RolGroupsIncorrect.where( { $_.Problem -contains 'Mail' })
            }

            Write-Verbose 'Get ROL group GroupScope'
            $AllObjects['RolGroup - GroupScope'] = @{
                Description      = "'GroupScope' not 'Universal'"
                WorksheetName    = 'GroupScopeNotUniversal'
                PropertyToExport = 'Name', 'Description', 'OU', 'GroupScope'
                Type             = $RolGroupType
                Data             = $RolGroupsIncorrect.where(
                    {
                        $_.Problem -contains 'GroupScope'
                    }
                )
            }

            Write-Verbose 'Get ROL group GroupCategory'
            $AllObjects['RolGroup - GroupCategory'] = @{
                Description      = "'GroupCategory' not 'Security'"
                WorksheetName    = 'GroupCategoryNotSecurity'
                PropertyToExport = 'Name', 'Description', 'OU', 'GroupCategory'
                Type             = $RolGroupType
                Data             = $RolGroupsIncorrect.where( { $_.Problem -contains 'GroupCategory' })
            }

            Write-Verbose 'Get ROL group CN'
            $AllObjects['RolGroup - CN'] = @{
                Description      = "'CN' not equal to 'Name'"
                WorksheetName    = 'NameToSameAsCN'
                PropertyToExport = 'Name', 'CN', 'Description', 'OU'
                Type             = $RolGroupType
                Data             = $RolGroupsIncorrect.where( { $_.Problem -contains 'CN' })
            }

            Write-Verbose 'Get ROL group DisplayName'
            $AllObjects['RolGroup - DisplayName'] = @{
                Description      = "'DisplayName' not equal to 'Name'<br>
        (Where the word 'ROL' is not replaced with the word 'DIS')"
                WorksheetName    = 'DisplayNameNotSameAsName'
                PropertyToExport = 'Name', 'DisplayName', 'CN', 'Description', 'OU'
                Type             = $RolGroupType
                Data             = $RolGroupsIncorrect.where( { $_.Problem -contains 'DisplayName' })
            }

            Write-Verbose 'Get ROL group ManagedBy'
            $AllObjects['RolGroup - ManagedBy'] = @{
                Description      = "'ManagedBy' blank"
                WorksheetName    = 'ManagedByBlank'
                PropertyToExport = 'Name', 'OU', 'ManagedBy'
                Type             = $RolGroupType
                Data             = $RolGroupsIncorrect.where( { $_.Problem -contains 'ManagedBy' })
            }
        }
        #endregion
        #endregion

        #region Users
        $M = 'Get all users with issues'
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M

        Write-Verbose 'Get user CountryNotMatchingOU'
        $AllObjects['User - CountryNotMatchingOU'] = @{
            Description      = 'Country name not equal to the OU country name'
            WorksheetName    = 'CountryNotMatchingOU'
            PropertyToExport = 'SamAccountName', 'Name', 'Description', 'EmployeeType', 'co', 'OU'
            Type             = 'User'
            Data             = foreach ($User in $Users) {
                if (($OuCountry.where( { $User.OU -like "$($_.OU)\*" }).Country) -ne $User.co) {
                    $User
                }
            }
        }

        Write-Verbose 'Get user DescriptionWrong'
        $AllObjects['User - DescriptionWrong'] = @{
            Description      = "'Description' not compliant with naming convention<br>(Only for 'EmployeeType': Plant, Kiosk, Vendor, Service or Resource)"
            WorksheetName    = 'DescriptionWrong'
            PropertyToExport = 'SamAccountName', 'Name', 'Description', 'EmployeeType', 'co', 'OU'
            Type             = 'User'
            Data             = foreach ($User in $Users) {
                $DescriptionWrong = switch -Regex ($User.OU) {
                    '\\Resource Accounts$' {
                        -not (($User.Description -ceq 'Shared mailbox') -or ($User.Description -clike 'Shared mailbox - ?*') -or
                            ($User.Description -ceq 'Meeting room') -or ($User.Description -clike 'Meeting room - ?*'))
                        break
                    }
                    '\\Service Accounts$' {
                        -not (($User.Description -ceq 'Service') -or ($User.Description -clike 'Service - ?*'))
                        break
                    }
                    '\\Users$' {
                        switch -Regex ($User.EmployeeType) {
                            'Vendor|Kiosk|Plant' {
                                -not (($User.Description -ceq $_) -or ($User.Description -clike "$_ - ?*"))
                                break
                            }
                        }
                    }
                }

                if ($DescriptionWrong) {
                    $User
                }
            }
        }

        Write-Verbose 'Get user DisplayNameNotUnique'
        $AllObjects['User - DisplayNameNotUnique'] = @{
            Description      = "'DisplayName' not unique"
            WorksheetName    = 'DisplayNameNotUnique'
            PropertyToExport = 'SamAccountName', 'DisplayName', 'Description', 'EmployeeType', 'co', 'OU'
            Type             = 'User'
            Data             = $Users | Group-Object -Property { $_.DisplayName } |
            Where-Object { $_.Count -ge 2 } | Select-Object -ExpandProperty Group
        }

        Write-Verbose 'Get user DisplayNameWrong'
        $AllObjects['User - DisplayNameWrong'] = @{
            Description      = "'DisplayName' not compliant with naming convention<br>(Only for 'EmployeeType': Plant, Kiosk, Service or Resource)"
            WorksheetName    = 'DisplayNameWrong'
            PropertyToExport = 'SamAccountName', 'Name', 'Description', 'EmployeeType', 'co', 'OU'
            Type             = 'User'
            Data             = foreach ($User in $Users) {
                $DisplayNameWrong = switch -Regex ($User.OU) {
                    '\\Resource Accounts$' {
                        $(Compare-ADobjectNameHC $User.DisplayName -Type ResourceAccount).Valid
                        break
                    }
                    '\\Service Accounts$' {
                        $(Compare-ADobjectNameHC $User.DisplayName -Type ServiceAccount).Valid
                        break
                    }
                    '\\Users$' {
                        if (($User.EmployeeType -eq 'Kiosk') -or ($User.EmployeeType -eq 'Plant')) {
                            if (-not $User.DisplayName) {
                                $false; break
                            }
                            $(Compare-ADobjectNameHC $User.DisplayName -Type User).Valid
                        }
                        break
                    }
                }

                if ($DisplayNameWrong -eq $false) {
                    $User
                }
            }
        }

        Write-Verbose 'Get user DisplayNameVsName'
        $AllObjects['User - DisplayNameVsName'] = @{
            Description      = "'DisplayName' not equal to 'Name'"
            WorksheetName    = 'DisplayNameVsName'
            PropertyToExport = 'SamAccountName', 'Name', 'DisplayName', 'Description', 'EmployeeType', 'co', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.Name -ne $_.DisplayName })
        }

        Write-Verbose 'Get user EmployeeTypeBlank'
        $AllObjects['User - EmployeeTypeBlank'] = @{
            Description      = "'EmployeeType' blank"
            WorksheetName    = 'EmployeeTypeBlank'
            PropertyToExport = 'SamAccountName', 'Name', 'Description', 'EmployeeType', 'co', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.EmployeeType -eq $null })
        }

        if ($AllowedEmployeeType) {
            Write-Verbose 'Get user EmployeeTypeNotAllowed'
            $AllObjects['User - EmployeeTypeNotAllowed'] = @{
                Description      = "'EmployeeType' not allowed, only the following values are valid: $($AllowedEmployeeType -join ', ')"
                WorksheetName    = 'EmployeeTypeNotAllowed'
                PropertyToExport = 'SamAccountName', 'Name', 'Description', 'EmployeeType', 'co', 'OU'
                Type             = 'User'
                Data             = $Users.Where( { $AllowedEmployeeType -notcontains $_.EmployeeType })
            }
        }

        Write-Verbose 'Get user StartEndWithSpaces'
        $AllObjects['User - StartEndWithSpaces'] = @{
            Description      = "'FirstName' or 'LastName' starting or ending with a space"
            WorksheetName    = 'StartEndWithSpaces'
            PropertyToExport = 'SamAccountName', 'DisplayName', 'FirstName', 'LastName', 'OU'
            Type             = 'User'
            Data             = $Users.Where( {
                    ($_.FirstName -match '^\s|\s$') -or
                    ($_.LastName -match '^\s|\s$') })
        }

        Write-Verbose 'Get user EmployeeTypeVendor'
        $AllObjects['User - EmployeeTypeVendor'] = @{
            Description      = "All users with 'EmployeeType' Vendor"
            WorksheetName    = 'Vendors'
            PropertyToExport = 'SamAccountName', 'DisplayName', 'FirstName', 'LastName', 'Manager', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.EmployeeType -eq 'Vendor' })
        }

        Write-Verbose 'Get user HomeDirGrouphc'
        $AllObjects['User - HomeDirGrouphc'] = @{
            Description      = "'HomeDirectory' starting with '\\GROUPHC\' instead of '\\GROUPHC.NET\'"
            WorksheetName    = 'HomeDirGrouphc'
            PropertyToExport = 'SamAccountName', 'Name', 'HomeDirectory', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.HomeDirectory -match '^\\\\GROUPHC\\' })
        }

        Write-Verbose 'Get user HomeDirWrong'
        $AllObjects['User - HomeDirWrong'] = @{
            Description      = "'HomeDirectory' not starting with '\\GROUPHC.NET\BNL\HOME\'<br>(Excluding EmployeeType 'Service accounts' and 'Resource accounts')"
            WorksheetName    = 'HomeDirWrong'
            PropertyToExport = 'SamAccountName', 'Name', 'HomeDirectory', 'EmployeeType', 'OU'
            Type             = 'User'
            Data             = $Users.Where( {
                    ($_.HomeDirectory -notlike '\\Grouphc.net\bnl\HOME\*') -and
                    ($_.EmployeeType -ne 'Resource') -and
                    ($_.EmployeeType -ne 'Service') })
        }

        Write-Verbose 'Get user HomeDirNotNeeded'
        $AllObjects['User - HomeDirNotNeeded'] = @{
            Description      = "'HomeDirectory' set for EmployeeType 'Service accounts' or 'Resource accounts')"
            WorksheetName    = 'HomeDirNotNeeded'
            PropertyToExport = 'SamAccountName', 'Name', 'HomeDirectory', 'EmployeeType', 'OU'
            Type             = 'User'
            Data             = $Users.Where( {
                    ($_.HomeDirectory) -and
                    (($_.EmployeeType -eq 'Resource') -or ($_.EmployeeType -eq 'Service')) })
        }

        Write-Verbose 'Get user Inactive'
        $AllObjects['User - Inactive'] = @{
            Description      = "'LastLogonDate' over $InactiveDays days<br>(Excluding 'EmployeeType' Resource and the OU's 'Terminated users' and 'Disabled\Users')"
            WorksheetName    = 'Inactive'
            PropertyToExport = 'SamAccountName', 'DisplayName', 'LastLogonDate',
            'EmployeeType', 'ManagerDisplayName', 'whenCreated', 'OU'
            Type             = 'User'
            Data             = $Users.where( {
                    ($_.EmployeeType -ne 'Resource') -and
                    (($_.whenCreated -le $CompareDate) -and
                    (($_.LastLogonDate -eq $null) -or ($_.LastLogonDate -le $CompareDate))) })
        }

        Write-Verbose 'Get user LogonScriptBlank'
        $AllObjects['User - LogonScriptBlank'] = @{
            Description      = "'LogonScript' blank<br>excluding EmployeeType 'Resource' and 'Service'."
            WorksheetName    = 'LogonScriptBlank'
            PropertyToExport = 'SamAccountName', 'Name', 'LogonScript', 'EmployeeType', 'OU'
            Type             = 'User'
            Data             = $Users.where( {
                    ($_.LogonScript -eq $null) -and
                    ($_.EmployeeType -ne 'Resource') -and
                    ($_.EmployeeType -ne 'Service') })
        }

        Write-Verbose 'Get user LogonScriptNotExisting'
        $AllObjects['User - LogonScriptNotExisting'] = @{
            Description      = "'LogonScript' not found"
            WorksheetName    = 'LogonScriptNotExisting'
            PropertyToExport = 'SamAccountName', 'Name', 'LogonScript', 'ManagerDisplayName', 'OU'
            Type             = 'User'
            Data             = @($Users.Where( { $_.LogonScript }) | Group-Object LogonScript).Where( { $_.Name }).foreach( {
                    try {
                        if (-not (Test-Path -Path "\\$env:USERDNSDOMAIN\NETLOGON\$($_.Name)" -PathType Leaf)) {
                            $_.Group
                        }
                    }
                    catch {
                        Write-Warning "Access denied for logon script '$($_.Name)'"
                        $_.Group | Select-Object -ExcludeProperty LogonScript -Property *,
                        @{N = 'LogonScript'; E = { ('ACCESS DENIED:' + $_.LogonScript) } }
                    }
                })
        }

        Write-Verbose 'Get user LogonScriptResourceService'
        $AllObjects['User - LogonScriptResourceService'] = @{
            Description      = "'LogonScript' set for 'EmployeeType' Resource and Service"
            WorksheetName    = 'LogonScriptResourceService'
            PropertyToExport = 'SamAccountName', 'Name', 'LogonScript', 'EmployeeType', 'OU'
            Type             = 'User'
            Data             = $Users.Where( {
                    ($_.LogonScript -ne $null) -and
                    (($_.EmployeeType -eq 'Resource') -or ($_.EmployeeType -eq 'Service')) })
        }

        Write-Verbose 'Get user ManagerOfSelf'
        $AllObjects['User - ManagerOfSelf'] = @{
            Description      = 'Manager same as user account'
            WorksheetName    = 'ManagerOfSelf'
            PropertyToExport = 'SamAccountName', 'Name', 'DisplayName', 'ManagerDisplayName', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.DistinguishedName -eq $_.Manager })
        }

        Write-Verbose 'Get user NoManager'
        $AllObjects['User - NoManager'] = @{
            Description      = "'Manager' blank<br>excluding EmployeeType 'Resource', 'Employee' and 'Service'."
            WorksheetName    = 'NoManager'
            PropertyToExport = 'SamAccountName', 'Name', 'DisplayName', 'ManagerDisplayName', 'OU'
            Type             = 'User'
            Data             = $Users.Where( {
                    (-not $_.Manager) -and
                    ($_.EmployeeType -ne 'Employee') -and
                    ($_.EmployeeType -ne 'Resource') -and
                    ($_.EmployeeType -ne 'Service')
                }
            )
        }

        Write-Verbose 'Get user NoManager type Employee'
        $AllObjects['User - NoManagerEmployee'] = @{
            Description      = "'Manager' blank<br>Only EmployeeType 'Employee'."
            WorksheetName    = 'NoManagerEmployee'
            PropertyToExport = 'SamAccountName', 'Name', 'DisplayName', 'ManagerDisplayName', 'OU'
            Type             = 'User'
            Data             = $Users.Where( {
                    (-not $_.Manager) -and
                    ($_.EmployeeType -eq 'Employee')
                }
            )
        }

        Write-Verbose 'Get user SamNameWithNr'
        $AllObjects['User - SamNameWithNr'] = @{
            Description      = "'SamAccountName' containing a number"
            WorksheetName    = 'SamNameWithNr'
            PropertyToExport = 'SamAccountName', 'Name', 'DisplayName', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.SamAccountName -match '\d' })
        }

        Write-Verbose 'Get user TSHomeDirNotExist'
        $AllObjects['User - TSHomeDirNotExist'] = @{
            Description      = "'TSHomeDirectory' not found"
            WorksheetName    = 'TSHomeDirNotExist'
            PropertyToExport = 'SamAccountName', 'Name', 'TSHomeDirectory', 'TSHomeDirExist', 'OU'
            Type             = 'User'
            Data             = foreach ($User in $Users) {
                $TSHomeDirExist = $null

                if (($null -ne $User.TSHomeDirectory) -and ($User.TSHomeDirectory -ne '')) {
                    try {
                        $TSHomeDirExist = Test-Path -Path $User.TSHomeDirectory -PathType Container
                    }
                    catch {
                        $TSHomeDirExist = $false
                        $Error.RemoveAt(0)
                        Write-Warning "Access denied on the TS Home directory '$($User.TSHomeDirectory)' of user '$($User.DisplayName)'"
                    }
                }

                if ($TSHomeDirExist -eq $false) {
                    $User
                }
            }
        }

        Write-Verbose 'Get user TSHomeDirGrouphc'
        $AllObjects['User - TSHomeDirGrouphc'] = @{
            Description      = "'TSHomeDirectory' starting with '\\GROUPHC\' instead of '\\GROUPHC.NET\'"
            WorksheetName    = 'TSHomeDirGrouphc '
            PropertyToExport = 'SamAccountName', 'Name', 'TSHomeDirectory', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.TSHomeDirectory -match '^\\\\GROUPHC\\' })
        }

        Write-Verbose 'Get user TSHomeDirVsHomeDir'
        $AllObjects['User - TSHomeDirVsHomeDir'] = @{
            Description      = "'TSHomeDirectory' not equal to 'HomeDirectory'"
            WorksheetName    = 'TSHomeDirVsHomeDir'
            PropertyToExport = 'SamAccountName', 'Name', 'TSHomeDirectory', 'HomeDirectory', 'OU'
            Type             = 'User'
            Data             = $Users.where( {
                    ($_.TSHomeDirectory -ne $_.HomeDirectory) -and
                    (-not(($_.TSHomeDirectory -eq $null) -and ($_.HomeDirectory -eq $null)))
                })
        }

        Write-Verbose 'Get user TSProfileGrouphc'
        $AllObjects['User - TSProfileGrouphc'] = @{
            Description      = "'TSUserProfile' starting with '\\GROUPHC\' instead of '\\GROUPHC.NET\'"
            WorksheetName    = 'TSProfileGrouphc'
            PropertyToExport = 'SamAccountName', 'Name', 'TSUserProfile', 'OU'
            Type             = 'User'
            Data             = $Users.Where( { $_.TSUserProfile -match '^\\\\GROUPHC\\' })
        }

        Write-Verbose 'Get user TSProfileNotExisting'
        $AllObjects['User - TSProfileNotExisting'] = @{
            Description      = "'TSUserProfile' not found"
            WorksheetName    = 'TSProfileNotExisting'
            PropertyToExport = 'SamAccountName', 'Name', 'TSUserProfile', 'TSUserProfileExist', 'ManagerDisplayName', 'OU'
            Type             = 'User'
            Data             = foreach ($User in $Users) {
                $TSUserProfileExist = $TSUserProfileV2Exist = $null

                if (($null -ne $User.TSUserProfile) -and ($User.TSUserProfile -ne '')) {
                    #region Srv 2003
                    try {
                        $TSUserProfileExist = Test-Path -Path $User.TSUserProfile -PathType Container
                    }
                    catch {
                        $TSUserProfileExist = $false
                        $Error.RemoveAt(0)
                        Write-Warning "Access denied on the TS User Profile '$($User.TSUserProfile)' of user '$($User.DisplayName)'"
                    }
                    #endregion

                    #region Srv 2008
                    try {
                        $TSUserProfileV2Exist = Test-Path -Path "$($User.TSUserProfile).V2"-PathType Container
                    }
                    catch {
                        $TSUserProfileV2Exist = $false
                        $Error.RemoveAt(0)
                        Write-Warning "Access denied on the TS User Profile.V2 '$($User.TSUserProfile).V2 of user '$($User.DisplayName)'"
                    }
                    #endregion
                }

                if (($TSUserProfileExist -eq $false) -or
                    ($TSUserProfileV2Exist -eq $false)) {
                    $User
                }
            }
        }

        Write-Verbose 'Get user VendorsNotExpiring'
        $AllObjects['User - VendorsNotExpiring'] = @{
            Description      = "'EmployeeType' Vendor with 'AccountExpirationDate' set for over one year or none at all"
            WorksheetName    = 'VendorsNotExpiring'
            PropertyToExport = 'DisplayName', 'SamAccountName', 'AccountExpirationDate', 'EmployeeType', 'ManagerDisplayName', 'OU'
            Type             = 'User'
            Data             = $Users.Where( {
                    ($_.EmployeeType -eq 'Vendor') -and
                    (($_.AccountExpirationDate -eq $null) -or ($_.AccountExpirationDate -gt $YearAheadDate)) })
        }

        #region Quota management
        Write-Verbose 'Get quota management groups'
        if ($QuotaGroupNameBegin = $File.QuotaGroupNameBegin) {
            $UserProperties = @('Name', 'Description', 'Enabled', 'OperatingSystem',
                'LastLogonDate', 'Created', 'Creator', 'Location', 'ManagedByDisplayName', 'OU')

            $QuotaUsers = foreach ($G in (Get-ADGroup -Filter "Name -like '$QuotaGroupNameBegin*'")) {
                Write-Verbose "Quota group '$($G.SamAccountName)'"
                # avoid pipeline with AD CmdLets for Pester tests, known limitation in Pester 4.0.8
                $Members = Get-ADGroupMember $G.SamAccountName -Recursive
                foreach ($M in ($Members | Where-Object ObjectClass -EQ User)) {
                    Get-ADUser $M.DistinguishedName -Properties HomeDirectory |
                    Where-Object { $_.Enabled -and $_.HomeDirectory } |
                    Select-Object SamAccountName, @{N = 'GroupName'; E = { $G.SamAccountName } }
                }
            }

            Write-Verbose 'Get user QuotaMultiGroup'
            $AllObjects['User - QuotaMultiGroup'] = @{
                Description      = "Quota groups 'HomeDirectory', member of multiple groups"
                WorksheetName    = 'QuotaMultiGroup'
                PropertyToExport = $UserProperties
                Type             = 'User'
                Data             = $QuotaUsers | Group-Object SamAccountName |
                Where-Object Count -GT 1 | Select-Object -ExpandProperty Group
            }

            Write-Verbose 'Get user QuotaNotOnHomeDir'
            $AllObjects['User - QuotaNotOnHomeDir'] = @{
                Description      = "Quota groups 'HomeDirectory', not member"
                WorksheetName    = 'QuotaNotOnHomeDir'
                PropertyToExport = $UserProperties
                Type             = 'User'
                Data             = $Users.where( {
                        $_.Enabled -and $_.HomeDirectory -and
                        (($QuotaUsers.SamAccountName | Select-Object -Unique) -notcontains $_.SamAccountName)
                    })
            }
        }
        #endregion

        #endregion

        #region GIT users
        $M = 'Get all GIT users with issues'
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M

        if ($GitOU -and $GitSearchCountries) {
            Write-Verbose "Get users from GIT OU '$GitOU'"

            $GitUsers = @(Get-ADUser -SearchBase $GitOU -Filter $GitSearchCountries -Properties LastLogonDate,
                WhenCreated, Country, CanonicalName, manager |
                Select-Object *,
                @{N = 'OU'; E = { ConvertTo-OuNameHC $_.CanonicalName } },
                @{
                    N = 'ManagerDisplayName'
                    E = { if ($_.manager) { Get-ADDisplayNameHC $_.manager } }
                }
            )

            Write-Verbose 'Get GIT user Inactive'
            $AllObjects['GitUser - Inactive'] = @{
                Description      = "'LastLogonDate' over $InactiveDays days"
                WorksheetName    = 'Inactive'
                PropertyToExport = 'Name', 'SamAccountName', 'ManagerDisplayName', 'LastLogonDate',
                'WhenCreated', 'Country', 'Enabled', 'OU'
                Type             = 'GitUser'
                Data             = $GitUsers.where( {
                        (($_.LastLogonDate -le $CompareDate) -or ($_.LastLogonDate -eq $null)) -and
                        ($_.WhenCreated -le $CompareDate) -and
                        ($_.Enabled) })
            }

            Write-Verbose 'Get GIT user NoManger'
            $AllObjects['GitUser - NoManger'] = @{
                Description      = "'Manager' blank"
                WorksheetName    = 'NoManger'
                PropertyToExport = 'Name', 'SamAccountName', 'ManagerDisplayName', 'LastLogonDate',
                'WhenCreated', 'Country', 'Enabled', 'OU'
                Type             = 'GitUser'
                Data             = $GitUsers.where(
                    {
                        ($_.Enabled) -and (-not $_.manager)
                    }
                )
            }

            Write-Verbose 'Get GIT user NotOwnManager'
            $AllObjects['GitUser - NotOwnManager'] = @{
                Description      = "'Manager' is not own account"
                WorksheetName    = 'NotOwnManager'
                PropertyToExport = 'SamAccountName', 'Name', 'ManagerName', 'LastLogonDate', 'WhenCreated', 'Country', 'Enabled', 'OU'
                Type             = 'GitUser'
                Data             = foreach ($G in $GitUsers.where( { $_.manager })) {
                    $ManagerWrong = $true

                    $G | Add-Member -NotePropertyName ManagerName -NotePropertyValue $null

                    try {
                        $Manager = Get-ADUser $G.Manager

                        $G.ManagerName = $Manager.Name

                        if (($G.Name -replace '...$') -eq ($Manager.Name -replace '...$')) {
                            $ManagerWrong = $false
                        }
                    }
                    catch {
                        Write-Verbose 'Manager is not a user or is not found'
                    }

                    if ($ManagerWrong) {
                        $G
                    }
                }
            }
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; exit 1
    }
}

end {
    try {
        $MailParams = @{
            Attachments = @()
        }

        $ExcelParams = @{
            Path         = "$LogFile - Source data.xlsx"
            AutoSize     = $true
            FreezeTopRow = $true
        }

        if (-not $NoEmail) {
            #region Export source data to Excel
            $M = "Export source data to Excel file '$($ExcelParams.Path)'"
            Write-Verbose $M
            Write-EventLog @EventVerboseParams -Message $M

            $MailParams.Attachments += $ExcelParams.Path

            if ($Computers) {
                $M = "Export $(@($Computers).Count) computers to Excel file '$($ExcelParams.Path)'"
                Write-Verbose $M
                Write-EventLog @EventOutParams -Message $M
                $Computers | Export-Excel @ExcelParams -TableName Computers -WorksheetName Computers
            }
            if ($Groups) {
                $M = "Export $(@($Groups).Count) groups to Excel file '$($ExcelParams.Path)'"
                Write-Verbose $M
                Write-EventLog @EventOutParams -Message $M
                $Groups | Export-Excel @ExcelParams -TableName Groups -WorksheetName Groups
            }
            if ($RolGroups) {
                $M = "Export $(@($RolGroups).Count) ROL groups to Excel file '$($ExcelParams.Path)'"
                Write-Verbose $M
                Write-EventLog @EventOutParams -Message $M
                $RolGroups | Export-Excel @ExcelParams -TableName RolGroups -WorksheetName GroupsROL
            }
            if ($Users) {
                $M = "Export $(@($Users).Count) users to Excel file '$($ExcelParams.Path)'"
                Write-Verbose $M
                Write-EventLog @EventOutParams -Message $M
                $Users | Export-Excel @ExcelParams -TableName Users -WorksheetName Users
            }
            if ($GitUsers) {
                $M = "Export $(@($GitUsers).Count) GIT users to Excel file '$($ExcelParams.Path)'"
                Write-Verbose $M
                Write-EventLog @EventOutParams -Message $M
                $GitUsers | Export-Excel @ExcelParams -TableName GitUsers -WorksheetName UsersGIT
            }
            #endregion
        }

        $UsersHtmlList = $ComputersHtmlList = $GroupsHtmlList =
        $GitUsersHtmlList = $GroupMembersHtmlList = @()

        #region Test for incorrect ticket topic
        $File.Tickets | Get-Member -MemberType NoteProperty -EA Ignore |
        Where-Object {
            $AllObjects.GetEnumerator().Name -notcontains $_.Name
        } | ForEach-Object {
            throw "The parameter 'Tickets' in the file '$ImportFile' contains an invalid topic name '$($_.Name)', valid topic names are: '$(($AllObjects.GetEnumerator().Name | Sort-Object) -join "', '")'"
        }
        #endregion

        foreach (
            $A in
            $AllObjects.GetEnumerator() | Sort-Object { $_.Value.WorkSheetName }
        ) {
            Write-Verbose "Type '$($A.Value.Type)' worksheet '$($A.Value.WorksheetName)' item '$($A.Key)' data '$(@($A.Value.Data).Count)'"

            #region Test missing properties
            if (-not (
                    $A.Value.ContainsKey('Data') -and
                    $A.Value.ContainsKey('PropertyToExport') -and
                    $A.Value.Description -and
                    $A.Value.WorkSheetName -and
                    $A.Value.Type)
            ) {
                throw "Missing a property for worksheet '$($A.Value.WorkSheetName)' with description '$($A.Value.Description)'"
            }
            #endregion

            #region Test create tickets
            $createTicket = $false

            if (
                $File.Tickets |
                Get-Member -Name $A.Name -MemberType NoteProperty -EA Ignore
            ) {
                $createTicket = $true
            }
            #endregion

            #region Create email table rows and Excel file name
            $HtmlListItem = '<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f
            $(
                if ($A.Value['Count']) { $A.Value['Count'] }
                else { @($A.Value.Data).Count }
            ),
            $A.Value.WorksheetName,
            $(
                if ($createTicket) { $A.Value.Description + '<br><b>(AUTO TICKET)</b>' }
                else { $A.Value.Description }
            )

            switch ($A.Value.Type) {
                'User' {
                    $ExcelParams.Path = "$LogFile - Users.xlsx"
                    $UsersHtmlList += $HtmlListItem
                    break
                }
                'GitUser' {
                    $ExcelParams.Path = "$LogFile - GIT Users.xlsx"
                    $GitUsersHtmlList += $HtmlListItem
                    break
                }
                'Computer' {
                    $ExcelParams.Path = "$LogFile - Computers.xlsx"
                    $ComputersHtmlList += $HtmlListItem
                    break
                }
                'Group' {
                    $ExcelParams.Path = "$LogFile - Groups.xlsx"
                    $GroupsHtmlList += $HtmlListItem
                    break
                }
                'RolGroup' {
                    $ExcelParams.Path = "$LogFile - ROL Groups.xlsx"
                    $RolGroupsHtmlList += $HtmlListItem
                    break
                }
                'ListGroupMembers' {
                    $ExcelParams.Path = "$LogFile - Group members.xlsx"
                    $GroupMembersHtmlList += $HtmlListItem
                    break
                }
                default {
                    throw "The custom object type '$_' is not supported. Please implement this feature."
                }
            }
            #endregion

            if ($A.Value.Data -and $A.Value.PropertyToExport) {
                #region Create ticket
                if ($createTicket) {
                    $params = @{
                        ScriptName       = $ScriptName 
                        ServiceNow       = $file.ServiceNow 
                        TopicName        = $A.Name
                        TopicDescription = $A.Value.Description
                        TicketFields     = $File.Tickets."$($A.Name)" |
                        Select-Object -ExcludeProperty 'Exclude' 
                        Data             = $A.Value.Data | Select-Object (
                            @('SamAccountName') +
                            (
                                $A.Value.PropertyToExport |
                                Where-Object { $_ -ne 'SamAccountName' }
                            )
                        )
                        ScriptAdmin      = $ScriptAdmin
                    }
                    & $scriptCreateTicketsItem.FullName @params
                }
                #endregion

                if (-not $NoEmail) {
                    #region Export to Excel file
                    $M = "Export '$($A.Key)' with $(@($A.Value.Data).Count) objects to worksheet '$($A.Value.WorksheetName)' in Excel file '$($ExcelParams.Path)'"
                    Write-Verbose $M
                    Write-EventLog @EventOutParams -Message $M

                    $ExcelParams.TableName = $A.Value.WorksheetName
                    $ExcelParams.WorkSheetName = $A.Value.WorksheetName
                    $A.Value.Data | Select-Object $A.Value.PropertyToExport |
                    Export-Excel @ExcelParams
                    #endregion

                    $MailParams.Attachments += $ExcelParams.Path
                }
            }
        }

        #region Create HTML tables for the email
        Write-Verbose 'Create HTML for e-mail'

        $UsersHtmlTable = $ComputersHtmlTable = $GroupsHtmlTable = $RolGroupsHtmlTable =
        $GitUsersHtmlTable = $GroupMembersHtmlTable = $null

        if ($UsersHtmlList) {
            $UsersHtmlTable = "
            <h3>Users</h3>
                <table>
                    <tr><th>Quantity</th><th>Sheet name</th><th>Description</th></tr>
                    $($UsersHtmlList -join "`r`n")
                </table>
            "
        }
        if ($GitUsersHtmlList) {
            $GitUsersHtmlTable = "
            <h3>GIT Users</h3>
                <table>
                    <tr><th>Quantity</th><th>Sheet name</th><th>Description</th></tr>
                    $($GitUsersHtmlList -join "`r`n")
                </table>
            "
        }
        if ($ComputersHtmlList) {
            $ComputersHtmlTable = "
            <h3>Computers</h3>
                <table>
                    <tr><th>Quantity</th><th>Sheet name</th><th>Description</th></tr>
                    $($ComputersHtmlList -join "`r`n")
                </table>
            "
        }
        if ($GroupsHtmlList) {
            $GroupsHtmlTable = "
            <h3>Groups</h3>
                <table>
                    <tr><th>Quantity</th><th>Sheet name</th><th>Description</th></tr>
                    $($GroupsHtmlList -join "`r`n")
                </table>
            "
        }
        if ($RolGroupsHtmlList) {
            $RolGroupsHtmlTable = "
            <h3>ROL Groups</h3>
                <table>
                    <tr><th>Quantity</th><th>Sheet name</th><th>Description</th></tr>
                    $($RolGroupsHtmlList -join "`r`n")
                </table>
            "
        }
        if ($GroupMembersHtmlList) {
            $GroupMembersHtmlTable = "
            <h3>Group members</h3>
                <table>
                    <tr><th>Members</th><th>Sheet name</th><th>Name</th></tr>
                    $($GroupMembersHTMLList -join "`r`n")
                </table>
            "
        }
        #endregion

        #region Send the summary email
        $Issues = $AllObjects.GetEnumerator().where{ ($_.Value.Data) }

        $MailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Priority  = if ($Issues) { 'High' } else { 'Normal' }
            Subject   = if ($Issues) { "$($Issues.count) issues found" }
            else { 'Success - No issues found' }
            Save      = "$LogFile - Mail.html"
            LogFolder = $LogFolder
            Header    = $ScriptName
            Message   = @"
<p>$(if ($Issues) { 'Inconsistencies found in the active directory:' }
    else { 'No inconsistencies found in the active directory:' })</p>
$ComputersHtmlTable
$GroupsHtmlTable
$RolGroupsHtmlTable
$UsersHtmlTable
$GitUsersHtmlTable
$GroupMembersHtmlTable
$(if($ExcludedGroups) {
"<p>The members of the following groups were excluded from the user inconsistency checks: $($ExcludedGroups.Name -join ', ').</p>"})
<p><i>* Check the attachments for details</i></p>
"@
        }

        $MailParams.Attachments = $MailParams.Attachments | Select-Object -Unique

        if ($error) {
            # $Error | Select-Object @{N = 'M'; E = { $_.Exception.Message + $_.ScriptStackTrace } } | Select-Object -ExpandProperty M |
            $HTMLErrors = $Error.Exception.Message | Sort-Object -Unique |
            ConvertTo-HtmlListHC -Spacing Wide -Header 'Errors detected:'
            $MailParams.Message += $HTMLErrors

            $MailParams.Subject = "Failure - $($error.Count) errors"
        }

        $MailParams.Message += $OU | ConvertTo-OuNameHC -OU | Sort-Object |
        ConvertTo-HtmlListHC -Header 'Organizational units:'

        Remove-EmptyParamsHC -Name $MailParams
        Get-ScriptRuntimeHC -Stop
        if (-not $NoEmail) {
            Send-MailHC @MailParams
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        exit 1
    }
    finally {
        Write-EventLog @EventEndParams
    }
}