﻿#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testUser = Get-ADUser $env:USERNAME

    $testDate = Get-Date

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and ($Subject -eq 'FAILURE')
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $InactiveDays = 40

    $testInputFile = @{
        MailTo              = @('bob@contoso.com')
        InactiveDays        = $InactiveDays
        RolGroup            = @{
            Prefix             = "BEL ROL-"
            PlaceHolderAccount = "belsrvc"
        }
        Prefix              = @{
            QuotaGroup = "BEL ATT Quota home"
        }
        OU                  = @('OU=BEL,OU=EU,DC=contoso,DC=com')
        AllowedEmployeeType = @("Vendor", "Plant", "Kiosk")
        Group               = @(
            @{
                Name        = "BEL ATT Leaver"
                Type        = "Exclude"
                ListMembers = $true
            }
            @{
                Name        = "BEL ATT User No OCS"
                Type        = "NoOCS"
                ListMembers = $false
            }
            @{
                Name        = "BEL ATT No Deprovision"
                Type        = $null
                ListMembers = $true
            }
        )
        Git                 = @{
            OU          = "OU=GIT,DC=contoso,DC=net"
            CountryCode = @( "BE", "LU", "NL")
        }
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = (New-Item "TestDrive:/log" -ItemType Directory).FullName
    }

    Mock Send-MailHC
    Mock Write-EventLog
    Mock Get-ADCircularGroupsHC
    Mock Get-ADComputer
    Mock Get-ADGroup
    Mock Get-ADGroupMember
    Mock Get-ADDisplayNameHC
    Mock Get-ADDisplayNameFromSID
    Mock Get-ADObject { $true }
    Mock Get-ADOrganizationalUnit
    Mock Get-ADUser
    Mock Get-ADTSProfileHC
    Mock Test-ADOUExistsHC { $true }
}
Describe 'Prerequisites' {
    Context 'ImportFile' {
        It 'mandatory parameter' {
            (Get-Command $testScript).Parameters['ImportFile'].Attributes.Mandatory | 
            Should -BeTrue
        } 
        It 'file not existing' {
            .$testScript -ScriptName $testParams.ScriptName -LogFolder $testParams.LogFolder -ImportFile 'NotExisting.txt'
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "Cannot find path*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } 
        It 'OU missing' {
            $testNewFile = Copy-ObjectHC $testInputFile
            $testNewFile.OU = $null
            $testNewFile | ConvertTo-Json | Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'OU' found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } 
        It 'OU incorrect' {
            Mock Get-ADOrganizationalUnit { throw 'OU not found' }

            $testInputFile | ConvertTo-Json | Out-File @testOutParams

            .$testScript @testParams
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*OU*does not exist*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } 
        It 'OU has no country name set' {
            Mock Get-ADOrganizationalUnit {
                [PSCustomObject]@{
                    CanonicalName = 'contoso.com/EU/BEL'
                    Description   = $null
                    Country       = $null
                }
            }

            $testInputFile | ConvertTo-Json | Out-File @testOutParams

            .$testScript @testParams
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "The AD Organizational Unit*is incomplete*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } 
        It 'MailTo missing' {
            $testNewFile = Copy-ObjectHC $testInputFile
            $testNewFile.MailTo = $null
            $testNewFile | ConvertTo-Json | Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$ImportFile*MailTo*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } 
        It 'GITCountryCode missing' {
            $testNewFile = Copy-ObjectHC $testInputFile
            $testNewFile.Git.CountryCode = $null
            $testNewFile | ConvertTo-Json | Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$ImportFile*GITcountryCode*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } 
    }
    Context 'LogFolder' {
        It 'parameter optional' {
            (Get-Command $testScript).Parameters['LogFolder'].Attributes.Mandatory | 
            Should -BeFalse
        } 
        It 'send error mail when folder is not found' {
            $testInputFile | ConvertTo-Json | Out-File @testOutParams

            .$testScript -ScriptName $testParams.ScriptName -LogFolder 'NonExisting' -ImportFile $testParams.ImportFile

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Path*not found")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } 
    }
}
Describe 'Computers' {
    BeforeAll {
        Mock Get-ADGroup
        Mock Get-ADUser
        Mock Get-ADOrganizationalUnit {
            [PSCustomObject]@{
                CanonicalName = 'contoso.com/EU/BEL'
                Description   = 'Belgium'
            }
        }
        
        $testInputFile | ConvertTo-Json | Out-File @testOutParams
    }
    BeforeEach {
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse
    }
    It 'Last LogonDate over x days' {
        Mock Get-ADComputer {
            [PSCustomObject]@{
                Name          = 'PC1'
                Description   = 'Not ok, 3 months ago'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $true
                LastLogonDate = ($testDate).AddMonths( - 3)
            }
            [PSCustomObject]@{
                Name          = 'PC2'
                Description   = 'Not ok, just over the treshold'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $true
                LastLogonDate = ($testDate).AddDays(-$InactiveDays)
            }
            [PSCustomObject]@{
                Name          = 'PC3'
                Description   = 'Not ok, never logged on'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $true
                LastLogonDate = $null
            }
            [PSCustomObject]@{
                Name          = 'PC4'
                Description   = 'Not ok, never logged on'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $true
                LastLogonDate = ''
            }
            [PSCustomObject]@{
                Name          = 'PC5'
                Description   = 'Not ok, 1 day overdue'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $true
                LastLogonDate = ($testDate).AddDays( - ($InactiveDays + 1))
            }
            [PSCustomObject]@{
                Name          = 'PC6'
                Description   = 'Ok, enabled is false'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $false
                LastLogonDate = ($testDate).AddYears(-$InactiveDays)
            }
            [PSCustomObject]@{
                Name          = 'PC7'
                Description   = 'Ok, logged on today'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $true
                LastLogonDate = $testDate
            }
            [PSCustomObject]@{
                Name          = 'PC8'
                Description   = 'Ok, ignored because of disabled OU'
                CanonicalName = 'contoso.com/EU/BEL/Disabled/Computers/PC'
                Enabled       = $true
                LastLogonDate = $null
            }
        }

        .$testScript @testParams

        $AllObjects['Computer - Inactive'].Data.Name | 
        Should -Be @('PC1', 'PC2', 'PC3', 'PC4', 'PC5')
    } -Tag test
    It 'enabled in OU disabled' {
        Mock Get-ADComputer {
            [PSCustomObject]@{
                Name          = 'PC1'
                Description   = 'Ok, disabled in normal OU'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $false
            }
            [PSCustomObject]@{
                Name          = 'PC2'
                Description   = 'Ok, in disabled OU'
                CanonicalName = 'contoso.com/EU/BEL/Disabled/Computers/PC'
                Enabled       = $false
            }
            [PSCustomObject]@{
                Name          = 'PC3'
                Description   = 'Ok, in normal OU'
                CanonicalName = 'contoso.com/EU/BEL/Computers/PC'
                Enabled       = $true
            }
            [PSCustomObject]@{
                Name          = 'PC4'
                Description   = 'Not ok, enabled in disabled OU'
                CanonicalName = 'contoso.com/EU/BEL/Disabled/Computers/PC'
                Enabled       = $true
            }
            [PSCustomObject]@{
                Name          = 'PC5'
                Description   = 'Ok, disabled in disabled ou'
                CanonicalName = 'contoso.com/EU/BEL/Disabled/Computers/PC'
                Enabled       = $false
            }
        }

        .$testScript @testParams

        $AllObjects['Computer - EnabledInDisabledOU'].Data.Name | Should -Be @('PC4')
    } 
}
Describe 'ROL Groups' {
    BeforeAll {
        Mock Get-ADDisplayNameHC
        Mock Get-ADComputer
        Mock Get-ADGroupMember
        Mock Get-ADUser
        Mock Get-ADOrganizationalUnit {
            [PSCustomObject]@{
                CanonicalName = 'contoso.com/EU/BEL'
                Description   = 'Belgium'
            }
        }

        $testInputFile | ConvertTo-Json | Out-File @testOutParams
    }
    BeforeEach {
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse
    }
    It 'filter out non ROL groups' {
        Mock Get-ADGroup {
            $GroupName = 'BEL ROL-AGG-SAGREX Plant Manager'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'ROL group'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
            }

            $GroupName = 'BEL No rol group'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'No ROL group'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
            }

            $GroupName = 'BEL Rolos lovers'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'No ROL group'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
            }
        }

        .$testScript @testParams

        $Groups | Should -HaveCount 3
        $RolGroups.SamAccountName | Should -Be 'BEL ROL-AGG-SAGREX Plant Manager'
    } 
    It 'Mail address cannot be blank' {
        Mock Get-ADGroup {
            'Alain@gmail.com', 'Chuck.Norris@hc.com', '165465@something.com',
            'bob.dylan@heidelbergcement.com' | ForEach-Object {
                $GroupName = 'BEL ROL-STAFF-IT CorrectGroup{0}'
                New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                    SamAccountName = $GroupName
                    Mail           = $_
                    Description    = 'With mail address'
                    CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                }
            }
            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Mail           = ''
                Description    = 'Blank mail address'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 2'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Mail           = $null
                Description    = 'Blank mail address'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 3'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Mail           = ' '
                Description    = 'Blank mail address'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
            }
        }

        .$testScript @testParams

        $AllObjects['RolGroup - Mail'].Data.SamAccountName | Should -HaveCount 3
        @(
            'BEL ROL-STAFF-IT IncorrectGroup 1',
            'BEL ROL-STAFF-IT IncorrectGroup 2',
            'BEL ROL-STAFF-IT IncorrectGroup 3'
        ) | ForEach-Object {
            $AllObjects['RolGroup - Mail'].Data.SamAccountName | Should -Contain $_
        }
    } 
    It "GroupScope needs to be 'Universal'" {
        Mock Get-ADGroup {
            $GroupName = 'BEL ROL-STAFF-IT CorrectGroup'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'GroupScope ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupScope     = 'Universal'
                GroupCategory  = 'Security'
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'GroupScope not ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupScope     = 'Global'
                GroupCategory  = 'Security'
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 2'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'GroupScope not ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupScope     = 'DomainLocal'
                GroupCategory  = 'Security'
            }
        }

        .$testScript @testParams

        $AllObjects['RolGroup - GroupScope'].Data.SamAccountName | Should -HaveCount 2
        @(
            'BEL ROL-STAFF-IT IncorrectGroup 1',
            'BEL ROL-STAFF-IT IncorrectGroup 2'
        ) | ForEach-Object {
            $AllObjects['RolGroup - GroupScope'].Data.SamAccountName | Should -Contain $_
        }
    } 
    It "GroupCategory needs to be 'Security'" {
        Mock Get-ADGroup {
            $GroupName = 'BEL ROL-STAFF-IT CorrectGroup'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupScope     = 'Universal'
                GroupCategory  = 'Security'
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupScope     = 'Universal'
                GroupCategory  = 'Distribution'
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 2'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupScope     = 'Universal'
                GroupCategory  = 'Distribution'
            }
        }

        .$testScript @testParams

        $AllObjects['RolGroup - GroupCategory'].Data.SamAccountName | Should -HaveCount 2
        @(
            'BEL ROL-STAFF-IT IncorrectGroup 1',
            'BEL ROL-STAFF-IT IncorrectGroup 2'
        ) | ForEach-Object {
            $AllObjects['RolGroup - GroupCategory'].Data.SamAccountName | Should -Contain $_
        }
    } 
    It "'CN' needs to be equal to 'Name'" {
        Mock Get-ADGroup {
            # Not possible to use the Name field to create a fake group
            # $GroupName = 'BEL ROL-STAFF-IT CorrectGroup'
            # New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
            #     CN             = $GroupName
            #     SamAccountName = $GroupName
            #     Description    = 'CN ok'
            #     CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
            # }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                CN             = 'BEL DIS-STAFF-IT IncorrectGroup 1'
                SamAccountName = $GroupName
                Description    = 'CN not ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/1' -f $GroupName
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 2'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                CN             = 'BEL ROL-STAFF-IT kiwi'
                SamAccountName = $GroupName
                Description    = 'CN not ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
            }
        }

        .$testScript @testParams

        $AllObjects['RolGroup - CN'].Data.SamAccountName | Should -HaveCount 2
        @(
            'BEL ROL-STAFF-IT IncorrectGroup 1',
            'BEL ROL-STAFF-IT IncorrectGroup 2'
        ) | ForEach-Object {
            $AllObjects['RolGroup - CN'].Data.SamAccountName | Should -Contain $_
        }
    } 
    It "'DisplayName' not equal to 'Name' when the word 'ROL' is replaced with 'DIS'" {
        Mock Get-ADGroup {
            # Not possible to use the Name field to create a fake group
            # $GroupName = 'BEL ROL-STAFF-IT CorrectGroup'
            # New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
            #     DisplayName    = 'BEL DIS-STAFF-IT CorrectGroup'
            #     SamAccountName = $GroupName
            #     Description    = 'DisplayName ok'
            #     CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
            # }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                DisplayName    = $GroupName
                SamAccountName = $GroupName
                Description    = 'DisplayName not ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
            }

            $GroupName = 'BEL ROL-STAFF-IT IncorrectGroup 2'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                DisplayName    = 'BEL ROL-STAFF-IT kiwi'
                SamAccountName = $GroupName
                Description    = 'DisplayName not ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
            }
        }

        .$testScript @testParams

        $AllObjects['RolGroup - DisplayName'].Data.SamAccountName | Should -HaveCount 2
        @(
            'BEL ROL-STAFF-IT IncorrectGroup 1',
            'BEL ROL-STAFF-IT IncorrectGroup 2'
        ) | ForEach-Object {
            $AllObjects['RolGroup - DisplayName'].Data.SamAccountName | Should -Contain $_
        }
    } 
    It "'ManagedBy' blank" {
        Mock Get-ADGroup {
            $GroupName = 'BEL ROL-STAFF-IT With manager'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                CN             = $GroupName
                SamAccountName = $GroupName
                ManagedBy      = 'Bob'
                Description    = 'Ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
            }

            $GroupName = 'BEL ROL-STAFF-IT No manager'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                CN             = $GroupName
                SamAccountName = $GroupName
                ManagedBy      = $null
                Description    = 'Not ok'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
            }
        }

        .$testScript @testParams

        $AllObjects['RolGroup - ManagedBy'].Data.SamAccountName |
        Should -Be 'BEL ROL-STAFF-IT No manager'
    } 
}
Describe 'Groups' {
    BeforeAll {
        $testNewFile = @{
            MailTo                    = @('bob@soemwhere.com')
            InactiveDays              = 40
            QuotaGroupNameBegin       = "BEL Quota"
            PlaceHolderSamAccountName = 'keepMe'
            OU                        = @('OU=BEL,OU=EU,DC=contoso,DC=com')
            Group                     = @(
                @{
                    Name        = "Group1"
                    Type        = "Exclude"
                    ListMembers = $true
                }
                @{
                    Name        = "Group2"
                    Type        = "NoOCS"
                    ListMembers = $false
                }
                @{
                    Name        = "Group3"
                    Type        = $null
                    ListMembers = $true
                }
            )
            Git                       = @{
                OU          = "OU=GIT,DC=contoso,DC=net"
                CountryCode = @( "BE", "LU", "NL")
            }
        }
        $testNewFile | ConvertTo-Json | Out-File @testOutParams
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse

        Mock Get-ADComputer
        Mock Get-ADGroup
        Mock Get-ADObject { $true }
        Mock Get-ADOrganizationalUnit {
            [PSCustomObject]@{
                CanonicalName = 'contoso.com/EU/BEL'
                Description   = 'Belgium'
            }
        }
        Mock Get-ADGroupMember {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'cnorris'
                CanonicalName     = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                DistinguishedName = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
            }
        }
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'cnorris'
                CanonicalName     = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                DistinguishedName = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
            }
        }

        .$testScript @testParams
    }
    It 'excluded group members' {
        $ExcludedGroups.Members.SamAccountName | Should -Be 'cnorris'
        $ExcludedGroups.Name | Should -Be 'Group1'
    } 
    It 'list group members' {
        $GroupMembers.Members.SamAccountName | Should -Be @('cnorris', 'cnorris')
        $GroupMembers.Name | Should -Be @('Group1', 'Group3')
    } 
    It 'distribution list no manager' {
        Mock Get-ADGroup {
            $GroupName = 'BEL DIS-AGG-SAGREX Plant Manager'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'DIS group no maanager'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Distribution'
                GroupScope     = 'Universal'
                ManagedBy      = $null
            }

            $GroupName = 'BEL DIS-AGG-SAGREX Employee'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'No DIS group no maanager'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
                ManagedBy      = $null
            }

            $GroupName = 'BEL DIS-AGG-SAGREX District Manager'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'DIS group with maanager'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Distribution'
                GroupScope     = 'Universal'
                ManagedBy      = 'Bob'
            }

            $GroupName = 'BEL ROL-RMC-IB Plant Managers'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'no DIS group with manager'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Security'
                GroupScope     = 'Universal'
                ManagedBy      = 'Jake'
            }
        }

        .$testScript @testParams

        $AllObjects['Group - DisListNoManager'].Data.SamAccountName |
        Should -Be 'BEL DIS-AGG-SAGREX Plant Manager'
    } 
    It 'member not in OU' {
        Mock Get-ADGroup {
            $GroupName = 'Group1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'DIS group no maanager'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Distribution'
                GroupScope     = 'Universal'
                ManagedBy      = $null
            }
        }

        Mock Get-ADGroupMember {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'mike'
                ObjectClass       = 'User'
                CanonicalName     = 'CN=Mike,CN=Users,OU=BEL,OU=EU,DC=contoso,DC=com'
                distinguishedName = 'CN=Mike,CN=Users,OU=BEL,OU=EU,DC=contoso,DC=com'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'bob'
                ObjectClass       = 'User'
                CanonicalName     = 'CN=Bob,CN=Users,OU=XXX,OU=EU,DC=contoso,DC=com'
                distinguishedName = 'CN=Bob,CN=Users,OU=XXX,OU=EU,DC=contoso,DC=com' # incorrect OU
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'jake'
                ObjectClass       = 'User'
                CanonicalName     = 'CN=Jake,CN=Users,OU=BEL,OU=EU,DC=contoso,DC=com'
                distinguishedName = 'CN=Jake,CN=Users,OU=BEL,OU=EU,DC=contoso,DC=com'
            }
        }

        .$testScript @testParams

        $AllObjects['Group - MembersNotInOU'].Data.GroupName |
        Should -Be 'Group1'

        $AllObjects['Group - MembersNotInOU'].Data.UserSamAccountName |
        Should -Be 'bob'
    } 
    It 'orphaned members' {
        Mock Get-ADGroup {
            $GroupName = 'Group1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'Incorrect group'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Distribution'
                GroupScope     = 'Universal'
                ManagedBy      = $null
            }
        }

        Mock Get-ADGroupMember {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'mike'
                ObjectClass       = 'User'
                CanonicalName     = 'CN=Mike,CN=Users,OU=BEL,OU=EU,DC=contoso,DC=com'
                distinguishedName = $null # orphaned because of no distinguishedName
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'bob'
                ObjectClass       = 'User'
                CanonicalName     = 'CN=Bob,CN=Users,OU=XXX,OU=EU,DC=contoso,DC=com'
                distinguishedName = 'CN=Bob,CN=Users,OU=XXX,OU=EU,DC=contoso,DC=com'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'jake'
                ObjectClass       = 'User'
                CanonicalName     = 'CN=Jake,CN=Users,OU=BEL,OU=EU,DC=contoso,DC=com'
                distinguishedName = 'CN=Jake,CN=Users,OU=BEL,OU=EU,DC=contoso,DC=com'
            }
        }

        .$testScript @testParams

        $AllObjects['Group - GroupsWithOrphans'].Data.SamAccountName |
        Should -Be 'Group1'
    } 
    It 'non traversable groups' {
        Mock Get-ADGroup {
            $GroupName = 'Group1'
            New-Object Microsoft.ActiveDirectory.Management.ADGroup Identity -Property @{
                SamAccountName = $GroupName
                Description    = 'Incorrect group'
                CanonicalName  = 'contoso.com/EU/BEL/Groups/{0}' -f $GroupName
                GroupCategory  = 'Distribution'
                GroupScope     = 'Universal'
                ManagedBy      = $null
            }
        }

        Mock Get-ADGroupMember {
            throw 'Group members cannot be retrieved'
        } -ParameterFilter {
            'Group1' -eq $Identity.SamAccountName
        }

        .$testScript @testParams

        $AllObjects['Group - NonTraversableGroups'].Data.SamAccountName |
        Should -Be 'Group1'
    } 
}
Describe 'Exclude users' {
    BeforeEach {
        Mock Get-ADComputer
        Mock Get-ADGroup
        Mock Get-ADObject { $true }
        Mock Get-ADOrganizationalUnit {
            [PSCustomObject]@{
                CanonicalName = 'contoso.com/EU/BEL'
                Description   = 'Belgium'
            }
        }

        $testInputFile | ConvertTo-Json | Out-File @testOutParams
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse
    }
    It "that are member of a group in 'ExcludedGroups'" {
        Mock Get-ADGroupMember {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'jlewis'
                GivenName         = 'John'
                Surname           = 'Doe'
                DistinguishedName = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                CanonicalName     = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                scriptPath        = 'logon.bat'
            }
        }
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'jlewis'
                GivenName         = 'John'
                Surname           = 'Doe'
                DistinguishedName = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                CanonicalName     = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                scriptPath        = 'logon.bat'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'mtyson'
                GivenName         = 'John'
                Surname           = 'Doe'
                DistinguishedName = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                CanonicalName     = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                scriptPath        = ''
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'norrisc'
                GivenName         = 'John'
                Surname           = 'Doe'
                DistinguishedName = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                CanonicalName     = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                scriptPath        = ''
            }
        }
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'jlewis'
                GivenName         = 'John'
                Surname           = 'Doe'
                DistinguishedName = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                CanonicalName     = 'CN=Doe John,CN=Users,DC=contoso,DC=com'
                scriptPath        = ''
            }
        } -ParameterFilter {
            'jlewis' -eq $Identity.SamAccountName
        }

        .$testScript @testParams

        $Users.SamAccountName | Should -Be @(
            'mtyson',
            'norrisc'
        )

        $ExcludedGroups.Members.SamAccountName | Should -Be 'jlewis'
        $MailParams.Message | Should -Match  ($testInputFile.Group.Where( { $_.Type -eq 'Exclude' })).Name
    } 
    It "in the OU 'Disabled'" {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'DisregardedUser'
                EmployeeType      = 'Employee'
                Description       = 'Disabled OU user'
                DisplayName       = 'Dummy, DisregardedUser (Somewhere) BEL'
                CanonicalName     = 'contoso.com/EU/BEL/Disabled/Users/Dummy, DisregardedUser (Somewhere) BEL'
                DistinguishedName = 'contoso.com/EU/BEL/Disabled/Users/Dummy, DisregardedUser (Somewhere) BEL'
                ScriptPath        = ''
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Resource Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, Correct (Somewhere) BEL'
                employeeType      = 'Resource'
            }
        }

        .$testScript @testParams

        $Users.SamAccountName | Should -Be 'Correct'
    } 
    It "in the OU 'Terminated Users'" {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'DisregardedUser'
                EmployeeType      = 'Employee'
                Description       = 'Terminated user'
                DisplayName       = 'Dummy, DisregardedUser (Somewhere) BEL'
                CanonicalName     = 'contoso.com/EU/BEL/Terminated Users/Dummy, DisregardedUser (Somewhere) BEL'
                DistinguishedName = 'contoso.com/EU/BEL/Terminated Users/Dummy, DisregardedUser (Somewhere) BEL'
                ScriptPath        = ''
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, Correct (Somewhere) BEL'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Resource Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                employeeType      = 'Resource'
            }
        }

        .$testScript @testParams

        $Users.SamAccountName | Should -Be 'Correct'
    } 
}
Describe 'Users' {
    BeforeAll {
        Mock Get-ADComputer
        Mock Get-ADGroup
        Mock Get-ADGroupMember
        Mock Get-ADOrganizationalUnit {
            [PSCustomObject]@{
                CanonicalName = 'contoso.com/EU/BEL'
                Description   = 'Belgium'
            }
        }
    }
    BeforeEach {
        $testInputFile | ConvertTo-Json | Out-File @testOutParams
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse
    }
    It 'country code is not matching the OU country code' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            'France', 'Germany', 'Luxembourg' | ForEach-Object {
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName = 'IncorrectUser'
                    co             = $_
                    Description    = 'Country incorrect'
                    DisplayName    = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName  = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    ScriptPath     = ''
                }
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'IncorrectUser'
                co                = $null
                Description       = 'Country incorrect'
                DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                ScriptPath        = ''
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'CorrectUser'
                co                = 'Belgium'
                Description       = 'Country correct'
                DisplayName       = 'Dummy, CorrectUser (Somewhere) BEL'
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                ScriptPath        = ''
            }
        }

        .$testScript @testParams

        $AllObjects['User - CountryNotMatchingOU'].Data.SamAccountName | Should -Be @(
            'IncorrectUser',
            'IncorrectUser',
            'IncorrectUser',
            'IncorrectUser'
        )
    } 
    It 'manager of self' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'norrisc'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                Manager           = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                ScriptPath        = ''
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'lswagger'
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                Manager           = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                ScriptPath        = ''
            }
        }

        .$testScript @testParams

        $AllObjects['User - ManagerOfSelf'].Data.SamAccountName | Should -eq 'norrisc'
    } 
    It 'display name wrong' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris, Chuck (Braine L’Alleud) BEL"
                EmployeeType      = 'Employee'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect plant user'
                DisplayName       = "Norris, Chuck (Braine L’Alleud) BEL"
                EmployeeType      = 'Plant'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect kiosk user'
                DisplayName       = "Norris, Chuck (Braine L’Alleud) BEL"
                EmployeeType      = 'Kiosk'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect resource account'
                DisplayName       = 'wrong'
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Resource accounts/Dummy, CorrectUser (Somewhere) BEL'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect service account'
                DisplayName       = 'wrong'
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Service accounts/Dummy, CorrectUser (Somewhere) BEL'
            }
        }
        Mock Compare-ADobjectNameHC {
            @{
                Valid = $false
            }
        }

        .$testScript @testParams

        $AllObjects['User - DisplayNameWrong'].Data.SamAccountName | Should -Be @(
            'Incorrect plant user',
            'Incorrect kiosk user',
            'Incorrect resource account',
            'Incorrect service account'
        )
    } 
    It 'duplicate display name' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = 'Marley, Bob'
                EmployeeType      = 'Employee'
                DistinguishedName = "CN=Marley\, Bob,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Service accounts/Dummy, User'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = 'Craig, Daniel'
                DistinguishedName = "CN=Daniel\, Craig,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Service accounts/Dummy, User'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect'
                DisplayName       = "Norris, Chuck (Braine L’Alleud) BEL"
                EmployeeType      = 'Plant'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Service accounts/Dummy, User'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect'
                DisplayName       = "Norris, Chuck (Braine L’Alleud) BEL"
                EmployeeType      = 'Kiosk'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Service accounts/Dummy, User'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect'
                DisplayName       = "Norris, Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Service accounts/Dummy, User'
            }
        }

        .$testScript @testParams

        $AllObjects['User - DisplayNameNotUnique'].Data.SamAccountName | Should -Be @(
            'Incorrect',
            'Incorrect',
            'Incorrect'
        )
    } 
    It 'TS Home directory does not exist' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                HomeDirectory     = $here
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect'
                HomeDirectory     = 'Does not exit'
                DistinguishedName = "CN=Incorrect,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
        }
        Mock Get-ADTSProfileHC {
            'Does not exit'
        } -ParameterFilter {
            $DistinguishedName -eq "CN=Incorrect,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
        }

        .$testScript @testParams

        $AllObjects['User - TSHomeDirNotExist'].Data.SamAccountName | Should -Be 'Incorrect'
    } 
    It 'TS profile does not exist' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                HomeDirectory     = $here
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Incorrect'
                HomeDirectory     = 'Does not exit'
                DistinguishedName = "CN=Incorrect,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
        }
        Mock Get-ADTSProfileHC {
            'Does not exit'
        } -ParameterFilter {
            $DistinguishedName -eq "CN=Incorrect,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
        }

        .$testScript @testParams

        $AllObjects['User - TSProfileNotExisting'].Data.SamAccountName | Should -Be 'Incorrect'
    } 
    It 'employeeType not allowed' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'norrisc'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                employeeType      = 'Vendor'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'lswagger'
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                employeeType      = 'Unknown'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'cdaniel'
                DistinguishedName = "CN=Daniel Craig\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                employeeType      = $null
            }
        }

        $testNewFile = Copy-ObjectHC $testInputFile
        $testNewFile.AllowedEmployeeType = @('Vendor', 'Plant')
        $testNewFile | ConvertTo-Json | Out-File @testOutParams

        .$testScript @testParams

        $AllObjects['User - EmployeeTypeNotAllowed'].Data.SamAccountName | Should -Be @('lswagger', 'cdaniel')
    } 
    It 'employeeType Vendor' {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'cnorris'
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                employeeType      = 'Vendor'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'lswagger'
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                employeeType      = 'Plant'
            }
        }

        .$testScript @testParams

        $AllObjects['User - EmployeeTypeVendor'].Data.SamAccountName | Should -eq 'cnorris'
    } 
    It "HomeDirectory not starting with '\\GROUPHC.NET\BNL\HOME\'
            and excluding EmployeeType 'Service accounts' and 'Resource accounts'" {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'InCorrect'
                DisplayName       = "Lee Swagger, Bob (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\grouphc.net\bnl\lixhe\home\bbartels"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, InCorrect (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'InCorrect'
                DisplayName       = "Lee Swagger, Bob (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\GROUPHC.NET\BNL\Centralized\HOME"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, InCorrect (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\grouphc.net\bnl\home\Correct"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\GROUPHC.NET\BNL\HOME\Centralized\Correct"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Service Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\grouphc.net\bnl\wrong"
                CanonicalName     = 'contoso.com/EU/BEL/Service Accounts/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Service'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Resource Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\GROUPHC.NET\BNL\wrong"
                CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Resource'
            }
        }

        .$testScript @testParams

        $AllObjects['User - HomeDirWrong'].Data.SamAccountName | Should -Be @('InCorrect', 'InCorrect')
    } 
    It "HomeDirectory set for EmployeeType 'Service' and 'Resource' when it's not needed" {
        Mock Get-ADGroupMember
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'InCorrect'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Service Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\grouphc.net\bnl\wrong"
                CanonicalName     = 'contoso.com/EU/BEL/Service Accounts/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Service'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'InCorrect'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Resource Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\GROUPHC.NET\BNL\wrong"
                CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Resource'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Lee Swagger, Bob (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\grouphc.net\bnl\lixhe\home\bbartels"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, InCorrect (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Lee Swagger, Bob (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Lee Swagger\, Bob (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\GROUPHC.NET\BNL\Centralized\HOME"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, InCorrect (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\grouphc.net\bnl\home\Correct"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Users,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = "\\GROUPHC.NET\BNL\HOME\Centralized\Correct"
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Employee'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Service Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = $null
                CanonicalName     = 'contoso.com/EU/BEL/Service Accounts/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Service'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'Correct'
                DisplayName       = "Norris Chuck (Braine L’Alleud) BEL"
                DistinguishedName = "CN=Norris\, Chuck (Braine L’Alleud) BEL,OU=Resource Accounts,OU=BEL,OU=EU,DC=contoso,DC=com"
                homeDirectory     = ''
                CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, Correct (Somewhere) BEL'
                ScriptPath        = ''
                employeeType      = 'Resource'
            }
        }

        .$testScript @testParams

        $AllObjects['User - HomeDirNotNeeded'].Data.SamAccountName | Should -Be @('InCorrect', 'InCorrect')
    } 
    Context 'description wrong' {
        It "OU 'Service accounts' needs to be 'Service' or 'Service - Description'" {
            Mock Get-ADUser {
                'Some stuff that is not ok', 'Service account', 'apples', 'Service-' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'IncorrectUser'
                        EmployeeType      = 'Service'
                        Description       = $_
                        DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Service Accounts/Dummy, IncorrectUser (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Service Accounts/Dummy, IncorrectUser (Somewhere) BEL'
                        ScriptPath        = ''
                    }
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Service'
                    Description       = ''
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Service Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Service Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    ScriptPath        = ''
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Service'
                    Description       = $null
                    DisplayName       = 'Dummy, IncorrectUser{0} (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Service Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Service Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    ScriptPath        = ''
                }
                'Service', 'Service - Something useful' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'CorrectUser'
                        EmployeeType      = 'Service'
                        Description       = $_
                        DisplayName       = 'Dummy, CorrectUser{0} (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Service Accounts/Dummy, CorrectUser{0} (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Service Accounts/Dummy, CorrectUser{0} (Somewhere) BEL'
                        ScriptPath        = ''
                    }
                }
            }

            .$testScript @testParams

            $AllObjects['User - DescriptionWrong'].Data.SamAccountName | Should -Be @(0..5).ForEach( { 'IncorrectUser' })
        } 
        It "OU 'Resource accounts' needs to be
                'Shared mailbox' or 'Shared mailbox  - Description',
                'Meeting room' or 'Meeting room - Description',
                'Shared mailbox' or 'Shared mailbox - Description'" {
            Mock Get-ADUser {
                'Some stuff that is not ok', 'Resource', 'Room', 'Meeting', 'Resource - ' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'IncorrectUser'
                        EmployeeType      = 'Resource'
                        Description       = $_
                        DisplayName       = 'Dummy, IncorrectUser{0} (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Resource Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                        ScriptPath        = ''
                    }
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Resource'
                    Description       = ''
                    DisplayName       = 'Dummy, IncorrectUser{0} (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Resource Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    ScriptPath        = ''
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Resource'
                    Description       = $null
                    DisplayName       = 'Dummy, IncorrectUser{0} (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Resource Accounts/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    ScriptPath        = ''
                }
                'Shared mailbox', 'Shared mailbox - Used for something', 'Meeting room', 'Meeting room - Room for meetings' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'CorrectUser'
                        EmployeeType      = 'Resource'
                        Description       = $_
                        DisplayName       = 'Dummy, CorrectUser{0} (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Resource Accounts/Dummy, CorrectUser{0} (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Resource Accounts/Dummy, CorrectUser{0} (Somewhere) BEL'
                        ScriptPath        = ''
                    }
                }
            }

            .$testScript @testParams

            $AllObjects['User - DescriptionWrong'].Data.SamAccountName | Should -Be @(0..6).ForEach( { 'IncorrectUser' })
        } 
        It "OU 'Users' with EmployeeType 'Kiosk' needs to be 'Kiosk' or 'Kiosk - Description'" {
            Mock Get-ADUser {
                'Some stuff that is not ok', 'Kiosk account', 'kiosk', 'Kiosk - ' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'IncorrectUser'
                        EmployeeType      = 'Kiosk'
                        Description       = $_
                        DisplayName       = 'Dummy, IncorrectUser{0} (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    }
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Kiosk'
                    Description       = ''
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Kiosk'
                    Description       = $null
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                }
                'Kiosk', 'Kiosk - Used by Bob' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName = 'CorrectUser'
                        EmployeeType   = 'Kiosk'
                        Description    = $_
                        DisplayName    = 'Dummy, CorrectUser (Somewhere) BEL'
                        CanonicalName  = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                    }
                }
            }

            .$testScript @testParams

            $AllObjects['User - DescriptionWrong'].Data.SamAccountName | Should -Be @(0..5).ForEach( { 'IncorrectUser' })
        } 
        It "OU 'Users' with EmployeeType 'Plant' needs to be 'Plant' or 'Plant - Description'" {
            Mock Get-ADUser {
                'Some stuff that is not ok', 'Plant account', 'plant', 'Plant - ' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'IncorrectUser'
                        EmployeeType      = 'Plant'
                        Description       = $_
                        DisplayName       = 'Dummy, IncorrectUser{0} (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser{0} (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser{0} (Somewhere) BEL'
                    }
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Plant'
                    Description       = ''
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Plant'
                    Description       = $null
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                }
                'Plant', 'Plant - Used by Bob' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'CorrectUser'
                        EmployeeType      = 'Plant'
                        Description       = $_
                        DisplayName       = 'Dummy, CorrectUser (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                    }
                }
            }

            .$testScript @testParams

            $AllObjects['User - DescriptionWrong'].Data.SamAccountName | Should -Be @(0..5).ForEach( { 'IncorrectUser' })
        } 
        It "OU 'Users' with EmployeeType 'Vendor' needs to be 'Vendor' or 'Vendor - Description'" {
            Mock Get-ADUser {
                'Some stuff that is not ok', 'Vendor account', 'vendor', 'Vendor - ' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'IncorrectUser'
                        EmployeeType      = 'Vendor'
                        Description       = $_
                        DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    }
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Vendor'
                    Description       = ''
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'IncorrectUser'
                    EmployeeType      = 'Vendor'
                    Description       = $null
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                }
                'Vendor', 'Vendor - Used by Bob' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'CorrectUser'
                        EmployeeType      = 'Vendor'
                        Description       = $_
                        DisplayName       = 'Dummy, CorrectUser (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                    }
                }
            }

            .$testScript @testParams

            $AllObjects['User - DescriptionWrong'].Data.SamAccountName | Should -Be @(0..5).ForEach( { 'IncorrectUser' })
        } 
        It "OU 'Users' with EmployeeType 'Employee' has no standard" {
            Mock Get-ADUser {
                'Some stuff that is ok', 'All text is good', 'kiwis', 'Kiosk - ',
                'Vendor', 'Plant - Used by Bob' | ForEach-Object {
                    New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                        SamAccountName    = 'CorrectUser'
                        EmployeeType      = 'Employee'
                        Description       = $_
                        DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                        CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                        DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                    }
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'CorrectUser'
                    EmployeeType      = 'Employee'
                    Description       = ''
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                }
                New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                    SamAccountName    = 'CorrectUser'
                    EmployeeType      = 'Employee'
                    Description       = $null
                    DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                    CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                    DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, CorrectUser (Somewhere) BEL'
                }
            }

            .$testScript @testParams

            $AllObjects['User - DescriptionWrong'].Data.SamAccountName | Should -BeNullOrEmpty
        } 
    }
}
Describe 'GIT users' {
    BeforeAll {
        Mock Get-ADComputer
        Mock Get-ADGroup
        Mock Get-ADGroupMember
        Mock Get-ADUser
        Mock Get-ADOrganizationalUnit {
            [PSCustomObject]@{
                CanonicalName = 'contoso.com/EU/BEL'
                Description   = 'Belgium'
            }
        }
    }
    BeforeEach {
        $testInputFile | ConvertTo-Json | Out-File @testOutParams
        Remove-Item -Path "$($testParams.LogFolder)\*" -Recurse
    }
    It 'no manager' {
        Mock Get-ADUser {
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'IncorrectUser'
                Country           = 'BE'
                Manager           = ''
                Enabled           = $true
                DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
            New-Object Microsoft.ActiveDirectory.Management.ADUser Identity -Property @{
                SamAccountName    = 'IncorrectUser'
                Country           = 'BE'
                Manager           = $null
                Enabled           = $true
                DisplayName       = 'Dummy, IncorrectUser (Somewhere) BEL'
                CanonicalName     = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
                DistinguishedName = 'contoso.com/EU/BEL/Users/Dummy, IncorrectUser (Somewhere) BEL'
            }
        } -ParameterFilter {
            $SearchBase -eq $GITOU
        }

        .$testScript @testParams

        $AllObjects['GitUser - NoManger'].Data.SamAccountName | Should -Be @(0..1).ForEach( { 'IncorrectUser' })
    } 
}    
