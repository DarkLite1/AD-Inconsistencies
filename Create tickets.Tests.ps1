#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        ScriptName   = 'Test'
        ServiceNow   = [PSCustomObject]@{
            CredentialsFilePath = (New-Item -Path 'TestDrive:\a.json' -ItemType File).FullName
            Environment         = 'Test'
            TicketFields        = [PSCustomObject]@{
                Caller = 'x'
            }
        }
        TicketFields = [PSCustomObject]@{
            ShortDescription = 'x'
        }
        TopicName    = 'Computer - Inactive'
        Data         = @(
            [PSCustomObject]@{
                SamAccountName = 'PC1'
            }
        )
        ScriptAdmin  = 'bob@contoso.com'
    }

    @{
        Test = @{
            Uri          = 'testUri'
            UserName     = 'testUserName'
            Password     = 'testPassword'
            ClientId     = 'testClientId'
            ClientSecret = 'testClientSecret'
        }
        Prod = @{
            Uri          = 'prodUri'
            UserName     = 'prodUserName'
            Password     = 'prodPassword'
            ClientId     = 'prodClientId'
            ClientSecret = 'prodClientSecret'
        }
    } | ConvertTo-Json | 
    Out-File -FilePath $TestParams.ServiceNow.CredentialsFilePath

    function Copy-ObjectHC {
        <#
        .SYNOPSIS
            Make a deep copy of an object using JSON serialization.

        .DESCRIPTION
            Uses ConvertTo-Json and ConvertFrom-Json to create an independent
            copy of an object. This method is generally effective for objects
            that can be represented in JSON format.

        .PARAMETER InputObject
            The object to copy.

        .EXAMPLE
            $newArray = Copy-ObjectHC -InputObject $originalArray
        #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $jsonString = $InputObject | ConvertTo-Json -Depth 100

        $deepCopy = $jsonString | ConvertFrom-Json -AsHashtable

        return $deepCopy
    }

    Mock Get-ServiceNowRecord
    Mock New-ServiceNowIncident { 
        @{
            number = 1
        } 
    }
    Mock New-ServiceNowSession { $true }
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'ScriptName', 'ServiceNow',
        'ScriptAdmin',
        'TopicName', 'Data', 'TicketFields'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'an error is thrown when' {
    BeforeEach {
        $ServiceNowSession = $null
        $testNewParams = Copy-ObjectHC $TestParams
    }
    Context 'property' {
        It 'ServiceNow.<_> is not found' -ForEach @(
            'CredentialsFilePath', 'Environment', 'TicketFields'
        ) {
            $testNewParams.ServiceNow.$_ = $null

            .$testScript @testNewParams

            Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
                ($Message -like "*Property 'ServiceNow.$_' not found*")
            }

            $LASTEXITCODE | Should -Be 1
        }
        It 'ServiceNow.CredentialsFilePath does not exist' {
            $testNewParams.ServiceNow.CredentialsFilePath = 'TestDrive:\NotExisting.json'

            .$testScript @testNewParams

            Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
                ($Message -like "*Failed to import the ServiceNow environment file 'TestDrive:\NotExisting.json': *")
            }

            $LASTEXITCODE | Should -Be 1
        }
        It 'ServiceNow.TicketFields is missing a mandatory property' {
            $testNewParams.ServiceNow.TicketFields = [PSCustomObject]@{
                # Caller missing
                ShortDescription = 'short description'
            }

            .$testScript @testNewParams

            Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
                ($Message -like "*Field 'Caller' not found. This field is mandatory to create a ticket in ServiceNow*")
            }

            $LASTEXITCODE | Should -Be 1
        }
    }
}
Describe 'when a ticket was already created for an issue and not closed' {
    BeforeAll {
        $testUniqueAdObjectIssueID = 'PSID_AD-Inconsistencies_Computer---Inactive_PC1'

        Mock Get-ServiceNowRecord {
            @{
                number = 5 
            }
        } -ParameterFilter {
            ($Table -eq 'incident' ) -and
            ($Filter[0][2] -like "$testUniqueAdObjectIssueID")
        }

        $testNewParams = $testParams.Clone()
        $testNewParams.Data = @(
            [PSCustomObject]@{
                SamAccountName = 'PC1'
            }
        )

        .$testScript @testNewParams
    }
    It 'create an ID to identify the AD object and its issue' {
        $adObjectIssueId | Should -BeExactly $testUniqueAdObjectIssueID
    }
    It 'check if a ticket is created for this ID' {
        Should -Invoke Get-ServiceNowRecord -Scope Describe -Times 1 -Exactly -ParameterFilter {
            ($Table -eq 'incident' ) -and
            ($Filter[0][2] -like "$testUniqueAdObjectIssueID")
        }
    }
    It 'do not create a new ticket' {
        Should -Not -Invoke New-ServiceNowIncident -Scope Describe
    }
}
Describe 'when no ticket was created or the ticket was closed' {
    BeforeAll {
        $testUniqueAdObjectIssueID = 'PSID_AD-Inconsistencies_Computer---Inactive_PC2'

        Mock Get-ServiceNowRecord {
        } -ParameterFilter {
            ($Table -eq 'incident' ) -and
            ($Filter[0][2] -like "$testUniqueAdObjectIssueID")
        }

        $testNewParams = $testParams.Clone()
        $testNewParams.Data = @(
            [PSCustomObject]@{
                SamAccountName = 'PC2'
            }
        )
        $testNewParams.ServiceNow.TicketFields = [PSCustomObject]@{
            Caller           = 'bob'
            ShortDescription = 'short description'
            Description      = 'description'
            Subcategory      = 'MS Office'
            InputData        = [PSCustomObject]@{
                service_offering = 'service offering'
                impact           = 2
            }
        }
        $testNewParams.TicketFields = [PSCustomObject]@{
            ShortDescription = 'short description winner'
            Category         = 'Software'
            InputData        = [PSCustomObject]@{
                service_offering = 'service offering winner'
            }
        }

        .$testScript @testNewParams
    }
    It 'create an ID to identify the AD object and its issue' {
        $adObjectIssueId | Should -BeExactly $testUniqueAdObjectIssueID
    }
    It 'check if a ticket is created for this ID' {
        Should -Invoke Get-ServiceNowRecord -Scope Describe -Times 1 -Exactly -ParameterFilter {
            ($Table -eq 'incident' ) -and
            ($Filter[0][2] -like "$testUniqueAdObjectIssueID")
        }
    }
    It 'create a new ticket' {
        Should -Invoke New-ServiceNowIncident -Scope Describe -Times 1 -Exactly -ParameterFilter {
            ($Caller -eq 'bob') -and
            ($ShortDescription -eq 'short description winner') -and
            ($Description -like '*description*SamAccountName*PC2*>PowerShell ID: PSID_AD-Inconsistencies_Computer---Inactive_PC2 (do not remove)*') -and
            ($Subcategory -eq 'MS Office') -and
            ($InputData['service_offering'] -eq 'service offering winner') -and
            ($InputData['impact'] -eq 2)
        }
    }
} -Tag test