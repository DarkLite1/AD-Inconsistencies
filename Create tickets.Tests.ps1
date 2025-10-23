#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        ScriptName       = 'Test'
        ServiceNow       = [PSCustomObject]@{
            CredentialsFilePath = (New-Item -Path 'TestDrive:\a.json' -ItemType File).FullName
            Environment         = 'Test'
        }
        SQLDatabase      = 'Test'
        TopicName        = 'Computer - Inactive'
        TopicDescription = "'LastLogonDate' over x days"
        Data             = @(
            [PSCustomObject]@{
                SamAccountName = 'PC1'
            }
        )
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

    Mock New-CherwellTicketHC { 1 }
    Mock Save-TicketInSqlHC
    Mock Send-MailHC
    Mock Write-EventLog
    Mock Invoke-Sqlcmd
    Mock Invoke-Sqlcmd -ParameterFilter {
        $Query -like "*FROM $SQLTableTicketsDefaults*"
    } -MockWith {
        [PSCustomObject]@{
            Requester          = 'jack'
            SubmittedBy        = 'mike'
            ServiceCountryCode = 'BNL'
        }
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'ScriptName', 'Environment', 'TopicName', 'TopicDescription', 'Data'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'an error is thrown when' {
    It 'no ticket default values are found in SQL' {
        Mock Invoke-Sqlcmd -ParameterFilter {
            $Query -like "*FROM $SQLTableTicketsDefaults*"
        }

        .$testScript @TestParams

        Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
            ($Message -like '*No ticket default values found*')
        }
        
        $LASTEXITCODE | Should -Be 1
    }
    Context 'the import file property' {
        It 'TicketFields contains unknown ticket fields' {
            $testNewParams = Copy-ObjectHC $TestParams
            $testNewParams.TicketFields = [PSCustomObject]@{
                incorrectFieldName = 'x'
            }

            .$testScript @testNewParams

            Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
                ($Message -like "*Field name 'incorrectFieldName' not valid*")
            }

            $LASTEXITCODE | Should -Be 1
            Should -Not -Invoke New-CherwellTicketHC
        }
    }
    Context 'property' {
        It 'ServiceNow.<_> is not found' -ForEach @(
            'CredentialsFilePath', 'Environment'
        ) {
            $testNewParams = Copy-ObjectHC $TestParams
            $testNewParams.ServiceNow.$_ = $null

            $ServiceNowSession = $null

            .$testScript @testNewParams

            Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
                ($Message -like "*Property 'ServiceNow.$_' not found*")
            }

            $LASTEXITCODE | Should -Be 1
        }
    }
} -Tag test
Describe 'create no ticket' {
    BeforeAll {
        Mock Invoke-Sqlcmd -ParameterFilter {
            $Query -like "*FROM $SQLTableAdInconsistencies*"
        } -MockWith {
            [PSCustomObject]@{
                SamAccountName = 'PC1'
            }
        }

        $testNewParams = $testParams.Clone()
        $testNewParams.Data = @(
            [PSCustomObject]@{
                SamAccountName = 'PC1'
            }
        )

        .$testScript @testNewParams
    }
    It 'when a ticket was already created and it is still open' {
        Should -Not -Invoke New-CherwellTicketHC -Scope Describe
    }
}
Describe 'create a new ticket' {
    BeforeAll {
        Mock Invoke-Sqlcmd -ParameterFilter {
            $Query -like "*FROM $SQLTableTicketsDefaults*"
        } -MockWith {
            [PSCustomObject]@{
                Requester          = 'jack'
                SubmittedBy        = 'mike'
                ServiceCountryCode = 'BNL'
            }
        }
        Mock Invoke-Sqlcmd -ParameterFilter {
            $Query -like "*FROM $SQLTableAdInconsistencies*"
        } -MockWith {
            [PSCustomObject]@{
                SamAccountName = 'PC1'
            }
        }
        $testNewParams = $testParams.Clone()
        $testNewParams.Data = @(
            [PSCustomObject]@{
                SamAccountName = 'PC2'
            }
        )
    }
    It 'when no ticket was created before or it was closed' {
        .$testScript @testNewParams

        Should -Invoke New-CherwellTicketHC -Times 1 -Exactly
    }
    Context 'with properties from' {
        It 'SQL table ticketsDefaults when there are none in the .json file' {
            $testNewParams.TicketFields = $null

            .$testScript @testNewParams

            Should -Invoke New-CherwellTicketHC -Times 1 -Exactly -ParameterFilter {
                ($KeyValuePair.RequesterSamAccountName -eq 'jack') -and
                ($KeyValuePair.SubmittedBySamAccountName -eq 'mike') -and
                ($KeyValuePair.ServiceCountryCode -eq 'BNL')
            }
        }
        It 'the .json file, they overwrite the SQL ticketsDefaults' {
            $testNewParams.TicketFields = [PSCustomObject]@{
                RequesterSamAccountName   = 'picard'
                SubmittedBySamAccountName = 'kirk'
                ServiceCountryCode        = $null
            }

            .$testScript @testNewParams

            Should -Invoke New-CherwellTicketHC -Times 1 -Exactly -ParameterFilter {
                ($KeyValuePair.RequesterSamAccountName -eq 'picard') -and
                ($KeyValuePair.SubmittedBySamAccountName -eq 'kirk') -and
                ($KeyValuePair.ServiceCountryCode -eq 'BNL')
            }
        }
    }
}