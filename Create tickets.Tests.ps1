#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        ScriptName        = 'Test'
        Environment       = 'Test'
        SQLDatabase       = 'Test'
        TopicName         = 'Computer - Inactive'
        TopicDescription  = "'LastLogonDate' over x days"
        DistinguishedName = 'bob'
    }

    Mock New-CherwellTicketHC { 1 }
    Mock Save-TicketInSqlHC
    Mock Send-MailHC
    Mock Write-EventLog
    Mock Invoke-Sqlcmd2
    Mock Invoke-Sqlcmd2 -ParameterFilter {
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
    It "<_>" -ForEach @(
        'ScriptName', 'Environment', 'TopicName', 'TopicDescription', 'DistinguishedName'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'an error is thrown when' {
    It 'no ticket default values are found in SQL' {
        Mock Invoke-Sqlcmd2 -ParameterFilter {
            $Query -like "*FROM $SQLTableTicketsDefaults*"
        }

        .$testScript @TestParams

        Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
            ($EntryType -eq 'Error') -and
            ($Message -like "*No ticket default values found*")
        }
    }
}
Describe 'create no ticket when' {
    It 'a ticket was already created and it is still open' {
        Mock Invoke-Sqlcmd2 -ParameterFilter {
            $Query -like "*FROM $SQLTableAdInconsistencies*"
        } -MockWith {
            [PSCustomObject]@{
                DistinguishedName = 'a'
            }
        }

        $testNewParams = $testParams.Clone()
        $testNewParams.DistinguishedName = 'a'

        .$testScript @testNewParams

        Should -Invoke New-CherwellTicketHC -Times 0 -Exactly
    }
}
Describe 'create a new ticket' {
    BeforeAll {
        Mock Invoke-Sqlcmd2 -ParameterFilter {
            $Query -like "*FROM $SQLTableTicketsDefaults*"
        } -MockWith {
            [PSCustomObject]@{
                Requester          = 'jack'
                SubmittedBy        = 'mike'
                ServiceCountryCode = 'BNL'
            }
        }
        Mock Invoke-Sqlcmd2 -ParameterFilter {
            $Query -like "*FROM $SQLTableAdInconsistencies*"
        } -MockWith {
            [PSCustomObject]@{
                DistinguishedName = 'a'
            }
        }
    }
    It 'when no ticket was created before or it was closed' {
        $testNewParams = $testParams.Clone()
        $testNewParams.DistinguishedName = 'b'

        .$testScript @testNewParams

        Should -Invoke New-CherwellTicketHC -Times 1 -Exactly
    }
    Context 'with properties from' {
        It 'SQL table ticketsDefaults when there are none in the .json file' {
            $testNewParams = $testParams.Clone()
            $testNewParams.DistinguishedName = 'b'
            $testNewParams.TicketFields = $null
            
            .$testScript @testNewParams
            
            Should -Invoke New-CherwellTicketHC -Times 1 -Exactly -ParameterFilter {
                ($KeyValuePair.RequesterSamAccountName -eq 'jack') -and
                ($KeyValuePair.SubmittedBySamAccountName -eq 'mike') -and
                ($KeyValuePair.ServiceCountryCode -eq 'BNL')
            }
        }
        It 'the .json file, they overwrite the SQL ticketsDefaults' {
            $testNewParams = $testParams.Clone()
            $testNewParams.DistinguishedName = 'b'
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
        } -Tag test
    }
}