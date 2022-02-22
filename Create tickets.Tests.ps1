#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        Environment       = 'Test'
        SQLDatabase       = 'Test'
        TopicName         = 'Computer - Inactive'
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
            Requester          = 'testScriptAccount'
            ServiceCountryCode = 'BNL'
        }
    }
}
Describe 'the mandatory parameters are' {
    It "<_>" -ForEach @(
        'TopicName', 'DistinguishedName'
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
    It 'when no ticket was created before or it was closed' {
        

        $testNewParams = $testParams.Clone()
        $testNewParams.DistinguishedName = 'b'

        .$testScript @testNewParams

        Should -Invoke New-CherwellTicketHC -Times 1 -Exactly
    }
}