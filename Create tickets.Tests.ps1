#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        ScriptName        = 'Test'
        Environment       = 'Test'
        SQLDatabase       = 'Test'
        TopicName         = 'Computer - Inactive'
        DistinguishedName = 'bob'
    }

    Mock Get-ADGroup
    Mock New-CherwellTicketHC { 1 }
    Mock Save-TicketInSqlHC
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
        'ScriptName', 'TopicName', 'DistinguishedName', 'Environment'
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