#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $TestParams = @{
        ScriptName       = 'Test'
        Environment      = 'Test'
        SQLDatabase      = 'Test'
        TopicName        = 'Computer - Inactive'
        TopicDescription = "'LastLogonDate' over x days"
        Data             = @(
            [PSCustomObject]@{
                Name              = 'bob'
                DistinguishedName = 'DC=bob,CN=contoso,CN=net'
            }
        )
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
        'ScriptName', 'Environment', 'TopicName', 'TopicDescription', 'Data'
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
    It 'the .json file contains unknown ticket fields' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Data = @(
            [PSCustomObject]@{
                Name              = 'jack'
                DistinguishedName = 'DC=jack,CN=contoso,CN=net'
            }
        )
        $testNewParams.TicketFields = [PSCustomObject]@{
            incorrectFieldName = 'x'
        }
        
        .$testScript @testNewParams
        
        Should -Invoke Write-EventLog -Times 1 -Exactly -ParameterFilter {
            ($EntryType -eq 'Error') -and
            ($Message -like "*Field name 'incorrectFieldName' not found*")
        }

        Should -Not -Invoke New-CherwellTicketHC
    }
}
Describe 'create no ticket' {
    BeforeAll {
        Mock Invoke-Sqlcmd2 -ParameterFilter {
            $Query -like "*FROM $SQLTableAdInconsistencies*"
        } -MockWith {
            [PSCustomObject]@{
                DistinguishedName = 'DC=jack,CN=contoso,CN=net'
            }
        }

        $testNewParams = $testParams.Clone()
        $testNewParams.Data = @(
            [PSCustomObject]@{
                Name              = 'jack'
                DistinguishedName = 'DC=jack,CN=contoso,CN=net'
            }
        )

        .$testScript @testNewParams
    }
    It 'when a ticket was already created and it is still open' {
        Should -Not -Invoke New-CherwellTicketHC -Scope Describe
    }
    It 'and register this in the event log' {
        Should -Invoke Write-EventLog -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($EntryType -ne 'Error') -and
            ($Message -like 'No ticket created')
        }
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
                DistinguishedName = 'DC=chuck,CN=contoso,CN=net'
            }
        }
    }
    It 'when no ticket was created before or it was closed' {
        $testNewParams = $testParams.Clone()
        $testNewParams.Data = @(
            [PSCustomObject]@{
                Name              = 'jack'
                DistinguishedName = 'DC=jack,CN=contoso,CN=net'
            }
        )

        .$testScript @testNewParams

        Should -Invoke New-CherwellTicketHC -Times 1 -Exactly
    }
    Context 'with properties from' {
        It 'SQL table ticketsDefaults when there are none in the .json file' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Data = @(
                [PSCustomObject]@{
                    Name              = 'jack'
                    DistinguishedName = 'DC=jack,CN=contoso,CN=net'
                }
            )
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
            $testNewParams.Data = @(
                [PSCustomObject]@{
                    Name              = 'jack'
                    DistinguishedName = 'DC=jack,CN=contoso,CN=net'
                }
            )
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