<# 
    .SYNOPSIS
        Create tickets when needed
        
    .DESCRIPTION
        Check if a ticket is already created for a specific topic and
        distinguished name. Only create a new ticket when there is no ticket
        yet in the database or when the previous ticket has been closed.

    .EXAMPLE
        $params = @{
            ScriptName        = 'AD Inconsistencies (BNL)'
            Environment       = 'Stage'
            SQLDatabase       = 'Powershell TEST'
            TopicName         = 'Computer - Inactive'
            TopicDescription  = 'LastLogonDate over 40 days'
            Data = @(
                [PSCustomObject]@{
                    Name                  = 'Bob Lee Swagger'
                    SamAccountName        = 'swagger'
                    DisplayName           = 'Bob Lee Swagger'
                    AccountExpirationDate = (Get-Date).AddYears(-2)
                    EmployeeType          = 'Sniper'
                    ManagerDisplayName    = ''
                    OU                    = 'contoso.com\USA\Users'
                },
                [PSCustomObject]@{
                    Name                  = 'Chuck Norris'
                    SamAccountName        = 'norris'
                    DisplayName           = 'Chuck Norris'
                    AccountExpirationDate = (Get-Date).AddYears(-3)
                    EmployeeType          = 'Actor'
                    ManagerDisplayName    = ''
                    OU                    = 'contoso.com\USA\Users\actors\hero'
                }
            ) 
            TicketFields      = (
                [PSCustomObject]@{
                    ShortDescription          = 'AD Inconsistency: Vendor account not expiring'
                    Description               = 'Please set the expiration date within 1 year'
                    SubmittedBySamAccountName = 'bob'
                }
            )
        }
        & $script @params

        Create tickets for Bob Lee Swagger and Chuck Norris in case there aren't
        any tickets created yet for them with the issue 'Computer - Inactive'.
#>
[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$Environment,
    [Parameter(Mandatory)]
    [String]$TopicName,
    [Parameter(Mandatory)]
    [String]$TopicDescription,
    [Parameter(Mandatory)]
    [PSCustomObject[]]$Data,
    [PSCustomObject]$TicketFields,
    [DateTime]$TicketRequestedDate = (Get-Date),

    [String]$SQLServerInstance = 'GRPSDFRAN0049',
    [String]$SQLDatabase = 'PowerShell',
    [String]$SQLTableTicketsDefaults = 'TicketsDefaults',
    [String]$SQLTableAdInconsistencies = 'AdInconsistencies',

    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        $M = "TopicName '$TopicName' TopicDescription '$TopicDescription'"
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M

        $SQLParams = @{
            ServerInstance    = $SQLServerInstance
            Database          = $SQLDatabase
            QueryTimeout      = '1000'
            ConnectionTimeout = '20'
            ErrorAction       = 'Stop'
        }

        #region Get SQL ticket default values
        $SQLTicketDefaults = Invoke-Sqlcmd2 -As PSObject @SQLParams -Query "
            SELECT *
            FROM $SQLTableTicketsDefaults
            WHERE ScriptName = '$ScriptName'"

        if (-not $SQLTicketDefaults) {
            throw "No ticket default values found in SQL table '$SQLTableTicketsDefaults' for ScriptName '$ScriptName'"
        }
        #endregion

        #region Overwrite with json file default values
        $KeyValuePair = @{
            ServiceCountryCode        = $SQLTicketDefaults.ServiceCountryCode
            RequesterSamAccountName   = $SQLTicketDefaults.Requester
            SubmittedBySamAccountName = $SQLTicketDefaults.SubmittedBy
            OwnedByTeam               = $SQLTicketDefaults.OwnedByTeam
            OwnedBySamAccountName     = $SQLTicketDefaults.OwnedBy
            ShortDescription          = $SQLTicketDefaults.ShortDescription
            Description               = 'Please correct the following:'
            Service                   = $SQLTicketDefaults.Service
            Category                  = $SQLTicketDefaults.Category
            SubCategory               = $SQLTicketDefaults.SubCategory
            Source                    = $SQLTicketDefaults.Source
            IncidentType              = $SQLTicketDefaults.IncidentType
            Priority                  = $SQLTicketDefaults.Priority
        }

        foreach (
            $field in 
            $TicketFields.PSObject.Properties | Where-Object { $_.Value }
        ) {
            if (-not $KeyValuePair.containsKey($field.Name)) {
                throw "Field name '$($field.Name)' not found in Cherwell, valid fields are '$($KeyValuePair.Keys)'"
            }
            $KeyValuePair[$field.Name] = $field.Value
        }
        #endregion

        #region Get open tickets
        $openTickets = Invoke-Sqlcmd2 @SQLParams -As PSObject -Query "
            SELECT SamAccountName
            FROM $SQLTableAdInconsistencies
            WHERE 
                TopicName = '$TopicName' AND 
                TicketRequestedDate IS NOT NULL AND 
                TicketCloseDate IS NULL"

        $M = "Found $($openTickets.count) open tickets"
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit
    }
}

Process {
    Try {
        $PSCode = $null
        $ticketDescription = $KeyValuePair.Description
        Foreach (
            $D in 
            $Data | Where-Object { 
                $openTickets.SamAccountName -notContains $_.SamAccountName 
            }
        ) {
            Try {
                #region Create ticket
                $PSCode = New-PSCodeHC $SQLTicketDefaults.ServiceCountryCode

                $KeyValuePair.Description = $ticketDescription + "
                <br><br>
                $TopicDescription
                <br><br>
                <table style=`"border:none`">
                $($D.PSObject.Properties | ForEach-Object {
                    '<tr style="border:none;text-align:left;">
                        <th style="border:none;width:62px;color:lightGray;">{0}</th>
                        <td style="border:none;"><b>{1}</b></td>
                    </tr>' -f $_.Name, $_.Value
                })
                </table>"
                Remove-EmptyParamsHC -Name $KeyValuePair

                $TicketParams = @{
                    Environment  = $Environment
                    KeyValuePair = $KeyValuePair
                    ErrorAction  = 'Stop'
                }
                $TicketNr = New-CherwellTicketHC @TicketParams

                Write-EventLog @EventOutParams -Message "Created ticket '$TicketNr' for '$($D.SamAccountName)' with short description '$($KeyValuePair.ShortDescription)'"
                #endregion

                #region Save details in SQL
                $SaveTicketParams = @{
                    Database     = $SQLDatabase
                    ScriptName   = $ScriptName
                    KeyValuePair = $KeyValuePair
                    PSCode       = $PSCode
                    TicketNr     = $TicketNr
                }
                Save-TicketInSqlHC @SaveTicketParams
                
                Invoke-Sqlcmd2 @SQLParams -Query "
                    INSERT INTO $SQLTableAdInconsistencies
                    (PSCode, SamAccountName, TopicName, 
                    TicketRequestedDate, TicketNr)
                    VALUES(
                    '$PSCode', $(FSQL $D.SamAccountName), 
                    $(FSQL $TopicName), $(FSQL $TicketRequestedDate), '$TicketNr')"
                #endregion
            }
            Catch {
                throw "Failed creating a ticket for TopicName '$TopicName' SamAccountName '$($_.Name)': $_"
            }
        }

        if (-not $PSCode) {
            $M = 'No ticket created'
            Write-Verbose $M
            Write-EventLog @EventVerboseParams -Message $M
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        $Error.RemoveAt(0); $Error.RemoveAt(0)
    }
}