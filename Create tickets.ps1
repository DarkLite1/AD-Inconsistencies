<# 
    .SYNOPSIS
        Create tickets when needed
        
    .DESCRIPTION
        Check if a ticket is already created for a specific topic and
        distinguished name. Only create a new ticket when there is no ticket
        yet in the database or when the previous ticket has been closed.
#>
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
    [String[]]$DistinguishedName,
    [PSCustomObject]$TicketFields,

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
            SELECT DistinguishedName
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
        
        Foreach (
            $Name in 
            $DistinguishedName | 
            Where-Object { $openTickets.DistinguishedName -notContains $_ }
        ) {
            Try {
                #region Create ticket
                $M = "Create ticket for '$Name'"
                Write-Verbose $M
                Write-EventLog @EventVerboseParams -Message $M

                $PSCode = New-PSCodeHC $SQLTicketDefaults.ServiceCountryCode

                $KeyValuePair.Description = $KeyValuePair.Description += "

                - DistinguishedName: $Name

                - Description: $TopicDescription"
                Remove-EmptyParamsHC -Name $KeyValuePair

                $TicketParams = @{
                    Environment  = $Environment
                    KeyValuePair = $KeyValuePair
                    ErrorAction  = 'Stop'
                }
                $TicketNr = New-CherwellTicketHC @TicketParams

                Write-EventLog @EventOutParams -Message "Created ticket '$TicketNr'"
                #endregion

                #region Save details in SQL
                $SaveTicketParams = @{
                    KeyValuePair = $KeyValuePair
                    PSCode       = $PSCode
                    TicketNr     = $TicketNr
                    ScriptName   = $ScriptName
                }
                Save-TicketInSqlHC @SaveTicketParams
                
                Invoke-Sqlcmd2 @SQLParams -Query "
                    INSERT INTO $SQLTableAdInconsistencies
                    (PSCode, DistinguishedName, TopicName, 
                    TicketRequestedDate, TicketNr)
                    VALUES('$PSCode', '$Name', '$TopicName', 
                    $(FSQL $Now), '$TicketNr')"
                #endregion
            }
            Catch {
                throw "Failed creating a ticket for TopicName '$TopicName' DistinguishedName '$Name': $_"
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