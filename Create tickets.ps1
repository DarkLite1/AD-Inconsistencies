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
    [String]$TopicName,
    [Parameter(Mandatory)]
    [String[]]$DistinguishedName,
    [PSCustomObject]$TicketFields,
    [String]$Environment = 'Prod',

    [String]$SQLServerInstance = 'GRPSDFRAN0049',
    [String]$SQLDatabase = 'PowerShell',
    [String]$SQLTableTicketsDefaults = 'TicketsDefaults',
    [String]$SQLTableAdInconsistencies = 'AdInconsistencies',

    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        $SQLParams = @{
            ServerInstance    = $SQLServerInstance
            Database          = $SQLDatabase
            QueryTimeout      = '1000'
            ConnectionTimeout = '20'
            ErrorAction       = 'Stop'
        }

        #region Get ticket default values
        $SQLTicketDefaults = Invoke-Sqlcmd2 -As PSObject @SQLParams -Query "
            SELECT *
            FROM $SQLTableTicketsDefaults
            WHERE ScriptName = '$ScriptName'"

        if (-not $SQLTicketDefaults) {
            throw "No ticket default values found in SQL table '$SQLTableTicketsDefaults' for ScriptName '$ScriptName'"
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
    }
}

Process {
    Try {
        $openTickets = Invoke-Sqlcmd2 @SQLParams -As PSObject -Query "
            SELECT DistinguishedName
            FROM $SQLTableAdInconsistencies
            WHERE 
                TopicName = '$TopicName' AND 
                TicketRequestedDate IS NOT NULL AND 
                TicketCloseDate IS NULL"

        Foreach (
            $Name in 
            $DistinguishedName | 
            Where-Object { $openTickets.DistinguishedName -notContains $_ }
        ) {
            Try {
                $M = "Create ticket for TopicName '$TopicName' DistinguishedName '$Name'"
                Write-Verbose $M
                Write-EventLog @EventVerboseParams -Message $M

                $PSCode = New-PSCodeHC $SQLTicketDefaults.ServiceCountryCode

                $Description = "
                    Please add the user '$PlaceHolderAccount'"

                $KeyValuePair = @{
                    ServiceCountryCode        = $SQLTicketDefaults.ServiceCountryCode
                    RequesterSamAccountName   = $SQLTicketDefaults.Requester
                    SubmittedBySamAccountName = $SQLTicketDefaults.SubmittedBy
                    OwnedByTeam               = $SQLTicketDefaults.OwnedByTeam
                    OwnedBySamAccountName     = $SQLTicketDefaults.OwnedBy
                    ShortDescription          = $SQLTicketDefaults.ShortDescription
                    Description               = $Description
                    Service                   = $SQLTicketDefaults.Service
                    Category                  = $SQLTicketDefaults.Category
                    SubCategory               = $SQLTicketDefaults.SubCategory
                    Source                    = $SQLTicketDefaults.Source
                    IncidentType              = $SQLTicketDefaults.IncidentType
                    Priority                  = $SQLTicketDefaults.Priority
                }
                Remove-EmptyParamsHC -Name $KeyValuePair

                $TicketParams = @{
                    Environment  = $Environment
                    KeyValuePair = $KeyValuePair
                    ErrorAction  = 'Stop'
                }
                $TicketNr = New-CherwellTicketHC @TicketParams

                Write-EventLog @EventOutParams -Message "Created ticket '$TicketNr'"

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
            }
            Catch {
                throw "Failed creating a ticket for TopicName '$TopicName' DistinguishedName '$Name': $_"
            }
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
    }
}