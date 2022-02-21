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
    [String]$TopicName,
    [Parameter(Mandatory)]
    [String[]]$DistinguishedName,
    [PSCustomObject]$TicketFields,
    [String]$Environment,


    [String]$SQLServerInstance = 'GRPSDFRAN0049',
    [String]$SQLDatabase = 'PowerShell',
    [String]$SQLTableTicketsDefaults = 'TicketsDefaults',
    [String]$SQLTableAdInconsistencies = 'AdInconsistencies',

    [String]$ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'
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

        $SQLTicketDefaults = Invoke-Sqlcmd2 -As PSObject @SQLParams -Query "
            SELECT *
            FROM $SQLTableTicketsDefaults
            WHERE ScriptName = '$ScriptName'"

        if (-not $SQLTicketDefaults) {
            throw "No ticket default values found in SQL table '$SQLTableTicketsDefaults' for ScriptName '$ScriptName'"
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
