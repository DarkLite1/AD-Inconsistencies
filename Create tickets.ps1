#Requires -Version 7
#Requires -Modules Toolbox.HTML, Toolbox.EventLog
#Requires -Modules SqlServer, ServiceNow

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
            PSCustomObject    = @{
                CredentialsFilePath = 'C:\PasswordsServiceNow.json'
                Environment         = 'Test'
            }
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
param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [PSCustomObject]$ServiceNow,
    [Parameter(Mandatory)]
    [String]$TopicName,
    [Parameter(Mandatory)]
    [String]$TopicDescription,
    [Parameter(Mandatory)]
    [PSCustomObject[]]$Data,
    [PSCustomObject]$TicketFields,
    [DateTime]$TicketRequestedDate = (Get-Date),
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

begin {
    function New-ServiceNowSessionHC {
        param (
            [parameter(Mandatory)]
            [String]$Uri,
            [parameter(Mandatory)]
            [String]$UserName,
            [parameter(Mandatory)]
            [String]$Password,
            [parameter(Mandatory)]
            [String]$ClientId,
            [parameter(Mandatory)]
            [String]$ClientSecret   
        )
        try {
            $userCred = New-Object System.Management.Automation.PSCredential(
                $UserName, 
                ($Password | ConvertTo-SecureString -AsPlainText -Force)
            )
        
            $clientCred = New-Object System.Management.Automation.PSCredential(
                $ClientId, 
                ($ClientSecret | ConvertTo-SecureString -AsPlainText -Force)
            )
            
            $params = @{
                Url              = $Uri
                Credential       = $userCred
                ClientCredential = $clientCred
            }
            New-ServiceNowSession @params
        }
        catch {
            $errorMessage = $_; $Error.RemoveAt(0)
            throw "Failed to create a ServiceNow session: $errorMessage"
        }           
    }

    try {
        $M = "TopicName '$TopicName' TopicDescription '$TopicDescription'"
        Write-Verbose $M
        Write-EventLog @EventVerboseParams -Message $M

        #region Create ServiceNow session
        if (-not $ServiceNowSession) {
            #region Test ServiceNow parameters
            @(
                'CredentialsFilePath', 'Environment', 'TicketFields'
            ).where(
                { -not $ServiceNow.$_ }
            ).foreach(
                { throw "Property 'ServiceNow.$_' not found" }
            )

            try {
                $serviceNowJsonFileContent = Get-Content $ServiceNow.CredentialsFilePath -Raw -EA Stop | ConvertFrom-Json
            }
            catch {
                throw "Failed to import the ServiceNow environment file '$($ServiceNow.CredentialsFilePath)': $_"
            }

            $serviceNowEnvironment = $serviceNowJsonFileContent.($ServiceNow.Environment)

            if (-not $serviceNowEnvironment) {
                throw "Failed to find environment '$($ServiceNow.Environment)' in the ServiceNow environment file '$($ServiceNow.CredentialsFilePath)'"
            }

            @(
                'Uri', 'UserName', 'Password', 'ClientId', 'ClientSecret'
            ).where(
                { -not $serviceNowEnvironment.$_ }
            ).foreach(
                { 
                    throw "Property '$_' not found for environment '$($ServiceNow.Environment)' in file '$($ServiceNow.CredentialsFilePath)'"
                }
            )
            #endregion

            #region Create global variable $ServiceNowSession
            $params = @{
                Uri          = $serviceNowEnvironment.Uri
                UserName     = $serviceNowEnvironment.UserName
                Password     = $serviceNowEnvironment.Password
                ClientId     = $serviceNowEnvironment.ClientId
                ClientSecret = $serviceNowEnvironment.ClientSecret
            }
            New-ServiceNowSessionHC @params
            #endregion
        }
        #endregion

        #region Get SQL ticket default values
        $SQLTicketDefaults = Invoke-Sqlcmd @SQLParams -Query "
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
                throw "Field name '$($field.Name)' not valid, valid fields are '$($KeyValuePair.Keys)'"
            }
            $KeyValuePair[$field.Name] = $field.Value
        }
        #endregion

        #region Get open tickets
        $openTickets = Invoke-Sqlcmd @SQLParams -Query "
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
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message "FAiled creating tickets: $_" -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; exit 1
    }
}

process {
    try {
        $ticketDescription = $KeyValuePair.Description

        foreach (
            $D in
            $Data | Where-Object {
                $openTickets.SamAccountName -notcontains $_.SamAccountName
            }
        ) {
            try {
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

                Invoke-Sqlcmd @SQLParams -Query "
                    INSERT INTO $SQLTableAdInconsistencies
                    (PSCode, SamAccountName, TopicName,
                    TicketRequestedDate, TicketNr)
                    VALUES(
                    '$PSCode', $(FSQL $D.SamAccountName),
                    $(FSQL $TopicName), $(FSQL $TicketRequestedDate), '$TicketNr')"
                #endregion
            }
            catch {
                throw "Failed creating a ticket for TopicName '$TopicName' SamAccountName '$($D.SamAccountName)': $_"
            }
        }
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message "FAiled creating tickets: $_" -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; exit 1
    }
}