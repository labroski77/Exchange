
<#
    .Synopsis
    Exchange-ConnectorManagement.ps1
    .DESCRIPTION

    PowerShell script to Exchange Receive Connectors.
    

    .EXAMPLE
    .\Exchange-ConnectorManagement.ps1

    ==========================================================================
    Created by: Carlos Annes
    09/2023
    ==========================================================================
#>


function Export-ReceiveConnectors { 

<#
.SYNOPSIS
Exports Receive Connectors to a CSV file based on user choice.
.DESCRIPTION
This function exports either all the Receive Connectors or a single Receive Connector to a CSV file based on user input.

#>

    # Ask the user for export type
    $exportType = Read-Host "Do you want to export (A)ll connectors or a (S)ingle connector? (A/S)"

    if ($exportType -eq "A") {
        # Initialize an array to hold all connectors' information 
        $allConnectors = @() 

        try {  
            # Get all Exchange servers
            $exchangeServers = Get-ExchangeServer | Select-Object -ExpandProperty Name

            # Initialize an array to hold Receive Connector details
            $allReceiveConnectors = @()

            # Loop through each Exchange server to get Receive Connectors
            foreach ($server in $exchangeServers) {
                $receiveConnectors = Get-ReceiveConnector -Server $server
                $allReceiveConnectors += $receiveConnectors
            }

            # Export details to CSV
            $allReceiveConnectors  | Select-Object Identity,AuthMechanism,PermissionGroups,{$_.Bindings}, {$_.RemoteIPRanges} | Sort-object Identity | Export-Csv -Path ".\AllReceiveConnectors.csv" -NoTypeInformation

            Write-Host "Exported all Receive Connectors to .\AllReceiveConnectors.csv"
        }
        catch {
            Write-Error $_.Exception.Message
        }
    } elseif ($exportType -eq "S") {
        try {
            # Ask for the connector identity
            $connectorIdentity = Read-Host "Enter the connector Identity"

            # Get the specific Receive Connector
            $singleReceiveConnector = Get-ReceiveConnector -Identity $connectorIdentity

           # Replace the backslash with an underscore for the file name
        $safeConnectorIdentity = $connectorIdentity -replace '\\', '_'

        # Export details to CSV
        $singleReceiveConnector | Select-Object Identity,AuthMechanism,PermissionGroups,{$_.Bindings}, {$_.RemoteIPRanges} | Sort-object Identity | Export-Csv -Path ".\${safeConnectorIdentity}_ReceiveConnector.csv" -NoTypeInformation

        Write-Host "Exported Receive Connector $connectorIdentity to .\${safeConnectorIdentity}_ReceiveConnector.csv"
    }
        catch {
            Write-Error $_.Exception.Message
        }
    } else {
        Write-Host "Option is not valid, please provide a valid option."
    }
}



# Function 2: Show and Export AD ExtendedRights 

Function Export-ExtendedRights { 

<#

.SYNOPSIS
Exports the Extended Rights of a specified Receive Connector in Microsoft Exchange to a CSV file.

.DESCRIPTION
The Export-ExtendedRights function allows the user to export the Extended Rights settings of a specified Receive Connector. The function prompts the user for the following:

Receive Connector: The name of the Receive Connector whose Extended Rights settings are to be exported.
After gathering this information, the function retrieves the Extended Rights settings using the Get-ADPermission cmdlet and filters out inherited and denied permissions. The function then displays these settings in a table format and exports them to a CSV file named ${receiveConnector}_ExtendedRights.csv in the current directory.

#>
    
    try {
        $receiveConnector = Read-Host "Enter the connector identity.Lke (server\connectorName) without quotations "
        
        $extendedRightsData = Get-ReceiveConnector -Identity $receiveConnector | Get-ADPermission | Where-Object {($_.Deny -eq $false) -and ($_.IsInherited -eq $false)}
        
        # Handle multi-value data for ExtendedRights
        $processedData = $extendedRightsData | Select-Object @{Name='ReceiveConnector'; Expression={$_.Identity}}, 
                                                     @{Name='User'; Expression={$_.User}}, 
                                                     @{Name='ExtendedRights'; Expression={($_.ExtendedRights -join ',')}}
        
        # Display in table format
        $processedData | Format-Table ReceiveConnector, User, ExtendedRights
        
        
        # Replace the backslash with an underscore for the file name
        $safeConnectorIdentity = $receiveConnector -replace '\\', '_'

        # Export to CSV
        $processedData | Export-Csv -Path ".\${safeConnectorIdentity}_ExtendedRights.csv" -NoTypeInformation
        
        Write-Host "Exported Extended Rights to .\${safeConnectorIdentity}_ExtendedRights.csv"
    }
    catch {
        Write-Error $_.Exception.Message
    }
}





# Function 3 Change Receive Connector Settings


Function Change-ReceiveConnectorSettings {

<#
.SYNOPSIS
Modifies the settings of a specified Receive Connector in Microsoft Exchange.

.DESCRIPTION
The Change-ReceiveConnectorSettings function allows the user to modify various settings of a specified Receive Connector. The function prompts the user for the following:

Receive Connector: The name of the Receive Connector to modify.
Remote IP Ranges: New Remote IP Ranges, separated by commas. Leave empty to keep the current setting.
Bindings: New Bindings in the format IP:Port. Leave empty to keep the current setting.
Authentication Mechanisms: New Authentication mechanisms, separated by commas. Leave empty to keep the current setting.
Permissions Group: New Permissions Group, separated by commas. Leave empty to keep the current setting.
After gathering this information, the function updates the Receive Connector with the new settings provided by the user. If any field is left empty, the current setting for that field will be retained.

#>

try {

    $receiveConnector = Read-Host "Enter the connector identity.Lke (server\connectorName) without quotations "
    #$connectorIdentity = Read-Host "Enter the Identity of the connector"

    # Change Remote IP Range
    $remoteIPs = Read-Host "Enter the new Remote IP Ranges (separated by commas, leave empty to keep current)"
    if ($remoteIPs -ne '') {
        $remoteIPArray = $remoteIPs -split ','
        Set-ReceiveConnector -Identity $receiveConnector -RemoteIPRanges $remoteIPArray
    }

    # Change Bindings
    $bindings = Read-Host "Enter the new Bindings (Format: IP:Port, leave empty to keep current)"
    if ($bindings -ne '') {
        Set-ReceiveConnector -Identity $receiveConnector -Bindings $bindings
    }

    # Change Authentication
    $authentication = Read-Host "Enter the new Authentication mechanisms (separated by commas, leave empty to keep current)"
    if ($authentication -ne '') {
        $authenticationArray = $authentication -split ','
        Set-ReceiveConnector -Identity $receiveConnector -AuthMechanism $authenticationArray
    }

    # Change Permissions Group
    $permissionsGroup = Read-Host "Enter the new Permissions Group (separated by commas, leave empty to keep current)"
    if ($permissionsGroup -ne '') {
        $permissionsGroupArray = $permissionsGroup -split ','
        Set-ReceiveConnector -Identity $receiveConnector -PermissionGroups $permissionsGroupArray
    }

    Write-Host "Changed settings for $receiveConnector."

    return

}

catch {
        Write-Error $_.Exception.Message
    }

}


# Function 4: Change Connector ExtendedRights

Function Change-ExtendedRights { 

<#
.SYNOPSIS
Modifies the Extended Rights of a specified Receive Connector in Microsoft Exchange.

.DESCRIPTION
The Change-ExtendedRights function allows the user to add or remove Extended Rights from a specified Receive Connector. The function prompts the user for the following:

Action: Whether to add or remove permissions. The user can enter "A" for Add or "R" for Remove.
Receive Connector: The name of the Receive Connector to modify.
User: The user whose permissions will be modified.
Extended Rights: The Extended Rights to add or remove, separated by commas if multiple.
After gathering this information, the function performs the specified action (add or remove) for each Extended Right. It then restarts the MSExchangeTransport service to apply the changes and displays the updated Extended Rights for the specified Receive Connector.


#>

    $action = Read-Host "Do you want to Add or Remove permissions? ((A)dd/(R)emove)" 

try {

    $receiveConnector = Read-Host "Enter the connector identity.Like (server\connectorName) without quotations "
    #$connectorIdentity = Read-Host "Enter the Identity of the connector"

    $user = Read-Host "Enter the User quotations "" " 

    $extendedRights = Read-Host "Enter the ExtendedRights without quotations (separate multiple values with commas)" 

    $extendedRightsArray = $extendedRights -split ',' 

  

    if ($action -eq "A") { 

        foreach ($right in $extendedRightsArray) { 

            Get-ReceiveConnector -Identity $receiveConnector | Add-ADPermission -User $user -ExtendedRights $right.Trim() 

        } 

    } elseif ($action -eq "R") { 

        foreach ($right in $extendedRightsArray) { 

            Get-ReceiveConnector -Identity $receiveConnector |Remove-ADPermission -User $user -ExtendedRights $right.Trim() -Confirm:$false 

        } 

        else {
        Write-Host "Option is not valid, please provide a valid option."

        }

    } 

  

    # Restart the Transport Service 

    Restart-Service MSExchangeTransport 

  

    # Show the updated ExtendedRights 

      Get-ReceiveConnector -Identity $receiveConnector | Get-ADPermission | Where-Object {($_.Deny -eq $false) -and ($_.IsInherited -eq $false)} | Format-Table ReceiveConnector, User, ExtendedRights 

return
}

catch {
        Write-Error $_.Exception.Message
    }

} 


# Function 4: Copy Receive Connector

Function Copy-ReceiveConnector {

<#
.SYNOPSIS
This function copies a Receive Connector from a source Exchange server to a target Exchange server.

.DESCRIPTION
The function performs the following tasks:
- Prompts the user for the source and target servers, and the Domain Controller.
- Prompts the user for the name of the Receive Connector to be copied from the source server.
- Checks if the Receive Connector exists on the source server.
- Handles special cases where the AuthMechanism is set to 'ExternalAuthoritative'.
- Prompts the user if any settings (Connector Name, Remote IPs, Bindings, Authentication, Permission Groups) need to be changed.
- Creates a new Receive Connector on the target server with either the changed settings or the original settings.
- Copies the Extended Rights from the source Receive Connector to the new Receive Connector on the target server.


.NOTES
- The function uses the Get-ReceiveConnector, New-ReceiveConnector, and Set-ReceiveConnector cmdlets for managing Receive Connectors.
- It also uses Get-ADPermission and Add-ADPermission for managing Extended Rights.
- Error handling is implemented to catch exceptions and display an error message.

#>
 

Import-Module -Name ActiveDirectory 

try {

# Specify Source and Target Servers
$sourceServer = Read-Host "Enter the name of the source server"
$targetServer = Read-Host "Enter the name of the target server"
$DomainControler = Read-Host "Enter the name of Domain Controler"

# Ask for the source connector name
$connectorName = Read-Host "Enter the name of the source connector to be copied"

# Fetch the source connector
$connector = Get-ReceiveConnector -Server $sourceServer -DomainController $DomainController| Where-Object {$_.Name -eq $connectorName}


if ($null -eq $connector) {
            Write-Error "Receive Connector '$connectorName' not found on source server '$sourceServer'."
            return
        }

        $FixAuthMechanism = $false
        if(([string]$connector.AuthMechanism).Split(',').Trim().Contains('ExternalAuthoritative')) {
            $AuthMechanism = $connector.AuthMechanism
            $PermissionGroups = $connector.PermissionGroups
            $connector.PermissionGroups = 'ExchangeServers'
            $FixAuthMechanism = $true
        }

# Ask the user if any settings need to be changed
$changeSettings = Read-Host "Do you want to change any settings? ((Y)es/(N)o)"

if ($changeSettings -eq "Y") {
    $newConnectorName = Read-Host "Enter the new Connector Name (leave empty to keep current)"
    $newRemoteIPs = Read-Host "Enter the new Remote IP address (if multiple, separate by commas; leave empty to keep current)"
    $newBindings = Read-Host "Enter the new Bindings (leave empty to keep current)"
    $newAuthentication = Read-Host "Enter the new Authentication mechanisms (leave empty to keep current)"
    $newPermissionGroups = Read-Host "Enter the new Permission Groups (leave empty to keep current)"
    $NewTransportRole = Read-Host "Enter the new Transport Role (leave empty to keep current)"


    # Use new values if provided, otherwise use existing values
    $newConnectorName = if ($newConnectorName) { $newConnectorName } else { $connector.Name }
    $newRemoteIPs = if ($newRemoteIPs) { $newRemoteIPs -split ',' } else { $connector.RemoteIPRanges }
    $newBindings = if ($newBindings) { $newBindings } else { $connector.Bindings }
    $newAuthentication = if ($newAuthentication) { $newAuthentication -split ',' } else { $connector.AuthMechanism }
    $newPermissionGroups = if ($newPermissionGroups) { $newPermissionGroups -split ',' } else { $connector.PermissionGroups }
    $NewTransportRole = if ($NewTransportRole) { $NewTransportRole } else { $connector.TransportRole }


} elseif ($changeSettings -eq "N") {
    # Use existing values for direct copy
    $newConnectorName = $connector.Name
    $newRemoteIPs = $connector.RemoteIPRanges
    $newBindings = $connector.Bindings
    $newAuthentication = $connector.AuthMechanism
    $newPermissionGroups = $connector.PermissionGroups
    $NewTransportRole = $connector.TransportRole
}

else {
    Write-Host "Option is not valid, please provide a valid option."
    }


# Create the new connector on the target server
New-ReceiveConnector -Name $newConnectorName -Server $targetServer -RemoteIPRanges $newRemoteIPs -Bindings $newBindings -AuthMechanism $newAuthentication -PermissionGroups $newPermissionGroups -DomainController $DomainController -TransportRole $NewTransportRole


if($FixAuthMechanism) {
            $newConnector = Get-ReceiveConnector -Identity ("{0}\{1}" -f $targetServer, $newConnectorName)
            $newConnector | Set-ReceiveConnector -PermissionGroups $PermissionGroups 
        }


        # Copy the Extended Rights from source to destination
        $sourceRights =  Get-ReceiveConnector -Identity "$sourceServer\$connectorName" | Get-ADPermission  | Where-Object {($_.Deny -eq $false) -and ($_.IsInherited -eq $false)}
        foreach ($right in $sourceRights) {
            Get-ReceiveConnector -Identity "$targetServer\$newConnectorName" | Add-ADPermission -User $right.User -ExtendedRights $right.ExtendedRights
        }

        # Ask user if Extended Rights permissions need to be changed
        $changeRights = Read-Host "Do you want to change Extended Rights permissions on the destination server? (Y/N)"
        
        if ($changeRights -eq "Y") {
            $action = Read-Host "Do you want to Add or Remove permissions? (A/R)"
            $user = Read-Host "Enter the User for Extended Rights (leave empty to keep current)"
            $extendedRights = Read-Host "Enter the Extended Rights (leave empty to keep current)"
            
            # Convert to array if multiple values are provided
            $extendedRightsArray = if ($extendedRights) { $extendedRights -split ',' } else { $right.ExtendedRights }

            # Add or Remove Extended Rights based on user input
            if ($action -eq "A") {
                foreach ($right in $extendedRightsArray) {
                    Get-ReceiveConnector -Identity "$targetServer\$newConnectorName" | Add-ADPermission -User $user -ExtendedRights $right.Trim()
                }
            } elseif ($action -eq "R") {
                foreach ($right in $extendedRightsArray) {
                    Get-ReceiveConnector -Identity "$targetServer\$newConnectorName" | Remove-ADPermission -User $user -ExtendedRights $right.Trim() -Confirm:$false
                }
            } else {
                Write-Host "Options are invalid, please provide a valid option."
            }
        } elseif ($changeRights -ne "N") {
            Write-Host "Options are invalid, please provide a valid option."
        }

        Write-Host "Successfully copied the Receive Connector to $targetServer with any changes you specified."
    }
    catch {
        Write-Error $_.Exception.Message
    }
}




# Create-NewReceiveConnector.ps1

# Function to create a new Receive Connector
Function Create-NewReceiveConnector {
    try {
        # Prompt user for parameters
        $Name = Read-Host "Enter the name of the new Receive Connector"
        $Server = Read-Host "Enter the name of the Exchange Server where the connector will be created"
        $Role = Read-Host "Enter the role for the Receive Connector (HubTransport, FrontendTransport)"
        $Bindings = Read-Host "Enter the IP address bindings for the Receive Connector (Format: IP:Port)"
        $RemoteIPRanges = Read-Host "Enter the remote IP ranges that can connect to this Receive Connector (Format: IP1,IP2)"
        
        # Convert RemoteIPRanges to array
        $RemoteIPArray = $RemoteIPRanges -split ','
        
        # Create the new Receive Connector
        New-ReceiveConnector -Name $Name -Server $Server -transportRole $Role -Bindings $Bindings -RemoteIPRanges $RemoteIPArray

        Write-Host "Successfully created new Receive Connector: $Name on server $Server."
    }
    catch {
        Write-Error $_.Exception.Message
    }
}


# Main Menu Loop
while ($true) {
    Clear-Host
    Write-Host "Menu Options:"
    Write-Host "1) Export-ReceiveConnectors  -  Exports Existing Receive Connectors" -ForegroundColor Blue
    Write-Host " .DESCRIPTION
    This function exports all the Receive Connectors from all servers to CSV.."  -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "2) Export-ExtendedRights - Select a connector and exports to csv it's Extended Rights"  -ForegroundColor Blue
    Write-Host " .DESCRIPTION
    This function exports extended rights of a specific connector. It asks for the connector name "  -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "3) Change-ReceiveConnectorSettings - Select a connector and change Connector Settings"  -ForegroundColor Blue
    Write-Host " .DESCRIPTION
    This function allows the user to change settings like Remote IP Ranges, Bindings, Authentication mechanisms, and Permission Groups of a specified Receive Connector. "  -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "4) Change-ExtendedRights  -  Select a connector and change Extended Rights permissions"  -ForegroundColor Blue
    Write-Host " .DESCRIPTION
    This function allows the user to either add or remove specified Extended Rights from a Receive Connector."  -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "5) Copy-ReceiveConnector -  Copies a ReceiveConnector to another server and aloows to change settings and extended permissions for the destination connector"  -ForegroundColor Blue
    Write-Host " .DESCRIPTION
    This function copies a specified Receive Connector from a source Exchange server to a target Exchange server, with an option to change settings during the copy."  -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "6) Create-NewReceiveConnector  -  Creates a New ReceiveConnector"  -ForegroundColor Blue
    Write-Host " .DESCRIPTION
    This function creates a new receive connector. It asks for name, server, role, bindings and remote IP address "  -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "7) Exit"
    $choice = Read-Host "Please select an option (1-6)"

  switch ($choice) {
        '1' { Export-ReceiveConnectors }
        '2' { Export-ExtendedRights }
        '3' { Change-ReceiveConnectorSettings }
        '4' { Change-ExtendedRights }
        '5' { Copy-ReceiveConnector }
        '6' { Create-NewReceiveConnector }
        '7' { Write-Host "Exiting..."; break }
        default { Write-Host "Invalid option. Please try again." }
    }

    if ($choice -eq '7') {
        break
    }

    Read-Host "Press Enter to continue..."
}
