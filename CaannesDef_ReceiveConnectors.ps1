

# Read the CSV file containing server and binding information
$csvFilePath = Read-Host "Enter the path to the CSV file:"
$settings = Import-Csv $csvFilePath -Delimiter ';'
$thumbprint = ""
$cert = Get-ExchangeCertificate -Thumbprint $thumbprint

# Check if $cert is null
if ($cert -eq $null) {
    $certThumbprint = Read-Host "Thumbprint is required for the TLS namespace. Press (Y) to enter Thumbprint or (C) to cancel."

    if ($certThumbprint -eq "Y") {
        # Prompt the user for the Thumbprint
        $thumbprint = Read-Host "Enter the Thumbprint for TLS Certificate"
    } elseif ($certThumbprint -eq "C") {
        Write-Host "Operation canceled. Exiting..."
        exit
    } else {
        Write-Host "Invalid choice. Exiting..."
        exit
    }
}

$cert = Get-ExchangeCertificate -Thumbprint $thumbprint
$tlscertificatename = "<i>$($cert.Issuer)<s>$($cert.Subject)"




function Export-ConnectorsTOCsv{


# Get the newly created LEGACY connector
$NewLegacyConnector = Get-ReceiveConnector -identity $ConnectorID

# Convert the settings to a PSObject
$LegacyConnectorObject = [PSCustomObject]@{
    Name = $NewLegacyConnector.Name
    Server = $NewLegacyConnector.Server
    TransportRole = $NewLegacyConnector.TransportRole
    AuthMechanism = $NewLegacyConnector.AuthMechanism
    PermissionGroups = $NewLegacyConnector.PermissionGroups
    Bindings = $NewLegacyConnector.Bindings
    RemoteIpRanges = $NewLegacyConnector.RemoteIpRanges
    DefaultDomain = $NewLegacyConnector.DefaultDomain
    Fqdn = $NewLegacyConnector.Fqdn
    MaxInboundConnection = $NewLegacyConnector.MaxInboundConnection
    MaxInboundConnectionPerSource = $NewLegacyConnector.MaxInboundConnectionPerSource
    MaxMessageSize = $NewLegacyConnector.MaxMessageSize
    ProtocolLoggingLevel = $NewLegacyConnector.ProtocolLoggingLevel
    Usage = $NewLegacyConnector.Usage
    TlsCertificateName = $NewLegacyConnector.TlsCertificateName
    RequireTLS = $NewLegacyConnector.RequireTLS
}

#CSV
$server = $NewLegacyConnector.Server
$Name = $NewLegacyConnector.Name
$CSV = "$Server_$Name.csv"
$LegacyConnectorObject | Export-Csv .\$CSV

}




# Loop through each server and create connectors
foreach ($setting in $settings) {
    $ServerName = $setting.ServerName


#Legacy Connector
    # Legacy Connector Settings
    $LegacyConnectorSettingsArray = @()
    $ConectorsNameLEGACY = "SMTP LEGACY $ServerName"
    $BindingsLEGACY = $setting.LEGACYIP
    $LegacyConnectorSettings = @{
        Name = $ConectorsNameLEGACY
        Server = $ServerName
        TransportRole = "FrontendTransport"
        AuthMechanism = "Tls"
        PermissionGroups = "AnonymousUsers"
        Bindings = $BindingsLEGACY
        RemoteIpRanges = "153.88.142.0/24"
        DefaultDomain = "caannesit.info"
        Fqdn = "smtp-legacy.internal.caannesit.info"
        MaxInboundConnection = 10000
        MaxInboundConnectionPerSource = 600
        MaxMessageSize = "35MB"
        ProtocolLoggingLevel = "Verbose"
        Usage = "Internal"
        TlsCertificateName = $tlscertificatename
        RequireTLS = $false
    }

# Display connector settings and ask for confirmation
    Write-Host ""
    Write-Host "Server: $ServerName" -ForegroundColor Yellow
    Write-Host "LEGACY Connector Settings:" -ForegroundColor Yellow
    $LegacyConnectorSettings.GetEnumerator() | ForEach-Object {
        Write-Host "- $($_.Key): $($_.Value)"
    }

    $confirmChoice = Read-Host "Do you want to create LEGACY Connector on this server? (Y/N)"
    if ($confirmChoice -eq "Y") {
        
         try {
        # Create LEGACY connector using splatting
        Write-Host "Creating connector LEFGACY on server $ServerName" -ForegroundColor DarkBlue -BackgroundColor White
        New-ReceiveConnector @LegacyConnectorSettings -Verbose
        Write-Host "LEGACY Connector created successfully on server $ServerName" -ForegroundColor Green

        #Add Permissions
        Write-Host "Adding Permission on server $ServerName for LEGACY connector" -ForegroundColor DarkBlue -BackgroundColor White
        Get-ReceiveConnector "$ServerName\$ConectorsNameLEGACY" | Add-ADPermission -User "NT AUTHORITY\ANONYMOUS LOGON" -ExtendedRights "ms-Exch-SMTP-Accept-Any-Recipient" -Verbose
        Write-Host "Successfully added permissions on server $ServerName for LEGACY Connector " -ForegroundColor Green 
    }

    catch {
            Write-Error "Error creating LEGACY Connector on server $ServerName $($_.Exception.Message)"
        }

        $ConnectorID = "$ServerName\$ConectorsNameLEGACY"
        Export-ConnectorsTOCsv {$ConnectorID}

        }


#Print Connector
    # Printer Connector Settings
    $ConectorsNamePRINT = "SMTP PRINT $ServerName"
    $BindingsPRINT = $setting.PRINTIP
    $PrinterConnectorSettings = @{
        Name = $ConectorsNamePRINT
        TransportRole = "FrontendTransport"
        AuthMechanism = "Tls"
        PermissionGroups = "AnonymousUsers"
        Bindings = $BindingsPRINT
        RemoteIpRanges = "153.88.142.0/24"
        DefaultDomain = "caannesit.info"
        Fqdn = "smtp-print.internal.caannesit.info"
        MaxInboundConnection = 10000
        MaxInboundConnectionPerSource = 600
        MaxMessageSize = "35MB"
        ProtocolLoggingLevel = "Verbose"
        TlsCertificateName = $tlscertificatename
        RequireTLS = $false
    }

# Display connector settings and ask for confirmation
    Write-Host ""
    Write-Host "Server: $ServerName" -ForegroundColor Yellow
    Write-Host "PRINTE Connector Settings:" -ForegroundColor Yellow
    $LegacyConnectorSettings.GetEnumerator() | ForEach-Object {
        Write-Host "- $($_.Key): $($_.Value)"
    }

    $confirmChoice = Read-Host "Do you want to create Printer Connector on this server? (Y/N)"
    if ($confirmChoice -eq "Y") {

    Try {
        # Create Printer connector using splatting
        Write-Host "Creating connector PRINTE on server $ServerName" -ForegroundColor DarkBlue -BackgroundColor White
        New-ReceiveConnector -Server $ServerName @PrinterConnectorSettings -Verbose
        Write-Host "Printer Connector created successfully on server $ServerName" -ForegroundColor Green
    }
    catch {
            Write-Error "Error creating PRINTE Connector on server $ServerName $($_.Exception.Message)"
       
    }

        $ConnectorID = "$ServerName\$ConectorsNamePRINT"
        Export-ConnectorsTOCsv {$ConnectorID}

    }

# Central Connector
    # Central Connector Settings
    $ConectorsNameCENTRAL = "SMTP CENTRAL $ServerName"
    $BindingsCENTRAL = $setting.CENTRALIP
    $CentralConnectorSettings = @{
        Name = $ConectorsNameCENTRAL
        Server = $ServerName
        Usage = "Custom"
        TransportRole = "HubTransport"
        PermissionGroups = "ExchangeServers"
        Bindings = $BindingsCENTRAL
        RemoteIpRanges = "153.88.142.0/24"
        DefaultDomain = "caannesit.info"
        Fqdn = "smtp-central.internal.caannesit.info"
        MaxInboundConnection = 10000
        MaxInboundConnectionPerSource = 600
        MaxMessageSize = "35MB"
        ProtocolLoggingLevel = "Verbose"
        RequireTLS = $false
    }

# Display connector settings and ask for confirmation
    Write-Host ""
    Write-Host "Server: $ServerName" -ForegroundColor Yellow
    Write-Host "CENTRAL Connector Settings:" -ForegroundColor Yellow
    $LegacyConnectorSettings.GetEnumerator() | ForEach-Object {
        Write-Host "- $($_.Key): $($_.Value)"
    }

    $confirmChoice = Read-Host "Do you want to create Central Connector on this server? (Y/N)"
    if ($confirmChoice -eq "Y") {

    Try {
        # Create Central Connector using splatting
        Write-Host "Creating connector CENTRAL on server $ServerName" -ForegroundColor DarkBlue -BackgroundColor White
        New-ReceiveConnector @CentralConnectorSettings -Verbose
        Write-Host "Central Connector created successfully on server $ServerName" -ForegroundColor Green

        #Remove Permiassions
        Write-Host "Removing permissions on server $ServerName for CENTRAL connector" -ForegroundColor DarkBlue -BackgroundColor White
        Get-ReceiveConnector "$ServerName\$ConectorsNameCENTRAL" | Remove-ADPermission -User "MS Exchange\Externally Secured Servers" -ExtendedRights "ms-Exch-SMTP-Accept-Any-Sender"

        Write-Host "Successfully removed permissions on server $ServerName for CENTRAL Connector" -ForegroundColor Green
    }

     
    catch {
            Write-Error "Error creating LEGACY Connector on server $ServerName $($_.Exception.Message)"
       
    }

        $ConnectorID = "$ServerName\$ConectorsNameCENTRAL"
        Export-ConnectorsTOCsv {$ConnectorID}

    }

# Internal Connector
    # Internal Connector Settings
    $ConectorsNameINTERNAL = "SMTP INTERNAL $ServerName"
    $BindingsINTERNAL = $setting.INTERNALIP
    $InternalConnectorSettings = @{
        Name = $ConectorsNameINTERNAL
        Server = $ServerName
        Usage = "Custom"
        TransportRole = "FrontendTransport"
        PermissionGroups = "ExchangeServers"
        AuthMechanism = "ExternalAuthoritative"
        Bindings = $BindingsINTERNAL
        RemoteIpRanges = "153.88.142.0/24"
        DefaultDomain = "caannesit.info"
        Fqdn = "smtp.internal.caannesit.info"
        MaxInboundConnection = 10000
        MaxInboundConnectionPerSource = 600
        MaxMessageSize = "35MB"
        ProtocolLoggingLevel = "Verbose"
        RequireTLS = $false
    }

# Display connector settings and ask for confirmation
    Write-Host ""
    Write-Host "Server: $ServerName" -ForegroundColor Yellow
    Write-Host "INTERNAL Connector Settings:" -ForegroundColor Yellow
    $LegacyConnectorSettings.GetEnumerator() | ForEach-Object {
        Write-Host "- $($_.Key): $($_.Value)"
    }

    $confirmChoice = Read-Host "Do you want to create Internal Connector on this server? (Y/N)"
    if ($confirmChoice -eq "Y") {

    try{
        # Create Internal Connector using splatting
        Write-Host "Creating connector INTERNAL on server $ServerName" -ForegroundColor DarkBlue -BackgroundColor White
        New-ReceiveConnector @InternalConnectorSettings -Verbose
        
        Write-Host "Internal Connector created successfully on server $ServerName" -ForegroundColor Green

         #Remove Permiassions
        Write-Host "Removing permissions on server $ServerName for INTERNAL connector" -ForegroundColor DarkBlue -BackgroundColor White
        Get-ReceiveConnector "$ServerName\$ConectorsNameINTERNAL" | Remove-ADPermission -User "MS Exchange\Externally Secured Servers" -ExtendedRights "ms-Exch-SMTP-Accept-Any-Sender"

        Write-Host "Successfully removed permissions on server $ServerName for INTERNAL Connector " -ForegroundColor Green
    }
     
    catch {
            Write-Error "Error creating LEGACY Connector on server $ServerName $($_.Exception.Message)"
       
    }

        $ConnectorID = "$ServerName\$ConectorsNameINTERNAL"
        Export-ConnectorsTOCsv {$ConnectorID}


}



}


Write-Host "ALL CONNECTORS CREATED CORRECTLY." -ForegroundColor DarkGreen
Write-Host "An export has been done for every connector created to the location from where you ran the script" -ForegroundColor DarkYellow