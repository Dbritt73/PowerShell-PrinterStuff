function Install-NetworkPrinter {
<#
.Synopsis
   Install Network printer

.DESCRIPTION
   Install-NetworkPrinter utilizes Cmdlets from the PrintManagement module in conjunction with PNPUtil to create a TCP/IP
   port, ingest the printer driver to the driver store, Install printer driver from driver store, and finally create the
   printer with these properties.

.EXAMPLE
   Install-NetworkPrinter -INF "FolderPath\subfolder\hpbuio170l.inf" -Driver 'HP Universal Printing PCL 5' -IP '127.0.0.1' -PrinterName 'Name of Printer' -Verbose

#>
    [CmdletBinding()]
    Param (

        [Parameter( Mandatory = $true,
                    HelpMessage = 'File path to inf file')]
        [String]$InfPath,

        [Parameter( Mandatory = $true,
                    HelpMessage = 'Name of driver as found in inf file')]
        [String]$Driver,

        [Parameter( Mandatory = $true,
                    HelpMessage = 'IP address of network printer to install')]
        [String]$IP,

        [Parameter( Mandatory = $true,
                    HelpMessage = 'What you want to name the printer')]
        [String]$PrinterName

    )

    Begin {}

    Process {

        Try {

            Write-Verbose -Message "Creating TCP/IP Printer Port for $PrinterName with address of $IP"
            Add-PrinterPort -Name "$IP" -PrinterHostAddress "$IP"

            Write-Verbose -Message 'Import INF file to Windows Driver Store'
            $argumentlist = @(

                '-I'
                '-A'
                "$InfPath"

            )

            Start-Process -FilePath "pnputil.exe" -ArgumentList $argumentlist -Wait -NoNewWindow

            Write-Verbose -Message "Install printer driver for $printername"
            Add-PrinterDriver -Name "$Driver"

            Write-Verbose -Message "Configuring $PrinterName"
            Add-Printer -Name "$PrinterName" -DriverName "$Driver" -PortName "$IP"

        } Catch {}

    }

    End {}

}