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

        [String]$INF,

        [String]$Driver,

        [String]$IP,

        [String]$PrinterName

    )

    Begin {}

    Process {

        Write-Verbose -Message "Creating TCP/IP Printer Port for $PrinterName"
        Add-PrinterPort -Name "$IP" -PrinterHostAddress "$IP"

        Write-Verbose -Message 'Import INF file to Windows Driver Store'
        $argumentlist = @(

            '-I'
            '-A'
            "$INF"

        )
        #pnputil.exe -I -A "$INF"
        Start-Process -FilePath "pnputil.exe" -ArgumentList $argumentlist -Wait -NoNewWindow

        Write-Verbose -Message "Install printer driver for $printername"
        Add-PrinterDriver -Name "$Driver"

        Write-Verbose -Message "Configuring $PrinterName"
        Add-Printer -Name "$PrinterName" -DriverName "$Driver" -PortName "$IP"

    }

    End {}

}