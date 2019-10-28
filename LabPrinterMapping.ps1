Function Get-ComputerOU {
    <#
    .Synopsis
       Finds the Organizational Unit location of the Computer queried
    .DESCRIPTION
       Get-ComputerOU uses a System.DirectoryServices object to traverse
       Active Directory for the location in the structure of the computer.
       Information output is the computer name and the last 3 instances of
       the OU structure the computer object resides.
    .EXAMPLE
       Get-ComputerOU -ComputerName SERVER1
    .EXAMPLE
        Get-ComputerOU -ComputerName SERVER1, SERVER2, SERVER3
    #>
    [CmdletBinding()]
    Param (

        [Parameter( Mandatory=$True,
                    ValueFromPipelineByPropertyName=$True,
                    HelpMessage='ComputerName of Computer you want to lookup')]
        [String[]]$ComputerName

    )

    Begin {}

    Process {

        Try {

            ForEach ($Computer in $ComputerName) {

                $Filter = "(&(objectCategory=Computer)(Name=$Computer))"
                $DirectorySearcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
                $DirectorySearcher.Filter = $Filter
                $SearcherPath = $DirectorySearcher.FindOne()
                $DistinguishedName = $SearcherPath.GetDirectoryEntry().DistinguishedName

                $OUName_Last = ($DistinguishedName.Split(','))[1]
                $OUName_Second = ($DistinguishedName.Split(','))[2]
                $OUName_First = ($DistinguishedName.Split(','))[3]

                $OUName1 = $OUName_Last.SubString($OUName_Last.IndexOf('=')+1)
                $ouname2 = $OUName_Second.SubString($OUName_Second.IndexOf('=')+1)
                $ouname3 = $OUName_First.SubString($OUName_First.IndexOf('=')+1)

                $Properties = @{

                    'ComputerName' = $Computer
                    'OrgUnit'      = $OUName3 + '-' + $OUName2 + '-' + $ouname1

                }

                $Obj = New-Object -TypeName PSObject -Property $Properties
                $Obj.PSObject.TypeNames.Insert(0,'PC.OrgUnit')
                Write-Output -InputObject $Obj

                #if ($obj -eq $Null) {
                if ($Null -eq $obj) {

                    Write-Error -Message 'Unable to find Computer OU based off computername' -ErrorAction 'Stop'

                }

            }

        } Catch {

            # get error record
            [Management.Automation.ErrorRecord]$e = $_

            # retrieve information about runtime error
            $info = [PSCustomObject]@{

                Date         = Get-Date
                ComputerName = $ENV:computerName
                Exception    = $e.Exception.Message
                Reason       = $e.CategoryInfo.Reason
                Target       = $e.CategoryInfo.TargetName
                Script       = $e.InvocationInfo.ScriptName
                Line         = $e.InvocationInfo.ScriptLineNumber
                Column       = $e.InvocationInfo.OffsetInLine

            }

            # output information. Post-process collected info, and log info (optional)
            Write-output -InputObject $info

        }

    }

    End {}

} #Get-ComputerOU


Function Set-LabPrinters {
    #Requires -Version 3
    <#
    .SYNOPSIS
    Map network printers from server to computer

    .DESCRIPTION
    Set-LabPrinters uses Add-printer to connect to existing printers on a print server

    .EXAMPLE
    Set-LabPrinters -ServerAddress \\Server -DefaultPrinter 'Prhh112'

    #>

    [CmdletBinding()]
    Param (

        [Parameter( Mandatory=$True,
                    ValueFromPipelineByPropertyName=$True,
                    HelpMessage='Lab identifier - change value type based on needs')]
        [String]$Org,

        [Parameter( Mandatory = $True,
                    HelpMessage='UNC path for Print Server')]
        [string]$ServerAddress,

        [Parameter( Mandatory=$True,
                    HelpMessage='Printer name on the server of the primary printer based on physical location')]
        [String]$DefaultPrinter,

        [string]$PrinterSecondary,

        [string]$PrinterThird,

        [String]$ErrorLog = "$ENV:WINDIR\Temp\LabPrinterMapping.log"

    )

    Begin {}

    Process {

        Try {

            #Add default printer
            if ($DefaultPrinter -ne '') {

                $PrinterConnections = @{

                    'ConnectionName' = "$ServerAddress\$DefaultPrinter"
                    'ErrorAction'    = 'stop'

                }

                Write-Verbose -Message "Installing default printer $($DefaultPrinter)"
                Add-Printer @PrinterConnections

            } Else {

                Write-Error -Message 'No Default Printer Found'

            }

            if ($PrinterSecondary -ne '') {

                $PrinterConnections = @{

                    'ConnectionName' = "$ServerAddress\$PrinterSecondary"
                    'ErrorAction'    = 'Stop'

                }

                Write-Verbose -Message "Installing printer $($PrinterSecondary)"
                Add-Printer @PrinterConnections

            }

            if ($PrinterThird -ne '') {

                $PrinterConnections = @{

                    'ConnectionName' = "$ServerAddress\$PrinterThird"
                    'ErrorAction'    = 'Stop'

                }

                Write-Verbose -Message "Installing printer $($PrinterThird)"
                Add-Printer @PrinterConnections

            }

        } Catch {

           [Management.Automation.ErrorRecord]$e = $_

            # retrieve information about runtime error
            $info = [PSCustomObject]@{

                Date         = Get-Date
                ComputerName = $ENV:computerName
                UserName     = $env:USERNAME
                Exception    = $e.Exception.Message
                Reason       = $e.CategoryInfo.Reason
                Target       = $e.CategoryInfo.TargetName
                Script       = $e.InvocationInfo.ScriptName
                Line         = $e.InvocationInfo.ScriptLineNumber
                Column       = $e.InvocationInfo.OffsetInLine

            }

            # output information. Post-process collected info, and log info (optional)
            Write-output -InputObject $info | Out-file -FilePath $ErrorLog -Append

        }

    }

    End {}

}


Function Set-Win7Printers {
    <#
    .SYNOPSIS
    Map network printers from server to computer

    .DESCRIPTION
    Set-Win7Printers uses Wscript network object  methods to connect to existing printers on a print server

    .EXAMPLE
    Set-Win7Printers -ServerAddress \\Server -DefaultPrinter 'Prhh112'

    #>
    [cmdletBinding()]
    Param (

        [Parameter( Mandatory=$True,
                    ValueFromPipelineByPropertyName=$True,
                    HelpMessage='Lab identifier - change value type based on needs')]
        [String]$Org,

        [Parameter( Mandatory = $True,
                    HelpMessage='UNC path for Print Server')]
        [string]$ServerAddress,

        [Parameter( Mandatory=$True,
                    HelpMessage='Printer name on the server of the primary printer based on physical location')]
        [String]$DefaultPrinter,

        [string]$PrinterSecondary,

        [string]$Printerthird,

        [String]$ErrorLog = "$ENV:WINDIR\Temp\LabPrinterMapping.log"

    )

    Begin {}

    Process {

        Try {

            if ($DefaultPrinter -ne '') {

                $NetObj = New-object -ComObject Wscript.Network
                $Netobj.AddWindowsPrinterConnection("$ServerAddress\$DefaultPrinter")

            } else {

                Write-Error -Message 'No reference to Default Printer'

            }

            if ($Printersecond -ne '') {

                $NetObj = New-object -ComObject Wscript.Network
                $Netobj.AddWindowsPrinterConnection("$ServerAddress\$PrinterSecondary")

            }

            if ($Printerthird -ne '') {

                $NetObj = New-object -ComObject Wscript.Network
                $Netobj.AddWindowsPrinterConnection("$ServerAddress\$PrinterThird")

            }

        } Catch {

            [Management.Automation.ErrorRecord]$e = $_

            # retrieve information about runtime error
            $info = [PSCustomObject]@{

                Date         = Get-Date
                ComputerName = $ENV:computerName
                UserName     = $env:USERNAME
                Exception    = $e.Exception.Message
                Reason       = $e.CategoryInfo.Reason
                Target       = $e.CategoryInfo.TargetName
                Script       = $e.InvocationInfo.ScriptName
                Line         = $e.InvocationInfo.ScriptLineNumber
                Column       = $e.InvocationInfo.OffsetInLine

            }

            # output information. Post-process collected info, and log info (optional)
            Write-output -InputObject $info | Out-file -FilePath $ErrorLog -Append

        }

    }

    End {}
}


#controller Script
$ScriptPath = $MyInvocation.MyCommand.Path
$CurrentDir = Split-Path -Path $ScriptPath
[String]$CsvPath = "$currentdir\LabPrinters.csv"

[String]$Log = "$env:windir\Temp\LabPrinterMapping.log"

$PrinterConfig = Import-Csv -Path "$CSVPath"
$OU = Get-ComputerOU -ComputerName "$ENV:COMPUTERNAME"

$PrinterConfig | ForEach-Object {

    if ($OU.orgunit -eq $_.Laborg ) {

        Do {

            Try {

                $LabPrinterParams = @{

                    'ServerAddress'    = $_.ServerAddress
                    'DefaultPrinter'   = $_.DefaultPrinter
                    'PrinterSecondary' = $_.PrinterSecondary
                    'PrinterThird'     = $_.PrinterThird
                    'ErrorLog'         = "$ENV:WINDIR\Temp\LabPrinterMapping.log"
                    'ErrorAction'      = 'Stop'


                }

                $OSWMI = @{

                    'class'       = 'Win32_operatingSystem'
                    'ErrorAction' = 'stop'

                }

                $OS = Get-wmiobject @OSWMI

                if ($OS.Version -like '10.*' ) {

                    Set-LabPrinters @LabPrinterParams

                } else {

                    Set-Win7Printers @LabPrinterParams

                }

                Write-Output -InputObject "$(Get-Date) : $ENV:COMPUTERNAME : Username $ENV:USERNAME : Printers to be installed $($_.DefaultPrinter, $_.PrinterSecondary, $_.PrinterThird)" | Out-file -FilePath $Log -Append
                Write-Output -InputObject "$(Get-Date) : $ENV:COMPUTERNAME : Username $ENV:USERNAME : Default Printer $($_.DefaultPrinter)" | Out-file -FilePath $Log -Append

                $PrinterWMI = @{

                    'Class'       = 'Win32_Printer'
                    'ErrorAction' = 'Stop'

                }

                $InstalledPrinters  = Get-WmiObject @PrinterWMI | Where-object {$_.Name -like '\\*'}
                Write-output -InputObject "$(Get-Date) : $ENV:COMPUTERNAME : Username $ENV:USERNAME : Installed Printers $($InstalledPrinters.Name)" | Out-file -FilePath $Log -Append

                if ($InstalledPrinters -eq $Null) {

                    Write-Error -Message "$(Get-Date) : $ENV:COMPUTERNAME : Username $ENV:USERNAME : Error - Printers did not get installed from $($_.ServerAddress)"

                }

            } Catch {

                [Management.Automation.ErrorRecord]$e = $_

                # retrieve information about runtime error
                $info = [PSCustomObject]@{

                    Date         = Get-Date
                    ComputerName = $ENV:computerName
                    UserName     = $env:USERNAME
                    Exception    = $e.Exception.Message
                    Reason       = $e.CategoryInfo.Reason
                    Target       = $e.CategoryInfo.TargetName
                    Script       = $e.InvocationInfo.ScriptName
                    Line         = $e.InvocationInfo.ScriptLineNumber
                    Column       = $e.InvocationInfo.OffsetInLine

                }

                # output information. Post-process collected info, and log info (optional)
                Write-output -InputObject $info | Out-file -FilePath $Log -Append

            }

        } Until($InstalledPrinters -ne $Null)

    }

}