Function Search-RegistryHive {
<#

    .SYNOPSIS
    Search Registry path for specific keyword

    .DESCRIPTION
    Search-RegistryHive will search the given path for specifc keywords and output the full path to Registry keys
    that contain that keyword 

    .EXAMPLE
    Search-RegistryHive -path 'HKLM:\SYSTEM\ControlSet001\Control\Print\Printers' -searchText $Port

    Example searching the registry for printers that match the $port search text criteria

    .NOTES
    Works on PowerShell 3.0 and higher

#>
    [CmdletBinding()]
    Param (

        [Parameter( Mandatory = $true,
                    HelpMessage = 'Keyword to search Registry for')]
        [string]$searchText,

        [Parameter()]
        [String]$path

    )

    Get-ChildItem -Path $path -Recurse | ForEach-Object {

        if ((Get-itemproperty -Path $_.pspath) -match $searchText) {

            $_.Name

        }

    }

}


Function Get-WSDPortIP {
<#

    .SYNOPSIS
    Resolve IP Address of a network printer mapped via a WSD port

    .DESCRIPTION
    If a network printer is configured on a computer using a WSD port, Get-WSDPortIP will take the portname of the printer
    which should look something like 'WSD-96d31e...' and search the HKLM registry hive for the particualr key that contains the
    IP Address of the network printer. 

    .PARAMETER WSDPort
    Portname of printer(s) configured via WSD

    .EXAMPLE
    Get-WSDPort -WSDPort 'WSD-96d31ef9-e17b-41a4-9adc-ad699617382c.0069'

    Example using a single WSD port identifier

    .EXAMPLE
    Get-wsdport -WSDPort (Get-Printer | Where-Object {$_.Portname -like "WSD*"}).PortName

    Example of use case if there are muultple WSD printers on a computer, the above will get all WSD portnames and resolve
    the IP Address of eachone.

    .NOTES
    Writtren as a helper function for a larger script to query printer information on computers in various OU's and output
    to database so that we can create print queues on a print server and Print mapping GPO's 
    
    Works on PowerShell 3.0 and higher

#>

    [CmdletBinding()]
    Param (

        [Parameter( Mandatory = $true,
                    ValueFromPipelineByPropertyName = $true,
                    ValueFromPipeline = $True,
                    Position = 0,
                    HelpMessage = 'Portname of printer configured via WSD - Helpful cmdlet -> Get-printer | Select name, portname')]
        [string[]]$WSDPort

    )

    Begin  {}

    Process {

        foreach ($port in $WSDPort) {
            
            Try {

                $Subkeys = (Search-RegistryHive -path 'HKLM:\SYSTEM\ControlSet001\Control\Print\Printers' -searchText $Port) -replace '^[^\\]*', 'HKLM:'


                $Subkeys | ForEach-Object {
        
                    if ($_ -like '*PrinterDriverData') {
        
                        $KeyProps = Get-ItemProperty -Path $_
        
                        $props = @{
        
                            'WsdPort' = $Port

                            'IPAddress' = ($KeyProps.HPEWSIPAddress).Split(',')[0]
        
                        }
        
                        $Object = New-Object -TypeName PSObject -Property $props
                        $object.PSObject.typenames.insert(0,'WSDPort.IPAddress')
                        Write-Output -InputObject $Object
        
                    }
        
                }

            } Catch {
            
                # get error record
                [Management.Automation.ErrorRecord]$e = $_

                # retrieve information about runtime error
                $info = [PSCustomObject]@{
                
                  Exception = $e.Exception.Message
                  Reason    = $e.CategoryInfo.Reason
                  Target    = $e.CategoryInfo.TargetName
                  Script    = $e.InvocationInfo.ScriptName
                  Line      = $e.InvocationInfo.ScriptLineNumber
                  Column    = $e.InvocationInfo.OffsetInLine
                  
                }
                
                # output information. Post-process collected info, and log info (optional)
                $info
                
            }

        }

    }

    End {}

}
