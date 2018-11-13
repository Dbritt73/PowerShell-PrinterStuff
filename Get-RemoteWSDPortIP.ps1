Function Search-RemoteRegistryHive {
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
        [String]$path,

        [Parameter()]
        [String]$ComputerName 

    )

    Invoke-Command -ComputerName $ComputerName -ScriptBlock {Get-ChildItem -Path $Using:path -Recurse | ForEach-Object {

            if ((Get-itemproperty -Path $_.pspath) -match $Using:searchText) {

                $_.Name

            }

        }

    }

}

Function Get-RemoteWSDPortIP {
  <#
      .SYNOPSIS
      Describe purpose of "Get-RemoteWSDPortIP" in 1-2 sentences.

      .DESCRIPTION
      Add a more complete description of what the function does.

      .PARAMETER ComputerName
      Describe parameter -ComputerName.

      .EXAMPLE
      Get-RemoteWSDPortIP -ComputerName 'SERVER01'
      Fetches any instance of a pritner with a WSD named port from SERVER01 and resolves the IP Address(es)

      .NOTES
      Place additional notes here.

      .LINK
      URLs to related sites
      The first link is opened by Get-Help -Online Get-RemoteWSDPortIP

      .INPUTS
      List of input types that are accepted by this function.

      .OUTPUTS
      List of output types produced by this function.
  #>


    [CmdletBinding()]
    Param (

        [String[]]$ComputerName

    )

    Begin {}

    Process {

        Foreach ($Computer in $ComputerName) {

            Try {

                $Printer = Get-Printer -ComputerName $Computer | Where-Object {$_.PortName -like 'WSD*'}

                foreach ($port in $Printer.PortName) {

                    $Subkeys = (Search-RegistryHive -path 'HKLM:\SYSTEM\ControlSet001\Control\Print\Printers' -searchText "$Port" -ComputerName "$Computer") -replace '^[^\\]*', 'HKLM:'

                    $Subkeys | ForEach-Object {
            
                        if ($_ -like '*PrinterDriverData') {
            
                            $KeyProps = Invoke-Command -ScriptBlock {Get-ItemProperty -Path $Using:_} -ComputerName $Computer
            
                            $props = [Ordered]@{

                                'ComputerName' = $Computer

                                'Printer' = ($Printer | Where-Object {$_.PortName -eq $Port}).Name
            
                                'WsdPort' = $Port
    
                                'IPAddress' = ($KeyProps.HPEWSIPAddress).Split(',')[0]
            
                            }
            
                            $Object = New-Object -TypeName PSObject -Property $props
                            $object.PSObject.typenames.insert(0,'WSDPort.IPAddress')
                            Write-Output -InputObject $Object
            
                        }
            
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