
#requires -Version 5.1
Function Get-PKPublicIP {<#
.SYNOPSIS
    Retrieves the public IP address of the machine via an API call to the ifconfig.me service

.DESCRIPTION
    Retrieves the public IP address of the machine via Invoke-WebRequest to 'https://ifconfig.me/ip'
    It returns the IP address as a trimmed string

.NOTES
    Name    : Function_Get-PKPublicIP.ps1
    Created : 2025-04-28
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2025-04-28 - Created script

.LINK
    https://www.powershellgallery.com/packages/WinTZ/0.1.0/Content/Public%5CGet-IANATimeZone.ps1        

.EXAMPLE
    PS C:\> Get-PKPublicIP -Verbose

        VERBOSE: PSBoundParameters: 

        Key           Value
        ---           -----
        Verbose       True
        ScriptName    Get-PKPublicIP
        ScriptVersion 1.0.0

        VERBOSE: Getting public IP address from 'https://ifconfig.me/ip' and ISP details from 'http://ip-api.com/json'
                                                                                                                                
        IPAddress   : 151.101.130.132
        ISP         : Telecom R Us
        City        : New York
        Region      : NY
        Coordinates : 40.7339,-74.0054
        TimeZone    : America/New_York

#>

    [CmdletBinding()]
    Param()
    Begin {
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.0.0000"

        # How did we get here
        $ScriptName = $MyInvocation.MyCommand.Name

        # Show our settings
        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        #$CurrentParams.Add("PipelineInput", $PipelineInput)
        #$CurrentParams.Add("ParameterSetName", $Source)
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    }
    Process {
        Try {
            $IPURI = "https://ifconfig.me/ip"
            $JSONURI ="http://ip-api.com/json"
            $Msg = "Getting public IP address from '$IPURI' and ISP details from '$JSONURI'"
            Write-Verbose $Msg
            $IPAddress = Invoke-WebRequest -Uri $IPURI -ContentType 'text/plain' -ErrorAction Stop -Verbose:$False | Select-Object -ExpandProperty Content
            
            $Results = Invoke-RestMethod -Method Get -Uri $JSONURI -ErrorAction Stop -Verbose:$False
            [PSCustomObject]@{
                IPAddress   = $IPAddress
                ISP         = $Results.ISP
                City        = $Results.City
                Region      = $Results.Region
                Coordinates = "$($Results.Lat),$($Results.Lon)"
                TimeZone    = $Results.TimeZone
            }
        }
        Catch {
            $Msg = "Operation failed! $($_.Exception.Message)"
            Write-Error $Msg
        }
    }
} #end Get-PKPublicIP