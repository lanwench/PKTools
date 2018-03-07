#Requires -version 2.0
Function Convert-PKBytesToSize {
<#
.SYNOPSIS 
    Converts any integer size given to a user friendly size

.DESCRIPTION
    Converts any integer size given to a user friendly size

.NOTES        
    Name    : Function_Convert-PKBytesToSize.ps1
    Created : 2018-02-13
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-02-13 - Created script based on Boe Prox's original

.LINK
    https://learn-powershell.net/2010/08/29/convert-bytes-to-highest-available-unit/

.PARAMETER size
    Mandatory; integer to conver to more readable format
    
.EXAMPLE
    PS C:\> Convert-PKBytesToSize -Size 40000 -Verbose
        
        VERBOSE: PSBoundParameters: 
	
        Key           Value                
        ---           -----                
        Verbose       True                 
        Size          40000                    
        PipelineInput False
        ScriptName    Convert-PKBytesToSize
        ScriptVersion 1.0.0           

        VERBOSE: Convert to KB
        39.06KB

.EXAMPLE
    PS C:\> Get-Childitem c:\windows\temp -File | Select -ExpandProperty Length | Convert-PKBytesToSize -Verbose
    
        VERBOSE: PSBoundParameters: 
	
        Key           Value                
        ---           -----                
        Verbose       True                 
        Size          0                    
        PipelineInput True
        ScriptName    Convert-PKBytesToSize
        ScriptVersion 1.0.0                

        VERBOSE: Convert to Bytes
        608Bytes
        VERBOSE: Convert to KB
        3.85KB
        VERBOSE: Convert to Bytes
        606Bytes
        VERBOSE: Convert to Bytes
        102Bytes
        VERBOSE: Convert to MB
        1.38MB
    
#>

[CmdletBinding()]
Param(
    [parameter(
        Mandatory=$True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        Position=0,
        HelpMessage = "Size (int64)"
    )]
    [ValidateNotNullOrEmpty()]
    [int64]$Size
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("Size")) -and ((-not $ComputerName))
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"

}
Process {

    Switch ($Size){
        {$Size -gt 1PB}{
            Write-Verbose "Convert to PB"
            $NewSize = "$([math]::Round(($Size / 1PB),2))PB"
            Break
        }
        {$Size -gt 1TB} {
            Write-Verbose "Convert to TB"
            $NewSize = "$([math]::Round(($Size / 1TB),2))TB"
            Break
        }
        {$Size -gt 1GB} {
            Write-Verbose "Convert to GB"
            $NewSize = "$([math]::Round(($Size / 1GB),2))GB"
            Break
        }
        {$Size -gt 1MB} {
            Write-Verbose "Convert to MB"
            $NewSize = "$([math]::Round(($Size / 1MB),2))MB"
            Break
        }
        {$Size -gt 1KB} {
            Write-Verbose "Convert to KB"
            $NewSize = "$([math]::Round(($Size / 1KB),2))KB"
            Break
        }
        Default {
            Write-Verbose "Convert to Bytes"
            $NewSize = "$([math]::Round($Size,2))Bytes"
            Break
        }
    }
    Write-Output $NewSize
}
} #end Convert-PKBytesToSize