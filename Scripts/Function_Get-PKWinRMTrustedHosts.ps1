#Requires -Version 3
Function Get-PKWinRMTrustedHosts {
<#
.SYNOPSIS
    Uses Get-Item to return the trusted hosts for WinRM configured on the local computer

.DESCRIPTION
    Uses Get-Item to return the trusted hosts for WinRM configured on the local computer
    Returns a PSobject

.NOTES        
    Name    : Function_Get-PKWinRMTrustedHosts.ps1
    Created : 2019-02-25
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-02-25 - Created script

.LINK
    https://devblogs.microsoft.com/scripting/powertip-use-powershell-to-view-trusted-hosts/

.EXAMPLE
    PS C:\> Get-PKWinRMTrustedHosts -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                  
        ---           -----                  
        Verbose       True                   
        ScriptName    Get-PKWinRMTrustedHosts
        ScriptVersion 1.0.0                  

        VERBOSE: Get TrustedHosts for local computer

        ComputerName    Name         SourceOfValue Value                    
        ------------    ----         ------------- -----                    
        WORKSTATION12   TrustedHosts               *.domain.local


#>
[CmdletBinding()]
Param()
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

}
Process {    
    
    $Msg = "Get TrustedHosts for local computer"
    Write-Verbose $Msg
    Get-Item WSMan:\localhost\Client\TrustedHosts | Select @{N="ComputerName";E={$env:COMPUTERNAME}},Name,SourceOfValue,Value
}
} #end function

$null = New-Alias Get-TrustedHosts -Value Get-PKWinRMTrustedHosts -Description "Easier to remember!" -Force -Confirm:$False
$null = New-Alias Get-PKTrustedHosts -Value Get-PKWinRMTrustedHosts -Description "Easier to remember!" -Force -Confirm:$False