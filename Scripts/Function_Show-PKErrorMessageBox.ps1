#Requires -Version 3
function Show-PKErrorMessageBox {
<#
.SYNOPSIS
    Uses Windows Forms to display an error message in a message box (defaults to most recent error)

.DESCRIPTION
    Uses Windows Forms to display an error message in a message box (defaults to most recent error)
    Accepts pipeline input
    No output

.NOTES
    Name    : Function_Show-PKErrorMessageBox.ps1 
    Created : 2019-03-04
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.01.0000 - 2019-03-04 - Created from ITMicaH's original

.LINK
    https://raw.githubusercontent.com/ITMicaH/Powershell-functions/master/Active-Directory/OUs/ChooseADOrganizationalUnit.ps1

.EXAMPLE
    PS C:\> Show-PKErrorMessageBox -Verbose
    
        VERBOSE: PSBoundParameters: 
	
        Key           Value                                                          
        ---           -----                                                          
        Verbose       True                                                           
        Message       Cannot find an overload for "Show" and the argument count: "1".
        ScriptName    Show-PKErrorMessageBox                                         
        ScriptVersion 1.0.0  

        VERBOSE: Display error message in Windows Form message box:
            Cannot find an overload for "Show" and the argument count: "1".

.EXAMPLE
    PS C:\> Show-PKErrorMessageBox -Message $Error[13]
#>

[CmdletBinding()]
Param(
    [Parameter(
        HelpMessage = "Error message",
        ValueFromPipeline = $True
    )]
    [string]$Message = $Error[0]
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Display our parameters
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
    
    $Msg = "Display error message in Windows Form message box:`n`t$Message"
    Write-Verbose $Msg
    $msgbox = [System.Windows.Forms.MessageBox]::Show($Message,"Exception Report",0,16)
    
}
}