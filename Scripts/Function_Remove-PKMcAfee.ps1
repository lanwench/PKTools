#requires -Version 3
Function Remove-PKMcAfee {
<#
.SYNOPSIS
    Removes McAfee Enterprise endpoint client from local computer without a key

.DESCRIPTION
    Removes McAfee Enterprise endpoint client from local computer without a key
    Returns a string

.NOTES
    Name    : Function_Remove-PKMcAfee.ps1 
    Created : 2019-05-06
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.00.0000 - 2019-05-06 - Created script

.PARAMETER Silent
    Run uninstall in silent mode

.EXAMPLE
    PS C:\> Remove-PKMcAfee -Verbose


#>
[Cmdletbinding(
    SupportsShouldProcess=$True,
    ConfirmImpact="High"
)]
Param(
    [Parameter(
        HelpMessage = "Run uninstall in silent mode"
    )]
    [switch]$Silent
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    If ($PipelineInput.IsPresent) {$CurrentParams.InputObject = $Null}
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    $File = "C:\Program Files (x86)\McAfee\Common Framework\x86\FrmInst.exe"

    # Console output
    $Activity = "Remove McAfee enterprise product(s)"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    $Host.UI.WriteLine($FGColor,$BGColor,"$Msg")
}
Process {    
    
    $Msg = "Search for executable"
    Write-Progress -Activity $Activity -CurrentOperation $Msg
    
    If ($Null = Get-Item $File -EA SilentlyContinue) {
        
        $Msg = "Found '$File'"
        Write-Verbose "[$Env:ComputerName] $Msg"
        
        $Msg = "Remove agent"
        If ($Silent.IsPresent) {
            $Cmd = '"C:\Program Files (x86)\McAfee\Common Framework\x86\FrmInst.exe" /Remove=Agent /Silent'
            $Msg += " (silent mode)"
        }
        Else {
            $Cmd = '"C:\Program Files (x86)\McAfee\Common Framework\x86\FrmInst.exe" /Remove=Agent'
        }
        Write-Progress -Activity $Activity -CurrentOperation $Msg
        Write-Verbose "[$Env:ComputerName] $Msg"
        $ConfirmMsg = "`n`n`t$Msg`n`n"
        If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$ConfirmMsg)) {
            Try {
                Invoke-Expression $Cmd
            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += ($ErrorDetails)}
                $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
            }
        }
        Else {
            $Msg = "Operation cancelled by user"
            Write-Verbose "[$Env:ComputerName] $Msg"
        }
    }
    Else {
        $Msg = "'$File' not found"
        $Host.UI.WriteErrorLine("[$Env:ComputerName] $Msg")
    }
}
End {
    Write-Progress -Activity $Activity -Completed
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    $Host.UI.WriteLine($FGColor,$BGColor,"$Msg")
}
} #end Remove-PKMcAfee