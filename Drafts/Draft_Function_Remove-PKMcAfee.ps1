Function Remove-PKMcAfee {
    [Cmdletbinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
    Param([switch]$Silent)
    
    $ProgressPreference = "Continue"
    $Activity = "Look for and remove McAfee enterprise product(s)"
    $Msg = "Search for executable"
    Write-Progress -Activity $Activity
    If ($Null = Get-Item "C:\Program Files (x86)\McAfee\Common Framework\x86\FrmInst.exe") {
        $Msg = "Found C:\Program Files (x86)\McAfee\Common Framework\x86\FrmInst.exe"
        Write-Verbose $Msg
        $Msg = "Remove agent"
        If ($Silent.IsPresent) {
            $Cmd = '"C:\Program Files (x86)\McAfee\Common Framework\x86\FrmInst.exe" /Remove=Agent /Silent'
            $Msg += " (silent mode)"
        }
        Else {
            $Cmd = '"C:\Program Files (x86)\McAfee\Common Framework\x86\FrmInst.exe" /Remove=Agent'
        }
        Write-Verbose $Msg
        $ConfirmMsg = "`n`n`t$Msg`n`n"
        If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$ConfirmMsg)) {
            Try {
                Invoke-Expression $Cmd
            }
            Catch {
                Throw $_.Exception.Message
            }
        }
    }

}