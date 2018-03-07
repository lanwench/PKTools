#requires -Version 3
Function Restore-PKISESession {
<#
.SYNOPSIS 
    Restores tabs/files from text file created using Save-PKISESession

.DESCRIPTION
    Restores tabs/files from text file created using Save-PKISESession
    
.NOTES
    Name    : Function_Restore-PKISESession.ps1
    Created : 2016-05-29
    Author  : Paula Kingsley
    Version : 02.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2016-05-29 - Created script based on links
        v02.00.0000 - 2018-02-14 - Updated/made consistent with new Save-PKISESession
        
.LINK
    https://itfordummies.net/2014/10/27/save-restore-powershell-ise-opened-scripts/

.LINK
    https://stackoverflow.com/questions/3710374/get-encoding-of-a-file-in-windows   

.EXAMPLE
    PS C:\> Restore-PKISESession -Verbose

#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Mandatory = $False,
        Position=0,
        ValueFromPipeline = $True,
        HelpMessage = "Full path to text file (created using Save-PKISESession)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (Test-Path $_) {$True}})]
    [String]$SessionFile = "$Home\PSISESession.txt"
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "02.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    
    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    # Make sure we're using the ISE
    If ($Host.Name -ne "Windows PowerShell ISE Host") {
        $Msg = "This script requires the PowerShell ISE; current  host is '$($Host.Name)'"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }
}
Process{
    
    # Make sure the file is ok.... 
    Try {
        If ([byte[]]$bytes = Get-Content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $SessionFile -ErrorAction SilentlyContinue){
                    
            Switch -regex ('{0:x2}{1:x2}{2:x2}{3:x2}' -f $bytes[0],$bytes[1],$bytes[2],$bytes[3]) {
                '^efbbbf'   { $Type = 'UTF8' }
                '^2b2f76'   { $Type = 'UTF7' }
                '^fffe'     { $Type = 'Unicode' }
                '^feff'     { $Type = 'BigendianUnicode' }
                '^0000feff' { $Type = 'UTF32' }
                default     { $Type = 'ASCII' }
            }
            If ($Type -eq "ASCII") {
                $Msg = "Can't import invalid file type '$SessionFile'"
                $Host.UI.WriteErrorLine($Msg)
                Break
            }
            Else {

                $Msg = "Open $((Get-Content $SessionFile).Split("/").Count) saved tab(s) from $Type file $SessionFile"
                Write-Verbose $Msg

                $ConfirmMsg = "`n`n`t$Msg`n`n"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
                    Try {
                        
                        [array]$FileNames = ( (Get-Content $SessionFile | Where-Object {$_} ) -split('/')) -replace('"','') 
                        
                        #Verify files exist
                        [array]$Reachable = ($FileNames | Where-Object { (Test-Path $_)})
                        [array]$Unreachable = ($FileNames | Where-Object {-not (Test-Path $_)})
                        
                        If ($Unreachable.Count -gt 0) {
                            $Msg = "File $SessionFile contains $($Unreachable.Count) unreachable file(s) that will not be opened:`n$($Unreachable | Format-List | Out-String)"
                            Write-Warning $Msg
                        }

                        $ToImport = """$($Reachable -join(","))"""

                        Invoke-Expression ("ise $ToImport") @StdParams
                    }
                    Catch {
                        $Msg = "Operation failed"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                        $Host.UI.WriteErrorLine($Msg)
                    }
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    $Host.UI.WriteErrorLine($Msg)
                }

            } #end if valid content type

        } # end get content

        Else {
            $Msg = "No content available for $SessionFile"
            $Host.UI.WriteErrorLine($Msg)
        }
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine($Msg)

    }
}
} #end Restore-PKISESession


