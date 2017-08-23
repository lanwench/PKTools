#requires -Version 3
Function Restore-PKISESession {
<#
.SYNOPSIS 
    Restores tabs/files from file saved using Save-PKISESession

.DESCRIPTION
    Restores tabs/files from file saved using Save-PKISESession
    
.NOTES
    Name    : Function_Restore-PKISESession.ps1
    Author  : Paula Kingsley
    Version : 1.0.0
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v1.0.0 - 2016-05-29 - Created script based on links
        
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
        HelpMessage = "Full path to text file (created using Save-PKISESession)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (Test-Path $_) {$True}})]
    [String]$ImportPath = "$env:temp\ISESession.txt"
)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "1.0.0"

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

    Try {
        If ([byte[]]$bytes = Get-Content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $ImportPath -ErrorAction SilentlyContinue){
                    
            Switch -regex ('{0:x2}{1:x2}{2:x2}{3:x2}' -f $bytes[0],$bytes[1],$bytes[2],$bytes[3]) {
                '^efbbbf'   { $Type = 'UTF8' }
                '^2b2f76'   { $Type = 'UTF7' }
                '^fffe'     { $Type = 'Unicode' }
                '^feff'     { $Type = 'BigendianUnicode' }
                '^0000feff' { $Type = 'UTF32' }
                default     { $Type = 'ASCII' }
            }
            If ($Type -eq "ASCII") {
                $Msg = "Can't import invalid file type '$SavePath'"
                $Host.UI.WriteErrorLine($Msg)
                Break
            }
            Else {

                $Msg = "Open $((Get-Content $ImportPath).Split(",").Count) saved tab(s) from $Type file $ImportPath"
                Write-Verbose $Msg

                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$Msg)) {
                    Try {
                        Invoke-Expression (Get-Content $ImportPath) @StdParams
                    }
                    Catch {
                        $Msg = $_.Exception.Message
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
            $Msg = "No content available for $ImportPath"
            $Host.UI.WriteErrorLine($Msg)
        }
    }
    Catch {
        $Msg = $_.Exception.Message
        $Host.UI.WriteErrorLine($Msg)
    }
}
} #end Restore-PKISESession


