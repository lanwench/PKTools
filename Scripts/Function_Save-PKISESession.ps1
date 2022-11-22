#requires -version 4
Function Save-PKISESession {
<#
.SYNOPSIS 
    Saves open tabs in current ISE session to a file

.DESCRIPTION
    Saves open tabs in current ISE session to a file
    
.NOTES
    Name    : Function_Save-PKISESession.ps1
    Created : 2016-05-29
    Author  : Paula Kingsley
    Version : 03.00.000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v1.0.0      - 2016-05-29 - Created script based on links
        v02.00.0000 - 2018-02-14 - Updated with default path, standardization, changed force to noclobber
        v02.01.0000 - 2019-10-08 - Minor cosmetic updates
        v03.00.0000 - 2022-10-10 - Cosmetic updates & standardization
        
.LINK
    https://itfordummies.net/2014/10/27/save-restore-powershell-ise-opened-scripts/

.LINK
    https://stackoverflow.com/questions/3710374/get-encoding-of-a-file-in-windows   

.EXAMPLE
    PS C:\> Save-PKISESession -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                             
        ---           -----                             
        Verbose       True                              
        SessionFile   C:\Users\gkravitz\PSISESession.txt
        Delimiter     |                                 
        NoClobber     False                             
        ScriptName    Save-PKISESession                 
        ScriptFile                                      
        ScriptVersion 3.0.0                             

        VERBOSE: [BEGIN:] Save-PKISESession Save open ISE session tabs to file
        VERBOSE: [C:\Users\gkravitz\PSISESession.txt] Looking for existing file object
        VERBOSE: [C:\Users\gkravitz\PSISESession.txt] Existing file found (last saved 2022-10-10 17:13:23Z)
        VERBOSE: [C:\Users\gkravitz\PSISESession.txt] Verifying compatible file type
        VERBOSE: [C:\Users\gkravitz\PSISESession.txt] Saving 7 current tab(s) to session file:
	        * C:\Users\gkravitz\git\Scripts\Untitled1.ps1
	        * C:\Users\gkravitz\git\Personal\HelperModule\Scripts\Function_AllFunctions.ps1
	        * C:\Users\gkravitz\git\Personal\HelperModule\Scripts\Untitled2.ps1
	        * C:\Users\gkravitz\git\WorkModules\Tools\Scripts\Function_Get-PKFunctionInfo.ps1
	        * C:\Users\gkravitz\git\WorkModules\Tools\Scripts\Function_Restore-PKISESession.ps1
	        * C:\Users\gkravitz\git\WorkModules\Tools\Scripts\Function_Save-PKISESession.ps1
	        * C:\Users\gkravitz\git\Personal\Profiles\ise-prof.ps1

        VERBOSE: [C:\Users\gkravitz\PSISESession.txt] 7 tab(s) saved to session file; use Restore-PKISesession to recover
        VERBOSE: [END:] Save-PKISESession Save open ISE session tabs to file

#>

[CmdletBinding(
    SupportsShouldProcess=$True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Position=0,
        HelpMessage = "Output file for session storage (will be created if nonexistent)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (Test-Path $_ -PathType Leaf -IsValid -ErrorAction SilentlyContinue) {$True}})]
    [String]$SessionFile = "$Home\PSISESession.txt",

    [Parameter(
        HelpMessage = "Character to parse filenames in text file; default is '|' as this character is guaranteed not to match a valid file path"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Delimiter = "|",

    [Parameter(
        HelpMessage = "Don't overwrite an existing file"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$NoClobber

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "03.00.0000"

   # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    $CurrentParams = $PSBoundParameters
    $ScriptName = $MyInvocation.MyCommand.Name
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path Variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptFile",$ScriptFile)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
   
    # Make sure we're using the ISE
    If ($Host.Name -ne "Windows PowerShell ISE Host") {
        $Msg = "This script requires the PowerShell ISE; current  host is '$($Host.Name)'"
        Throw $Msg
        Break
    }

    # Function to test file type
    Function TestFileType {
        [Cmdletbinding()]
        Param([Parameter(ValueFromPipeline,Position=0)]$Item)
        If ([byte[]]$bytes = Get-Content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $FileObj.FullName -ErrorAction SilentlyContinue){
            Switch -regex ('{0:x2}{1:x2}{2:x2}{3:x2}' -f $bytes[0],$bytes[1],$bytes[2],$bytes[3]) {
                '^efbbbf'   {'UTF8'}
                '^2b2f76'   {'UTF7'}
                '^fffe'     {'Unicode'}
                '^feff'     {'BigendianUnicode'}
                '^0000feff' {'UTF32'}
                default     {'ASCII'}
            }
        }
        Else {
            Write-Warning "Failed to get bytes for file type check"
        }
    }

    $Activity = "Save open ISE session tabs to file"
    Write-Verbose "[BEGIN:] $ScriptName $Activity"

}
Process{
    
    [switch]$Continue = $False

    Try {
        $Msg = "Looking for existing file object"
        Write-Verbose "[$SessionFile] $Msg"
        Write-Progress -Activity $Activity -CurrentOperation $Msg
        If (($FileObj = Get-Item $SessionFile -ErrorAction SilentlyContinue) -and (-not $FileObj.PSIsContainer)) {

            $Msg = "Existing file found (last saved $(Get-Date ($FileObj.LastWriteTime) -f u))"
            If ($NoClobber.IsPresent) {
                $Msg += " ; -NoClobber specified"
                Write-Warning "[$SessionFile] $Msg"
            }
            Else {
                Write-Verbose "[$SessionFile] $Msg"
                $Msg = "Verifying compatible file type"
                Write-Verbose "[$SessionFile] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg
                $Type = TestFileType -Item $FileObj
            
                If ($Type -eq "ASCII") {
                    $Msg = "Invalid file type '$Type'; cannot overwrite"
                    Write-Warning "[$SessionFile] $Msg"
                }
                Else {$Continue = $True}
            }
        } # end if file found
        Else {
            $Msg = "No current file object found"
            Write-Verbose "[$SessionFile] $Msg"
            $Continue = $True
        }
    }
    Catch {
        $Msg = "File lookup operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
        Write-Warning "[$SessionFile] $Msg"
    }

    If ($Continue.IsPresent) {
        
        $Msg = "Saving $($psISE.PowerShellTabs.Files.Count) current tab(s) to UTF8-formatted session file"
        Write-Progress -Activity $Activity -CurrentOperation $Msg

        $Tabs = "`t* $(($psISE.PowerShellTabs.Files.FullPath | Sort-Object) -join("`n`t* "))`n"
        $Msg = "Saving $($psISE.PowerShellTabs.Files.Count) current tab(s) to UTF8-formatted session file:`n$($Tabs | Format-List | Out-String)"
        Write-Verbose "[$SessionFile] $Msg"
        
        $Msg = "Save $($psISE.PowerShellTabs.Files.Count) current tab(s) to UTF8-formatted session file:`n$($Tabs | Format-List | Out-String)"
        If ($PSCmdlet.ShouldProcess($SessionFile,$Msg)) {
            Try {
                $psISE.CurrentPowerShellTab.Files | Foreach-Object {$_.SaveAs($_.FullPath)}
            
                # Join with custom delimiter (we may have commas or semicolons in the filenames)
                $($psISE.PowerShellTabs.Files.FullPath -join$Delimiter) | Out-File -Encoding UTF8 -FilePath $SessionFile -Force -Confirm:$False -ErrorAction Stop -Verbose:$False

                $Msg = "$($psISE.CurrentPowerShellTab.Files.count) tab(s) saved to session file; use Restore-PKISesession to recover"
                Write-Verbose "[$SessionFile] $Msg"
            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                Write-Warning "[$SessionFile] $Msg"
            }
        }
        Else {
            $Msg = "Operation cancelled by user"
            Write-Verbose "[$SessionFile] $Msg"
        }
    } #end if continue
}
End {

    Write-Verbose "[END:] $ScriptName $Activity"
    $Null = Write-Progress -Activity * -Completed

}
} #end Save-PKISESession






    
