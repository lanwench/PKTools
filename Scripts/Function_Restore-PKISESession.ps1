#requires -Version 4
Function Restore-PKISESession {
<#
.SYNOPSIS 
    Restores tabs/files from text file created using Save-PKISESession

.DESCRIPTION
    Restores tabs/files from text file created using Save-PKISESession
    Works only in ISE, because of course it does
    
.NOTES
    Name    : Function_Restore-PKISESession.ps1
    Created : 2016-05-29
    Author  : Paula Kingsley
    Version : 03.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2016-05-29 - Created script based on links
        v02.00.0000 - 2018-02-14 - Updated/made consistent with new Save-PKISESession
        v02.01.0000 - 2019-10-08 - Minor cosmetic updates
        v03.00.0000 - 2022-10-10 - Mainly cosmetic updates/standardization
        
.LINK
    https://itfordummies.net/2014/10/27/save-restore-powershell-ise-opened-scripts/

.LINK
    https://stackoverflow.com/questions/3710374/get-encoding-of-a-file-in-windows   

.EXAMPLE
    PS C:\> Restore-PKISESession -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                             
        ---           -----                             
        Verbose       True                              
        SessionFile   C:\Users\kipa7003\PSISESession.txt
        Delimiter     |                                 
        ScriptName    Restore-PKISESession              
        ScriptFile                                      
        ScriptVersion 3.0.0                             

        VERBOSE: [BEGIN:] Restore-PKISESession Restore ISE session tabs from file
        VERBOSE: [C:\Users\kipa7003\PSISESession.txt] Getting file object
        VERBOSE: [C:\Users\kipa7003\PSISESession.txt] File last saved 2022-10-10 17:28:14Z
        VERBOSE: [C:\Users\kipa7003\PSISESession.txt] Verifying compatible file type
        VERBOSE: [C:\Users\kipa7003\PSISESession.txt] Getting content from UTF8 file
        VERBOSE: [C:\Users\kipa7003\PSISESession.txt] 7 file name(s) found; testing path
        VERBOSE: [C:\Users\kipa7003\PSISESession.txt] 7 valid/reachable file path(s)s found:
	        * C:\Users\gkravitz\git\Scripts\Untitled1.ps1
	        * C:\Users\gkravitz\git\Personal\HelperModule\Scripts\Function_AllFunctions.ps1
	        * C:\Users\gkravitz\git\Personal\HelperModule\Scripts\Untitled2.ps1
	        * C:\Users\gkravitz\git\WorkModules\Tools\Scripts\Function_Get-PKFunctionInfo.ps1
	        * C:\Users\gkravitz\git\WorkModules\Tools\Scripts\Function_Restore-PKISESession.ps1
	        * C:\Users\gkravitz\git\WorkModules\Tools\Scripts\Function_Save-PKISESession.ps1
	        * C:\Users\gkravitz\git\Personal\Profiles\ise-prof.ps1
        
        VERBOSE: [C:\Users\kipa7003\PSISESession.txt] Importing 7 saved tab(s) from file


            


#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Position=0,
        ValueFromPipeline = $True,
        HelpMessage = "Full path to text file (created using Save-PKISESession)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (Test-Path -PathType Leaf $_ -ErrorAction SilentlyContinue) {$True}})]
    [String]$SessionFile = "$Home\PSISESession.txt",

    [Parameter(
        HelpMessage = "Character to parse filenames in text file; default is '|' as this character is guaranteed not to match a valid file path"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Delimiter = "|"


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

    $Activity = "Restore ISE session tabs from file"
    Write-Verbose "[BEGIN:] $ScriptName $Activity"
}
Process{
    
    Try {
        $Msg = "Getting file object"
        Write-Verbose "[$SessionFile] $Msg"
        Write-Progress -Activity $Activity -CurrentOperation $Msg
        $FileObj = Get-Item $SessionFile -ErrorAction Stop

        $Msg = "File last saved $(Get-Date ($FileObj.LastWriteTime) -f u)"
        Write-Verbose "[$SessionFile] $Msg"

        $Msg = "Verifying compatible file type"
        Write-Verbose "[$SessionFile] $Msg"
        Write-Progress -Activity $Activity -CurrentOperation $Msg
        
        $Type = TestFileType -Item $FileObj
            
        If ($Type -eq "ASCII") {
            $Msg = "Invalid file type '$Type'; please create a new UTF/unicode file using Save-PKISESession"
            Write-Warning "[$SessionFile] $Msg"
        }
        Else {
            $Msg = "Getting content from $Type file"
            Write-Verbose "[$SessionFile] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg

            [string[]]$FileNames = (((Get-Content -path $FileObj.FullName | Where-Object {$_} ).split($Delimiter)) -replace('"','') | Sort-Object)

            $Msg = "$($FileNames.Count) file name(s) found; testing path"
            Write-Verbose "[$SessionFile] $Msg"
            [string[]]$Valid = ($FileNames | Where-Object { (Test-Path $_)})
            [string[]]$Invalid = ($FileNames | Where-Object {-not (Test-Path $_)})

            If ($Invalid.Count -gt 0) {
                $InvalidPaths = "`t* $(($Invalid | Sort-Object) -join("`n`t* "))`n"
                $Msg = "File contains $($Invalid.Count) unreachable file(s) that will not be opened:`n$($InvalidPaths | Format-List | Out-String)"
                Write-Warning "[$SessionFile] $Msg"
            }

            If ($Valid.count -gt 0) {
                
                $ValidPaths = "`t* $(($Valid | Sort-Object) -join("`n`t* "))`n"
                $Msg = "$($Valid.Count) valid/reachable file path(s)s found:`n$($ValidPaths | Format-List | Out-String)"
                Write-Verbose "[$SessionFile] $Msg"
                    
                $Msg = "Importing $($Valid.Count) saved tab(s) from file"
                Write-Verbose "[$SessionFile] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg

                $Msg = "Import $($Valid.Count) saved tab(s) from file:`n$($ValidPaths | Format-List | Out-String)"
                If ($PSCmdlet.ShouldProcess($SessionFile,$Msg)) {
                    Try {
                        $ToImport = """$($Valid -join(","))"""
                        Invoke-Expression ("ise $ToImport") -ErrorAction Stop

                        If (-not ($Valid | Where-Object {$psISE.PowerShellTabs.Files.FullPath -notcontains $_})) {
                            $Msg = "Successfully imported tabs"
                            Write-Verbose "[$SessionFile] $Msg"
                        }                    }
                    Catch {
                        $Msg = "Failed to import tabs"
                        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                        Write-Warning "[$SessionFile] $Msg"
                    }
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    Write-Verbose "[$SessionFile] $Msg"
                }
            } # end if valid file paths found
        } #end if valid content type
    }
    Catch {
        $Msg = "Failed to get file object"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
        Write-Warning "[$SessionFile] $Msg"
    }

}
} #end Restore-PKISESession


