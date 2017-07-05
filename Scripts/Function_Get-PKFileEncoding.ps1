function Get-PKFileEncoding {
<#
.SYNOPSIS
    Returns the encoding type for one or more files

.DESCRIPTION
    Returns the encoding types for one or more files
    Accepts pipeline input
    Returns a PSObject or strin

.NOTES        
    Name    : Function_Get-PKFileEncoding.ps1
    Version : 1.0.0
    Author  : Paula Kingsley
    Created : 2017-07-05

    History:

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v1.0.0 - 2017-07-05 - Created script based on yzorg's StackOverflow answer

.LINK
    https://stackoverflow.com/questions/3710374/get-encoding-of-a-file-in-windows

.EXAMPLE
    $ Get-PKFileEncoding -FilePath $Arr -OutputType TypeOnly -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                                                                                                                             
        ---           -----                                                                                                                                             
        FilePath      {C:\Temp\chef-client-12.18.31-1-x64.msi, C:\Distrib\Filezilla\FileZillaPortable_3.8.0.paf.exe, C:\Users\jbloggs\Tracing\LyncUninstall101253.log}
        OutputType    TypeOnly                                                                                                                                          
        Verbose       True                                                                                                                                              
        ScriptName    Get-PKFileEncoding                                                                                                                                
        ScriptVersion 1.0.0                                                                                                                                             

        VERBOSE: C:\Temp\chef-client-12.18.31-1-x64.msi
        ASCII
        VERBOSE: C:\Distrib\Filezilla\FileZillaPortable_3.8.0.paf.exe
        ASCII
        VERBOSE: C:\Users\jbloggs\Tracing\LyncUninstall101253.log
        Unicode

.EXAMPLE
    PS C:\> (Get-ChildItem C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory -Recurse).FullName | Get-PKFileEncoding

        VERBOSE: PSBoundParameters: 
	
        Key           Value             
        ---           -----             
        FilePath                        
        OutputType    Full              
        ScriptName    Get-PKFileEncoding
        ScriptVersion 1.0.0             

        VERBOSE: OpsActiveDirectory.psd1
        VERBOSE: OpsActiveDirectory.psm1
        VERBOSE: README.md
        VERBOSE: Backup_Function_New-OpsADSecurityGroups.ps1
        VERBOSE: Draft_Function_Get-OpsADComputerOU.ps1
        VERBOSE: Draft_Function_Get-OpsADDNSForwarders.ps1
        VERBOSE: Draft_Function_Get-OpsADSite.ps1
        VERBOSE: Draft_Function_Get-OpsADUserLockout.ps1
        VERBOSE: Draft_Function_Get-OpsADUserLockoutReport.ps1

        Name                                            Type    FullName                                                                                                                               
        ----                                            ----    --------                                                                                                                               
        OpsActiveDirectory.psd1                         Unicode C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\OpsActiveDirectory.psd1                              
        OpsActiveDirectory.psm1                         UTF8    C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\OpsActiveDirectory.psm1                              
        README.md                                       ASCII   C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\README.md                                              
        Backup_Function_New-OpsADSecurityGroups.ps1     UTF8    C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\Scripts\Backup_Function_New-OpsADSecurityGroups.ps1  
        Draft_Function_Get-OpsADComputerOU.ps1          UTF8    C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\Scripts\Draft_Function_Get-OpsADComputerOU.ps1          
        Draft_Function_Get-OpsADDNSForwarders.ps1       UTF8    C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\Scripts\Draft_Function_Get-OpsADDNSForwarders.ps1       
        Draft_Function_Get-OpsADSite.ps1                UTF8    C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\Scripts\Draft_Function_Get-OpsADSite.ps1             
        Draft_Function_Get-OpsADUserLockout.ps1         UTF8    C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\Scripts\Draft_Function_Get-OpsADUserLockout.ps1      
        Draft_Function_Get-OpsADUserLockoutReport.ps1   UTF8    C:\users\jbloggs\git\Corp\Infrastructure\PowerShell\OpsActiveDirectory\Scripts\Draft_Function_Get-OpsADUserLockoutReport.ps1

#>


[CmdletBinding()]
Param(
    [Parameter(
        Position=0,
        Mandatory=$True,
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({If (Test-Path $_) {$True}})]
    [Alias("Name","FullName")]
    [string[]]$FilePath,

    [Parameter(
        HelpMessage = "Display full output with file path, or just type"
    )]
    [ValidateSet("Full","TypeOnly")]
    [string]$OutputType = "Full"

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
    
    $Activity = "Get file encoding"
    $Results = @()
}
    
Process {
        
    $Total = $FilePath.Count
    $Current = 0

    Foreach ($File in $FilePath) {
        
        If (-not ($Obj = Get-Item $File -ErrorAction SilentlyContinue).PSIsContainer) {
            
            Try {
                $Current ++
                Write-Progress -Activity $Activity -CurrentOperation $Obj.FullName -Percentcomplete ($Current / $Total * 100)
                
                If ($OutputType -eq "Typeonly") {$Msg = $Obj.FullName}
                Else {$Msg = $Obj.Name}
                    
                Write-Verbose $Msg

                If ($Obj -is [system.io.fileinfo]) {
 
                    Try {
                        If ($Bytes = [byte[]](Get-Content -Encoding Byte -ReadCount 4 -TotalCount 4 -Path $Obj -ErrorAction SilentlyContinue)) {
                            Switch -regex ('{0:x2}{1:x2}{2:x2}{3:x2}' -f $Bytes[0],$Bytes[1],$Bytes[2],$Bytes[3]) {
                                '^efbbbf'   { $Type = 'UTF8' }
                                '^2b2f76'   { $Type = 'UTF7' }
                                '^fffe'     { $Type = 'Unicode' }
                                '^feff'     { $Type = 'BigendianUnicode' }
                                '^0000feff' { $Type = 'UTF32' }
                                default     { $Type = 'ASCII' }
                            }
                            $Output = New-Object PSObject -Property ([ordered] @{
                                Name     = $Obj.Name
                                Type     = $Type  
                                FullName = $Obj.FullName
                            })

                            If ($OutputType -eq "Typeonly") {Write-Output $Output.Type} #return in pipeline
                            Else {$Results += $Output} # will add up and return at the end
                        }
                        Else {
                            $Msg = "No content available for $($Obj.FullName)"
                            Write-Warning $Msg
                        }
                    }
                    Catch {
                        $Msg = "Get-Content failed for '$($Obj.FullName)'"
                        $ErrorDetails = $_.Exception.Message
                        $Host.UI.WriteWarningLine("$Msg; $ErrorDetails")
                    }
                }
                Else {Write-Warning "Skipping directory $Obj"}
            }
            Catch {}

        } # end if not directory

    } #end for each

}
End {
    Write-Progress -Activity $Activity -Completed
    If ($Results) {Write-Output $Results}
}
}
