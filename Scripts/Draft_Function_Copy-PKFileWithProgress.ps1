#requires -version 3.0
Function Copy-PKFileWithProgress {
<#
.SYNOPSIS
    Copies one or more items from source to destination, displaying a progress bar

.DESCRIPTION
    Copies one or more items from source to destination, displaying a progress bar

.NOTES 
    Name    : Function_Copy-PKFileWithProgress.ps1
    Author  : Paula Kingsley
    Version : 1.0.0
    Created : 2017-06-01
    History :  
        
        ** PLEASE KEPEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v1.0.0 - 2017-06-01 - Created script based on original (see link)

.LINK
    https://blogs.technet.microsoft.com/heyscriptingguy/2015/12/20/build-a-better-copy-item-cmdlet-2/

.EXAMPLE
    PS C:\> Copy-PKFileWithProgress

#>
 [CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
 )]
 Param(
  [Parameter(
        Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        Position=0
    )]
    [ValidateScript({Test-Path $_})]
    [string]$SourcePath,

    [Parameter(
        Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        Position=1
    )]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath

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

    # For console output 
    $BGColor = $Host.UI.RawUI.BackgroundColor

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

}
 Process {
    
    $AllCopiedFiles = @()
    $AllSkippedFiles = @()

    If (-not ($Null = Test-Path $DestinationPath)) {
        $Msg = "Create directory '$DestinationPath'"
        Write-Verbose $Msg
        If ($PSCmdlet.ShouldProcess($Null,$Msg)) {
            Try {
                $Null = New-Item -Path $DestinationPath -ItemType Directory -Force @StdParams
            }
            Catch {
                $Msg = "Directory creation failed for '$DestinationPath'"
                $ErrorDetails = $_.Exception.Message
                $Host.UI.WriteErrorLine("ERROR: $Msg`n $ErrorDetails")
                Break
            }
        }
        Else {
            $Msg = "Directory creation cancelled by user"
            $Host.UI.WriteErrorLine("$Msg")
            Break
        }
    }

    Try {

        #If ($SourcePath -notlike "*\") {$SourcePath = "$SourcePath\"}
        #If ($DestinationPath -notlike "*\") {$DestinationPath = "$DestinationPath\"}


        $Filelist = Get-Childitem $SourcePath -File -Exclude ".DS_Store" -Attributes !hidden –Recurse -ErrorAction Stop
        $Total = $Filelist.count
        $TotalSize = "$([math]::round(($FileList | Measure-Object -Property Length -Sum).Sum/1MB)) MB"

        $Msg = "Copy $Total files ($TotalSize) from '$SourcePath' to '$DestinationPath'"
        Write-Verbose $Msg

        If ($PSCmdlet.ShouldProcess($Null,$Msg)) {
        #If ($Foo.IsPresent) {
            $Current = 0
            $SW = [system.diagnostics.stopwatch]::StartNew()

            Foreach ( $File in ($Filelist | Where-Object {($_ -is [System.IO.FileInfo])} )) {
            
                $Current ++
                $FileName = $File.Name
                $DestinationFile = "$DestinationPath\$Filename"
                $Size = "$([math]::round($File.Length/1MB)) MB"
            
                Write-Verbose $FileName

                $Activity = "Copying data from $SourcePath to $DestinationFile"
                $Percent = (($Current/$total)*100)
                $Status = "$Percent % [$Current / $Total]: $($File.Name) ($Size)"

                Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent
                If (-not (Test-Path $DestinationFile -ErrorAction SilentlyContinue -Verbose:$False)) {
                    $AllCopiedFiles += Copy-Item $File.FullName -Destination $DestinationPath -Recurse -Force -PassThru -ErrorAction Stop
                }
                Else {
                    $AllSkippedFiles += $File
                }
            } #end foreach

            $SW.Stop()
        
            $Host.UI.WriteLine()
            If ($AllSkippedFiles.Count -gt 0) {
                $Msg = "$($AllSkippedFiles.Count) existing target file(s) skipped:`n`n$($AllSkippedFiles.FullName | Format-Table -AutoSize | Out-String)"
                Write-Verbose $Msg
            }
            If ($AllCopiedFiles.Count -gt 0) {

                $Msg = "Copied $($AllCopiedFiles.Count) from $SourcePath to $DestinationPath in $($SW.Elapsed.ToString())"

                #$Msg = "Copied $($AllCopiedFiles.Count) from $SourcePath to $DestinationPath in  in $($SW.Elapsed.Hours)h, $($SW.Elapsed.Minutes)m, $($SW.Elapsed.Seconds)s"
                #$Msg = "$($AllCopiedFiles.Count) file(s) copied from $SourcePath to $DestinationPath"
                Write-Verbose $Msg

                $Msg = $AllCopiedFiles.FullName
                Write-Output $Msg
            }
            Else {
                $Msg = "No files copied"
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Else {
            $Msg = "File copy cancelled by user"
            $Host.UI.WriteErrorLine("$Msg")
        }
    }
    Catch {
        $Msg = "Operation failed"
        $ErrorDetails = $_.Exception.Message
        $Host.UI.WriteErrorLine("ERROR: $Msg; $ErrorDetails")
        Break
    }

 }

 } #end Copy-PKFileWithProgress


 # 
 $exclude = @("main.js")
$excludeMatch = @("app1", "app2", "app3")
[regex] $excludeMatchRegEx = ‘(?i)‘ + (($excludeMatch |foreach {[regex]::escape($_)}) –join “|”) + ‘’
Get-ChildItem -Path $from -Recurse -Exclude $exclude | 
 where { $excludeMatch -eq $null -or $_.FullName.Replace($from, "") -notmatch $excludeMatchRegEx} |
 Copy-Item -Destination {
  if ($_.PSIsContainer) {
   Join-Path $to $_.Parent.FullName.Substring($from.length)
  } else {
   Join-Path $to $_.FullName.Substring($from.length)
  }
 } -Force -Exclude $exclude




$exclude = @("main.js")
$excludeMatch = @("app1", "app2", "app3")
[regex] $excludeMatchRegEx = ‘(?i)‘ + (($excludeMatch |foreach {[regex]::escape($_)}) –join “|”) + ‘’


$from = 'c:\sources'
$to = 'c:\build'


$AllCopiedFiles = @()
$DestinationPath = Get-Item $DestinationPath

Foreach ($File in $FileList) {

    Write-Verbose $File.FullName
    If ($File.PSIsContainer) {
        $Destination = Join-Path $SourcePath $File.Parent.FullName.Substring($DestinationPath.length)
    } 
    Else {
        $Destination = Join-Path $SourcePath $File.FullName.Substring($DestinationPath.length)
    }

    $AllCopiedFiles += Copy-Item $File -Destination $Destination -Force -PassThru
}