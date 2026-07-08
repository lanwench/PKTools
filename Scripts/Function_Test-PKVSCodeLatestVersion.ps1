#requires -version 4
Function Test-PKVSCodePortableVersion {
    <#
.SYNOPSIS
    Gets the latest VSCode portable version available at code.visualstudio.com and compares it to the current local version

.DESCRIPTION
    Gets the latest VSCode portable version available at code.visualstudio.com
    Compares it to the locally installed version unless -SkipCurrentVersionCheck is specified
    Custom URI and local path are configurable
    Windows only

.NOTES
    Name    : Function_Test-PKVSCodePortableVersion.ps1
    Created : 2025-03-03
    Author  : Paula Kingsley
    Version : 02.00
    History :

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00 - 2025-03-03 - Created script
        v02.00 - 2026-06-17 - Renamed, cosmetic updates; converted version format to major.minor
        
.LINK
    https://stackoverflow.com/questions/25125818/powershell-invoke-webrequest-how-to-automatically-use-original-file-name

.LINK 
    https://sqladm.in/posts/update-portable-vscode/   

.PARAMETER URI
    URI for portable VSCode zip file download; change at your peril!

.PARAMETER SkipCurrentVersionCheck
    Don't look for current version of VSCode Portable for comparison    

.PARAMETER CurrentPath
    Current folder for VSCode Portable for version comparison (used if -Skip not specified; default is $Home\VSCode)    

.EXAMPLE 
    PS C:\> Test-PKVSCodePortableVersion
    Looks for local version of VSCode Portable in $Home\VSCode, then checks for latest version on the web, returning comparison results

.EXAMPLE 
    PS C:\> Test-PKVSCodePortableVersion-SkipCurrentVersionCheck
    Checks for latest version of VSCode Portable on the web

.EXAMPLE
    PS C:\> Test-PKVSCodePortableVersion-CurrentPath "D:\VSCode" -URI "https://example.com/foobar"
    Checks for latest version of VSCode Portable on the web using a custom URI, and checking the specified path for an existing version for comparison        

#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High",DefaultParameterSetName = "Default")]
Param(
    
    [Parameter(
        HelpMessage = "URI for portable VSCode zip file download; change at your peril!"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$URI = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-archive",

    [Parameter(
        ParameterSetName = "Skip",
        HelpMessage = "Don't look for current version of VSCode Portable for comparison"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$SkipCurrentVersionCheck,

    [Parameter(
        HelpMessage = "Current folder for VSCode Portable for version comparison (used if -Skip not specified; default is `$Home\VSCode)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$CurrentPath = "$Home\VSCode"
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01"

    # How did we get here?
    $ScriptName = $MyInvocation.MyCommand.Name
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput

    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
    Where-Object { Test-Path Variable:$_ } | Foreach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ParameterSetName", $Source)
    $CurrentParams.Add("ScriptVersion", $Version)
    $CurrentParams.Add("PipelineInput", $PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    If ($IsWindows -eq $False) {
        $Msg = "This function requires Windows"
        Write-Error $Msg
        Return
    }

    $ComputerName = &hostname
    #region Inner functions
    
    # find the current filename
    Function _GetRedirectedURL {
        Param ([Parameter(Mandatory = $true)][String]$URL)
        $request = [System.Net.WebRequest]::Create($url)
        $request.AllowAutoRedirect = $false
        $response = $request.GetResponse()
        If ($response.StatusCode -eq "Found") {$response.GetResponseHeader("Location")}
        Else {Write-Warning "Failed to get filename from $URL"; $False}
    }

    #endregion Inner functions

    $Activity = "Look for latest available version of VSCode Portable"
    Write-Verbose "[BEGIN: $ScriptName] $Activity"

} #end begin

Process {
    
    [switch]$IsInstalled = $False
    If ($Source -ne "Skip") {
        Try {
            $Msg =  "Looking for existing VSCode Portable executable in $CurrentPath" 
            Write-Host "[$ComputerName] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg
            If ($CurrentFile = Get-Item "$CurrentPath\code.exe" -ErrorAction SilentlyContinue) {
                [version]$CurrentVersion = $CurrentFile.VersionInfo.FileVersionRaw
                $Msg = "VSCode Portable $CurrentVersion found in $CurrentPath"
                Write-Host "[$ComputerName] $Msg" -ForegroundColor Cyan
                $IsInstalled = $True
            }
            Else {
                $Msg = "No current version of VSCode detected in $CurrentPath"
                Write-Host "[$ComputerName] $Msg"
            }
        }
        Catch {
            $Msg = "Failed to look for current version of VSCode Portable in $CurrentPath"
            Write-Warning "[$ComputerName] $Msg"
        }
    }

    Try {
        $Msg =  "Looking for latest version available of VSCode Portable on $URI" 
        Write-Host "[$ComputerName] $Msg"
        Write-Progress -Activity $Activity -CurrentOperation $Msg
        $NewFileName = [System.IO.Path]::GetFileName((_GetRedirectedURL $URI -ErrorAction Stop)) | Where-Object {$_ -match "^VSCode-Win32-x64"}

        # Parse version from filename
        If ([version]$NewVersion = [regex]::Matches($NewFileName, "(\d+\.)?(\d+\.)?(\*|\d+)").value | Where-Object { $_ -match '(.*?(?=\.\d))\.(.+)' }) {  
            # Compare versions 
            If ($IsInstalled.IsPresent) {
                If (-not ($Newversion -gt $CurrentVersion)) {
                    $Msg = "You're currently running the latest VSCode Portable version, $CurrentVersion"
                    Write-Host "[$ComputerName] $Msg" -ForegroundColor Cyan
                }
                Else {
                    $Msg = "Newer version VSCode Portable $NewVersion is available!"
                    Write-Host "[$ComputerName] $Msg" -ForegroundColor Green
                }
            }
            Else {
                $Msg = "VSCode Portable $NewVersion is available!"
                Write-Host "[$ComputerName] $Msg" -ForegroundColor Cyan
            }
        }
        Else {
            $Msg = "No matching file found! Please verify your URI."
            Write-Host "[$ComputerName] $Msg" -ForegroundColor Red
        }
    }
    Catch {
        $Msg = "Failed to get latest version of VSCode Portable from $URI"
        Write-Warning "[$ComputerName] $Msg"
    }

} #end process

End {
    Write-Progress -Activity * -Completed 
    Write-Verbose "[END: $ScriptName] Script ran successfully"
}
} #end Install-PKVSCodePortable


$Null = New-Alias -Name Get-PKVSCodeLatestVersion -Value Test-PKVSCodePortableVersion -Force:$True -ErrorAction SilentlyContinue