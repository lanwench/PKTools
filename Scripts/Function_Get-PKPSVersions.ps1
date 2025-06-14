#required -Version 3
Function Get-PKPSVersions {
<#
.SYNOPSIS
    Retrieves the installed versions and paths of Windows PowerShell and PowerShell Core on the local computer.

.DESCRIPTION
    The Get-PKPSVersions function scans the local computer for installations of Windows PowerShell (powershell.exe) and PowerShell Core (pwsh.exe).
    Sure, it's kind of silly, but it was created as an exercise when trying to interrogate remote computers & there's no straightforward way to 
    tet this info via the registry, so I decided to keep it just for no reason, really.
    It returns a list of objects containing the computer name, PowerShell version, and the executable path for each discovered instance.

.PARAMETER PSCorePaths
    One or more paths to search for PowerShell Core (pwsh.exe) installations; default are the standard installation directories for PowerShell Core

.OUTPUTS
    [PSCustomObject]
    Returns one or more custom objects with the following properties:
        - ComputerName: The name of the local computer.
        - Version:      The version of PowerShell found.
        - Path:         The full path to the PowerShell executable.

.EXAMPLE
    PS C:\> Get-PKPSVersions

    Lists all installed versions of Windows PowerShell and PowerShell Core on the local machine.

.NOTES        
    Name    : Function_Get-PKPSVersions.ps1
    Created : 2025-05-22
    Author  : Paula Kingsley
    Version : 012.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2025-05-22 - Created script

.EXAMPLE
    PS C:\> Get-PKPSVersions -Verbose 
        VERBOSE: PSBoundParameters: 

        Key              Value
        ---              -----
        Verbose          True
        UseSearchPath    False
        PSCorePaths      {C:\Program Files\PowerShell, C:\Program Files (x86)\PowerShell}
        ParameterSetNAme Custom
        ScriptName       Get-PKPSVersions
        ScriptVersion    1.0.0

        VERBOSE: Looking for PowerShell versions on LAPTOP145
        VERBOSE: Current script is running in version 7.5.1
        VERBOSE: Looking for Windows PowerShell (powershell.exe)
        VERBOSE: ...searching in C:\Program Files\PowerShell
        VERBOSE: ...searching in C:\Program Files (x86)\PowerShell
        WARNING: Invalid path 'C:\Program Files (x86)\PowerShell'

        ComputerName Version   Path
        ------------ -------   ----
        LAPTOP145    5.1.22621 C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe
        LAPTOP145    7.5.1     C:\Program Files\PowerShell\7\pwsh.exe

.EXAMPLE
    PS C:\> Get-PKPSVersions -Verbose -UseSearchPath

        VERBOSE: PSBoundParameters: 

        Key              Value
        ---              -----
        Verbose          True
        UseSearchPath    True
        PSCorePaths      {C:\Program Files\PowerShell, C:\Program Files (x86)\PowerShell}
        ParameterSetNAme Path
        ScriptName       Get-PKPSVersions
        ScriptVersion    1.0.0

        VERBOSE: Looking for PowerShell versions on LAPTOP19
        VERBOSE: Current script is running in version 7.5.1
        VERBOSE: Looking for Windows PowerShell (powershell.exe)
        VERBOSE: Looking for PowerShell Core (pwsh.exe)...searching in PATH
        VERBOSE: PowerShell Core version 7.5.1 found in C:\Program Files\PowerShell\7\pwsh.exe

        ComputerName Version   Path
        ------------ -------   ----
        LAPTOP19     5.1.22621 C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe
        LAPTOP19     7.5.1     C:\Program Files\PowerShell\7\pwsh.exe
#>

    [CmdletBinding(DefaultParameterSetName = "Custom")]
    param(
        [Parameter(
            ParameterSetName = "Path",
            Position = 0,
            HelpMessage = "Use paths to search for PowerShell Core installations (pwsh.exe)"
        )]
        [switch]$UseSearchPath,

        [Parameter(
            ParameterSetName = "Custom",
            HelpMessage = "Paths to search for PowerShell Core (pwsh.exe) installations"
        )]
        [ValidateNotNullOrEmpty()]
        [string[]]$PSCorePaths =  @("${env:ProgramFiles}\PowerShell", "${env:ProgramFiles(x86)}\PowerShell")
    )
    Begin {
    
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # How did we get here
        $ScriptName = $MyInvocation.MyCommand.Name
        
        # Show our settings
        $Source = $PSCmdlet.ParameterSetName
        $CurrentParams = $PSBoundParameters
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ParameterSetNAme", $Source)
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        $RegexPattern = '^(\d+\.\d+\.\d+)' # Defines the pattern to capture Major.Minor.Build

        $Activity = "Retrieving PowerShell versions"
        Write-Verbose "[BEGIN: $ScriptName] $Activity"
    }
    Process {

        $Results = @()
        [string]$Version = $Null
        $ComputerName = $Env:ComputerName

        $Msg =  "Looking for PowerShell versions on $ComputerName"
        Write-Verbose $Msg
        $CurrentVersion = $PSVersionTable.PSVersion.ToString()
        $Msg =  "Current script is running in version $CurrentVersion"
        Write-Verbose $Msg
        
        Try {
            $Msg = "Looking for Windows PowerShell (powershell.exe)"
            Write-Verbose $Msg
            $Exe = (Get-Command powershell.exe).Source
            if (Test-Path $Exe) {
                $Version = & $Exe -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
                $Results += [PSCustomObject]@{
                    ComputerName = $ComputerName
                    Version      = If ($Version -match $RegexPattern) { $Matches[0] } else { "Error" }
                    Path         = $Exe
                }
            }
        }
        Catch {
            $Msg = "Error determining Windows Powershell version - $($_.Exception.Message))"
            Write-Error $Msg
        }
        
        $Msg = "Looking for PowerShell Core (pwsh.exe)"
        Switch ($Source) {
            Path {
                $Msg += "...searching in PATH"
                Write-Verbose $Msg
                Try {
                    Get-Command pwsh.exe -ErrorAction Stop | ForEach-Object {
                        $Exe = $_.Source
                        If ($Version = & $Exe -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()` # | Select-Object @{N="V";E={"$($_.Major).$($_.Minor).$($_.Patch)"}} | Select-Object -ExpandProperty V') {
                            $Msg = "PowerShell Core version $Version found in $Exe"
                            Write-Verbose $Msg
                                $Results += [PSCustomObject]@{
                                ComputerName = $ComputerName
                                Version      = If ($Version -match $RegexPattern) { $Matches[0] } else { "Error" }
                                Path         = $Exe
                            }
                        }
                    }
                }
                Catch {
                    $Msg = "PowerShell Core not found in path - $($_.Exception.Message))"
                    Write-Warning $Msg
                }
            }
            Custom {
                Try {
                    Foreach ($PSPath in $PSCorePaths) {
                        $Msg =  "...searching in $PSPath"
                        Write-Verbose $Msg
                        If (Test-Path $PSPath -ErrorAction SilentlyContinue) {
                            Get-ChildItem -Path $PSPath -Filter pwsh.exe -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                                $Exe = $_.FullName
                                $Version = & $Exe -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 
                                $Results += [PSCustomObject]@{
                                    ComputerName = $ComputerName
                                    Version      = If ($Version -match $RegexPattern) { $Matches[0] } else { "Error" }
                                    Path         = $Exe
                                }
                            }
                        }
                        Else {
                            $Msg = "Invalid path '$PSPath'"
                            Write-Warning $Msg
                        }
                    }
                }
                Catch {
                    $Msg = "Error looking for PowerShell Core - $($_.Exception.Message))"
                    Write-Error $Msg
                }
            }
        } # end switch

        Write-Output $Results
    }
    End {
        Write-Verbose "[END: $ScriptName] Operation complete"
    }
} # End Get-PKPSVersions

