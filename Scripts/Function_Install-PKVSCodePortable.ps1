#requires -version 4
Function Install-PKVSCodePortable {
    <#
.SYNOPSIS 
    Downloads and installs or updates VSCode Portable in a specified target directory, since Portable can't update itself! 

.DESCRIPTION
    Downloads and installs or updates VSCode Portable in a specified target directory, since Portable can't update itself
    Skips existing /Data folder and code.lnk shortcut file in the target directory if present
    Won't run inside Visual Studio Code
    First validates target path, then downloads and installs the latest version of VSCode Portable from the URI specified in the script if a newer version is available
    -ForceUpdate parameter will download/overwrite even if current version is the latest; -KillRunningProcess will stop any running VSCode Portable process in the specified path (requires elevation).
    Supports ShouldProcess
    Writes to a log file in the target directory for tracking

.NOTES
    Name    : Function_Install-PKVSCodePortable.ps1
    Created : 2024-04-17
    Author  : Paula Kingsley
    Version : 01.02.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2024-04-17 - Created script
        v01.01.0000 - 2024-07-22 - Updates and fixes
        v01.02.0000 - 2024-10-25 - Fixed error in continue

.LINK
    https://stackoverflow.com/questions/25125818/powershell-invoke-webrequest-how-to-automatically-use-original-file-name

.LINK 
    https://sqladm.in/posts/update-portable-vscode/   

.PARAMETER TargetPath
    Target folder for VSCode Portable (default is $Home\VSCode)

.PARAMETER URI
    URI for portable VSCode zip file download; change at your peril!

.PARAMETER ForceUpdate
    Update current installation even if no newer version was found

.PARAMETER KillRunningProcess
    Stop any running VSCode Portable process in the specified path (requires elevation)

.EXAMPLE 
    PS C:\> Install-PKVSCodePortable -Verbose
    Downloads and installs or updates VSCode Portable in C:\VSCode

.EXAMPLE 
    PS C:\> Install-PKVSCodePortable -TargetPath "C:\MyPath" -URI "https://this.seems-like-a-dodgy-path.right" -Verbose
    Downloads and installs or updates VSCode Portable in a non-default path from a non-default URI

#>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High")]
    Param(
        [Parameter(
            HelpMessage = "Target folder for VSCode Portable (default is `$Home\VSCode)"
        )]
        [ValidateNotNullOrEmpty()]
        [string]$TargetPath = "$Home\VSCode",

        [Parameter(
            HelpMessage = "URI for portable VSCode zip file download; change at your peril!"
        )]
        [ValidateNotNullOrEmpty()]
        [string]$URI = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-archive",

        [Parameter(
            HelpMessage = "Update current installation even if no newer version was found"
        )]
        [switch]$ForceUpdate,

        [Parameter(
            HelpMessage = "Stop any running VSCode Portable process in the specified path (requires elevation)"
        )]
        [switch]$KillRunningProcess
    )
    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "01.02.0000"

        # How did we get here?
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $Source = $PSCmdlet.ParameterSetName

        $CurrentParams = $PSBoundParameters
        $ScriptName = $MyInvocation.MyCommand.Name
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path Variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("ScriptVersion", $Version)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        # We can't commit suicide - although this could be a different executable, I guess. Let's be safe anyway.
        If ($Host.Name -match "Visual Studio") {
            $Msg = "You can't run this from *inside* Visual Studio Code! Please close it and try again from a regular pwsh shell." 
            Write-Warning $Msg
            Break
        }

        #region Inner functions
        
        # check for session elevation
        Function _IsElevated {
            If ((New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                [boolean]$True
            }
            Else {[boolean]$False}
        }

        # find the current filename
        Function _GetRedirectedURL {
            Param ([Parameter(Mandatory = $true)][String]$URL)
            $request = [System.Net.WebRequest]::Create($url)
            $request.AllowAutoRedirect = $false
            $response = $request.GetResponse()
            If ($response.StatusCode -eq "Found") {$response.GetResponseHeader("Location")}
            Else {Write-Warning "Failed to get filename from $URL"; $False}
        }

        # Purge code folder except for date folder/shortcut file, if present
        Function _PurgeFolder {
            [Cmdletbinding(SupportsShouldProcess, ConfirmImpact = "High")]
            Param($Source, $Exclude = @("Data", "code.lnk"))
            If (Get-Item $Source -ErrorAction SilentlyContinue) {
                $Msg = "Purging directory except for $($Exclude -join(', and '))"
                Write-Verbose "[$Source] $Msg"
                Try {
                    If ($PSCmdlet.ShouldProcess($Source, $Msg)) {Get-ChildItem -Path $Source -Exclude $Exclude -ErrorAction Stop | Remove-Item -Recurse -Force -Confirm:$False; $True}
                    Else { $False }
                }
                Catch {Write-Warning $_.Exception.Message; $False}
            }
            Else {$True}
        }

        # Create a new shortcut and pin to the taskbar
        Function _AddShortcut {
            [CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High")]
            param(
                [Parameter(Mandatory)][object]$SourceFile,
                [Parameter()][object]$TargetFolder
            )
            If ($Sourcefile -is [System.IO.FileSystemInfo]) { $FileObj = $SourceFile }
            Elseif ($SourceFile -is [string]) { $FileObj = Get-Item $SourceFile | Where-Object { -not $_.PSIsContainer } }
            If (-not $PSBoundParameters.TargetFolder) { $Folderobj = Get-Item ($FileObj.FullName | Split-Path -Parent) }
            ElseIf ($TargetFolder -is [System.IO.DirectoryInfo]) { $FolderObj = $TargetFolder }
            Elseif ($TargetFolder -is [string]) { $FolderObj = Get-Item $TargetFolder | Where-Object { -not $_.PSIsContainer } }

            If ($FolderObj -and $FileObj) {
                $NewName = ($FileObj.Name) -Replace ("(\.\w+$)", ".lnk")
                $Msg = "Creating shortcut $($FolderObj.FullName)\$NewName from $($FileObj.FullName)"
                Write-Verbose "[$Env:ComputerName] $Msg"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName, $Msg)) {
                    Try {
                        $WshShell = New-Object -comObject WScript.Shell
                        $Shortcut = $WshShell.CreateShortcut("$($FolderObj.FullName)\$NewName")
                        $Shortcut.TargetPath = $FileObj.FullName
                        If ($FileObj.VersionInfo) { $Shortcut.Description = "$($File.Name.SourceFile) ($($FileObj.VersionInfo.ProductVersion))" }
                        $Shortcut.Save()
                        $Output = Get-Item "$($FolderObj.FullName)\$NewName" -ErrorAction Stop
                        $Msg = "Created shortcut $($FolderObj.FullName)\$NewName pointing to $($FileObj.FullName)"
                        Write-Verbose "[$Env:ComputerName] $Msg"
                        Write-Output $Output
                    }
                    Catch { Throw $_.Exception.Message }
                }   
            }
        }

        Function _UpdatePath {
            [CmdletBinding(SupportsShouldProcess,ConfirmImpact = "High")]
            Param($NewPath)
            $CurrentPath = ([environment]::GetEnvironmentVariable("PATH",[EnvironmentVariableTarget]::User))
            $PathArr = $CurrentPath -split(";") | Sort-Object
            If ($PathArr -notContains $NewPath) {
                $Msg = "Current user environment path:`n$($PathArr | Format-List | out-string)"
                Write-Verbose "[$Env:ComputerName] $Msg"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName, "Append $NewPath to current path?")) {
                    Write-Verbose "[$Env:ComputerName] $Msg"
                    [Environment]::SetEnvironmentVariable("PATH", $PathArr -join(";") + ";$NewPath", [EnvironmentVariableTarget]::User)
                    $UpdatedPath = ([environment]::GetEnvironmentVariable("PATH",[EnvironmentVariableTarget]::User)) -split(";") | Sort-Object
                    Write-Verbose "[$Env:ComputerName] Updated path:`n$($UpdatedPath | Format-List | out-string)"
                }   
            }
            Else {
                $Msg = "Path '$NewPath' is already present in user environment path"
                Write-Verbose "[$Env:ComputerName] $Msg"
            }
        }

        #endregion Inner functions

        # Make sure we can do this 
        If ($KillRunningProcess.IsPresent -and -not (_IsElevated)) {
            $Msg = "Session elevation required for -KillRunningProcess! Note that you can manually stop VSCode Portable and re-run without elevation..."
            Write-Warning $Msg 
            Break
        }

        $Activity = "Download and install latest version of VSCode Portable"
        Write-Verbose "[BEGIN: $ScriptName] $Activity"

    } #end begin

    Process {
        
        Try {
            Set-Location $Home
            [switch]$IsInstalled = $False
            [switch]$Continue = $True
            
            $Msg =  "Looking for existing installation" 
            Write-Host "[$Env:ComputerName] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg
            If ($CurrentFile = Get-Item "$TargetPath\code.exe" -ErrorAction SilentlyContinue) {
                [version]$CurrentVersion = $CurrentFile.VersionInfo.FileVersionRaw
                $Msg = "VSCode Portable $CurrentVersion found in $TargetPath"
                Write-Host "[$Env:ComputerName] $Msg"
                $IsInstalled = $True
                Try {
                    $Msg =  "Looking for running VSCode Portable process" 
                    Write-Host "[$Env:ComputerName] $Msg"
                    If ([switch]$IsRunning = (Get-Process -Name code -ErrorAction SilentlyContinue | 
                        Where-Object {$_.Path.ToString() -eq $CurrentFile.FullName} -ErrorAction SilentlyContinue).Count -gt 0) {
                        $Msg = "VSCode Portable is currently running from $TargetPath"
                        If ($KillRunningProcess.IsPresent) {
                            $Msg += "; -KillRunningProcess detected"
                            Write-Host "[$Env:ComputerName] $Msg"
                        }
                        Else {
                            $Msg += "; -KillRunningProcess not detected"
                            Write-Warning "[$Env:ComputerName] $Msg"
                        }
                    }
                }Catch {}
            }
            Else {
                $Msg = "No current version of VSCode detected in $TargetPath"
                Write-Host "[$Env:ComputerName] $Msg"
                $TargetPath = New-Item "$Home\VSCode" -ItemType Directory -Force -ErrorAction Stop -Confirm:$False | Select-Object -ExpandProperty FullName 
            }
            
            If ($Continue.IsPresent) {
                $Continue = $False 

                $Msg =  "Getting latest version available of VSCode Portable" 
                Write-Host "[$Env:ComputerName] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg
                $NewFileName = [System.IO.Path]::GetFileName((_GetRedirectedURL $URI -ErrorAction Stop)) | Where-Object {$_ -match "^VSCode-Win32-x64"}
            
                # Compare versions 
                If ([version]$NewVersion = [regex]::Matches($NewFileName, "(\d+\.)?(\d+\.)?(\*|\d+)").value | Where-Object { $_ -match '(.*?(?=\.\d))\.(.+)' }) {

                    If ($IsInstalled.IsPresent) {
                        If ((-not ($Newversion -gt $CurrentVersion)) -and (-not $ForceUpdate.IsPresent)) {
                            $Msg = "No newer version available; -ForceUpdate not detected"
                            Write-Host "[$Env:ComputerName] $Msg"
                            $Continue = $False
                        }
                        Else {
                            $Msg = "VSCode Portable $NewVersion is available for download"
                            Write-Host "[$Env:ComputerName] $Msg"
                            $Continue = $True
                        }
                    }
                    Else {$Continue = $True}
                }
                Else {
                    $Msg = "No download found! Please verify your URI is correct."
                    Write-Warning "[$Env:ComputerName] $Msg"
                }
            }

            If ($Continue.IsPresent) {
                $Continue = $False

                If ($PSCmdlet.ShouldProcess($TargetPath, "Download and extract VSCode portable $($NewVersion.ToString())"))  {
                    
                    # Create random name
                    [string]$DownloadFile = "$Env:Temp\$(-join (65..90 | ForEach-Object { [char]$_ } | Get-Random -Count 12)).zip"

                    $Msg = "Downloading VSCode Portable $($NewVersion.ToString()) to $DownloadFile"
                    Write-Host "[$Env:ComputerName] $Msg"
                    Write-Progress -Activity $Activity -CurrentOperation $Msg
                    Invoke-WebRequest -Uri $uri -OutFile $DownloadFile -UseBasicParsing -Verbose:$False
                        
                    If ($FileObj = Get-Item $DownloadFile -ErrorAction SilentlyContinue) {

                        $Msg = "Successfully downloaded $DownloadFile"
                        Write-Host "[$Env:ComputerName] $Msg"

                        If ($IsRunning.IsPresent) {
                            Switch ($KillRunningProcess) {
                                $True {
                                    $Msg = "; killing running $TargetPath processes"
                                    Write-Host "[$Env:ComputerName] $Msg" -ForegroundColor Yellow
                                    If ($PSCmdlet.ShouldProcess($Env:ComputerName, "Kill running code.exe processes from $TargetPath?")) {
                                        Try {
                                            $Null = Get-Process code -ErrorAction SilentlyContinue | Where-Object {$_.Path -eq "$TargetPath\code.exe"} -ErrorAction SilentlyContinue | 
                                                Stop-Process -PassThru -Force -Confirm:$False -ErrorAction Stop
                                                $Continue = $True
                                        }
                                        Catch {
                                            $Msg = "[$Env:ComputerName] $($_.Exception.Message)"
                                            Write-Warning "[$Env:ComputerName] $Msg"
                                        }
                                    }
                                    Else {
                                        $Msg = "Operation cancelled by user; script will now exit"
                                        Write-Warning "[$Env:ComputerName] $Msg" 
                                    }
                                }
                                $False {
                                    $Msg = "VSCode Portable is already running from $TargetPath; -KillRunningProcess not specified"
                                    Write-Warning "[$Env:ComputerName] $Msg" 
                                }
                            } # end switch 
                        } # End if running
                        Else {$Continue = $True}

                        If ($Continue.IsPresent) {
            
                            $Null = $FileObj | Unblock-File -Confirm:$False
                            If ($IsInstalled.IsPresent) {    
                                $Msg = "Removing current version (excluding shortcuts and \data subfolder)"
                                Write-Host "[$Env:ComputerName] $Msg" 
                                Get-ChildItem -Path  $TargetPath -exclude "data", "*.lnk" -Force | Remove-Item -force -recurse -Confirm:$False
                            }
                            Try {
                                $Msg = "Expanding $DownloadFile to $TargetPath"
                                Write-Host "[$Env:ComputerName] $Msg" 
                                $Null = Add-Type -AssemblyName System.IO.Compression.FileSystem
                                [System.IO.Compression.ZipFile]::ExtractToDirectory($DownloadFile, $TargetPath)
                                $Msg = "Downloaded and expanded VSCode Portable $($NewVersion.ToString()) in $TargetPath"
                                "[$(Get-Date -f u)] $Msg" | Out-File $TargetPath\InstallLog.txt -Force
                                Write-Host "[$Env:ComputerName] $Msg" -ForegroundColor Green
                            } 
                            Catch {
                                Write-Warning "[$Env:ComputerName] $($_.Exception.Message)"
                            }
                            Finally {
                                $Msg = "Removing temporary file $DownloadFile"
                                Write-Host "[$Env:ComputerName] $Msg" 
                                $Null = Remove-Item $DownloadFile -Force -Confirm:$False -ErrorAction SilentlyContinue
                            }
                        }
                    }
                    Else {
                        $Msg = "Failed to download $URI to $DownloadFile"
                        Write-Warning "[$Env:ComputerName] $Msg"
                    }
                }  # end if shouldprocess
                Else {
                    $Msg = "Operation cancelled by user"
                    Write-Host "[$Env:ComputerName] $Msg"
                }    
            } #end if new continue
            
        }
        Catch {
            Write-Warning "[$Env:ComputerName] $($_.Exception.Message)"
        }
    } #end process

    End {
        Write-Progress -Activity * -Completed 
        Write-Verbose "[END $ScriptName]"
    }
} #end Install-PKVSCodePortable


$Null = New-Alias Update-PKVSCodePortable -Value Install-PKVSCodePortable -Force -ErrorAction SilentlyContinue -Description "Alias for Install-PKVSCodePortable for easier navigation/findability"

