#requires -version 4
Function Install-PKVSCodePortable {
    <#
.SYNOPSIS 
    Downloads and installs or updates VSCode Portable in a specified target directory, since Portable can't update itself! 

.DESCRIPTION
    Downloads and installs or updates VSCode Portable in a specified target directory, since Portable can't update itself! 
    First validates target path, then downloads and installs the latest version of VSCode Portable from the URI specified in the script if a newer version is available.
    -Force parameter will download/overwrite even if current version is the latest
    Supports ShouldProcess
    Writes to a log file

.NOTES
    Name    : Function_Install-PKVSCodePortable.ps1
    Created : 2024-04-17
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2024-04-17 - Created script

.LINK
    https://stackoverflow.com/questions/25125818/powershell-invoke-webrequest-how-to-automatically-use-original-file-name

.LINK 
    https://sqladm.in/posts/update-portable-vscode/   

.PARAMETER TargetPath
    Target folder for VSCode Portable (default is $Home\VSCode)

.PARAMETER URI
    URI for VSCode zip file download; change at your peril!

.PARAMETER Force
    Force download and update even if current version matches new version

.EXAMPLE 
    Install-PKVSCodePortable -Verbose
    Downloads and installs or updates VSCode Portable in C:\VSCode

.EXAMPLE 
    PS C:\> Install-PKVSCodePortable -TargetPath "C:\MyPath" -URI "https://this.seemslike.adodgypath.what" -Verbose
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
            HelpMessage = "URI for VSCode zip file download; change at your peril!"
        )]
        [ValidateNotNullOrEmpty()]
        [string]$URI = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-archive",

        [Parameter(
            HelpMessage = "Force download and update even if current version matches new version"
        )]
        [switch]$Force

    )
    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

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
        If ($Host.Name -match "Visual Studio") {$Msg = "Dude, don't run this from *inside* Visual Studio Code! Please close it and try again from a regular pwsh shell."; Throw $Msg}

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

        $Activity = "Download and install VSCode Portable"
        Write-Verbose "[BEGIN: $ScriptName] $Activity"

    } #end begin

    Process {
        # Set the flag (because too many nested try/catch statements are annoying)
        [switch]$Continue = $True

        # Create log file
        $Logfile = New-Item -ItemType File -Path "$TargetPath\$(get-date -f yyyy-MM-dd)_vscode_portable.log" -Force -ErrorAction Stop
        
        #region Quickly check newest version available
        If ($Continue.IsPresent) {
            $Continue = $False
            $Msg = "Getting latest VSCode zip file name"
            Write-Verbose "[$Env:ComputerName] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg
            Try {
                $FileName = [System.IO.Path]::GetFileName((_GetRedirectedURL $URI -ErrorAction Stop))
                If ($Filename -match "^VSCode-Win32-x64") {
                    $Version = [regex]::Matches($FileName, "(\d+\.)?(\d+\.)?(\*|\d+)").value | 
                        Where-Object { $_ -match '(.*?(?=\.\d))\.(.+)' }
                    If ($CurrentExecutable = Get-command "$TargetPath\code.exe" -ErrorAction SilentlyContinue) {
                        $CurrentVersion = $CurrentExecutable.Version.ToString()
                        If ([version]$Version -gt [version]$CurrentVersion) {
                            $Msg = "Update availalable! Latest version is $Version; you are currently running $CurrentVersion"                            
                            Write-Verbose "[$Env:ComputerName] $Msg"
                            "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                            $Continue = $True
                        }
                        Else {
                            $Msg = "No newer version available; latest available is $Version and you are currently running $CurrentVersion"
                            If ($Force.IsPresent) {
                                $Msg += "; -Force detected" 
                                $Continue = $True
                            }
                            "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                            Write-Warning "[$Env:ComputerName] $Msg"
                        }
                    }
                    Else {
                        $Msg = "Found VSCode $Version available for download from $URI"
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                        Write-Verbose "[$Env:ComputerName] $Msg"
                        $Continue = $True
                    }
                }
            }
            Catch {
                $Msg = "[$Env:ComputerName] Failed to get latest VSCode file name $($_.Exception.Message)"
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                Write-Warning "[$Env:ComputerName] $Msg"
            }    
        }
        #endregion Quickly check newest version available

        #region Download file
        If ($Continue.IsPresent) {
            $Continue = $False

            Try {
                $OutFile = "$([System.IO.Path]::GetTempFileName().Replace('.tmp','.zip'))"
                $Null = Get-Item $Outfile -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction Stop

                $Msg = "Downloading zip file to temporary file using Invoke-WebRequest"
                Write-Verbose "[$Env:ComputerName] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg
                If ($PSCmdlet.ShouldProcess($Env:COMPUTERNAME, $Msg)) {
                    Try {
                        $Null = Invoke-WebRequest -Uri $Uri -UseBasicParsing -Method Get -OutFile $Outfile -ErrorAction Stop 
                        If ($Null = Get-Item $Outfile -ErrorAction SilentlyContinue) {
                            $Msg = "Successfully downloaded $Filename to $Outfile ($([math]::round((get-item $Outfile).length/1mb)) MB)"
                            Write-Verbose "[$Env:ComputerName] $Msg"
                            "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                            $Continue = $True
                        }
                        Else {
                            $Msg = "[$Env:ComputerName] Failed to download file! $($_.Exception.Message)"
                            "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                            Write-Warning "[$Env:ComputerName] $Msg"
                        }
                    }
                    Catch {
                        $Msg = "[$Env:ComputerName] Failed to download file! $($_.Exception.Message)"
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                        Write-Warning "[$Env:ComputerName] $Msg"
                    }
                }        
            }
            Catch {
                $Msg = "[$Env:ComputerName] Error! $($_.Exception.Message)"
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                Write-Warning "[$Env:ComputerName] $Msg"
            }
        }
        #endregion Download file 

        #region Validate or create target path
        If ($Continue.IsPresent) {
            $Continue = $False

            $Msg = "Validating target path and data subfolder"
            Write-Verbose "[$Env:ComputerName] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg
            If (Get-Item -Path $TargetPath -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }) {
                If ($DataFolder = Get-Item -Path "$TargetPath\Data" -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }) {
                    $Msg = "Found $($Datafolder.FullName), lastwritetime $(Get-Date $Datafolder.LastWriteTime -f u)"
                    Write-Verbose "[$Env:ComputerName] $Msg"
                    "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                    $Continue = $True
                }
                Else {
                    If ($PSCmdlet.ShouldProcess($TargetPath, "Create new data subfolder?")) {
                        Try {
                            $NewFolder = New-Item -Path "$TargetPath\Data" -ItemType Directory -Confirm:$False -ErrorAction Stop
                            $Msg = "Created $($NewFolder.FullName)"
                            Write-Verbose "[$Env:ComputerName] $Msg"
                            "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                            $Continue = $True
                        }
                        Catch {
                            $Msg = "Error! failed to create '$TargetPath\Data' $($_.Exception.Message)"; 
                            "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                            Write-Warning "[$Env:ComputerName] $Msg"
                        }
                    }
                }
            } #end if path exists
            Else {
                If ($PSCmdlet.ShouldProcess($Env:COMPUTERNAME, "Create target path and data subfolder?")) {
                    Try {
                        $NewFolder = New-Item -Path "$TargetPath\Data" -ItemType Directory -Confirm:$False -ErrorAction Stop
                        $Msg = "Created $($NewFolder.FullName)"
                        Write-Verbose "[$Env:ComputerName] $Msg"
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                        $Continue = $True
                    }
                    Catch {
                        $Msg = "Error! failed to create target path '$TargetPath\Data' $($_.Exception.Message)"
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                        Write-Warning "[$Env:ComputerName] $Msg"
                    }
                }
            }
        }

        #region Stop any running processes
        If ($Continue.IsPresent) {
            $Continue = $False
            Try {
                $Process = Get-Process code -ErrorAction SilentlyContinue | Where-Object {$_.Path -eq "$TargetPath\code.exe"} -ErrorAction SilentlyContinue
                $Msg = "Detected running VSCode processes from target folder $TargetPath"
                Write-Warning "[$Env:ComputerName] $Msg"
                $ConfirmMsg = "Stop all current code.exe processes from $TargetPath?"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
                    Try {
                        $Process | Stop-Process -PassThru -Force -Confirm:$False
                        $Continue = $True
                        $Msg = "Stopped running code.exe processes in $TargetPath"
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                        Write-Warning "[$Env:ComputerName] $Msg"
                    }
                    Catch {
                        $Msg = "Error! Failed to stop current process(es) $($_.Exception.Message)"
                        If ($Force.IsPresent) {
                            $Msg += "; -Force detected" 
                            $Continue = $True
                        }
                        "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                        Write-Warning "[$Env:ComputerName] $Msg"
                    }
                }
            }
            Catch {
                $Msg = "Failed to get current process(es) $($_.Exception.Message)"
                Write-Warning "[$Env:ComputerName] $($_.Exception.Message)"
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
            }
        }
        #endregion Stop any running processes

        #region Purge current folder except data folder & shortcut
        If ($Continue.IsPresent) {
            $Continue = $False
            $Msg = "Removing prior version in $Targetpath, if present"
            If ($Null = Get-Childitem "$TargetPath\Data" -Recurse) { $Msg += " (excluding shortcut file & \data subfolder contents, if applicable)" }
            Write-Verbose "[$Env:ComputerName] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg
            Try {
                $Null = _PurgeFolder -Source $TargetPath -Verbose -ErrorAction Stop
                $Msg = "Successfully removed prior version content, if any (excluding data subfolder and shortcut file)"
                Write-Verbose "[$Env:ComputerName] $Msg"
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                $Continue = $True
            }
            Catch {
                $Msg = "Failed to purge existing folder contents! $($_.Exception.Message)"
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                Write-Warning "[$Env:ComputerName] $Msg"
            }
        }        
        #endregion Purge current folder except data folder & shortcut

        #region Expand download file to target path
        If ($Continue.IsPresent) {
            $Continue = $False
            $Msg = "Extracting $Outfile contents to $Targetpath"
            Write-Verbose "[$Env:ComputerName] $Msg"
            Write-Progress -Activity $Activity -CurrentOperation $Msg
            If ($PSCmdlet.ShouldProcess($Env:COMPUTERNAME, $Msg)) {
                Try {
                    $Null = Expand-Archive -Path $Outfile -DestinationPath $Targetpath -force -ErrorAction Stop
                    $Executable = Get-Childitem $Targetpath -Recurse -filter "code.exe" -ErrorAction Stop
                    $Version = ($Executable.VersionInfo).ProductVersion
                    $Logfile = New-Item -ItemType File -Path "$TargetPath\$(get-date -f yyyy-MM-dd)_vscode-$Version`_install.log" -Force -ErrorAction Stop
                    "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] Successfully updated VSCode Portable to $Version in $TargetPath on $Env:ComputerName" | Out-File $Logfile -Append
                    $Msg = "Successfully expanded VSCode Portable to $Targetpath"
                    Write-Verbose "[$Env:ComputerName] $Msg"
                    "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                    $Continue = $True
                }
                Catch {
                    $Msg = "Error! Failed to expand archive $($_.Exception.Message)"
                    "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
                    Write-Warning "[$Env:ComputerName] $Msg"
                }
            }
        }
        #endregion Expand download file to target path

        #region Update path and create shortcut 
        
        If ($Continue.IsPresent) {
            $Continue = $False

            _UpdatePath -NewPath $TargetPath

            If (-not ($Null = Get-Item "$Targetpath\code.lnk" -ErrorAction SilentlyContinue)) {
                $Msg = "Creating shortcut to code.exe"
                Write-Verbose "[$Env:ComputerName] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg
                _AddShortcut -SourceFile $Executable.FullName -Verbose 
                $Msg = "Added shortcut to $Targetpath\code.exe"
                Write-Verbose "[$Env:ComputerName] $Msg"
                "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $Msg" | Out-File $Logfile -Append
            }
        }
        #endregion Create shortcut 

    }
    End {

        Try {
            If (Get-Item $LogFile -ErrorAction SilentlyContinue) {
                $Msg = "Script is complete; log file can be found at $LogFile (contents below)"
                Write-Verbose "[$Env:ComputerName] $Msg"
                Get-Content $LogFile
            }
        }
        Catch {}

        $Null = Get-Item $Outfile -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Write-Verbose "[END $ScriptName]"
    }
} #end function

