#requires -Version 3
Function Backup-PKChrome {
[CmdletBinding(SupportsShouldProcess,ConfirmImpact="High")]
Param(
    [Parameter(
        Mandatory = $False,
        HelpMessage = "Absolute path for backup files"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({Test-Path $_})]
    [string]$TargetPath = $env:USERPROFILE,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Copy all folder contents, or copy only folders matching 'Default' and 'Profile'"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("All","ProfilesOnly")]
    [string]$SourceType = "All",

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Exclude cache files"
    )]
    [switch]$ExcludeCache, 

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Suppress all non-verbose/non-error console output"
    )]
    [switch]$SuppressConsoleOutput 

)
Begin{
    
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }

    $FileDate = Get-Date -f yyyy-MM-dd_hh-mm-ss @StdParams
    $Target = "$TargetPath\$FileDate`_Chrome_$Env:ComputerName"

    Switch ($SourceType) {
        All {$Msg = "Copy all Chrome user data to $Target"}
        ProfilesOnly {$Msg = "Copy all Chrome user profile data folders to $Target"}
    }
    If ($ExcludeCache.IsPresent) {$Msg += " (excluding cache files)"}

    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Msg"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}
    $Host.UI.WriteLine()

}
Process {

    Try {

        [switch]$Continue = $False
        
        $Source = ("$env:LOCALAPPDATA\Google\Chrome\User Data").ToLower()

        If ($Null = Test-Path $Source -ErrorAction SilentlyContinue -Verbose:$False) {
            
            $FileList = @()

            Switch ($SourceType) {
                All {
                    $Filelist = Get-Childitem $Source –Recurse @StdParams
                }
                ProfilesOnly {
                    $FileList += (Get-ChildItem -Path $Source -Filter "default" -Recurse -Force @StdParams)
                    $FileList += (Get-ChildItem -Path $Source -Filter "profile*" -Recurse -Force @StdParams)                        
                }
            }
            If ($ExcludeCache.IsPresent) {
                $FileList = $FileList | Where-Object {$_.fullname -notmatch "cache"}
            }

            <#

            Switch ($SourceType) {
                All {
                    If ($ExcludeCache.IsPresent) {
                        $Filelist = Get-Childitem $Source –Recurse @StdParams | Where-Object {$_.fullname -notmatch "cache"}
                    }
                    Else {
                        $Filelist = Get-Childitem $Source –Recurse @StdParams
                    }
                
                }
                Profiles{
                    $Directories = "user data\default","user data\profile"
                    $Include = [regex]::escape($Directories)

                    #$FilterSource = (Get-ChildItem -Path $Source -Filter "default" -Recurse @StdParams),(Get-ChildItem -Path $Source -Filter "profile*" -Recurse @StdParams)
                    
                    If ($ExcludeCache.IsPresent) {
                        $Exclude = "cache"
                        $FileList = (Get-ChildItem -Path $Source -Filter "default" -Recurse @StdParams),(Get-ChildItem -Path $Source -Filter "profile*" -Recurse @StdParams) | Where-object {$_.fullname -notmatch $Exclude
                        #$Filelist = Get-Childitem $Source –Recurse @StdParams | 
                        #    Where-object {$_.fullname -notmatch $Exclude -and $_.FullName -match $Include}
                    }
                    Else {
                        $FileList = (Get-ChildItem -Path $Source -Filter "default" -Recurse @StdParams),(Get-ChildItem -Path $Source -Filter "profile*" -Recurse @StdParams) | Where-object {$_.fullname -notmatch $Exclude
                        #$Filelist = Get-Childitem $Source –Recurse @StdParams | 
                        #    Where-object {$_.FullName -match $Include}
                    }
                }
            }

            #>

            $Total = $FileList.Count

            $Msg = "Found Chrome profile path '$Source'"
            Write-Verbose $Msg
            $Continue = $True
        }
        Else {
            $Msg = "Invalid path '$("$env:LOCALAPPDATA\Google\Chrome\User Data")'"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
        }
        

        If ($Continue.IsPresent) {

            $Continue = $False
        
            If ($ChromeProc = Get-Process chrome -ErrorAction SilentlyContinue) {
            
                $Msg = "Stop running Chrome process(es)"
                Write-Verbose $Msg

                $ConfirmMsg = "`n$Msg`n`n"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$Msg)) {
                    $StopProc = $ChromeProc | Stop-Process -Force -PassThru @StdParams
                    $Msg = "Stopped Chrome"
                    Write-Verbose $Msg
                    $Continue = $True
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    Write-Verbose $Msg
                }
            }
            Else {
                $Continue = $True
            }
        }


        If ($Continue.IsPresent) {

            Try {

                $Msg = "Back up '$Source' files to '$Target'"
                Write-Verbose $Msg
                Write-Progress -Activity $Msg -Status $Env:ComputerName
                
                $ConfirmMsg = "`n$Msg`n`n"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$Msg)) {    

                    $Current = 0
                    $CopyFiles = Foreach ($File in $Filelist) {
                        
                        $Current ++
                        $TargetFile = $File.Fullname.tolower().replace($Source,'')

                        $DestinationFile = ($Target+$TargetFile)
                        
                        Write-Progress -Activity "Backup Chrome data files on $Env:ComputerName to '$Target'" -Status "Copying $File" -PercentComplete ($Current/$Total*100)
                        
                        Write-Verbose $File.FullName

                        $Null = Copy-Item $File.FullName -Destination $DestinationFile -Force -PassThru -Confirm:$False @StdParams
                    
                    }

                    #


                    #$CopyFiles = Copy-Item -Path $Source -Recurse -Destination $Target -Force -PassThru @StdParams

                    #

                    $GetFiles = Get-ChildItem $Target -Recurse -File @StdParams
                    $Msg = "$('{0:N0}' -f $($GetFiles.Count)) file(s) copied to '$Target'"
                    Write-Verbose $Msg
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    Write-Verbose $Msg
                }
            }
            Catch {
                $Msg = "File backup failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
            }
        }
    }
    Catch {
        $Msg = "Operation failed"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
    }

}

}
