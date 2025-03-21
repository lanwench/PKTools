#requires -version 4
Function Backup-PKVSCodeData {
    <#
.SYNOPSIS 
    Backs up the user \Data folder for VSCode to a date-named compressed file in the target path of your choice

.DESCRIPTION
    Backs up the user \Data folder for VSCode to a date-named compressed file in the target path of your choice

.NOTES
    Name    : Function_Backup-PKVSCodeData.ps1
    Created : 2025-01-31
    Author  : Paula Kingsley
    Version : 01.0.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2025-01-31 - Created script
        
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
            HelpMessage = "Source folder for VSCode (if not specified, looks for parent folder of code.exe)"
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path $_ -PathType Container})]
        [string]$VSCodePath, # = "$Home\VSCode", # (Get-Item (Get-Command code -ErrorAction Stop).Source -ErrorAction Stop).Directory, # "$Home\VSCode",

        [Parameter(
            HelpMessage = "Target folder for backups (default us user's temp directory)"
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path $_ -PathType Container})]
        [string]$BackupPath = $Env:TEMP

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
        $CurrentParams.Add("ParameterSetName", $Source)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        $Activity = "Back up VSCode \Data folder to $BackupPath"
        Write-Verbose "[BEGIN: $ScriptName] $Activity"

        # Make sure files aren't in use!
        If ($Host.Name -match "Visual Studio") {
            $Msg = "Sorry, you can't run this from *inside* Visual Studio Code! Please close it and try again from a regular pwsh shell." 
            Write-Warning $Msg
            Break
        }

        # Get the path if it wasn't specified
        If (-not $VSCodePath) {
            Try {
                $VSCodePath = (Get-Item (Get-Command code -ErrorAction Stop).Source -ErrorAction Stop).Directory
                $Msg = "Setting VSCodePath based on code.exe location"
                Write-Verbose "[$VSCodePath] $Msg"
            }
            Catch {
                $Msg = "VSCodePath not specified; code.exe not found in path"
                Write-Warning $Msg
                Break
            }
        }

        # Make sure there's a subfolder for data
        If (-not (Get-Item "$VSCodePath\Data")) {
            $Msg = "No \Data subfolder found"
            Write-Warning "[$VSCodePath] $Msg"
            Break
        }

        # Create the filename
        $BackupFile = "$BackupPath\VSCodeBackup-$((Get-Date).ToString('yyyyMMdd-HHmmss')).zip"       

    } #end begin

    Process {

        
        Try {
            $Msg = "Get data subfolder"
            Write-Verbose "[$VSCodePath] $Msg"
            $DataFolder = Get-Item "$VSCodePath\Data" -ErrorAction Stop
            $Msg = "Compress VSCode data folder to $BackupFile"
            Write-Verbose "[$VSCodePath] $Msg"
            

            <#
            If ($PSCmdlet.ShouldProcess($Env:ComputerName, $Msg)) {

                $Msg = "Compressing $VSCodePath\Data to $BackupFile"
                Write-Verbose "[$VSCodePath] $Msg"
            
                # Start the timer
                $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            
                # Create a new zip archive object
                $zip = [System.IO.Compression.ZipFile]::Open($BackupFile, "Create")
            
                # Get all files in the source folder (including subfolders if needed)
                $files = Get-ChildItem -Path $DataFolder -File -Recurse
            
                # Loop through each file, create the entry name, and add it to the zip archive with its relative path, then close zip/finalize compression
                foreach ($file in $files) {
                    $FileStr = $File.FullName -replace [regex]::Escape($DataFolder.FullName), ""
                    Write-Progress -Activity "Compressing VSCode data folder to $BackupFile" -Status "Adding $FileStr" -PercentComplete ($files.IndexOf($file) / $files.Count * 100)
            
                    $entryName = $file.FullName.Substring($DataFolder.Length).TrimStart("\") # Use $DataFolder here
            
                    $entry = $zip.CreateEntry($entryName)
                    $fileStream = [System.IO.File]::OpenRead($file.FullName)
                    $entryStream = $entry.Open()
                    $fileStream.CopyTo($entryStream)
                    $fileStream.Close()
                    $entryStream.Close()
                }
                $zip.Dispose()
                $Msg = "Activity completed in $($Stopwatch.Elapsed)"
                Write-Verbose "[$VSCodePath] $Msg"
                Write-Progress -Activity * -Completed

            } 
            Else {
                $Msg = "Operation cancelled by user"
                Write-Verbose "[$VSCodePath] $Msg"
            }
            #>

            <#
                Backip before changing to show job status
                If ($PSCmdlet.ShouldProcess($Env:ComputerName, $Msg)) {

                # Start the timer
                $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            
                $Msg = "Creating $BackupFile object"
                Write-Verbose "[$VSCodePath] $Msg"
                $zip = [System.IO.Compression.ZipFile]::Open($BackupFile, "Create")
            
                $Msg = "Getting child items (this can take a while!)"
                Write-Verbose "[$VSCodePath] $Msg"
                $files = Get-ChildItem -Path $DataFolder -File -Recurse

                $Msg = "Creating a queue to hold the file information"
                Write-Verbose "[$VSCodePath] $Msg"
                $fileQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
                foreach ($file in $files) { $fileQueue.Enqueue($file) }

                $Msg = "Setting maximum concurrent threads"
                Write-Verbose "[$VSCodePath] $Msg"
                $maxThreads = [Environment]::ProcessorCount  # Or a smaller number if needed

                # Create a thread job for each file
                $Msg = "Creating a thread job for each file"
                Write-Verbose "[$VSCodePath] $Msg"
                $jobs = 1..$maxThreads | ForEach-Object {
                    Start-Job -ScriptBlock {
                        param($fileQueue, $zip, $DataFolder) # Pass parameters to the job

                        while ($fileQueue.Count > 0) {
                            if ($fileQueue.TryDequeue([ref]$file)) { # Try to get a file from the queue
                                $entryName = $file.FullName.Substring($DataFolder.Length).TrimStart("\")
                                try {
                                    $entry = $zip.CreateEntry($entryName)
                                    $fileStream = [System.IO.File]::OpenRead($file.FullName)
                                    $entryStream = $entry.Open()
                                    $fileStream.CopyTo($entryStream)
                                } finally { # Important: Ensure streams are closed even on error
                                    if ($fileStream) { $fileStream.Close() }
                                    if ($entryStream) { $entryStream.Close() }
                                }
                            } else {
                                Start-Sleep -Milliseconds 10 # Avoid busy-waiting
                            }
                        }
                    } -ArgumentList $fileQueue, $zip, $DataFolder
                }

                # Wait for all jobs to complete
                $Msg = "Waiting for all jobs to complete"
                Write-Verbose "[$VSCodePath] $Msg"
                Wait-Job $jobs | Receive-Job
                
                $zip.Dispose()
                $Msg = "Activity completed in $($Stopwatch.Elapsed)"
                Write-Verbose "[$VSCodePath] $Msg"
                Write-Progress -Activity * -Completed

                $Msg = "Successfully created output file"
                Write-Verbose "[$VSCodePath] $Msg"
                Get-Item $BackupFile

            } 
            Else {
                $Msg = "Operation cancelled by user"
                Write-Verbose "[$VSCodePath] $Msg"
            }
            #>
            If ($PSCmdlet.ShouldProcess($Env:ComputerName, $Msg)) {

                # Start the timer
                $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            
                $Msg = "Creating $BackupFile object"
                Write-Verbose "[$VSCodePath] $Msg"
                $zip = [System.IO.Compression.ZipFile]::Open($BackupFile, "Create")
            
                $Msg = "Getting child items (this can take a while!)"
                Write-Verbose "[$VSCodePath] $Msg"
                $files = Get-ChildItem -Path $DataFolder -File -Recurse
                $totalFiles = $files.Count # Store total file count
            
                $Msg = "Creating a queue to hold the file information"
                Write-Verbose "[$VSCodePath] $Msg"
                $fileQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
                foreach ($file in $files) { $fileQueue.Enqueue($file) }
            
                $Msg = "Setting maximum concurrent threads"
                Write-Verbose "[$VSCodePath] $Msg"
                $maxThreads = [Environment]::ProcessorCount  # Or a smaller number if needed
            
                # Create a thread job for each file
                $Msg = "Creating a thread job for each file"
                Write-Verbose "[$VSCodePath] $Msg"
                $jobs = 1..$maxThreads | ForEach-Object {
                    Start-Job -ScriptBlock {
                        param($fileQueue, $zip, $DataFolder, $totalFiles) # Add $totalFiles parameter
            
                        # Hashtable to track progress for this job
                        $jobProgress = @{ filesProcessed = 0 }
            
                        while ($fileQueue.Count > 0) {
                            if ($fileQueue.TryDequeue([ref]$file)) {
                                $entryName = $file.FullName.Substring($DataFolder.Length).TrimStart("\")
                                try {
                                    $entry = $zip.CreateEntry($entryName)
                                    $fileStream = [System.IO.File]::OpenRead($file.FullName)
                                    $entryStream = $entry.Open()
                                    $fileStream.CopyTo($entryStream)
                                } finally {
                                    if ($fileStream) { $fileStream.Close() }
                                    if ($entryStream) { $entryStream.Close() }
                                }
            
                                # Increment and report progress
                                $jobProgress.filesProcessed++
                                $percentComplete = [Math]::Round(($jobProgress.filesProcessed / $totalFiles) * 100)
                                Write-Progress -Activity "Compressing $DataFolder" -Status "Thread $($_.Id): $($jobProgress.filesProcessed) of $totalFiles files" -PercentComplete $percentComplete -Id $_.Id
                            } else {
                                Start-Sleep -Milliseconds 10
                            }
                        }
                    } -ArgumentList $fileQueue, $zip, $DataFolder, $totalFiles # Pass $totalFiles
                }
            
                # Wait for all jobs to complete (no Receive-Job needed)
                $AllJobs = (Get-Job).Count
                $Msg = "Waiting for $AllJobs jobs to complete"
                Write-Verbose "[$VSCodePath] $Msg"
                Wait-Job $jobs
            
                $zip.Dispose()
                $Msg = "Activity completed in $($Stopwatch.Elapsed)"
                Write-Verbose "[$VSCodePath] $Msg"
                Write-Progress -Activity * -Completed
            
                $Msg = "Successfully created output file"
                Write-Verbose "[$VSCodePath] $Msg"
                Get-Item $BackupFile
            
            } 
            Else {
                $Msg = "Operation cancelled by user"
                Write-Verbose "[$VSCodePath] $Msg"
            }


            <#
            
            If ($PSCmdlet.ShouldProcess($DataFolder, $Msg)) {
                #Write-Progress -Activity $Activity -Status "Compressing $DataFolder to $BackupFile" 

                $Msg = "Compressing data folder to $BackupFile"
                Write-Verbose "[$VSCodePath] $Msg"

#                [IO.Compression.ZipFile]::CreateFromDirectory( $Datafolder, $Backupfile, 'Fastest', $false )

                # Start the timer
                $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()                
                
                # Create a new zip archive object
                $zip = [System.IO.Compression.ZipFile]::Open($BackupFile, "Create")

                # Get all files in the source folder (including subfolders if needed)
                $files = Get-ChildItem -Path $DataFolder -File -Recurse

                # Loop through each file , create the entry name, and add it to the zip archive with its relative path, then close zip/finalize compression
                foreach ($file in $files) {
                    Write-Progress -Activity "Compressing $DataFolder to $BackupFile" -Status "Adding $file" -PercentComplete ($files.IndexOf($file) / $files.Count * 100)
                    $entryName = $file.FullName.Substring($sourceFolder.Length).TrimStart("\")
                    $zip.CreateEntryFromFile($file.FullName, $entryName)
                }
                $zip.Dispose()
                $Msg = "Activity completed in $($Stopwatch.Elapsed)"
                Write-Verbose "[$VSCodePath] $Msg"
                Write-Progress -Activity * -Completed
            }   
            Else {
                $Msg = "Operation cancelled by user"    
                Write-Verbose "[$VSCodePath] $Msg"
            }
                #>
        }
        Catch {
            Write-Progress -Activity * -Completed
            $Msg = "Operation failed! $($_.Exception.Message)"
            Write-Warning "[$VSCodePath] $Msg"
        }
    } #end process
    End {
        Write-Progress -Activity * -Completed 
        $Stopwatch.Stop()
        Write-Verbose "[END $ScriptName]"
    }
} #end Backup-VSCodeData

