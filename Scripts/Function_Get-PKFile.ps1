#requires -Version 3
Function Get-PKFile {
<#
.SYNOPSIS
    Uses a queue object and .NET [IO.Directory] to search a path for files matching a string

.DESCRIPTION
    Uses a queue object and .NET [IO.Directory] to search a path for files matching a string
    Significantly faster than Get-ChildItem -Recurse
    Accepts pipeline input
    Returns an array of strings

.NOTES
    Name    : Function_Get-PKFile.ps1
    Author  : Paula Kingsley
    Created : 2019-04-25
    Version : 01.00.0000    
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2019-04-25 - Created script based on Lee Holmes' original (see link)

.PARAMETER Directory
    Directory to scan (default is current directory)

.PARAMETER SearchString
    String to search for; unless -ExactMatchOnly is specified, will treat as wildcard (default is '*')

.PARAMETER ExactMatchOnly
    No wildcard matches

.PARAMETER Quiet
    Suppress all non-verbose console output

.LINK
    https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/using-a-queue-instead-of-a-recursion

.EXAMPLE
    PS C:\Users\JBloggs\Repos\Misc> Get-PKFile -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                  
        ---            -----                                  
        Verbose        True                                   
        Directory      C:\Users\JBloggs\Repos\Misc
        SearchString   *                                      
        ExactMatchOnly False                                  
        Quiet          False                                  
        PipelineInput  False                                  
        ScriptName     Get-PKFile                             
        ScriptVersion  1.0.0                                  

        BEGIN  : Get files matching search pattern '*'

        C:\Users\JBloggs\Repos\MiscTools\MiscTools.psd1
        C:\Users\JBloggs\Repos\MiscTools\MiscTools.psm1
        C:\Users\JBloggs\Repos\MiscTools\README.md
        C:\Users\JBloggs\Repos\MiscTools\.git\COMMIT_EDITMSG
        C:\Users\JBloggs\Repos\MiscTools\.git\config
        C:\Users\JBloggs\Repos\MiscTools\.git\description
        C:\Users\JBloggs\Repos\MiscTools\.git\FETCH_HEAD
        C:\Users\JBloggs\Repos\MiscTools\.git\HEAD
        C:\Users\JBloggs\Repos\MiscTools\.git\index
        C:\Users\JBloggs\Repos\MiscTools\.git\ORIG_HEAD
        C:\Users\JBloggs\Repos\MiscTools\.git\packed-refs
        C:\Users\JBloggs\Repos\MiscTools\.vscode\settings.json
        C:\Users\JBloggs\Repos\MiscTools\Scripts\Function_Get-WindowsLogonInfo.ps1
        C:\Users\JBloggs\Repos\MiscTools\Scripts\Function_Get-NestedGroupMember.ps1
        C:\Users\JBloggs\Repos\MiscTools\Scripts\Function_Repair-DISM.ps1
        C:\Users\JBloggs\Repos\MiscTools\Scripts\Scratchpad.txt
        C:\Users\JBloggs\Repos\MiscTools\Scripts\Test_Function_DoThings.ps1

        END    : Get files matching search pattern '*'

.EXAMPLE
    PS C:\> Get-PKFile -Directory 'G:\My Drive\' -SearchString "new server request form*"
        
        BEGIN  : Get files matching search pattern 'new server request form*'
        
        G:\My Drive\Misc\New server request form 4.2.docm
        G:\My Drive\Misc\New server request form 4.1.docm
        G:\My Drive\Misc\New server request form 1.0.pdf
        G:\My Drive\Misc\New server request form 4.0.pdf
        G:\My Drive\Misc\New server request form 4.0.docm
        G:\My Drive\Misc\New server request form 2.0.docm
        G:\My Drive\Misc\New server request form 3.0.pdf
        G:\My Drive\Misc\New server request form 3.0.docm
        G:\My Drive\Misc\New server request form 1.0.docm
        G:\My Drive\Misc\New server request form v1.0.pdf
        G:\My Drive\Misc\New server request form.docx
        
        END    : Get files matching search pattern 'new server request form*'


.EXAMPLE
    PS C:\> Get-PKFile -Directory 'G:\My Drive\reports' -SearchString wsus -Quiet 

        G:\My Drive\reports\2016-06-13_OpsWindowsWSUSGroup.csv
        G:\My Drive\reports\2016-02-01_AllWSUSGroupMembers.csv
        G:\My Drive\reports\2014-08-27_AllWSUS.csv
        G:\My Drive\reports\WSUS02.csv
        G:\My Drive\reports\WSUS01.csv
        G:\My Drive\reports\GPOReports\2014-07-08 - wsus_qa.html
        G:\My Drive\reports\GPOReports\2014-07-08 - wsus_test.html
        G:\My Drive\reports\GPOReports\2014-07-08 - wsus_gen.html
        G:\My Drive\reports\GPOReports\2014-07-08 - wsus_dev.html
        G:\My Drive\reports\GPOReports\2014-07-08 - wsus_blockautoupdate.html
        G:\My Drive\reports\NewVMs\WSUSSERVER_BuildReport_.CSV
        G:\My Drive\reports\NewVMs\WSUSSERVER_BuildConfig_2016-09-19_06-18.csv


.EXAMPLE
    PS C:\> Get-PKFile -Directory 'G:\My Drive\' -SearchString login.bat -ExactMatchOnly -Quiet

        G:\My Drive\Misc\Login.bat
    
#>
[CmdletBinding()]
Param(
    
    [Parameter(
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True,
        HelpMessage = "Directory to scan (default is current directory)"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Name","FullName")]
    $Directory = $PWD,

    [Parameter(
        HelpMessage = "String to search for; unless -ExactMatchOnly is specified, will treat as wildcard (default is '*')"
    )]
    [ValidateNotNullOrEmpty()]
    $SearchString = '*',
    
    [Parameter(
        HelpMessage = "No wildcard matches"
    )]
    [switch]$ExactMatchOnly,

    [Parameter(
        HelpMessage = "Suppress all non-verbose/non-error console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    If ($PipelineInput) {$CurrentParams.Directory = $Null}
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    If ($PipelineInput.IsPresent) {$CurrentParams.InputObject = $Null}
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "SilentlyContinue"
    $ProgressPreference    = "Continue"

    If ($SearchString -notmatch "\*") {
        If (-not $ExactMatchOnly.IsPresent) {
            $SearchString = "*$SearchString*"
        }
    }

    # Console output
    $Activity = "Get files matching search pattern '$SearchString'"    
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}
}
Process {
    
    # create a new queue
    $dirs = [System.Collections.Queue]::new()

    # add an initial path to the queue; any folder path in the queue will later be processed
    $dirs.Enqueue($Directory)

    # process all elements on the queue until all are taken
    While ($current = $dirs.Dequeue()){
        # find subfolders of current folder, and if present, add them all to the queue
        try {
            foreach ($_ in [IO.Directory]::GetDirectories($current)){
                Write-Progress -Activity $Activity -CurrentOperation $_
                $dirs.Enqueue($_)
            }
        } catch {}

        try{
            # find all files in the folder currently processed
            [IO.Directory]::GetFiles($current, $SearchString) 
        } catch { }
    
    } #end while 

    <#

    # create a new queue
    $DirQueue = [System.Collections.Queue]::new()

    # add an initial path to the queue
    # any folder path in the queue will later be processed
    $DirQueue.Enqueue($Directory)

    
    # process all elements on the queue until all are taken
    While ($CurrentDir = $DirQueue.Dequeue()){

        # find subfolders of current folder, and if present, add them all to the queue
        Try {
            If ($Directories = [IO.Directory]::GetDirectories($CurrentDir)) {
                $Counter = 0
                foreach ($Dir in $Directories) {
                    $Counter ++
                    Write-Progress -Activity $Actvity -CurrentOperation $Dir -PercentComplete ($Counter/($Directories -as [array]).Count )
                    $DirQueue.Enqueue($Dir)
                }
            }
        } catch {}
        
        # find all files in the folder currently processed
        try {
            If ($Results = [IO.Directory]::GetFiles($CurrentDir, $SearchString) ) {
                Write-Verbose "$($Results.Count) matching file(s) found" 
                Write-Output $Results
            }
            Else {
                Write-Warning "No match found"
            }
           # [IO.Directory]::GetFiles($CurrentDir, $SearchString) 
        } 
        catch {}
    } # While there's something to do
    #>
}
End {
    Write-Progress -Activity $Activity -Complete

    # Console output
    $Msg = "END    : $Activity"
    $FGColor = "Yellow"
    If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg")}
    Else {Write-Verbose $Msg}

}
} # end Get-PKFile

