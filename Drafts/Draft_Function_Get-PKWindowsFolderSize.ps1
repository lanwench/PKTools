#requires -version 3
Function Get-PKWindowsFolderSize {
<#
.SYNOPSIS
    Gets the name, number and size for one or more folders, up to a specified depth

.DESCRIPTION
    Gets the name, number and size for one or more folders, up to a specified depth
    Accepts pipeline input
    Returns a PSObject

.NOTES        
    Name    : Get-GNOpsWindowsMeltdownRegKey.ps1
    Created : 2018-01-29
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2018-01-29 - Created script based on Bill Stewart's original (see link)


.PARAMETER Path
    Specifies a path to one or more file system directories (default is current directory, and wildcards are not permitted)

.PARAMETER LiteralPath
    Specifies a path to one or more file system directories (unlike Path, the value of LiteralPath is used exactly as it is typed)

.PARAMETER Only
    Outputs statistics for a directory, but not any of its subdirectories

.PARAMETER Every
    Outputs statistics for every directory in the specified path instead of only the first level of directories

.PARAMETER FormatNumbers
    Formats numbers in the output object to include thousands separators.

.PARAMETER IncludeTotal
    Outputs a summary object after all other output, summing all statistics

.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Outputs-directory-size-964d07ff

.LINK
    https://stackoverflow.com/questions/46308030/handling-path-too-long-exception-with-new-psdrive

.EXAMPLE
    PS C:\Temp> Get-PKWindowsFolderSize -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                  
        ---                   -----                  
        Verbose               True                   
        Path                  C:\Temp                
        IsLiteralPath         False                  
        Depth                 0                      
        AllSubdirectories     False                  
        NumberFormat          Raw                    
        IncludeTotal          False                  
        SuppressConsoleOutput False                  
        PipelineInput         False                  
        ParameterSetName      Depth                  
        ScriptName            Get-PKWindowsFolderSize
        ScriptVersion         1.0.0                  

        Action: Get folder sizes (output size in bytes)
        VERBOSE: C:\Temp

        Path                            Folders Files     Size
        ----                            ------- -----     ----
        C:\Temp\adlockout                     1     9  2714656
        C:\Temp\adlockout5                    1     0        0
        C:\Temp\adobjpicker                   1     1     1279
        C:\Temp\cygwin64                      1     3   210498
        C:\Temp\dns                           1     3   160278
        C:\Temp\foo                           1     3    12879
        C:\Temp\PSWindowsUpdate               1    13    57957
        C:\Temp\google_profile                1     0        0
        C:\Temp\OSCSpecs                      1    41   164224
        C:\Temp\google_profile2               1     0        0
        C:\Temp\PoshWSUS                      1     5     8644
        C:\Temp\robocopylog_2018-01-23        1     0        0
        C:\Temp\shutdown                      1     2     7606
        C:\Temp\Source                        1     3    24913
        C:\Temp\Tests                         1     2      512

.EXAMPLE
    PS C:\> Get-PKWindowsFolderSize c:\temp\level0 -Depth 2 -NumberFormat Units -IncludeTotal -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                                     Value
        ---                                     -----
        Depth                                       2
        NumberFormat                            Units
        IncludeTotal                             True
        Verbose                                  True
        Path                           c:\temp\level0
        IsLiteralPath                           False
        AllSubdirectories                       False
        SuppressConsoleOutput                   False
        PipelineInput                           False
        ParameterSetName                        Depth
        ScriptName            Get-PKWindowsFolderSize
        ScriptVersion                           1.0.0

        Action: Get folder sizes up to 2 subdirectory level(s); output size in units and include total
        VERBOSE: c:\temp\level0

        Path                                Folders Files Size    
        ----                                ------- ----- ----    
        C:\temp\level0\Level1                     1 5     56.11 KB
        C:\temp\level0\Level1\Level2              1 3     91.68 MB
        C:\temp\level0\Level1\Level2\Level3       1 2     5.01 KB 

#>

[CmdletBinding(
    DefaultParameterSetName="Depth"
)]
param(
    [parameter(
        Position=0,
        Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true,
        HelpMessage = "Path to scan"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Path = (get-location).Path,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Treat path as literal path to scan"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$IsLiteralPath,
    
    [Parameter(
        ParameterSetName="Depth",
        Mandatory = $False,
        HelpMessage = "Depth to scan (default is 0, for top-level only)"
    )]
    [int]$Depth = 0,

    [Parameter(
        ParameterSetName="All",
        Mandatory = $False,
        HelpMessage = "Output statistics for every directory in the specified path instead of only the first level of directories."
    )]
    [Switch]$AllSubdirectories,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Format numbers in the output object to include thousands separators or units"
    )]
    [ValidateSet("Raw","Commas","Units")]
    [String]$NumberFormat = "Raw",
    
    [Parameter(
        Mandatory = $False,
        HelpMessage = "Output a summary object after all other output, summing all statistics"
    )]
    [Switch] $IncludeTotal,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)

begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Current script name
    $ScriptName = $MyInvocation.MyCommand.Name

    # Detect pipeline input and save parametersetname
    $Source = $PSCmdlet.ParameterSetName
    Switch ($Source) {
        Path {$PipelineInput = ( -not $PSBoundParameters.ContainsKey("Path") ) -and ( -not $Path )}
        Default {$PipelineInput = $false}
    }

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    # Don't want this in -Requires as it will prevent the parent module from loading
    If ($PSVersionTable.PSVersion -lt "5.0") {
        $CurrVer = "{0}.{1}.{2}" -f $($PSVersionTable.PSVersion.Major),$($PSVersionTable.PSVersion.Minor),$($PSVersionTable.PSVersion.Build)
        $Msg = "PowerShell version $CurrVer detected`nThis script requires PowerShell 5.0 at minimum (it uses the -Depth parameter now available in Get-ChildItem)"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

    #region Variables for total

    [UInt64[]] $script:totalfolders = 0
    [UInt64[]] $script:totalfiles = 0
    [UInt64[]] $script:totalbytes = 0

    #endregion Variables for total

    #region Inner functions

    # Function to return a [System.IO.DirectoryInfo] object if it exists
    function Get-Directory {
        [CmdletBinding()]
        param($Path,[switch]$IsLiteral)
        $DirectoryObj = $Null
        If ($IsLiteral.IsPresent) {
            If ($Path -match '\\\\?\\' -and ($PSVersionTable.PSVersion -lt "5.1")) {
                $Msg = "'The use of \\?\' in LiteralPath requires PowerShell 5.1 at minimum (you are on $($PSVersionTable.PSVersion.ToString()))`nIllegal characters will be stripped out"
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                $DirectoryObj = $Path.Replace('\\?\',$Null)   
            }
            If ( Test-Path -LiteralPath $Path -PathType Container -ErrorAction SilentlyContinue) {
                $DirectoryObj = Get-Item -LiteralPath $Path -Force -EA Stop
            }
        }
        Else {
            If ( Test-Path -Path $Path -PathType Container -ErrorAction SilentlyContinue ) {
                $DirectoryObj = Get-Item -Path $Path -Force -EA Stop
            }
        }        
        if ( $DirectoryObj -and ($DirectoryObj -is [System.IO.DirectoryInfo]) ) {
            Write-Output $DirectoryObj
        }
    } #end Get-Directory


    # Function to get directory contents
    Function Get-DirectoryContents {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline =$True)]
        $Path,
        [switch]$IsLiteral,
        [switch]$Recurse,
        [int]$Depth
    )
    Process {
     
        $Param_GCI1 = @{
            ErrorAction = "Stop"
            Verbose = $False
            Directory = $True
        }
        $Param_GCI2 = @{
            Path = $Null
            ErrorAction = "SilentlyContinue"
            Verbose = $False
            #Recurse = $True
            #File = $True
        }

        If ($Depth -gt 0) {
            $Param_GCI1.Add("Depth",$Depth)
        }
        Elseif ($Recurse.IsPresent) {
            $Param_GCI1.Add("Recurse",$True)
            $Param_GCI2.Add("Recurse",$True)
        }
        If ($IsLiteral.IsPresent) {
            $Param_GCI1.Add("LiteralPath",$Path)
        }
        Else{
            $Param_GCI1.Add("Path",$Path)
        }
        
        # This is redundant, but we will test the path anyway
        If ($Null = Test-Path $Path -PathType Container -EA SilentlyContinue) {
    
            # Get subdirectories up to the specified depth
            $Directories = Get-ChildItem @Param_GCI1 | Sort-Object
            
            $Total = $Directories.Count
            $Current = 0

            # Recurse through subdirectories
            $Output = Foreach ($Dir in ($Directories | Sort-Object FullName)) {
                $Current ++
                Write-Progress -Activity "Get directory contents for $Total subfolder(s)" -CurrentOperation $Dir.FullName -PercentComplete ($Current/$Total*100)
                
                $Param_GCI2.Path = $Dir.FullName

                If ($subFolderItems = Get-ChildItem @Param_GCI2) {
    
                    $FileCount = $subFolderItems | Where-Object {(-not $_.PSIsContainer)}.Count
                    If (-not ($Sum = ($subFolderItems | Measure-Object -property Length -sum).Sum)) {$Sum = 0}
                    If (-not ($DirCount = ($subFolderItems | Where-Object {$_.PSIsContainer} | Select -Unique).Count)) {$DirCount = 0}

                    New-Object PSObject -Property ([ordered]@{
                        Path    = $Dir.FullName
                        Folders = $DirCount
                        Files   = $FileCount
                        Size    = $Sum
                    })
                }
                Else {
                    New-Object PSObject -Property ([ordered]@{
                        Path    = $Dir.FullName
                        Folders = $DirCount
                        Files   = 0
                        Size    = 0
                    })
                }
            } # end for each top level folder
            #Write-Progress -Activity $Complet
            Write-Output $Output| Select Path,Folders,Files,Size

        } #end if dir exists
    }
    } #end Get-DirectoryContents


    # Function to format numbers (comma separator or unit)
    function Format-Number([switch]$Unit) {
        Process{
            If ($Unit.IsPresent) {
                Function GetUnit ($Size) {
                    If ($Size -gt 0) {
                        switch -Regex ([math]::truncate([math]::log($Size,1024))) {
                            '^0' {"{0:n2} B"  -f $Size}
                            '^1' {"{0:n2} KB" -f ($Size / 1KB)}
                            '^2' {"{0:n2} MB" -f ($Size / 1MB)}
                            '^3' {"{0:n2} GB" -f ($Size / 1GB)}
                            '^4' {"{0:n2} TB" -f ($Size / 1TB)}
                            '^5' {"{0:n2} PB" -f ($Size / 1PB)}
                            Default {"{0:n2} PB" -f ($Size / 1pb)}
                        }
                    }
                    Else {"0.00 B"}
                }
                $_ | Select-Object Path,
                    @{Name="Files"   ; Expression={"{0:N0}" -f $_.Files}},
                    @{Name="Folders" ; Expression={$_.Folders}},
                    @{Name="Raw"     ; Expression={$_.Size}},
                    @{Name="Size"    ; Expression={"{0:N0}" -f (GetUnit $_.Size)}}
            }
            Else {
                $_ | Select-Object Path,
                    @{Name="Files"   ; Expression={"{0:N0}" -f $_.Files}},
                    @{Name="Folders" ; Expression={$_.Folders}},
                    @{Name="Raw"     ; Expression={$_.Size}},
                    @{Name="Size"    ; Expression={"{0:N0}" -f $_.Size}}
            }
        }
    } #end Format-Number

    #endregion Inner functions

    #region Splats

    # For Write-Progress
    Switch ($Source) {
        Depth {
            If ($Depth -gt 0) {$Activity = "Get folder sizes up to $Depth subdirectory level(s)"}
            Else {$Activity = "Get folder sizes"}
        }
        All {$Activity = "Get folder sizes"}
    }

    # Splat for Write-Progress
    $Param_WP = @{
        Activity         = $Activity
        Status           = "Working"
        CurrentOperation = $Null
        PercentComplete  = $Null
    }

    # Splat for Get-Directory
    $Param_GetDir = @{}
    $Param_GetDir = @{
        Path = $Null
        ErrorAction = "SilentlyContinue"
        Verbose = $False
    }

    # Splat for Get-DirectoryContents
    $Param_GetDirContents = @{}
    $Param_GetDirContents = @{
        ErrorAction = "SilentlyContinue"
        Verbose = $True
    }
    Switch ($Source) {
        Depth {
            If ($Depth -gt 0) {$Param_GetDirContents.Add("Depth",$Depth)}
        }
        All {$Param_GetDirContents.Add("Recurse",$True)}
    }

    #endregion Splats

   # Console output
    Switch ($NumberFormat) {
        Raw    {$Activity += "; output size in bytes"}
        Commas {$Activity += "; output size with thousands separator"}
        Units  {$Activity += "; output size in units"}
    }
    If ($IncludeTotal.IsPresent) {
        $Activity += " and include total"
    }

    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}

Process {
    
    # Get the item to process, whether the input comes from the pipeline or not
    If ( $PipelineInput ) {
        $Item = $_
    }
    else {
        $Item = $Path
    }
    
    $Total = $Item.Count
    $Current = 0

    $Msg = $Item
    Write-Verbose $Msg
    $Param_WP.PercentComplete
    $Param_WP.CurrentOperation = $Msg

    Try {
        $Param_GetDir.Path = $Item

        #$Param_GetDir.ErrorAction = "Inquire"
        #$Param_GetDirContents.ErrorAction = "Inquire"
        
        If ($Dir = Get-Directory @Param_GetDir)  {

            Try {
                $Output = @()
                $Output = $Dir | Get-DirectoryContents @Param_GetDirContents
                $script:totalfolders += $Output.Folders
                $script:totalfiles += $Output.Files
                $script:totalbytes += $Output.Size
                
                Switch ($NumberFormat) {
                    Raw    {Write-Output ($Output | Select Path,Folders,Files,Size)}
                    Commas {Write-Output ($Output | Format-Number | Select Path,Folders,Files,Size)}
                    Units  {Write-Output ($Output | Format-Number -Unit | Select Path,Folders,Files,Size)}
                }
            }
            Catch {
                $Msg = "Failed to get directory contents for '$Item'"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Else {
            $Msg = "Failed to get directory '$Item'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine($Msg)
        }

    }
    Catch {
        $Msg = "Failed to get directory '$P'"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
        $Host.UI.WriteErrorLine($Msg)
    }

}

End {
    
    Write-Progress -Activity $Activity -Completed

    If ($IncludeTotal.IsPresent) {

        $TotFolders = ($script:totalfolders | Measure-Object -sum ).sum
        $TotFiles   = ($script:totalfiles   | Measure-Object -sum ).sum
        $TotBytes   = ($script:totalbytes   | Measure-Object -sum ).sum

        $TotalObj = New-Object PSObject -Property ([ordered]@{
            Path = "<Total>"
            Folders = $TotFolders
            Files = $TotFiles
            Size = $TotBytes
        })

        Switch ($NumberFormat) {
            Raw    {Write-Output ($TotalObj | Select Path,Folders,Files,Size)}
            Commas {Write-Output ($TotalObj | Format-Number | Select Path,Folders,Files,Size)}
            Units  {Write-Output ($TotalObj | Format-Number -Unit | Select Path,Folders,Files,Size)}
        }
    }

}
} #end Get-PKWindowsFolderSize

<#

$Path = "\\?\c:\users\pkingsley\appdata"
[switch]$IsLiteralPath = $True
[switch]$IncludeTotal = $True
[switch]$AllSubdirectories = $True
$NumberFormat = "Unit"


$Path,$IsLiteralPath,$IncludeTotal,$AllSubdirectories,$NumberFormat

$RemoteDef = "function Get-PKWindowsFolderSize { ${function:Get-PKWindowsFolderSize} }"
Invoke-Command -ArgumentList $RemoteDef, -AsJob -ComputerName ops-pktest-1 -JobName mymymy -ScriptBlock {
    . ([ScriptBlock]::Create($RemoteDef))
    Get-PKWindowsFolderSize -Path "\\?\c:\users\pkingsley\appdata" -IsLiteralPath -IncludeTotal -AllSubdirectories -NumberFormat Units
}

#>