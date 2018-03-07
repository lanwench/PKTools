# Draft_Function-GetPKWindowsTreeSize
# https://www.reddit.com/r/PowerShell/comments/66nsr9/treesize_script/

function Get-PKFolderSize{
<##>
[CmdletBinding()]
param(
    [parameter(
        Position = 0,
        Mandatory=$False,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage = "Target folder (default is current user home directory)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$TargetFolder,

    [parameter(
        Position = 1,
        Mandatory=$False,
        HelpMessage = "Levels to recurse in target (default is 1, maximum is 10)"
    )]
    [ValidateRange(1,10)]
    [int]$Depth = 1,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Return size formatted as 123.45 MB (default is Double)"
    )]
    [switch]$FormatSize,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Detect pipeline input and save parametersetname
    $Source = $PSCmdlet.ParameterSetName
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("TargetFolder")) -and ((-not $TargetFolder)) # -or (-not $InputObj -eq $Env:ComputerName))

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # If we didn't supply anything
    If (-not $TargetFolder) {$TargetFolder = $Home}

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    #Splat for Write-Progress
    $Activity = "Get folder sizes up to $Depth level(s)"
    $Param_WP = @{
        Activity         = $Activity
        Status           = "Working"
        CurrentOperation = $Null
        PercentComplete  = $Null
    }

    # Inner function
    function Get-DisplayFileSize {
        # Based on http://community.idera.com/powershell/powertips/b/tips/posts/formatting-numbers-part-1
        param([Double]$Number)
        $NewNumber = $Number
        $Unit = ',KB,MB,GB,TB,PB,EB,ZB' -split ','
        $i = $null
        While ($newNumber -ge 1KB -and $i -lt $Unit.Length){
            $newNumber /= 1KB
           $i++
        }
        If ($i -eq $null) {
            Write-Output $Number
        }
        Else {
           $DisplayText = "'{0:N2} {1}'" -f $NewNumber , $unit[$i]
           $Output = $Number | Add-Member -MemberType ScriptMethod -Name ToString -Value ([Scriptblock]::Create($DisplayText)) -Force -PassThru
           Write-Output $Output
       }
    }


    Function GetSize {
        Param($TargetFolder,$Depth,$FormatSize)
        Try {
            $TargetObj = Get-Item $TargetFolder -ErrorAction Stop
            $Output = New-Object PSObject -Property ([ordered] @{
                Name = $TargetObj.FullName
                NumItems = 0
                Size = 0
                Subs = @()
            })

            Switch ($Depth) {
                1 {
                    $Subs = (Get-ChildItem $TargetObj.FullName -Force -File -ErrorAction SilentlyContinue)
                    $Output.NumItems = $Subs.Count
                    $Size = ($Subs | Measure-Object -Sum -Property Length).Sum
                    $Output.Size = $Size
                    Write-Output $Output
                }
                Default {
                    $Subs = Get-ChildItem $TargetObj.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    $Output.NumItems = $Subs.Count
                    $Output.Subs = Foreach ($S in $Subs){
                        Write-Verbose
                        If ($S.PSIsContainer){
                            $Tmp = $S.FullName | Get-PKFolderSize -depth ($Depth -1) -Verbose:$False
                            $Output.Size += $tmp.Size
                            Write-Output $Tmp
                        }
                        else{
                            $Output.Size += $S.length
                        }
                    }
                    Write-Output $Output
                }
            }
        }
        Catch {
            $Msg = "Failed to get target '$Target'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine($Msg)
        }
    } #end inner function

}

Process {

    $Total = $TargetFolder.Count
    $Current = 0

    Foreach ($Target in $TargetFolder) {
    
        $Current ++ 
        $Msg = $Target
        
        $Param_WP.CurrentOperation = $Msg
        $Param_WP.PercentComplete = $Current/$Total*100
        Write-Progress @Param_WP

        
        Try {
            $TargetObj = Get-Item $Target -ErrorAction Stop

            $Msg = $TargetObj.FullName
            $Param_WP.CurrentOperation = $Msg
            Write-Progress @Param_WP
            Write-Verbose $Msg

            $Output = New-Object PSObject -Property ([ordered] @{
                Name = $TargetObj.FullName
                NumItems = 0
                Size = 0
                Subs = @()
            })


            $SubItems = Get-ChildItem $TargetObj.FullName -Force -Recurse -Depth $Depth -File -ErrorAction SilentlyContinue
            $Output.NumItems = $SubsItems.Count
            $Size = ($SubItems | Measure-Object -Sum -Property Length).Sum
            If ($FormatSize.IsPresent) {$Output.Size = Get-DisplayFileSize -Number $Size}
            Else {$Output.Size = $Size}
            Write-Output $Output

            <#

            Foreach ($Item in $SubItems) {
            
            
            }



            Switch ($Depth) {
                1 {
                    $Subs = (Get-ChildItem $TargetObj.FullName -Force -File -ErrorAction SilentlyContinue)
                    $Output.NumItems = $Subs.Count
                    $Size = ($Subs | Measure-Object -Sum -Property Length).Sum
                    If ($FormatSize.IsPresent) {$Output.Size = Get-DisplayFileSize -Number $Size}
                    Else {$Output.Size = $Size}
                    Write-Output $Output
                }
                Default {
                    #we are not at the depth limit, keep recursing
                    $Output.Subs = Foreach ($S in Get-ChildItem $TargetObj.FullName -Force -ErrorAction SilentlyContinue){
                        If ($S.PSIsContainer){
                            $Tmp = $S.FullName | Get-PKFolderSize -depth ($Depth -1) -Verbose:$False
                            $Output.Size += $tmp.Size
                            Write-Output $Tmp
                        }
                        else{
                            $Output.Size += $S.length
                        }
                    }
                    Write-Output $Output
                }
            }



            #>
        }

        Catch {
            $Msg = "Failed to get target '$Target'"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += "`n$ErrorDetails"}
            $Host.UI.WriteErrorLine($Msg)
        }

    } #end foreach

}

End {

}
}




function Print-Results{
    param(
        [parameter(mandatory=$true, position=1)]
        $Data,

        [parameter(mandatory=$true, position=2)]
        [int]$IndentDepth        
    )



    "{0:N2}" -f ($Data.Size / 1GB) + " GB",  (' '*($IndentDepth+2)) + $Data.Name

    foreach($S in $Data.Subs){
        Print-Results $S ($IndentDepth+1)
    }
}




function getSubFolderSizes{
    [cmdletbinding()]
    param(
        [parameter(mandatory=$true, position=1)]
        [string]$targetFolder,

        [int]$DepthLimit=3
    )


    if(-not (Test-Path $targetFolder)){
        Write-Error "The target [$targetFolder] does not exist"
        exit
    }

}

<#

$Data=Get-SizeInfo $targetFolder $DepthLimit

    #returning $data will provide a useful PS object rather than plain text
    #return $Data

    #generate a human fraindly listing instead
    Print-Results $Data 0

#>