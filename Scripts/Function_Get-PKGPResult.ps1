#requires -version 4
Function Get-PKGPResult {
<#
.SYNOPSIS 
    Executes gpresult to an HTML or XML file, with the option to modify html font name/size and launch the file with the associated handler

.DESCRIPTION
    Executes gpresult to an HTML or XML file, with the option to modify html font name/size and launch the file with the associated handler
    Defaults to HTML format if -XML not specified
    If current PowerShell session is not elevated: prompts to run as elevated, unless -NoElevationPrompt is present, in which case only user information is returned

.OUTPUTS
    File

.NOTES
    Name    : Function_Get-PKGPResult.ps1
    Created : 2024-04-03
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2024-04-03 - Created script
        v01.01.0000 - 2025-10-27 - Changed FontName to FontFamily with validateset & HTML will select best available; other cosmetic changes

.PARAMETER OutputPath
    Absolute or UNC path for output file (default is user's temp directory)

.PARAMETER FontSizeIncrease
    For HTML report, an integer between 1 and 10 by which to increment all font size percentages

.PARAMETER FontName
    For HTML report, name of replacement font (must exist on system)

.PARAMETER NoElevationPrompt
    Run gpresult in current session, instead of prompting for elevation if not present

.PARAMETER XML
    Generate XML file instead of default HTML

.EXAMPLE
    PS C:\> Get-PKGPResult -Verbose -FontName Verdana -OutputPath c:\temp -FontSizeIncrease 5
        VERBOSE: PSBoundParameters: 

        Key              Value
        ---              -----
        Verbose          True
        FontName         Verdana
        OutputPath       c:\temp
        FontSizeIncrease 5
        XML              False
        ScriptName       Get-PKGPResult
        ScriptVersion    1.0.0

        VERBOSE: [USTI9EKM0PQ] Generating group policy results to file c:\temp\GPResult_USTI9EKM0PQ_2024-04-16-15-16-15.html
        VERBOSE: [USTI9EKM0PQ] Current PowerShell session is not elevated; would you like to launch gpresult in a separate elevated process?
        Confirm
        Are you sure you want to perform this action?
        Performing the operation "Current PowerShell session is not elevated; would you like to launch gpresult in a separate elevated process?" on target "USTI9EKM0PQ".
        [Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "Y"):
        VERBOSE: [USTI9EKM0PQ] Please wait...
        VERBOSE: [USTI9EKM0PQ] Successfully created C:\temp\GPResult_USTI9EKM0PQ_2024-04-16-15-16-15.html
        VERBOSE: [USTI9EKM0PQ] Increasing all font size percentages by 5
        VERBOSE: [USTI9EKM0PQ] Changing all font family references to Verdana
        VERBOSE: [USTI9EKM0PQ] Open c:\temp\GPResult_USTI9EKM0PQ_2024-04-16-15-16-15.html with default application handler?
        Confirm
        Are you sure you want to perform this action? 
        Performing the operation: "Open c:\temp\GPResult_USTI9EKM0PQ_2024-04-16-15-16-15.html with default application handler?"
        [Y] Yes  [N] No  [S] Suspend  [?] Help (default is "Y"):

#>
[CmdletBinding(DefaultParameterSetName = "Raw", SupportsShouldProcess, ConfirmImpact = "High")]
Param(
    [Parameter(
        HelpMessage = "Absolute or UNC path for output file (default is user's temp directory)",
        Position = 0
    )]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$OutputPath = $Env:Temp,

    [Parameter(
        HelpMessage = "Generate XML file instead of default HTML",    
        ParameterSetName = "XML"
    )]
    [switch]$XML,

    [Parameter(
        ParameterSetName = "HTML",
        HelpMessage = "For HTML reports increase font size by specified percentages to improve readability; default is 15 (note that larger sizes can cause problems!)"
    )]
    [ValidateSet("01%","03%","05%","07%","10%","15%","20%","30","None")]
    [string]$FontSizeIncrease = "15%",

    [Parameter(
        ParameterSetName = "HTML",
        HelpMessage = "For an HTML report, the style of font (Default, Serif, SansSerif, Monospace) - default is SansSerif"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet('Default','Serif','SansSerif','Monospace')]
    [string]$FontFamily = "SansSerif",

    [Parameter(
        HelpMessage = "Run gpresult in current session, instead of prompting for elevation if not present (in which case only user results will be returned)"
    )]
    [switch]$NoElevationPrompt

)
Begin {
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"
    
    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters

    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
    Where-Object { Test-Path variable:$_ } | ForEach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("ParameterSetName", $Source)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$(($CurrentParams.GetEnumerator() | Sort-Object) | Format-Table -AutoSize | out-string )"

    # variables & session check
    If ($XML.IsPresent) { $Extension = "xml" }
    Else { $Extension = "html" }
    $FileName = "GPResult_$Env:ComputerName`_$(Get-Date -f "yyyy-MM-dd-HH-mm-ss").$Extension"
    $OutFile = "$OutputPath\$FileName"
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    # Fonts
    If ($FontFamily -ne "Default") {
        switch ($FontFamily) {
            'Serif'     { $ValidFont = "Georgia, 'Times New Roman', Times, serif" }
            'SansSerif' { $ValidFont = "Segoe UI, Roboto, 'Helvetica Neue', Arial, sans-serif" }
            'Monospace' { $ValidFont = "'Consolas', 'Courier New', 'Liberation Mono', monospace" }
        }
        Write-Verbose "[$Env:ComputerName] Using font family: $($ValidFont.replace("'",$Null) -join(", "))"
    }
    
    # Font size
    If ($FontSizeIncrease -ne "None") {
        #[int]$Increase = $FontSizeIncrease.TrimEnd("%")
        $Increase = $FontSizeIncrease.TrimEnd("%") -as [int]
        Write-Verbose "[$Env:ComputerName] Increasing font size by $FontSizeIncrease"
    } 
    Write-Verbose "[BEGIN: $ScriptName] Generate GPResult with custom formatting"
}
Process {
    $Msg = "Generating group policy results to file $OutFile"
    Write-Verbose "[$Env:ComputerName] $Msg"
    Try {
        
        If (-not $IsAdmin) {
            $Msg = "Current PowerShell session is not elevated" 
            If ($NoElevationPrompt.IsPresent) {
                $Msg += " (-NoElevationPrompt detected)"
                Write-Verbose "[$Env:ComputerName] $Msg"
                $Null = gpresult.exe /h  $OutFile
            }
            Else {
                Write-Warning "[$Env:ComputerName] $Msg"
                $ConfirmMsg = "Launch gpresult in a separate elevated process?"
                If ($PSCmdlet.ShouldProcess($Env:ComputerName, $ConfirmMsg)) {
                    $Msg = "Please wait..."
                    Write-Verbose "[$Env:ComputerName] $Msg"
                    $Null = Start-Process cmd.exe -Verb runas -Wait -WindowStyle Hidden -PassThru -ArgumentList "/c gpresult.exe /h $OutFile"
                }
                Else {
                    $ConfirmMsg = "Run gpresult in non-elevated context?"
                    If ($PSCmdlet.ShouldProcess($Env:ComputerName, $ConfirmMsg)) {
                        $Msg = "Please wait..."
                        Write-Verbose "[$Env:ComputerName] $Msg"
                        $Null = gpresult.exe /h  $OutFile
                    }
                    Else {
                        $Msg = "Operation cancelled by user"
                        Write-Verbose "[$Env:ComputerName] $Msg"
                        Break
                    }
                }
            }
        }
        Else {
            $Null = gpresult.exe /h  $OutFile
        }
        
        Try {
            $FileObj = Get-Item $OutFile -ErrorAction Stop
            $Msg = "[$Env:ComputerName] Successfully created $($FileObj.FullName)"
            Write-Verbose $Msg

            Try {
                If ($ValidFont -or $Increase) {
                    $Content = Get-Content -Path $OutFile -Raw    # ...getting content as a single string to improve performance
                }

                If ($Increase -gt 0) {
                    $Msg = "Increasing all font size percentages by $FontSizeIncrease"
                    Write-Verbose "[$Env:ComputerName] $Msg"
                    # Thanks, Jeff Hicks! Using named captures to find all references to the font-size & increment the percentage, updating the file
                    [regex]$rx = "font-size:\s*(?<size>\d+(?=%))"
                    $values = $rx.matches($Content).groups | Where-Object { $_.Name -eq 'size' } | Select-Object Value -Unique
                    foreach ($item in $values) {
                        $new = $n + ([int]$item.value)
                        $Content = $Content -replace "font-size:$($item.Value)", "font-size:$new"
                    }
                    $Content | Out-File $OutFile -ErrorAction Stop -Confirm:$False -Verbose:$False
                }
                
                If ($ValidFont) {
                    # Ensure replacement is a single CSS font-family string
                    $FontReplacement = if ($ValidFont -is [array]) { ($ValidFont -join ", ") } else { $ValidFont }

                    $Msg = "Changing all font family references to $FontReplacement"
                    Write-Verbose "[$Env:ComputerName] $Msg"

                    # Match the CSS font-family value (up to the next semicolon) and trim whitespace/quotes
                    [regex]$rx = "(?<=font-family\s*:\s*)(?<font>[^;]+)"
                    $OldFonts = $rx.Matches($Content) | ForEach-Object { $_.Groups['font'].Value.Trim() } | Select-Object -Unique

                    foreach ($OldFont in $OldFonts) {
                        # Replace only the font-family value (preserves surrounding CSS)
                        $escaped = [regex]::Escape($OldFont)
                        $Content = [regex]::Replace($Content, "(?<=font-family\s*:\s*)$escaped", $FontReplacement)
                    }

                    $Content | Out-File $OutFile -ErrorAction Stop -Confirm:$False -Verbose:$False
                }
                
                $Msg = "Open $Outfile with default application handler?"
                Write-Verbose "[$Env:ComputerName] $Msg"
                If ($PSCmdlet.ShouldContinue($Msg, $Env:ComputerName)) {
                    Invoke-Item $OutFile -ErrorAction Stop -Confirm:$False -Verbose:$False
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    Write-Verbose "[$Env:ComputerName] $Msg"
                }
            }
            Catch {
                $Msg = "Operation failed! $($_.Exception.Message)"
                Write-Warning "[$Env:ComputerName] $Msg"
            }
        }
        Catch {
            $Msg = "Unable to generate $Outfile! $($_.Exception.Message)"
            Write-Warning "[$Env:ComputerName] $Msg"
        }
    }
    Catch {
        $Msg = "Failed to generate group policy results! $($_.Exception.Message)"
        Write-Warning "[$Env:ComputerName] $Msg"
    }
}
End {
    Write-Verbose "[END: $Scriptname] Script operation complete"
}
} #end Get-PKGPResult