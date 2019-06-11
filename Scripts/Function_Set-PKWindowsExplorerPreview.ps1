#requires -Version 3
Function Set-PKWindowsExplorerPreview {
<#
.SYNOPSIS
    Enables the local computer's Windows Explorer preview pane for an additional file type
    
.DESCRIPTION
    Enables the local computer's Windows Explorer preview pane for an additional file type
    by modifying the registry to set the percieved file type/content type as text
            
.NOTES
    Name    : Function_Set-PKWindowsExplorerPreview.ps1
    Author  : Paula Kingsley
    Created : 2017-08-22
    Version : v01.01.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v1.0.0      - 2017-08-22 - Created script
        v01.01.0000 - 2019-03-11 - Fixed erroneous Win10 message, other general updates
        v01.02.0000 - 2019-06-11 - Renamed to Set-PKWindowsExplorerPreview, fixed issue where it ...
                                   um...didn't actually do anything to change settings.
        
.LINK
    https://blogs.technet.microsoft.com/bshukla/2010/03/30/script-to-enable-preview-pane-for-powershell-scripts/

.EXAMPLE
    PS C:\> Set-PKWindowsExplorerPreview -FileExtension ps1,psd1,psm1 -Verbose

        VERBOSE: PSBoundParameters: 
 	
        Key           Value                       
        ---           -----                       
        FileExtension {ps1, psd1, psm1}           
        Verbose       True                        
        Force         False                       
        Quiet         False                       
        ScriptName    Set-PKWindowsExplorerPreview
        ScriptVersion 1.2.0                       

        VERBOSE: [Prerequisites] Check PowerShell session mode
        VERBOSE: [Prerequisites] PowerShell is running in Elevated mode
        
        WARNING: This script tells Windows to treat the specified file extension as a text file,
        allowing it to be viewed in the Windows Explorer preview pane.
        It does not verify that the file type is compatible with this setting; please proceed with caution!

        BEGIN  : Enable Windows Explorer preview pane to view file extensions

        [Registry::HKEY_CLASSES_ROOT\.ps1] Check registry for current settings
        [Registry::HKEY_CLASSES_ROOT\.ps1] PerceivedType for extension already set to 'text'

        Extension    : ps1
        Path         : Registry::HKEY_CLASSES_ROOT\.ps1
        IsSuccessful : False
        ComputerName : PKINGSLEY-05122
        Messages     : PerceivedType for extension already set to 'text'

        [Registry::HKEY_CLASSES_ROOT\.psd1] Check registry for current settings
        [Registry::HKEY_CLASSES_ROOT\.psd1] Set PerceivedType for extension to 'text' for Windows Explorer preview pane
        [Registry::HKEY_CLASSES_ROOT\.psd1] Successfully set PerceivedType for extension to 'text'
        
        Extension    : psd1
        Path         : Registry::HKEY_CLASSES_ROOT\.psd1
        IsSuccessful : True
        ComputerName : PKINGSLEY-05122
        Messages     : Successfully set PerceivedType for extension to 'text'

        [Registry::HKEY_CLASSES_ROOT\.psm1] Check registry for current settings
        [Registry::HKEY_CLASSES_ROOT\.psm1] PerceivedType for extension already set to 'text'
        
        Extension    : psm1
        Path         : Registry::HKEY_CLASSES_ROOT\.psm1
        IsSuccessful : False
        ComputerName : PKINGSLEY-05122
        Messages     : PerceivedType for extension already set to 'text'

        END    : Enable Windows Explorer preview pane to view file extensions

.EXAMPLE
    PS C:\> Set-PKWindowsExplorerPreview -FileExtension md -Verbose

        VERBOSE: PSBoundParameters: 
 	
        Key           Value                       
        ---           -----                       
        FileExtension {md}                      
        Verbose       True                        
        Force         False                       
        Quiet         False                       
        ScriptName    Set-PKWindowsExplorerPreview
        ScriptVersion 1.2.0                       

        VERBOSE: [Prerequisites] Check PowerShell session mode
        VERBOSE: [Prerequisites] PowerShell is running in Elevated mode
        WARNING: This script tells Windows to treat the specified file extension as a text file,
        allowing it to be viewed in the Windows Explorer preview pane.
        It does not verify that the file type is compatible with this setting; please proceed with caution!

        BEGIN  : Enable Windows Explorer preview pane to view file extensions

        [Registry::HKEY_CLASSES_ROOT\.md] Check registry for current settings
        [Registry::HKEY_CLASSES_ROOT\.md] Set PerceivedType for extension to 'text' for Windows Explorer preview pane
        [Registry::HKEY_CLASSES_ROOT\.md] Operation cancelled by user

        Extension    : md
        Path         : Registry::HKEY_CLASSES_ROOT\.md
        IsSuccessful : False
        ComputerName : PKINGSLEY-05122
        Messages     : Operation cancelled by user

        END    : Enable Windows Explorer preview pane to view file extensions


.EXAMPLE
    PS C:> Set-PKWindowsExplorerPreview -FileExtension ps1 -Force -Quiet

        WARNING: This script tells Windows to treat the specified file extension as a text file,
        allowing it to be viewed in the Windows Explorer preview pane.
        It does not verify that the file type is compatible with this setting; please proceed with caution!

        Extension    : ps1
        Path         : Registry::HKEY_CLASSES_ROOT\.ps1
        IsSuccessful : True
        ComputerName : PKINGSLEY-05122
        Messages     : Successfully set PerceivedType for extension to 'text'


#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        ValueFromPipeline = $True,
        Mandatory = $True,
        HelpMessage = "File extension (e.g., 'txt' or '.txt')"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$FileExtension,

    [Parameter(
        HelpMessage = "Force registry change"
    )]
    [switch]$Force,

    [Parameter(
        HelpMessage = "Suppress all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.02.0000"

    
    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.FileExtension = $FileExtension
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n `t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # General-purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }
    
    #region Inner functions
    
    # Function to write a console message or a verbose message
    Function Write-MessageInfo {
        Param([Parameter(ValueFromPipeline)]$Message,$FGColor,[switch]$Title)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {
            If ($Title.IsPresent) {$Message = "`n$Message`n"}
            $Host.UI.WriteLine($FGColor,$BGColor,"$Message")
        }
        Else {Write-Verbose "$Message"}
    }

    # Function to write an error or a verbose message
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)#,[switch]$Quiet = $Quiet)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        Else {Write-Verbose "$Message"}
    }

    # Function to check if PowerShell is running in elevated mode
    function Test-Elevated{
        [CmdletBinding()]
        $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
        $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
        $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
        [switch]$IsAdmin=$prp.IsInRole($adm)
        if ($IsAdmin.IsPresent)  {
            $Msg = "[Prerequisites] PowerShell is running in Elevated mode"
            Write-Verbose $Msg
            $True
        }
        Else {
            #$Msg = "PowerShell must be running in Elevated mode; please re-launch as Administrator"
            $Msg = "[Prerequisites] PowerShell is not running in Elevated mode"
            #$Host.UI.WriteErrorLine("ERROR: $Msg")
            Write-Warning $Msg
            $False
        }
    }

    #endregion Inner functions

    #region Prerequisites

    $Msg = "Check PowerShell session mode"    
        Write-Verbose "[Prerequisites] $Msg"
    If (-not (Test-Elevated -Verbose:$False)) {
        $Msg = "[Prerequisites] PowerShell must be running in Elevated mode; please re-launch as Administrator"

    }

    #[Switch]$Continue = $False
    <#
    If (-not $SkipOSVersionCheck.IsPresent) {
        $Msg = "Checking OS version compatibility"    
        Write-Verbose "[Prerequisites] $Msg"
        
        Try {
            $OS = (Get-WmiObject -Class win32_OperatingSystem @StdParams).caption
            switch -wildcard ($OS){
                "*Windows 7*" {
                    $Msg = "Verified OS $OS"
                    Write-Verbose "[Prerequisites] $Msg"
                    $Null = Test-Elevated
                    $Continue = $True
                }
                "*Windows Server 2008 R2*" {
                    $Msg = "Verified OS $OS"
                    Write-Verbose "[Prerequisites] $Msg"
                    $Null = Test-Elevated
                    $Continue = $True
                }
                default {
                    $Msg = "$Env:ComputerName is running $OS. This function requires Windows 7 or Windows Server 2008 R2. "
                    $Host.UI.WriteErrorLine("ERROR  : [Prerequisites] $Msg")
                    Break
                }
            }
        }
        Catch {
            $Msg = "OS version test failed"
            $ErrorDetails = $_.Exception.Message
            $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
            Break
        }
    }
    Else {$Continue = $True}

    If (-not $Continue.IsPresent) {
        Break
    }
    #>

    #endregion Prerequisites

    #region Output object

    $OutputTemplate = [pscustomobject]@{
        Extension    = $Null
        Path         = "Error"
        IsSuccessful = "Error"
        ComputerName = $Env:ComputerName
        Messages     = "Error"
    }

    #endregion Output object

    #region Splats

    # Splat for Get-ItemProperty
    $Param_GetItem = @{}
    $Param_GetItem = @{
        Path        = $Null
        Verbose     = $False
        ErrorAction = "SilentlyContinue"
    }

    # Splat for Set-ItemProperty
    $Param_SetItem = @{}
    $Param_SetItem = @{
        Path        = $Null
        Name        = "PerceivedType"
        Value       = "text"
        Passthru    = $True
        Force       = $True
        Confirm     = $False
        Verbose     = $False
        ErrorAction = "Stop"
    }

    # Splat for Write-Progress
    $Activity = "Add file extension to Windows Explorer preview pane"
    $Param_WP1 = @{}
    $Param_WP1 = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = $Null
        PercentComplete  = $Null
    }

    #endregion Splats

    $Msg = "This script tells Windows to treat the specified file extension as a text file,`nallowing it to be viewed in the Windows Explorer preview pane.`nIt does not verify that the file type is compatible with this setting; please proceed with caution!"
    Write-Warning $Msg
    
    $Activity = "Enable Windows Explorer preview pane to view file extensions"   
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}

Process {
    

    $Total = $FileExtension.Count
    $Current = 0

    Foreach ($Ext in $FileExtension) {
        
        If ($Ext -match "^.") {$Ext = $Ext.Replace(".",$Null)}
        $Path = "Registry::HKEY_CLASSES_ROOT\.$Ext" 

        $Msg = "Check registry for current settings"
        "[$Path] $Msg" | Write-MessageInfo -FGColor White

        $Current ++
        $Param_WP.CurrentOperation = $Msg
        $Param_WP.Status = ".$Ext"
        $Param_WP.PercentComplete = ($Current/$Total*100)
        Write-Progress @Param_WP

        $Output = $OutputTemplate.PSObject.Copy()
        $Output.Extension = "$Ext"
        $Output.Path = $Path

        [switch]$Continue = $False

        Try {
            $Param_GetItem.Path = $Path
            If ($GetReg = Get-ItemProperty @Param_GetItem) {
                If ($GetReg.PerceivedType -eq "Text") {
                    
                    If ($Force.IsPresent) {
                        $Msg = "PerceivedType for extension already set to 'text'; -Force detected"
                        "[$Path] $Msg" | Write-MessageInfo -FGColor White
                        $ConfirmMsg = "`n`nContinue to change PerceivedType`n`n"
                        If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
                            $Continue = $True
                        }
                        Else {
                            $Msg = "Operation cancelled by user"
                            "[$Path] $Msg" | Write-MessageInfo -FGColor Red
                            $Output.IsSuccessful = $False
                            $Output.Messages = $Msg
                        }
                    }
                    Else {
                        $Msg = "PerceivedType for extension already set to 'text'"
                        "[$Path] $Msg" | Write-MessageInfo -FGColor Green
                        $Output.IsSuccessful = $False
                        $Output.Messages = $Msg
                    }   
                }
                Else {
                    $Continue = $True
                }
            } # if found
            Else {
                $Continue = $True
            }
            
            If ($Continue.IsPresent) {
            
                $Msg = "Set PerceivedType for extension to 'text' for Windows Explorer preview pane"
                "[$Path] $Msg" | Write-MessageInfo -FGColor White
                $ConfirmMsg = "`n`nSet PerceivedType for extension .$Ext to 'text' for Windows Explorer preview pane`n`n"
                
                If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg )) {
                    Try {
                        $Param_SetItem.Path = $Path
                        $Setit = Set-ItemProperty @Param_SetItem
                        If ($SetIt.PerceivedType -eq 'text') {
                            $Msg = "Successfully set PerceivedType for extension to 'text'"
                            "[$Path] $Msg" | Write-MessageInfo -FGColor Green
                            $Output.IsSuccessful = $True
                            $Output.Messages = $Msg
                        }
                    }
                    Catch {
                        $Msg = "Failed to change PerceivedType to 'text'"
                        "[$Path] $Msg" | Write-MessageError
                        $Output.IsSuccessful = $False
                        $Output.Messages = $Msg
                    }
                }
                Else {
                    $Msg = "Operation cancelled by user"
                    "[$Path] $Msg" | Write-MessageInfo -FGColor Red
                    $Output.IsSuccessful = $False
                    $Output.Messages = $Msg
                }
            }
        }
        Catch {
            $Msg = "Failed to get registry data"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
            "[$Path] $Msg" | Write-MessageError
            $Output.IsSuccessful = $False
            $Output.Messages = $Msg
        }

        Write-Output $Output

    } #end foreach
}

End {
    
    Write-Progress -Activity * -Completed
    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow -Title

}

} #end Set-PKExplorerPreview