#requires -Version 3
Function Set-PKExplorerPreview {
<#
.SYNOPSIS
    Enables the local computer's Windows Explorer preview pane for an additional file type
    
.DESCRIPTION
    Enables the local computer's Windows Explorer preview pane for an additional file type
    by modifying the registry to set the percieved file type/content type as text
            
.NOTES
    Name    : Function_Set-PKExplorerPreview.ps1
    Version : 1.0.0
    Author  : Paula Kingsley
    Created : 2017-08-22
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v1.0.0 - 2017-08-22 - Created script
        
.LINK
    https://blogs.technet.microsoft.com/bshukla/2010/03/30/script-to-enable-preview-pane-for-powershell-scripts/

.EXAMPLE
    PS C:\> @("log","foo","txt") | Set-PKExplorerPreview -Verbose

        VERBOSE: PSBoundParameters: 
 	
        Key           Value                
        ---           -----                
        Verbose       True                 
        FileExtension                      
        ScriptName    Set-PKExplorerPreview
        ScriptVersion 1.0.0                


        VERBOSE: Checking OS version compatibility
        VERBOSE: Verified OS Microsoft Windows 7 Enterprise 
        VERBOSE: PowerShell is running in Elevated mode
        WARNING: This script tells Windows to treat the specified file extension as a text file,
        allowing it to be viewed in the Windows Explorer preview pane.
        It does not verify that the file type is compatible with this setting; please proceed with caution!

        VERBOSE: Open registry key/subkey for .log as read/write
        VERBOSE: Current settings for file extension 'log'
	        PerceivedType: (none)
	        Content Type : (none)
        VERBOSE: Enable Windows Explorer preview for file extension 'log'
        Registry subkeys successfully set for '.log'
	        PerceivedType: text
	        Content Type : text/plain
        VERBOSE: Open registry key/subkey for .foo as read/write
        No registry subkey found for file extension '.foo'
        VERBOSE: Open registry key/subkey for .txt as read/write
        VERBOSE: Current settings for file extension 'txt'
	        PerceivedType: text
	        Content Type : text/plain
        Operation cancelled by user for file extension '.txt'


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
    [string[]]$FileExtension

)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "1.0.0"

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    

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

    # General-purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose = $False
    }
    
    # Function to check if PowerShell is running in elevated mode
    function Test-Elevated{
        [CmdletBinding()]
        $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
        $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
        $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
        [switch]$IsAdmin=$prp.IsInRole($adm)
        if ($IsAdmin.IsPresent)  {
            $Msg = "PowerShell is running in Elevated mode"
            Write-Verbose $Msg
            $True
        }
        Else {
            $Msg = "PowerShell must be running in Elevated mode; please re-launch as Administrator"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            $False
        }
    }

    If (-not $SkipOSVersionCheck.IsPresent) {
        $Msg = "Checking OS version compatibility"    
        Write-Verbose $Msg
        
        Try {
            $OS = (Get-WmiObject -Class win32_OperatingSystem @StdParams).caption
            switch -wildcard ($OS){
                "*Windows 7*" {
                    $Msg = "Verified OS $OS"
                    Write-Verbose $Msg
                    $Null = Test-Elevated
                }
                "*Windows Server 2008 R2*" {
                    $Msg = "Verified OS $OS"
                    Write-Verbose $Msg
                    $Null = Test-Elevated
                }
                default {
                    $Msg = "You are running $OS. This function requires Windows 7 or Windows Server 2008 R2. "
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
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

    $Msg = "This script tells Windows to treat the specified file extension as a text file,`nallowing it to be viewed in the Windows Explorer preview pane.`nIt does not verify that the file type is compatible with this setting; please proceed with caution!"
    Write-Warning $Msg
    $Host.UI.WriteLine()

}

Process {
    
    Foreach ($REG_KEY in $FileExtension) {

        Try {
            
            # Normalize input
            $REG_KEY = $REG_KEY.Replace(".","")
            $REG_KEY = ".$REG_KEY"

            # Set flag
            [switch]$Change = $False

	        # Open registry
	        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('ClassesRoot', ".")
	
            # Open the targeted remote registry key/subkey as read/write
            $Msg = "Open registry key/subkey for $REG_KEY as read/write"
            Write-Verbose $Msg	        
            If (-not ($regKey = $reg.OpenSubKey($REG_KEY,$true))) {
                $Msg = "No registry subkey found for file extension '$REG_KEY'"
                $Host.UI.WriteErrorLine($MSG)
            }
            Else {
                # Test current settings
                If (-not ($Perceived = $Regkey.GetValue("PerceivedType"))) {$Perceived = "(none)"}
                If (-not ($Content = $Regkey.GetValue("Content Type"))) {$Content = "(none)"}

                $Msg = "Current settings for file extension '$FileExtension'`n`tPerceivedType: $Perceived`n`tContent Type : $Content"
                Write-Verbose $Msg

                If ("(none)" -in @($Perceived,$Content)) {
                    $Change = $True
                }
                Else {
                    $Msg = "`n`nModify registry subkey for file extension '$REG_KEY' `n`n`tChange PerceivedType from '$Perceived' to 'text'`n`tChange Content Type from '$Content' to 'text/plain'`n`n`tThis is not recommended!`n`n"
                    If ($PSCmdlet.ShouldProcess($Env:ComputerName,$Msg)) {
                        $Change = $True
                    }
                    Else {
                        $Msg = "Operation cancelled by user for file extension '$REG_KEY'"
                        $Host.UI.WriteLine($Msg)
                        $Change = $False
                        Break
                    }
                }
            }
        }
        Catch {
            Break
        }

        If ($Change.IsPresent) {
            $Msg = "Enable Windows Explorer preview for file extension '$FileExtension'"
            Write-Verbose $Msg
            If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$Msg)) {
                Try {
                    $regKey.Setvalue('PerceivedType', 'text', 'String')
		            $regKey.Setvalue('Content Type', 'text/plain', 'String')
		        
                    # Get new settings
                    $Perceived = $Regkey.GetValue("PerceivedType")
                    $Content = $Regkey.GetValue("Content Type")

		            # Close the Reg Key
		            $regKey.Close()

                    $Msg = "`Registry subkeys successfully set for '$REG_KEY'`n`tPerceivedType: $Perceived`n`tContent Type : $Content"
                    $Host.UI.WriteLine($Msg)
                }
                Catch {
                    $Msg = "Operation failed for file extension '$REG_KEY'"
                    $ErrorDetails = $_.Exception.Message
                    $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
                    Break
                }
            }
            Else {
                $Msg = "Operation cancelled by user for file extension '$REG_KEY'"
                $Host.UI.WriteLine($Msg)
            }
    
        }
    } #end foreach
}

End {}

} #end Set-PKExplorerPreview