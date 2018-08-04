#requires -Version 3

Function Enable-PKExplorerPreview {
<#
.Synopsis
    Modifies the registry to enable file content viewing in the Windows Explorer preview pane
   
.DESCRIPTION
    Modifies the registry to enable file content viewing in the Windows Explorer preview pane
    Requires an elevated shell and administrative privileges
    Outputs a psobject
    Accepts pipeline input

.NOTES
    Name    : Function_Enable-PKExplorerPreview.ps1 
    Created : 2018-07-17
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.00.0000 - 2018-07-17 - Created script

.EXAMPLE
    PS C:\> Enable-PKExplorerFilePreview -FileExtension psd1 -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key           Value                       
        ---           -----                       
        FileExtension {psd1}                      
        Verbose       True                        
        ScriptName    Enable-PKExplorerFilePreview
        ScriptVersion 1.0.0                       

        VERBOSE: Modify registry to allow Windows Explorer preview pane for file extensions
        WARNING: This script requires administrative privileges and an elevated shell. It does not verify that the file type can be viewed!
        VERBOSE: Enable Windows Explorer preview pane content view for '.psd1'

        PerceivedType : text
        PSPath        : Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\.psd1
        PSParentPath  : Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT
        PSChildName   : .psd1
        PSProvider    : Microsoft.PowerShell.Core\Registry

.EXAMPLE
    PS C:\> Enable-PKExplorerPreview -FileExtension .rb

        WARNING: This script requires administrative privileges and an elevated shell. It does not verify that the file type can be viewed!
        Preview pane already enabled for file extension '.rb'


#>
[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact="High")]
Param(
    [Parameter(
        Mandatory = $True,
        ValueFromPipeline = $True,
        HelpMessage = "One or more file extensions, with or without leading dots"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$FileExtension
)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Display our parameters
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #endregion Show parameters

    # Preference
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"


    # Function to check if PowerShell is running elevated
    function Check-Elevated{
      $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
      $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
      $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
      $IsAdmin=$prp.IsInRole($adm)
        if ($IsAdmin){
            Set-Variable -Name elevated -Value $true -Scope 1
        }
    }

    # Console output
    #$Host.UI.WriteLine() 
    $BGColor = $Host.UI.RawUI.BackgroundColor
    $Msg = "Modify registry to allow Windows Explorer preview pane for file extensions"
    #$FGColor = "Yellow"
    #If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    #Else {Write-Verbose $Msg}
    Write-Verbose $Msg

    $Msg = "This script requires administrative privileges and an elevated shell. It does not verify that the file type can be viewed!"
    Write-Warning $Msg

}
Process {

    Foreach ($Ext in $FileExtension) {
        
        If ($Ext -notmatch '^.[a-zA-Z0-9]') {$Ext = ".$Ext"}

        $Msg = "Enable Windows Explorer preview pane content view for '$Ext'"
        Write-Verbose $Msg 

        If ($Null = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$Ext" -Name PerceivedType -ErrorAction SilentlyContinue).PerceivedType -ne "Text") {

            If ($PSCmdlet.ShouldProcess($Env:ComputerName,"`n`n`t$Msg`n")) {
                Set-ItemProperty "Registry::HKEY_CLASSES_ROOT\$Ext" -Name PerceivedType -Value text -ErrorAction Stop -PassThru -Force
            }
            Else {
                $Msg = "Operation cancelled by user"
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Else {
            $Msg = "Preview pane already enabled for file extension '$Ext'"
            $Host.UI.WriteErrorLine($Msg)
        }
    }

}

} #end Enable-PKExplorerFilePreview



<#

# Enable-ps1Preview.ps1
# This script will enable preview for ps1 files in Windows Explorer.
# Special thanks to Nate Bruneau for the idea.
#
# Created by 
# Bhargav Shukla
# http://www.bhargavs.com
# 
# DISCLAIMER
# ==========
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
# RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#############################################################################

# Check if OS is Windows 7 or Windows Server 2008 R2, quit if not.
$OS = (Get-WmiObject -Class win32_OperatingSystem).caption
switch -wildcard ($OS){
  "*Windows 7*" {"`nChecking Elevation..."}
  "*Windows Server 2008 R2*" {"`nChecking Elevation..."}
  default {Write-Host -ForegroundColor Red "`nYou are not running Windows 7 or Windows Server 2008 R2. You can't use this feature on older OS."; exit}
}

# Function to check if PowerShell is running elevated
function Check-Elevated{
  $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
  $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
  $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
  $IsAdmin=$prp.IsInRole($adm)
  if ($IsAdmin){
    Set-Variable -Name elevated -Value $true -Scope 1
  }
}

# Make registry changes if running elevated, throw error if not
Check-Elevated

If ($elevated -eq $true){
	
# Set Registry Key variables
	$REG_KEY = ".ps1"
			
	# Open remote registry
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('ClassesRoot', ".")
	
	# Open the targeted remote registry key/subkey as read/write
	$regKey = $reg.OpenSubKey($REG_KEY,$true)
		
	# Set PerceivedType to "text"
	if ($regKey -ne $null){
		$regKey.Setvalue('PerceivedType', 'text', 'String')
		$regKey.Setvalue('Content Type', 'text/plain', 'String')
		
		# Close the Reg Key
		$regKey.Close()
	
		Write-Host -ForegroundColor Green -BackgroundColor Black "Preview for .ps1 files is now enabled. Enable preview pane in Windows Explorer.`n"
			
	}
}

else{
  Write-Host -ForegroundColor Red -BackgroundColor Black "Please run PowerShell as administrator before you run this script.`n"
}

#>