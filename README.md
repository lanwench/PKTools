# Module PKTools

## About
|||
|---|---|
| **Name** | PKTools |
| **Author** | Paula Kingsley |
| **Type** | Script |
| **Version** | 1.5.0 |
| **Date** | README.md file generated on Thursday, November 30, 2017 15:42:38 |

This module contains 17 PowerShell functions or commands

All functions should have reasonably detailed comment-based help, accessible via Get-Help ... e.g., 
  * `Get-Help Do-Something`
  * `Get-Help Do-Something -Examples`
  * `Get-Help Do-Something -ShowWindow`

## Prerequisites

Computers must:
  * be running PowerShell 3.0.0 or later

## Installation

Clone/copy entire module directory into a valid PSModules folder on your computer and run `Import-Module PKTools`

## Notes

_All code should be presumed to be written by Paula Kingsley unless otherwise specified (see the context help within each function for more information, including credits)._

## Commands
|**Command**|**Synopsis**|
|---|---|
| **ConvertTo-PKRegexArray** | Converts a simple array of strings to a regular expression with escaped characters |
| **Get-PKADComputerMiniReport** | <br/>Get-PKADComputerMiniReport [[-ComputerName] <string[]>] [[-SizeLimit] <int>] [[-DomainDN] <string>] [[-Credential] <Object>] [-GetADSite] [-TestConnection] [-SuppressConsoleOutput] [<CommonParameters>]<br/> |
| **Get-PKADDomainController** | Returns a domain controller for the current computer site or a named site in the current or named domain |
| **Get-PKADOrganizationalUnit** | <br/>Get-PKADOrganizationalUnit [[-SearchBase] <string>] [[-MatchPattern] <string>] [[-ExcludePattern] <string>] [-SuppressConsoleOutput] [<CommonParameters>]<br/> |
| **Get-PKFileEncoding** | Returns the encoding type for one or more files |
| **Get-PKNetworkPortAssignments** | Downloads and returns a list of known port assignments from IANA |
| **Get-PKWindowsPath** | Gets the environment path value for a computer or user or both |
| **Install-PKChocoPackage** | Uses Invoke-Command to install Chocolatey packages on remote computers as PSJobs |
| **New-PKISEInvokeCommandSnippet** | <br/>New-PKISEInvokeCommandSnippet [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]<br/> |
| **Open-PKWindowsMSC** | Launches an MMC snapin using current or alternate credentials |
| **Reset-PKPSModule** | Removes and re-imports named PowerShell modules |
| **Restore-PKISESession** | Restores tabs/files from file saved using Save-PKISESession |
| **Save-PKISESession** | Saves open tabs in current ISE session to a file |
| **Set-PKExplorerPreview** | Enables the local computer's Windows Explorer preview pane for an additional file type |
| **Show-PKNestedProgressBars** | <br/>Show-PKNestedProgressBars [<CommonParameters>]<br/> |
| **Test-PKConnection** | <br/>Test-PKConnection [-ComputerName] <string[]> [-Credential <pscredential>] [-Tests <string>] [-SuppressConsoleOutput] [<CommonParameters>]<br/> |
| **Test-PKWindowsPendingReboot** | Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer |
