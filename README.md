# Module PKTools

## About
|||
|---|---|
|**Name** |PKTools|
|**Author** |Paula Kingsley|
|**Type** |Script|
|**Version** |1.19.0|
|**Description**|Various PowerShell tools, functions, demos, stuff, things|
|**Date**|README.md file generated on Friday, June 21, 2019 01:54:07 PM|

This module contains 65 PowerShell functions or commands

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
|**Backup-PKChromeProfile**|Backs up Chrome profiles to file|
|**ConvertFrom-PKErrorRecord**|Converts an error record or stop exception to a more intelligible format|
|**Convert-PKBytesToSize**|Converts any integer size given to a user friendly size|
|**Convert-PKSubnetCIDR**|Converts between dotted-decimal subnet mask and CIDR length, or displays full table with usable networks/hosts|
|**Convert-PKSubnetMask**|Converts a dotted-decimal subnet mask to CIDR index, or vice versa|
|**ConvertTo-PKCSV**|Performs ConvertTo-CSV on an input object, with customizeable delimiter and options to remove header row/quotes|
|**ConvertTo-PKRegexArray**|Converts a simple array of strings to a regular expression with escaped characters|
|**Enable-PKExplorerPreview**|Modifies the registry to enable file content viewing in the Windows Explorer preview pane|
|**Format-PKBytes**|Converts bytes to human-readable form--detecting B,KB,MB,GB,TB,PB--and returning a PSObject or string|
|**Format-PKTestWSMANError**|Formats error messages from Test-WSMAN into human-readable strings|
|**Get-PKADComputerMiniReport**|Uses the ADSI type accelerator to return AD computer object AD details (no ActiveDirectory module required)|
|**Get-PKADDomainController**|Returns a domain controller for the current computer site or a named site in the current or named domain|
|**Get-PKADOrganizationalUnit**|Uses the ADSI type accelerator to return a menu of Organizational Units in an external gridview (no ActiveDirectory module required)|
|**Get-PKADOrganizationalUnitLinkedGPOs**|Gets details or counts of Group Policy object linked to Active Directory organizational units|
|**Get-PKChocoPackages**|Gets a list of locally installed Chocolatey packages, interactively or as a PSJob|
|**Get-PKCompletedJobOutput**|Gets the results for completed PowerShell jobs by name or ID, with option to keep or remove results, or remove job entirely|
|**Get-PKDateTimeFormats**|Lists Powershell date/time format examples (standard/custom/all, or a legend)|
|**Get-PKErrorMessage**|Returns details about errors from ErrorRecord objects|
|**Get-PKFile**|Uses a queue object and .NET [IO.Directory] to search a path for files matching a string|
|**Get-PKFileEncoding**|Returns the encoding type for one or more files|
|**Get-PKLocalGroupMember**|Uses 'net localgroup' to return members of a named local group or all groups|
|**Get-PKNetworkPortAssignments**|Downloads and returns a list of known port assignments from IANA|
|**Get-PKPSObjectProperties**|Returns property names from a PSCustomObject in order as an array, suitable for Select-Object|
|**Get-PKTaniumActionLog**|Invokes a scriptblock to parse content from the Tanium client action log, returning a PSObject|
|**Get-PKTimeZones**|Returns all time zone info|
|**Get-PKWindowsDateTime**|Returns various date / time / time zone settings for a computer|
|**Get-PKWindowsHardware**|Does something cool, interactively or as a PSJob|
|**Get-PKWindowsHotfix**|Invokes a scriptblock to return installed Windows hotfixes using WMI (all or by KB number), interactively or as a PSJob|
|**Get-PKWindowsLicenseInfo**|Uses good old slmgr.vbs & WMI to retrieve Windows licensing on a local or remote computer, interactively or as a PSJob|
|**Get-PKWindowsPath**|Gets the environment path value for a computer or user or both|
|**Get-PKWindowsProductKey**|Uses WMI to retrieve a Windows product key on a local or remote computer, interactively or as a PSJob|
|**Get-PKWindowsRegistryDN**|Gets the DistinguishedName value of an AD-joined Windows computer from its registry, interactively or as a PSJob|
|**Get-PKWindowsReport**|Returns Windows computer report data, interactively or as a PSJob|
|**Get-PKWindowsTPMChipInfo**|Gets TPM chip data for a local or remote computer, interactively or as a PSJob|
|**Get-PKWinRMTrustedHosts**|Uses Get-Item to return the trusted hosts for WinRM configured on the local computer|
|**Install-PKChocoPackage**|Uses Invoke-Command to install Chocolatey packages on one or more computers, interactively or as PSJobs|
|**Install-PKWindowsPythonPackage**|Installs a Python pip package via Invoke-Command, synchronously or in a PSJob|
|**New-PKCodeSigningCert**|Creates a new self-signed certificate on the local computer in the current user's certificate store|
|**New-PKISESnippetFunctionAD**|Adds a new PS ISE snippet containing a template function using the ActiveDirectory module|
|**New-PKISESnippetFunctionGeneric**|Adds a new PS ISE snippet containing a template function|
|**New-PKISESnippetFunctionInvokeCommand**|Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more computers interactively or as a job)|
|**New-PKISESnippetFunctionVMware**|Adds a new PS ISE snippet containing a template function using the PowerCLI module|
|**New-PKPSModuleManifest**|Creates a new PowerShell module manifest using New-ModuleManifest|
|**New-PKRandomPassword**|Generates a random password string, with option to select length|
|**Remove-PKMcAfee**|Removes McAfee Enterprise endpoint client from local computer without a key|
|**Reset-PKPSModule**|Removes and re-imports named PowerShell modules|
|**Resolve-PKDNSName**|Uses Resolve-DNSName to test lookups on a DNS server, with options for record type and truncated output|
|**Resolve-PKDNSNameByNET**|Uses [System.Net.Dns]::GetHostEntryAsync to resolves a hostname to an IP address, or an IP address to a hostname|
|**Restore-PKISESession**|Restores tabs/files from text file created using Save-PKISESession|
|**Save-PKISESession**|Saves open tabs in current ISE session to a file|
|**Select-PKActiveDirectoryOU**|Uses Windows Forms to generate a selectable tree view of Active Directory containers/organizational units|
|**Set-PKWindowsExplorerPreview**|Enables the local computer's Windows Explorer preview pane for an additional file type|
|**Set-PKWinRMTrustedHosts**|Uses Set-Item to modify the trusted hosts for WinRM configured on the local computer|
|**Show-PKErrorMessageBox**|Uses Windows Forms to display an error message in a message box (defaults to most recent error)|
|**Show-PKNestedProgressBars**|Demonstrates four nested progress bars displaying the hour, minute, second, and millisecond in a countdown to midnight|
|**Show-PKObjectTaggingDemo**|Demonstrates the differences between using Select-Object and Add-Member to add properties to an output object|
|**Show-PKPSConsoleColors**|Displays console colors, with option to show as a grid with different foreground/background combinations|
|**Show-PKSubnetMaskTable**|Displays a table of prefix lengths and dotted decimal subnet masks|
|**Test-Department**|Demonstration of dynamic dynamic parameters|
|**Test-PKConnection**|Uses Get-WMIObject and the Win32_PingStatus class to quickly ping one or more targets, either interactively or as a PSJob|
|**Test-PKDNSServer**|Uses Resolve-DNSName to perform DNS lookups on one or more servers, with connectivity tests and option to return lookup output|
|**Test-PKDynamicParameter**|Demonstration of dynamic dynamic parameters|
|**Test-PKNetworkConnections**|Performs various connectivity tests to remote computers|
|**Test-PKWindowsPendingReboot**|Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer|
|**Test-PKWSMan**|Test WinRM connectivity to a remote computer using various protocols|
