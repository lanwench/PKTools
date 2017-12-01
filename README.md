# Module PKTools

## About
|||
|---|---|
|**Name** |PKTools|
|**Author** |Paula Kingsley|
|**Type** |Script|
|**Version** |1.7.0|
|**Description**|Various PowerShell tools, functions, demos|
|**Date**|README.md file generated on Friday, December 1, 2017 12:33:51 PM|

This module contains 18 PowerShell functions or commands

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
|**ConvertTo-PKRegexArray**|Converts a simple array of strings to a regular expression with escaped characters|
|**Get-PKADComputerMiniReport**|Uses the ADSI type accelerator to return AD computer object AD details (no ActiveDirectory module required)|
|**Get-PKADDomainController**|Returns a domain controller for the current computer site or a named site in the current or named domain|
|**Get-PKADOrganizationalUnit**|Uses the ADSI type accelerator to return a menu of Organizational Units in an external gridview (no ActiveDirectory module required)|
|**Get-PKFileEncoding**|Returns the encoding type for one or more files|
|**Get-PKNetworkPortAssignments**|Downloads and returns a list of known port assignments from IANA|
|**Get-PKWindowsPath**|Gets the environment path value for a computer or user or both|
|**Install-PKChocoPackage**|Uses Invoke-Command to install Chocolatey packages on remote computers as PSJobs|
|**New-PKISEInvokeCommandSnippet**|Adds a new PS ISE snippet containing a template function (function runs Invoke-Command on one or more remote computers interactively or as a job)|
|**Open-PKWindowsMSC**|Launches an MMC snapin using current or alternate credentials|
|**Reset-PKPSModule**|Removes and re-imports named PowerShell modules|
|**Restore-PKISESession**|Restores tabs/files from file saved using Save-PKISESession|
|**Save-PKISESession**|Saves open tabs in current ISE session to a file|
|**Set-PKExplorerPreview**|Enables the local computer's Windows Explorer preview pane for an additional file type|
|**Show-PKNestedProgressBars**|Demonstrates four nested progress bars displaying the hour, minute, second, and millisecond in a countdown to midnight|
|**Show-PKObjectTaggingDemo**|Demonstrates the differences between using Select-Object and Add-Member to add properties to an output object|
|**Test-PKConnection**|Performs various connectivity tests to remote computers|
|**Test-PKWindowsPendingReboot**|Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer|
