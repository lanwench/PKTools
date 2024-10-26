# Module PKTools

## About
|||
|---|---|
|**Name** |PKTools|
|**Author** |Paula Kingsley|
|**Type** |Script|
|**Version** |2.19.0|
|**Description**|Various PowerShell tools, functions, demos, stuff, things|
|**Date**|README.md file generated on Friday, October 25, 2024 7:34:46 PM|

This module contains 30 PowerShell functions or commands

All functions should have reasonably detailed comment-based help, accessible via Get-Help ... e.g., 
  * `Get-Help Do-Something`
  * `Get-Help Do-Something -Examples`
  * `Get-Help Do-Something -ShowWindow`

## Prerequisites

Computers must:

  * be running PowerShell 4.0.0 or later

## Installation

Clone/copy entire module directory into a valid PSModules folder on your computer and run `Import-Module PKTools`

## Notes

_All code should be presumed to be written by Paula Kingsley unless otherwise specified (see the context help within each function for more information, including credits)._

_Changelogs are generally found within individual functions, not per module._

## Commands

|**Command**|**Version**|**Synopsis**|
|---|---|---|
|**Backup-PKChromeProfile**|01.01.0000|Backs up Chrome profiles to file|
|**Convert-PKBytesToSize**|01.00.0000|Converts any integer size given to a user friendly size|
|**Convert-PKEXchangeSMTPLog**|01.00.0000|Parses an Exchange send or receive connector log from a file (string or object) and returns a PSObject|
|**Convert-PKIISLog**|01.00.0000|Parses an IIS log from a file (string or object) and returns a PSObject|
|**ConvertTo-PKCSV**|01.00.0000|Performs ConvertTo-CSV on an input object, with customizeable delimiter and options to remove header row/quotes|
|**ConvertTo-PKRegex**|02.00.0000|Escapes characters in one or more strings for nefarious regex purposes|
|**Format-PKBytes**|01.00.0000|Converts bytes to human-readable form--detecting B,KB,MB,GB,TB,PB--and returning a PSObject or string|
|**Format-PKLDAPFilter**|01.00.0000|Formats an LDAP filter string with line breaks and indentation, outputting as a more visually readable string|
|**Get-PKADDisabledObject**|01.00.0000|Retrieves details of disabled Active Directory objects|
|**Get-PKADUserDisabledDate**|01.00.0000|Uses Get-ADUser and Get-ADReplicationAttributeMetadata to return the date user objects were disabled|
|**Get-PKDateTimeExamples**|01.00.0000|Returns standard or unix format date/time formatting options with examples and descriptions|
|**Get-PKGPResult**|01.00.0000|Executes gpresult to an HTML or XML file, with the option to modify html font name/size and launch the file with the associated handler|
|**Get-PKInputObjectType**|01.00.0000|Uses regex to check the type of the input object, in friendly and full name/type formats.|
|**Get-PKJOTD**|01.00.0000|Retrieves jokes from the v2.jokeapi.dev API based on specified parameters|
|**Get-PKSID**|01.00.0000|Gets the SID for one or more local or domain users or groups via .NET|
|**Get-PKTaniumClient**|01.02.0000|Gets the Tanium Client service and registry configuration from one or more computers, using Get-WMIObject for downlevel compatibility|
|**Install-PKVSCodePortable**|01.02.0000|Downloads and installs or updates VSCode Portable in a specified target directory, since Portable can't update itself!|
|**New-PKCodeSigningCert**|01.01.0000|Creates a new self-signed certificate on the local computer in the current user's certificate store|
|**New-PKComplexPassword**|01.00.0000|Uses Get-Random and defined character sets to generate a password between 10 and 265 characters, with option to return secure string or plain text|
|**New-PKFakeIdentity**|01.00.0000|Generates one or more random identities using Invoke-WebRequest and API call to publicapis.io, with option to return only basic details|
|**New-PKJargonIpsum**|01.00.0000|Want a wall of mission statements? This function generates jargon-filled Lorem Ipsum text from internal dictionary arrays of words, <br/>so you can fit in at your next meeting!|
|**New-PKPassphrase**|02.00.0000|Uses REST API calls to generates one or more passphrases of English or Lorem Ipsum, <br/>with the option to pick the count, the word count, the separator character, and number of integers|
|**Open-PKChrome**|01.00.0000|Launches a URL in Chrome, with options for default profile/new window|
|**Remove-PKAttributeBit**|01.00.0000|Removes one or more filesystem attribute bits from one or more files or folders (recursive)|
|**Remove-PKMcAfee**|01.00.0000|Removes McAfee Enterprise endpoint client from local computer without a key|
|**Resolve-PKDNSName**|01.01.0000|Performs forward and reverse lookups of one or more names or IP addresses, optionally testing for forward/reverse name match|
|**Restore-PKISESession**|03.00.0000|Restores tabs/files from text file created using Save-PKISESession|
|**Save-PKISESession**|03.00.0000|Saves open tabs in current ISE session to a file|
|**Test-PKLdapSSLConnection**|01.01.0000|Tests an LDAPS connection, returning information about the negotiated SSL connection including the server certificate.|
|**Update-PKADDisabledObjDescription**|01.00.0000|Updates the description field of disabled Active Directory objects with the object's disabled and last logon dates.|
