# Module PKTools

## About
|||
|---|---|
|**Name** |PKTools|
|**Author** |Paula Kingsley|
|**Type** |Script|
|**Version** |2.35|
|**Description**|Various PowerShell tools, functions, demos, stuff, things|
|**Date**|README.md file generated on Wednesday, July 8, 2026 1:53:33 PM|

This module contains 38 PowerShell functions or commands

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
|<span style="white-space: nowrap">**Backup-PKChromeProfile**</span>|01.02|Backs up Google Chrome user profile data to a date-stamped folder|
|<span style="white-space: nowrap">**Backup-PKVSCodeData**</span>|01.01|Backs up the VSCode \Data folder to a date-stamped zip file|
|<span style="white-space: nowrap">**Convert-PKBytes**</span>|02.00|Converts one or more byte counts to a human-readable unit, auto-detecting or using a specified target unit|
|<span style="white-space: nowrap">**Convert-PKCollectionToString**</span>|01.02|Converts object properties containing collections (arrays) into flattened strings<br/>for safe and clean export to CSV, avoiding the [system]System.Object[] value|
|<span style="white-space: nowrap">**Convert-PKEXchangeSMTPLog**</span>|01.01|Parses an Exchange send or receive connector log from a file (string or object) and returns a PSObject|
|<span style="white-space: nowrap">**Convert-PKIISLog**</span>|01.01|Parses an IIS log from a file (string or object) and returns a PSObject|
|<span style="white-space: nowrap">**ConvertTo-PKCSV**</span>|01.01|Performs ConvertTo-CSV on an input object, with customizeable delimiter and options to remove header row/quotes|
|<span style="white-space: nowrap">**ConvertTo-PKRegex**</span>|02.01|Escapes special characters in one or more strings for use in regex patterns|
|<span style="white-space: nowrap">**Get-PKCertificate**</span>|01.01|Retrieves SSL/TLS certificate details from one or more remote hosts by performing a TCP connection and SSL handshake.|
|<span style="white-space: nowrap">**Get-PKColorInfo**</span>|01.01|Look up color information by Hex code or Name with ANSI color examples, via REST API (default) or local query.|
|<span style="white-space: nowrap">**Get-PKDadJoke**</span>|02.01|Fetches random dad jokes from icanhazdadjoke.com|
|<span style="white-space: nowrap">**Get-PKDadJoke2**</span>|02.00|Retrieves dad jokes from the icanhazdadjoke API.|
|<span style="white-space: nowrap">**Get-PKDateTimeExamples**</span>|01.02|Returns standard or unix format date/time formatting options with examples and descriptions|
|<span style="white-space: nowrap">**Get-PKDomainInfo**</span>|01.00|Gets details for one or more domains using the RDAP server for that TLD, retrieved via IANA RDAP bootstrap registry|
|<span style="white-space: nowrap">**Get-PKFileReport**</span>|01.01|Generates an HTML report of files in a specified directory, including summary statistics and detailed file information.|
|<span style="white-space: nowrap">**Get-PKGoogleFSLogErrors**</span>|01.01|Scans local computer Google Drive FileSync log files for errors and returns matching entries.|
|<span style="white-space: nowrap">**Get-PKInputObjectType**</span>|01.01|Uses regex to check the type of the input object, in friendly and full name/type formats.|
|<span style="white-space: nowrap">**Get-PKJOTD**</span>|02.01|Retrieves jokes from the v2.jokeapi.dev API based on specified parameters|
|<span style="white-space: nowrap">**Get-PKPSVersions**</span>|01.01|Retrieves the installed versions and paths of Windows PowerShell and PowerShell Core on the local computer.|
|<span style="white-space: nowrap">**Get-PKPublicIP**</span>|02.01|Retrieves the public IP address of the machine via an API call to the ifconfig.me service|
|<span style="white-space: nowrap">**Get-PKSpecialChar**</span>|01.01|Retrieves Unicode characters and code points by named range, decimal code, or character|
|<span style="white-space: nowrap">**Get-PKTaniumClient**</span>|01.03|Gets the Tanium Client service and registry configuration from one or more computers, using Get-WMIObject for downlevel compatibility|
|<span style="white-space: nowrap">**Get-PKTimeZones**</span>|02.01|Retrieves system time zones with UTC offset, DST support, and current local time|
|<span style="white-space: nowrap">**Get-PKTree**</span>|01.01|Lists files and folders recursively from the current or specified directory, with filtering and output options|
|<span style="white-space: nowrap">**Get-PKWeather**</span>|01.01|Retrieves current weather conditions for a specified location using the OpenWeatherMap API.|
|<span style="white-space: nowrap">**Install-PKVSCodePortable**</span>|01.04|Downloads and installs or updates VSCode Portable from code.visualstudio.com|
|<span style="white-space: nowrap">**New-PKCodeSigningCert**</span>|01.02|Creates a new self-signed certificate on the local computer in the current user's certificate store|
|<span style="white-space: nowrap">**New-PKComplexPassword**</span>|01.01|Uses Get-Random and defined character sets to generate a password between 10 and 265 characters, with option to return secure string or plain text|
|<span style="white-space: nowrap">**New-PKFakeIdentity**</span>|01.01|Generates one or more random fake identities via the randomuser.me API|
|<span style="white-space: nowrap">**New-PKJargonIpsum**</span>|01.01|Generates jargon-filled Lorem Ipsum-style corporate text using built-in word arrays|
|<span style="white-space: nowrap">**New-PKPassphrase**</span>|03.01|Uses REST API calls to generate one or more passphrases from English or Lorem Ipsum word lists|
|<span style="white-space: nowrap">**Open-PKChrome**</span>|01.01|Launches a URL in Google Chrome with options for new window and default profile|
|<span style="white-space: nowrap">**Remove-PKAttributeBit**</span>|01.01|Removes one or more filesystem attribute bits from one or more files or folders (recursive)|
|<span style="white-space: nowrap">**Resolve-PKDNSName**</span>|02.00|Performs forward and reverse DNS lookups with optional matching check, on Windows (DNSClient module) or on Mac/Linux (DNSClient-PS module)|
|<span style="white-space: nowrap">**Restore-PKISESession**</span>|03.01|Restores open ISE tabs from a session file created by Save-PKISESession|
|<span style="white-space: nowrap">**Save-PKISESession**</span>|03.01|Saves open PowerShell ISE tabs to a file for later restoration via Restore-PKISESession|
|<span style="white-space: nowrap">**Test-PKLdapSSLConnection**</span>|01.02|Tests an LDAPS connection and returns SSL parameters and the server certificate|
|<span style="white-space: nowrap">**Test-PKVSCodePortableVersion**</span>|01.01|Gets the latest VSCode portable version available at code.visualstudio.com and compares it to the current local version|
