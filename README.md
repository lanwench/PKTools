# Module PKTools

## About
|||
|---|---|
|**Name** |PKTools|
|**Author** |Paula Kingsley|
|**Type** |Script|
|**Version** |1.39.0|
|**Date**|README.md file generated on Wednesday, May 4, 2022 4:39:09 PM|

This module contains 102 PowerShell functions or commands

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

_Changelogs are generally found within individual functions, not per module._

## Commands

|**Command**|**Version**|**Synopsis**|
|---|---|---|
|**Backup-PKChromeProfile**|01.01.0000|Backs up Chrome profiles to file|
|**Compare-PKHostnameToPTR**|01.00.0000|Compares a hostname to the PTR of its IP, using the .NET system.net.dns class|
|**ConvertFrom-PKErrorRecord**|01.00.0000|Converts an error record or stop exception to a more intelligible format|
|**Convert-PKBytesToSize**|01.00.0000|Converts any integer size given to a user friendly size|
|**Convert-PKDistinguishedNameToJSON**|01.00.0000|Uses Zachary Loeber's Get-ChildOUStructure to output the CanonicalName format of a container/OU as JSON|
|**Convert-PKDistinguishedNameToOrganizationalUnit**|01.00.0000|Converts an Active Directory object's DistinguishedName to its parent Organizational Unit or Container, displayed in DN or CanonicalName format|
|**Convert-PKDNSRecordData**|-|Converts the RecordData output from Get-DNSZoneResourceRecord to a string based on record type|
|**Convert-PKEXchangeSMTPLog**|01.00.0000|Parses an Exchange send or receive connector log from a file (string or object) and returns a PSObject|
|**Convert-PKIISLog**|01.00.0000|Parses an IIS log from a file (string or object) and returns a PSObject|
|**Convert-PKLastLogon**|01.00.0000|Converts an Active Directory object's LastLogonTimestamp attribute value to human-readable datetime format|
|**Convert-PKSubnetCIDR**|01.00.0000|Converts between dotted-decimal subnet mask and CIDR length, or displays full table with usable networks/hosts|
|**Convert-PKSubnetMask**|01.00.0000|Converts a dotted-decimal subnet mask to CIDR index, or vice versa|
|**ConvertTo-PKCSV**|01.00.0000|Performs ConvertTo-CSV on an input object, with customizeable delimiter and options to remove header row/quotes|
|**ConvertTo-PKMarkdownTable**|01.00.0000|Converts a PSObject to a markdown table|
|**ConvertTo-PKRegexArray**|01.00.0000|Converts a simple array of strings to a regular expression with escaped characters|
|**Enable-PKExplorerPreview**|01.00.0000|Modifies the registry to enable file content viewing in the Windows Explorer preview pane|
|**Format-PKBytes**|01.00.0000|Converts bytes to human-readable form--detecting B,KB,MB,GB,TB,PB--and returning a PSObject or string|
|**Format-PKTestWSMANError**|01.00.0000|Formats error messages from Test-WSMAN into human-readable strings|
|**Get-PKADComputerMiniReport**|01.01.0000|Uses the ADSI type accelerator to return AD computer object AD details (no ActiveDirectory module required)|
|**Get-PKADDomainController**|01.01.0000|Returns a domain controller for the current computer site or a named site in the current or named domain|
|**Get-PKADOrganizationalUnit**|01.01.0000|Uses the ADSI type accelerator to return a menu of Organizational Units in an external gridview (no ActiveDirectory module required)|
|**Get-PKADOrganizationalUnitLinkedGPOs**|01.00.0000|Gets details or counts of Group Policy object linked to Active Directory organizational units|
|**Get-PKADReplicationFailures**|01.00.0000|Gets AD replication errors for domain controllers in a forest|
|**Get-PKBatteryReport**|01.00.0000|Runs powercfg to create an HTML report on the utilization and history of the local computer battery|
|**Get-PKChocoPackages**|01.00.0000|Gets a list of locally installed Chocolatey packages, interactively or as a PSJob|
|**Get-PKCompletedJobOutput**|01.00.0000|Gets the results for completed PowerShell jobs by name or ID, with option to keep or remove results, or remove job entirely|
|**Get-PKCurrentUser**|01.00.0000|Gets details for the currently logged-in user, including group membership, using .NET DirectoryServices.AccountManagement|
|**Get-PKDateTimeFormats**|01.00.0000|Lists Powershell date/time format examples (standard/custom/all, or a legend)|
|**Get-PKErrorMessage**|01.01.0000|Returns details about errors from ErrorRecord objects|
|**Get-PKFile**|01.00.0000|Uses a queue object and .NET [IO.Directory] to search a path for files matching a string|
|**Get-PKFileEncoding**|1.0.0|Returns the encoding type for one or more files|
|**Get-PKFunctionPath**|01.00.0000|Returns the underlying filename(s) for a PowerShell alias or function, either by command name or module name|
|**Get-PKIANAPorts**|01.00.0000|Uses Invoke-Webrequest to get a CSV file from iana.org and creates a PSObject of TCP/UDP names, port numbers, and descriptions|
|**Get-PKIPConfig**|01.00.0000|Gets IPv4 information for one or more Windows computers, interactively or as a PSJob|
|**Get-PKLocalGroupMember**|01.00.0000|Uses 'net localgroup' to return members of a named local group or all groups|
|**Get-PKNestedGroupMembers**|01.00.0000|Returns direct and nested members for an Active Directory group, including the parent group & depth/level number|
|**Get-PKNetworkPortAssignments**|01.00.0000|Downloads and returns a list of known port assignments from IANA|
|**Get-PKPSObjectProperties**|01.00.0000|Returns property names from a PSCustomObject in order as an array, suitable for Select-Object|
|**Get-PKSitesAndSubnets**|01.00.0000|Returns AD sites and subnets using .NET|
|**Get-PKTaniumActionLog**|02.02.0000|Invokes a scriptblock to parse content from the Tanium client action log, returning a PSObject|
|**Get-PKTaniumConfig**|01.00.0000|Invokes a scriptblock to get Tanium client config details, interactively or as a PSJob|
|**Get-PKTimeZones**|01.00.0000|Returns all time zone info|
|**Get-PKvCenterMetrics**|01.00.0000|Lists all available vCenter performance metrics, optionally sorted or grouped by property names|
|**Get-PKVMIPConfig**|-|Returns IPv4 configuration data for one or more guest VMs|
|**Get-PKWindowsActivationStatus**|01.00.0000|Gets the activation status for one or more remote Windows computers|
|**Get-PKWindowsDateTime**|01.00.0000|Returns various date / time / time zone settings for a computer|
|**Get-PKWindowsEvent**|01.00.0000|Uses Get-WinEvent to return events from Windows event logs on one or more computers|
|**Get-PKWindowsEventLogSettings**|01.00.0000|Returns information on Windows event logs on one or more computers|
|**Get-PKWindowsHardware**|01.00.0000|Does something cool, interactively or as a PSJob|
|**Get-PKWindowsHotfix**|01.00.0000|Invokes a scriptblock to return installed Windows hotfixes using WMI (all or by KB number), interactively or as a PSJob|
|**Get-PKWindowsLicenseInfo**|01.00.0000|Uses good old slmgr.vbs & WMI to retrieve Windows licensing on a local or remote computer, interactively or as a PSJob|
|**Get-PKWindowsPath**|01.00.0000|Gets the environment path value for a computer or user or both|
|**Get-PKWindowsProductKey**|01.00.0000|Uses WMI to retrieve a Windows product key on a local or remote computer, interactively or as a PSJob|
|**Get-PKWindowsRegistryDN**|01.01.0000|Gets the DistinguishedName value of an AD-joined Windows computer from its registry, interactively or as a PSJob|
|**Get-PKWindowsReport**|01.00.0000|Returns Windows computer report data, interactively or as a PSJob|
|**Get-PKWindowsRoute**|01.00.0000|Invokes a scriptblock to get network routes on one or more computers using Get-NetRoute (available in PowerShell 4 on Windows 8/2012 and newer)|
|**Get-PKWindowsShutdown**|01.00.0000|Invokes a scriptblock to query the Windows event log for shutdown/startup events via Get-WinEvent, or last boot time data via Get-WMIObject, interactively or as a PSJob|
|**Get-PKWindowsSoftware**|01.00.0000|Invokes a scriptblock to return the installed software on one or more Windows computers, using registry lookups instead of WMI|
|**Get-PKWindowsTPMChipInfo**|01.00.0000|Gets TPM chip data for a local or remote computer, interactively or as a PSJob|
|**Get-PKWinRMTrustedHosts**|01.00.0000|Uses Get-Item to return the trusted hosts for WinRM configured on the local computer|
|**Install-PKChocoPackage**|01.01.000|Uses Invoke-Command to install Chocolatey packages on one or more computers, interactively or as PSJobs|
|**Install-PKWindowsPythonPackage**|01.00.0000|Installs a Python pip package via Invoke-Command, synchronously or in a PSJob|
|**Install-PKWMIExporter**|01.00.0000|Invokes a scriptblock to download and install WMI Exporter, with specified collectors/listening port - defaults to AD collectors|
|**Invoke-PKTelnet**|01.00.0000|Uses System.Net.Sockets.TcpClient to test telnet connectivity on a specified port, to one or more computers, with a timeout|
|**Invoke-PKWin2019Activation**|01.00.0000|Changes the product key and activates Windows 2019 Standard or Datacenter|
|**New-PKCodeSigningCert**|01.00.0000|Creates a new self-signed certificate on the local computer in the current user's certificate store|
|**New-PKISESnippetFunction**|-|Adds a new PS ISE snippet containing a template function for ActiveDirectory, Invoke-Command, VMware, or generic use|
|**New-PKPSModuleManifest**|01.00.0000|Creates a new PowerShell module manifest using New-ModuleManifest|
|**New-PKRandomPassword**|01.00.0000|Generates a random password string, with option to select length|
|**Remove-GNOpsWindowsGraphiteFolder**|01.00.0000|Invokes a scriptblock to remove the GraphitePowershell folder|
|**Remove-PKAttributeBit**|01.00.0000|Removes one or more filesystem attribute bits from one or more files or folders (recursive)|
|**Remove-PKISECache**|01.00.0000|Removes PowerShell ISE cache files for the local user|
|**Remove-PKMcAfee**|01.00.0000|Removes McAfee Enterprise endpoint client from local computer without a key|
|**Reset-PKPSModule**|03.00.0000|Removes and re-imports named PowerShell modules|
|**Resolve-PKDNSName**|01.01.0000|Uses Resolve-DNSName to test lookups on a DNS server, with options for record type and truncated output|
|**Resolve-PKDNSNameByNET**|03.00.0000|Uses [System.Net.Dns]::GetHostEntryAsync to resolves a hostname to an IP address, or an IP address to a hostname|
|**Resolve-PKIPtoPTR**|01.00.0000|Performs a reverse DNSlookup on one or more IPv4 addresses, against one or more DNS server names or IP addresses|
|**Restore-PKISESession**|02.01.0000|Restores tabs/files from text file created using Save-PKISESession|
|**Save-PKISESession**|02.01.0000|Saves open tabs in current ISE session to a file|
|**Select-PKActiveDirectoryOU**|01.00.0000|Uses Windows Forms to generate a selectable tree view of Active Directory containers/organizational units|
|**Send-PKHTMLEmail**|01.00.0000|Send a nicely-formatted HTML email from an existing PSObject|
|**Set-PKWindowsExplorerPreview**|01.03.0000|Enables the local computer's Windows Explorer preview pane for an additional file type|
|**Set-PKWinRMTrustedHosts**|01.00.0000|Uses Set-Item to modify the trusted hosts for WinRM configured on the local computer|
|**Set-PKWMIPermissions**|01.00.0000|Invokes a scriptblock to set WMI permissions, interactively or as a PSJob, with a huge HT to Graeme Bray & Steve Lee|
|**Show-PKErrorMessageBox**|01.00.0000|Uses Windows Forms to display an error message in a message box (defaults to most recent error)|
|**Show-PKNestedProgressBars**|01.00.0000|Demonstrates four nested progress bars displaying the hour, minute, second, and millisecond in a countdown to midnight|
|**Show-PKObjectTaggingDemo**|01.00.0000|Demonstrates the differences between using Select-Object and Add-Member to add properties to an output object|
|**Show-PKPSConsoleColors**|01.00.0000|Displays console colors, with option to show as a grid with different foreground/background combinations|
|**Show-PKSubnetMaskTable**|01.00.0000|Displays a table of prefix lengths and dotted decimal subnet masks|
|**Test-Department**|-|Demonstration of dynamic dynamic parameters|
|**Test-PKAccountLockout**|01.00.0000|A simple little function that uses .NET / ADSI searcher to return the account lockout status of one or more AD users in the current user's domain|
|**Test-PKADConnection**|01.00.0000|Tests connectivity to an Active Directory domain (and optional domain controller) without the ActiveDirectory module|
|**Test-PKConnection**|01.00.0000|Uses Get-WMIObject and the Win32_PingStatus class to quickly ping one or more targets, either interactively or as a PSJob|
|**Test-PKDNSServer**|02.00.0000|Uses Resolve-DNSName to perform record lookups on one or more DNS servers, with optional connectivity tests and option to return lookup output|
|**Test-PKDomainCredential**|01.00.0000|Tests authentication to Active Directory|
|**Test-PKDynamicParameter**|-|Demonstration of dynamic dynamic parameters|
|**Test-PKLdapSslConnection**|01.00.0000|Tests an LDAPS connection, returning information about the negotiated SSL connection including the server certificate.|
|**Test-PKNetworkConnections**|02.00.0000|Performs various connectivity tests to remote computers|
|**Test-PKPasswordPolicy**|01.00.0000|Tests a string against a domain password policy for length and complexity|
|**Test-PKWindowsPendingReboot**|1.00.0000|Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer|
|**Test-PKWinRM**|02.00.0000|Test WinRM connectivity to a remote computer using various protocols|
|**Test-Port**|-|Tests port on computer.|
