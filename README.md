# Module PKTools

## About
|||
|---|---|
|**Name** |PKTools|
|**Author** |Paula Kingsley|
|**Type** |Script|
|**Version** |1.38.0|
|**Date**|README.md file generated on Wednesday, February 17, 2021 4:40:37 PM|

This module contains 94 PowerShell functions or commands

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

|**Command**|**Synopsis**|
|---|---|
|**Backup-PKChromeProfile**|Backs up Chrome profiles to file|
|**Compare-PKHostnameToPTR**|Compares a hostname to the PTR of its IP, using the .NET system.net.dns class|
|**ConvertFrom-PKErrorRecord**|Converts an error record or stop exception to a more intelligible format|
|**Convert-PKBytesToSize**|Converts any integer size given to a user friendly size|
|**Convert-PKDistinguishedNameToJSON**|Uses Zachary Loeber's Get-ChildOUStructure to output the CanonicalName format of a container/OU as JSON|
|**Convert-PKDistinguishedNameToOrganizationalUnit**|Converts an Active Directory object's DistinguishedName to its parent Organizational Unit or Container, displayed in DN or CanonicalName format|
|**Convert-PKDNSRecordData**|Converts the RecordData output from Get-DNSZoneResourceRecord to a string based on record type|
|**Convert-PKEXchangeSMTPLog**|Parses an Exchange send or receive connector log from a file (string or object) and returns a PSObject|
|**Convert-PKIISLog**|Parses an IIS log from a file (string or object) and returns a PSObject|
|**Convert-PKLastLogon**|Converts an Active Directory object's LastLogonTimestamp attribute value to human-readable datetime format|
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
|**Get-PKADReplicationFailures**|Gets AD replication errors for domain controllers in a forest|
|**Get-PKBatteryReport**|Runs powercfg to create an HTML report on the utilization and history of the local computer battery|
|**Get-PKChocoPackages**|Gets a list of locally installed Chocolatey packages, interactively or as a PSJob|
|**Get-PKCompletedJobOutput**|Gets the results for completed PowerShell jobs by name or ID, with option to keep or remove results, or remove job entirely|
|**Get-PKCurrentUser**|Gets details for the currently logged-in user, including group membership, using .NET DirectoryServices.AccountManagement|
|**Get-PKDateTimeFormats**|Lists Powershell date/time format examples (standard/custom/all, or a legend)|
|**Get-PKErrorMessage**|Returns details about errors from ErrorRecord objects|
|**Get-PKFile**|Uses a queue object and .NET [IO.Directory] to search a path for files matching a string|
|**Get-PKFileEncoding**|Returns the encoding type for one or more files|
|**Get-PKIANAPorts**|Uses Invoke-Webrequest to get a CSV file from iana.org and creates a PSObject of TCP/UDP names, port numbers, and descriptions|
|**Get-PKIPConfig**|Gets IPv4 information for one or more Windows computers, interactively or as a PSJob|
|**Get-PKLocalGroupMember**|Uses 'net localgroup' to return members of a named local group or all groups|
|**Get-PKNetworkPortAssignments**|Downloads and returns a list of known port assignments from IANA|
|**Get-PKPSObjectProperties**|Returns property names from a PSCustomObject in order as an array, suitable for Select-Object|
|**Get-PKTaniumActionLog**|Invokes a scriptblock to parse content from the Tanium client action log, returning a PSObject|
|**Get-PKTaniumConfig**|Invokes a scriptblock to get Tanium client config details, interactively or as a PSJob|
|**Get-PKTimeZones**|Returns all time zone info|
|**Get-PKvCenterMetrics**|Lists all available vCenter performance metrics, optionally sorted or grouped by property names|
|**Get-PKVMIPConfig**|Returns IPv4 configuration data for one or more guest VMs|
|**Get-PKWindowsActivationStatus**|Gets the activation status for one or more remote Windows computers|
|**Get-PKWindowsDateTime**|Returns various date / time / time zone settings for a computer|
|**Get-PKWindowsHardware**|Does something cool, interactively or as a PSJob|
|**Get-PKWindowsHotfix**|Invokes a scriptblock to return installed Windows hotfixes using WMI (all or by KB number), interactively or as a PSJob|
|**Get-PKWindowsLicenseInfo**|Uses good old slmgr.vbs & WMI to retrieve Windows licensing on a local or remote computer, interactively or as a PSJob|
|**Get-PKWindowsPath**|Gets the environment path value for a computer or user or both|
|**Get-PKWindowsProductKey**|Uses WMI to retrieve a Windows product key on a local or remote computer, interactively or as a PSJob|
|**Get-PKWindowsRegistryDN**|Gets the DistinguishedName value of an AD-joined Windows computer from its registry, interactively or as a PSJob|
|**Get-PKWindowsReport**|Returns Windows computer report data, interactively or as a PSJob|
|**Get-PKWindowsRoute**|Invokes a scriptblock to get network routes on one or more computers using Get-NetRoute (available in PowerShell 4 on Windows 8/2012 and newer)|
|**Get-PKWindowsShutdown**|Invokes a scriptblock to query the Windows event log for shutdown/startup events via Get-WinEvent, or last boot time data via Get-WMIObject, interactively or as a PSJob|
|**Get-PKWindowsSoftware**|Invokes a scriptblock to return the installed software on one or more Windows computers, using registry lookups instead of WMI|
|**Get-PKWindowsTPMChipInfo**|Gets TPM chip data for a local or remote computer, interactively or as a PSJob|
|**Get-PKWinRMTrustedHosts**|Uses Get-Item to return the trusted hosts for WinRM configured on the local computer|
|**Install-PKChocoPackage**|Uses Invoke-Command to install Chocolatey packages on one or more computers, interactively or as PSJobs|
|**Install-PKWindowsPythonPackage**|Installs a Python pip package via Invoke-Command, synchronously or in a PSJob|
|**Install-PKWMIExporter**|Invokes a scriptblock to download and install WMI Exporter, with specified collectors/listening port - defaults to AD collectors|
|**Invoke-PKTelnet**|Uses System.Net.Sockets.TcpClient to test telnet connectivity on a specified port, to one or more computers, with a timeout|
|**Invoke-PKWin2019Activation**|Changes the product key and activates Windows 2019 Standard or Datacenter|
|**New-PKCodeSigningCert**|Creates a new self-signed certificate on the local computer in the current user's certificate store|
|**New-PKISESnippetFunction**|Adds a new PS ISE snippet containing a template function for ActiveDirectory, Invoke-Command, VMware, or generic use|
|**New-PKPSModuleManifest**|Creates a new PowerShell module manifest using New-ModuleManifest|
|**New-PKRandomPassword**|Generates a random password string, with option to select length|
|**Remove-GNOpsWindowsGraphiteFolder**|Invokes a scriptblock to remove the GraphitePowershell folder|
|**Remove-PKAttributeBit**|Removes one or more filesystem attribute bits from one or more files or folders (recursive)|
|**Remove-PKISECache**|Removes PowerShell ISE cache files for the local user|
|**Remove-PKMcAfee**|Removes McAfee Enterprise endpoint client from local computer without a key|
|**Reset-PKPSModule**|Removes and re-imports named PowerShell modules|
|**Resolve-PKDNSName**|Uses Resolve-DNSName to test lookups on a DNS server, with options for record type and truncated output|
|**Resolve-PKDNSNameByNET**|Uses [System.Net.Dns]::GetHostEntryAsync to resolves a hostname to an IP address, or an IP address to a hostname|
|**Resolve-PKIPtoPTR**|Performs a reverse DNSlookup on one or more IPv4 addresses, against one or more DNS server names or IP addresses|
|**Restore-PKISESession**|Restores tabs/files from text file created using Save-PKISESession|
|**Save-PKISESession**|Saves open tabs in current ISE session to a file|
|**Select-PKActiveDirectoryOU**|Uses Windows Forms to generate a selectable tree view of Active Directory containers/organizational units|
|**Set-PKWindowsExplorerPreview**|Enables the local computer's Windows Explorer preview pane for an additional file type|
|**Set-PKWinRMTrustedHosts**|Uses Set-Item to modify the trusted hosts for WinRM configured on the local computer|
|**Set-PKWMIPermissions**|Invokes a scriptblock to set WMI permissions, interactively or as a PSJob, with a huge HT to Graeme Bray & Steve Lee|
|**Show-PKErrorMessageBox**|Uses Windows Forms to display an error message in a message box (defaults to most recent error)|
|**Show-PKNestedProgressBars**|Demonstrates four nested progress bars displaying the hour, minute, second, and millisecond in a countdown to midnight|
|**Show-PKObjectTaggingDemo**|Demonstrates the differences between using Select-Object and Add-Member to add properties to an output object|
|**Show-PKPSConsoleColors**|Displays console colors, with option to show as a grid with different foreground/background combinations|
|**Show-PKSubnetMaskTable**|Displays a table of prefix lengths and dotted decimal subnet masks|
|**Test-Department**|Demonstration of dynamic dynamic parameters|
|**Test-PKADConnection**|Tests connectivity to an Active Directory domain (and optional domain controller) without the ActiveDirectory module|
|**Test-PKConnection**|Uses Get-WMIObject and the Win32_PingStatus class to quickly ping one or more targets, either interactively or as a PSJob|
|**Test-PKDNSServer**|Uses Resolve-DNSName to perform record lookups on one or more DNS servers, with optional connectivity tests and option to return lookup output|
|**Test-PKDomainCredential**|Tests authentication to Active Directory|
|**Test-PKDynamicParameter**|Demonstration of dynamic dynamic parameters|
|**Test-PKLdapSslConnection**|Tests an LDAPS connection, returning information about the negotiated SSL connection including the server certificate.|
|**Test-PKNetworkConnections**|Performs various connectivity tests to remote computers|
|**Test-PKPasswordPolicy**|Tests a string against a domain password policy for length and complexity|
|**Test-PKWindowsPendingReboot**|Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer|
|**Test-PKWinRM**|Test WinRM connectivity to a remote computer using various protocols|
|**Test-Port**|Tests port on computer.|
