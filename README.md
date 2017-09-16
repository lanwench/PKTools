# Module PKTools

## About
| | | |
|-|-|-|
Module name | PKTools
Author  | Paula Kingsley
Module version | 1.0.0
Module type | Script
Minimum PowerShell version | 3.0.0
README.md generated | Saturday, September 16 09:39:07 

This module contains 8 PowerShell function(s)

_Functions should be presumed to be authored by Paula Kingsley unless otherwise specified (see the context help within each function for more information, including credits)._

All functions should have reasonably detailed comment-based help, accessible via Get-Help, e.g.,
  * `Get-Help Do-Something`
  * `Get-Help Do-Something -Examples`
  * `Get-Help Do-Something -ShowWindow`

## Prerequisites ##

Computers must:
  * be running PowerShell 3.0.0 or later

## Installation ##

Clone/copy entire module directory into a valid PSModules folder on your computer and run `Import-Module PKTools`

## Functions

#### Get-PKADDomainController ####
Returns a domain controller for the current computer site or a named site in the current or named domain

#### Get-PKFileEncoding ####
Returns the encoding type for one or more files

#### Install-PKChocoPackage ####
Uses Invoke-Command to install Chocolatey packages on remote computers as PSJobs

#### Reset-PKPSModule ####
Removes and re-imports named PowerShell modules

#### Restore-PKISESession ####
Restores tabs/files from file saved using Save-PKISESession

#### Save-PKISESession ####
Saves open tabs in current ISE session to a file

#### Set-PKExplorerPreview ####
Enables the local computer's Windows Explorer preview pane for an additional file type

#### Test-PKWindowsPendingReboot ####
Invokes a PSJob using Invoke-Command to test the pending reboot status of a computer
