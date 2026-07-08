#requires -Version 4.0
Function Get-PKDomainInfo {
	<#
	.SYNOPSIS
		Gets details for one or more domains using the RDAP server for that TLD, retrieved via IANA RDAP bootstrap registry 

	.DESCRIPTION
		Gets details for one or more domains using the RDAP server for that TLD, retrieved via IANA RDAP bootstrap registry 

		For thin registries (like .com, .net), the registry record includes a "related" link pointing to the registrar's 
			own RDAP endpoint - we follow that link and use the registrar record as the primary source for client* lock 
			status codes and contact details

		For thick registries (like .org, .io, most ccTLDs, new gTLDs), the registry record contains everything and no 
			second hop is needed

		Returns one object per domain with:
			- DomainName, Registrar, Nameservers
			- WHOIS: registrant email if exposed, otherwise "Blocked"
			- ClientDeleteProhibited, ClientTransferProhibited, 
				ClientRenewProhibited, ClientUpdateProhibited (explicit booleans)
			- DNSSECSigned (boolean)
			- Expiration and LastChanged (ISO 8601)
			- AllStatusCodes, StatusSource, RegistryUri, RegistrarUri

		EPP status flags absent from the RDAP status array are returned as an explicit $False (not $Null, as returned by the RDAP results)

		Handles both RDAP status string formats:
			- spaced lowercase: "client transfer prohibited" (e.g. Verisign)
			- camelCase: "clientTransferProhibited" (e.g. GoDaddy)

	.NOTES
		Name    : Function_Get-PKDomainInfo.ps1
		Created : 2026-06-15
		Author  : Paula Kingsley
		Version : 01.00
		History :
			** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

			v01.00 - 2026-06-15 - Created script

	.PARAMETER Domain
		One or more domain names to query (e.g., example.com)

	.EXAMPLE
		PS > Get-PKDomainInfo wired.com -Verbose    
		Gets the registrar & DNS hosting info for a single domain

			VERBOSE: PSBoundParameters: 

			Key           Value
			---           -----
			Verbose       True
			Domain        {wired.com}
			ScriptName    Get-PKDomainInfo
			ScriptVersion 1.0
			PipelineInput False

			VERBOSE: [BEGIN: Get-PKDomainInfo] Get domain info using IANA RDAP bootstrap registry
			VERBOSE: [Prerequisite] Downloading IANA RDAP bootstrap file
			VERBOSE: [wired.com] Processing domain
			VERBOSE: [wired.com] Querying registry RDAP at https://rdap.verisign.com/com/v1/domain/wired.com
			VERBOSE: [wired.com] Found registrar RDAP link, querying https://rdap.cscglobal.com/dbs/rdap-api/v1/domain/WIRED.COM

			DomainName               : wired.com
			Registrar                : CSC Corporate Domains, Inc.
			Nameservers              : {NS-1116.AWSDNS-11.ORG, NS-1935.AWSDNS-49.CO.UK, NS-28.AWSDNS-03.COM, NS-836.AWSDNS-40.NET}
			WHOISContact             : Blocked
			ClientDeleteProhibited   : False
			ClientTransferProhibited : True
			ClientRenewProhibited    : False
			ClientUpdateProhibited   : False
			DNSSECSigned             : False
			AllStatusCodes           : {server delete prohibited, server update prohibited, server transfer prohibited, client transfer prohibited}
			StatusSource             : Registrar
			RegistryUri              : https://rdap.verisign.com/com/v1/domain/wired.com
			RegistrarUri             : https://rdap.cscglobal.com/dbs/rdap-api/v1/domain/WIRED.COM
			DateRegistered           : 1992-11-20T00:00:00Z
			DateExpires              : 2027-11-19T05:00:00Z
			DateLastChanged          : 2026-02-02T16:24:26Z

			VERBOSE: [wired.com] Sleeping for 500 seconds to avoid angering the rate-limit bear
			VERBOSE: [END: Get-PKDomainInfo] Script ran successfully

.EXAMPLE
	PS > "nytimes.com","bogus.thisisfake.info","boingboing.net" | Get-PKDomainInfo | Select-Object DomainName,Registrar,DateRegistered

	Gets the registrar & registration date for domain names in the pipeline

		WARNING: [bogus.thisisfake.info] Registry RDAP query failed: Response status code does not indicate success: 400 (Bad Request).
		
		DomainName     Registrar                  DateRegistered
		----------     ---------                  --------------
		nytimes.com    Markmonitor Inc.           1994-01-18T05:00:00Z
		boingboing.net Squarespace Domains II LLC 1998-05-06T04:00:00Z

	#>

	[CmdletBinding()]
	Param(
		[Parameter(
			Mandatory,
			Position = 0,
			ValueFromPipeline,
			HelpMessage = "Domain name to query (e.g., example.com)"
		)]
		[ValidateNotNullOrEmpty()]
		[string[]]$Domain
	)

	Begin {

		[version]$Version = "01.00"

		# How did we get here?
		$ScriptName = $MyInvocation.MyCommand.Name
		$Source = $PSCmdlet.ParameterSetName
		[switch]$PipelineInput = $MyInvocation.ExpectingInput
		$CurrentParams = $PSBoundParameters
		$MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } |
			Where-Object { Test-Path Variable:$_ } | ForEach-Object {
				$CurrentParams.Add($_, (Get-Variable $_).value)
			}
		$CurrentParams.Add("ParameterSetName", $Source)
		$CurrentParams.Add("PipelineInput", $PipelineInput)
		$CurrentParams.Add("ScriptName", $ScriptName)
		$CurrentParams.Add("ScriptVersion", $Version)
		Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | Out-String)"
		
		Write-Verbose "[BEGIN: $ScriptName] Get domain info using IANA RDAP bootstrap registry"

		# IANA bootstrap registry maps TLD to its RDAP base URL
		[string]$BootstrapUri = "https://data.iana.org/rdap/dns.json"
		Try {
			Write-Verbose "[Prerequisite] Downloading IANA RDAP bootstrap file"
			$Bootstrap = Invoke-RestMethod -Uri $BootstrapUri -ErrorAction Stop -Verbose:$False
		}
		Catch {
			$PSCmdlet.ThrowTerminatingError($_)
		}

		#region Inner functions

		# Checks a status array for named EPP flags in either known RDAP string format
		Function _TestPKStatusFlag {
			Param([string[]]$StatusArray)
			# Normalize the entire array to spaced lowercase once, so we can match against both formats
			[string[]]$Normalized = $StatusArray | ForEach-Object { ($_ -creplace '([A-Z])', ' $1').Trim().ToLower() }
			[pscustomobject]@{
				ClientDeleteProhibited   = $Normalized -contains 'client delete prohibited'
				ClientTransferProhibited = $Normalized -contains 'client transfer prohibited'
				ClientRenewProhibited    = $Normalized -contains 'client renew prohibited'
				ClientUpdateProhibited   = $Normalized -contains 'client update prohibited'
			}
		}

		# Returns registrant email from an RDAP entity's vcardArray (first email, then fallback to contact-uri, or $Null if not found)
		Function _GetVcardEmail {
			Param([object[]]$Entities, [string]$Role)
			$Entity = $Entities | Where-Object { $_.roles -contains $Role } | Select-Object -First 1
			If (-not $Entity) { Return $Null }
			$VcardProps = $Entity.vcardArray[1]
			$EmailProp = $VcardProps | Where-Object { $_[0] -eq 'email' } | Select-Object -First 1
			If ($EmailProp) { Return $EmailProp[3] }
			# contact-uri is a mailto: URI used in place of email by some registrars (e.g. GoDaddy)
			$ContactProp = $VcardProps | Where-Object { $_[0] -eq 'contact-uri' } | Select-Object -First 1
			If ($ContactProp -and $ContactProp[3] -like 'mailto:*') {
				Return $ContactProp[3] -replace '^mailto:', ''
			}
			Return $Null
		}

		# Returns ISO 8601 date string from RDAP events array, trying each eventAction in order, returningfirst match
		Function _GetEventRDAPDate {
			Param([object[]]$Events, [string[]]$EventAction)
			ForEach ($Action in $EventAction) {
				$Match = $Events | Where-Object { $_.eventAction -eq $Action } | Select-Object -First 1
				If ($Match) {
					Try { Return ([datetime]$Match.eventDate).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") }
					Catch { Return $Match.eventDate }
				}
			}
		}
		
		# Get dates from either events array, primary first
		Function _GetEventRDAPDates {
			Param([object[]]$PrimaryEvents,[object[]]$RegistryEvents)
			[object[]]$AllEvents = $PrimaryEvents + $RegistryEvents
			[pscustomobject]@{
				Registered  = ($AllEvents | Where-Object { $_.eventAction -eq 'registration' } | Select-Object -First 1).eventDate
				Expiration  = ($AllEvents | Where-Object { $_.eventAction -in 'registrar expiration','expiration' } | Select-Object -First 1).eventDate
				LastChanged = ($AllEvents | Where-Object { $_.eventAction -eq 'last changed' } | Select-Object -First 1).eventDate
			}
		}

		#endregion Inner functions

	}
	Process {

		ForEach ($D in $Domain) {

			Write-Verbose "[$D] Processing domain"
			[string]$TLD = ($D -split '\.')[-1].ToLower()
			
			# Find the RDAP service entry matching this TLD
			[string]$RdapBase =$Null
			ForEach ($Service in $Bootstrap.services) {
				If ($Service[0] -contains $TLD) {
					$RdapBase = $Service[1][0].TrimEnd('/')
					Break
				}
			}

			If (-not $RdapBase) {
				Write-Warning "[$D] No RDAP service found for TLD '$TLD'"
				Continue
			}

			[string]$RegistryUri = "$RdapBase/domain/$D"
			Write-Verbose "[$D] Querying registry RDAP at $RegistryUri"

			[object]$RegistrarResponse = [string]$RegistrarUri = [string]$StatusSource = $Null
			Try {
				$RegistryResponse = Invoke-RestMethod -Uri $RegistryUri -ErrorAction Stop -Verbose:$False
			}
			Catch {
				Write-Warning "[$D] Registry RDAP query failed: $($_.Exception.Message)"
				Continue
			}

			# "related" link points to registrar's RDAP record fpr thin registries - if found, all client status codes/contacts live there
			If ($RelatedLink = $RegistryResponse.links | Where-Object { $_.rel -eq 'related' } | Select-Object -First 1) {
				$RegistrarUri = $RelatedLink.href
				Write-Verbose "[$D] Found registrar RDAP link, querying $RegistrarUri"
				Try {
					$RegistrarResponse = Invoke-RestMethod -Uri $RegistrarUri -ErrorAction Stop -Verbose:$False
					$StatusSource = "Registrar"
				}
				Catch {
					Write-Warning "[$D] Registrar RDAP query failed; falling back to registry data: $($_.Exception.Message)"
				}
			}

			# Use registrar record if we got one, otherwise fall back to registry record
			[object]$PrimaryResponse = If ($RegistrarResponse) { $RegistrarResponse } Else { $RegistryResponse }
			If (-not $StatusSource) { $StatusSource = "Registry" }

			[string[]]$StatusArray = $PrimaryResponse.status

			# Get the registrar name from entity vcard 'fn'
			[string]$RegistrarName = $Null
			$RegistrarEntity = $PrimaryResponse.entities | Where-Object { $_.roles -contains 'registrar' } | Select-Object -First 1
			If ($RegistrarEntity) {
				$FnProp = $RegistrarEntity.vcardArray[1] | Where-Object { $_[0] -eq 'fn' } | Select-Object -First 1
				If ($FnProp) { $RegistrarName = $FnProp[3] }
			}

			# Get the nameservers (ideally from registry, as registrar RDAP often omits info)
			[string[]]$Nameservers = $RegistryResponse.nameservers.ldhName
			If (-not $Nameservers) { $Nameservers = $PrimaryResponse.nameservers.ldhName }

			# Get the email contact in WHOIS, or return "Blocked" if redacted/proxied/absent
			[string]$RegistrantEmail = _GetVcardEmail -Entities $PrimaryResponse.entities -Role 'registrant'
			[string]$Whois = If ([string]::IsNullOrWhiteSpace($RegistrantEmail)) { "Blocked" } Else { $RegistrantEmail }

			# For DNSSEC delegationSigned = $True means DS records are published and the zone is signed (prefer primary, fall back to registry)
			[bool]$DNSSECSigned = If ($PrimaryResponse.secureDNS) {
				[bool]$PrimaryResponse.secureDNS.delegationSigned
			}
			ElseIf ($RegistryResponse.secureDNS) {
				[bool]$RegistryResponse.secureDNS.delegationSigned
			}
			Else { $False }

			# Get dates for registered/expiration/last updated
			$Dates = _GetEventRDAPDates -PrimaryEvents $PrimaryResponse.events -RegistryEvents $RegistryResponse.events

			# Get domain transfer lock status by checking everything in the status property
			$LockStatus = _TestPKStatusFlag -StatusArray $StatusArray

			[PSCustomObject]@{
				DomainName               = $D
				Registrar                = $RegistrarName
				Nameservers              = $Nameservers
				WHOISContact             = $Whois
				ClientDeleteProhibited   = $LockStatus.ClientDeleteProhibited
				ClientTransferProhibited = $LockStatus.ClientTransferProhibited
				ClientRenewProhibited    = $LockStatus.ClientRenewProhibited
				ClientUpdateProhibited   = $LockStatus.ClientUpdateProhibited
				DNSSECSigned             = $DNSSECSigned
				AllStatusCodes           = [array]$StatusArray
				StatusSource             = $StatusSource
				RegistryUri              = $RegistryUri
				RegistrarUri             = $RegistrarUri
				DateRegistered  		 = & {Try { ([datetime]$Dates.Registered).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } Catch { $Dates.Registered }}
				DateExpires  			 = & {Try { ([datetime]$Dates.Expiration).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } Catch { $Dates.Expiration }}
				DateLastChanged 		 = & {Try { ([datetime]$Dates.LastChanged).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") } Catch { $Dates.LastChanged }}
			}
			Write-Verbose "[$D] Sleeping for 500 seconds to avoid angering the rate-limit bear"
			Start-Sleep -Milliseconds 500
		} # end foreach

	}
	End {
		Write-Verbose "[END: $ScriptName] Script ran successfully"
	}
} #end Get-PKDomainInfo