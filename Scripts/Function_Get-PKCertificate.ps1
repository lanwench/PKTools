
#requires -Version 4
Function Get-PKCertificate {
<#
.SYNOPSIS
    Retrieves SSL/TLS certificate details from one or more remote hosts by performing a TCP connection and SSL handshake.

.DESCRIPTION
    Connects to specified hostnames on a given TCP port (default 443), performs an SSL/TLS handshake using System.Net.Security.SslStream, and returns key certificate metadata such as subject, issuer, validity period, thumbprint and Subject Alternative Names. 
    The function accepts pipeline input and will normalize incoming values (stripping URL schemes and ports) before attempting connection.
    This function intentionally bypasses certificate chain/hostname validation during the handshake to allow retrieving certificates from hosts regardless of trust status. 
    It is intended for information-gathering and monitoring purposes only and should not be used as an indicator of certificate trust.

    - Invalid hostname syntax: The function warns and skips hostnames containing illegal characters.
    - Connection timeout: If the TCP connection cannot be established within TimeoutSeconds, the operation throws a timeout error.
    - Handshake failure: SSL/TLS handshake errors (protocol mismatch, abrupt closure, etc.) are reported as errors for that host.
    - No certificate available: If a connection succeeds but no remote certificate is presented, an error is emitted.
    
    Designed for remote server interrogation. To inspect certificates from Windows certificate stores, use the PKI / Cert: provider or related cmdlets.

.NOTES        
    Name    : Function_Get-PKCertificate.ps1
    Created : 2025-12-19
    Author  : Paula Kingsley
    Version : 01.00.1000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2025-12-19- Created script

.PARAMETER Hostname
    One or more hostnames (or URLs) to query. Accepts pipeline input (ValueFromPipeline and ValueFromPipelineByPropertyName). Aliases: Name, ComputerName, ServerName, Domain, URL.
    The value will be normalized by removing an optional leading "http(s)://" and any appended port or path. Hostnames must match alphanumeric characters, hyphens and periods.

.PARAMETER Port
    TCP port to connect to on the remote host. Default is 443.

.PARAMETER TimeoutSeconds
    Timeout in seconds for establishing the TCP connection. Default is 5 seconds. A connection attempt that exceeds this timeout will fail with a timeout error.

.INPUTS
    String (hostname) - Accepted via pipeline input.

.OUTPUTS
    System.Management.Automation.PSCustomObject
    Properties:
    - Host        : Normalized hostname that was contacted.
    - Subject     : Certificate subject distinguished name.
    - FriendlyName: Comma-separated friendly names for enhanced key usage (if available).
    - Issuer      : Certificate issuer distinguished name.
    - ValidFrom   : Certificate NotBefore (DateTime).
    - Expiration  : Certificate NotAfter (DateTime).
    - Thumbprint  : Certificate thumbprint string.
    - SANs        : Formatted Subject Alternative Name extension contents.

.EXAMPLE 
    # Query a single host (default port 443)
    PS C:> Get-PKCertificate -Hostname "example.com"

    # Query multiple hosts
    PS C:> Get-PKCertificate -Hostname "example.com","api.example.org","10.0.0.5"

    # Use pipeline input (one hostname per line in a file)
    PS C:> Get-Content .\hosts.txt | Get-PKCertificate

    # Query a host on a custom port
    PS C:> Get-PKCertificate -Hostname "domaincontroller.example.com" -Port 636

    # Increase timeout to allow for slower hosts
    PS C:> Get-PKCertificate -Hostname "slow.example.com" -TimeoutSeconds 15

#>
    [CmdletBinding()]
    Param(
        [Parameter(
            Position = 0,
            Mandatory,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName,
            HelpMessage = "One or more hostnames to retrieve the certificate from"
        )]
        [Alias("Name","ComputerName","ServerName","Domain","URL")]
        [string[]]$Hostname,

        [Parameter(
            HelpMessage = "TCP port for connection (default is 443)"
        )][int]$Port = 443,
        [Parameter(
            HelpMessage = "Timeout in seconds for connection (default is 5)"
        )]
        [int]$TimeoutSeconds = 5

    )
    Begin {
        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        $ScriptName = $MyInvocation.MyCommand.Name
        $CurrentParams = $PSBoundParameters
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
        Where-Object { Test-Path variable:$_ } | Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
        $CurrentParams.Add("ScriptName", $ScriptName)
        $CurrentParams.Add("PipelineInput", $PipelineInput)
        $CurrentParams.Add("ScriptVersion", $Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

        Write-Verbose "[BEGIN: $ScriptName] Get certificate details on port $Port"

    }
    Process {

        Foreach ($Name in $Hostname) {
            Write-Verbose "[$Name] Validating hostname syntax"
            $Normalized = $Name.Trim() -replace '^(?:https?://)?([^:/]+).*$', '$1' # getting rid of extra junk
            If ($Normalized -notmatch '^[a-zA-Z0-9.-]+$') {
                Write-Warning "[$Normalized] Normalized hostname contains illegal characters (anything other than alphanumeric, hyphens, and periods)"
            }
            Else {
                $Target = $Normalized 

                Write-Verbose "[$Target] Creating TCP socket"
                Try { 
                    $TimeoutMS = $TimeoutSeconds * 1000
                    $TCPClient = New-Object System.Net.Sockets.TcpClient
                    
                    Write-Verbose "[$Target] Attempting connection with a $TimeoutSeconds second timeout"
                    $connectTask = $tcpClient.ConnectAsync($Target,$Port)  # Connect with a timeout (using a Task to ensure it doesn't hang)
                    If (-not $connectTask.Wait($TimeoutMS)) {
                        Throw "[$Target] Connection timed out after $TimeoutSeconds seconds."
                    }
                    Else {
                        Write-Verbose "[$Target] Creating SSL stream"

                        $SSLStream = New-Object System.Net.Security.SslStream($TCPClient.GetStream(), $false, { $true }) # Adding 'true' aallows the cert to stay open after we close the stream

                        Write-Verbose "[$Target] Forcing handshake"
                        $SSLStream.AuthenticateAsClient($Target)

                        Write-Verbose "[$Target] Extracting certificate data"
                        If ($Certificate = $SSLStream.RemoteCertificate) {
                            $SANs = ($Certificate.Extensions | Where-Object { $_.Oid.FriendlyName -match "Subject Alternative Name" }).Format($true)

                            [PSCustomObject]@{
                                Host         = $Target
                                Subject      = $Certificate.Subject
                                FriendlyName = $Certificate.EnhancedKeyUsageList.FriendlyName -join(", ")
                                Issuer       = $Certificate.Issuer
                                ValidFrom    = $Certificate.NotBefore
                                Expiration   = $Certificate.NotAfter
                                Thumbprint   = $Certificate.Thumbprint
                                SANs         = $SANs
                            }
                        }
                        Else {
                            Throw "[$Target] No certificate data available"
                        }
                    } # end else
                }
                Catch {
                    Write-Error "[$Target] Operation failed! $($_.Exception.Message)"
                }
                Finally {
                    Write-Verbose "[$Target] Closing any remaining connections"
                    If ($SSLStream) {$SSLStream.Close()}
                    If ($TCPClient) {$TCPClient.Close()}
                    
                } 
            } # end valid hostname
        } # end loop      
    } # End Process
    End {
        Write-Verbose "[END: $ScriptName] Operation complete"
    }
} # end Get-PKCertificate