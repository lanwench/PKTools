#requires -Version 4
Function Test-PKLdapSSLConnection {
<#
.SYNOPSIS
    Tests an LDAPS connection, returning information about the negotiated SSL connection including the server certificate.

.DESCRIPTION
    Test an LDAP connection, returning information about the negotiated SSL connection including the server certificate
    Permits credentials
    Defaults to port 636
    The state message "The LDAP server is unavailable" indicates the server is either offline or unwilling to negotiate an SSL connection

.NOTES
    .NOTES
    Name    : Function_Test-PKLdapSslConnection.ps1
    Created : 2020-09-15
    Author  : Paula Kingsley
    Version : 01.01.0000
    History :
        
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **
        
        2020-09-15 - v01.00.0000 - Forked by Paula from Chris Dent's original, adding Wintelrob's comments from original gist that correct enumeration 
                                   errors in output object when converting $Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm
                                   Made computername mandatory, added begin block, changed and added other stuff 
        2022-10-18 - v01.01.0000 - Updates, cleanup, added more comments

.INPUTS
    System.String

.LINK
    https://gist.github.com/indented-automation/f3b252e5e2689efe51a0f9627e0e3a0e

.LINK
    https://github.com/WintelRob

.PARAMETER Name
        One or more computer/server/target names to test

.PARAMETER Port
    Connection port (default is 636)
    
.PARAMETER Credential
    Credentials to use for the bind attempt (default is none; this command requires no special privileges)

.EXAMPLE
    PS C:\> Test-PKLdapSslConnection -computername ldaps-test.domain.local -Credential (Get-Credential DOMAIN\user) -Verbose
        
        VERBOSE: PSBoundParameters: 	
        Key              Value                                    
        ---              -----                                    
        ComputerName     ldaps-test.domain.local                    
        Credential       System.Management.Automation.PSCredential
        Verbose          True                                     
        Port             636                                      
        PipelineInput    False                                    
        ScriptName       Test-PKLdapSslConnection                 
        ScriptVersion    1.0.0    

        ComputerName         : ldaps-test.domain.local
        Port                 : 636
        State                : Connected
        Protocol             : 2048
        AlgorithmIdentifier  : Aes128
        CipherStrength       : 128
        Hash                 : Sha256
        HashStrength         : 0
        KeyExchangeAlgorithm : 44550 (ECDH_Ephem)
        ExchangeStrength     : 256
        X509Certificate      : [Subject]
                                 CN=ldaps-test.domain.local, O="Test Co", OU=Security, L=San Francisco, S=California, C=US
                       
                               [Issuer]
                                 CN=GlobalSign RSA OV SSL CA 2018, O=GlobalSign nv-sa, C=BE
                       
                               [Serial Number]
                                 0B810839C3C25D7rr61EC25E
                       
                               [Not Before]
                                 2020-09-15 12:16:08 PM
                       
                               [Not After]
                                 2021-10-17 12:16:08 PM
                       
                               [Thumbprint]
                                 FB3E093D97BE601E1BD4228B984ADF6E6123BBC2
                                 
#>

[CmdletBinding()]
[OutputType('Indented.LDAP.ConnectionInformation')]

Param (
    
    [Parameter(
        Mandatory = $True,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromPipeline = $true,
        HelpMessage = "The name of a computer/server to test"
    )]
    [Alias("DnsHostName","ComputerName")]
    [String[]]$Name,

    [Parameter(
        HelpMessage = "Connection port (default is 636)"
    )]
    [UInt16]$Port = 636,

    [Parameter(
        HelpMessage = "Credentials to use for the bind attempt (default is none; this command requires no special privileges)"
    )]
    [PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty

)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.01.0000"

    # How did we get here
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    # Display our parameters
    $CurrentParams = $PSBoundParameters
    
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    $Activity = "Test LDAPS connection"
    $Msg = "[BEGIN: $Scriptname] $Activity" 
    Write-Verbose $Msg

}
process {
    
    $Total = $Name.Count
    $Current = 0
    
    Foreach ($Computer in $Name) {
        
        $Current ++
        
        $Msg = "Creating new DirectoryServices object using port $Port"
        Write-Verbose "[$Computer] $Msg"
        Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete ($Current/$Total*100)
        
        Try {
            $DirectoryIdentifier = New-Object DirectoryServices.Protocols.LdapDirectoryIdentifier($Computer, $Port)
            $Auth = "Kerberos"
            $Args = @($DirectoryIdentifier)
            If ($CurrentParams.Credential.Username) {
                $Args += $Credential.GetNetworkCredential()
                $Auth = "Basic"
            }
            $Connection = New-Object DirectoryServices.Protocols.LdapConnection($Args)
            $Connection.AuthType = [DirectoryServices.Protocols.AuthType]::$Auth
            
            $Connection.SessionOptions.ProtocolVersion = 3
            $Connection.SessionOptions.SecureSocketLayer = $True

            Try {

                # Declare a script level variable which can be used to return information from the delegate.
                $Msg = "Creating script variable to hold information"
                Write-Verbose "[$Computer] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete ($Current/$Total*100)
                New-Variable LdapCertificate -Scope Script -Force

                $Msg = "Creating callback delegate to retrieve certificate"
                Write-Verbose "[$Computer] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete ($Current/$Total*100)
                <# Create a callback delegate to retrieve the negotiated certificate.
                    Note:
                    * The certificate is unlikely to return the subject.
                    * The delegate is documented as using the X509Certificate type, automatically casting this to X509Certificate2 allows access to more information.
                #>
                $Connection.SessionOptions.VerifyServerCertificate = {
                    param(
                        [DirectoryServices.Protocols.LdapConnection]$Connection,
                        [Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
                    )
                    $Script:LdapCertificate = $Certificate
                    return $true
                }

                $Msg = "Testing connection"
                Write-Verbose "[$Computer] $Msg"
                Write-Progress -Activity $Activity -CurrentOperation $Msg -Status $Computer -PercentComplete ($Current/$Total*100)

                $state = "Connected"  
                Try {
                    $Connection.Bind()
                    $Msg = "LDAPS connection successful"
                    Write-Verbose "[$Computer] $Msg"
                    $Messages += $Msg
                } 
                Catch {
                    $state = "Connection failed ($($_.Exception.InnerException.Message.Trim()))"
                    $Msg = "LDAPS connection failed"
                    Write-Warning "[$Computer] $Msg"
                    $Messages += $Msg
                }

                <# Wintelrob's addition:
                    Code to handle when the system cannot interpret the newer code '44550'
                    to convert to the algorithm name. We just don't get the type name and
                    convert the number to a string to combine with the name we think it 
                    should be.
                #>
                If ([Int]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm -eq 44550){
                    $KeyExchangeAlgo = [string]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm + ' (ECDH_Ephem)'
                }
                Else {
                    If (-not ($KeyExchangeAlgo = [Security.Authentication.ExchangeAlgorithmType][Int]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm)) {
                        $KeyExchangeAlgo = $Null
                    }
                }
            }
            Catch {
                $Msg = "Failed to create delegate $($_.Exception.Message)"
                Write-Warning "[$Computer] $Msg"
                $Messages += $Msg
            }
        }
        Catch {
            $Msg = "Failed to create DirectoryServices object $($_.Exception.Message)"
            Write-Warning "[$Computer] $Msg"
            $Messages += $Msg
        }


        [PSCustomObject]@{
            ComputerName         = $Computer
            Port                 = $Port
            State                = $State
            Protocol             = $Connection.SessionOptions.SslInformation.Protocol
            AlgorithmIdentifier  = $Connection.SessionOptions.SslInformation.AlgorithmIdentifier
            CipherStrength       = $Connection.SessionOptions.SslInformation.CipherStrength
            Hash                 = $Connection.SessionOptions.SslInformation.Hash
            HashStrength         = $Connection.SessionOptions.SslInformation.HashStrength
            #KeyExchangeAlgorithm = [Security.Authentication.ExchangeAlgorithmType][Int]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm
            KeyExchangeAlgorithm = $KeyExchangeAlgo
            ExchangeStrength     = $Connection.SessionOptions.SslInformation.ExchangeStrength
            X509Certificate      = $Script:LdapCertificate
            PSTypeName           = 'Indented.LDAP.ConnectionInformation'
            Messages             = $Messages          
        }

        Try {
            $Connection.Dispose()
        }
        Catch {}
    
    } #end foreach

}
End {
    
    Write-Progress -Activity * -Completed
    $Msg = "END: $Scriptname] $Activity" 
    Write-Verbose $Msg

}
} # End Test-PKLDAPSSLConnection