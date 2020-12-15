#requires -Version 4
Function Test-PKLdapSslConnection {
<#
.SYNOPSIS
    Tests an LDAPS connection, returning information about the negotiated SSL connection including the server certificate.

.DESCRIPTION
    Test an LDAP connection, returning information about the negotiated SSL connection including the server certificate
    Permits credentials
    Defaults to port 636
    The state message "The LDAP server is unavailable" indicates the server is either offline or unwilling to negotiate an SSL connection

.NOTES
    2020-09-15 - v01.00.0000 - Forked by Paula from Chris Dent's original, adding Wintelrob's comments from original gist that correct enumeration 
                               errors in output object when converting $Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm
                               Made computername mandatory, added begin block, changed and added other stuff 
.INPUTS
    System.String

.LINK
    https://gist.github.com/indented-automation/f3b252e5e2689efe51a0f9627e0e3a0e

.LINK
    https://github.com/WintelRob

.EXAMPLE
    PS C:\> Test-LdapSslConnection -computername ldaps-test.nlsn.media -Credential $PKCredMedia -Verbose
        
        VERBOSE: PSBoundParameters: 	
        Key              Value                                    
        ---              -----                                    
        ComputerName     ldaps-test.nlsn.media                    
        Credential       System.Management.Automation.PSCredential
        Verbose          True                                     
        Port             636                                      
        PipelineInput    False                                    
        ScriptName       Test-PKLdapSslConnection                 
        ScriptVersion    1.0.0    

        ComputerName         : ldaps-test.nlsn.media
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
                                 CN=ldaps-test.nlsn.media, O="Gracenote, Inc", OU=Information Security, L=Emeryville, S=California, C=US
                       
                               [Issuer]
                                 CN=GlobalSign RSA OV SSL CA 2018, O=GlobalSign nv-sa, C=BE
                       
                               [Serial Number]
                                 0B940839C3C25D7CC61EC29B
                       
                               [Not Before]
                                 2020-09-15 12:16:08 PM
                       
                               [Not After]
                                 2021-10-17 12:16:08 PM
                       
                               [Thumbprint]
                                 FB3E093D97BE601E1BD4228B984ADF6E6123BBC2
                                 Test-LdapSSLConnection

    Attempt to bind using SSL and serverless binding.

.EXAMPLE
    Test-LdapSSLConnection -ComputerName servername
        
    Attempt to negotiate SSL with "servername".




#>

[CmdletBinding()]
[OutputType('Indented.LDAP.ConnectionInformation')]

Param (
    
    [Parameter(
        Mandatory = $True,
        ValueFromPipelineByPropertyName = $true, 
        ValueFromPipeline = $true,
        HelpMessage = "The name of a computer to test"
    )]
    [Alias("DnsHostName")]
    [String]$ComputerName = "",

    [Parameter(
        HelpMessage = "Connection port (default is 636)"
    )]
    [UInt16]$Port = 636,

    [Parameter(
        HelpMessage = "Credentials to use for the bind attempt (default is none; this command requires no special privileges)"
    )]
    [PSCredential]$Credential,

    [Parameter(
        HelpMessage = "Suppress non-verbose console output"
    )]
    [switch]$Quiet
)

Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

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

    # Function to write verbose message, collecting error data, and optional prefixes
    Function Write-MessageVerbose {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Verbose $Message
    }

    # Function to write a console message or a verbose message
    Function Write-MessageInfo {
        Param([Parameter(ValueFromPipeline)]$Message,$FGColor,[switch]$Title)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {
            If ($Title.IsPresent) {$Message = "`n$Message`n"}
            $Host.UI.WriteLine($FGColor,$BGColor,"$Message")
        }
        Else {Write-Verbose "$Message"}
    }

    # Function to write an error as a string (no stacktrace), or an error, with optional prefixes
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        If (-not $Quiet.IsPresent) {
            $Host.UI.WriteErrorLine("$Message")
        }
        Else {Write-Error "$Message"}
    }
    
    # Function to write an error/warning, collecting error data, with optional prefixes
    Function Write-MessageWarning {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message,[switch]$PrefixPrerequisites,[switch]$PrefixError)
        If ($PrefixPrerequisites.IsPresent) {$Message = "[Prerequisites] $Message"}
        Elseif ($PrefixError.IsPresent) {$Message = "ERROR  :  $Message"}
        If ($ErrorDetails = $_.Exception.Message) {$Message += " ($ErrorDetails)"}
        Write-Warning $Message
    }

    #endregion Splats



}
process {

    Foreach ($Computer in $ComputerName) {
        
        $Msg = "[$Computer] Test LDAPS connection"
        $Msg | Write-MessageVerbose

        $DirectoryIdentifier = New-Object DirectoryServices.Protocols.LdapDirectoryIdentifier($Computer, $Port)
        if ($PSBoundParameters.ContainsKey("Credential")) {
            $Connection = New-Object DirectoryServices.Protocols.LdapConnection($DirectoryIdentifier, $Credential.GetNetworkCredential())
            $Connection.AuthType = [DirectoryServices.Protocols.AuthType]::Basic
        } 
        Else {
            $Connection = New-Object DirectoryServices.Protocols.LdapConnection($DirectoryIdentifier)
            $Connection.AuthType = [DirectoryServices.Protocols.AuthType]::Kerberos
        }
        $Connection.SessionOptions.ProtocolVersion = 3
        $Connection.SessionOptions.SecureSocketLayer = $True

        # Declare a script level variable which can be used to return information from the delegate.
        New-Variable LdapCertificate -Scope Script -Force

        # 
        # Create a callback delegate to retrieve the negotiated certificate.
        # Note:
        #   * The certificate is unlikely to return the subject.
        #   * The delegate is documented as using the X509Certificate type, automatically casting this to X509Certificate2 allows access to more information.
        $Connection.SessionOptions.VerifyServerCertificate = {
            param(
                [DirectoryServices.Protocols.LdapConnection]$Connection,
                [Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
            )
            $Script:LdapCertificate = $Certificate
            return $true
        }

        $state = "Connected"  
        Try {
            $Connection.Bind()
            $Msg = "[$Computer] LDAPS connection successful"
            $Msg | Write-MessageInfo -FGColor Green
        } 
        Catch {
            $state = "Connection failed ($($_.Exception.InnerException.Message.Trim()))"
            $Msg = "[$Computer] LDAPS connection failed"
            $Msg | Write-MessageError
        }

        # Wintelrob's addition:
        # Code to handle when the system cannot interpret the newer code '44550'
        # to convert to the algorithm name. We just don't get the type name and
        # convert the number to a string to combine with the name we think it 
        # should be.
        If ([Int]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm -eq 44550){
            $KeyExchangeAlgo = [string]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm + ' (ECDH_Ephem)'
        }
        Else{
            #$KeyExchangeAlgo = [Security.Authentication.ExchangeAlgorithmType][Int]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm
            If (-not ($KeyExchangeAlgo = [Security.Authentication.ExchangeAlgorithmType][Int]$Connection.SessionOptions.SslInformation.KeyExchangeAlgorithm)) {
                $KeyExchangeAlgo = $Null
            }
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
        }

        Try {
            $Connection.Dispose()
        }
        Catch {}
    
    
    
    
    
    
    
    
    
    
    
    }


    
}
}