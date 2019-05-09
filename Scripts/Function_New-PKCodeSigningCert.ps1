#Requires -version 3
function New-PKCodeSigningCert {
<#
.SYNOPSIS
    Creates a new self-signed certificate on the local computer in the current user's certificate store

.DESCRIPTION
    Creates a new self-signed certificate on the local computer in the current user's certificate store
    Checks for existing certificate and exits if true
    Supports ShouldProcess
    Returns a string

.NOTES
    Name    : Function_New-PKCodeSigningCert.ps1 
    Created : 2019-05-08
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK ** 

        v01.00.0000 - 2019-05-08 - Created script based on Tobias Weltner's original

.LINK
    https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/creating-code-signing-certificates

.PARAMETER FriendlyName
    Friendly name for certificate

.PARAMETER Name
    Certificate name (will become CN=Name)

.PARAMETER YearsValid
    Certificate validity in years (default 5)

.PARAMETER Quiet
    Suppress all non-verbose console output

.EXAMPLE
    PS C:\> New-PKCodeSigningCert -FriendlyName "Test cert" -Name PKTestCert1 -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                
        ---           -----                
        FriendlyName  Test cert            
        Name          PKTestCert1          
        Verbose       True                 
        YearsValid    5                    
        Quiet         False                
        PipelineInput False                
        ScriptName    New-PKCodeSigningCert
        ScriptVersion 1.0.0                

        BEGIN  : Create self-signed certificate for code signing
        
        [WORKSTATION99] Test for existing self-signed certificate
        [WORKSTATION99] Create new self-signed certificate
        [WORKSTATION99] Successfully created self-signed certificate on WORKSTATION99 in Cert:\CurrentUser\My
        
        [Subject]
          CN=PKTestCert1

        [Issuer]
          CN=PKTestCert1

        [Serial Number]
          1F9942DAE28887AE4BEECD790775AF61

        [Not Before]
          2019-05-08 05:29:58 PM

        [Not After]
          2024-05-08 05:39:57 PM

        [Thumbprint]
          1AF19EE2EBB15033C5EB86E3641384AEE7323E5E

        END    : Create self-signed certificate for code signing

.EXAMPLE
    PS C:\> New-PKCodeSigningCert -FriendlyName "Test cert" -Name PKTestCert1
        
        BEGIN  : Create self-signed certificate for code signing
        [WORKSTATION99] Test for existing self-signed certificate
        [WORKSTATION99] Certificate already exists; please remove it and run this script again

           PSParentPath: Microsoft.PowerShell.Security\Certificate::CurrentUser\My

        Thumbprint                                Subject                                         
        ----------                                -------                                         
        1AF19EE2EBB15033C5EB86E3641384AEE7323E5E  CN=PKTestCert1                                  
        
        END    : Create self-signed certificate for code signing

.EXAMPLE
    PS C:\> New-PKCodeSigningCert -FriendlyName "Test cert 2" -Name PKTestCert2 -YearsValid 1 -Quiet

        [Subject]
          CN=PKTestCert2

        [Issuer]
          CN=PKTestCert2

        [Serial Number]
          772D76B4AF0F9FA7409A4D9474164A2C

        [Not Before]
          2019-05-08 05:33:05 PM

        [Not After]
          2020-05-08 05:43:00 PM

        [Thumbprint]
          A0E4DB3434EB775AC58B6F6187E7775768C49A46

.EXAMPLE 
    $Cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Out-GridView -Title 'Select Certificate' -OutputMode Single 
    # (Post-function task) Gets a cert in your personal store
    
    $Path = "C:\path\to\your\script.ps1"
    Set-AuthenticodeSignature -Certificate $Cert -FilePath $Path 
    # (Post-function task) Signs a single script

    Set-AuthenticodeSignature -Certificate $cert -TimestampServer http://timestamp.digicert.com -FilePath $Path  
    # (Post-function task) Signs a script with a timestamped signature
    
    Get-ChildItem -Path "$home\Documents" -Filter *.ps1 -Include *.ps1 -Recurse | Set-AuthenticodeSignature -Certificate $Cert 
    # (Post-function task) Signs multiple scripts

#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
param (
    [Parameter(
        Mandatory = $True,
        Position=0,
        HelpMessage = "Friendly name for certificate"
    )]
    [System.String]$FriendlyName,
        
    [Parameter(
        Mandatory = $True, 
        Position = 1,
        HelpMessage = "Certificate name (will become CN=Name)"
    )]
    [ValidateScript({If ($_ -notmatch "CE=") {$True}})]
    [System.String]$Name,

    [Parameter(
        HelpMessage = "Certificate validity in years (default 5)"
    )]
    [ValidateRange(1,10)]
    [int]$YearsValid = 5,

    [Parameter(
        HelpMessage = "Suppress all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [switch]$Quiet
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ScriptName",$ScriptName)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
    #region Verify OS version

    If (-not ($Null = [Environment]::OSVersion.Version -ge (new-object 'Version' 10,0))) {
        $Msg = "This script requires Windows 10 or Windows Server 2016 at minimum; you are running $((Get-CimInstance -ClassName Win32_OperatingSystem -Property Caption).Caption)"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }
    
    #endregion Verify OS version

    #region Functions
    
    # Function to write a console message in color, or write a verbose message if Quiet
    Function Write-MessageInfo {
        Param([Parameter(ValueFromPipeline)]$Message,$FGColor)
        $BGColor = $host.UI.RawUI.BackgroundColor
        If (-not $Quiet.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Message")}
        Else {Write-Verbose "$Message"}
    }

    # Function to write a console error message, or write a warning if Quiet
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        If (-not $Quiet.IsPresent) {$Host.UI.WriteErrorLine("$Message")}
        Else {Write-Warning "$Message"}
    }
    
    #endregion Functions

    #region Splats

    # Splat for Write-Progress
    $Activity = "Create self-signed certificate for code signing"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = "Create certificate $Name"
        Status           = $Env:ComputerName
    }

    # Splat for New-SelfSignedCertificate
    $Param_Cert = @{}
    $Param_Cert = @{
        KeyUsage          = "DigitalSignature"
        KeySpec           = "Signature"
        FriendlyName      = $FriendlyName
        Subject           = "CN=$Name"
        KeyExportPolicy   = "ExportableEncrypted"
        CertStoreLocation = "Cert:\CurrentUser\My"
        NotAfter          = (Get-Date).AddYears($YearsValid)
        TextExtension     = @('2.5.29.37={text}1.3.6.1.5.5.7.3.3')
        Confirm           = $False
        Verbose           = $False
        ErrorAction       = "Stop"
    }

    #endregion Splats

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "BEGIN  : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow  

}    

Process {
    
    Write-Progress @Param_WP

    Try {
        $Msg = "Test for existing self-signed certificate"
        "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor White

        If ($ExistingCert = Get-Childitem Cert:\CurrentUser\My -ErrorAction Stop | Where-Object {$_.Subject -eq "CN=$Name"}) {
            
            $Msg = "Certificate already exists; please remove it and run this script again"
            "[$Env:ComputerName] $Msg" | Write-MessageError
            $ExistingCert
        }
        Else {
            $Msg = "Create new self-signed certificate"
            "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor White

            $ConfirmMsg = "`n`n`tCreate new self-signed certificate for $Env:Username`n`n`tPath: Cert:\CurrentUser\My`n`tSubject: 'CN=$Name`n`tFriendly name: '$FriendlyName'`n`n"
            If ($PSCmdlet.ShouldProcess($Env:ComputerName,$ConfirmMsg)) {
                Try {
                    $NewCert = New-SelfSignedCertificate @Param_Cert
                    $Msg = "Successfully created self-signed certificate in Cert:\CurrentUser\My"
                    "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor Green
                    $NewCert.ToString()
                }
                Catch {
                    $Msg = "Failed to create certificate 'CN=$Name'"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
                    "[$Env:ComputerName] $Msg" | Write-MessageError
                }
            }
            Else {
                $Msg = "Operation cancelled by user"
                "[$Env:ComputerName] $Msg" | Write-MessageInfo -FGColor White
            }
        }
    }
    Catch {
        $Msg = "Failed to test for existing self-signed certificate"
        If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
        "[$Env:ComputerName] $Msg" | Write-MessageError
    }
  
}
End {
    
    Write-Progress -Activity * -Completed
    $Msg = "END    : $Activity"
    $Msg | Write-MessageInfo -FGColor Yellow  
}
}

<#

# To create a cert in personal store:
$Cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | 
    Out-GridView -Title 'Select Certificate' -OutputMode Single 

# To sign a single cert
$Path = "C:\path\to\your\script.ps1"
Set-AuthenticodeSignature -Certificate $Cert -FilePath $Path 


# To sign with a timestamped signature,
Set-AuthenticodeSignature -Certificate $cert -TimestampServer http://timestamp.digicert.com -FilePath $Path  

# To sign multiple certs
Get-ChildItem -Path "$home\Documents" -Filter *.ps1 -Include *.ps1 -Recurse |
    Set-AuthenticodeSignature -Certificate $cert 


#>