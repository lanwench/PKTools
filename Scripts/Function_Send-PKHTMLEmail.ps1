#Requires -Version 3.0
function Send-PKHTMLEmail {
<#
.SYNOPSIS
    Send a nicely-formatted HTML email from an existing PSObject

.DESCRIPTION
    Send a nicely-formatted HTML email from an existing PSObject

.NOTES
    Name    : Function_Send-PKHTMLEmail.ps1
    Version : 01.00.0000
    Created : 2022-04-11
    Author  : Paula Kingsley
    History : 
        
        (please keep up to date with $Version in Begin block!)

        v01.00.0000 - 2022-04-11 - Created script


.PARAMETER InputObject
    PSCustomObject or PSObject

.PARAMETER AsList
    Convert PSObject to list/table format

.PARAMETER Subject
    Text for subject line

.PARAMETER To
    One or more recipient SMTP addresses, separated by commas

.PARAMETER CC
    One or more carbon copy SMTP addresses, separated by commas

.PARAMETER From
    SMTP address of the sender (defaults to current user, if available)

.PARAMETER SMTPServer
    Hostname, FQDN or IP address of the SMTP relay server 

.PARAMETER Port
    TCP port for SMTP (default is 25)

.PARAMETER UseSSL
    Use SSL for SMTP server

.PARAMETER CSS
    Path to Cascading Style Sheet file for formatting table (optional; basic internal CSS provided in here-string)

.PARAMETER TestSMTPServer
    Test TCP connectivity prior to sending message

.EXAMPLE
    PS C:\> Send-PKHTMLEmail -InputObject $Output -Subject "I love kittens" -To jane.bloggs@domain.com -SMTPServer relay.domain.com -Verbose -TestSMTPServer -UseSSL

        VERBOSE: PSBoundParameters: 
	
        Key            Value                                                                                                                                                                                 
        ---            -----                                                                                                                                                                                 
        InputObject    @{Pets=cats; Color=blue; Numbers=4, 5, 6...}
        Subject        I love kittens
        To             {jane.bloggs@domain.com}                                                                                                                                                          
        SMTPServer     relay.domain.com                                                                                                                                                                       
        Verbose        True                                                                                                                                                                                  
        TestSMTPServer True                                                                                                                                                                                  
        UseSSL         True                                                                                                                                                                                  
        AsList         False                                                                                                                                                                                 
        From           joe.bloggs@domain.com                                                                                                                                                            
        CC                                                                                                                                                                                                   
        Port           25                                                                                                                                                                                    
        Credential                                                                                                                                                                                           
        CSS                                                                                                                                                                                                  
        ScriptName     Send-PKHTMLEmail                                                                                                                                                                      
        ScriptVersion  1.0.0                                                                                                                                                                                 

        VERBOSE: Testing connectivity to SMTP server 'relay.domain.com' on TCP port 25
        VERBOSE: Connection to relay.domain.com (10.24.40.246) on port 25 successful
        VERBOSE: Create message body; convert input object to HTML & replace line break literals with line break tabs
        VERBOSE: Confirm email
        VERBOSE: Message submitted to SMTP server


.OUTPUTS
    None
#> 
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
 Param 
   ([Parameter(
        Mandatory=$True,
        Position = 0,
        ValueFromPipeline = $true,
        HelpMessage = "PSCustomObject or PSObject"
    )]
    [ValidateNotNullOrEmpty()]
    [object]$InputObject,
    
        [Parameter(
        Mandatory = $False,
        HelpMessage = "Convert PSObject to list/table format "
    )]    
    [switch]$AsList,

    [Parameter(
        Mandatory = $True,
        HelpMessage = "Text for subject line"
    )]
    [ValidateNotNullOrEmpty()]
    [String]$Subject,    

    [Parameter(
        Mandatory = $True,
        HelpMessage = "One or more recipient SMTP addresses, separated by commas"
    )]    
    [ValidateNotNullOrEmpty()]
    [String[]]$To,

    [Parameter(
        HelpMessage = "Sender SMTP address (defaults to your own, if available)"
    )]
    [ValidateNotNullOrEmpty()]
    [String]$From = $(([adsi]"LDAP://$(whoami /fqdn)").mail),
    
    [Parameter(
        HelpMessage = "One or more CC addresses, separated by commas"
    )]    
    [ValidateNotNullOrEmpty()]
    [String[]]$CC,

    [Parameter(
        Mandatory = $True,
        HelpMessage = "SMTP server name or FQDN"
    )]    
    [ValidateNotNullOrEmpty()]
    [String]$SMTPServer,

    [Parameter(
        HelpMessage = "TCP port for SMTP (default is 25)"
    )]    
    [ValidateNotNullOrEmpty()]
    [int]$Port = 25,

    [Parameter(
        HelpMessage = "Use SSL for SMTP server"
    )]    
    [ValidateNotNullOrEmpty()]
    [switch]$UseSSL,

    [Parameter(
        HelpMessage = "Sender credentials for SMTP server (if required)"
    )]    
    [PSCredential]$Credential,

    [Parameter(
        HelpMessage = "Path to Cascading Style Sheet file for formatting table (optional; basic internal CSS provided in here-string)"
    )]
    [ValidateNotNullOrEmpty()]    
    [String]$CSS,

    [Parameter(
        HelpMessage = "Test TCP connectivity prior to sending message"
    )]    
    [switch]$TestSMTPServer

)
Begin {
    
    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # CSS
$Head = @"
<style type="text/css">
    table {
    	font-family: Verdana;	    
        padding: 5px;
        border-style: solid;
    	border-width: 3px;
        border-collapse: collapse;
    	background-color: #eef3f6; 
    	table-layout: auto;
    	text-align: left;
    	font-size: 10pt;
    }
    table th {
        border-width: 1px;
        border-style: solid;
        border-color: black;
        border-collapse: collapse;
        padding: 5px;
        background-color: #578ca9;
        text-align: left;
        font: bold
    }
    table td {
        border-width: 1px;
        padding: 5px;
        border-style: solid;
        border-color: black;
        border-collapse: collapse;
        text-align: left;
    }
    .style1 {
        font-family: Calibri;
        font-weight:bold;
        font-size:small;
    }
    </style>
"@

# Yellow:
#FFFFCC
 
    
    # Splat for ConvertTo-HTML
    $HTMLDetails = @{}    
    $HTMLDetails = @{
        Title      = $Subject
        Head       = $Head
        As         = &{If ($AsList.IsPresent) {"List"} Else {"Table"}}
        PreContent = "<P>Generated on $(Get-Date)</P><br>"
        PostContent = "<br>Message generated by $(whoami) on $Env:ComputerName"
    }
    If ($CurrentParams.CSS) {
        $HTMLDetails.CSSURI = $CSS
    }

    # Splat for Send-MailMessage
    $Param_SendMail = @{}
    $Param_SendMail = @{
        To          = $To
        Body        = $Null
        Subject     = $Subject
        SmtpServer  = $SMTPServer
        From        = $From
        Port        = $Port
        UseSSL      = $UseSSL
        BodyAsHtml  = $True
        ErrorAction = "Stop"
        Verbose     = $False
    }
    If ($CurrentParams.CC) {
        $Param_SendMail.add("CC",$CC)
    }
    If ($CurrentParams.Credential) {
        $Param_SendMail.add("Credential",$Credential)
    }

    If ($TestSMTPServer.IsPresent) {
        $Msg = "Testing connectivity to SMTP server '$SMTPServer' on TCP port $Port"
        Write-Verbose $Msg
        Try {
            $Test = Test-NetConnection -ComputerName $SMTPServer -Port $Port -InformationLevel Detailed -ErrorAction Stop
            If ($Test.TCPTestSucceeded) {
                $Msg = "Connection to $SMTPServer ($($Test.RemoteAddress)) on port $($Test.RemotePort) successful"
                Write-Verbose $Msg
            }
            Else {
                $Msg = "Connection failed"
                If ($ErrorDetails = $_.ExceptionMessage) {$Msg += " ($ErrorDetails)"}
                Write-Error $Msg
                break
            }
        }
        Catch {
            Throw $_.ExceptionMessage
        }
    }

}
Process {

    $Msg = "Create message body; convert input object to HTML & replace line break literals with line break tabs"
    Write-Verbose $Msg
    
    #If ($Aslist.IsPresent) {
    #    $Body = "$(($InputObject | ConvertTo-Html -as list @HTMLDetails) -replace '&lt;br&gt;', '<br>')"
    #}
    #Else {
        $Body = "$(($InputObject | ConvertTo-Html @HTMLDetails) -replace '&lt;br&gt;', '<br>')"
    #}
    $Param_SendMail.Body = $Body

    $Msg =  "Confirm email"
    Write-Verbose $Msg
    
    If ($PSCmdlet.ShouldProcess($Null,"$($Msg)`n$(New-Object PSObject -Property $Param_SendMail | 
        Select To,From,Subject,SMTPServer,BodyAsHTML,@{N="MessageBody";E={ if ($_.Body.length -gt 500) { "$($_.Body.substring(0, 500))`n`n<snip>" } else { $_.Body }}} | Out-String)")) {
        Try {
            Send-MailMessage @Param_SendMail
            $Msg = "Message submitted to SMTP server"
            Write-Verbose $Msg
        }
        catch {
            $Msg = "Operation failed"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
            Write-Warning $Msg
        }
    }
    Else {
        $Msg = "Operation cancelled by user"
        Write-Verbose $Msg
    }

} #end process
End {}
    
}#Send-PKHTMLEmail

