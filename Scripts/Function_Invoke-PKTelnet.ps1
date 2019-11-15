#requires -Version 3
Function Invoke-PKTelnet {
<#
.SYNOPSIS
    Uses System.Net.Sockets.TcpClient to test telnet connectivity on a specified port, to one or more computers, with a timeout

.DESCRIPTION
    Uses System.Net.Sockets.TcpClient to test telnet connectivity on a specified port, to one or more computers, with a timeout
    Created because the telnet-client is, irritatingly, usually disabled
    Accepts pipeline input
    Returns a PSObject or a Boolean

.NOTES        
    Name    : Function_Invoke-PKTelnet .ps1
    Created : 2019-11-14
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2019-11-14 - Created script based on link

.LINK
    https://www.techtutsonline.com/powershell-alternative-telnet-command/

.PARAMETER ComputerName
    One or more computer names (default is 'localhost')

.PARAMETER Port
    Port to test (e.g., 22, 3389, 80)

.PARAMETER Timeout 
    Connection timeout in milliseconds (default is 1000)

.PARAMETER BooleanOutput
    Return True/False only (default is psobject)

.PARAMETER Quiet
    Hide all non-verbose console output

.EXAMPLE
    PS C:\> Invoke-PKTelnet -Port 47001 -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                     Value
        ---                     -----
        Port                    47001
        Verbose                  True
        ComputerName        Localhost
        Timeout                 10000
        Quiet                   False
        ScriptName    Invoke-PKTelnet
        ScriptVersion           1.0.0
        PipelineInput           False

        BEGIN: Test telnet on port 47001 with 10000-millisecond timeout 

        [Localhost] Connection successful


        ComputerName : Localhost
        IsSuccessful : True
        Port         : 47001
        Timeout      : 10000
        Messages     : Connection successful


        END  : Test telnet on port 47001 with 10000-millisecond timeout 

.EXAMPLE
    PS C:\> $Arr | Invoke-PKTelnet -port 3389 -Timeout 500 -Quiet | Format-Table -Autosize

        ComputerName             IsSuccessful Port Timeout Messages             
        ------------             ------------ ---- ------- --------             
        bastion12.domain.local   True         3389     500 Connection successful
        legacysqlbackup          False        3389     500 Connection failed    
        dc04                     True         3389     500 Connection successful

.EXAMPLE
    PS C:\> Invoke-PKTelnet -Computername 8.8.8.8 -Port 53 -BooleanOutput -Quiet
        True


#>
[CmdletBinding()]
param(
    [Parameter(
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "One or more computer names (default is 'localhost')"
    )]
    [Alias ('HostName','cn','Host','Computer','DNSHostname','FQDN','Name')]
    [String[]]$ComputerName,

    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Port to test (e.g., 22, 3389, 80)"
    )]
    [int32]$Port,

    [Parameter(
        HelpMessage = "Connection timeout in milliseconds (default is 1000)"
    )]
    [int32] $Timeout = 10000,

    [Parameter(
        HelpMessage = "Return True/False only (default is psobject)"
    )]
    [Switch] $BooleanOutput,

    [Parameter(
        HelpMessage = "Hide all non-verbose console output"
    )]
    [Alias("SuppressConsoleOutput")]
    [Switch] $Quiet

)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    If (-not $PipelineInput.IsPresent -and -not $CurrentParams.ComputerName) {
        $ComputerName = $CurrentParams.ComputerName = 'Localhost'
    }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    $Activity = "Test telnet on port $Port with $Timeout`-millisecond timeout "

    #region Functions

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

    # Function to write an error message (as a string with no stacktrace info)
    Function Write-MessageError {
        [CmdletBinding()]
        Param([Parameter(ValueFromPipeline)]$Message)
        $Host.UI.WriteErrorLine("$Message")
    }

    #endregion Functions

    # Console output
    "BEGIN: $Activity" | Write-MessageInfo -FGColor Yellow -Title

}

Process {
    
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        #$Msg = "Attempting to connect on port $Port"
        #Write-Verbose "[$Computer] $Msg"
        Write-Progress -Activity $Activity -CurrentOperation $Computer -PercentComplete ($Current/$Total*100) 

        $Output = [PSCustomObject]@{
            ComputerName = $Computer
            IsSuccessful = $Null
            Port         = $Port
            Timeout      = $Timeout
            Messages     = $Null
        }

        Try {
            $TCP = New-Object System.Net.Sockets.TcpClient
            $Connection = $TCP.BeginConnect($Computer, $Port, $Null, $Null)
            $Connection.AsyncWaitHandle.WaitOne($Timeout,$False)  | Out-Null 
            
            If ($TCP.Connected -eq $true) {
                $Msg = "Connection successful"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Green
                $Output.IsSuccessful = $True
                $Output.Messages = $Msg
            }
            Else {
                $Msg = "Connection failed"
                "[$Computer] $Msg" | Write-MessageInfo -FGColor Red
                $Output.IsSuccessful = $False
                $Output.Messages = $Msg
            }
        }
        Catch {
            $Msg = "Operation failed"
            If ($ErrorDetails = $_.Exception.Message) {$Msg += " ($ErrorDetails)"}
            "[$Computer] $Msg" | Write-MessageError
            $Output.IsSuccessful = $False
            $Output.Messages = $Msg
        }

        If ($BooleanOutput.IsPresent) {
            Write-Output $Output.IsSuccessful
        }
        Else {
            Write-Output $Output
        }
    }
}
End {
    
    "END  : $Activity" | Write-MessageInfo -FGColor Yellow -Title
    Write-Progress -Activity $Activity -Completed
}
} #end Invoke-PKTelnet

$Null = New-Alias -Name Test-PKTelnet -Value Invoke-PKTelnet -Description "For guessability" -Force -Confirm:$False


