Function Test-PKConnection {
<#
SYNOPSIS
    Reports Windows disk information using Invoke-Command as a job, with option to wait for output

.DESCRIPTION
    Reports Windows disk information using Invoke-Command as a job, with option to wait for output
    Accepts pipeline input
    Returns a PSobject or PSJob

.NOTES        
    Name    : Function_Test-PKConnection.ps1
    Author  : Paula Kingsley
    Created : 2017-11-20
    Version : 01.00.0000
    History:

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - 2017-11-20 - Created script based on jrich's original
        
.LINK
    https://gallery.technet.microsoft.com/scriptcenter/Powershell-Test-Server-e0cdea9a?ranMID=24542&ranEAID=je6NUbpObpQ&ranSiteID=je6NUbpObpQ-1p7VW5KEESFnLgSaMed_Bw&tduid=(05cc9a118b47a445311a4d27bb47d63a)(256380)(2459594)(je6NUbpObpQ-1p7VW5KEESFnLgSaMed_Bw)()

.PARAMETER ComputerName
    Name of computer to check; separate multiple names with commas

.PARAMETER CredSSP
    Include CredSSP test (requires credential)

.PARAMETER Credential
    Valid credential on target

.PARAMETER Tests
    Tests to perform (Default is DNS, Ping, RDP, Registry, SSH, WinRM and WMI; ShowMenu allows for selection, including CredSSP)

.PARAMETER SuppressconsoleOutput
    Suppress all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Test-PKConnection server222 -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                   Value                                    
        ---                   -----                                    
        Verbose               True                                     
        ComputerName          {server222}                           
        Credential            System.Management.Automation.PSCredential
        Tests                 Default                                  
        SuppressConsoleOutput False                                    
        ScriptName            Test-PKConnection                        
        ScriptVersion         1.0.0                                    
        PipelineInput         False                                    

        Action: Test connectivity
        VERBOSE: ComputerName : server222
        VERBOSE: DNS          : 0.02
        VERBOSE: RDP          : 0.03
        VERBOSE: Ping         : 0.05
        VERBOSE: WinRM        : 0.02
        VERBOSE: Registry     : 12.44
        VERBOSE: WMI          : 0.89
        SSH test failed on server222
        VERBOSE: SSH          : 1.04
        VERBOSE: TOTAL        : 14.50

        VERBOSE: Testing for 1 computer(s) completed in 14.50s


        ComputerName   : server222
        IP             : 10.11.12.13
        Domain         : server222.domain.local
        DNS            : True
        Ping           : True
        WSMAN          : True
        CredSSP        : -
        RemoteReg      : True
        RPC            : True
        RDP            : True
        SSH            : False
        ElapsedSeconds : 14.49
        Messages       : 1 connection test(s) failed

.EXAMPLE
    PS C:\>$ Get-VM ops*  | Test-PKConnection | Format-Table -AutoSize

        Action: Perform 7 network connectivity test(s)
        RDP test failed on ops-monitor-1
        WinRM test failed on ops-monitor-1
        Registry test failed on ops-monitor-1
        WMI test failed on ops-monitor-1
        RDP test failed on ops-ftpeval-fw-1
        WinRM test failed on ops-ftpeval-fw-1
        Registry test failed on ops-ftpeval-fw-1
        WMI test failed on ops-ftpeval-fw-1
        RDP test failed on ops-device42-2
        WinRM test failed on ops-device42-2
        Registry test failed on ops-device42-2
        WMI test failed on ops-device42-2
        SSH test failed on ops-device42-2
        SSH test failed on ops-pdamweb-1
        SSH test failed on ops-sgwin-1
        RDP test failed on ops-jenkins-dev-1
        WinRM test failed on ops-jenkins-dev-1
        Registry test failed on ops-jenkins-dev-1
        WMI test failed on ops-jenkins-dev-1
        SSH test failed on ops-windev-1


        ComputerName      IP            Domain                                DNS Ping WSMAN CredSSP RemoteReg   RPC   RDP
        ------------      --            ------                                --- ---- ----- ------- ---------   ---   ---
        ops-monitor-1     10.11.178.29  globix-sc.gracenote.com              True True False -           False False False
        ops-ftpeval-fw-1  10.11.86.17   globix-sc.gracenote.com              True True False -           False False False
        ops-device42-2    10.11.128.204 globix-sc.gracenote.com              True True False -           False False False
        ops-pdamweb-1     10.11.178.103 gracenote.gracenote.com              True True  True -            True  True  True
        ops-sgwin-1       10.11.178.67  gracenote.gracenote.com              True True  True -            True  True  True
        ops-jenkins-dev-1 10.11.178.119 globix-sc.gracenote.com              True True False -           False False False
        ops-windev-1      10.11.178.120 OPS-WINDEV-1.gracenote.gracenote.com True True  True -            True  True  True

#>
[cmdletBinding()]
param(
	[parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage = "Computer name(s)"
    )]
    [Alias("Name","DNSHostName","FQDN","HostName")]
    [ValidateNotNullOrEmpty()]
	[string[]]$ComputerName,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Credentials on target computer"
    )]
	
    [pscredential] $Credential = [System.Management.Automation.PSCredential]::Empty ,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Tests to perform (Default or ShowMenu)"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Default","ShowMenu")]
	[string]$Tests = "Default",

	##[parameter(
    #    Mandatory=$False
    #)]
	#[switch]$IncludeCredSSP,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Hide all non-verbose/non-error console output"
    )]
    [Switch] $SuppressConsoleOutput

)    
	
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Detect pipeline input
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("ComputerName")) -and (-not $ComputerName)

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference    = "Continue"
    
    # For time tally, output object
    $ScriptStartTime = Get-Date
	
    #region parameter cleanup & selection

    # Can't use implicit credentials with CredSSP
	If ($IncludeCredSSP.IsPresent -and (-not $Credential)){
        $Msg = "Credentials must be provided when testing CredSSP"
        $Host.UI.WriteErrorLine("ERROR: $MSg")
        Break
    }
    
    [array]$Menu = "CredSSP","DNS","Ping","RDP","Registry","SSH","WinRM","WMI"
    If ($Tests -eq "ShowMenu") {
        If (([array]$Selection = ($Menu | Sort | Out-GridView -Title "Select one or more tests (DNS test is mandatory and will be added if not selected)" -OutputMode Multiple)).Count -eq 0){
            $Msg = "At least one selection is mandatory"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
        Else {
            If ($Selection -notcontains "DNS") {$Selection += "DNS"}
            $Msg = "$($Selection.Count) test(s) selected"
            Write-Verbose $Msg
            $Msg = $Selection -join(", ")
            Write-Verbose $Msg
        }
    }
    Else {
        [array]$Selection = $Menu | Where-Object {$_ -ne "CredSSP"}
    }

    #endregion Parameter cleanup & selection

    #region splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Splat for Write-Progress
    $Activity = "Perform $($Selection.Count) network connectivity test(s)"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        CurrentOperation = $Null
        Status           = "Working"
        PercentComplete  = $Null
    }
    
    #endregion Splats
        
    #region functions

    # Mini function to avoid DNS lookup errors
    function LookupDNS {
        Param($Name)
        Try {
            If ($Lookup = [Net.Dns]::GetHostEntry($Name)) {
                Write-Output $Lookup
            }
        }
        Catch {}
    }

    #endregion functions

    #region  Output
    $InitialValue = "Error"
    $OutputTemplate = New-Object PSObject -Property ([ordered]@{
        ComputerName   = $InitialValue
        IP             = $InitialValue
        Domain         = $InitialValue
        DNS            = $InitialValue
        Ping           = $InitialValue
        WSMAN          = $InitialValue
        CredSSP        = $InitialValue
        RemoteReg      = $InitialValue
        RPC            = $InitialValue
        RDP            = $InitialValue
        SSH            = $InitialValue
        ElapsedSeconds = $InitialValue
        Messages       = $InitialValue
    })

    $Results = @()

    # To line up verbose output 
    $LongestLabel = "ComputerName"
    $Length = $LongestLabel.Length
    $Pad = $Length + 1

    #endregion Output
    
    # Console output
    $BGColor = $Host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg}

}
process{
    
    $Total = $ComputerName.Count
    [int]$Current = 0

    Foreach ($Computer in $ComputerName){
        
        $Current ++
        $Param_WP.CurrentOperation = $Computer
        $Param_WP.PercentComplete = ($Current / $Total * 100)

        $Output = $OutputTemplate.PSObject.Copy()
        $Output.ComputerName = $Computer

        $Failures = 0
        $TimeTally = 0

        $ComputerStartTime = Get-Date
	    $DT = $CDT = Get-Date

        $Label = "ComputerName"
        $Msg = "$($Label.PadRight($Pad, ' ')): $Computer"
        Write-Verbose $Msg

        $failed = 0

        $Msg = "Look up IP in DNS"
        $Param_WP.Status = $Msg
        Write-Progress @Param_WP

        [switch]$Continue = $False

        # Lookup in DNS
	    Try{
	        If ($DNSEntity = LookupDNS -Name $Computer) {
	            $Domain = ($DNSEntity.hostname).replace("$Computer.","")
	            [string[]]$IPs = $DNSEntity.AddressList | Foreach-Object {$_.IPAddressToString}
                $Output.Domain = $Domain

                $Label = "DNS"
                $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                $TimeTally += $Time
                $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	            Write-verbose $Msg
            
                $Output.DNS = $True

                $Continue = $True
            }
            Else {
                $Msg = "DNS lookup failed for $Computer"
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                $Output.Messages = $Msg		        

                $Output.Ping = $Output.WSMAN = $Output.RDP = $Output.RemoteReg = $Output.RPC = $False
                If ($Output.CredSSP -ne "-") {$Output.CredSSP = $False}

                $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	            Write-verbose $Msg

                $OutputTotal = $Time
                $Output.DNS = $False
                
		    
                $Results += $Output
            }
	    }
	    Catch {
		    $Msg = "DNS lookup failed for $Computer"
            If ($ErrorDetails = $_.Exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            $Output.Messages = $Msg
		    
            $Output.Ping = $Output.WSMAN = $Output.RDP = $Output.RemoteReg = $Output.RPC = $False
            If ($Output.CredSSP -ne "-") {$Output.CredSSP = $False}

            $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
            $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	        Write-verbose $Msg

            $OutputTotal = $Time
            $Output.DNS = $False
		    
            $Results += $Output
	    }
        
        
	    If ($Continue.IsPresent) {
    	    
            foreach ($IP in $IPs) {

                $Output1 = $Output.PSObject.Copy()	            
                $Output1.IP = $IP

                $Label = "Ping"
                If ($Selection -contains $Label) {
                    
                    $Msg = "Test $Label"
                    $Param_WP.Status = $Msg
                    Write-Progress @Param_WP

                    Try {
	                    If (Test-Connection $IP -count 1 -Quiet) {
			                $Output1.Ping = $True
                        }
			            Else {
                            $Output1.Ping = $False
                            $Msg = "$Label test failed on $Computer"
                            $Host.UI.WriteErrorLine($Msg)
                            $Failures ++
                        }
                    }
                    Catch {
                        $Output1.Ping = $False
                        $Msg = "$Label test failed on $Computer"
                        $Host.UI.WriteErrorLine($Msg)
                        $Failures ++
                    }
                    
                    $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                    $TimeTally += $Time
                    $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	                Write-verbose $Msg
                }
                Else {
                    $Output1.Ping = "-"
                }

                $Label = "RDP"
                If ($Selection -contains $Label) {
                
                    $Msg = "Test $Label"
                    $Param_WP.Status = $Msg
                    Write-Progress @Param_WP
		            Try {
                        $Socket = New-Object Net.Sockets.TcpClient($IP, 3389)
		                If ($Socket -eq $null) {
			                $Output1.RDP = $False
                            $Msg = "$Label test failed on $Computer"
                            $Host.UI.WriteErrorLine($Msg)
                            $Failures ++
		                }
		                else {
			                $Output1.RDP = $True
			                $socket.close()
		                }
                    }
                    Catch {
                        $Output1.RDP = $False
                        $Msg = "$Label test failed on $Computer"
                        $Host.UI.WriteErrorLine($Msg)
                        $Failures ++
                    }
		        
                    $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                    $TimeTally += $Time
                    $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	                Write-verbose $Msg
                }
                Else {
                    $Output1.RDP = "-"
                }

                
                $Label = "WinRM"
                If ($Selection -contains $Label) {

                    $Msg = "Test $Label"
                    $Param_WP.Status = $Msg
                    Write-Progress @Param_WP

                    Try {
                        #Test-WSMan $IP -Credential $Credential | Out-Null				        
                        Test-WSMan $IP | Out-Null
				        $Output1.WSMAN = $True
				    }
			        Catch {
                        $Output1.WSMAN = $False
                        $Msg = "$Label test failed on $Computer"
                        $Host.UI.WriteErrorLine($Msg)
                        $Failures ++
                    }

                    $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                    $TimeTally += $Time
                    $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	                Write-verbose $Msg
                }
                Else {
                    $Output1.WSMAN = "-"
                }
                
                $Label = "CredSSP"
                If ($Selection -contains $Label) {
                    
                    $Msg = "Test $Label"
                    $Param_WP.Status = $Msg
                    Write-Progress @Param_WP				    
                        
                    Try {
					    Test-WSMan $ip -Authentication Credssp -Credential $Credential
					    $Output1.CredSSP = $True
					}
				    Catch {
                        $Output1.CredSSP = $False
                        $Msg = "$Label test failed on $Computer"
                        $Host.UI.WriteErrorLine($Msg)
                        $Failures ++
                    }
                
                    $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                    $TimeTally += $Time
                    $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	                Write-verbose $Msg
			    }
                Else {
                    $Output1.CredSSP = "-"
                }
                
                $Label = "Registry"   
                If ($Selection -contains $Label) {    

                    $Msg = "Test $Label"
                    $Param_WP.Status = $Msg
                    Write-Progress @Param_WP
                    Try {
				        [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $IP) | Out-Null
				        $Output1.RemoteReg = $True
			        }
			        Catch {
                        $Output1.RemoteReg = $False
                        $Msg = "$Label test failed on $Computer"
                        $Host.UI.WriteErrorLine($Msg)
                        $Failures ++
                    }
		            	        
                    $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                    $TimeTally += $Time
                    $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	                Write-verbose $Msg
                }
                Else {
                    $Output1.RemoteReg = "-"
                }
                
                $Label = "WMI"
                If ($Selection -contains $Label) {                    
                    $Msg = "Test $Label"
                    $Param_WP.Status = $Msg
                    Write-Progress @Param_WP
                    Try {	
				        $w = [wmi] ''
				        #$w.psbase.options.timeout = 15000000
                        $w.psbase.options.timeout = 15000
				        $w.path = "\\$Computer\root\cimv2:Win32_ComputerSystem.Name='$Computer'"
				        $w | select none | Out-Null
				        $Output1.RPC = $True
			        }
			        Catch {
                        $Output1.RPC = $False
                        $Msg = "$Label test failed on $Computer"
                        $Host.UI.WriteErrorLine($Msg)
                        $Failures ++
                    }

                    $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                    $TimeTally += $Time
                    $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	                Write-verbose $Msg
                }
                Else {
                    $Output1.RPC = "-"
                }

                $Label = "SSH"
                If ($Selection -contains $Label) {
                    
                    $Msg = "Test $Label"
                    $Param_WP.Status = $Msg
                    Write-Progress @Param_WP
                    Try {
                        $Test = New-Object System.Net.Sockets.TcpClient
                        If ($Test.Connect($Computer, 22) -eq $Null) {
                            $Output1.SSH = $True
                        }
                        Else {
                            $Output1.SSH = $False
                            $Msg = "$Label test failed on $Computer"
                            $Host.UI.WriteErrorLine($Msg)
                            $Failures ++
                        }
			        }
			        Catch {
                        $Output1.SSH = $False
                        $Msg = "$Label test failed on $Computer"
                        $Host.UI.WriteErrorLine($Msg)
                        $Failures ++
                    }
		     
                    $Time = "{0:N2}" -f $((New-TimeSpan $DT ($DT = get-date)).totalseconds)
                    $TimeTally += $Time
                    $Msg = "$($Label.PadRight($Pad, ' ')): $Time"
	                Write-verbose $Msg
                }
                Else {
                    $Output1.SSH = "-"
                }

                # Total time
                $Time = "{0:N2}" -f $((New-TimeSpan $ScriptStartTime ($DT)).totalseconds)
                $Label = "TOTAL"
                $Msg = "$($Label.PadRight($Pad, ' ')): $Time`n"
        
                $Output1.ElapsedSeconds = "{0:N2}" -f $TimeTally

                If ($Failures.Count -gt 0) {
                    $Output1.Messages = "$($Failures.Count) connection test(s) failed"
                }
                Else {
                    $Output1.Messages = "All tests completed successfully"
                }

                # Add them up
                $Results += $Output1
            
            } #end foreach IP
    
        } #end If DNS lookup successful

        
        Write-Verbose $Msg
        
    } #end foreach computer
}
end{
    
    Write-Progress -Activity $Activity -Completed

    $Host.UI.WriteLine()

    $TotalSec = "{0:N2}" -f $((New-TimeSpan $ScriptStartTime ($DT)).totalseconds)
    $Msg = "Testing for $Total computer(s) completed in $TotalSec`s"
	Write-Verbose $Msg

    Write-Output $Results 

}
} #end Test-PKConnection