#requires -Version 3
Function Install-PKChocoPackage {
<#
.Synopsis
    Uses Invoke-Command to install Chocolatey packages on remote computers as PSJobs

.DESCRIPTION
    Uses Invoke-Command to install Chocolatey packages on remote computers as PSJobs
    Accepts pipeline input
    Returns a PSJob object
    

.NOTES
    Name    : Function_Install-PKChocoPackage.ps1
    Created : 2017-06-27
    Version : 1.0.0
    Author  : Paula Kingsley
    History :

        # PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK

        v1.0.0 - 2017-06-27 - Created script

.EXAMPLE
    PS C:\> Install-PKChocoPackage -ComputerName ops-pktest-1 -Package sysinternals -Verbose
        VERBOSE: PSBoundParameters: 
	
        Key              Value                                    
        ---              -----                                    
        ComputerName     {ops-pktest-1}                           
        Package          sysinternals                             
        Verbose          True                                     
        Credential       System.Management.Automation.PSCredential
        Force            False                                    
        JobName          Choco                                    
        ScriptName       Install-PKChocoPackage                   
        ScriptVersion    1.0.0                                                        

        VERBOSE: ops-pktest-1

        Id     Name            PSJobTypeName   State         HasMoreData     Location             Command                  
        --     ----            -------------   -----         -----------     --------             -------                  
        30     Choco           RemoteJob       Running       True            ops-pktest-1         ...                      


        PS C:\> Get-Job 30 | Receive-Job 

        ComputerName     : OPS-PKTEST-1
        Package          : sysinternals
        AlreadyInstalled : False
        Installed        : True
        Messages         : Package installed successfully
        Output           : {Chocolatey v0.10.3, Installing the following packages:, sysinternals, By installing you accept licenses for the 
                           packages....}
        PSComputerName   : ops-pktest-1
        RunspaceId       : 64263ddd-1156-4c26-8544-53c3e9dd2586

.EXAMPLE
    PS C:\> (Get-ADComputer ops-pktest-1).DNSHostName | 
        Install-PKChocoPackage -Package nodejs.install -Credential $Credential -JobName "MyJob" -Verbose | Get-Job | Wait-Job | Receive-Job
        
        VERBOSE: PSBoundParameters: 
	
        Key           Value                                    
        ---           -----                                    
        Package       nodejs.install                           
        Credential    System.Management.Automation.PSCredential
        JobName       MyJob                                    
        Verbose       True                                     
        ComputerName                                           
        Force         False                                    
        ScriptName    Install-PKChocoPackage                   
        ScriptVersion 1.0.0                                    

        VERBOSE: ops-pktest-1.domain.local

        ComputerName     : OPS-PKTEST-1
        Package          : nodejs.install
        AlreadyInstalled : False
        Installed        : True
        Messages         : Package installed successfully
        Output           : {Chocolatey v0.10.3, Installing the following packages:, nodejs.install, By installing you accept licenses for the 
                           packages....}
        PSComputerName   : ops-pktest-1.domain.local
        RunspaceId       : b38f981a-1044-4fd2-97b7-d2464054526d

#>
[CmdletBinding(
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [parameter(
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage = "Name of target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("NodeName","Hostname","Name","FQDN")]
    [string[]]$ComputerName,

    [parameter(
        Mandatory=$True,
        HelpMessage = "Package name"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Package,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Valid credentials on target computer"
    )]
    [ValidateNotNullOrEmpty()]
    [PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Force install"
    )]
    [ValidateNotNullOrEmpty()]
    [switch]$Force,

    [parameter(
        Mandatory=$False,
        HelpMessage = "Name of job"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$JobName = "Choco"

)

Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "1.0.0"

    # Generalpurpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Scriptblock
    $SB = {
        $Output = New-Object PSObject -Property @{
            ComputerName     = $Env:ComputerName
            Package          = $Using:Package
            AlreadyInstalled = "Error"
            Installed        = "Error"
            Messages         = "Error"
            Output           = "Error"
        }
        $Select = "ComputerName","Package","AlreadyInstalled","Installed","Messages","Output"

        #Make sure chocolatey is installed 
        Try {
            $Null = Get-Command choco.exe -ErrorAction Stop

            [switch]$F = $Using:Force
            $Command = "choco install $Using:Package -y"
            If ($F.IsPresent) {
                $Command = "$Command --force" 
                $Output.AlreadyInstalled = "[-Force specified]"
            }
            Try {
                $Install = Invoke-Expression -Command $Command -ErrorAction Stop
        
                $Output.Output = $Install
        
                If ($Install | Select-String -Pattern "The install of $Using:Package was successful") {
                    $Output.AlreadyInstalled = $False
                    $Output.Installed = $True
                    $Output.Messages = "Package installed successfully"
                }
                Elseif ($Install | Select-String -Pattern 'already installed') {
                    $Output.AlreadyInstalled = $True
                    $Output.Installed = "(n/a)"
                    $Output.Message = "Package already installed"
                }
                Elseif ($Install | Select-String -Pattern 'The package was not found') {
                    $Output.AlreadyInstalled = "(n/a)"
                    $Output.Installed = $False
                    $Output.Messages = "Package not found"
                }
                Else {
                    $Output.Installed = $False
                    $Output.Messages = "Package failed to install"
                }
            }
            Catch {
                $Output.Messages = $_.Exception.Message
            }
        }
        Catch {
            $Output.Messages = "Chocolatey not found; see https://chocolatey.org"
        }
        Write-Output ($Output | Select-Object $Select)
        
    }

    # Splat
    $Param_IC = @{
        ComputerName = $Null
        Credential   = $Credential
        ScriptBlock  = $SB 
        ArgumentList = $Package,$Force
        AsJob        = $True
        JobName      = $JobName
        ErrorAction  = "Stop"
        Verbose      = $False
    }
    
    $Activity = "Install Chocolatey package"

}
Process {
    
    $Total = $ComputerName.Count
    $Current = 0
    
    Foreach ($computer in $computerName) {
        $Current ++
        $Msg = $Computer
        Write-Verbose $Msg

        Write-Progress -Activity $Activity -CurrentOperation $Current -PercentComplete ($Current /$Total * 100)
        
        $Msg = "Install package '$Sysinternals'"
        If ($PSCmdlet.ShouldProcess($Computer,$Msg)) {

            Try {
                $Param_IC.ComputerName = $Computer
                Invoke-Command @Param_IC
            }
            Catch {
                $Msg = "Job invocation failed on '$Computer'"
                $ErrorDetails = $_.Exception.Message
                $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
            }
        }
        Else {
            $Msg = "Job execution on $Computer cancelled by user"
            Write-Warning $Msg
        }
    } # end foreach 
}
End {
    Write-Progress -Activity $Activity -Completed
}

} #end Install-PKChocoPackage

