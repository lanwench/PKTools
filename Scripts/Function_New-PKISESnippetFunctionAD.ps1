#requires -Version 3
Function New-PKISESnippetFunctionAD {
<#
.SYNOPSIS
    Adds a new PS ISE snippet containing a template function using the ActiveDirectory module

.DESCRIPTION
    Adds a new PS ISE snippet containing a template function using the ActiveDirectory module
    SupportsShouldProcess
    Returns a file object

.NOTES
    Name    : Function_New-PKISESnippetFunctionAD.ps1
    Created : 2018-10-19
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UPDATED IN BEGIN BLOCK **

        v01.00.0000 - 2018-10-19 - Created script

.PARAMETER Author
    Author name

.PARAMETER AutoDetectAuthorFullName
    Attempt to match the current username to their full name via the registry & WMI

.PARAMETER Force
    Forces creation even if snippet name exists

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionAD -AutoDetectAuthorFullName -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                             
        ---                      -----                             
        AutoDetectAuthorFullName True                              
        Verbose                  True                              
        Author                                                     
        Force                    False                             
        Confirm                  True
        ScriptName               New-PKISESnippetFunctionAD        
        ScriptVersion            1.0.0                             



        VERBOSE: Setting author to current user's full name, 'Paula Kingsley'
        Action: Create ISE Snippet 'PK AD Function'
        VERBOSE: Snippet 'PK AD Function' created successfully


            Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                                        
        ----                -------------         ------ ----                                                                                        
        -a----       2018-10-19  05:40 PM          15278 PK AD Function.snippets.ps1xml 

.EXAMPLE
    PS C\> New-PKISESnippetFunctionAD -Author "Empress Josephine" -Force -Verbose 

        VERBOSE: PSBoundParameters: 
	
        Key                      Value                     
        ---                      -----                     
        Author                   Empress Josephine         
        Force                    True                      
        Verbose                  True                      
        AutoDetectAuthorFullName False                     
        Confirm                                            
        ScriptName               New-PKISESnippetFunctionAD
        ScriptVersion            1.0.0                     

        VERBOSE: Setting author name to 'Empress Josephine'
        Action: Create ISE Snippet 'PK AD Function'
        VERBOSE: Snippet 'PK AD Function' created successfully


            Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                                        
        ----                -------------         ------ ----                                                                                        
        -a----       2018-10-19  05:42 PM          15298 PK AD Function.snippets.ps1xml      

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionAD -Force -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                        
        ---           -----                        
        Force         True                         
        Verbose       True                         
        ScriptName    New-PKISESnippetFunctionAD
        ScriptVersion 1.0.0                        

        ACTION: Create ISE Snippet 'PK AD Function'
        VERBOSE: Snippet 'PK AD Function' created successfully

            Directory: C:\Users\pkingsley\documents\WindowsPowerShell\Snippets


        Mode                LastWriteTime         Length Name                                                                           
        ----                -------------         ------ ----                                                                           
        -a----       2017-11-28     12:10          12156 PK AD Function.snippets.ps1xml 

.EXAMPLE
    PS C:\> New-PKISESnippetFunctionAD -Author foo  -Verbose 

    VERBOSE: PSBoundParameters: 
	
    Key                      Value                     
    ---                      -----                     
    Author                   foo                       
    Verbose                  True                      
    AutoDetectAuthorFullName False                     
    Force                    False                     
    Confirm                                            
    ScriptName               New-PKISESnippetFunctionAD
    ScriptVersion            1.0.0                     

    VERBOSE: Setting author name to 'foo'
    Action: Create ISE Snippet 'PK AD Function'
    Snippet 'PK AD Function' already exists; specify -Force to overwrite
    VERBOSE: 

        Directory: C:\Users\pkingsley\Documents\WindowsPowerShell\Snippets


    Mode                LastWriteTime         Length Name                                                                                        
    ----                -------------         ------ ----                                                                                        
    -a----       2018-10-19  05:42 PM          15298 PK AD Function.snippets.ps1xml 


#>
[Cmdletbinding(
    DefaultParameterSetName = "Name",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
Param(
    [Parameter(
        Mandatory = $True,
        ParameterSetName = "Name",
        HelpMessage = "Author name"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$Author,

    [Parameter(
        ParameterSetName = "Detect",
        HelpMessage = "Attempt to detect author name (currently logged-in user's full name)"
    )]
    [switch]$AutoDetectAuthorFullName,
    
    [Parameter(
        Mandatory = $False,
        HelpMessage = "Force creation of snippet even if name already exists"
    )]
    [switch]$Force
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

    # Preferences 
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }
    
    If (-not $PSISE) {
        $Msg = "This function requires the PowerShell ISE environment"
        $Host.UI.WriteErrorLine($Msg)
        Break
    }

    #region Snippet info
    
    $SnippetName = "PK AD Function"
    $Description = "Snippet to create a new generic Active Directory function; created using New-PKISESnippetFunctionAD"

    If ($AutoDetectAuthorFullName.IsPresent) {
        
        Function GetFullName {
            $SIDLocalUsers = Get-WmiObject Win32_UserProfile -EA Stop | select-Object Localpath,SID
            $UserName = (Get-WMIObject -class Win32_ComputerSystem -Property UserName -ErrorAction Stop).UserName
            $UserOnly = $UserName.Split("\")[1]
            Foreach ($Profile in $SIDLocalUsers) {
            
                # Match profile to current user
                If ($Profile.localpath -like "*$UserOnly"){

                    # Look up path in registry/Group Policy by SID
                    $SID = $Profile.sid
                    $DN = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\$SID" -ErrorAction SilentlyContinue | Select -ExpandProperty Distinguished-Name)
                    
                    If ($DN -match "(CN=)(.*?),.*") {
                        $Matches[2]        
                    }
                } #end if matching profile path
            }
        } #end function
        
        If (-not ($AuthorName = GetFullName)) {
            $Msg = "Failed to get current user's full name; please set Author manually"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
        Else {
            $Msg = "Setting author to current user's full name, '$AuthorName'"
            Write-Verbose $Msg
        }
    }
    Else {
        $AuthorName = $Author
        $Msg = "Setting author name to '$AuthorName'"
        Write-Verbose $Msg
    }
    
    #endregion Snippet info

    #region Here-string for function content    

$Body = @'
#Requires -version 4
Function Do-SomethingCool {
<# 
.SYNOPSIS
    Generic function using Active Directory module

.DESCRIPTION
    Generic function using Active Directory module
    Accepts pipeline input
    Returns a PSobject

.NOTES        
    Name    : Function_Do-SomethingCool.ps1
    Created : ##CREATEDATE##
    Author  : ##AUTHOR##
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **
        
        v01.00.0000 - ##CREATEDATE## - Created script
        
.PARAMETER ComputerName         
    One or more computernames (wildcards permitted)

.PARAMETER ADDomain             
    Active Directory domain name or FQDN (default is current user's domain)

.PARAMETER BaseDN               
    Starting Active Directory path/organizational unit (default is root of current user's domain)

.PARAMETER Server               
    Domain controller name or FQDN (default is first available)

.PARAMETER SuppressConsoleOutput
    Hide all non-verbose/non-error console output

.EXAMPLE
    PS C:\> Do-SomethingCool -ComputerName foo

#> 

[CmdletBinding(
    DefaultParameterSetName = "__DefaultParameterSet",
    SupportsShouldProcess = $True,
    ConfirmImpact = "High"
)]
[OutputType([Array])]
Param (
    [Parameter(
        Position = 0,
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="One or more computer names"
    )]
    [ValidateNotNullOrEmpty()]
    [Alias("Computer","Name","HostName","FQDN")]
    [String[]] $ComputerName,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Active Directory domain name or FQDN (default is current user's domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [string] $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name,

    [Parameter(
        Position = 1,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage="Starting Active Directory path/organizational unit (default is root of current user's domain)"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$BaseDN,
 
    [Parameter(
        HelpMessage="Domain controller name or FQDN (default is first available)"
    )]
    [ValidateNotNullOrEmpty()]
    [String] $Server,

    [Parameter(
        HelpMessage = "Suppress non-verbose, non-error console output"
    )]
    [Switch]$SuppressConsoleOutput


)
Begin {
    
     # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # How did we get here
    $PipelineInput = (-not $PSBoundParameters.ContainsKey("VM")) -and (-not $VM)
    $Source = $PSCmdlet.ParameterSetName

    # Show our settings
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
        Where-Object {Test-Path variable:$_}| Foreach-Object {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("PipelineInput",$PipelineInput)
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    # Preferences
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "Continue"
    $WarningPreference = "Continue"

    #region Prerequisites

    $Activity = "Prerequisites"

    # Make sure AD module is loaded
    # Use this only if not using a -Requires statement, which you may not want to do if this is part of a module
    # with non-AD-related functions

    $Msg = "Verify ActiveDirectory module"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    Try {
        If ($Module = Get-Module -Name ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue -Verbose:$False) {
            $Msg = "Successfully located ActiveDirectory module version $($Module.Version.ToString())"
            Write-Verbose "[Prerequisites] $Msg"
        }
        Else {
            $Msg = "Failed to locate ActiveDirectory module in PSModule path"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
    }
    Catch {
        $Msg = "Failed to locate ActiveDirectory module"
        If ($ErrorDetails = $_.exception.message) {$Msg += "; $ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        Break
    }


    $Msg = "Connect to Active Directory"
    Write-Verbose "[Prerequisites] $Msg"
    Write-Progress -Activity $Activity -CurrentOperation $Msg

    Try {
        $Param_GetAD = @{}
        $Param_GetAD = @{
            Identity    = $ADDomain
            ErrorAction = "Stop"
            Verbose     = $False
        }
        If ($CurrentParams.Credential) {
            $Param_GetAD.Add("Credential",$Credential)
        }
        $ADConfirm = Get-ADDomain @Param_GetAD
        $Msg = "Successfully connected to '$($ADConfirm.DNSRoot.Tolower())'"
        Write-Verbose "[Prerequisites] $Msg"
        
        # Get the domain controller
        If (-not $CurrentParams.Server) {
            
            $Msg = "Find nearest domain controller"
            Write-Verbose "[Prerequisites] $Msg"
            
            Try {        
                $Param_GetDC = @{}
                $Param_GetDC = @{
                    DomainName      = $ADDomain
                    Discover        = $True
                    NextClosestSite = $True
                    ErrorAction     = "Stop"
                    Verbose         = $False
                }
                $DCObj = Get-ADDomainController @Param_GetDC
                $DC = $($DCObj.HostName)
                $Msg = "Successfully connected to '$DC'"
                Write-Verbose "[Prerequisites] $Msg"
            }
            Catch {
                $Msg = "Failed to find domain controller for '$($ADConfirm.DNSRoot.Tolower())'"
                If ($ErrorDetails = $_.exception.message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                Break
            }    
        }
        Else {
            $Msg = "Connect to named domain controller"
            Write-Verbose "[Prerequisites] $Msg"

            Try {        
                $Param_GetDC = @{}
                $Param_GetDC = @{
                    Identity    = $Server
                    ErrorAction = "Stop"
                    Verbose     = $False
                }
                If ($CurrentParams.Credential) {
                    $Param_GetDC.Add("Credential",$Credential)
                }
                $DCObj = Get-ADDomainController @Param_GetDC
                If ($DCObj.Domain -eq $ADConfirm.DNSRoot) {
                    $DC = $($DCObj.HostName)
                    $Msg = "Successfully connected to '$DC'"
                    Write-Verbose "[Prerequisites] $Msg"
                }
                Else {
                    $Msg = "Domain controller '$($DCObj.HostName)' is not in domain '$($ADConfirm.DNSRoot)'"
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    Break
                }
            }
            Catch {
                $Msg = "Failed to find domain controller'$Server' in '$($ADConfirm.DNSRoot.Tolower())'"
                If ($ErrorDetails = $_.exception.message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("ERROR: $Msg")
                Break
            }
        }    
        # Get or verify the searchbase
        If ($DCObj) {
                
            If ($CurrentParams.BaseDN) {
                $Msg = "Verify searchbase"
                Write-Verbose "[Prerequisites] $Msg"
                Write-Progress -Activity "Prerequisites" -CurrentOperation $Msg

                $Param_GetBaseDN = @{}
                $Param_GetBaseDN = @{
                    Identity    = $BaseDN
                    Server      = $DC
                    ErrorAction = "SilentlyContinue"
                    Verbose     = $False
                }

                Try {
                    If ([string]$Searchbase = Get-ADObject @Param_GetBaseDN | Select-Object -ExpandProperty DistinguishedName) {
                        $Separator = ""
                        $Msg = "Searchbase set to '$SearchBase'"
                        Write-Verbose "[Prerequisites] $Msg"
                    }
                    Else {
                        $Msg = "Failed to validate searchbase '$BaseDN' in domain $($ADConfirm.DNSRoot)'"
                        $Host.UI.WriteErrorLine("ERROR: $Msg")
                        Break
                    }
                }
                Catch {
                    $Msg = "Failed to validate searchbase '$BaseDN'"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine("ERROR: $Msg")
                    Break
                }
            }    
            Else {
                $Msg = "Set searchbase"
                Write-Verbose "[Prerequisites] $Msg"
                Write-Progress -Activity "Prerequisites" -CurrentOperation $Msg

                [string]$Searchbase = $ADConfirm.DistinguishedName
                $Msg = "Successfully created searchbase as '$SearchBase'"
                Write-Verbose "[Prerequisites] $Msg"
            } 
        } 
    }
    Catch [exception] {
        $Msg = "Failed to connect to AD domain '$ADDomain'"
        If ($ErrorDetails = $_.exception.Message) {$Msg = "$Msg`n$ErrorDetails"}
        $Host.UI.WriteErrorLine("ERROR: $Msg")
        Break
    }

    #endregion Prerequisites

    #region Output object and property order

    # Will be returned if no match found
    $InitialValue = "Error"
    $OutputTemplate = New-Object PSObject -Property ([ordered]@{
        Computername = $InitialValue        
        Messages     = $InitialValue
    })
    $Select = $OutputTemplate.PSObject.Properties.Name

    #endregion Output object and property order
    
    #region Splats

    # General purpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
        Debug       = $False
    }

    # Splat for Get-ADComputer
    $Props = "Name","IPv4Address","Description","Location","OperatingSystem","DistinguishedName","CanonicalName","SID","WhenCreated","ServicePrincipalNames"
    $Select = "Name","IPv4Address","Description","Location","OperatingSystem","DistinguishedName","CanonicalName","SID","WhenCreated","ServicePrincipalNames",@{N="OU";E={$_.DistinguishedName -replace '^.+?(?<!\\),',''}}
    $Param_AD = @{}
    $Param_AD = @{
        Filter      = $Null
        Properties  = $Props
        Searchbase  = $SearchBase
        SearchScope = "Subtree"
        Server      = $DC
        ErrorAction = "Stop"
        Verbose     = $False
        Debug       = $False
    }

    # Splat for write-progress
    $Activity = "Generic Active Directory function"
    $Param_WP = @{}
    $Param_WP = @{
        Activity         = $Activity
        PercentComplete  = $Null
        CurrentOperation = $Null
        Status           = "Working"
    }

    #endregion Splats

    #region Functions

    # Convert DN to CN
    function Get-CanonicalName ([string[]]$DistinguishedName) {    
        foreach ($dn in $DistinguishedName) {      
            $d = $dn.Split(',') ## Split the dn string up into it's constituent parts 
            $arr = (@(($d | Where-Object { $_ -notmatch 'DC=' }) | ForEach-Object { $_.Substring(3) }))  ## get parts excluding the parts relevant to the FQDN and trim off the dn syntax 
            [array]::Reverse($arr)  ## Flip the order of the array. 
 
            ## Create and return the string representation in canonical name format of the supplied DN 
            $("{0}/{1}" -f  (($d | Where-Object { $_ -match 'dc=' } | ForEach-Object { $_.Replace('DC=','') }) -join '.'), ($arr -join '/')).TrimEnd("/") 
        } 
    }

    #endregion Functions  

    # Console output
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,"$Msg`n")}
    Else {Write-Verbose $Msg}


} #end begin

Process {

    # Counter for progress bar
    $Total = $ComputerName.Count
    $Current = 0

    Foreach ($Computer in $ComputerName) {
        
        $Computer = $Computer.Trim()
        $Current ++ 

        [int]$percentComplete = ($Current/$Total* 100)
        $Param_WP.PercentComplete = $PercentComplete
        $Param_WP.Status = $Computer
            
        $ConfirmMsg = "`n`n`t$Activity`n`n"
        If ($PSCmdlet.ShouldProcess($Computer,$ConfirmMsg)) {
                
            Try {
                $Msg = $Activity
                $Param_WP.CurrentOperation = $Msg
                Write-Verbose "[$Computer] $Msg"
                Write-Progress @Param_WP

                If ($Computer -match "\*") {
                    $Param_AD.Filter =  "Name -like '$Computer'"
                }
                Else {
                    $Param_AD.Filter =  "Name -eq '$Computer'"
                }
                $ADObj = Get-ADComputer @Param_AD | Select-Object $Select
                $Results += $ADObj
            }
            Catch {
                $Msg = "Operation failed"
                If ($ErrorDetails = $_.Exception.Message) {$Msg += "; $ErrorDetails"}
                $Host.UI.WriteErrorLine("[$Computer] $Msg")
            }
        }
        Else {
            $Msg = "Operation cancelled by user"
            $Host.UI.WriteErrorLine("[$Computer] $Msg")
        }
        
        
        
    } #end for each computer
        
}
End {
    
    $Null = Write-Progress -Activity $Activity -Completed

     If ($Results.Count -eq 0) {
        $Msg = "No output returned"
        Write-Warning $Msg
    }
    Else {
        $Msg = "$($Results.Count) result(s) found"
        Write-Verbose $Msg
        Write-Output ($Results | Select -Property *)
    }
}

} # end Do-SomethingCool


'@
    
    $Body = $Body.Replace("##AUTHOR##",$AuthorName)
    $Body = $Body.Replace("##CREATEDATE##",(Get-Date -f yyyy-MM-dd))
   
    #endregion Here-string for function content    

    #region Splat

    $Param_Snip = @{
        Title       = $SnippetName
        Description = $Description
        Text        = $Body
        Author      = $Author
        Verbose     = $False
        ErrorAction = "Stop"
    }
    If ($CurrentParams.ContainsKey("Force")) {
        $Param_Snip.Add("Force",$Force)
    }

    #endregion Splat

    # What are we doing
    $Activity = "Create ISE Snippet '$SnippetName'"
    $BGColor = $host.UI.RawUI.BackgroundColor
    $Msg = "Action: $Activity"
    $FGColor = "Yellow"
    If (-not $SuppressConsoleOutput.IsPresent) {$Host.UI.WriteLine($FGColor,$BGColor,$Msg)}
    Else {Write-Verbose $Msg} 

}
Process {

    # Filename
    $SnippetFile = $SnippetName +".snippets.ps1xml"

    If (($Test = Get-ISESnippet -ErrorAction SilentlyContinue | 
        Where-Object {$_.Name -eq $SnippetFile}) -and (-not $Force.IsPresent)) {
            $Msg = "Snippet '$SnippetName' already exists; specify -Force to overwrite"
            $Host.UI.WriteErrorLine("$Msg")
            Write-Verbose ($Test | Out-String)
    }    
    
    Else {
        $ConfirmMsg = "`n`n`t$Activity`n`n"
        If ($PSCmdlet.ShouldProcess($env:COMPUTERNAME,$ConfirmMsg)) {
        
            Try {
                
                # Create snippet
                $Create = New-IseSnippet @Param_Snip

                If ($IsSuccess = Get-ISESnippet  -ErrorAction SilentlyContinue | 
                    Where-Object {($_.Name -eq $SnippetFile) -and ($_.CreationTime -lt (get-date))}) {
                    $Msg = "Snippet '$SnippetName' created successfully"
                    Write-Verbose $Msg
                    Write-Output $IsSuccess
                }
                Else {
                    $Msg = "Failed to create snippet"
                    If ($ErrorDetails = $_.Exception.Message) {$Msg = "; $ErrorDetails"}
                    $Host.UI.WriteErrorLine($Msg)
                }
            }
            Catch {
                $Msg = "Error creating snippet"
                If ($ErrorDetails = $_.Exception.Message) {$Msg = "; $ErrorDetails"}
                $Host.UI.WriteErrorLine($Msg)
            }
        }
        Else {
            $Msg = "Snippet creation cancelled by user"
            $Host.UI.WriteErrorLine($Msg)
        }
    }
}
} #end New-PKISESnippetFunctionAD