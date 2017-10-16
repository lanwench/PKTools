#requires -version 3
Function Get-PKNetworkPortAssignments {
<#
.Synopsis
   Downloads and returns a list of known port assignments from IANA 

.DESCRIPTION
   Downloads and returns a list of known port assignments from IANA 
   Allows filtering on service name, port, and transport protocol
   Returns a PSObject

.NOTES
    Name    : Function_Get-PKNetworkPortAssignments.ps1
    Version : 01.00.0000
    Author  : Paula Kingsley
    Created : 2017-06-26

    History:

        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK *

        v01.00.0000 - 2017-06-26 - Created script based on Idera community link

.LINK
    http://community.idera.com/powershell/powertips/b/tips/posts/get-list-of-port-assignments

        
.EXAMPLE
    PS C:\> Get-PKNetworkPortAssignments -Verbose | Format-Table -AutoSize

        VERBOSE: PSBoundParameters: 
	
        Key           Value                                                                                     
        ---           -----                                                                                     
        Verbose       True                                                                                      
        Type          All                                                                                       
        URL           https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.csv
        ScriptName    Get-PKNetworkPortAssignments                                                              
        ScriptVersion 1.0.0                                                                                     



        VERBOSE: 14012 result(s) found for all transport protocols

        PortNumber ServiceName Protocol Description                      Assignee        Contact         DateRegistered DateModified Reference ServiceCode
        ---------- ----------- -------- -----------                      --------        -------         -------------- ------------ --------- -----------
        0                      TCP      Reserved                         Jon_Postel      Jon_Postel                                                       
        0                      UDP      Reserved                         Jon_Postel      Jon_Postel                                                       
        1          tcpmux      TCP      TCP Port Service Multiplexer     Mark_Lottor     Mark_Lottor                                                      
        1          tcpmux      UDP      TCP Port Service Multiplexer     Mark_Lottor     Mark_Lottor                                                      
        2          compressnet TCP      Management Utility                                                                                                
        2          compressnet UDP      Management Utility                                                                                                
        3          compressnet TCP      Compression Process              Bernie_Volz     Bernie_Volz                                                      
        3          compressnet UDP      Compression Process              Bernie_Volz     Bernie_Volz                                                      
        4                      TCP      Unassigned                                                                                                        
        4                      UDP      Unassigned                                                                                                        
        5          rje         TCP      Remote Job Entry                 Jon_Postel      Jon_Postel                                                       
        5          rje         UDP      Remote Job Entry                 Jon_Postel      Jon_Postel                                                       
        6                      TCP      Unassigned                                                                                                        
        6                      UDP      Unassigned                                                                                                        
        7          echo        TCP      Echo                             Jon_Postel      Jon_Postel                                                       
        7          echo        UDP      Echo                             Jon_Postel      Jon_Postel                                                       
        8                      TCP      Unassigned                                                                                                        
        8                      UDP      Unassigned                                                                                                        
        9          discard     TCP      Discard                          Jon_Postel      Jon_Postel                                                       
        9          discard     UDP      Discard                          Jon_Postel      Jon_Postel                                                       
        9          discard     SCTP     Discard                          Randall_Stewart Randall_Stewart                             [RFC4960]            
        9          discard     DCCP     Discard                          Eddie_Kohler    Eddie_Kohler                                [RFC4340] 1145656131 
        10                     TCP      Unassigned                                                                                                        
        10                     UDP      Unassigned                                                                                                        
        11         systat      TCP      Active Users                     Jon_Postel      Jon_Postel                                                       
        11         systat      UDP      Active Users                     Jon_Postel      Jon_Postel                                                       
        12                     TCP      Unassigned                                                                                                        
        12                     UDP      Unassigned                                                                                                        
        13         daytime     TCP      Daytime                          Jon_Postel      Jon_Postel                                  [RFC867]             
        13         daytime     UDP      Daytime                          Jon_Postel      Jon_Postel                                  [RFC867]             
        14                     TCP      Unassigned                                                                                                        
        14                     UDP      Unassigned                                                                                                        
        15                     TCP      Unassigned [was netstat]                                                                                          
        15                     UDP      Unassigned                                                                                                        
        16                     TCP      Unassigned                                                                                                        
        16                     UDP      Unassigned                                                                                                        
        17         qotd        TCP      Quote of the Day                 Jon_Postel      Jon_Postel                                                       
        17         qotd        UDP      Quote of the Day                 Jon_Postel      Jon_Postel                                                       
        18         msp         TCP      Message Send Protocol (historic) Rina_Nethaniel  Rina_Nethaniel                                                   
        18         msp         UDP      Message Send Protocol (historic) Rina_Nethaniel  Rina_Nethaniel                                                

        <snip>


.EXAMPLE
    PS C:\> Get-PKNetworkPortAssignments -Name smtp,pop3,dns -Type TCP -Verbose | Format-Table

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                                                                     
        ---              -----                                                                                     
        Name             {smtp, pop3, dns}                                                                         
        Type             TCP                                                                                       
        Verbose          True                                                                                      
        Port                                                                                                       
        URL              https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.csv
        ScriptName       Get-PKNetworkPortAssignments                                                              
        ScriptVersion    1.0.0                                                                                     
        ParameterSetName ByName                                                                                    

        VERBOSE: 2 matching result(s) found for transport protocol TCP

        PortNumber ServiceName Protocol Description                      Assignee      Contact       DateRegistered DateModified Reference ServiceCode
        ---------- ----------- -------- -----------                      --------      -------       -------------- ------------ --------- -----------
        25         smtp        TCP      Simple Mail Transfer             IESG          IETF Chair                   2017-06-05   [RFC5321]            
        110        pop3        TCP      Post Office Protocol - Version 3 Marshall Rose Marshall Rose                                                  


.EXAMPLE
    PS C:\> Get-PKNetworkPortAssignments -Port 80 -Verbose | Format-Table

        VERBOSE: PSBoundParameters: 
	
        Key              Value                                                                                     
        ---              -----                                                                                     
        Port             {80}                                                                                      
        Verbose          True                                                                                      
        Name                                                                                                       
        Type             All                                                                                       
        URL              https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.csv
        ScriptName       Get-PKNetworkPortAssignments                                                              
        ScriptVersion    1.0.0                                                                                     
        ParameterSetName ByPort                                                                                    


        VERBOSE: 7 matching result(s) found for all transport protocols

        PortNumber ServiceName Protocol Description         Assignee        Contact         DateRegistered DateModified Reference ServiceCode
        ---------- ----------- -------- -----------         --------        -------         -------------- ------------ --------- -----------
        80         http        TCP      World Wide Web HTTP                                                                                  
        80         http        UDP      World Wide Web HTTP                                                                                  
        80         www         TCP      World Wide Web HTTP                                                                                  
        80         www         UDP      World Wide Web HTTP                                                                                  
        80         www-http    TCP      World Wide Web HTTP Tim Berners Lee Tim Berners Lee                                                  
        80         www-http    UDP      World Wide Web HTTP Tim Berners Lee Tim Berners Lee                                                  
        80         http        SCTP     HTTP                Randall Stewart Randall Stewart                             [RFC4960]   


.EXAMPLE
    PS C:\> Get-PKNetworkPortAssignments -Port 4060
    
        PortNumber      : 4060
        ServiceName     : dsmeter-iatc
        Protocol        : TCP
        Description     : DSMETER Inter-Agent Transfer Channel
                          IANA assigned this well-formed service name as a replacement for "dsmeter_iatc".
        Assignee        : John McCann
        Contact         : John McCann
        DateRegistered  : 2006-12
        DateModified    : 
        Reference       : 
        ServiceCode     : 
        AssignmentNotes : 

        PortNumber      : 4060
        ServiceName     : dsmeter_iatc
        Protocol        : TCP
        Description     : DSMETER Inter-Agent Transfer Channel
        Assignee        : John McCann
        Contact         : John McCann
        DateRegistered  : 2006-12
        DateModified    : 
        Reference       : 
        ServiceCode     : 
        AssignmentNotes : This entry is an alias to "dsmeter-iatc".
                          This entry is now historic, not usable for use with many
                          common service discovery mechanisms.

        PortNumber      : 4060
        ServiceName     : dsmeter-iatc
        Protocol        : UDP
        Description     : DSMETER Inter-Agent Transfer Channel
                          IANA assigned this well-formed service name as a replacement for "dsmeter_iatc".
        Assignee        : John McCann
        Contact         : John McCann
        DateRegistered  : 2006-12
        DateModified    : 
        Reference       : 
        ServiceCode     : 
        AssignmentNotes : 

        PortNumber      : 4060
        ServiceName     : dsmeter_iatc
        Protocol        : UDP
        Description     : DSMETER Inter-Agent Transfer Channel
        Assignee        : John McCann
        Contact         : John McCann
        DateRegistered  : 2006-12
        DateModified    : 
        Reference       : 
        ServiceCode     : 
        AssignmentNotes : This entry is an alias to "dsmeter-iatc".
                          This entry is now historic, not usable for use with many
                          common service discovery mechanisms.

#>
[Cmdletbinding(DefaultParameterSetName="All")]
Param(
    [Parameter(
        ParameterSetName = "ByPort",
        Mandatory = $True,
        HelpMessage = "Port number(s)"
    )]
    [ValidateNotNullOrEmpty()]
    [int[]]$Port,

    [Parameter(
        ParameterSetName = "ByName",
        Mandatory = $True,
        HelpMessage = "Service name(s)"
    )]
    [ValidateNotNullOrEmpty()]
    [string[]]$Name,

    [Parameter(
        Mandatory = $False,
        HelpMessage = "Type of port data to return - TCP, UDP, or All"
    )]
    [ValidateSet("TCP","UDP","All")]
    [ValidateNotNullOrEmpty()]
    [string]$Type = "All",

    [Parameter(
        Mandatory = $False,
        HelpMessage = "IANA port assignment CSV URL"
    )]
    [ValidateNotNullOrEmpty()]
    [string]$URL = 'https://www.iana.org/assignments/service-names-port-numbers/service-names-port-numbers.csv'

)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Generalpurpose splat
    $StdParams = @{}
    $StdParams = @{
        ErrorAction = "Stop"
        Verbose     = $False
    }

    # Show our settings
    $Source = $PSCmdlet.ParameterSetName
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where {$CurrentParams.keys -notContains $_} | 
        Where {Test-Path variable:$_}| Foreach {
            $CurrentParams.Add($_, (Get-Variable $_).value)
        }
    $CurrentParams.Add("ScriptName",$MyInvocation.MyCommand.Name)
    $CurrentParams.Add("ScriptVersion",$Version)
    $CurrentParams.Add("ParameterSetName",$Source)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
    
}
Process {
    
    $OutFile = "$env:TEMP\portlist.csv"

    If ($Null = Test-Path $OutFile) {
        Try {
            $Null = Get-Item $OutFile @StdParams | Remove-Item -Force @StdParams
        }
        Catch {
            $Msg = "Can't remove existing file '$OutFile'"
            $ErrorDetails = $_.Exception.Message
            $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
            Break
        }
    }

    Try {
        $Msg = "Retrieving service name/port number CSV from IANA.org"
        Write-Progress -Activity $Msg
        
        # Download as a temp file, then import from CSV 
        $Web = Invoke-WebRequest -Uri $URL -UseBasicParsing -OutFile $OutFile @StdParams
        $Import = Import-Csv -Path $OutFile -Encoding UTF8 
        
        # Kill the temp file
        Try {
            $Null = Get-Item $OutFile @StdParams | Remove-Item -Force -ErrorAction SilentlyContinue -Verbose:$False -Confirm:$False
        }
        Catch {
        }

        Write-Progress -Activity $Msg -Completed

        # Filter on name/port or all
        Switch ($Source){
            ByName {
                $Results =  ($Import | Where-Object {$($_."Service Name") -in $Name})
            }
            ByPort {
                $Results = ($Import | Where-Object {$($_."Port Number") -in $Port})

            }
            Default {
                $Results = $Import
            }
        }

        # Filter on transport proptocol, or all
        Switch ($Type) {
            All {
                $Msg = "$($Results.Count) matching result(s) found for all transport protocols" 
            }
            TCP {
                $Results = $Results | Where-Object {$($_."Transport Protocol") -eq "TCP"}
                $Msg = "$($Results.Count) matching result(s) found for transport protocol TCP" 
            }
            UDP {
                $Results = $Results | Where-Object {$($_."Transport Protocol") -eq "UDP"}
                $Msg = "$($Results.Count) matching result(s) found for transport protocol UDP" 
            }
        }
        If ($Results.Count -eq 0) {
            $Host.UI.WriteErrorLine($Msg)
        }
        
        Else {
            Write-Verbose $Msg

            # Normalize/format and return the output
            Write-Output ($Results | 
                Select @{N="PortNumber";E={$($_."Port Number")}},
                @{N="ServiceName";E={$($_."Service Name")}},
                @{N="Protocol";E={$($_."Transport Protocol").ToUpper()}},
                Description,
                @{N="Assignee";E={$($_.Assignee.Replace("_"," ").Replace("[",$Null).Replace("]",$Null))}},
                @{N="Contact";E={$($_.Contact.Replace("_"," ").Replace("[",$Null).Replace("]",$Null))}},
                @{N="DateRegistered";E={$($_."Registration Date")}},
                @{N="DateModified";E={$($_."Modification Date")}},
                Reference,
                @{N="ServiceCode";E={$($_."Service Code")}},
                @{N="AssignmentNotes";E={$($_."Assignment Notes")}})
        }
    }
    Catch {
        $Msg = "Download failed"
        $ErrorDetails = $_.Exception.Message
        $Host.UI.WriteErrorLine("ERROR: $Msg`n$ErrorDetails")
    }

}
} #end Get-PKNetworkPortAssignments