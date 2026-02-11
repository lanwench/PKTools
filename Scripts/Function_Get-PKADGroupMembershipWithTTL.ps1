#requires -version 4
Function Get-PKADGroupMembershipWithTTL {
    <#
    .SYNOPSIS
        Audits a group to identify members, their managers, and any TTL expiration metadata.
    
    .DESCRIPTION
        Retrieves the membership of an AD Group. If the Privileged Access Management (PAM) feature is 
        active and members have a Time To Live (TTL), this function parses that metadata.
        It also attempts to find the exact "Date Added" using replication metadata.
    
    .NOTES
        v01.00.0000 - Created script
        
    .PARAMETER Name
        The Identity, Name, or DistinguishedName of the group(s) to audit.

    .PARAMETER ExpandMembers
        If specified, outputs a detailed object for each TTL member instead of a summary group object
    
    .PARAMETER Server
        The AD Domain or Domain Controller to connect to (defaults to the current user's domain)
    
    .PARAMETER Credential
        Optional credentials for Active Directory operations

    .EXAMPLE
        PS C:\> Get-PKADGroupMembershipWithTTL -Name "Temp local admins"  -Verbose
        # Returns a single summary object for a group with the number of members with & without TTLs, and a nested list of TTL members with their expiration details.

            VERBOSE: PSBoundParameters: 

            Key              Value
            ---              -----
            Name             {Temp local admins}
            Server           megacorp.local
            Verbose          True
            ExpandMembers    False
            Credential       
            PipelineInput    False
            ScriptName       Get-PKADGroupMembershipWithTTL
            ScriptVersion    1.0.0

            VERBOSE: [BEGIN: Get-PKADGroupMembershipWithTTL] Retrieving group membership and TTL expiration details
            VERBOSE: Getting PDC emulator for replication metadata queries
            VERBOSE: [CN=Temp local admins,OU=GroupPolicy,OU=Groups,DC=megacorp,DC=local] Fetching group membership details with TTL metadata
            VERBOSE: [CN=Temp local admins,OU=GroupPolicy,OU=Groups,DC=megacorp,DC=local] Found 1 members with TTLs and 1195 permanent members
            VERBOSE: [CN=Temp local admins,OU=GroupPolicy,OU=Groups,DC=megacorp,DC=local] Parsing TTL member details and fetching additional attributes
            VERBOSE: [CN=Temp local admins,OU=GroupPolicy,OU=Groups,DC=megacorp,DC=local] <TTL=7764876>,CN=LAPTOP17,OU=Laptops,DC=megacorp,DC=local

            Group             : Temp local admins
            GroupCategory     : Security
            DistinguishedName : CN=Temp local admins,OU=GroupPolicy,OU=Groups,DC=megacorp,DC=local
            Mail              : 
            ManagedBy         : CN=Jennifer Collegio,OU=IT,OU=Staff,DC=megacorp,DC=local
            Description       : Computers in this group will have the ManagedBy user added to the local Administrators group
            NumMembers        : 6
            NumTtlMembers     : 1
                TTLMembers    : @{Name=LAPTOP17; Enabled=True; ObjectClass=computer; DistinguishedName=CN=LAPTOP17,OU=Laptops,DC=megacorp,DC=local; Manager=CN=Fanny Price,
                                OU=Enabled Users,OU=Accounts,DC=megacorp,DC=local; DateAddedToGroup=2026-02-10 12:59:56; ExpirationDate=2026-05-11 13:00:02; SecondsLeft=7764876;
                                TimeLeft=89 days, 20 hours, 54 minutes, 35 seconds; MetadataFound=True}
        
    .EXAMPLE
        PS C:\> Get-PKADGroupMembershipWithTTL -Name "temp local admins"  -ExpandMembers        

        # Returns individual objects for each TTL member instead of a summary group object that requires -expandproperty

            Name              : LAPTOP17
            Enabled           : True
            ObjectClass       : computer
            DistinguishedName : CN=LAPTOP17,OU=Laptops,DC=megacorp,DC=local
            Manager           : CN=Fanny Price,OU=Enabled Users,OU=Accounts,DC=megacorp,DC=local
            DateAddedToGroup  : 2026-02-10 12:59:56
            ExpirationDate    : 2026-05-11 13:00:02
            TimeLeft          : 89 days, 20 hours, 48 minutes, 25 seconds
            MetadataFound     : True
        

    #>
    [CmdletBinding()]
    Param(
        [Parameter(
            Mandatory = $True,
            Position = 0,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            HelpMessage = "Specify the identity, name, or distinguishedname of the group(s) to audit"
        )]
        [object[]]$Name,
        
        [Parameter(
            HelpMessage = "Returns only the TTLmember list without additional group details"
        )]
        [switch]$ExpandMembers,

        [Parameter(
            HelpMessage = "Specify the AD domain or domain controller to connect to (defaults to the current user's domain)"
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Server = $env:USERDNSDOMAIN,

        [Parameter(
            HelpMessage = "Specify credentials for Active Directory operations"
        )]
        [ValidateNotNullOrEmpty()]
        [PSCredential]$Credential
    )

    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.000"

        # How did we get here?
        $PipelineInput = $PSCmdlet.MyInvocation.ExpectingInput
        $CurrentParams = $PSBoundParameters
        $ScriptName = $MyInvocation.MyCommand.Name
        $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} | 
            Where-Object {Test-Path variable:$_}| ForEach-Object {
                $CurrentParams.Add($_, (Get-Variable $_).value)
            }
        $CurrentParams.Add("PipelineInput",$PipelineInput)
        $CurrentParams.Add("ScriptName",$ScriptName)
        $CurrentParams.Add("ScriptVersion",$Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
        
        $Activity = "Retrieving group membership and TTL expiration details"
        Write-Verbose "[BEGIN: $ScriptName] $Activity"
        
        If (-not ($Null = Get-Module ActiveDirectory -ErrorAction SilentlyContinue)) {
            Throw "The ActiveDirectory module is required for this function to run. Please install RSAT tools or run on a Domain Controller."
        }

        # Splats & PDC emulator retrieval for replication metadata queries (to find DateAdded for TTL members)
        $ADParams = @{ Server = $Server }
        If ($PSBoundParameters.ContainsKey('Credential') -and $null -ne $Credential) {$ADParams.Add('Credential', $Credential)}
        Write-Verbose "Getting PDC emulator for replication metadata queries"
        $DomainObj = Get-ADDomain @ADParams -ErrorAction Stop
        $PDCParams = @{
            Server = $DomainObj.PDCEmulator
            ErrorAction = 'Stop'
        }
        If ($ADParams.Credential) {$PDCParams.Add('Credential', $Credential)}

        $Output = [System.Collections.Generic.List[PSObject]]::new()
        
    }
    Process {
        Foreach ($GroupIdentity In $Name) { 

            Try {

                Write-Verbose "[$($GroupIdentity.ToString())] Fetching group membership details with TTL metadata"
                
                # Have to use -ShowMemberTimeToLive in order to see the <TTL=...> prefix in the DN
                $GroupObj = Get-ADGroup -Identity $GroupIdentity @ADParams -Properties Member, ManagedBy, Mail, Description, GroupCategory -ShowMemberTimeToLive -ErrorAction Stop
                $TtlMembers = $GroupObj.Member | Where-Object { $_ -like "<TTL=*" }
                $NoTTL = $GroupObj.Member | Where-Object { $_ -notlike "<TTL=*" }
                Write-Verbose "[$($GroupIdentity.ToString())] Found $($TtlMembers.Count) members with TTLs and $($NoTTL.Count) permanent members"

                # Initialize the report array
                $TtlMemberReport = @()

                # Parse TTL members into a detailed report
                If ($TtlMembers) {
                    Write-Verbose "[$($GroupIdentity.ToString())] Parsing TTL member details and fetching additional attributes"

                    $TtlMemberReport = $TtlMembers | ForEach-Object {
                        $Me = $_
                        Write-Verbose "[$($GroupIdentity.ToString())] $Me"

                        # Extract the TTL value and the DN using Regex
                        If ($Me -match "<TTL=(?<Seconds>\d+)>,(?<DN>.*)") {
                            $CleanDN = $Matches['DN']
                            Try {
                                # Fetch only necessary properties to optimize speed. Get-ADObject doesn't calculate 'Enabled', so we fetch UserAccountControl.
                                $MemberObj = Get-ADObject -Identity $CleanDN -Properties ManagedBy, Manager, UserAccountControl, WhenChanged, ObjectClass @ADParams -ErrorAction Stop
                                
                                # Get Date Added from Replication Metadata (Forensics) to see when membership was created, using the PDC emulator
                                # We MUST use -ShowAllLinkedValues to find metadata for specific group member links (especially ephemeral ones)
                                $Meta = Get-ADReplicationAttributeMetadata -Object $GroupObj -Attribute member @PDCParams -ShowAllLinkedValues | 
                                        Where-Object { $_.AttributeValue -eq $CleanDN }
                                
                                # Calculate Expiration and TimeLeft string
                                $ExpDate = (Get-Date).AddSeconds([int]$Matches['Seconds'])
                                $Span = $ExpDate - (Get-Date)
                                $TimeLeftStr = "{0:D2} days, {1:D2} hours, {2:D2} minutes, {3:D2} seconds" -f $Span.Days, $Span.Hours, $Span.Minutes, $Span.Seconds

                                # Return the object to the pipeline so it's captured in $TtlMemberReport
                                [PSCustomObject]@{
                                    Name              = $MemberObj.Name
                                    Enabled           = If ($Null -ne $MemberObj.UserAccountControl) { -not ($MemberObj.UserAccountControl -band 2) } Else { "n/a" }
                                    ObjectClass       = $MemberObj.ObjectClass
                                    DistinguishedName = $CleanDN
                                    Manager           = If ($MemberObj.ManagedBy) { $MemberObj.ManagedBy } Else { $MemberObj.Manager }
                                    DateAddedToGroup  = If ($Meta) { $Meta.LastOriginatingChangeTime } Else { $Null }
                                    ExpirationDate    = $ExpDate
                                    TimeLeft          = $TimeLeftStr
                                    Messages          = If ($Meta) { "DateAdded found in replication metadata" } Else { "Replication metadata not found" }
                                }
                            }
                            Catch {
                                Write-Warning "Could not resolve member object details for: $CleanDN"
                            }
                        }
                    }
                } # end if members with TTL were found
                If ($TtlMemberReport.Count -eq 0) {
                    Write-Warning "[$($GroupIdentity.ToString())] No TTL members found or failed to parse TTL member details"
                }
                ElseIf ($ExpandMembers.IsPresent) {
                    $TtlMemberReport | ForEach-Object { $Output.Add($_) }
                }
                Else {
                    # Construct the final group report object
                    $GroupSummary = [PSCustomObject]@{
                        Group             = $GroupObj.Name
                        GroupCategory     = $GroupObj.GroupCategory
                        DistinguishedName = $GroupObj.DistinguishedName
                        Mail              = $GroupObj.Mail
                        ManagedBy         = $GroupObj.ManagedBy
                        Description       = $GroupObj.Description
                        NumMembers        = $GroupObj.Member.Count
                        NumTtlMembers     = $TtlMembers.Count
                        TTLMembers        = $TtlMemberReport
                    }
                    $Output.Add($GroupSummary)
                }
            }
            Catch {
                Write-Error "[$($GroupIdentity.ToString())] Failed to process group! $($_.Exception.Message)"
            }
        }
        
        # Output the results to the pipeline
        Write-Output $Output
    }

    End {
        Write-Verbose "[END: $ScriptName] Group audit complete"
    }
}#end function Get-PKADGroupMembershipWithTTL