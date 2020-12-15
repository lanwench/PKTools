#requires -version 4
Function Convert-PKDNSRecordData {
<#
.SYNOPSIS
    Converts the RecordData output from Get-DNSZoneResourceRecord to a string based on record type

.DESCRIPTION
    Converts the RecordData output from Get-DNSZoneResourceRecord to a string based on record type
    The RecordData property is a rich object and doesn't export to CSV nicely
    Accepts pipeline input
    Returns a string

.NOTES        
    Name    : Function_Convert-PKDNSRecordData.ps1
    Created : 2020-09-09
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :

        ** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2020-09-09 - Created script

.EXAMPLE
    PS C:\> Get-DnsServerResourceRecord -ZoneName testdomain.com -ComputerName $DNSServer | Select -First 10 | Convert-PKDNSRecordData -Verbose

        VERBOSE: Hostname 154.227.38.10, type PTR
        exchange02.testdomain.com.
        VERBOSE: Hostname *.allapps, type CNAME
        appserver.testdomain.com.
        VERBOSE: Hostname hrserver, type A
        10.32.8.14
        VERBOSE: Hostname dev.testdomain.com, type TXT
        v=spf1 include:_spf.mailhost.com ~all
        VERBOSE: Hostname _autodiscover._tcp, type SRV
        autodiscover.testdomain.com.
        VERBOSE: Hostname 14.8.32.10, type PTR
        hrserver.testdomain.com.

.EXAMPLE
    PS C:\> Get-DnsServerResourceRecord -ZoneName testdomain.com -Computername dns1.internal.lan | Select-Object Hostname,
        RecordType,
        Type,
        Timestamp,
        TimeToLive,
        @{N="RecordData";E={$_ | Convert-PKDNSRecordData}},
        @{N="DNSServer";E={"dns1.internal.lan"}}


        Hostname   : 154.227.38.10
        RecordType : PTR
        Type       : 12
        Timestamp  : 
        TimeToLive : 01:00:00
        RecordData : server12345.testdomain.com.
        DNSServer  : dnsserver1.internal.lan
        
#>
[CmdletBinding()]
Param(
    [Parameter(
        ValueFromPipeline = $True,
        ValueFromPipelineByPropertyName = $True
    )]
    $DNSRecord
)
Begin {}
Process {

    Foreach ($Record in $DNSRecord) {
        Write-Verbose "Hostname $($Record.HostName), type $($Record.RecordType)"
        Switch ($Record.RecordType) {
            A     {$Record.RecordData.ipv4address.IPAddressToString}
            NS    {$Record.RecordData.nameserver}
            SOA   {$Record.RecordData.PrimaryServer}
            MX    {$Record.RecordData.MailExchange}
            TXT   {$Record.RecordData.DescriptiveText}
            PTR   {$Record.RecordData.PtrDomainName}
            CNAME {$Record.RecordData.HostNameAlias}
            SRV   {$Record.RecordData.DomainName}
        }
    }
}
}
