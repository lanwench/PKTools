#requires -Version 3

Function Get-PKWindowsNetIP {
[Cmdletbinding()]
Param($ComputerName = $env:COMPUTERNAME)

$ScriptBlock = {
    $ErrorActionPreference = "Continue"
    $NICs = Get-NetAdapter
    Foreach ($NICObj in $NICS) {
        Try {
            $NICObj | Get-NetIPConfiguration | Where-Object {$_.ipv4address} | Select `
                @{Name="Name"; expression={$NicObj.name}},
                InterfaceDescription,
                InterfaceIndex,
                MACAddress,
                @{N="Status";E={$NICObj.Status}},
                @{N="DHCPEnabled";E={
                    If (($NICObj | Get-NetIPInterface).DHCP -eq "Enabled") {$True}
                    Else {$False}
                }},
                @{N="IPAddress";E={$_.IPv4Address.IPAddress}},
                @{N="DefaultGateway";E={$_.IPv4DefaultGateway.NextHop}},
                @{N="SubnetMask";E={
                    $Index = ($NICObj | Get-NetIPAddress -AddressFamily IPv4).PrefixLength
                    (('1'*$Index+'0'*(32-$Index)-split'(.{8})')-ne''| Foreach-Object {[convert]::ToUInt32($_,2)})-join'.'}
                },
                @{N="DNSServers";E={($_.DNSServer | Where-Object {$_.AddressFamily -eq 2}).ServerAddresses}}
        }    
        Catch {
            $_.Exception.Message
        }
    } #end for each
}
Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock


}


<#
    
    QBYCTX01



$Switches = @($CheckTools,$CheckGuestOS,$CheckVMScript) 
If ((($Switches | Foreach-Object {$_.IsPresent} ) -contains $True)) {

}
$Steps = (($Switches | Where-Object {$_.IsPresent} ).Count)





$AliasNum = $Nics[0].Name | Select-String '(?<num>\d{1,5})' |Foreach-Object {$_.matches[0].Groups['num'].value}



"$($NIC.MacAddress.Replace(':','-'))"

"`$NICObj = Get-NetAdapter | Where-Object {`$_.MACAddress -eq '$($NIC.MacAddress.Replace(':','-'))'}"


$NICObj | Get-NetIPAddress | where {$_.ipv4address} |
    select `
        @{Name="Adapter"; expression={$NicObj.name}},
        IPv4Address,
        @{n="SubnetMask"; e={
            $prefix = ($_| select -expand prefixlength)
            write-output (
                '{0}.{1}.{2}.{3}' -f ([math]::Truncate(([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) / 16777216)), 
                ([math]::Truncate((([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) % 16777216) / 65536)), 
                ([math]::Truncate((([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) % 65536)/256)), 
                ([math]::Truncate(([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) % 256))
            )
        }
    }

    
$NICs = Get-NetAdapter | Where {$_.Status -ne "Disabled"}
Foreach ($NICObj in $NICS) {
$NICObj | Get-NetIPConfiguration | Where-Object {$_.ipv4address} |
    select `
        @{Name="Adapter"; expression={$NicObj.name}},
        InterfaceIndex,
        InterfaceDescription,
        @{N="IPAddress";E={$_.IPv4Address.IPAddress}},
        @{N="DefaultGateway";E={$_.IPv4DefaultGateway.NextHop}},
        @{N="DNSServers";E={($_.DNSServer | Where-Object {$_.AddressFamily -eq 2}).ServerAddresses}},
        @{n="SubnetMask"; e={
            $prefix = ($_ | select -expand prefixlength)
            write-output (
                '{0}.{1}.{2}.{3}' -f ([math]::Truncate(([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) / 16777216)), 
                ([math]::Truncate((([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) % 16777216) / 65536)), 
                ([math]::Truncate((([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) % 65536)/256)), 
                ([math]::Truncate(([convert]::ToInt64(('1' * $prefix + '0' * (32 - $prefix)), 2)) % 256))
            )
        }
    }
}    


#>