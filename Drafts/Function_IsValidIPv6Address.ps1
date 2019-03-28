﻿#https://www.powershelladmin.com/wiki/PowerShell_.NET_regex_to_validate_IPv6_address_(RFC-compliant)
# Joakim Svendsen
function Test-IsValidIPv6Address {
    param(
        [Parameter(Mandatory=$true,HelpMessage='Enter IPv6 address to verify')] [string] $IP)
    $IPv4Regex = '(((25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2})\.){3}(25[0-5]|2[0-4][0-9]|[0-1]?[0-9]{1,2}))'
    $G = '[a-f\d]{1,4}'
    # In a case sensitive regex, use:
    #$G = '[A-Fa-f\d]{1,4}'
    $Tail = @(":",
        "(:($G)?|$IPv4Regex)",
        ":($IPv4Regex|$G(:$G)?|)",
        "(:$IPv4Regex|:$G(:$IPv4Regex|(:$G){0,2})|:)",
        "((:$G){0,2}(:$IPv4Regex|(:$G){1,2})|:)",
        "((:$G){0,3}(:$IPv4Regex|(:$G){1,2})|:)",
        "((:$G){0,4}(:$IPv4Regex|(:$G){1,2})|:)")
    [string] $IPv6RegexString = $G
    $Tail | foreach { $IPv6RegexString = "${G}:($IPv6RegexString|$_)" }
    $IPv6RegexString = ":(:$G){0,5}((:$G){1,2}|:$IPv4Regex)|$IPv6RegexString"
    $IPv6RegexString = $IPv6RegexString -replace '\(' , '(?:' # make all groups non-capturing
    [regex] $IPv6Regex = $IPv6RegexString
    if ($IP -imatch "^$IPv6Regex$") {
        $true
    } else {
        $false
    }
}