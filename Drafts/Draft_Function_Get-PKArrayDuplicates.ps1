$List1 = @(
 [PSCustomObject]@{Alias = 1; Place = 1; Extra = 'c'}
 [PSCustomObject]@{Alias = 2; Place = 3; Extra = 'a'}
 [PSCustomObject]@{Alias = 3; Place = 2; Extra = 'c'}
 [PSCustomObject]@{Alias = 4; Place = 1; Extra = 'a'}
 [PSCustomObject]@{Alias = 22; Place = 3; Extra = 'g'}
 [PSCustomObject]@{Alias = 2; Place = 3; Extra = 'a'}
 [PSCustomObject]@{Alias = 5; Place = 6; Extra = 'e'}
 [PSCustomObject]@{Alias = 4; Place = 2; Extra = 'c'}
 [PSCustomObject]@{Alias = 1; Place = 6; Extra = 'b'}
 )
 
$List2 = @(
 [PSCustomObject]@{Name = 1; Place = 5; Somthing = 'a1'}
 [PSCustomObject]@{Name = 1; Place = 1; Somthing = 'b6'}
 [PSCustomObject]@{Name = 5; Place = 1; Somthing = 'c3'}
 [PSCustomObject]@{Name = 2; Place = 4; Somthing = 'a3'}
 [PSCustomObject]@{Name = 12; Place = 6; Somthing = 'a1'}
 [PSCustomObject]@{Name = 1; Place = 2; Somthing = 'b1'}
 [PSCustomObject]@{Name = 2; Place = 7; Somthing = 'd4'}
 [PSCustomObject]@{Name = 44; Place = 2; Somthing = 'a5'}
 )


[array]$AllSPNList = ($Output | sort-object ServicePrincipalNames)
[array]$UniqueSPNs = ($AllSPNList.ServicePrincipalNames | Select -unique)
[array]$DuplicateSPNs = Compare-Object -ReferenceObject $UniqueSPNs -DifferenceObject ($AllSPNList.ServicePrincipalNames)

$Dupe = @()
 $Output | Sort ServicePrincipalNames | Where-Object {
    
    #$_.ServicePrincipalNames -match "dms-rndbqa-1"
    
    Foreach ($SPN in ($_.ServicePrincipalNames | sort)) {
        
        Write-Host -f cyan $SPN

        If ($DuplicateSPNs.IndexObject -contains $SPN) {
            $True
            $Dupe += $_
        }
        Else {$False}
    }


}



Function GetArrayDupes{
# http://www.mickputley.net/2013/06/find-duplicates-in-array-powershell.html
   param($array)
   $hash = @{}
   $array | %{ $hash[$_] = $hash[$_] + 1 }
   $result = $hash.GetEnumerator() | ?{$_.value -gt 1} | %{$_.key}
   Return $result
}


Function Get-PKArrayDuplicates{
# http://www.mickputley.net/2013/06/find-duplicates-in-array-powershell.html
[cmdletbinding()]
param(
    [Parameter(
        Mandatory=$True,
        HelpMessage="Array object"
    )]
    [ValAnimal4ateNotNullOrEmpty()]
    [object] $Array,

    [Parameter(
        Mandatory=$False,
        HelpMessage="Property name (default is root)"
    )]
    [ValAnimal4ateNotNullOrEmpty()]
    [string]$PropertyName

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
    $CurrentParams.Add("ParameterSetName",$Source)
    $CurrentParams.Add("PipelineInput",$PipelineInput)
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

    If ($CurrentParams.PropertyName) {
        If (($Array | Get-Member -MemberType NoteProperty).Name -notcontains $PropertyName) {
            $Msg = "Invalid property name $PropertyName"
            $Host.UI.WriteErrorLine("ERROR: $Msg")
            Break
        }
    }
   
    

}
Process {

    Write-Verbose "Converting object to hashtable and looking for duplicates"   
    $hash = @{}
    
    $array | Foreach-Object { 
        Write-Host $_
        $hash[$_] = $hash[$_] + 1 
    }
   

    $Array | Foreach-Object { 
        


        #Write-Host $_
        $hash[$_.$PropertyName] = $hash[$_] + 1 
    }
   
    $result = $hash.GetEnumerator() | Where-Object {$_.value -gt 1} | Foreach-Object {$_.key}
    
    Return $result
}
}


$hash.GetEnumerator() | Foreach-Object {
    $Curr = $_
    $Curr.Name
}




$Output | Select-Object -Property ServicePrincipalNames -Unique


    $hash = @{            
        
        Animal1    = "kitten"
        Animal2    = "puppy"
        Animal3      = "bunny"
        Animal4               = "snake"
    }                           
                                    
    $Object = New-Object PSObject -Property $hash     

    [pscustomobject]$Allmyobj = @()
      
      $allmyobjpart = New-Object PSObject -Property ([ordered] @{            
        
        Animal1 = "kitten"
        Animal2 = "puppy"
        Animal3 = "bunny"
        Animal4 = "snake"
    })
        $Allmyobj += $allmyobjpart

                                    
    $Object = New-Object PSObject -Property $hash 


    $Obj = New-Object PSObject -Property ([ordered] @{            
        
        Animal1 = "kitten"
        Animal2 = "puppy"
        Animal3 = "bunny"
        Snack   = "doritos","zinfandel"
    })

    $Array = @()
    $Array += New-Object PSObject -Property ([ordered] @{            
        Animal1 = "elephant"
        Animal2 = "giraffe"
        Animal3 = "lion"
        Snack   = "cupcake","milk"
    })

    $Array += New-Object PSObject -Property ([ordered] @{            
        
        Animal1 = "kitten"
        Animal2 = "puppy"
        Animal3 = "bunny"
        Snack   = "doritos","zinfandel"
    })

    $Array += New-Object PSObject -Property ([ordered] @{            
        
        Animal1 = "ocelot"
        Animal2 = "guppy"
        Animal3 = "zebra"
        Snack   = "apple","water"
    })
    $Array += New-Object PSObject -Property ([ordered] @{            
        
        Animal1 = "rhino"
        Animal2 = "leopard"
        Animal3 = "mole"
        Snack   = "apple","water","doritos","zinfandel"
    })

    $Array.GetType()
    $Array | GM
    $Array.PSObject.Properties.Name


    #$L = @()

    $L += $K




      $Array = @()
    $Array +=  ([pscustomobject][ordered] @{            
        Animal1 = "elephant"
        Animal2 = "giraffe"
        Animal3 = "lion"
        Snack   = "cupcake","milk"
    })


        $Allmyobj += $allmyobjpart


Function Get-JuneDuplicates {
<#
    .SYNOPSIS
        Gets duplicates or unique values in a collection.

    .DESCRIPTION
        The Get-Duplicates.ps1 script takes a collection and returns
        the duplicates (by default) or unique members (use the Unique
        switch parameter).

    .PARAMETER  Items
        Enter a collection of items. You can also pipe the items to
        Get-Duplicates.ps1.

    .PARAMETER  Unique
        Returns unique items instead of duplicates. By default, Get-Duplicates.ps1
        returns only duplicates.

    .EXAMPLE
        PS C:\> .\Get-Duplicates.ps1 -Items 1,2,3,2,4
        2

    .EXAMPLE
        PS C:\> 1,2,3,2,4 | .\Get-Duplicates.ps1
        2

    .EXAMPLE
        PS C:\> .\Get-Duplicates.ps1 -Items 1,2,3,2,4 -Unique
        1
        2
        3
        4

    .INPUTS
        System.Object[]

    .OUTPUTS
        System.Object[]

    .NOTES
    ===========================================================================
     Created with:     SAPIEN Technologies, Inc., PowerShell Studio 2014 v4.1.72
     Created on:       10/15/2014 9:34 AM
     Created by:       June Blender (juneb)
#>

param
(
    [Parameter(
        Mandatory = $true,
        ValueFromPipeline = $true
    )]
    [Object[]]$Object,
    
    $PropertyName,

    [Parameter(
        Mandatory = $false
    )]
    [Switch]$Unique
)
Begin{
    $hash = [ordered]@{ }
    $duplicates = @()
}
Process{
    
    # For each 
    foreach ($item in ($Objects | Sort $PropertyName)){
        
        Foreach ($SubItem in $Item.$PropertyName) {
            
            try {
                $hash.add($SubItem, 0)
            }
            catch [System.Management.Automation.MethodInvocationException]{
                $duplicates += $SubItem
            }
        }
        
    }
}
End{

    if ($unique){
        return $hash.keys
        
    }
    elseif ($duplicates){
        return $duplicates
    }
}


}ForEach($obj in $Array)
{ 
    if (compare-object $obj $mycfg -Property Projname,procfg)
    {Write-Host "Unequal"}
    else
    {Write-Host "Equal"}
}



$difference = compare-object $DuplicateSPNs.InputObject $Output.ServicePrincipalNames |
    where-Object sideindicator -eq => | 
        Select @{N='Servi';E={$_.InputObject}}




[array]$AllSPNList = ($Output | sort-object ServicePrincipalNames)
[array]$UniqueSPNs = ($AllSPNList.ServicePrincipalNames | Select -Unique)
[array]$DuplicateSPNs = Compare-Object -ReferenceObject $UniqueSPNs -DifferenceObject ($AllSPNList.ServicePrincipalNames)
[array]$Culprits = @()


$Output | Foreach-Object {
    
    $Curr = $_
    Write-host $Curr.Objectname -ForegroundColor Yellow

    $Found = 0
    Foreach ($SPN in $Curr.ServicePrincipalNames) {
        If ($DuplicateSPNs.InputObject -Contains $SPN) {
            $Found ++
        }
    }
    If ($Found -gt 0) {$Culprits += $Curr}

}

$Culprits | Sort ServicePrincipalNames