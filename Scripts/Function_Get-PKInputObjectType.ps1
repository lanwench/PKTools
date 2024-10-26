

#requires -version 4
Function Get-PKInputObjectType {
    <#
    .SYNOPSIS
        Uses regex to check the type of the input object, in friendly and full name/type formats.

    .DESCRIPTION
        The Get-PKInputObjectType function determines and returns the type of the provided input object. 
        This is a very silly little function, but it can be useful for debugging or logging purposes where knowing the type of an object is necessary.
        It can recognize a variety of types, including AD objects, strings, integers, and file system objects.

    .NOTES
        Name    : Function_Get-PKInputObjectType.ps1
        Author  : Paula Kingsley
        Created : 2024-10-25
        Version : 01.00.0000
        History :
    
        ** PLEASE KEEP $VERSION UPDATED IN BEGIN  BLOCK! **
    
        v01.00.0000 - 2024-10-25 - Created script

        
    .PARAMETER InputObject
        The object whose type needs to be determined.

    .EXAMPLE
        PS C:\> "12345",54321,(Get-ADUser jbloggs),(Get-Item C:\temp\2024-05-14_Groups.csv) | Get-PKInputObjectType -Verbose
        Returns a friendly name as well as the name, fullname, and basetype for input objects in the pipeline


    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]$Identity
    )
    
    Begin {

        # Current version (please keep up to date from comment block)
        [version]$Version = "01.00.0000"

        # Show our settings
        $ScriptName = $MyInvocation.MyCommand.Name
        $CurrentParams = $PSBoundParameters
        $Source = $PSCmdlet.ParameterSetName
        [switch]$PipelineInput = $MyInvocation.ExpectingInput
        $MyInvocation.MyCommand.Parameters.keys | Where-Object {$CurrentParams.keys -notContains $_} |
            Where-Object {Test-Path variable:$_}| ForEach-Object {
                $CurrentParams.Add($_, (Get-Variable $_).value)
            }
        $CurrentParams.Add("ParameterSetName",$Source)
        $CurrentParams.Add("PipelineInput",$PipelineInput)
        $CurrentParams.Add("ScriptName",$Scriptname)
        $CurrentParams.Add("Version",$Version)
        Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"
        

        Write-Verbose "[BEGIN: $Scriptname] Identifying certain types of trees from very, very far away..."
    }
    
    Process {

        Foreach ($i in $Identity) {
    
            switch -regex ($i.GetType().Name) {
                "ADObject"                { $Type = "AD Object" }
                "ADUser"                  { $Type = "AD user" }
                "ADGroup"                 { $Type = "AD group" }
                "ADComputer"              { $Type = "AD computer" }
                "ADOrganizationalUnit"    { $Type = "AD Organizational Unit" }
                "ADDomainController"      { $Type = "AD Domain Controller" }
                "ActiveDirectorySecurity" { $Type = "AD object ACL" }
                "String"                  { $Type = "String" }
                "Int16"                   { $Type = "16-bit integer" }
                "Int32"                   { $Type = "32-bit integer" }
                "Int64"                   { $Type = "64-bit integer" }
                "UInt16"                  { $Type = "16-bit unsigned integer" }
                "UInt32"                  { $Type = "32-bit unsigned integer" }
                "UInt64"                  { $Type = "64-bit unsigned integer" }
                "Single"                  { $Type = "Single" }
                "Double"                  { $Type = "Double" }
                "FileInfo"                { $Type = "File" }
                "DirectoryInfo"           { $Type = "Directory" }
                "FileSystemSecurity"      { $Type = "File ACL" }
                "DirectorySecurity"       { $Type = "Directory ACL" }
                default                   { $Type = "Not really sure what to call this?" }
            }
            
            Write-Verbose "[$i] Input object is a $Type!"

            [PSCustomObject]@{
                InputObject = $i
                WhatIsIt    = $Type
                Name        = $Identity.GetType().Name
                FullName    = $Identity.GetType().FullName.ToString()
                BaseType    = $Identity.GetType().BaseType.Name
            }
    }
}

} #end Get-PKADInputObjectType


