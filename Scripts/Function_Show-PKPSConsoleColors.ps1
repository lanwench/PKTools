<#
.SYNOPSIS
    Displays console colors, with option to show as a grid with different foreground/background combinations

.DESCRIPTION
    Displays console colors, with option to show as a grid with different foreground/background combinations
    Uses Write-Host. Hush. It needs to.
    
.NOTES
    Name    : Function_Show-PKPSConsoleColors.ps1
    Author  : Paula Kingsley
    Created : 2019-06-10
    Version : 01.00.0000    
    History :

        *** PLEASE KEEP $VERSION UP TO DATE IN BEGIN BLOCK **

        v01.00.0000 - 2019-06-10 - Created script based on Emperor XLII & Tim Abell's StackOverflow answers

.LINK
    https://stackoverflow.com/questions/20541456/list-of-all-colors-available-for-powershell

.LINK
    https://gist.github.com/timabell/cc9ca76964b59b2a54e91bda3665499e

.EXAMPLE
    PS C:\> Show-PKPSConsoleColors -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                 
        ---           -----                 
        Verbose       True                  
        AsGrid        False                 
        ScriptName    Show-PKPSConsoleColors
        ScriptVersion 1.0.0                 

        VERBOSE: Display console colors

          0        Black Black
          1     DarkBlue DarkBlue
          2    DarkGreen DarkGreen
          3     DarkCyan DarkCyan
          4      DarkRed DarkRed
          5  DarkMagenta DarkMagenta
          6   DarkYellow DarkYellow
          7         Gray Gray
          8     DarkGray DarkGray
          9         Blue Blue
         10        Green Green
         11         Cyan Cyan
         12          Red Red
         13      Magenta Magenta
         14       Yellow Yellow
         15        White White

.EXAMPLE
    PS C:\> Show-PKPSConsoleColors -AsGrid -Verbose

        VERBOSE: PSBoundParameters: 
	
        Key           Value                 
        ---           -----                 
        AsGrid        True                  
        Verbose       True                  
        ScriptName    Show-PKPSConsoleColors
        ScriptVersion 1.0.0                 

        VERBOSE: Display console colors as a grid showing different foreground colors

        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Black
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on DarkBlue
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on DarkGreen
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on DarkCyan
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on DarkRed
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on DarkMagenta
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on DarkYellow
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Gray
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on DarkGray
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Blue
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Green
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Cyan
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Red
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Magenta
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on Yellow
        Black|DarkBlue|DarkGreen|DarkCyan|DarkRed|DarkMagenta|DarkYellow|Gray|DarkGray|Blue|Green|Cyan|Red|Magenta|Yellow|White| on White


#>
Function Show-PKPSConsoleColors {
[Cmdletbinding()]
Param(
    [Parameter(
        HelpMessage = "Display console colors as a grid showing different foreground colors"
    )]
    [switch]$AsGrid)

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
    If ($PipelineInput.IsPresent) {$CurrentParams.InputObject = $Null}
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

}
Process {
    
    If ($AsGrid.IsPresent) {

        $Msg = "Display console colors as a grid showing different foreground colors"
        Write-Verbose "$Msg`n"

        $colors = [enum]::GetValues([System.ConsoleColor])
        Foreach ($bgcolor in $colors){
            Foreach ($fgcolor in $colors) { 
                Write-Host "$fgcolor|"  -ForegroundColor $fgcolor -BackgroundColor $bgcolor -NoNewLine 
            }
            Write-Host " on $bgcolor"
        }
    }

    Else {

        $Msg = "Display console colors"
        Write-Verbose "$Msg`n"
        
        $colors = [Enum]::GetValues( [ConsoleColor] )
        $max = ($colors | foreach { "$_ ".Length } | Measure-Object -Maximum).Maximum
        foreach( $color in $colors ) {
            Write-Host (" {0,2} {1,$max} " -f [int]$color,$color) -NoNewline
            Write-Host "$color" -Foreground $color
        }
    }
}
End {}
} #End Show-PKPSConsoleColors