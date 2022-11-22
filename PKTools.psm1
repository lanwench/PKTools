# PKTools

$ModuleName = "PKTools"
$Activity = "Importing module $ModuleName"

If (Test-Path $PSScriptRoot\Scripts -ErrorAction SilentlyContinue){
    
    <#
    
    $Current = 0
    [object[]]$Functions = Get-ChildItem -Path (Join-Path $PSScriptRoot Scripts) -Filter function_*.ps1 -Recurse 
    $Functions | Sort-Object FullName | Foreach-Object {
        $Current ++
        Write-Verbose $_.FullName
        Write-Progress -Activity "Loading $ModuleName functions" -CurrentOperation $_.FullName -PercentComplete ($Current/$Functions.Count*100)
        
        . $_.FullName
    }

    #>

   











    Try {
        [object[]]$ScriptFiles = @( Get-ChildItem -Path $PSScriptRoot\Scripts -Filter "function_*.ps1" -ErrorAction SilentlyContinue )
        
        $Total = $ScriptFiles.Count
        $Current = 0
        
        Foreach($import in ($ScriptFiles | Sort-Object FullName)){
            $Current ++
            Write-Verbose "[$ModuleName] $($Import.FullName)"
            Try {
                Write-Progress -Activity "Importing module PKTools" -CurrentOperation $Import.Fullname -PercentComplete ($Current/$Total*100)
                If ($Import.FullName -match "session") {
                    write-host "beeeep" -ForegroundColor Red
                }
                . $import.fullname
            }
            Catch {
                Write-Error -Message "Failed to import function $($import.fullname): $_"
            }
        }
    }
    Catch {}
    Finally {
        Write-Progress -Activity $Activity -Completed
    }



    #>
}
