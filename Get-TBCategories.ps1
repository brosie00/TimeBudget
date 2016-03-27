

#requires -Version 1
function Get-TBCategories {
    
    $Output = $Namespace.Categories | Select-Object -ExpandProperty Name

    Write-Output -InputObject $Output
}

#>