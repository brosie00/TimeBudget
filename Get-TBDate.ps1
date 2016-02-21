function Get-TBDate {
    #requires -Version 1

    $StartDate = (Get-Date).AddDays(0)
    $EndDate = (Get-Date).AddDays(7)

    <#
    $a = for ($x = $StartDate ; $x -lt $EndDate ; $x = ($x).AddDays(1) )
    {
        for ($H = 5  ; $H -lt 23 ; $H += 1 ) 
        {
            for ($m = 0; $m -lt 59; $m += 15 ) 
            {$x | Get-Date  -Hour $H -Minute $m}
        }
    }

 
    $c = $a   | Sort-Object -Property Day,Hour, Minute  -Unique

    Write-Output -InputObject $c
   
} #>

for ($x = $StartDate ; $x -lt $EndDate ; $x = ($x).AddDays(1) )
    {
       $x | Get-Date -Minute 0 -Second 0
       
    }

} 



