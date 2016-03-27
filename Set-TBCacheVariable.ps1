

function Set-TBCacheVariable
{  
    [cmdletbinding()]

    Param(
        [int32]
        $Days = 90 

    )
    $Time = Measure-Command -Expression {
        $global:TimeBudgetCache = Get-Appointment -StartDate (Get-Date).AddDays(-$Days) -EndDate (Get-Date).AddDays(5)
    }
   
    
    if (  $PSBoundParameters.ContainsKey('Verbose') ) 
    {
        Write-Verbose -Message "$Time"
    }
}
#>