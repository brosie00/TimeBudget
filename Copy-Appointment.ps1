#requires -Version 2
function Copy-Appointment 
{  
    [cmdletbinding()]
    param(

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [Parameter(ParameterSetName = 'Pattern')]
    
        $Pattern
        ,

        [Parameter(ParameterSetName = 'Subject')]
        $Subject 
        ,

        [datetime]
        $Start = (Get-Date).addminutes(-30)
        ,

        [switch]
        $ShowAppointment = $False

    )
    #region  All Potential Matches Snippet
       
    try
    {
        $Namespace.Folders.Item($NamespaceFolderItemTitle).Folders
    }
    catch
    {
        Write-Host -ForegroundColor Red -Object 'The Com Object with Microsoft Outlook has broken. We will attempt to reimport the Module'
        Import-Module -Name TimeBudget -Force
    } 
   
   

    [array]$AllPossibleMatches = $TimeBudgetCache | Where-Object -FilterScript { $_.Subject -match $Pattern } 
       
    try
    {
        #If no possible matches were found, lets figure this out now
        $AllPossibleMatches[0]
    }
    catch
    {Write-Warning -Message "Explain to the user that you couldn't find something that matched the pattern. Tell them what to change in order to expand the search"}
       
    if ( $AllPossibleMatches.Count -gt 1 )
    {
        #The first synthetic property is needed to display a later Index Choice; the other properties are situational dependent
        $AllPossibleMatches |
        Select-Object -Property @{Name = 'Index';Expression = { [array]::IndexOf($AllPossibleMatches, $_)}}, Subject, Duration, Categories, @{ Name = 'Module'; Expression = { Split-Path -Path $_.Directory -Leaf } } |
        Format-Table -AutoSize
       
        [string]$SelectedIndexNumbers = Read-Host -Prompt 'Please Choose the Index Number of your Selected Appointment'
       
        [int32[]]$SelectedSubjects = $SelectedIndexNumbers -split { $_ -eq ' ' -or $_ -eq ',' } |
        Where-Object -FilterScript { $_ } |
        ForEach-Object -Process { $_.trim() } 
   
        foreach ( $Selection in $SelectedSubjects )
        {[array]$ChosenSubjects += $AllPossibleMatches[$Selection]}
    } 
    else 
    {
        [int32]$Selection = 0
        [array]$ChosenSubjects = $AllPossibleMatches[$Selection]
    }
   
    Write-Output -InputObject $ChosenSubjects 
   
    #endregion Snippet
       
    foreach ( $Choice in $ChosenSubjects ) 
    {
        $Properties = @{}
        $Properties.Start = $Start
        $Properties.Subject = $Choice.Subject 
        $Properties.Minutes = $Choice.Duration
        $Properties.Categories = @($Choice.Categories) 
        $Properties.Location = $Choice.Location
        $Properties.Body = $Choice.Body
        $Properties.ReminderMinutes = $Choice.ReminderSet
        $Properties.ShowAppointment = $ShowAppointment

        New-Appointment @Properties
    }
}

#>