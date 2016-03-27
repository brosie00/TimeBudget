
#requires -Version 2
function Get-Appointment 
{
    [cmdletbinding()]
    param ( 
 
        [Parameter(ParameterSetName = 'Start_End')]
        [Parameter(ParameterSetName = 'Days_Start')]
        [datetime]$StartDate = (Get-Date)
        ,
        [Parameter(ParameterSetName = 'Start_End')]
        [Parameter(ParameterSetName = 'Days_End')]
        [datetime]$EndDate = (Get-Date)
        ,
        [Parameter(ParameterSetName = 'Days_Start')]
        [Parameter(ParameterSetName = 'Days_End')]
        [int]$Days 
    )

    try
    {
        $Namespace.Folders.Item($NamespaceFolderItemTitle).Folders
    }
    catch
    {
        Write-Host -ForegroundColor Red -Object 'The Com Object with Microsoft Outlook has broken. We will attempt to reimport the Module'
        Import-Module -Name TimeBudget -Force
    } 

    switch ($PSCmdlet.ParameterSetName) 
    { 
        'Start_End' 
        {
            $rangeStart = $StartDate | Get-Date -Hour 0 -Minute 0
            $rangeEnd   = $EndDate   | Get-Date -Hour 11 -Minute 59
            break
        } 
        'Days_Start'     
        {
            $rangeStart = $StartDate | Get-Date -Hour 0 -Minute 0
            $rangeEnd   = $rangeStart.AddDays($Days) | Get-Date -Hour 11 -Minute 59
            break
        } 
        'Days_End'   
        {
            $rangeEnd = $EndDate | Get-Date -Hour 11 -Minute 59
            $rangeStart = $EndDate.AddDays(-$Days) | Get-Date -Hour 0 -Minute 00
            break
        } 
    }

   

    $apptItems.Sort('[Start]')
    $apptItems.IncludeRecurrences = $true

    $restriction = "[End] >= '{0}' AND [Start] <= '{1}'" -f $rangeStart.ToString('g'), $rangeEnd.ToString('g')

    $Array = @()
    foreach($item in $apptItems.Restrict($restriction))
    {  
        $Props = @{}
        $Props.Start                = $item.Start
        $Props.Duration             = $item.Duration
        $Props.End                  = $item.End
        $Props.Subject              = $item.Subject
        $Props.Body                 = $item.Body
        $Props.ReminderSet          = $item.ReminderSet
        $Props.Location             = $item.Location
        $Props.Categories           = $item.Categories
        $Props.CreationTime         = $item.CreationTime
        $Props.LastModificationTime = $item.LastModificationTime
        $Props.IsRecurring          = $item.IsRecurring
        $Props.ConversationIndex    = $item.ConversationIndex
        $Props.EntryID              = $item.EntryID
        $Props.GlobalAppointmentID  = $item.GlobalAppointmentID

        $Obj = New-Object -TypeName PsObject -Property $Props
    }#foreach (item in apptitems)

    $Array += $Obj 

Write-Output -InputObject $Array

}
#>