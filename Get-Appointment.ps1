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

    $outlook = New-Object -ComObject Outlook.Application

    # Ensure we are logged into a session
    $session = $outlook.Session
    $session.Logon()

    $olFolderCalendar = 9
    $apptItems = $session.GetDefaultFolder($olFolderCalendar).Items
    $apptItems.Sort('[Start]')
    $apptItems.IncludeRecurrences = $true

    $restriction = "[End] >= '{0}' AND [Start] <= '{1}'" -f $rangeStart.ToString('g'), $rangeEnd.ToString('g')
    Write-Debug -Message $restriction
    $Array = @()
    foreach($item in $apptItems.Restrict($restriction))
    {
        $NewObject = New-Object  -TypeName PSObject -Property @{
            Start                = $item.Start
            Duration             = $item.Duration
            End                  = $item.End
            Subject              = $item.Subject
            Body                 = $item.Body
            ReminderSet          = $item.ReminderSet
            Location             = $item.Location
            Categories           = $item.Categories
            CreationTime         = $item.CreationTime
            LastModificationTime = $item.LastModificationTime
            IsRecurring          = $item.IsRecurring
            ConversationIndex    = $item.ConversationIndex
            EntryID              = $item.EntryID
            GlobalAppointmentID  = $item.GlobalAppointmentID
        }
        $Array += $NewObject
    }
    Write-Output -InputObject $Array
    #$outlook = $session = $null
}
