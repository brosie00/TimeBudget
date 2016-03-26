#requires -Version 2
<#
        .Synopsis
        Builds an Outlook appointment or meeting from the powershell command line.
        .DESCRIPTION
        This creates a new Appointment from the command line and displays it in the Outlook application after the item is created.  The display action is
        a parameter and can be turned off.

        Outlook appointments and meetings are the same Outlook type. Appointments become meeting once someone is invited to the invite list. 
        
        I use this tool in conjunction with my Outlook account to contribute to a journal of my workday, without needing to shift my concentration to 
        Outlook.

        Should be a good case for use with about_Parameters_Default_Values. ( Use your home email address as a recipient, and keep your calendars synced )

        .EXAMPLE
        New-Appointment -Start '8:30 am' -AllDayEvent -Subject 'Do Something' -Recipients Alice@acme.com

        .EXAMPLE
        New-Appointment -Start '10/5/2015 8:30 am' -AllDayEvent -Subject 'Do Something' -Recipients Alice@acme.com

        .EXAMPLE
        New-Appointment -End ( Get-Date ) -Minutes ( -24 ) -Subject 'Just Finished This Random Assignment'

        .EXAMPLE
        New-Appointment
        
        When run with no parameters, creates a new appointment starting and ending at the time the cmdlet was run.  The ShowAppointment
        defaults to true, and the Outlook windows pops up allowing changes.

       
        .NOTES
        
        The meeting BusyStatus is set ( without a parameter ) to 0 (Free) .  Other options are 
        {1=Tentative; 2 = Busy; 3 = Out of Office;  4 = Working Elsewhere;} 

        Assumes  the following variables were created when the module was imported
        $global:olAppointmentItem = 1
        $global:Outlook = New-Object -ComObject outlook.application 
        $global:Namespace = $Outlook.GetNamespace('MAPI')
#>

function New-Appointment 
{
    [cmdletbinding(DefaultParameterSetName = 'End_Minutes' )]
    param (

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [Parameter(ParameterSetName = 'End_Start')]
        [Parameter(ParameterSetName = 'Start_Minutes')]
        [Parameter(ParameterSetName = 'Start_AllDayEvent')]
        [datetime]
        $Start = (Get-Date)
        , 

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [Parameter(ParameterSetName = 'End_Start'        )]
        [Parameter(ParameterSetName = 'End_Minutes'      )]
        [datetime]
        $End = (Get-Date).AddMinutes(30)  
        ,  

        [Parameter(ParameterSetName = 'End_Minutes')]
        [Parameter(ParameterSetName = 'Start_Minutes')]
        [Alias('Duration')]
        [int32]
        $Minutes  = (Get-Date).Minute
        ,

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [Alias('Reminder')]
        [int32]
        $ReminderMinutes = 20 
        , 

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [Parameter(ParameterSetName = 'Start_AllDayEvent' )]
        [switch]
        $AllDayEvent   
        ,

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [string]
        $Subject = 'This is the Subject'
        ,
        
        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [string[]]
        $Body = 'This is the Body'
        ,

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [string[]]
        $Recipients = 'Brett.Osiewicz@irs.gov'
        ,

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [string]
        $Location
        ,

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [string[]]
        $Categories = @('Daily')
        ,

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [switch]
        $ShowAppointment
        ,

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [string]
        $Calendar = 'Calendar'
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

    $CalendarComObject = $Namespace.Folders.Item($NamespaceFolderItemTitle).Folders | Where-Object -FilterScript { $_.Name -ieq $Calendar }
    $global:objAppointment = $CalendarComObject.Items.Add($olAppointmentItem)

    #$objAppointment = $Outlook.CreateItem($olAppointmentItem) 
    $objAppointment.BusyStatus  = '0' 
    $objAppointment.Subject     = $Subject
    $objAppointment.Body        = $Body
    $objAppointment.Location    = $Location
    $objAppointment.Categories  = $Categories
  
    foreach ($Name in $Recipients ) 
    {
        $objAppointment.Recipients.Add($Name)
        $objAppointment.MeetingStatus = '1'
    }
        
    $null = $objAppointment.Recipients.ResolveAll()

    if ( $ReminderMinutes ) 
    {  
        $objAppointment.ReminderSet = $True
        $objAppointment.ReminderMinutesBeforeStart = $ReminderMinutes
    }

    switch ($PSCmdlet.ParameterSetName) 
    { 
        'End_Start'     
        {
            $objAppointment.Start = $Start
            $objAppointment.End   = $End
                
            break
        } 
        'End_Minutes'   
        {
            if ( $Minutes -lt 0 ) { $Start = $End.AddMinutes( $Minutes) }
            if ( $Minutes -eq 0 ) { $Start = $End                       }
            if ( $Minutes -gt 0 ) { $Start = $End.AddMinutes(-$Minutes) }

            $objAppointment.Start = $Start
            $objAppointment.End   = $End

            break
        }
        'Start_Minutes' 
        {
            if ( $Minutes -lt 0 ) {$End = $Start.AddMinutes(-$Minutes)}
            if ( $Minutes -eq 0 ) {$End = $Start}
            if ( $Minutes -gt 0 ) {$End = $Start.AddMinutes($Minutes)}

            $objAppointment.Start = $Start
            $objAppointment.End   = $End

            break
        } 
        'Start_AllDayEvent' 
        {
            $objAppointment.Start       = $Start
            $objAppointment.AllDayEvent = $AllDayEvent

            break
        }
    } 

   Write-Verbose $($objAppointment | Select-Object -Property Start, End, Duration, ReminderMinutes, AllDayEvent, Subject, Body, RequiredAttendees, Location, Categories | Format-List | Out-String )

    $null = $objAppointment.Send()
    $null = $objAppointment.Save()

    if ($ShowAppointment) 
    {$objAppointment.Display($True)}
}


