#requires -Version 2
<#

#>
function New-TBDeadline
{
    [cmdletbinding(DefaultParameterSetName = 'End_Minutes' )]
    param (

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [datetime]
        $Start = (Get-Date).AddHours(1)
        , 

        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [Alias('Reminder')]
        [int32]
        $ReminderMinutes = 20 
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
        $Categories 
        ,
        
        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [switch]
        $ShowAppointment = $False
        ,
        
        [Parameter(ValueFromPipelineByPropertyName = $True )]
        [string]
        $Calendar = 'Calendar'
    )



    if ( $PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent -eq $True ) 
    {
        $DebugBoundParameters = New-Object -TypeName PsObject -Property $PSCmdlet.MyInvocation.BoundParameters
         
        Write-Verbose -Message $( $PSCmdlet.MyInvocation.BoundParameters).Keys
    }

    Write-Debug -Message "Start = $Start"
    Write-Debug -Message "Duration = $Duration"
    Write-Debug -Message "End      = $End"

    $CalendarComObject = $Namespace.Folders.Item($NamespaceFolderItemTitle).Folders | Where-Object -FilterScript { $_.Name -ieq $Calendar }
      
    $objAppointment = $CalendarComObject.Items.Add($olAppointmentItem)
        
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

    if ( $ReminderMinutes ) {  
        $objAppointment.ReminderSet = $True
        $objAppointment.ReminderMinutesBeforeStart = $ReminderMinutes
    }

    $objAppointment.Start = $Start
    $objAppointment.End = $Start

    $null = $objAppointment.Send()
    $null = $objAppointment.Save()

    Write-Verbose -Message $objAppointment.Subject
    Write-Verbose -Message $objAppointment.Start
    Write-Verbose -Message $objAppointment.End

    if ($ShowAppointment) 
    {$objAppointment.Display($True)}

    Write-Verbose -Message $objAppointment.Subject
    Write-Verbose -Message $objAppointment.Start
    Write-Verbose -Message $objAppointment.End
}


