#requires -Version 2 -Modules TimeBudget
function Get-TBAppointment 
{
    [cmdletbinding()]
    param ( 

        [Parameter(ValueFromPipelineByPropertyName = $true,
        Position = 0)]
        [string[]]
        $Categories = ( Get-TBCategories )
        #   $Categories = 'Test' 
        ,

        [string[]]
        $Calendar = @('Calendar', 'Step', 'Phase')
        #   $Calendar = 'Step'   

    )


    try
    {$Namespace.Folders.Item($NamespaceFolderItemTitle).Folders}
    catch
    {
        Write-Host -ForegroundColor Red -Object 'The Com Object with Microsoft Outlook has broken. We will attempt to reimport the Module'
        Import-Module -Name TimeBudget -Force
    } 
   


    $Appointments = @() 
    
    foreach ( $SingleCalendar in $Calendar )# { $SingleCalendar }
    {
        #$SingleCalendar = $Calendar
        Write-Debug -Message "Calendar: $SingleCalendar"
        #$CalendarComObject = $Namespace.Folders.Item('1').Folders | Where-Object -FilterScript { $_.Name -ieq $SingleCalendar }
        # see https://msdn.microsoft.com/en-us/magazine/dn189202.aspx for a discussion of navigating in the MAPI namespace
        $CalendarComObject = $Namespace.Folders.Item($NamespaceFolderItemTitle).Folders.Item($SingleCalendar) 
 
        foreach ( $Category in $Categories  ) 
        {   
            $Restriction = "[Categories] = $Category"
            $apptItems = $CalendarComObject.Items

            foreach ( $Item in  $apptItems.Restrict($Restriction) ) 
            {
                $Props = @{}
                $Props.Calendar             = $SingleCalendar
                $Props.Start                = $Item.Start
                $Props.End                  = $Item.End
                $Props.Duration             = $Item.Duration
                $Props.Days                 = (New-TimeSpan -Start $Item.Start -End $Item.End).Days
                $Props.Weeks                = (New-TimeSpan -Start $Item.Start -End $Item.End).Days /7
                $Props.Subject              = $Item.Subject
                $Props.Body                 = $Item.Body
                $Props.ReminderSet          = $Item.ReminderSet
                $Props.Location             = $Item.Location
                $Props.Categories           = $Item.Categories
                $Props.CreationTime         = $Item.CreationTime
                $Props.LastModificationTime = $Item.LastModificationTime
                $Props.IsRecurring          = $Item.IsRecurring
                $Props.Deliverable          = $Item.UserProperties.Item('Deliverable').Value
                $Props.Reference            = $Item.UserProperties.Item('Reference').value
                $Props.GUID                 = $Item.UserProperties.Item('GUID').Value
                $Props.SortNumber           = $Item.UserProperties.Item('Sorting Number').Value

                New-Object -TypeName PsObject -Property $Props
            } #end foreach ( item in apptItems )
         
            $Appointments += $obj
        }#end foreach (Category in Categories)
    }#end foreach (SingleCalendar in Calendar)
   
    Write-Output -InputObject $Appointments
}


