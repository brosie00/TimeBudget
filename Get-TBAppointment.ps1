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

        # Param2 help description
        [string[]]
        $Calendar = @('Calendar', 'Step', 'Phase')
        #   $Calendar = 'Step'   
    )
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

            $RestrictedItemsCount = $apptItems.Restrict($Restriction).Count | Out-String

            foreach ( $Item in  $apptItems.Restrict($Restriction) ) 
            {
                Write-Verbose 'About to create an object'
                $obj = New-Object -TypeName PsObject -Property @{
                    Calendar              = $SingleCalendar
                    Start                 = $Item.Start
                    End                   = $Item.End   
                    Duration              = $Item.Duration    
                    Days                  = (New-TimeSpan -Start $Item.Start -End $Item.End).Days   
                    Weeks                 = (New-TimeSpan -Start $Item.Start -End $Item.End).Days /7  
                    Subject               = $Item.Subject               
                    Body                  = $Item.Body                  
                    ReminderSet           = $Item.ReminderSet           
                    Location              = $Item.Location              
                    Categories            = $Item.Categories            
                    CreationTime          = $Item.CreationTime          
                    LastModificationTime  = $Item.LastModificationTime  
                    IsRecurring           = $Item.IsRecurring           
                    Deliverable           = $Item.UserProperties.Item('Deliverable').Value
                    Reference             = $Item.UserProperties.Item('Reference').value
                    GUID                  = $Item.UserProperties.Item('GUID').Value
                    SortNumber            = $Item.UserProperties.Item('Sorting Number').Value
                }#end Object 
                Write-Verbose -Message $obj.Subject
                $Appointments += $obj
            }#end foreach Item in apptItems      
        }#end foreach Category in Categories
    }#end foreach SingleCalendar in Calendar
   
#    $Appointments = $Appointments | Sort-Object -Unique -Property 'EntryID'
    Write-Output -InputObject $Appointments
}


