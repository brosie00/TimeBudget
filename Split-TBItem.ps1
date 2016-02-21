#requires -Version 2

<#
        .Synopsis
        Short description
        .DESCRIPTION
        Long description
        .EXAMPLE
        Example of how to use this cmdlet
        .EXAMPLE
        Another example of how to use this cmdlet
#>
function Split-TBItem
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        $AppointmentItems
    )

    foreach($item in $AppointmentItems)
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
