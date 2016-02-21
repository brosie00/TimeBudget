#requires -Version 2


<#
        .Synopsis
        Renames Appointments in a Calendar that have been copied from another calendar and the subject line begins with 'Copy:'.
        .DESCRIPTION
        Used mainly as helper function this module and called by New-OPMAppointment. Restrict to a single Category. Select all items
        within that CALENDAR and CATEGORY. Modify SUBJECT line.

#>
function Reset-TimeBudgetSubject {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(ValueFromPipelineByPropertyName = $true,
        Position = 0)]
        $Categories = @('Test', 'Project27') 
        ,

        # Param2 help description
        [string[]]
        $Calendar = @('Step', 'Phase')
    )

    #Namespace is a global variable defined when the module was imported

    foreach (  $SingleCalendar in $Calendar ) 
    {
        #I think this is just for a non-Exchange system
        $CalendarComObject = $Namespace.Folders.Item($NamespaceFolderItemTitle).Folders | Where-Object -FilterScript { $_.Name -ieq $SingleCalendar }
        
        
        foreach ( $Category in $Categories ) {
            Write-Debug -Message $CalendarComObject.Name | Out-String -Debug
            $Restriction = "[Categories] = $Categories"
            $apptItems = $CalendarComObject.Items
        
            foreach ( $Item in  $apptItems.Restrict($Restriction) ) 
            {
                Write-Debug -Message "Looping through item `'$($Item.Subject)`' within the $Categories Category  within the $SingleCalendar Calendar"
                
                Write-Verbose -Message "$Item.Subject will be changed to (  $Item.Subject -replace 'Copy:', '' ).TrimStart()"
                $Item.Subject = (  $Item.Subject -replace 'Copy:', '' ).TrimStart()
                $Item.Save()
            }
        }
    } #end Calendar foreach
}