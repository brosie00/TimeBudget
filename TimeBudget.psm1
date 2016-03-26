Get-ChildItem $psscriptroot\*.ps1 | foreach { . $_.FullName }



#Remove-Variable -Name NamespaceFolderItemTitle
        $global:olAppointmentItem = 1
        $global:Outlook = New-Object -ComObject outlook.application 
        $global:Namespace = $Outlook.GetNamespace('MAPI')
        
        #The different versions of Outlook may use the 
        #Namespace component differently, so i need to specify differently for different versions of outlook. '



#  $TimeBudgetCache =  Get-TBAppointment 

switch ($env:COMPUTERNAME)
{
     'DCW139MA4228116'  { $global:NamespaceFolderItemTitle = 'brett.osiewicz@irs.gov' } #used with the IRS version of Outlook (version Office 2010)
     'Surface'          { $global:NamespaceFolderItemTitle = '1' }
     'DESKTOP-HBB93I3'  { $global:NamespaceFolderItemTitle = '1' }
     'WORKSTATION2'     { $global:NamespaceFolderItemTitle = '1' }
     'SURFACEPROFOUR'   { $global:NamespaceFolderItemTitle = '1' }
}
 
        
if (!( $NamespaceFolderItemTitle )) {
            Write-Warning "Please configure the switch code in the ..\TimeBudget.psm1 file to include your 
            ComputerName.  I can't figure out why this is necesary."
        }

 Write-Warning -Message 'The Outlook module has been imported.'