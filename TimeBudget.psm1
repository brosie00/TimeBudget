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
 
 
function New-CompletionResultTest
{
    param([Parameter(Position = 0, ValueFromPipelineByPropertyName, Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CompletionText,

        [Parameter(Position = 1, ValueFromPipelineByPropertyName)]
        [string]
        $ToolTip,

        [string]
        $ListItemText = $CompletionText,

        [System.Management.Automation.CompletionResultType]
    $CompletionResultType = [System.Management.Automation.CompletionResultType]::ParameterValue)
    
    if ($ToolTip -eq '')
    {$ToolTip = $CompletionText}

    if ($CompletionResultType -eq [System.Management.Automation.CompletionResultType]::ParameterValue)
    {
        # Add single quotes for the caller in case they are needed.
        # We use the parser to robustly determine how it will treat
        # the argument.  If we end up with too many tokens, or if
        # the parser found something expandable in the results, we
        # know quotes are needed.

        $tokens = $null
        $null = [System.Management.Automation.Language.Parser]::ParseInput("echo $CompletionText", [ref]$tokens, [ref]$null)
        if ($tokens.Length -ne 3 -or
            ($tokens[1] -is [System.Management.Automation.Language.StringExpandableToken] -and
        $tokens[1].Kind -eq [System.Management.Automation.Language.TokenKind]::Generic))
        {$CompletionText = "'$CompletionText'"}
    }
    return New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList ($CompletionText, $ListItemText, $CompletionResultType, $ToolTip.Trim())
}


Microsoft.PowerShell.Core\Register-ArgumentCompleter -Verbose -CommandName New-TBAppointment -ParameterName Categories -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    
    Get-TBCategories |
    ForEach-Object -Process {  New-CompletionResultTest -CompletionText $_ -ToolTip $_ -ListItemText $_ -CompletionResultType ParameterValue }
}

 
        
        if (!( $NamespaceFolderItemTitle )) {
            Write-Warning "Please configure the switch code in the ..\TimeBudget.psm1 file to include your 
            ComputerName.  I can't figure out why this is necesary."
        }

write-warning -Message 'The Outlook module has been imported.'

Microsoft.PowerShell.Core\Register-ArgumentCompleter -Verbose -CommandName New-TBAppointment -ParameterName Categories -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    
    Get-TBCategories |
    ForEach-Object -Process {  New-CompletionResultTest -CompletionText $_ -ToolTip $_ -ListItemText $_ -CompletionResultType ParameterValue }
}

Microsoft.PowerShell.Core\Register-ArgumentCompleter -Verbose -CommandName New-TBAppointment -ParameterName End -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    
    Get-TBDate |
    ForEach-Object -Process {  New-CompletionResultTest -CompletionText ( Get-Date -Date $_ -Format 'ddd' )  -ToolTip $_  -ListItemText ( Get-Date -Date $_ -Format 'ddd HH:mm') -CompletionResultType ParameterValue -Verbose }
}

Microsoft.PowerShell.Core\Register-ArgumentCompleter -Verbose -CommandName @( 'New-TBAppointment', 'New-TBDeadline') -ParameterName Start -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    
    Get-TBDate |
    ForEach-Object -Process {  New-CompletionResultTest -CompletionText ( Get-Date -Date $_ -Format 'ddd HH:mm')  -ToolTip $_  -ListItemText ( Get-Date -Date $_ -Format 'ddd') -CompletionResultType ParameterValue -Verbose }
}

Microsoft.PowerShell.Core\Register-ArgumentCompleter -Verbose -CommandName @( 'New-TBAppointment', 'New-TBDeadline') -ParameterName Categories -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    
    $TimeBudgetCache | 
    Sort-Object -Property Categories -Unique | #
    Where-Object { $_.Categories } |
    ForEach-Object -Process {  New-CompletionResultTest -CompletionText $_.Categories -ToolTip $_.Subject -ListItemText $_.Categories -CompletionResultType ParameterValue -Verbose }
}

Microsoft.PowerShell.Core\Register-ArgumentCompleter -Verbose -CommandName @( 'New-TBAppointment', 'New-TBDeadline') -ParameterName Subject -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    
    $TimeBudgetCache | 
    Where-Object { $_.Subject } |
    Sort-Object -Property Subject -Unique |
    ForEach-Object -Process {  New-CompletionResultTest -CompletionText $_.Subject -ToolTip $_.Start -ListItemText $_.Subject -CompletionResultType ParameterValue -Verbose }
}

Microsoft.PowerShell.Core\Register-ArgumentCompleter -Verbose -CommandName @( 'New-TBAppointment', 'New-TBDeadline') -ParameterName Location -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    
    $TimeBudgetCache | 
    Sort-Object -Property Location -Unique |
    Where-Object { $_.Location } |
    ForEach-Object -Process {  New-CompletionResultTest -CompletionText $_.Location -ToolTip $_.Subject -ListItemText $_.Location -CompletionResultType ParameterValue -Verbose }
}
