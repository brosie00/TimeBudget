
function New-TBBody {
param
(
    $Entry_1 = 'Entry_1'
    ,
    $Entry_2 = 'Entry_2'
)
    $StringArray = @()
    $StringArray += "Entry_1: $Entry_1"
    $stringArray += " "
    $StringArray += "Entry_2: $Entry_2"
    #$StringArray = $StringArray | Out-String

    Write-Output $StringArray 
}

function Split-TBBody {
param
(

)

$a = $StringArray | foreach { $_ -split '`\n`\r' }
$a[0]

}
