function Add-TBCategories {

param 
(
    [string[]]
    $Categories
)

$Categories | foreach { $_}

$Categories | foreach { $Namespace.Categories.Add($_) }

}

