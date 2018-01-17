$object = New-Object -comobject outlook.application
$namespace = $object.GetNamespace("MAPI")
Write-Host $namespace.GetDefaultFolder(1)
Write-Host $namespace.GetDefaultFolder(2)
Write-Host $namespace.GetDefaultFolder(3)
Write-Host $namespace.GetDefaultFolder(4)