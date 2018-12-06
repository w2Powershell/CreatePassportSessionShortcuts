$files = Get-ChildItem "\\softwareShare\files\software\passport\sessions"

foreach ($file in $files) {
$newName = [io.path]::GetFileNameWithoutExtension($file)
    if ($newName -ne "shortcuts") {    
    
    $AppLocation = "c:\program files\passport\passport.exe"
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("c:\users\public\Desktop\$newName.lnk")
    $Shortcut.TargetPath = "U:\sessions\$newName.zws"
    $Shortcut.IconLocation = "$AppLocation,1"
    $Shortcut.Description ="$newName"
    $Shortcut.WorkingDirectory ="U:\sessions\"
    $Shortcut.Save()    
    } 
}

