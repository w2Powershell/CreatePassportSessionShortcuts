del "C:\Users\Public\Desktop\SLSS - Train.lnk"
del "C:\Users\Public\Desktop\SLSS - DEV.lnk"
del "C:\Users\Public\Desktop\SLSS - Production.lnk"
del "C:\Users\Public\Desktop\SLSS - Primary.lnk"
del "C:\Users\Public\Desktop\SLSS - Additional.lnk"
del "C:\Users\Public\Desktop\SLSS - Com Additional.lnk"
del "C:\Users\Public\Desktop\SLSS - Commercial.lnk"

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