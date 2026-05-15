$WshShell = New-Object -ComObject WScript.Shell
$ShortcutPath = "$env:USERPROFILE\Desktop\Radar-Gestor.lnk"
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "C:\Users\PMASCLANS\Downloads\radar-gestor-app\start-app.bat"
$Shortcut.WorkingDirectory = "C:\Users\PMASCLANS\Downloads\radar-gestor-app"
$Shortcut.IconLocation = "C:\Windows\System32\imageres.dll,235"
$Shortcut.Description = "Arrenca Radar-Gestor en local (http://localhost:5173/)"
$Shortcut.Save()
if (Test-Path $ShortcutPath) {
    Write-Output "OK: Shortcut creat a $ShortcutPath"
} else {
    Write-Output "ERROR: no s'ha pogut crear el shortcut"
}
