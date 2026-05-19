# Instal.la l'arrencada automatica del servidor Radar-Gestor a l'inici
# de sessio de Windows. Crea una drecera (minimitzada) a la carpeta
# Startup que executa start-app.bat. Aixi el servidor sempre esta actiu
# i nomes cal obrir el navegador a http://localhost:5173/.
# Es pot tornar a executar sense problema (sobreescriu).

$ErrorActionPreference = "Stop"
$startup = [Environment]::GetFolderPath("Startup")
$lnkPath = Join-Path $startup "Radar-Gestor (servidor).lnk"
$target  = "C:\Users\PMASCLANS\Downloads\radar-gestor-app\start-app.bat"
$workdir = "C:\Users\PMASCLANS\Downloads\radar-gestor-app"

if (-not (Test-Path $target)) { Write-Output "ERROR: no existeix $target"; exit 1 }
if (Test-Path $lnkPath) { Remove-Item $lnkPath -Force }

$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($lnkPath)
$sc.TargetPath       = $target
$sc.WorkingDirectory = $workdir
$sc.IconLocation     = "C:\Windows\System32\imageres.dll,235"
$sc.Description       = "Servidor Radar-Gestor (arrenca sol amb Windows)"
$sc.WindowStyle      = 7   # 7 = minimitzada (no molesta, visible a la barra)
$sc.Save()

Start-Sleep -Milliseconds 400
if (Test-Path $lnkPath) {
    Write-Output "OK - Autostart instal.lat:"
    Write-Output ("  " + $lnkPath)
    Write-Output "  El servidor arrencara sol cada cop que iniciis sessio a Windows."
} else {
    Write-Output "ERROR: no s'ha pogut crear la drecera d'autostart"
    exit 1
}
