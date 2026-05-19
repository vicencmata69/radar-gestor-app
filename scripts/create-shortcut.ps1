# Recrea la drecera Radar-Gestor a l'escriptori de forma robusta.
# Es pot tornar a executar sempre que calgui (sobreescriu l'anterior).

$ErrorActionPreference = "Stop"
$desktop = [Environment]::GetFolderPath("Desktop")
$lnkPath = Join-Path $desktop "Radar-Gestor.lnk"
$target  = "C:\Users\PMASCLANS\Downloads\radar-gestor-app\start-app.bat"
$workdir = "C:\Users\PMASCLANS\Downloads\radar-gestor-app"

if (-not (Test-Path $target)) {
    Write-Output "ERROR: no existeix el target $target"
    exit 1
}

# Esborra la drecera anterior si hi era (per si estava corrupta)
if (Test-Path $lnkPath) { Remove-Item $lnkPath -Force }

$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($lnkPath)
$sc.TargetPath       = $target
$sc.WorkingDirectory = $workdir
$sc.IconLocation     = "C:\Windows\System32\imageres.dll,235"
$sc.Description       = "Arrenca Radar-Gestor en local (http://localhost:5173/)"
$sc.WindowStyle      = 1
$sc.Save()

Start-Sleep -Milliseconds 400

if (Test-Path $lnkPath) {
    $chk = $ws.CreateShortcut($lnkPath)
    Write-Output "OK - Drecera creada:"
    Write-Output ("  Ubicacio : " + $lnkPath)
    Write-Output ("  Target   : " + $chk.TargetPath)
    Write-Output ("  TargetOK : " + (Test-Path $chk.TargetPath))
} else {
    Write-Output "ERROR: no s'ha pogut crear la drecera a $lnkPath"
    exit 1
}
