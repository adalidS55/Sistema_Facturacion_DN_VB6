param([string]$Compilador = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE')
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path -LiteralPath $Compilador)) { throw 'No se encontro VB6.exe. Indique -Compilador con su ruta.' }
$registros = Join-Path $raiz 'registros'
New-Item -ItemType Directory -Path $registros -Force | Out-Null
$log = Join-Path $registros 'compilacion.log'
[IO.File]::WriteAllText($log, '')
$proyecto = Join-Path $raiz 'Facturacion.vbp'
$ejecutable = Join-Path $raiz 'Facturacion.exe'
$anterior = if (Test-Path -LiteralPath $ejecutable) { (Get-Item -LiteralPath $ejecutable).LastWriteTimeUtc } else { [datetime]::MinValue }
$argumentos = '/make "{0}" /out "{1}" /outdir "{2}"' -f $proyecto, $log, $raiz
$proceso = Start-Process -FilePath $Compilador -ArgumentList $argumentos -WindowStyle Hidden -PassThru
if (-not $proceso.WaitForExit(120000)) { throw "La compilacion sigue abierta. Revise VB6 y $log" }
Get-Content -LiteralPath $log
if (-not (Test-Path -LiteralPath $ejecutable)) { throw "No se genero el ejecutable. Revise $log" }
if ((Get-Item -LiteralPath $ejecutable).LastWriteTimeUtc -le $anterior) { throw "No se actualizo el ejecutable. Cierre el programa y revise $log" }
Write-Output "Ejecutable actualizado directamente: $ejecutable"