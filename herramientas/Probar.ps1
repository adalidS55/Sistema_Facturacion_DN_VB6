param([string]$Compilador = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE')
$ErrorActionPreference = 'Stop'
$raiz = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path -LiteralPath $Compilador)) { throw 'No se encontro VB6.exe. Indique -Compilador con su ruta.' }
$salida = Join-Path ([IO.Path]::GetTempPath()) ('Facturacion-Pruebas-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path (Join-Path $salida 'datos\pendientes') -Force | Out-Null
Write-Output "Prueba aislada en: $salida"
$lecturas = @()
try {
    foreach ($nombre in @('dondenacho1.mdb','Ingresos.mdb')) {
        $lecturas += [IO.File]::Open((Join-Path $raiz "datos\$nombre"), [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    }
    for ($i = 0; $i -lt $lecturas.Count; $i++) {
        $destino = [IO.File]::Create((Join-Path $salida ('datos\' + [IO.Path]::GetFileName($lecturas[$i].Name))))
        try { $lecturas[$i].CopyTo($destino) } finally { $destino.Dispose() }
    }
} finally { foreach ($lectura in $lecturas) { $lectura.Dispose() } }
$enc = [Text.Encoding]::GetEncoding(1252)
$lineas = foreach ($linea in [IO.File]::ReadAllLines((Join-Path $raiz 'Facturacion.vbp'), $enc)) {
    if ($linea -match '^(Form|Designer)=(.+)$') {
        $matches[1] + '=' + (Join-Path $raiz $matches[2])
    } elseif ($linea -match '^((?:Module|Class)=.+; )(.+)$') {
        $matches[1] + (Join-Path $raiz $matches[2])
    } elseif ($linea.StartsWith('Startup=')) { 'Startup="Sub Main"'
    } elseif ($linea.StartsWith('ExeName32=')) { 'ExeName32="PruebaPendientes.exe"'
    } else { $linea }
}
$lineas += 'Module=PruebaPendientes; ' + (Join-Path $raiz 'pruebas\PruebaPendientes.bas')
$proyecto = Join-Path $salida 'PruebaPendientes.vbp'
[IO.File]::WriteAllLines($proyecto, [string[]]$lineas, $enc)
$log = Join-Path $salida 'compilacion.log'
$argumentos = '/make "{0}" /out "{1}" /outdir "{2}"' -f $proyecto, $log, $salida
$proceso = Start-Process -FilePath $Compilador -ArgumentList $argumentos -WindowStyle Hidden -PassThru
if (-not $proceso.WaitForExit(120000)) { throw "Revise la compilacion de prueba en $salida" }
$ejecutable = Join-Path $salida 'PruebaPendientes.exe'
if (-not (Test-Path -LiteralPath $ejecutable)) { throw "No se compilo la prueba. Revise $log" }
$prueba = Start-Process -FilePath $ejecutable -WindowStyle Hidden -PassThru
if (-not $prueba.WaitForExit(60000)) { throw "La prueba sigue abierta. Revise la ventana de prueba y $salida" }
$resultado = Join-Path $salida 'resultado-pendientes.txt'
if (-not (Test-Path -LiteralPath $resultado)) { throw 'La prueba no genero un resultado.' }
$registros = Join-Path $raiz 'registros'
New-Item -ItemType Directory -Path $registros -Force | Out-Null
Copy-Item -LiteralPath $resultado -Destination (Join-Path $registros 'pruebas.log') -Force
$texto = [IO.File]::ReadAllText($resultado, $enc)
Write-Output $texto
if (-not $texto.StartsWith('OK:')) { throw 'La prueba de pendientes fallo.' }
