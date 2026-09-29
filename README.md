# Facturación Donde Nacho

Sistema de Visual Basic 6 con una sola ventana de facturación y facturas pendientes persistentes.

## Uso

Ejecutar **Facturacion.exe** desde esta carpeta. Para editar el proyecto, abrir **Facturacion.vbp** en Visual Basic 6.

La aplicación utiliza exclusivamente las bases de datos de `datos/`. Las pendientes están en `datos/pendientes/`. Se puede trasladar la carpeta completa: las rutas de datos se resuelven desde la ubicación del programa, también al ejecutar el proyecto en VB6.

## Estructura

| Ubicación | Contenido |
| --- | --- |
| `Facturacion.exe` | Programa listo para usar |
| `Facturacion.vbp` | Proyecto principal de VB6 |
| `codigo/formularios/` | Ventanas `.frm` y sus recursos `.frx` |
| `codigo/modulos/` | Rutinas comunes y configuración de rutas |
| `codigo/clases/` | Guardado y recuperación de pendientes |
| `codigo/informes/` | Diseñadores `.Dsr` y recursos `.dsx` incluidos en el proyecto |
| `datos/` | Bases de datos operativas |
| `datos/pendientes/` | Facturas pendientes `.pend` |
| `herramientas/` | Compilación y pruebas |
| `pruebas/` | Prueba de integración de facturas pendientes |
| `documentacion/` | Requisitos y mantenimiento |
| `registros/` | Resultados de compilación y pruebas; se crea al usar las herramientas |

## Compilar y comprobar

En PowerShell, desde esta carpeta:

```powershell
& .\herramientas\Compilar.ps1
& .\herramientas\Probar.ps1
```

Ambos scripts aceptan `-Compilador` para indicar otra ubicación de VB6.exe. Cerrar el sistema antes de compilar o probar. La prueba usa copias temporales de las bases y pendientes de prueba; no registra ventas en los datos operativos.

## Copias de seguridad

Cerrar el programa y respaldar la carpeta `datos` completa, incluidas las pendientes. Para un respaldo del proyecto y del sistema, copiar toda esta carpeta. Las bases y ejecutables se excluyen del control de versiones mediante `.gitignore`; el respaldo de datos debe hacerse por separado.

Esta copia es independiente del proyecto anterior. Después de comenzar a trabajar aquí, utilizar **Facturacion.exe** de esta ubicación para que todas las ventas queden en la misma base.

## Control de Inventario

Abrir **Inventarios → Control de Inventario**, con el mismo acceso de usuario de Productos.

- La cuadrícula muestra todos los productos. Solo se edita **Cantidad nueva**, después de pulsar **Guardar datos antiguos** en cada apertura de la ventana. Si se elimina esa copia, se bloquea la edición hasta guardar otra.
- Usar Enter o las flechas arriba/abajo para continuar con otro producto; RePág/AvPág saltan diez filas. Esc deshace la edición de la celda activa. Ctrl+S inicia el guardado.
- Se admiten cantidades con hasta cuatro decimales. Una celda vacía representa una cantidad sin dato (NULL), distinta de cero.
- **Guardar cambios** confirma las cantidades y aplica los cambios en una transacción, sin crear respaldos automáticos. **Guardar datos antiguos** guarda los datos actuales de la base; **Guardar datos nuevos** guarda códigos, nombres y cantidades de Cantidad nueva, sin aplicar cambios. Se guardan en `datos/respaldos-inventario/` y aparecen en la lista como `datosantiguos-fecha-hora` y `datosnuevos-fecha-hora`. Ante un error, se revierte toda la operación.
- Si una cantidad cambió desde la carga de la lista, se detiene el guardado. **Recargar** trae las cantidades actuales desde la base y conserva las nuevas editadas, identificando cada producto por su código. Los productos sin edición se actualizan con su cantidad actual. **Descartar** restablece todas las cantidades nuevas con las actuales que muestra la tabla, sin consultar ni modificar la base.
- Para recuperar, seleccionar un respaldo y pulsar **Revisar / Recuperar**. Siempre aparece una confirmación antes de reemplazar las cantidades nuevas. Se preparan las cantidades anteriores de los productos que todavía existen; no se cambian precios ni se crean o eliminan productos. Revisar y pulsar **Guardar cambios** para aplicarlas. La recuperación no genera respaldos automáticos. El respaldo seleccionado se elimina de `datos/respaldos-inventario/` únicamente después de confirmar y guardar correctamente la recuperación; si se cancela o falla, se conserva. Si las cantidades ya coinciden, Guardar cambios también permite completar y retirar ese respaldo.
- Al cerrar con cambios sin guardar, se solicita confirmar que se descartarán.

Los respaldos de inventario forman parte de la carpeta `datos` que debe incluirse en las copias de seguridad.
### Excel de control de inventario

**Generar Excel** permite elegir dónde guardar un `.xlsx` mediante el diálogo de Windows. El archivo se genera directamente, sin iniciar Excel ni depender de que esté instalado. Las columnas son: Código Producto, Nombre, Precio de venta, Cantidad antigua, Cantidades Nuevas, una columna vacía, Diferencia y Dinero.

- Diferencia contiene la fórmula cantidad nueva menos cantidad antigua.
- Dinero contiene la fórmula precio de venta por diferencia, con formato de lempiras.
- Los códigos se conservan como texto, incluidos ceros iniciales. Una cantidad sin dato deja vacíos los cálculos asociados; cero sigue siendo una cantidad válida.
- El Excel se genera exclusivamente con los archivos elegidos en **Excel: datos antiguos** y **Excel: datos nuevos**. Código, nombre, precio y cantidad antigua salen del archivo antiguo; cantidad nueva sale del nuevo. Se cruzan por código y se rechazan archivos con productos diferentes. Primero guardar datos antiguos, editar, guardar datos nuevos y seleccionar ambas copias. No se exportan cantidades sin respaldar ni datos actuales de la base.
- La exportación no guarda cambios de inventario ni consume respaldos.
### Pendientes y limpieza de respaldos

Facturación guarda una pendiente únicamente si el FlexGrid contiene datos de productos. Los campos de cliente, cobro o producto todavía sin agregar no generan una pendiente por sí solos. Esta condición también aplica al cerrar la ventana o recuperar otra pendiente.

En Control de Inventario, **Borrar respaldos** solicita confirmación antes de eliminar los archivos de respaldo de inventario (.adtg). No elimina productos, facturas pendientes ni cantidades en edición. Los respaldos eliminados no se pueden recuperar.

### Resumen del Excel

Después del último producto se deja una fila vacía, seguida de **Total**, otra fila vacía, **Sobrante**, otra fila vacía y **Diferencia**, en la columna G. En la columna H, Total suma el dinero de todos los productos; Sobrante queda vacío para ingreso manual; Diferencia suma Total y Sobrante.

La columna Dinero aplica formato condicional: regla mayor que 0 con relleno verde claro y texto verde oscuro; regla menor que 0 con relleno rojo claro y texto rojo oscuro. Los encabezados y las etiquetas del resumen conservan su fondo de color. Las reglas cubren la columna H desde la fila 2. Los ceros y las celdas vacías conservan el formato normal. Las reglas también cubren el resumen y reaccionan a los cambios introducidos en Excel.
Facturación permite ventas sin existencia suficiente: descuenta las unidades vendidas y admite saldos negativos. Una existencia NULL se toma como cero al vender.

## Presentaciones con inventario compartido

En Productos, crear un código para el paquete con su propio nombre, precio de venta y precio de costo (ambos por paquete). Marcar **Usar inventario de otro producto**, seleccionar el código base y escribir las unidades que consume. Por ejemplo, `022P5` usa `022` con factor `5`. Guardar o Actualizar conserva el producto y su vinculación juntos.

El paquete se guarda con cantidad propia cero; en Productos se muestra cantidad propia cero y en Facturación la disponibilidad se calcula a partir del producto base. Las entradas de mercancía y los ajustes de inventario se realizan en el código base. Control de Inventario e Inventario General muestran solo productos con inventario propio, para evitar contar o modificar dos veces las mismas existencias. Sus copias manuales y Excel mantienen ese mismo alcance.

Facturación admite códigos alfanuméricos. Una venta mixta de dos paquetes de cinco libras y tres libras sueltas descuenta trece libras del base. La advertencia de disponibilidad considera las otras líneas de la factura; se conserva la posibilidad de vender con saldo negativo.

La tabla `presentaciones` guarda código de venta, código base y factor. La anulación lee la cantidad vendida de `facturacion` y devuelve cantidad por factor al producto base; los productos sin vínculo usan factor 1. No se genera una tabla adicional de consumos. Mientras existan facturas sin anular para un código, no se permite cambiar ni quitar su vínculo o factor; para otro tamaño se crea otro código. Se pueden modificar nombre y precios. La anulación se realiza en una transacción y cancelar su confirmación no modifica nada.

No se permiten factores cero o negativos, más de cuatro decimales, vínculos consigo mismo ni cadenas de paquetes. No se convierte en paquete un producto que conserva existencia propia, abastece otros paquetes o debe conservar su inventario para anular ventas anteriores. Usar un código nuevo en esos casos. Los productos base referenciados no pueden eliminarse.
El selector de producto base muestra solo nombres: se puede escribir el nombre completo o su inicio y pulsar Enter. Si hay varias coincidencias, se debe completar el nombre o elegir de la lista. El codigo vinculado se conserva internamente.
