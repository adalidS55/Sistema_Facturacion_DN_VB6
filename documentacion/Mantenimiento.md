# Mantenimiento

## Requisitos

- Windows con soporte para aplicaciones de 32 bits y el runtime de Visual Basic 6.
- Microsoft Jet OLE DB 4.0 y ADO, utilizados por las bases Access `.mdb`.
- Controles registrados `MSFLXGRD.OCX` y `MSCOMCT2.OCX`.
- Bibliotecas de formato de datos y Data Report referenciadas en `Facturacion.vbp`.
- Para desarrollar y compilar: Visual Basic 6 y las referencias del proyecto instaladas.

Los componentes del sistema se referencian desde su instalación en Windows; no se copian DLL ni OCX sueltos. Esta carpeta no es un instalador para otro equipo.

## Organización y compatibilidad

Se conservaron los nombres internos de formularios, módulos y reportes para evitar cambios de comportamiento. El código VB6 mantiene la codificación Windows-1252 y los saltos de línea CRLF. Cada `.frx` o `.dsx` permanece junto a su formulario o diseñador.

El proyecto incluye únicamente los formularios, módulos, clases y diseñadores declarados en el `.vbp`, con sus recursos. Se excluyeron temporales, bloqueos de Access, cachés del IDE, registros antiguos, copias históricas, imágenes de diseño y documentos que no utiliza el programa.

Las rutas de las bases y pendientes se definen en `codigo/modulos/Module1.bas`, mediante `InicializarRutaBD`. Tanto el ejecutable como el `.vbp` deben permanecer en la raíz del sistema.

## Pendientes

Guardar pendiente conserva los campos y productos sin registrar una venta ni descontar inventario. Al recuperar se vuelve a consultar el correlativo. Cambiar de pendiente o cerrar la ventana conserva automáticamente el trabajo actual. Los archivos `.pend` forman parte de los datos que deben respaldarse.

## Pruebas

`herramientas/Probar.ps1` compila una aplicación de prueba en una carpeta temporal única. Copia las dos bases con acceso exclusivo para evitar una copia inconsistente; si están en uso, solicita cerrarlas mediante un error y no continúa. Comprueba ventana única, persistencia, recuperación de campos y filas, producto incompleto, descuento, correlativo disponible, intercambio y cierre. No prueba impresión ni registra ventas reales.

`registros/compilacion.log` y `registros/pruebas.log` contienen los últimos resultados. La herramienta muestra también la ubicación temporal de la prueba para investigar errores.
