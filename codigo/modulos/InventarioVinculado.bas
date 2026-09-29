Attribute VB_Name = "InventarioVinculado"
Option Explicit

Public Function SqlTexto(ByVal valor As String) As String
    SqlTexto = "'" & Replace$(valor, "'", "''") & "'"
End Function

Public Function SqlCantidad(ByVal valor As Currency) As String
    SqlCantidad = Trim$(Str$(valor))
End Function

Private Function ExisteTablaInventario(ByVal cn As ADODB.Connection, ByVal nombre As String) As Boolean
    Dim r As ADODB.Recordset
    Set r = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, nombre, "TABLE"))
    ExisteTablaInventario = Not r.EOF
    r.Close
End Function

Public Sub PrepararVinculaciones(ByVal cn As ADODB.Connection)
    Dim transaccion As Boolean, mensaje As String
    On Error GoTo Fallo
    If ExisteTablaInventario(cn, "presentaciones") Then Exit Sub
    cn.BeginTrans
    transaccion = True
    If Not ExisteTablaInventario(cn, "presentaciones") Then
        cn.Execute "CREATE TABLE presentaciones (codventa TEXT(255) CONSTRAINT pk_presentaciones PRIMARY KEY, codbase TEXT(255) NOT NULL, factor CURRENCY NOT NULL)"
    End If
    cn.CommitTrans
    Exit Sub
Fallo:
    mensaje = Err.Description
    On Error Resume Next
    If transaccion Then cn.RollbackTrans
    On Error GoTo 0
    Err.Raise vbObjectError + 2200, , "No se pudo preparar el inventario vinculado: " & mensaje
End Sub

Public Sub ResolverInventario(ByVal cn As ADODB.Connection, ByVal codigo As String, ByRef base As String, ByRef factor As Currency)
    Dim r As ADODB.Recordset
    base = codigo
    factor = 1
    Set r = cn.Execute("SELECT codbase,factor FROM presentaciones WHERE codventa=" & SqlTexto(codigo))
    If Not r.EOF Then
        base = CStr(r!codbase)
        factor = CCur(r!factor)
        If factor <= 0 Then Err.Raise 5, , "Factor de inventario no valido para " & codigo
    End If
    r.Close
End Sub

Public Function EsPresentacion(ByVal cn As ADODB.Connection, ByVal codigo As String) As Boolean
    Dim base As String, factor As Currency
    ResolverInventario cn, codigo, base, factor
    EsPresentacion = (StrComp(base, codigo, vbTextCompare) <> 0)
End Function

Public Function ExistenciaVenta(ByVal cn As ADODB.Connection, ByVal codigo As String) As Double
    Dim base As String, factor As Currency, r As ADODB.Recordset
    ResolverInventario cn, codigo, base, factor
    Set r = cn.Execute("SELECT cantprod FROM productos WHERE codproducto=" & SqlTexto(base))
    If r.EOF Then Err.Raise 5, , "No existe el producto de inventario " & base
    If Not IsNull(r!cantprod) Then ExistenciaVenta = CDbl(r!cantprod) / CDbl(factor)
    r.Close
End Function

Public Sub MoverExistenciaBase(ByVal cn As ADODB.Connection, ByVal base As String, ByVal cantidad As Currency)
    Dim afectados As Variant
    cn.Execute "UPDATE productos SET cantprod=IIF(cantprod IS NULL,0,cantprod)+(" & SqlCantidad(cantidad) & ") WHERE codproducto=" & SqlTexto(base) & " AND codproducto NOT IN (SELECT codventa FROM presentaciones)", afectados
    If afectados <> 1 Then Err.Raise 5, , "No se pudo actualizar la existencia del producto base " & base
End Sub

Public Sub ConsumirVenta(ByVal cn As ADODB.Connection, ByVal codigo As String, ByVal unidades As Currency)
    Dim base As String, factor As Currency, consumo As Currency, exacto As Variant
    If unidades <= 0 Then Err.Raise 5, , "La cantidad vendida debe ser mayor que cero."
    ResolverInventario cn, codigo, base, factor
    exacto = CDec(unidades) * CDec(factor)
    consumo = CCur(exacto)
    If CDec(consumo) <> exacto Then Err.Raise 5, , "El consumo excede cuatro decimales para " & codigo
    MoverExistenciaBase cn, base, -consumo
End Sub

Public Sub AnularVentaVinculada(ByVal cn As ADODB.Connection, ByVal factura As Long)
    Dim r As ADODB.Recordset, transaccion As Boolean, mensaje As String, afectados As Variant
    Dim base As String, factor As Currency, cantidad As Currency, exacto As Variant
    On Error GoTo Fallo
    cn.BeginTrans
    transaccion = True
    Set r = cn.Execute("SELECT COUNT(*) AS n FROM facturacion WHERE numfact=" & factura)
    If r!n = 0 Then Err.Raise 5, , "La factura no existe o ya fue anulada."
    r.Close
    Set r = cn.Execute("SELECT codprodfact,cantidad FROM facturacion WHERE numfact=" & factura)
    If r.EOF Then Err.Raise 5, , "La factura no tiene productos para devolver."
    Do While Not r.EOF
        If IsNull(r!codprodfact) Or IsNull(r!cantidad) Then Err.Raise 5, , "La factura tiene un producto o cantidad sin dato."
        ResolverInventario cn, CStr(r!codprodfact), base, factor
        exacto = CDec(r!cantidad) * CDec(factor)
        cantidad = CCur(exacto)
        If CDec(cantidad) <> exacto Then Err.Raise 5, , "La devolucion excede cuatro decimales."
        MoverExistenciaBase cn, base, cantidad
        r.MoveNext
    Loop
    r.Close
    cn.Execute "INSERT INTO facturacionnulas SELECT * FROM facturacion WHERE numfact=" & factura
    cn.Execute "DELETE FROM facturacion WHERE numfact=" & factura, afectados
    If afectados = 0 Then Err.Raise 5, , "La factura cambio durante la anulacion."
    cn.Execute "DELETE FROM facturareporte WHERE numfact=" & factura
    cn.CommitTrans
    Exit Sub
Fallo:
    mensaje = Err.Description
    On Error Resume Next
    If transaccion Then cn.RollbackTrans
    On Error GoTo 0
    Err.Raise vbObjectError + 2201, , mensaje
End Sub

Public Sub ValidarVinculo(ByVal cn As ADODB.Connection, ByVal codigo As String, ByVal base As String, ByVal factor As Currency)
    Dim r As ADODB.Recordset, anterior As String, factorAnterior As Currency, propuesta As String
    ResolverInventario cn, codigo, anterior, factorAnterior
    propuesta = base
    If Len(propuesta) = 0 Then propuesta = codigo
    If StrComp(anterior, propuesta, vbTextCompare) <> 0 Or (Len(base) > 0 And factorAnterior <> factor) Then
        Set r = cn.Execute("SELECT TOP 1 numfact FROM facturacion WHERE codprodfact=" & SqlTexto(codigo))
        If Not r.EOF Then Err.Raise 5, , "Este producto tiene facturas sin anular. No se puede cambiar ni quitar su vinculacion o factor; cree otro codigo para una presentacion distinta."
        r.Close
    End If
    If Len(base) = 0 Then Exit Sub
    If StrComp(codigo, base, vbTextCompare) = 0 Then Err.Raise 5, , "Un producto no puede usar su propio codigo como vinculacion."
    If factor <= 0 Then Err.Raise 5, , "Las unidades consumidas deben ser mayores que cero."
    Set r = cn.Execute("SELECT codproducto FROM productos WHERE codproducto=" & SqlTexto(base))
    If r.EOF Then Err.Raise 5, , "Seleccione un producto base existente."
    r.Close
    If EsPresentacion(cn, base) Then Err.Raise 5, , "El producto base debe tener inventario propio; no se permiten cadenas de vinculaciones."
    Set r = cn.Execute("SELECT codventa FROM presentaciones WHERE codbase=" & SqlTexto(codigo))
    If Not r.EOF Then Err.Raise 5, , "Este producto ya abastece otras presentaciones y debe conservar inventario propio."
    r.Close
    Set r = cn.Execute("SELECT cantprod FROM productos WHERE codproducto=" & SqlTexto(codigo))
    If Not r.EOF Then
        If Not IsNull(r!cantprod) Then
            If CCur(r!cantprod) <> 0 Then Err.Raise 5, , "El producto tiene existencia propia. Use un codigo nuevo o regularice primero su inventario."
        End If
    End If
    r.Close
End Sub

Public Sub GuardarProductoVinculado(ByVal cn As ADODB.Connection, ByVal nuevo As Boolean, ByVal codigo As String, ByVal nombre As String, ByVal ingreso As Currency, ByVal precio As Currency, ByVal costo As Currency, ByVal fecha As Date, ByVal base As String, ByVal factor As Currency)
    Dim r As ADODB.Recordset, transaccion As Boolean, mensaje As String, cantidad As Currency
    On Error GoTo Fallo
    cn.BeginTrans
    transaccion = True
    ValidarVinculo cn, codigo, base, factor
    If Len(base) > 0 And ingreso <> 0 Then Err.Raise 5, , "Ingrese las existencias en el producto base. El paquete no mantiene cantidad propia."
    Set r = New ADODB.Recordset
    r.Open "SELECT * FROM productos WHERE codproducto=" & SqlTexto(codigo), cn, adOpenKeyset, adLockOptimistic
    If nuevo Then
        If Not r.EOF Then Err.Raise 5, , "Ya existe ese codigo de producto."
        r.AddNew
        r!codproducto = codigo
    Else
        If r.EOF Then Err.Raise 5, , "No existe el producto que desea actualizar."
        If Not IsNull(r!cantprod) Then cantidad = CCur(r!cantprod)
    End If
    If Len(base) > 0 Then
        cantidad = 0
    Else
        cantidad = cantidad + ingreso
    End If
    r!nombreprod = nombre
    r!cantprod = cantidad
    r!preciov = precio
    r!precioc = costo
    r!fechav = fecha
    r.Update
    r.Close
    cn.Execute "DELETE FROM presentaciones WHERE codventa=" & SqlTexto(codigo)
    If Len(base) > 0 Then cn.Execute "INSERT INTO presentaciones (codventa,codbase,factor) VALUES (" & SqlTexto(codigo) & "," & SqlTexto(base) & "," & SqlCantidad(factor) & ")"
    If ingreso <> 0 Then
        Set r = New ADODB.Recordset
        r.Open "SELECT * FROM ingresos WHERE 1=0", cn, adOpenKeyset, adLockOptimistic
        r.AddNew
        r!codigo = codigo
        r!descripcion = nombre
        r!cantidad = ingreso
        r!fecha = Date
        r!hora = Time
        r.Update
        r.Close
    End If
    cn.CommitTrans
    Exit Sub
Fallo:
    mensaje = Err.Description
    On Error Resume Next
    If Not r Is Nothing Then
        If r.EditMode <> adEditNone Then r.CancelUpdate
    End If
    If transaccion Then cn.RollbackTrans
    On Error GoTo 0
    Err.Raise vbObjectError + 2202, , mensaje
End Sub

Public Sub ValidarBorradoProducto(ByVal cn As ADODB.Connection, ByVal codigo As String)
    Dim r As ADODB.Recordset
    Set r = cn.Execute("SELECT codventa FROM presentaciones WHERE codbase=" & SqlTexto(codigo) & " OR codventa=" & SqlTexto(codigo))
    If Not r.EOF Then Err.Raise 5, , "Quite primero las vinculaciones de este producto."
    r.Close
    Set r = cn.Execute("SELECT TOP 1 numfact FROM facturacion WHERE codprodfact=" & SqlTexto(codigo))
    If Not r.EOF Then Err.Raise 5, , "No puede eliminar un producto necesario para anular ventas existentes."
    r.Close
End Sub
