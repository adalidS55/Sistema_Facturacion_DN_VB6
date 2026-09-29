VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ReporteVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TAMANO_PAGINA As Long = 1000
Private datos As ADODB.Recordset
Private esDetalle As Boolean
Private registros As Long
Private registrosHistoricos As Long
Private paginaActual As Long
Private vendido As Double
Private ganancia As Double

Public Sub Limpiar()
    If Not datos Is Nothing Then
        If datos.State = adStateOpen Then datos.Close
    End If
    Set datos = Nothing
    registros = 0
    registrosHistoricos = 0
    paginaActual = 0
    vendido = 0
    ganancia = 0
End Sub

Private Function ConsultaFechas(ByVal cn As ADODB.Connection, ByVal sql As String, ByVal desde As Date, ByVal hasta As Date) As ADODB.Recordset
    Dim orden As New ADODB.Command, resultado As New ADODB.Recordset
    Set orden.ActiveConnection = cn
    orden.CommandText = sql
    orden.Parameters.Append orden.CreateParameter("desde", adDate, adParamInput, , DateValue(desde))
    orden.Parameters.Append orden.CreateParameter("hasta", adDate, adParamInput, , DateAdd("d", 1, DateValue(hasta)))
    resultado.CursorLocation = adUseClient
    resultado.Open orden, , adOpenStatic, adLockReadOnly
    Set resultado.ActiveConnection = Nothing
    Set ConsultaFechas = resultado
End Function

Public Sub Cargar(ByVal cn As ADODB.Connection, ByVal desde As Date, ByVal hasta As Date, ByVal detalle As Boolean)
    Dim tabla As String, filtro As String, campos As String, importes As String
    Dim resumen As ADODB.Recordset, historia As ADODB.Recordset
    Limpiar
    If DateValue(desde) > DateValue(hasta) Then Err.Raise 5, , "La fecha inicial debe ser anterior o igual a la final."
    esDetalle = detalle
    filtro = " WHERE fechafact >= ? AND fechafact < ?"
    If detalle Then
        tabla = "facturacion"
        campos = "numfact,fechafact,horafact,nombreprodfact,cantidad,precio,preciocosto,descuentoi,subtotal,totaldescuento,totalneto,totalganacia"
        importes = "SUM(CDbl(IIF(cantidad IS NULL,0,cantidad))*(CDbl(IIF(precio IS NULL,0,precio))-CDbl(IIF(descuentoi IS NULL,0,descuentoi)))) AS vendido, " & _
            "SUM(CDbl(IIF(cantidad IS NULL,0,cantidad))*(CDbl(IIF(precio IS NULL,0,precio))-CDbl(IIF(preciocosto IS NULL,0,preciocosto))-CDbl(IIF(descuentoi IS NULL,0,descuentoi)))) AS ganancia"
    Else
        tabla = "facturareporte"
        campos = "numfact,fechafact,horafact,subtotal,totaldescuento,totalneto,totalganacia"
        importes = "SUM(CDbl(IIF(totalneto IS NULL,0,totalneto))) AS vendido, SUM(CDbl(IIF(totalganacia IS NULL,0,totalganacia))) AS ganancia"
    End If
    Set resumen = ConsultaFechas(cn, "SELECT COUNT(*) AS cantidad," & importes & " FROM " & tabla & filtro, desde, hasta)
    registros = CLng(resumen!cantidad)
    If Not IsNull(resumen!vendido) Then vendido = CDbl(resumen!vendido)
    If Not IsNull(resumen!ganancia) Then ganancia = CDbl(resumen!ganancia)
    resumen.Close
    Set historia = cn.Execute("SELECT COUNT(*) AS cantidad FROM " & tabla)
    registrosHistoricos = CLng(historia!cantidad)
    historia.Close
    If registros = 0 Then Exit Sub
    Set datos = ConsultaFechas(cn, "SELECT " & campos & " FROM " & tabla & filtro & " ORDER BY fechafact,numfact,horafact", desde, hasta)
    If datos.RecordCount <> registros Then
        Limpiar
        Err.Raise 5, , "Las ventas cambiaron durante la consulta. Vuelva a generar el informe."
    End If
    paginaActual = 1
End Sub

Public Property Get Cantidad() As Long
    Cantidad = registros
End Property

Public Property Get CantidadHistorica() As Long
    CantidadHistorica = registrosHistoricos
End Property

Public Property Get TotalVendido() As Double
    TotalVendido = vendido
End Property

Public Property Get TotalGanancia() As Double
    TotalGanancia = ganancia
End Property

Public Property Get Pagina() As Long
    Pagina = paginaActual
End Property

Public Property Get Paginas() As Long
    If registros > 0 Then Paginas = ((registros - 1) \ TAMANO_PAGINA) + 1
End Property

Public Sub MoverPagina(ByVal cambio As Long)
    Dim destino As Long
    destino = paginaActual + cambio
    If destino >= 1 And destino <= Paginas Then paginaActual = destino
End Sub

Public Sub Mostrar(ByVal cuadricula As Object)
    Dim primero As Long, cuantos As Long, fila As Long, columna As Long, valor As Variant
    Dim mensaje As String
    On Error GoTo Fallo
    cuadricula.Redraw = False
    cuadricula.Rows = 2
    For columna = 0 To cuadricula.Cols - 1
        cuadricula.TextMatrix(1, columna) = vbNullString
    Next columna
    If registros = 0 Then
        cuadricula.Rows = 2
        cuadricula.Redraw = True
        Exit Sub
    End If
    primero = (paginaActual - 1) * TAMANO_PAGINA + 1
    cuantos = registros - primero + 1
    If cuantos > TAMANO_PAGINA Then cuantos = TAMANO_PAGINA
    cuadricula.Rows = cuantos + 1
    datos.AbsolutePosition = primero
    For fila = 1 To cuantos
        For columna = 0 To datos.Fields.Count - 1
            valor = datos.Fields(columna).Value
            If IsNull(valor) Then
                cuadricula.TextMatrix(fila, columna + 1) = vbNullString
            ElseIf columna = 1 Then
                cuadricula.TextMatrix(fila, columna + 1) = Format$(CDate(valor), "dd-mm-yyyy")
            ElseIf columna = 2 Then
                cuadricula.TextMatrix(fila, columna + 1) = Format$(CDate(valor), "hh:nn")
            Else
                cuadricula.TextMatrix(fila, columna + 1) = CStr(valor)
            End If
        Next columna
        datos.MoveNext
    Next fila
    cuadricula.Row = 1
    cuadricula.TopRow = 1
    cuadricula.Redraw = True
    Exit Sub
Fallo:
    mensaje = Err.Description
    cuadricula.Redraw = True
    Err.Raise vbObjectError + 2301, , mensaje
End Sub

Private Sub Class_Terminate()
    Limpiar
End Sub
