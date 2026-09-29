VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "FacturaPendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private datos As New PropertyBag

Public Sub Capturar(ByVal ventana As Facturacion)
    Dim i As Long, r As Long, c As Long
    Set datos = New PropertyBag
    datos.WriteProperty "Version", 1
    For i = 1 To 21
        If i <> 12 Then datos.WriteProperty "Text" & i, ventana.Controls("Text" & i).Text
    Next i
    datos.WriteProperty "Producto", ventana.Combo1.Text
    datos.WriteProperty "Descuento", CInt(ventana.Check1.Value)
    datos.WriteProperty "Fecha", CDate(ventana.DTPicker1.Value)
    datos.WriteProperty "Hora", CDate(ventana.DTPicker2.Value)
    With ventana.MSFlexGrid1
        datos.WriteProperty "Filas", CLng(.Rows)
        datos.WriteProperty "Fila", CLng(.Row)
        datos.WriteProperty "Columna", CLng(.Col)
        datos.WriteProperty "PrimeraFila", CLng(.TopRow)
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                datos.WriteProperty "Celda" & r & "_" & c, .TextMatrix(r, c)
            Next c
        Next r
    End With
End Sub

Public Sub Restaurar(ByVal ventana As Facturacion)
    Dim i As Long, r As Long, c As Long
    'El formulario suspende los eventos de calculo mientras restaura.
    ventana.Check1.Value = datos.ReadProperty("Descuento", 0)
    ventana.Text21.Enabled = (ventana.Check1.Value = vbChecked)
    For i = 1 To 21
        If i <> 12 Then ventana.Controls("Text" & i).Text = datos.ReadProperty("Text" & i, "")
    Next i
    ventana.Combo1.Text = datos.ReadProperty("Producto", "")
    ventana.DTPicker1.Value = datos.ReadProperty("Fecha", Date)
    ventana.DTPicker2.Value = datos.ReadProperty("Hora", Time)
    With ventana.MSFlexGrid1
        .Rows = datos.ReadProperty("Filas", 2)
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .TextMatrix(r, c) = datos.ReadProperty("Celda" & r & "_" & c, "")
            Next c
        Next r
        .Row = datos.ReadProperty("Fila", 0)
        .Col = datos.ReadProperty("Columna", 1)
        .TopRow = datos.ReadProperty("PrimeraFila", 1)
    End With
End Sub

Public Property Get Descripcion() As String
    Descripcion = datos.ReadProperty("Text11", "CLIENTE CONTADO") & _
        " | Productos: " & datos.ReadProperty("Text10", "0") & _
        " | Total: " & datos.ReadProperty("Text9", "0")
End Property

Public Sub Guardar(ByVal ruta As String)
    Dim archivo As Integer, contenido() As Byte
    Dim numeroError As Long, mensaje As String
    On Error GoTo Fallo
    contenido = datos.Contents
    archivo = FreeFile
    Open ruta & ".tmp" For Binary Access Write As #archivo
    Put #archivo, , contenido
    Close #archivo
    archivo = 0
    Name ruta & ".tmp" As ruta
    Exit Sub
Fallo:
    numeroError = Err.Number
    mensaje = Err.Description
    On Error Resume Next
    If archivo <> 0 Then Close #archivo
    On Error GoTo 0
    Err.Raise numeroError, "FacturaPendiente.Guardar", mensaje
End Sub

Public Sub Leer(ByVal ruta As String)
    Dim archivo As Integer, contenido() As Byte, filas As Long
    Dim numeroError As Long, mensaje As String
    On Error GoTo Fallo
    archivo = FreeFile
    Open ruta For Binary Access Read As #archivo
    If LOF(archivo) = 0 Then Err.Raise 5, , "El archivo pendiente esta vacio."
    ReDim contenido(0 To LOF(archivo) - 1)
    Get #archivo, , contenido
    Close #archivo
    archivo = 0
    Set datos = New PropertyBag
    datos.Contents = contenido
    If datos.ReadProperty("Version", 0) <> 1 Then Err.Raise 5, , "Formato pendiente no reconocido."
    filas = datos.ReadProperty("Filas", 0)
    If filas < 2 Or filas > 32767 Then Err.Raise 5, , "Cantidad de filas no valida."
    Exit Sub
Fallo:
    numeroError = Err.Number
    mensaje = Err.Description
    On Error Resume Next
    If archivo <> 0 Then Close #archivo
    On Error GoTo 0
    Err.Raise numeroError, "FacturaPendiente.Leer", mensaje
End Sub
