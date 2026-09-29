Attribute VB_Name = "PruebaPendientes"
Option Explicit

Public Sub Main()
    Dim ventana As Facturacion, i As Long, r As Long, c As Long
    Dim textos(1 To 21) As String, celdas(0 To 3, 0 To 8) As String
    Dim numero As String, producto As String, resultado As String
    On Error GoTo Fallo
    InicializarRutaBD
    Facturacion.Show
    Facturacion.Show
    Exigir Forms.Count = 1, "Una sola ventana"
    Set ventana = Facturacion
    numero = ventana.Text12.Text
    Agregar ventana, "101", 2
    Agregar ventana, "303", 1
    ventana.Text1.Text = "505"
    ventana.Combo1.Text = "Producto sin agregar, con comas y acentos: caf" & Chr$(233)
    ventana.Text2.Text = "2"
    ventana.Text4.Text = "12"
    ventana.Text5.Text = "100"
    ventana.Text6.Text = "5"
    ventana.Check1.Value = vbChecked
    ventana.Text21.Text = "2"
    ventana.Text16.Text = "50"
    For i = 1 To 21
        textos(i) = ventana.Controls("Text" & i).Text
    Next i
    producto = ventana.Combo1.Text
    For r = 0 To 3
        For c = 0 To 8
            celdas(r, c) = ventana.MSFlexGrid1.TextMatrix(r, c)
        Next c
    Next r
    ventana.cmdGuardarPendiente.Value = True
    Exigir ventana.cboPendientes.ListCount = 1, "Primera pendiente en disco"
    Exigir ventana.MSFlexGrid1.Rows = 2 And ventana.Text1.Text = "", "Nueva factura limpia"
    Exigir ventana.Text21.Text = "" And ventana.Check1.Value = 0, "Descuento limpio"
    Agregar ventana, "202", 4
    ventana.cmdGuardarPendiente.Value = True
    Exigir ventana.cboPendientes.ListCount = 2, "Dos pendientes"
    Unload ventana
    Set ventana = Nothing
    Facturacion.Show
    Set ventana = Facturacion
    Exigir ventana.cboPendientes.ListCount = 2, "Persistencia al reabrir"
    ventana.Text12.Text = "99999999"
    ventana.cboPendientes.ListIndex = 0
    ventana.cmdRecuperarPendiente.Value = True
    Exigir ventana.Text12.Text = numero, "Correlativo disponible y no almacenado"
    Exigir ventana.cboPendientes.ListCount = 1, "Pendiente retirada al recuperar"
    For i = 1 To 21
        If i <> 12 Then Exigir ventana.Controls("Text" & i).Text = textos(i), "Restaurar Text" & i
    Next i
    Exigir ventana.Check1.Value = vbChecked, "Casilla de descuento restaurada"
    Exigir ventana.Combo1.Text = producto, "Producto parcial restaurado"
    For r = 0 To 3
        For c = 0 To 8
            Exigir ventana.MSFlexGrid1.TextMatrix(r, c) = celdas(r, c), "Restaurar celda " & r & "," & c
        Next c
    Next r
    'Recuperar la otra guarda automaticamente el trabajo actual.
    ventana.cmdRecuperarPendiente.Value = True
    Exigir ventana.cboPendientes.ListCount = 1, "Intercambio sin perder actual"
    Exigir ventana.MSFlexGrid1.TextMatrix(1, 2) = "202", "Segunda factura recuperada"
    Agregar ventana, "606", 1
    Exigir ventana.MSFlexGrid1.Rows = 4, "Contador de filas restaurado"
    Exigir Val(ventana.Text9.Text) = 50, "Total al continuar pendiente"
    Unload ventana
    Set ventana = Nothing
    Facturacion.Show
    Set ventana = Facturacion
    Exigir ventana.cboPendientes.ListCount = 2, "Guardado automatico al cerrar"
    Unload ventana
    resultado = "OK: ventana unica, dos pendientes persistentes, restauracion de todos los campos y celdas, producto incompleto y descuento, correlativo actual, intercambio automatico, continuidad y cierre. Sin guardar ventas."
    GoTo Terminar
Fallo:
    resultado = "ERROR: " & Err.Description
Terminar:
    Open App.Path & "\resultado-pendientes.txt" For Output As #1
    Print #1, resultado
    Close #1
    End
End Sub

Private Sub Exigir(ByVal condicion As Boolean, ByVal mensaje As String)
    If Not condicion Then Err.Raise 5, , mensaje
End Sub

Private Sub Agregar(ByVal ventana As Facturacion, ByVal codigo As String, ByVal cantidad As Integer)
    ventana.Text1.Text = codigo
    ventana.Combo1.Text = "Producto de prueba"
    ventana.Text2.Text = CStr(cantidad)
    ventana.Text3.Text = "0"
    ventana.Text4.Text = "10"
    ventana.Text5.Text = "100"
    ventana.Text6.Text = "5"
    ventana.Command1.Value = True
End Sub
