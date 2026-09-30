VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ControlInventario 
   BackColor       =   &H00404000&
   Caption         =   "Control de Inventario"
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14880
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Guardar cambios"
      Height          =   480
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   2400
   End
   Begin VB.CommandButton cmdRecargar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Recargar"
      Height          =   480
      Left            =   2820
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1900
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Cerrar"
      Height          =   480
      Left            =   12540
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   2100
   End
   Begin MSFlexGridLib.MSFlexGrid grdProductos 
      Height          =   5040
      Left            =   240
      TabIndex        =   0
      Top             =   1710
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   8890
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FocusRect       =   0
      AllowUserResizing=   1
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1710
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.ComboBox cboRespaldos 
      Height          =   375
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   7200
      Width           =   4080
   End
   Begin VB.CommandButton cmdRestaurar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Revisar / &Recuperar"
      Height          =   480
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7140
      Width           =   3120
   End
   Begin VB.CommandButton cmdDescartar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Descartar"
      Height          =   480
      Left            =   4900
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Restablecer las cantidades nuevas con las cantidades actuales de la tabla"
      Top             =   720
      Width           =   1900
   End
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Generar &Excel"
      Height          =   480
      Left            =   6980
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Compara exclusivamente los archivos antiguos y nuevos seleccionados"
      Top             =   720
      Width           =   2400
   End
   Begin VB.CommandButton cmdBorrarRespaldos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Borrar respaldos"
      Height          =   480
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Eliminar todos los respaldos de inventario, previa confirmación"
      Top             =   7140
      Width           =   3120
   End
   Begin VB.CommandButton cmdDatosAntiguos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Guardar datos antiguos"
      Height          =   480
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6780
      Width           =   3360
   End
   Begin VB.CommandButton cmdDatosNuevos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Guardar datos nuevos"
      Height          =   480
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6780
      Width           =   3360
   End
   Begin VB.Label lblAyuda 
      BackStyle       =   0  'Transparent
      Caption         =   "Edite Cantidad nueva. Enter / flechas arriba y abajo: otro producto. Esc: deshacer la celda. Ctrl+S: guardar."
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   14400
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   240
      TabIndex        =   13
      Top             =   6300
      Width           =   14400
   End
   Begin VB.Label lblRespaldo 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos guardados: seleccione datosantiguos o datosnuevos para revisar / recuperar."
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   14
      Top             =   6810
      Width           =   14400
   End
   Begin VB.Label lblNota 
      BackStyle       =   0  'Transparent
      Caption         =   $"ControlInventario.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   7680
      Width           =   14400
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONTROL DE INVENTARIO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   240
      TabIndex        =   16
      Top             =   150
      Width           =   14400
   End
   Begin VB.Label lblExcelAntiguos
      Caption = "Excel: datos antiguos"
      BackStyle = 0
      ForeColor = &H00FFFFFF&
      Left = 240
      Top = 6600
      Width = 6000
      Height = 300
   End
   Begin VB.ComboBox cboExcelAntiguos
      Style = 2
      Left = 240
      Top = 6900
      Width = 7000
      Height = 360
      TabIndex = 12
   End
   Begin VB.Label lblExcelNuevos
      Caption = "Excel: datos nuevos"
      BackStyle = 0
      ForeColor = &H00FFFFFF&
      Left = 7560
      Top = 6600
      Width = 6000
      Height = 300
   End
   Begin VB.ComboBox cboExcelNuevos
      Style = 2
      Left = 7560
      Top = 6900
      Width = 7000
      Height = 360
      TabIndex = 13
   End
   Begin VB.Image imgFondo 
      Height          =   8700
      Left            =   0
      Picture         =   "ControlInventario.frx":0088
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14880
   End
End
Attribute VB_Name = "ControlInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private cn As ADODB.Connection
Private originales() As Variant
Private totalFilas As Long
Private filaEditada As Long
Private cargando As Boolean
Private ocupado As Boolean
Private enganchado As Boolean
Private respaldoRecuperado As String
Private editados As Object
Private respaldoEdicion As String
Private respaldoNuevo As String

Private Sub Form_Load()
    respaldoEdicion = vbNullString
    respaldoNuevo = vbNullString
    On Error GoTo Fallo
    cargando = True
    Set editados = CreateObject("Scripting.Dictionary")
    editados.CompareMode = vbTextCompare
    imgFondo.ZOrder 1
    lblAyuda.Font.Size = 9.75
    lblNota.Font.Size = 9.75
    If Len(rutaBD) = 0 Then InicializarRutaBD
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
    PrepararVinculaciones cn
    With grdProductos
        .Font.Name = "Times New Roman"
        .Font.Size = 9.75
        .Font.Bold = True
        .BackColorFixed = &HFFFFC0
        .ColWidth(0) = 1600
        .ColWidth(1) = 4600
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .RowHeightMin = 360
        .TextMatrix(0, 0) = "Código"
        .TextMatrix(0, 1) = "Producto"
        .TextMatrix(0, 2) = "Cantidad actual"
        .TextMatrix(0, 3) = "Cantidad nueva"
        .TextMatrix(0, 4) = "Precio venta"
        .TextMatrix(0, 5) = "Precio costo"
        .TextMatrix(0, 6) = "Fecha"
    End With
    AjustarColumnasProductos
    CargarProductos
    CargarRespaldos
    WheelHook Me
    enganchado = True
    Exit Sub
Fallo:
    MsgBox "No se pudo abrir Control de Inventario: " & Err.Description, vbExclamation
    Unload Me
End Sub

Private Function Texto(ByVal valor As Variant) As String
    If Not IsNull(valor) Then Texto = CStr(valor)
End Function

Private Function Iguales(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) Or IsNull(b) Then
        Iguales = IsNull(a) And IsNull(b)
    Else
        Iguales = (CCur(a) = CCur(b))
    End If
End Function

Private Function LeerProductos() As ADODB.Recordset
    Dim datos As New ADODB.Recordset
    datos.CursorLocation = adUseClient
    datos.Open "SELECT * FROM productos WHERE codproducto NOT IN (SELECT codventa FROM presentaciones) ORDER BY nombreprod, codproducto", cn, adOpenStatic, adLockReadOnly
    Set datos.ActiveConnection = Nothing
    Set LeerProductos = datos
End Function

Private Function IndiceProductos(ByVal datos As ADODB.Recordset) As Object
    Dim indice As Object, codigo As String
    Set indice = CreateObject("Scripting.Dictionary")
    indice.CompareMode = vbTextCompare
    If Not datos.EOF Then datos.MoveFirst
    Do While Not datos.EOF
        codigo = Texto(datos!codproducto)
        If Len(codigo) = 0 Then Err.Raise 5, , "Hay un producto sin código."
        If indice.Exists(codigo) Then Err.Raise 5, , "Código de producto duplicado: " & codigo
        indice.Add codigo, datos!cantprod.Value
        datos.MoveNext
    Loop
    Set IndiceProductos = indice
End Function

Private Sub CargarProductos(Optional ByVal conservarEdiciones As Boolean = False)
    Dim datos As ADODB.Recordset, indice As Object, i As Long
    Dim r As Long, c As Long, valores() As String, bases() As Variant
    Dim pendientes As Object, codigo As String, seleccion As String, clave As Variant
    Set pendientes = CreateObject("Scripting.Dictionary")
    pendientes.CompareMode = vbTextCompare
    If conservarEdiciones Then
        If filaEditada > 0 Then seleccion = grdProductos.TextMatrix(filaEditada, 0)
        For i = 1 To totalFilas
            codigo = grdProductos.TextMatrix(i, 0)
            If editados.Exists(codigo) Then pendientes.Add codigo, ValorFila(i)
        Next i
    End If
    Set datos = LeerProductos()
    Set indice = IndiceProductos(datos)
    For Each clave In pendientes.Keys
        If Not indice.Exists(CStr(clave)) Then
            Err.Raise 5, , "El producto editado " & clave & " ya no existe. No se recargó la lista para conservar sus ediciones."
        End If
    Next clave
    If datos.RecordCount > 32766 Then Err.Raise 5, , "La tabla supera el límite de filas de esta cuadrícula."
    r = datos.RecordCount
    ReDim valores(0 To r, 0 To 6)
    ReDim bases(0 To r)
    If r > 0 Then datos.MoveFirst
    For i = 1 To r
        bases(i) = datos!cantprod.Value
        codigo = Texto(datos!codproducto)
        valores(i, 0) = codigo
        valores(i, 1) = Texto(datos!nombreprod)
        valores(i, 2) = Texto(datos!cantprod)
        valores(i, 3) = Texto(datos!cantprod)
        If pendientes.Exists(codigo) Then valores(i, 3) = Texto(pendientes(codigo))
        valores(i, 4) = Texto(datos!preciov)
        valores(i, 5) = Texto(datos!precioc)
        valores(i, 6) = Texto(datos!fechav)
        datos.MoveNext
    Next i
    datos.Close
    cargando = True
    txtCantidad.Visible = False
    filaEditada = 0
    totalFilas = r
    originales = bases
    grdProductos.Rows = 2
    If r > 0 Then grdProductos.Rows = r + 1
    For i = 1 To grdProductos.Rows - 1
        For c = 0 To 6
            grdProductos.TextMatrix(i, c) = ""
            If i <= r Then grdProductos.TextMatrix(i, c) = valores(i, c)
        Next c
    Next i
    grdProductos.Col = 3
    grdProductos.Row = 1
    For i = 1 To r
        If valores(i, 0) = seleccion Then grdProductos.Row = i
    Next i
    Set editados = pendientes
    cargando = False
    If Not conservarEdiciones Then respaldoRecuperado = vbNullString
    If r > 0 Then
        If Not grdProductos.RowIsVisible(grdProductos.Row) Then grdProductos.TopRow = grdProductos.Row
        MostrarEditor
    End If
    ActualizarEstado
End Sub

Private Function CantidadValida(ByVal entrada As String, ByVal fila As Long, ByRef valor As Variant) As Boolean
    On Error GoTo Invalida
    entrada = Trim$(entrada)
    If Len(entrada) = 0 Then
        valor = Null
        CantidadValida = True
        Exit Function
    End If
    If Len(entrada) = 0 Or Not IsNumeric(entrada) Then Exit Function
    valor = CCur(entrada)
    If CDec(entrada) <> CDec(valor) Then Exit Function
    CantidadValida = True
Invalida:
End Function

Private Function ValorFila(ByVal fila As Long) As Variant
    Dim valor As Variant
    If Not CantidadValida(grdProductos.TextMatrix(fila, 3), fila, valor) Then
        Err.Raise 5, , "Cantidad no válida en " & grdProductos.TextMatrix(fila, 0)
    End If
    ValorFila = valor
End Function

Private Function TieneRespaldoAntiguo() As Boolean
    If Len(respaldoEdicion) = 0 Then Exit Function
    TieneRespaldoAntiguo = (Len(Dir$(respaldoEdicion)) > 0)
End Function

Private Function AceptarEdicion() As Boolean
    Dim valor As Variant
    AceptarEdicion = True
    If cargando Or filaEditada = 0 Then Exit Function
    If Not CantidadValida(txtCantidad.Text, filaEditada, valor) Then
        AceptarEdicion = False
        MsgBox "Escriba una cantidad numérica con un máximo de cuatro decimales (vacío = sin dato).", vbExclamation, "Cantidad"
        txtCantidad.Visible = True
        txtCantidad.SetFocus
        Exit Function
    End If
    If Not Iguales(ValorFila(filaEditada), valor) Then editados(grdProductos.TextMatrix(filaEditada, 0)) = True
    grdProductos.TextMatrix(filaEditada, 3) = Texto(valor)
    ActualizarEstado
End Function

Private Function Cambios() As Long
    Dim i As Long
    For i = 1 To totalFilas
        If Not Iguales(originales(i), ValorFila(i)) Then Cambios = Cambios + 1
    Next i
End Function

Private Function HayEdiciones() As Boolean
    If totalFilas = 0 Then Exit Function
    HayEdiciones = (Cambios > 0)
    If filaEditada > 0 Then
        If txtCantidad.Text <> grdProductos.TextMatrix(filaEditada, 3) Then HayEdiciones = True
    End If
End Function

Private Sub ActualizarEstado()
    ColorearCantidades
    lblEstado.Caption = totalFilas & " productos con inventario propio | " & Cambios & " cantidades modificadas sin guardar"
    If Not TieneRespaldoAntiguo Then lblEstado.Caption = lblEstado.Caption & " | Antes de guardar, guarde datos antiguos."

End Sub

Private Sub MostrarEditor()
    If cargando Or ocupado Or totalFilas = 0 Then Exit Sub
    If grdProductos.Row < 1 Then Exit Sub
    filaEditada = grdProductos.Row
    cargando = True
    grdProductos.Col = 3
    txtCantidad.Text = grdProductos.TextMatrix(filaEditada, 3)
    cargando = False
    PosicionarEditor
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
End Sub

Private Sub PosicionarEditor()
    If filaEditada = 0 Or ocupado Then txtCantidad.Visible = False: Exit Sub
    With grdProductos
        txtCantidad.Visible = .RowIsVisible(filaEditada) And .ColIsVisible(3)
        If txtCantidad.Visible Then
            txtCantidad.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
            txtCantidad.ZOrder 0
        End If
    End With
End Sub

Private Sub grdProductos_RowColChange()
    If cargando Or ocupado Or totalFilas = 0 Then Exit Sub
    If Not AceptarEdicion Then
        cargando = True
        grdProductos.Row = filaEditada
        grdProductos.Col = 3
        cargando = False
        PosicionarEditor
        Exit Sub
    End If
    MostrarEditor
End Sub

Private Sub grdProductos_Click()
    If txtCantidad.Visible Then txtCantidad.SetFocus
End Sub

Private Sub grdProductos_Scroll()
    If Not cargando Then PosicionarEditor
End Sub

Private Sub MoverFila(ByVal destino As Long)
    If Not AceptarEdicion Then Exit Sub
    If destino < 1 Then destino = 1
    If destino > totalFilas Then destino = totalFilas
    If destino < 1 Then Exit Sub
    grdProductos.Row = destino
    If Not grdProductos.RowIsVisible(destino) Then grdProductos.TopRow = destino
    MostrarEditor
    If txtCantidad.Visible Then txtCantidad.SetFocus
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            KeyCode = 0
            MoverFila filaEditada + 1
        Case vbKeyUp
            KeyCode = 0
            MoverFila filaEditada - 1
        Case vbKeyPageDown
            KeyCode = 0
            MoverFila filaEditada + 10
        Case vbKeyPageUp
            KeyCode = 0
            MoverFila filaEditada - 10
        Case vbKeyEscape
            KeyCode = 0
            txtCantidad.Text = grdProductos.TextMatrix(filaEditada, 3)
            txtCantidad.SelStart = 0
            txtCantidad.SelLength = Len(txtCantidad.Text)
    End Select
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then KeyAscii = 0
End Sub

Private Sub grdProductos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyF2 Then
        KeyCode = 0
        MostrarEditor
        If txtCantidad.Visible Then txtCantidad.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyS And (Shift And vbCtrlMask) <> 0 Then
        KeyCode = 0
        cmdGuardar_Click
    End If
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    If ocupado Or totalFilas = 0 Then Exit Sub
    If Not AceptarEdicion Then Exit Sub
    FlexGridScroll grdProductos, MouseKeys, Rotation, Xpos, Ypos
    PosicionarEditor
End Sub

Private Function CarpetaRespaldos() As String
    CarpetaRespaldos = App.Path & "\datos\respaldos-inventario"
End Function

Private Sub CargarRespaldos()
    Dim archivo As String, anterior As String, nueva As String, i As Long
    anterior = cboExcelAntiguos.Text
    nueva = cboExcelNuevos.Text
    cboRespaldos.Clear
    cboExcelAntiguos.Clear
    cboExcelNuevos.Clear
    If Len(Dir$(CarpetaRespaldos, vbDirectory)) > 0 Then
        archivo = Dir$(CarpetaRespaldos & "\*.adtg")
        Do While Len(archivo) > 0
            cboRespaldos.AddItem NombreArchivoVisible(archivo)
            If LCase$(Left$(archivo, 14)) = "datosantiguos-" Then cboExcelAntiguos.AddItem NombreArchivoVisible(archivo)
            If LCase$(Left$(archivo, 12)) = "datosnuevos-" Then cboExcelNuevos.AddItem NombreArchivoVisible(archivo)
            archivo = Dir$()
        Loop
    End If
    For i = 0 To cboExcelAntiguos.ListCount - 1
        If cboExcelAntiguos.List(i) = anterior Then cboExcelAntiguos.ListIndex = i
    Next i
    For i = 0 To cboExcelNuevos.ListCount - 1
        If cboExcelNuevos.List(i) = nueva Then cboExcelNuevos.ListIndex = i
    Next i
    If cboRespaldos.ListCount > 0 Then cboRespaldos.ListIndex = cboRespaldos.ListCount - 1
    cmdRestaurar.Enabled = (cboRespaldos.ListCount > 0)
    PosicionarEditor
    ActualizarEstado
End Sub

Private Function GuardarRespaldo(ByVal datos As ADODB.Recordset, ByVal tipo As String) As String
    Dim base As String, ruta As String, N As Long
    If Len(Dir$(CarpetaRespaldos, vbDirectory)) = 0 Then MkDir CarpetaRespaldos
    base = CarpetaRespaldos & "\" & tipo & "-" & Format$(Now, "dd-mm-yyyy hh-nn")
    Do
        N = N + 1
        ruta = base & "-" & Format$(N, "000") & ".adtg"
    Loop While Len(Dir$(ruta)) > 0 Or Len(Dir$(ruta & ".tmp")) > 0
    If datos.RecordCount > 0 Then datos.MoveFirst
    datos.Save ruta & ".tmp", adPersistADTG
    Name ruta & ".tmp" As ruta
    GuardarRespaldo = ruta
End Function

Private Sub cmdDatosAntiguos_Click()
    GuardarDatosManual False
End Sub

Private Sub cmdDatosNuevos_Click()
    GuardarDatosManual True
End Sub

Private Function DatosNuevosPantalla() As ADODB.Recordset
    Dim datos As New ADODB.Recordset, i As Long
    datos.CursorLocation = adUseClient
    datos.Fields.Append "codproducto", adVarWChar, 255
    datos.Fields.Append "nombreprod", adVarWChar, 255
    datos.Fields.Append "cantprod", adCurrency, , adFldIsNullable
    datos.Open
    For i = 1 To totalFilas
        datos.AddNew
        datos!codproducto = grdProductos.TextMatrix(i, 0)
        datos!nombreprod = grdProductos.TextMatrix(i, 1)
        datos!cantprod = ValorFila(i)
        datos.Update
    Next i
    Set DatosNuevosPantalla = datos
End Function

Private Sub GuardarDatosManual(ByVal nuevos As Boolean)
    Dim datos As ADODB.Recordset, ruta As String, tipo As String, i As Long
    On Error GoTo Fallo
    If ocupado Then Exit Sub
    If nuevos Then
        If Not TieneRespaldoAntiguo Then
            MsgBox "Primero guarde una copia de datos antiguos antes de guardar datos nuevos.", vbExclamation
            Exit Sub
        End If
    End If
    If Not AceptarEdicion Then Exit Sub
    ocupado = True
    Screen.MousePointer = vbHourglass
    If nuevos Then
        tipo = "datosnuevos"
        Set datos = DatosNuevosPantalla()
    Else
        tipo = "datosantiguos"
        Set datos = LeerProductos()
        ValidarDatosActuales datos
    End If
    ruta = GuardarRespaldo(datos, tipo)
    datos.Close
    If nuevos Then
        respaldoNuevo = ruta
    Else
        respaldoEdicion = ruta
        respaldoNuevo = vbNullString
    End If
    CargarRespaldos
    For i = 0 To cboRespaldos.ListCount - 1
        If CarpetaRespaldos & "\" & NombreArchivoReal(cboRespaldos.List(i)) = ruta Then cboRespaldos.ListIndex = i
    Next i
    ocupado = False
    Screen.MousePointer = vbDefault
    If nuevos Then
        For i = 0 To cboExcelNuevos.ListCount - 1
            If CarpetaRespaldos & "\" & NombreArchivoReal(cboExcelNuevos.List(i)) = ruta Then cboExcelNuevos.ListIndex = i
        Next i
    Else
        For i = 0 To cboExcelAntiguos.ListCount - 1
            If CarpetaRespaldos & "\" & NombreArchivoReal(cboExcelAntiguos.List(i)) = ruta Then cboExcelAntiguos.ListIndex = i
        Next i
    End If
    MostrarEditor
    ActualizarEstado
    MsgBox "Datos guardados en la lista: " & Mid$(ruta, InStrRev(ruta, "\") + 1) & vbCrLf & "Las cantidades de la base de datos no se han modificado.", vbInformation, "Guardar " & tipo
    Exit Sub
Fallo:
    ocupado = False
    Screen.MousePointer = vbDefault
    MsgBox "No se pudieron guardar los datos: " & Err.Description, vbExclamation, "Control de Inventario"
End Sub

Private Sub ActualizarCantidad(ByVal codigo As String, ByVal anterior As Variant, ByVal nueva As Variant)
    Dim orden As New ADODB.Command, afectados As Variant
    Set orden.ActiveConnection = cn
    orden.CommandType = adCmdText
    orden.CommandText = "UPDATE productos SET cantprod = ? WHERE codproducto = ? AND codproducto NOT IN (SELECT codventa FROM presentaciones)"
    orden.Parameters.Append orden.CreateParameter("nueva", adCurrency, adParamInput, , nueva)
    orden.Parameters.Append orden.CreateParameter("codigo", adVarWChar, adParamInput, 255, codigo)
    If IsNull(anterior) Then
        orden.CommandText = orden.CommandText & " AND cantprod IS NULL"
    Else
        orden.CommandText = orden.CommandText & " AND cantprod = ?"
        orden.Parameters.Append orden.CreateParameter("anterior", adCurrency, adParamInput, , anterior)
    End If
    orden.Execute afectados, , adExecuteNoRecords
    If CLng(afectados) <> 1 Then
        Err.Raise vbObjectError + 1100, , "El producto " & codigo & " cambió o fue eliminado. Recargue antes de volver a guardar."
    End If
End Sub

Private Sub ValidarDatosActuales(ByVal datos As ADODB.Recordset)
    Dim indice As Object, filas As Object, i As Long, codigo As String
    Set indice = IndiceProductos(datos)
    If indice.Count <> totalFilas Then Err.Raise 5, , "Se agregaron, eliminaron o vincularon productos fuera de esta ventana. Recargue antes de guardar."
    Set filas = CreateObject("Scripting.Dictionary")
    filas.CompareMode = vbTextCompare
    For i = 1 To totalFilas
        codigo = grdProductos.TextMatrix(i, 0)
        If Not indice.Exists(codigo) Then Err.Raise 5, , "El producto " & codigo & " cambió fuera de esta ventana. Recargue antes de guardar."
        If Not Iguales(originales(i), indice(codigo)) Then Err.Raise 5, , "La cantidad del producto " & codigo & " cambió por una factura u otra operación. Recargue antes de guardar."
        filas.Add codigo, i
    Next i
    If Not datos.EOF Or datos.RecordCount > 0 Then datos.MoveFirst
    Do While Not datos.EOF
        codigo = Texto(datos!codproducto)
        i = CLng(filas(codigo))
        If Texto(datos!nombreprod) <> grdProductos.TextMatrix(i, 1) Or _
           Texto(datos!preciov) <> grdProductos.TextMatrix(i, 4) Or _
           Texto(datos!precioc) <> grdProductos.TextMatrix(i, 5) Or _
           Texto(datos!fechav) <> grdProductos.TextMatrix(i, 6) Then
            Err.Raise 5, , "Los datos del producto " & codigo & " cambiaron fuera de esta ventana. Recargue antes de guardar."
        End If
        datos.MoveNext
    Loop
End Sub

Private Sub ValidarRespaldoNuevo()
    Dim datos As New ADODB.Recordset, indice As Object
    If Len(respaldoNuevo) = 0 Then Err.Raise 5, , "Primero guarde una copia de datos nuevos antes de aplicar los cambios."
    If Len(Dir$(respaldoNuevo)) = 0 Then Err.Raise 5, , "La copia de datos nuevos ya no existe. Guarde datos nuevos otra vez."
    datos.Open respaldoNuevo, , adOpenStatic, adLockReadOnly
    Set indice = IndiceProductos(datos)
    datos.Close
    ValidarCantidadesRespaldadas indice
End Sub

Private Sub ValidarCantidadesRespaldadas(ByVal indice As Object)
    Dim i As Long, codigo As String
    If indice.Count <> totalFilas Then Err.Raise 5, , "La lista cambió desde el respaldo. Guarde datos nuevos otra vez."
    For i = 1 To totalFilas
        codigo = grdProductos.TextMatrix(i, 0)
        If Not indice.Exists(codigo) Then Err.Raise 5, , "La lista cambió desde el respaldo. Guarde datos nuevos otra vez."
        If Not Iguales(ValorFila(i), indice(codigo)) Then Err.Raise 5, , "Hay cantidades nuevas modificadas después del respaldo. Guarde datos nuevos otra vez antes de aplicar los cambios."
    Next i
End Sub

Private Function ColorCantidad(ByVal valor As String) As Long
    ColorCantidad = grdProductos.ForeColor
    If IsNumeric(valor) Then
        If CDbl(valor) < 0 Then ColorCantidad = vbRed
    End If
End Function

Private Sub txtCantidad_Change()
    txtCantidad.ForeColor = ColorCantidad(txtCantidad.Text)
End Sub

Private Sub ColorearCantidades()
    Dim fila As Long, columna As Long, filaAnterior As Long, columnaAnterior As Long
    Dim seleccionFila As Long, seleccionColumna As Long, estabaCargando As Boolean
    If totalFilas = 0 Then Exit Sub
    estabaCargando = cargando
    cargando = True
    With grdProductos
        filaAnterior = .Row
        columnaAnterior = .Col
        seleccionFila = .RowSel
        seleccionColumna = .ColSel
        .Redraw = False
        For fila = 1 To totalFilas
            .Row = fila
            For columna = 2 To 3
                .Col = columna
                .CellForeColor = ColorCantidad(.TextMatrix(fila, columna))
            Next columna
        Next fila
        .Row = filaAnterior
        .Col = columnaAnterior
        .RowSel = seleccionFila
        .ColSel = seleccionColumna
        .Redraw = True
    End With
    cargando = estabaCargando
    txtCantidad.ForeColor = ColorCantidad(txtCantidad.Text)
End Sub

Private Sub cmdGuardar_Click()
    Dim i As Long, numero As Long, datos As ADODB.Recordset
    Dim codigo As String, mensaje As String, aviso As String
    Dim transaccion As Boolean, guardado As Boolean
    If ocupado Then Exit Sub
    On Error GoTo Fallo
    If Not TieneRespaldoAntiguo Then
        MsgBox "Primero guarde una copia de datos antiguos.", vbExclamation
        Exit Sub
    End If
    If Not AceptarEdicion Then Exit Sub
    numero = Cambios
    If numero = 0 And Len(respaldoRecuperado) = 0 Then
        MsgBox "No hay cantidades modificadas.", vbInformation
        Exit Sub
    End If
    ValidarRespaldoNuevo
    mensaje = vbNullString
    If Len(respaldoRecuperado) > 0 Then mensaje = vbCrLf & "Al completar la recuperación se eliminará el respaldo seleccionado."
    If MsgBox("Se actualizarán " & numero & " cantidades." & mensaje & vbCrLf & "¿Guardar los cambios?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Control de Inventario") <> vbYes Then Exit Sub
    ocupado = True
    Screen.MousePointer = vbHourglass
    cn.BeginTrans
    transaccion = True
    Set datos = LeerProductos()
    ValidarDatosActuales datos
    For i = 1 To totalFilas
        If Not Iguales(originales(i), ValorFila(i)) Then
            ActualizarCantidad grdProductos.TextMatrix(i, 0), originales(i), ValorFila(i)
        End If
    Next i
    cn.CommitTrans
    transaccion = False
    guardado = True
    aviso = EliminarRespaldoRecuperado()
    ocupado = False
    Screen.MousePointer = vbDefault
    CargarProductos
    CargarRespaldos
    MsgBox numero & " cantidades guardadas." & aviso, vbInformation, "Control de Inventario"
    Exit Sub
Fallo:
    mensaje = Err.Description
    On Error Resume Next
    If transaccion Then cn.RollbackTrans
    ocupado = False
    cargando = False
    Screen.MousePointer = vbDefault
    CargarRespaldos
    If guardado Then
        MsgBox "Los cambios se guardaron, pero no se pudo refrescar la lista: " & mensaje, vbExclamation
    Else
        MsgBox "No se guardó ningún cambio. " & mensaje, vbExclamation, "Control de Inventario"
    End If
End Sub

Private Sub cmdRestaurar_Click()
    Dim datos As New ADODB.Recordset, indice As Object
    Dim i As Long, codigo As String, encontrados As Long, faltantes As Long, rutaRecuperada As String
    On Error GoTo Fallo
    If ocupado Or cboRespaldos.ListIndex < 0 Then Exit Sub
    Dim confirmacion As String
    confirmacion = "¿Recuperar las cantidades del respaldo " & cboRespaldos.Text & "?"
    If HayEdiciones Then confirmacion = confirmacion & vbCrLf & "Se reemplazarán las cantidades nuevas que tiene sin guardar."
    confirmacion = confirmacion & vbCrLf & "Se prepararán para revisión. La base no cambiará hasta presionar Guardar cambios."
    If MsgBox(confirmacion, vbQuestion Or vbYesNo Or vbDefaultButton2, "Confirmar recuperación") <> vbYes Then Exit Sub
    rutaRecuperada = CarpetaRespaldos & "\" & NombreArchivoReal(cboRespaldos.Text)
    datos.Open rutaRecuperada, , adOpenStatic, adLockReadOnly
    Set indice = IndiceProductos(datos)
    'Validar todo el archivo antes de sustituir las ediciones actuales.
    Dim clave As Variant, comprobacion As Currency
    For Each clave In indice.Keys
        If Not IsNull(indice(clave)) Then comprobacion = CCur(indice(clave))
    Next clave
    CargarProductos
    For i = 1 To totalFilas
        codigo = grdProductos.TextMatrix(i, 0)
        If indice.Exists(codigo) Then
            grdProductos.TextMatrix(i, 3) = Texto(indice(codigo))
            editados(codigo) = True
            encontrados = encontrados + 1
        End If
    Next i
    faltantes = indice.Count - encontrados
    If encontrados > 0 Then respaldoRecuperado = rutaRecuperada
    datos.Close
    MostrarEditor
    ActualizarEstado
    MsgBox "Cantidades preparadas para " & encontrados & " productos." & vbCrLf & faltantes & " productos del respaldo ya no existen y no se crearán." & vbCrLf & "Revise la columna Cantidad nueva y presione Guardar cambios para confirmar la recuperación y eliminar el respaldo utilizado.", vbInformation, "Recuperar inventario"
    Exit Sub
Fallo:
    cargando = False
    MsgBox "No se pudo revisar el respaldo: " & Err.Description, vbExclamation
End Sub

Private Sub cmdRecargar_Click()
    On Error GoTo Fallo
    If ocupado Then Exit Sub
    If Not AceptarEdicion Then Exit Sub
    CargarProductos True
    CargarRespaldos
    Exit Sub
Fallo:
    cargando = False
    MsgBox Err.Description, vbExclamation, "Recargar inventario"
End Sub

Private Sub cmdDescartar_Click()
    Dim i As Long
    If ocupado Then Exit Sub
    If MsgBox("¿Descartar las cantidades nuevas y restablecerlas con las cantidades actuales de la tabla?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Descartar ediciones") <> vbYes Then Exit Sub
    cargando = True
    For i = 1 To totalFilas
        grdProductos.TextMatrix(i, 3) = Texto(originales(i))
    Next i
    editados.RemoveAll
    respaldoRecuperado = vbNullString
    cargando = False
    MostrarEditor
    ActualizarEstado
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ocupado Then Cancel = 1: Exit Sub
    If cargando Then Exit Sub
    If HayEdiciones Then
        If MsgBox("Hay cantidades sin guardar. ¿Cerrar y descartarlas?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Control de Inventario") <> vbYes Then Cancel = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If enganchado Then WheelUnHook Me
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cn = Nothing
End Sub

Private Sub AjustarColumnasProductos()
    Dim disponible As Long, numerica As Long, codigo As Long, nombre As Long
    'Reservar bordes y barra vertical para mantener las cinco columnas visibles.
    disponible = grdProductos.Width - (GetSystemMetrics(2) + 4) * Screen.TwipsPerPixelX
    If disponible < 6000 Then Exit Sub
    numerica = CLng(disponible * 0.12)
    If numerica < 1500 Then numerica = 1500
    codigo = CLng((disponible - 3 * numerica) * 0.3)
    nombre = disponible - 3 * numerica - codigo
    With grdProductos
        .ColWidth(0) = codigo
        .ColWidth(1) = nombre
        .ColWidth(2) = numerica
        .ColWidth(3) = numerica
        .ColWidth(4) = numerica
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then Exit Sub
    If Width < 12600 Then Width = 12600
    If Height < 7400 Then Height = 7400
    grdProductos.Width = ScaleWidth - 480
    grdProductos.Height = ScaleHeight - 5070
    AjustarColumnasProductos
    imgFondo.Move 0, 0, ScaleWidth, ScaleHeight
    lblTitulo.Width = ScaleWidth - 480
    lblAyuda.Width = ScaleWidth - 480
    cmdCerrar.Left = ScaleWidth - cmdCerrar.Width - 240
    lblEstado.Top = ScaleHeight - 3240
    cmdDatosAntiguos.Top = ScaleHeight - 2760
    cmdDatosNuevos.Top = ScaleHeight - 2760
    lblEstado.Width = ScaleWidth - 480
    lblExcelAntiguos.Top = ScaleHeight - 2190
    lblExcelNuevos.Top = ScaleHeight - 2190
    cboExcelAntiguos.Top = ScaleHeight - 1860
    cboExcelNuevos.Top = ScaleHeight - 1860
    cboExcelAntiguos.Width = (ScaleWidth - 720) / 2
    cboExcelNuevos.Width = cboExcelAntiguos.Width
    cboExcelNuevos.Left = cboExcelAntiguos.Left + cboExcelAntiguos.Width + 240
    lblExcelNuevos.Left = cboExcelNuevos.Left
    lblRespaldo.Top = ScaleHeight - 1290
    cboRespaldos.Top = ScaleHeight - 900
    cboRespaldos.Width = ScaleWidth - 7080
    cmdRestaurar.Left = ScaleWidth - 6840
    cmdBorrarRespaldos.Left = ScaleWidth - 3480
    cmdBorrarRespaldos.Top = ScaleHeight - 960
    cmdRestaurar.Top = ScaleHeight - 960
    lblNota.Top = ScaleHeight - 420
    lblNota.Width = ScaleWidth - 480
    If Not cargando Then PosicionarEditor
End Sub

Private Function EliminarRespaldoRecuperado() As String
    On Error GoTo Fallo
    If Len(respaldoRecuperado) = 0 Then Exit Function
    'Consumir el respaldo solamente despues del CommitTrans exitoso.
    If Len(Dir$(respaldoRecuperado)) > 0 Then Kill respaldoRecuperado
    respaldoRecuperado = vbNullString
    EliminarRespaldoRecuperado = vbCrLf & "El respaldo recuperado se retiró de la lista."
    Exit Function
Fallo:
    EliminarRespaldoRecuperado = vbCrLf & "La recuperación se guardó, pero no se pudo eliminar el respaldo utilizado: " & Err.Description
End Function
Private Function InformeArchivos() As Variant
    Dim antiguos As New ADODB.Recordset, nuevos As New ADODB.Recordset
    Dim indice As Object, comprobacion As Object, filas() As Variant
    Dim i As Long, codigo As String
    If cboExcelAntiguos.ListIndex < 0 Or cboExcelNuevos.ListIndex < 0 Then
        Err.Raise 5, , "Seleccione un archivo de datos antiguos y otro de datos nuevos en las listas de Excel."
    End If
    antiguos.Open CarpetaRespaldos & "\" & NombreArchivoReal(cboExcelAntiguos.Text), , adOpenStatic, adLockReadOnly
    nuevos.Open CarpetaRespaldos & "\" & NombreArchivoReal(cboExcelNuevos.Text), , adOpenStatic, adLockReadOnly
    Set comprobacion = IndiceProductos(antiguos)
    Set indice = IndiceProductos(nuevos)
    If comprobacion.Count <> indice.Count Then Err.Raise 5, , "Los archivos no contienen los mismos productos. Seleccione dos copias del mismo control."
    If antiguos.RecordCount = 0 Then Exit Function
    ReDim filas(1 To antiguos.RecordCount, 1 To 5)
    antiguos.MoveFirst
    i = 1
    Do While Not antiguos.EOF
        codigo = Texto(antiguos!codproducto)
        If Not indice.Exists(codigo) Then Err.Raise 5, , "El producto " & codigo & " no está en datos nuevos. Seleccione copias del mismo control."
        filas(i, 1) = codigo
        filas(i, 2) = Texto(antiguos!nombreprod)
        If Not IsNull(antiguos!preciov) Then filas(i, 3) = antiguos!preciov.Value
        If Not IsNull(antiguos!cantprod) Then filas(i, 4) = antiguos!cantprod.Value
        If Not IsNull(indice(codigo)) Then filas(i, 5) = indice(codigo)
        antiguos.MoveNext
        i = i + 1
    Loop
    antiguos.Close
    nuevos.Close
    InformeArchivos = filas
End Function

Private Sub cmdExcel_Click()
    Dim salida As String, filas As Variant, mensaje As String
    On Error GoTo Fallo
    If ocupado Then Exit Sub
    filas = InformeArchivos()
    If Not IsArray(filas) Then
        MsgBox "No hay productos para exportar.", vbInformation
        Exit Sub
    End If
    ocupado = True
    salida = ElegirExcelInventario(Me.hWnd)
    If Len(salida) = 0 Then
        ocupado = False
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    GuardarExcelInventario salida, filas
    ocupado = False
    Screen.MousePointer = vbDefault
    MsgBox "Excel generado: " & salida, vbInformation, "Control de Inventario"
    Exit Sub
Fallo:
    mensaje = Err.Description
    ocupado = False
    Screen.MousePointer = vbDefault
    MsgBox "No se pudo generar el Excel: " & mensaje, vbExclamation, "Control de Inventario"
End Sub

Private Sub cmdBorrarRespaldos_Click()
    Dim archivos As New Collection, nombre As String, ruta As Variant
    Dim borrados As Long, fallidos As Long
    On Error GoTo Fallo
    If ocupado Then Exit Sub
    If Len(Dir$(CarpetaRespaldos, vbDirectory)) > 0 Then
        nombre = Dir$(CarpetaRespaldos & "\*.adtg")
        Do While Len(nombre) > 0
            archivos.Add CarpetaRespaldos & "\" & nombre
            nombre = Dir$()
        Loop
    End If
    If archivos.Count = 0 Then
        MsgBox "No hay respaldos de inventario para borrar.", vbInformation
        Exit Sub
    End If
    If MsgBox("Se eliminarán permanentemente los " & archivos.Count & " respaldos anteriores de inventario." & vbCrLf & "Los productos y las cantidades en pantalla no se modificarán. Esta acción no se puede deshacer." & vbCrLf & "¿Borrar todos los respaldos?", vbExclamation Or vbYesNo Or vbDefaultButton2, "Borrar respaldos de inventario") <> vbYes Then Exit Sub
    ocupado = True
    For Each ruta In archivos
        On Error Resume Next
        Err.Clear
        Kill CStr(ruta)
        If Err.Number = 0 Then
            borrados = borrados + 1
            If StrComp(CStr(ruta), respaldoRecuperado, vbTextCompare) = 0 Then respaldoRecuperado = vbNullString
        Else
            fallidos = fallidos + 1
        End If
        Err.Clear
        On Error GoTo Fallo
    Next ruta
    ocupado = False
    CargarRespaldos
    If fallidos > 0 Then
        MsgBox borrados & " respaldos eliminados. No se pudieron eliminar " & fallidos & "; revise si están en uso o protegidos.", vbExclamation
    Else
        MsgBox borrados & " respaldos eliminados.", vbInformation
    End If
    Exit Sub
Fallo:
    ocupado = False
    MsgBox "No se pudo completar el borrado de respaldos: " & Err.Description, vbExclamation
End Sub
