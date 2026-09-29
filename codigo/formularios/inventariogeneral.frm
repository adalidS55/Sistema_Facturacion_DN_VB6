VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form inventariosg 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario General"
   ClientHeight    =   8610
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   14115
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   8100
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   8100
      Width           =   1500
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   20
      Top             =   8100
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BUSCAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   14880
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BLOQUEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   14880
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   8040
      TabIndex        =   16
      Top             =   8760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton botoneditar 
      Caption         =   "EDITAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   14880
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11040
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9480
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6945
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12250
      _Version        =   393216
      Cols            =   6
      Redraw          =   -1  'True
      ScrollTrack     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   7425
      Left            =   12360
      TabIndex        =   1
      Top             =   1080
      Width           =   1560
      Begin VB.Label Label16 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Negativos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   420
         TabIndex        =   29
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image9 
         Height          =   705
         Left            =   495
         Picture         =   "inventariogeneral.frx":0000
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   705
      End
      Begin VB.Label Label13 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   540
         TabIndex        =   26
         Top             =   960
         Width           =   615
      End
      Begin VB.Image Image5 
         Height          =   705
         Left            =   495
         Picture         =   "inventariogeneral.frx":8061
         Stretch         =   -1  'True
         Top             =   240
         Width           =   705
      End
      Begin VB.Image Image3 
         Height          =   705
         Left            =   495
         Picture         =   "inventariogeneral.frx":14056
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   705
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   420
         TabIndex        =   17
         Top             =   5760
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   495
         Picture         =   "inventariogeneral.frx":2821F
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   2160
         Width           =   735
      End
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   495
         Picture         =   "inventariogeneral.frx":2E01A
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   705
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   540
         TabIndex        =   3
         Top             =   6960
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   4560
         Width           =   735
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   495
         Picture         =   "inventariogeneral.frx":31521
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   705
      End
   End
   Begin VB.Label Label15 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bloquear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13125
      TabIndex        =   28
      Top             =   720
      Width           =   780
   End
   Begin VB.Image Image8 
      Height          =   510
      Left            =   13260
      Picture         =   "inventariogeneral.frx":34C38
      Stretch         =   -1  'True
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label14 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12405
      TabIndex        =   27
      Top             =   720
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   505
      Left            =   12400
      Picture         =   "inventariogeneral.frx":3D2C2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   505
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de Productos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   8160
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Precio Venta"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Precio Costo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   23
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Precio C."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11040
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Precio V."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cant"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   4215
      Left            =   0
      Picture         =   "inventariogeneral.frx":458E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "inventariosg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub botoneditar_Click()
Text2.Enabled = True
Text3.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text2.SetFocus
End Sub

Private Sub Command1_Click()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
End Sub


Private Sub Command3_Click()
If Len(Trim(Text2.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Inventario General"
        'Text2.SetFocus
        Exit Sub
    Else
        rs.Find "codproducto = '" & Text2.Text & "'", , , 1
    End If
    If rs.BOF = False And rs.EOF = False Then
        Text3.Text = rs.Fields("nombreprod")
        Text4.Text = rs.Fields("cantprod")
        Text5.Text = rs.Fields("preciov")
        Text6.Text = rs.Fields("precioc")
    ElseIf rs.BOF = True Or rs.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Inventario General"
        Image6_Click
    End If
End Sub


''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  If TypeOf Me.ActiveControl Is MSFlexGrid Then FlexGridScroll Me.ActiveControl, MouseKeys, Rotation, Xpos, Ypos
End Sub


Private Sub Form_Load()

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
  Dim N As Integer
  Dim I As Integer
  With MSFlexGrid1
    .Rows = 2
    .Cols = 6
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  Call WheelHook(Me)
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------


Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
    PrepararVinculaciones cn
rs.Source = "productos"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from productos WHERE codproducto NOT IN (SELECT codventa FROM presentaciones)", cn

cuadroinventario

Image4.Left = 0
Image4.Top = 0
Image4.Height = inventariosg.Height
Image4.Width = inventariosg.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Call WheelUnHook(Me)
End Sub

Private Sub Image1_Click()
' envirommen3 y datareport3
'DataReport3.Show 1
End Sub

Private Sub Image2_Click()
Dim v As Integer
Dim w As Integer
Dim a As Integer
'********************************************************************
Dim numreg2 As Integer
numreg2 = 0
rs.MoveFirst
Do While rs.EOF = False
    numreg2 = numreg2 + 1
    rs.MoveNext
Loop
Text1.Text = numreg2
MSFlexGrid1.Rows = Text1.Text + 1
'***************************************************************************************
rs.MoveFirst
For a = 1 To Text1.Text
        MSFlexGrid1.TextMatrix(a, 1) = rs.Fields("codproducto")
        MSFlexGrid1.TextMatrix(a, 2) = rs.Fields("nombreprod")
        MSFlexGrid1.TextMatrix(a, 3) = rs.Fields("cantprod")
        MSFlexGrid1.TextMatrix(a, 4) = rs.Fields("preciov")
        MSFlexGrid1.TextMatrix(a, 5) = rs.Fields("precioc")
        
        Text8.Text = Val(Text8.Text) + (Val(MSFlexGrid1.TextMatrix(a, 4)) * (MSFlexGrid1.TextMatrix(a, 3)))
        Text9.Text = Val(Text9.Text) + (Val(MSFlexGrid1.TextMatrix(a, 5)) * (MSFlexGrid1.TextMatrix(a, 3)))
    rs.MoveNext
Next
'*********************************************************************

MSFlexGrid1.Col = 2
MSFlexGrid1.Sort = flexSortGenericAscending

''Color del MSFlexGrid------------------------------------------------------------------------------------------------------------------------------------------------------
' Y el renglón que estás comparando.
' X la columna.
For y = 1 To Text1.Text
If Val(MSFlexGrid1.TextMatrix(y, 3)) < 0 Then
MSFlexGrid1.Row = y
MSFlexGrid1.Col = 3
MSFlexGrid1.CellForeColor = &HFF& ' Color rojo para el texto.
Else
MSFlexGrid1.Row = y
MSFlexGrid1.Col = 3
MSFlexGrid1.CellForeColor = 0 ' Color negro
End If
Next


End Sub

Private Sub Image3_Click()
On Error GoTo ErrorVinculo
If EsPresentacion(cn, Text2.Text) Then
    MsgBox "Modifique la existencia desde el producto base.", vbExclamation
    Exit Sub
End If
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or (Not (IsNumeric(Text4.Text))) Or Text5.Text = "" Or (Not (IsNumeric(Text5.Text))) Or Text6.Text = "" Or (Not (IsNumeric(Text6.Text))) Then
    MsgBox "No hay productos para Actualizar o Ingreso mal algunos Datos", vbCritical, "Inventario General"
    Exit Sub
Else
    'n = MSFlexGrid1.Rows - 1
    rs.Find "codproducto = '" & Text2.Text & "'", , , 1
    rs.Update Array("nombreprod", "cantprod", "preciov", "precioc"), Array(Text3.Text, Text4.Text, Text5.Text, Text6.Text)
    rs.Update
End If
MsgBox "Se ah Actualizado el registro", vbExclamation, "Inventario General"
Image2_Click
'MSFlexGrid1.TextMatrix(Text7.Text, 1)
Exit Sub
ErrorVinculo:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Image5_Click()
Command3_Click
End Sub

Private Sub Image6_Click()
limpiarcuadroinventario
cuadroinventario
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
Text7.Text = Empty
Text8.Text = Empty
Text9.Text = Empty
End Sub

Private Sub Image7_Click()
botoneditar_Click
End Sub

Private Sub Image8_Click()
Command1_Click
End Sub

Private Sub Image9_Click()
Dim v As Integer
Dim w As Integer
Dim a As Integer
Dim numreg2 As Integer
Dim fila As Integer
fila = 1 ' Controla la fila en la que se van a mostrar los datos en el MSFlexGrid
numreg2 = 0

rs.MoveFirst
Do While rs.EOF = False
    If rs.Fields("cantprod") < 0 Then ' Solo contar registros donde cantprod < 0
        numreg2 = numreg2 + 1
    End If
    rs.MoveNext
Loop
Text1.Text = numreg2
MSFlexGrid1.Rows = Text1.Text + 1

' Mostrar únicamente los registros con cantprod < 0
rs.MoveFirst
Do While Not rs.EOF
    If rs.Fields("cantprod") < 0 Then ' Condición de filtro
        MSFlexGrid1.TextMatrix(fila, 1) = rs.Fields("codproducto")
        MSFlexGrid1.TextMatrix(fila, 2) = rs.Fields("nombreprod")
        MSFlexGrid1.TextMatrix(fila, 3) = rs.Fields("cantprod")
        MSFlexGrid1.TextMatrix(fila, 4) = rs.Fields("preciov")
        MSFlexGrid1.TextMatrix(fila, 5) = rs.Fields("precioc")

        ' Calcular los totales con productos cuyo cantprod < 0
        Text8.Text = Val(Text8.Text) + (Val(MSFlexGrid1.TextMatrix(fila, 4)) * Val(MSFlexGrid1.TextMatrix(fila, 3)))
        Text9.Text = Val(Text9.Text) + (Val(MSFlexGrid1.TextMatrix(fila, 5)) * Val(MSFlexGrid1.TextMatrix(fila, 3)))
        
        fila = fila + 1 ' Incrementar la fila en la que se asignan los datos
    End If
    rs.MoveNext
Loop

' Ordenar y colorear el MSFlexGrid
MSFlexGrid1.Col = 2
MSFlexGrid1.Sort = flexSortGenericAscending

For y = 1 To Text1.Text
    If Val(MSFlexGrid1.TextMatrix(y, 3)) < 0 Then
        MSFlexGrid1.Row = y
        MSFlexGrid1.Col = 3
        MSFlexGrid1.CellForeColor = &HFF& ' Color rojo para el texto
    Else
        MSFlexGrid1.Row = y
        MSFlexGrid1.Col = 3
        MSFlexGrid1.CellForeColor = 0 ' Color negro
    End If
Next

End Sub

Private Sub MSFlexGrid1_Click()
Text7.Text = MSFlexGrid1.Row
Text2 = MSFlexGrid1.TextMatrix(Text7.Text, 1)
Text3 = MSFlexGrid1.TextMatrix(Text7.Text, 2)
Text4 = MSFlexGrid1.TextMatrix(Text7.Text, 3)
Text5 = MSFlexGrid1.TextMatrix(Text7.Text, 4)
Text6 = MSFlexGrid1.TextMatrix(Text7.Text, 5)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Command3_Click
End If
End Sub
