VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form productos 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6645
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInventarioBase
      Caption = "Usar inventario de otro producto"
      BackColor = &H00404000&
      ForeColor = &H00FFFFFF&
      Left = 240
      Top = 5040
      Width = 6000
      Height = 360
      TabIndex = 30
   End
   Begin VB.ComboBox cboProductoBase
      Style = 0
      Left = 240
      Top = 5460
      Width = 6000
      Height = 315
      TabIndex = 31
      Enabled = 0
   End
   Begin VB.Label lblFactorBase
      Caption = "Unidades del producto base por cada unidad vendida:"
      BackStyle = 0
      ForeColor = &H00FFFFFF&
      Left = 240
      Top = 5940
      Width = 4560
      Height = 300
   End
   Begin VB.TextBox txtFactorBase
      Text = "1"
      Left = 4920
      Top = 5880
      Width = 1320
      Height = 360
      TabIndex = 32
      Enabled = 0
   End
   Begin VB.Label lblNotaBase
      Caption = "Ejemplo: un paquete de 5 libras consume 5 del codigo por libra."
      BackStyle = 0
      ForeColor = &H00FFFFFF&
      Left = 240
      Top = 6300
      Width = 6000
      Height = 300
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7560
      Top             =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Productos.frx":0000
      Left            =   2040
      List            =   "Productos.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Propiedades"
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
      Height          =   2415
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   6255
      Begin VB.TextBox Text7 
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
         Left            =   4560
         TabIndex        =   26
         Top             =   1800
         Width           =   1455
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
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
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
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
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
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
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
         Left            =   960
         TabIndex        =   11
         Top             =   840
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   138215425
         CurrentDate     =   44558
      End
      Begin VB.Label Label16 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingresos"
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
         Left            =   2880
         TabIndex        =   27
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
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
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Publico"
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
         Left            =   2880
         TabIndex        =   18
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Precio Costo"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Vencimiento"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   10
         Top             =   900
         Width           =   735
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   7080
      TabIndex        =   28
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   138215425
      CurrentDate     =   44558
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   7080
      TabIndex        =   29
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   138215426
      CurrentDate     =   44558
   End
   Begin VB.Image Image8 
      Height          =   705
      Left            =   1350
      Picture         =   "Productos.frx":0004
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label Label15 
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
      Left            =   1440
      TabIndex        =   25
      Top             =   4560
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   705
      Left            =   270
      Picture         =   "Productos.frx":3C01
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label Label14 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nuevo"
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
      Left            =   360
      TabIndex        =   24
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label13 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   1080
      TabIndex        =   23
      Top             =   705
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   1080
      TabIndex        =   22
      Top             =   165
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar"
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
      Left            =   2385
      TabIndex        =   6
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Limpiar"
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
      Left            =   5745
      TabIndex        =   5
      Top             =   4560
      Width           =   735
   End
   Begin VB.Image Image5 
      Height          =   705
      Left            =   5775
      Picture         =   "Productos.frx":640F
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Eiminar"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   4560
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   705
      Left            =   4560
      Picture         =   "Productos.frx":9916
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3405
      TabIndex        =   3
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   705
      Left            =   3480
      Picture         =   "Productos.frx":CDAB
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   705
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   705
      Left            =   2400
      Picture         =   "Productos.frx":20F74
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   705
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   225
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   240
      Picture         =   "Productos.frx":245E1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Image Image6 
      Height          =   2775
      Left            =   0
      Picture         =   "Productos.frx":305D6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "productos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private codigosBase As Collection
Private codigoCargado As String
Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs As New ADODB.Recordset  'Creamos el objeto Recordset.
Dim cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Image1_Click
End If
End Sub

Private Sub Command1_Click()
LimpiarVinculo
CargarProductosBase
Text1.Text = Empty
Text1.Enabled = True
Text2.Text = 0
Text2.Enabled = False
Text3.Text = Empty
Text3.Enabled = True
Text4.Text = Empty
Text4.Enabled = True
Text5.Text = Empty
Text5.Enabled = True
Text6.Text = Empty
Text7.Text = Empty
Text7.Enabled = True
Combo1.Text = Empty
Image2.Enabled = True
DTPicker1.Enabled = True
Image3.Enabled = False
Image4.Enabled = False

DTPicker1.Value = Date
Text5.SetFocus

End Sub

Private Sub Command2_Click()
chkInventarioBase.Enabled = True
Text1.Enabled = True
Text2.Enabled = False
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = False
Text6.Text = Empty
Text7.Enabled = True
Combo1.Text = Empty
Image2.Enabled = True
DTPicker1.Enabled = True
Image3.Enabled = True
Image4.Enabled = True
chkInventarioBase_Click
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
    PrepararVinculaciones cn
rs.Source = "productos"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from productos", cn

Set rs2 = New ADODB.Recordset
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs2.Source = "ingresos"
rs2.CursorType = adOpenKeyset
rs2.LockType = adLockOptimistic
rs2.Open "select * from ingresos", cn2

While rs.EOF = False
Combo1.AddItem rs!nombreprod
rs.MoveNext
Wend

CargarProductosBase
LimpiarVinculo
Image6.ZOrder 1
Image6.Height = productos.Height
Image6.Width = productos.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
cn2.Close
End Sub

Private Sub Image1_Click()
If Text6.Text <> "" Then
    If Len(Trim(Text6.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Productos"
        Text6.SetFocus
        Exit Sub
    Else
        rs.Find "codproducto = '" & Replace$(Text6.Text, "'", "''") & "'", , , 1
    End If
    If rs.BOF = False And rs.EOF = False Then
        Text1.Text = rs.Fields("nombreprod")
        If IsNull(rs!cantprod) Then
            Text2.Text = "0"
        Else
            Text2.Text = CStr(rs!cantprod)
        End If
        Text3.Text = rs.Fields("precioc")
        Text4.Text = rs.Fields("preciov")
        DTPicker1.Value = rs.Fields("fechav")
        Text5.Text = rs("codproducto")
        codigoCargado = Text5.Text
        CargarVinculoProducto
        
        Image3.Enabled = True
        Image4.Enabled = True
        
    ElseIf rs.BOF = True Or rs.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Productos"
        Image5_Click
    End If
Else
If Combo1.Text <> "" Then
    If Len(Trim(Combo1.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Productos"
        Combo1.SetFocus
        Exit Sub
    Else
        rs.Find "nombreprod = '" & Replace$(Combo1.Text, "'", "''") & "'", , , 1
    End If
    If rs.BOF = False And rs.EOF = False Then
        Text1.Text = rs.Fields("nombreprod")
        If IsNull(rs!cantprod) Then
            Text2.Text = "0"
        Else
            Text2.Text = CStr(rs!cantprod)
        End If
        Text3.Text = rs.Fields("precioc")
        Text4.Text = rs.Fields("preciov")
        DTPicker1.Value = rs.Fields("fechav")
        Text5.Text = rs("codproducto")
        codigoCargado = Text5.Text
        CargarVinculoProducto
                
        Image3.Enabled = True
        Image4.Enabled = True
        
    ElseIf rs.BOF = True Or rs.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Productos"
        Image5_Click
    End If
End If
End If

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text7.Enabled = False
Text7.Text = 0
DTPicker1.Enabled = False

chkInventarioBase.Enabled = False
cboProductoBase.Enabled = False
txtFactorBase.Enabled = False
End Sub

Private Sub Image2_Click()
    GuardarProductoPantalla True
End Sub

Private Sub Image3_Click()
    GuardarProductoPantalla False
End Sub

Private Sub Image4_Click()
On Error GoTo ErrorBorrar
ValidarBorradoProducto cn, Text5.Text
If (Text1.Text = "") Or (Text2.Text = "") Or (Text3.Text = "") Or (Text4.Text = "") Or (Text5.Text = "") Then
    MsgBox "Primero Debe Buscar el Registro Para poder Eliminarlo", vbExclamation, "Productos"
Else
'**************************************************************************************
    If MsgBox("¿Esta seguro de Eliminar este Registro?, No se podran recuperra los datos.", vbYesNo, "Productos") = vbYes Then
        rs.Delete
        MsgBox "Registro Eliminado", vbInformation, "Productos"
        rs.MoveNext  'Nos movemos al siguiente registro para no provocar un error.
        If rs.EOF Then rs.MoveLast  'Si es el fin del archivo nos movemos al último registro.
        
        'ingreso en la base
        rs2.AddNew
        rs2("codigo") = Text5.Text
        rs2("fecha") = DTPicker2.Value
        rs2("hora") = DTPicker3.Value
        rs2("descripcion") = Text1.Text + " (ELIMINADO)"
        rs2("cantidad") = Text2.Text
        rs2.Update
        
        Image5_Click
    Else
        MsgBox "Los datos No se Eliminaron", vbCritical, "Productos"
    End If
'*************************************************************************************
End If
Exit Sub
ErrorBorrar:
    MsgBox Err.Description, vbExclamation, "Productos"
End Sub

Private Sub Image5_Click()
LimpiarVinculo
Text1.Text = Empty
Text1.Enabled = False
Text2.Text = Empty
Text2.Enabled = False
Text3.Text = Empty
Text3.Enabled = False
Text4.Text = Empty
Text4.Enabled = False
Text5.Text = Empty
Text5.Enabled = False
Text6.Text = Empty
Text7.Text = Empty
Text7.Enabled = False
Combo1.Text = Empty
DTPicker1.Enabled = False
End Sub

Private Sub Image7_Click()
Command1_Click
End Sub

Private Sub Image8_Click()
Command2_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
End Sub

Private Sub Text5_Change()
Text5.BackColor = vbWhite
End Sub

Private Sub Text6_Change()
Text6.BackColor = vbWhite

End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Image1_Click
End If
End Sub

Private Sub Timer1_Timer()
DTPicker2.Value = Date
DTPicker3.Value = Time
End Sub

Private Sub CargarProductosBase()
    Dim datos As ADODB.Recordset
    Set codigosBase = New Collection
    cboProductoBase.Clear
    Set datos = cn.Execute("SELECT codproducto,nombreprod FROM productos WHERE codproducto NOT IN (SELECT codventa FROM presentaciones) ORDER BY nombreprod")
    Do While Not datos.EOF
        cboProductoBase.AddItem CStr(datos!nombreprod)
        codigosBase.Add CStr(datos!codproducto)
        datos.MoveNext
    Loop
    datos.Close
End Sub

Private Sub LimpiarVinculo()
    codigoCargado = vbNullString
    chkInventarioBase.Value = vbUnchecked
    chkInventarioBase.Enabled = True
    cboProductoBase.ListIndex = -1
    txtFactorBase.Text = "1"
    cboProductoBase.Enabled = False
    txtFactorBase.Enabled = False
End Sub

Private Sub chkInventarioBase_Click()
    cboProductoBase.Enabled = (chkInventarioBase.Value = vbChecked)
    txtFactorBase.Enabled = cboProductoBase.Enabled
    If chkInventarioBase.Value = vbChecked Then
        Text7.Text = "0"
        Text7.Enabled = False
    Else
        Text7.Enabled = Text1.Enabled
    End If
End Sub

Private Sub CargarVinculoProducto()
    Dim base As String, factor As Currency, i As Long
    CargarProductosBase
    ResolverInventario cn, Text5.Text, base, factor
    chkInventarioBase.Value = vbUnchecked
    txtFactorBase.Text = CStr(factor)
    If StrComp(base, Text5.Text, vbTextCompare) <> 0 Then
        chkInventarioBase.Value = vbChecked
        For i = 1 To codigosBase.Count
            If StrComp(CStr(codigosBase(i)), base, vbTextCompare) = 0 Then cboProductoBase.ListIndex = i - 1
        Next i
    End If
End Sub

Private Sub GuardarProductoPantalla(ByVal nuevo As Boolean)
    Dim base As String, factor As Currency, valor As Variant
    On Error GoTo Fallo
    If Len(Trim$(Text5.Text)) = 0 Or Len(Trim$(Text1.Text)) = 0 Then Err.Raise 5, , "Escriba codigo y nombre."
    If Len(Text5.Text) > 255 Or Len(Text1.Text) > 255 Then Err.Raise 5, , "Codigo y nombre permiten hasta 255 caracteres."
    If Not IsNumeric(Text7.Text) Or Not IsNumeric(Text3.Text) Or Not IsNumeric(Text4.Text) Then Err.Raise 5, , "Escriba cantidades y precios numericos."
    If Not nuevo Then
        If StrComp(codigoCargado, Text5.Text, vbTextCompare) <> 0 Or Len(codigoCargado) = 0 Then Err.Raise 5, , "Busque el producto antes de actualizar; su codigo no se puede cambiar."
        If MsgBox("¿Actualizar este producto y su vinculacion?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Productos") <> vbYes Then Exit Sub
    End If
    factor = 1
    If chkInventarioBase.Value = vbChecked Then
        SeleccionarBasePorNombre False
        base = CStr(codigosBase(cboProductoBase.ListIndex + 1))
        If Not IsNumeric(txtFactorBase.Text) Then Err.Raise 5, , "Escriba un factor numerico mayor que cero."
        valor = CDec(txtFactorBase.Text)
        factor = CCur(valor)
        If CDec(factor) <> valor Or factor <= 0 Then Err.Raise 5, , "El factor debe ser positivo y tener hasta cuatro decimales."
    End If
    GuardarProductoVinculado cn, nuevo, Trim$(Text5.Text), Trim$(Text1.Text), CCur(Text7.Text), CCur(Text4.Text), CCur(Text3.Text), DTPicker1.Value, base, factor
    rs.Requery
    Combo1.Clear
    Do While Not rs.EOF
        Combo1.AddItem CStr(rs!nombreprod)
        rs.MoveNext
    Loop
    Image5_Click
    CargarProductosBase
    MsgBox "Producto y vinculacion guardados correctamente.", vbInformation, "Productos"
    Exit Sub
Fallo:
    MsgBox "No se guardo el producto: " & Err.Description, vbExclamation, "Productos"
End Sub

Private Sub SeleccionarBasePorNombre(ByVal permitirInicio As Boolean)
    Dim i As Long, coincidencias As Long, encontrado As Long, consulta As String
    If cboProductoBase.ListIndex >= 0 Then Exit Sub
    consulta = Trim$(cboProductoBase.Text)
    If Len(consulta) = 0 Then Err.Raise 5, , "Escriba o seleccione el nombre del producto base."
    encontrado = -1
    For i = 0 To cboProductoBase.ListCount - 1
        If StrComp(cboProductoBase.List(i), consulta, vbTextCompare) = 0 Then
            coincidencias = coincidencias + 1
            encontrado = i
        End If
    Next i
    If coincidencias = 0 And permitirInicio Then
        For i = 0 To cboProductoBase.ListCount - 1
            If StrComp(Left$(cboProductoBase.List(i), Len(consulta)), consulta, vbTextCompare) = 0 Then
                coincidencias = coincidencias + 1
                encontrado = i
            End If
        Next i
    End If
    If coincidencias = 0 Then Err.Raise 5, , "No se encontro ese nombre. Seleccione un producto de la lista."
    If coincidencias > 1 Then Err.Raise 5, , "Hay varios productos que coinciden. Complete el nombre o seleccione uno de la lista."
    cboProductoBase.ListIndex = encontrado
End Sub

Private Sub cboProductoBase_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Fallo
    If KeyCode <> vbKeyReturn Then Exit Sub
    KeyCode = 0
    SeleccionarBasePorNombre True
    txtFactorBase.SetFocus
    Exit Sub
Fallo:
    MsgBox Err.Description, vbExclamation, "Buscar producto base"
End Sub
