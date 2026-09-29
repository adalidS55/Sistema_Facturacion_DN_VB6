VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Ingresos 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresos de Productos"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar Base"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   10515
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   10485
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Top             =   10560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   10335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton Command5 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   25
         Top             =   6840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ingresar Base 2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         TabIndex        =   17
         Top             =   960
         Width           =   855
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
         ItemData        =   "ingresos.frx":0000
         Left            =   1920
         List            =   "ingresos.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox Text1 
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1620
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   8805
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   15531
         _Version        =   393216
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   7440
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   405
         Left            =   1080
         TabIndex        =   19
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   135135233
         CurrentDate     =   43596
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   405
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
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
         Format          =   135135234
         CurrentDate     =   43596
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6480
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "AGREGAR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8700
         TabIndex        =   16
         Top             =   2280
         Width           =   855
      End
      Begin VB.Image Image4 
         Height          =   705
         Left            =   8775
         Picture         =   "ingresos.frx":0004
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "BORRAR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8700
         TabIndex        =   15
         Top             =   3600
         Width           =   855
      End
      Begin VB.Image Image3 
         Height          =   705
         Left            =   8775
         Picture         =   "ingresos.frx":2812
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "LIMPIAR"
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
         Left            =   8640
         TabIndex        =   9
         Top             =   6240
         Width           =   960
      End
      Begin VB.Image Image5 
         Height          =   705
         Left            =   8775
         Picture         =   "ingresos.frx":5CA7
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   705
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   8760
         Picture         =   "ingresos.frx":91AE
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   705
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESAR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8520
         TabIndex        =   8
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Código Producto"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ingresos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7440
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "N. Productos:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   10575
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   0
      Picture         =   "ingresos.frx":C81B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "Ingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs As New ADODB.Recordset  'Creamos el objeto Recordset.
Dim cn2 As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs2 As New ADODB.Recordset
Dim cn3 As New ADODB.Connection
Dim rs3 As New ADODB.Recordset

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Len(Trim(Combo1.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Ingresos"
        Exit Sub
    Else
        rs.Find "nombreprod = '" & Combo1.Text & "'", , , 1
    End If
    If rs.BOF = False And rs.EOF = False Then
        Text1.Text = rs.Fields("codproducto")
        Text5.Text = CStr(ExistenciaVenta(cn, CStr(rs!codproducto)))
        Text2.SetFocus
    ElseIf rs.BOF = True Or rs.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Ingresos"
        Combo1.Text = Empty
        Combo1.SetFocus
    End If
End If
    If KeyCode = vbKeyLeft Then
    Text1.SetFocus
    End If
End Sub

Private Sub Command1_Click()
Dim i As Integer

    If Text2.Text = "" Then
        MsgBox "Valor No valido, Vacio", vbCritical, "Ingresos"
        Text2.SetFocus
        Exit Sub
    ElseIf Not (IsNumeric(Text2.Text)) Then
        MsgBox "Valores No validos. Solo Numeros", vbCritical, "Ingresos"
        Text2.SetFocus
        Exit Sub
    End If
    If Text2.Text = 0 Then
        MsgBox "Valores No validos", vbCritical, "Ingresos"
        Text2.SetFocus
        Exit Sub
    End If

If (Text1.Text = "") Or (Combo1.Text = "") Or (Text2.Text = "") Then
    MsgBox "Datos en BLANCO, llenar las casillas", vbExclamation, "Ingresos"
    Exit Sub
Else
    'For i = 0 To MSFlexGrid1.Rows - 1
        MSFlexGrid1.Col = 1
        MSFlexGrid1.Row = fila
        MSFlexGrid1.Text = Text1.Text
        MSFlexGrid1.Col = 2
        MSFlexGrid1.Row = fila
        MSFlexGrid1.Text = Combo1.Text
        MSFlexGrid1.Col = 3
        MSFlexGrid1.Row = fila
        MSFlexGrid1.Text = Text2.Text
        MSFlexGrid1.Col = 4
        MSFlexGrid1.Row = fila
        MSFlexGrid1.Text = DTPicker2.Value
        fila = fila + 1
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        Text3 = fila
    'Next
End If
borrartexto
Text4.Text = MSFlexGrid1.Rows - 2


''Color del MSFlexGrid------------------------------------------------------------------------------------------------------------------------------------------------------
' Y el renglón que estás comparando.
' X la columna.
For y = 1 To Text4.Text
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


Private Sub Command2_Click()
If MSFlexGrid1.TextMatrix(1, 2) = "" Then
    MsgBox "No hay producto para actualizar", vbCritical, "Ingresos"
    Exit Sub
Else

    Dim b As Integer
    Dim m As Integer
    
    'ingreso en la base
    m = MSFlexGrid1.Rows - 1
    For b = 1 To m - 1
    If (MSFlexGrid1.TextMatrix(b, 1) <> "") And (MSFlexGrid1.TextMatrix(b, 2) <> "") And (MSFlexGrid1.TextMatrix(b, 3) <> "") Then
                rs3.AddNew
                rs3("codigo") = MSFlexGrid1.TextMatrix(b, 1)
                rs3("fecha") = DTPicker1.Value
                rs3("hora") = MSFlexGrid1.TextMatrix(b, 4)
                rs3("descripcion") = MSFlexGrid1.TextMatrix(b, 2)
                rs3("cantidad") = MSFlexGrid1.TextMatrix(b, 3)
                rs3.Update
            Else
                Exit Sub
            End If
        Next
    
End If
MsgBox "Ingreso de producto Exitoso", vbExclamation, "Ingresos"
Image5_Click
fila = 1

End Sub

Private Sub Command4_Click()
borrartexto
'********************************************************************
Dim a As Integer
Dim numreg As Integer
Dim numreg2 As Integer
numreg = 0
numreg2 = 0
rs3.Find "fecha = '" & DTPicker1.Value & "'", , , 1

If rs3.EOF = True Or rs3.BOF = True Then
    MsgBox "No hay registros en Ingresos", vbInformation, "Reporte Ingresos"
    Image5_Click
    Exit Sub
    End If
    
rs3.MoveFirst
Do While rs3.EOF = False
    If rs3.Fields("fecha") = DTPicker1.Value Then
        numreg = numreg + 1
    Else
        numreg = numreg
    End If
    rs3.MoveNext
Loop
Text4.Text = numreg
MSFlexGrid1.Rows = Text4.Text + 1
rs3.MoveFirst
Do While rs3.EOF = False
    numreg2 = numreg2 + 1
    rs3.MoveNext
Loop

If rs3.EOF = True And rs3.BOF = True Then
    MsgBox "No hay registros en Ingresos", vbInformation, "Reporte Ingresos"
    Exit Sub
Else
'********************************************************************
    rs3.Find "fecha = '" & DTPicker1.Value & "'", , , 1
'*********************************************************************
    rs3.MoveFirst
    For a = 1 To Text4.Text
If rs3.Fields("fecha") = DTPicker1.Value Then

            MSFlexGrid1.TextMatrix(a, 1) = rs3.Fields("codigo")
            MSFlexGrid1.TextMatrix(a, 2) = rs3.Fields("descripcion")
            MSFlexGrid1.TextMatrix(a, 3) = rs3.Fields("cantidad")
            MSFlexGrid1.TextMatrix(a, 4) = rs3.Fields("hora")
            
        Else
            a = a - 1
        End If
        rs3.MoveNext
    Next
        
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1

End If

''Color del MSFlexGrid------------------------------------------------------------------------------------------------------------------------------------------------------
' Y el renglón que estás comparando.
' X la columna.
For y = 1 To Text4.Text
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


Private Sub Command3_Click()
If MSFlexGrid1.Row <= 0 Then
    MsgBox "Debe Seleccionar una fila", vbInformation, "Ingresos"
Else
    If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) = Empty Then
        MsgBox "No hay registro que eliminar", vbExclamation, "Ingresos"
    ElseIf MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> Empty Then
        fila = fila - 1
        Text4.Text = Text4.Text - 1
        MSFlexGrid1.RemoveItem (MSFlexGrid1.Row)
    End If
End If
End Sub



''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  If TypeOf Me.ActiveControl Is MSFlexGrid Then FlexGridScroll Me.ActiveControl, MouseKeys, Rotation, Xpos, Ypos
End Sub



Private Sub Command5_Click()

Dim v As Integer
Dim w As Integer
Dim a As Integer
'********************************************************************
Dim numreg As Integer
Dim numreg2 As Integer
numreg = 0
numreg2 = 0
rs2.Find "fecha = '" & DTPicker1.Value & "'", , , 1

If rs2.EOF = True Or rs2.BOF = True Then
    MsgBox "No hay registros en Ingresos", vbInformation, "Reporte Ingresos"
Exit Sub
End If
    
    rs2.MoveFirst
    Do While rs2.EOF = False
        If rs2.Fields("fecha") = DTPicker1.Value Then
            numreg = numreg + 1
        Else
            numreg = numreg
        End If
        rs2.MoveNext
    Loop
    Text4.Text = numreg
    MSFlexGrid1.Rows = Text4.Text + 1
    rs2.MoveFirst
    Do While rs2.EOF = False
        numreg2 = numreg2 + 1
        rs2.MoveNext
    Loop

If rs2.EOF = True And rs2.BOF = True Then
    MsgBox "No hay registros en Facturas", vbInformation, "Reporte Ventas"
    Exit Sub
Else
'********************************************************************
    rs2.Find "fecha = '" & DTPicker1.Value & "'", , , 1
'*********************************************************************
rs2.MoveFirst
    For a = 1 To Text4.Text
If rs2.Fields("fecha") = DTPicker1.Value Then

            MSFlexGrid1.TextMatrix(a, 1) = rs2.Fields("codigo")
            MSFlexGrid1.TextMatrix(a, 2) = rs2.Fields("descripcion")
            MSFlexGrid1.TextMatrix(a, 3) = rs2.Fields("cantidad")
            MSFlexGrid1.TextMatrix(a, 4) = rs2.Fields("hora")
            
        Else
            a = a - 1
        End If
        rs2.MoveNext
    Next
        
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
End If

End Sub

Private Sub Form_Load()

DTPicker1.Value = Date

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
  Dim N As Integer
  Dim i As Integer
  With MSFlexGrid1
    .Rows = 2
    .Cols = 5
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
rs.Open "select * from productos", cn

Set rs2 = New ADODB.Recordset
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs2.Source = "ingresos"
rs2.CursorType = adOpenKeyset
rs2.LockType = adLockOptimistic
rs2.Open "select * from ingresos", cn2


''Segunda Base Para Ingresos------------------------------------------------------------------------------------------------------------------------------------------------------
Set rs3 = New ADODB.Recordset
cn3.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD2 & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs3.Source = "ingresos"
rs3.CursorType = adOpenKeyset
rs3.LockType = adLockOptimistic
rs3.Open "select * from ingresos", cn3


While rs.EOF = False
Combo1.AddItem rs!nombreprod
rs.MoveNext
Wend

fila = 1

MSFlexGrid1.Col = 1
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "CODIGO"
MSFlexGrid1.ColAlignment(1) = 3

MSFlexGrid1.Col = 2
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "DESCRIPCION"
MSFlexGrid1.ColAlignment(2) = 3

MSFlexGrid1.Col = 3
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "CANT."
MSFlexGrid1.ColAlignment(3) = 3

MSFlexGrid1.ColWidth(4) = 0

Image1.Left = 0
Image1.Top = 0
Image1.Height = Ingresos.Height
Image1.Width = Ingresos.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
cn2.Close
cn3.Close


''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Call WheelUnHook(Me)

End Sub

Private Sub Image2_Click()
Dim filaVinculo As Long
On Error GoTo ErrorVinculo
For filaVinculo = 1 To MSFlexGrid1.Rows - 1
    If Len(MSFlexGrid1.TextMatrix(filaVinculo, 1)) > 0 Then
        If EsPresentacion(cn, MSFlexGrid1.TextMatrix(filaVinculo, 1)) Then
            MsgBox "Registre la entrada en el producto base, usando sus unidades. Los paquetes no tienen inventario propio.", vbExclamation
            Exit Sub
        End If
    End If
Next filaVinculo
If MSFlexGrid1.TextMatrix(1, 2) = "" Then
    MsgBox "No hay producto para actualizar", vbCritical, "Ingresos"
    Exit Sub
Else
    Dim a As Integer
    Dim N As Integer
    Dim mcinv As Double
    
    Dim b As Integer
    Dim m As Integer
    
    'n = MSFlexGrid1.Rows - 1
    For a = 1 To Text4.Text
        rs.Find "codproducto = '" & MSFlexGrid1.TextMatrix(a, 1) & "'", , , 1
        If rs.BOF = False And rs.EOF = False Then
            mcinv = rs.Fields("cantprod") + MSFlexGrid1.TextMatrix(a, 3)
            rs.Update Array("cantprod"), Array(mcinv)
            rs.Update
        ElseIf rs.BOF = True Or rs.EOF = True Then
            MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Inventarios"
        End If
    Next
    
    'ingreso en la base
    m = MSFlexGrid1.Rows - 1
    For b = 1 To m - 1
    If (MSFlexGrid1.TextMatrix(b, 1) <> "") And (MSFlexGrid1.TextMatrix(b, 2) <> "") And (MSFlexGrid1.TextMatrix(b, 3) <> "") Then
                rs2.AddNew
                rs2("codigo") = MSFlexGrid1.TextMatrix(b, 1)
                rs2("fecha") = DTPicker1.Value
                rs2("hora") = MSFlexGrid1.TextMatrix(b, 4)
                rs2("descripcion") = MSFlexGrid1.TextMatrix(b, 2)
                rs2("cantidad") = MSFlexGrid1.TextMatrix(b, 3)
                rs2.Update
            Else
                Exit Sub
            End If
        Next
    
    
    
End If
MsgBox "Ingreso de producto Exitoso", vbExclamation, "Ingresos"
Image5_Click
fila = 1
Exit Sub
ErrorVinculo:
    MsgBox Err.Description, vbExclamation
End Sub

Private Sub Image3_Click()
Command3_Click
End Sub

Private Sub Image4_Click()
Command1_Click
End Sub

Private Sub Image5_Click()
Text1.Text = Empty
Text2.Text = Empty
Combo1.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Dim v As Integer
Dim w As Integer
For v = 0 To MSFlexGrid1.Rows - 1
  For w = 0 To MSFlexGrid1.Cols - 1
  MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
MSFlexGrid1.Rows = 2
MSFlexGrid1.ColWidth(0) = 1

MSFlexGrid1.Col = 1
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "CODIGO"
MSFlexGrid1.ColAlignment(1) = 3

MSFlexGrid1.Col = 2
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "DESCRIPCION"
MSFlexGrid1.ColAlignment(2) = 3

MSFlexGrid1.Col = 3
MSFlexGrid1.Row = 0
MSFlexGrid1.Text = "CANT."
MSFlexGrid1.ColAlignment(3) = 3

fila = 1
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Len(Trim(Text1.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Ingresos"
        Text1.SetFocus
        Exit Sub
    Else
        rs.Find "codproducto = '" & Text1.Text & "'", , , 1
    End If
    If rs.BOF = False And rs.EOF = False Then
        Combo1.Text = rs.Fields("nombreprod")
        Text5.Text = CStr(ExistenciaVenta(cn, CStr(rs!codproducto)))
        Text2.SetFocus
    ElseIf rs.BOF = True Or rs.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Ingresos"
        Text1.Text = Empty
        Text1.SetFocus
    End If
End If

    If KeyCode = vbKeyRight Then
    Combo1.SetFocus
    End If
    
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Text2.Text = "" Then
        MsgBox "Valor No valido, Vacio", vbCritical, "Ingresos"
        Text2.SetFocus
        Exit Sub
    ElseIf Not (IsNumeric(Text2.Text)) Then
        MsgBox "Valores No validos. Solo Numeros", vbCritical, "Ingresos"
        Text2.SetFocus
        Exit Sub
    End If
    If Text2.Text = 0 Then
        MsgBox "Valores No validos", vbCritical, "Ingresos"
        Text2.SetFocus
        Exit Sub
    Else
        Command1_Click
        Text1.SetFocus
    End If
End If
End Sub


Private Sub Timer1_Timer()
'Text3.Text = MSFlexGrid1.Row
DTPicker2.Value = Time
End Sub
