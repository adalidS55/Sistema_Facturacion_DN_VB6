VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form reportingresos 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTE DE INGRESOS"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11760
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      TabIndex        =   4
      Top             =   7080
      Width           =   675
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10080
      TabIndex        =   3
      Top             =   7125
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
      ItemData        =   "reportingresos.frx":0000
      Left            =   240
      List            =   "reportingresos.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   5895
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   11000
      _ExtentX        =   19394
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   6
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114950145
      CurrentDate     =   45115
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114950145
      CurrentDate     =   45115
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114950145
      CurrentDate     =   45115
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Height          =   255
      Left            =   7680
      TabIndex        =   18
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTE DEL DIA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9840
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   705
      Left            =   9915
      Picture         =   "reportingresos.frx":0004
      Stretch         =   -1  'True
      Top             =   240
      Width           =   705
   End
   Begin VB.Label Label7 
      Caption         =   "contador todas las facturas"
      Height          =   495
      Left            =   11640
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   4275
      Picture         =   "reportingresos.frx":5DFF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "REPORTE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   5400
      Picture         =   "reportingresos.frx":BBFA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label5 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "BORRAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5325
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Registros"
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
      TabIndex        =   12
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de Productos"
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
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   7200
      Width           =   2295
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   11760
      Picture         =   "reportingresos.frx":F101
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label8 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11625
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final:"
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
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicio:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   0
      Picture         =   "reportingresos.frx":12818
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "reportingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset



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
rs.Source = "ingresos"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from ingresos", cn

Set rs2 = New ADODB.Recordset
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs2.Source = "productos"
rs2.CursorType = adOpenKeyset
rs2.LockType = adLockOptimistic
rs2.Open "select * from productos", cn2

While rs2.EOF = False
Combo1.AddItem rs2!nombreprod
rs2.MoveNext
Wend

cuadroingresos

Image3.Height = reportingresos.Height
Image3.Width = reportingresos.Width

DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
cn2.Close

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Call WheelUnHook(Me)

End Sub

Private Sub Image1_Click()
' envirommen1 y datareport1
'DataEnvironment1.Command1 Trim(DTPicker1.Value)
'DataReport1.Show 1
End Sub

Private Sub Image2_Click()
Text2.Text = 0
Dim v As Integer
Dim w As Integer
Dim a As Integer
Dim numreg As Integer
Dim numreg2 As Integer
numreg = 0
numreg2 = 0

If DTPicker1.Value > DTPicker2.Value Then
MsgBox "El rango de fechas no es correcto", vbInformation, "Reporte Producto"
Exit Sub
End If

rs.Find "fecha = '" & DTPicker1.Value & "'", , , 1
rs.Find "descripcion = '" & Combo1.Text & "'", , , 1

If rs.EOF = True Or rs.BOF = True Then
    MsgBox "No hay registros en Facturas", vbInformation, "Reporte Producto"
    limpiarcuadroingresos
    cuadroingresos
    Exit Sub
ElseIf rs.EOF = False And rs.BOF = False Then
    rs.MoveFirst
    Do While rs.EOF = False
        If rs.Fields("fecha") >= DTPicker1.Value And rs.Fields("fecha") <= DTPicker2.Value And rs.Fields("descripcion") = Combo1.Text Then
            numreg = numreg + 1
        Else
            numreg = numreg
        End If
        rs.MoveNext
    Loop
    Text1.Text = numreg
    MSFlexGrid1.Rows = Text1.Text + 1
    rs.MoveFirst
    Do While rs.EOF = False
        numreg2 = numreg2 + 1
        rs.MoveNext
    Loop
    Text4.Text = numreg2
    '********************************************************************
    rs.Find "fecha = '" & DTPicker1.Value & "'", , , 1
    rs.Find "descripcion = '" & Combo1.Text & "'", , , 1
    '*********************************************************************
    rs.MoveFirst
    For a = 1 To Text1.Text
        If rs.Fields("fecha") >= DTPicker1.Value And rs.Fields("fecha") <= DTPicker2.Value And rs.Fields("descripcion") = Combo1.Text Then
            MSFlexGrid1.TextMatrix(a, 1) = rs.Fields("codigo")
            MSFlexGrid1.TextMatrix(a, 2) = rs.Fields("fecha")
            MSFlexGrid1.TextMatrix(a, 3) = rs.Fields("hora")
            MSFlexGrid1.TextMatrix(a, 4) = rs.Fields("descripcion")
            MSFlexGrid1.TextMatrix(a, 5) = rs.Fields("cantidad")

            Text2.Text = Val(Text2.Text) + Val((MSFlexGrid1.TextMatrix(a, 5)))
        Else
            a = a - 1
            
        End If
        rs.MoveNext

        
    Next
    '*********************************************************************
    'Totales parte de abajo de MSFlexgrid1
    'Total precio Publico
    '*********************************************************************
End If

'***********************************************************************************

''Color del MSFlexGrid------------------------------------------------------------------------------------------------------------------------------------------------------
' Y el renglón que estás comparando.
' X la columna.
For y = 1 To Text1.Text
If Val(MSFlexGrid1.TextMatrix(y, 5)) < 0 Then
MSFlexGrid1.Row = y
MSFlexGrid1.Col = 5
MSFlexGrid1.CellForeColor = &HFF& ' Color rojo para el texto.
Else
MSFlexGrid1.Row = y
MSFlexGrid1.Col = 5
MSFlexGrid1.CellForeColor = 0 ' Color negro
End If
Next

End Sub

Private Sub Image4_Click()
Text1.Text = Empty
Text2.Text = Empty
Dim v As Integer
Dim w As Integer
Dim a As Integer
'********************************************************************
Dim numreg As Integer
Dim numreg2 As Integer
numreg = 0
numreg2 = 0
rs.Find "fecha = '" & DTPicker3.Value & "'", , , 1

If rs.EOF = True Or rs.BOF = True Then
    MsgBox "No hay registros en Ingresos", vbInformation, "Reporte Ingresos"
Exit Sub
End If
    
    rs.MoveFirst
    Do While rs.EOF = False
        If rs.Fields("fecha") = DTPicker3.Value Then
            numreg = numreg + 1
        Else
            numreg = numreg
        End If
        rs.MoveNext
    Loop
    Text1.Text = numreg
    MSFlexGrid1.Rows = Text1.Text + 1
    rs.MoveFirst
    Do While rs.EOF = False
        numreg2 = numreg2 + 1
        rs.MoveNext
    Loop
Text4.Text = numreg2
If rs.EOF = True And rs.BOF = True Then
    MsgBox "No hay registros en Facturas", vbInformation, "Reporte Ventas"
    Exit Sub
Else
'********************************************************************
    rs.Find "fecha = '" & DTPicker3.Value & "'", , , 1
'*********************************************************************
    rs.MoveFirst
    For a = 1 To Text1.Text
        If rs.Fields("fecha") = DTPicker3.Value Then
            MSFlexGrid1.TextMatrix(a, 1) = rs.Fields("codigo")
            MSFlexGrid1.TextMatrix(a, 2) = rs.Fields("fecha")
            MSFlexGrid1.TextMatrix(a, 3) = rs.Fields("hora")
            MSFlexGrid1.TextMatrix(a, 4) = rs.Fields("descripcion")
            MSFlexGrid1.TextMatrix(a, 5) = rs.Fields("cantidad")

            Text2.Text = Val(Text2.Text) + Val((MSFlexGrid1.TextMatrix(a, 5)))
        Else
            a = a - 1
        End If
        rs.MoveNext
    Next
End If

''Color del MSFlexGrid------------------------------------------------------------------------------------------------------------------------------------------------------
' Y el renglón que estás comparando.
' X la columna.
For y = 1 To Text1.Text
If Val(MSFlexGrid1.TextMatrix(y, 5)) < 0 Then
MSFlexGrid1.Row = y
MSFlexGrid1.Col = 5
MSFlexGrid1.CellForeColor = &HFF& ' Color rojo para el texto.
Else
MSFlexGrid1.Row = y
MSFlexGrid1.Col = 5
MSFlexGrid1.CellForeColor = 0 ' Color negro
End If
Next


End Sub

Private Sub Image6_Click()
limpiarcuadroingresos
cuadroingresos
End Sub

