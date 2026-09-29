VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form reporteventadetalladonulas 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTE FACTURAS ANULADAS"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   16005
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   6645
      Width           =   800
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF80&
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
      Left            =   9000
      TabIndex        =   1
      Top             =   6645
      Width           =   1695
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
      Left            =   5160
      TabIndex        =   0
      Top             =   6645
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   13
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   137232385
      CurrentDate     =   43605
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
      Left            =   12285
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   12360
      Picture         =   "reportefactanuladas.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label7 
      Caption         =   "contador todas las facturas"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Especifique Fecha:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   13635
      Picture         =   "reportefactanuladas.frx":5DFB
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
      Left            =   13560
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad de Registros"
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
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Ganancia"
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
      Left            =   7440
      TabIndex        =   8
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Vendido"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   14895
      Picture         =   "reportefactanuladas.frx":9302
      Stretch         =   -1  'True
      Top             =   120
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
      Left            =   14760
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   0
      Picture         =   "reportefactanuladas.frx":CA19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "reporteventadetalladonulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

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
    .Cols = 13
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  Call WheelHook(Me)
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------

Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs.Source = "facturacionnulas"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from facturacionnulas", cn
cuadrodetalladonulas

Image3.Height = reporteventadetalladonulas.Height
Image3.Width = reporteventadetalladonulas.Width

DTPicker1.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Call WheelUnHook(Me)

End Sub

Private Sub Image1_Click()
' envirommen1 y datareport1
'DataEnvironment1.Command1 Trim(DTPicker1.Value)
'DataReport1.Show 1
End Sub

Private Sub Image2_Click()
Text3.Text = 0
Text2.Text = 0
Dim v As Integer
Dim w As Integer
Dim a As Integer
Dim numreg As Integer
Dim numreg2 As Integer
numreg = 0
numreg2 = 0

rs.Find "fechafact = '" & DTPicker1.Value & "'", , , 1

If rs.EOF = True Or rs.BOF = True Then
    MsgBox "No hay registros en Facturas", vbInformation, "Reporte Detallado"
    limpiarcuadrodetalladonulas
    cuadrodetalladonulas
    Exit Sub
ElseIf rs.EOF = False And rs.BOF = False Then
    rs.MoveFirst
    Do While rs.EOF = False
        If rs.Fields("fechafact") = DTPicker1.Value Then
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
    rs.Find "fechafact = '" & DTPicker1.Value & "'", , , 1
    '*********************************************************************
    rs.MoveFirst
    For a = 1 To Text1.Text
        If rs.Fields("fechafact") = DTPicker1.Value Then
            MSFlexGrid1.TextMatrix(a, 1) = rs.Fields("numfact")
            MSFlexGrid1.TextMatrix(a, 2) = rs.Fields("fechafact")
            MSFlexGrid1.TextMatrix(a, 3) = rs.Fields("horafact")
            MSFlexGrid1.TextMatrix(a, 4) = rs.Fields("nombreprodfact")
            MSFlexGrid1.TextMatrix(a, 5) = rs.Fields("cantidad")
            MSFlexGrid1.TextMatrix(a, 6) = rs.Fields("precio")
            MSFlexGrid1.TextMatrix(a, 7) = rs.Fields("preciocosto")
            MSFlexGrid1.TextMatrix(a, 8) = rs.Fields("descuentoi")
            MSFlexGrid1.TextMatrix(a, 9) = rs.Fields("subtotal")
            MSFlexGrid1.TextMatrix(a, 10) = rs.Fields("totaldescuento")
            MSFlexGrid1.TextMatrix(a, 11) = rs.Fields("totalneto")
            MSFlexGrid1.TextMatrix(a, 12) = rs.Fields("totalganacia")
            Text3.Text = Val(Text3.Text) + ((Val(MSFlexGrid1.TextMatrix(a, 6)) - (MSFlexGrid1.TextMatrix(a, 7))) - Val(MSFlexGrid1.TextMatrix(a, 8))) * (MSFlexGrid1.TextMatrix(a, 5))
            Text2.Text = Val(Text2.Text) + Val((MSFlexGrid1.TextMatrix(a, 5)) * (MSFlexGrid1.TextMatrix(a, 6))) - (Val((MSFlexGrid1.TextMatrix(a, 5)) * (MSFlexGrid1.TextMatrix(a, 8))))
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
End Sub

Private Sub Image6_Click()
limpiarcuadrodetalladonulas
cuadrodetalladonulas
End Sub

