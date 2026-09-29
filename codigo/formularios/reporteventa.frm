VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form reporteventa 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTE DE VENTAS"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
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
      Height          =   405
      Left            =   1440
      TabIndex        =   12
      Top             =   6885
      Width           =   800
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   735
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
      Left            =   8160
      TabIndex        =   5
      Top             =   6885
      Width           =   1500
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
      Left            =   4440
      TabIndex        =   3
      Top             =   6885
      Width           =   1500
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   8
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
      TabIndex        =   7
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
      Format          =   131727361
      CurrentDate     =   43605
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cant. Trans."
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
      TabIndex        =   13
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   8895
      Picture         =   "reporteventa.frx":0000
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
      Left            =   8760
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "contador todas las facturas"
      Height          =   495
      Left            =   2760
      TabIndex        =   10
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   2055
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
      Left            =   6600
      TabIndex        =   6
      Top             =   6960
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
      Left            =   3000
      TabIndex        =   4
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   6375
      Picture         =   "reporteventa.frx":3717
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
      Left            =   6300
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   7635
      Picture         =   "reporteventa.frx":9512
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
      Left            =   7560
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   0
      Picture         =   "reporteventa.frx":CA19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "reporteventa"
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
    .Cols = 8
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  Call WheelHook(Me)
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------



Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs.Source = "facturareporte"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from facturareporte", cn

cuadroreporte

Image3.Height = reporteventa.Height
Image3.Width = reporteventa.Width

DTPicker1.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Call WheelUnHook(Me)
End Sub

Private Sub Image1_Click()
' envirommen2 y datareport2
'DataEnvironment2.Command1 Trim(DTPicker1.Value)
'DataReport2.Show 1
End Sub

Private Sub Image2_Click()
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Dim v As Integer
Dim w As Integer
Dim a As Integer
'********************************************************************
Dim numreg As Integer
Dim numreg2 As Integer
numreg = 0
numreg2 = 0
rs.Find "fechafact = '" & DTPicker1.Value & "'", , , 1

If rs.EOF = True Or rs.BOF = True Then
    MsgBox "No hay registros en Facturas", vbInformation, "Reporte Ventas"
    limpiarcuadroreporte
    cuadroreporte
    Text4.Text = Empty
    Exit Sub
    End If
    
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
If rs.EOF = True And rs.BOF = True Then
    MsgBox "No hay registros en Facturas", vbInformation, "Reporte Ventas"
    Exit Sub
Else
'********************************************************************
    rs.Find "fechafact = '" & DTPicker1.Value & "'", , , 1
'*********************************************************************
    rs.MoveFirst
    For a = 1 To Text1.Text
If rs.Fields("fechafact") = DTPicker1.Value Then
            MSFlexGrid1.TextMatrix(a, 1) = rs.Fields("numfact")
            MSFlexGrid1.TextMatrix(a, 2) = rs.Fields("fechafact")
            MSFlexGrid1.TextMatrix(a, 3) = rs.Fields("horafact")
            MSFlexGrid1.TextMatrix(a, 4) = rs.Fields("subtotal")
            MSFlexGrid1.TextMatrix(a, 5) = rs.Fields("totaldescuento")
            MSFlexGrid1.TextMatrix(a, 6) = rs.Fields("totalneto")
            MSFlexGrid1.TextMatrix(a, 7) = rs.Fields("totalganacia")
            Text3.Text = Val(Text3.Text) + Val(MSFlexGrid1.TextMatrix(a, 7))
            Text2.Text = Val(Text2.Text) + Val(MSFlexGrid1.TextMatrix(a, 6))
        Else
            a = a - 1
        End If
        rs.MoveNext
    Next
End If
End Sub



Private Sub Image6_Click()
limpiarcuadroreporte
cuadroreporte
Text4.Text = Empty
End Sub
