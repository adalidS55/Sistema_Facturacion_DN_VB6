VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form reporteventamensual 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORTE DE VENTA MENSUAL"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnterior
      Caption = "< Anterior"
      Left = 120
      Top = 6315
      Width = 1560
      Height = 360
      Enabled = 0
      TabIndex = 20
   End
   Begin VB.CommandButton cmdSiguiente
      Caption = "Siguiente >"
      Left = 1800
      Top = 6315
      Width = 1560
      Height = 360
      Enabled = 0
      TabIndex = 21
   End
   Begin VB.Label lblPagina
      Caption = "Sin consulta"
      BackStyle = 0
      ForeColor = &H00FFFFFF&
      Left = 3600
      Top = 6360
      Width = 6120
      Height = 300
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
      Height          =   405
      Left            =   1440
      TabIndex        =   11
      Top             =   6885
      Width           =   1200
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   1300
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
      Height          =   5055
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
      Format          =   130416641
      CurrentDate     =   43605
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2280
      TabIndex        =   13
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
      Format          =   130416641
      CurrentDate     =   43605
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
      Left            =   2280
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   120
      Width           =   1335
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
      TabIndex        =   12
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   8895
      Picture         =   "reporteventamensual.frx":0000
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
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "contador todas las facturas"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
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
      Picture         =   "reporteventamensual.frx":3717
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
      Picture         =   "reporteventamensual.frx":9512
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
      Picture         =   "reporteventamensual.frx":CA19
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "reporteventamensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Private informe As New ReporteVentas


''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  If TypeOf Me.ActiveControl Is MSFlexGrid Then FlexGridScroll Me.ActiveControl, MouseKeys, Rotation, Xpos, Ypos
End Sub


Private Sub Form_Load()

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
  Dim N As Long
  Dim I As Long
  With MSFlexGrid1
    .Rows = 2
    .Cols = 8
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  Call WheelHook(Me)
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------



cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"

cuadroreportemensual

Image3.Height = reporteventamensual.Height
Image3.Width = reporteventamensual.Width

DTPicker1.Value = Date
DTPicker2.Value = Date

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
    On Error GoTo Fallo
    Screen.MousePointer = vbHourglass
    informe.Cargar cn, DTPicker1.Value, DTPicker2.Value, False
    Text1.Text = CStr(informe.Cantidad)
    Text4.Text = CStr(informe.CantidadHistorica)
    Text2.Text = CStr(informe.TotalVendido)
    Text3.Text = CStr(informe.TotalGanancia)
    Text2.ToolTipText = "Total del periodo completo: " & Text2.Text
    Text3.ToolTipText = "Ganancia del periodo completo: " & Text3.Text
    MostrarPaginaInforme
    Screen.MousePointer = vbDefault
    If informe.Cantidad = 0 Then MsgBox "No hay ventas en las fechas seleccionadas.", vbInformation, "Reporte de ventas"
    Exit Sub
Fallo:
    Screen.MousePointer = vbDefault
    Dim mensaje As String
    mensaje = Err.Description
    Image6_Click
    MsgBox "No se pudo generar el informe: " & mensaje, vbExclamation, "Reporte de ventas"
End Sub



Private Sub Image6_Click()
    informe.Limpiar
    limpiarcuadroreportemensual
    cuadroreportemensual
    lblPagina.Caption = "Sin consulta"
    cmdAnterior.Enabled = False
    cmdSiguiente.Enabled = False
    Text2.ToolTipText = ""
    Text3.ToolTipText = ""
End Sub


Private Sub MostrarPaginaInforme()
    informe.Mostrar MSFlexGrid1
    cmdAnterior.Enabled = (informe.Pagina > 1)
    cmdSiguiente.Enabled = (informe.Pagina < informe.Paginas)
    lblPagina.Caption = "Pagina " & informe.Pagina & " de " & informe.Paginas & " | " & informe.Cantidad & " registros | Totales del periodo completo"
End Sub

Private Sub cmdAnterior_Click()
    On Error GoTo Fallo
    informe.MoverPagina -1
    MostrarPaginaInforme
    Exit Sub
Fallo:
    MsgBox Err.Description, vbExclamation, "Pagina del informe"
End Sub

Private Sub cmdSiguiente_Click()
    On Error GoTo Fallo
    informe.MoverPagina 1
    MostrarPaginaInforme
    Exit Sub
Fallo:
    MsgBox Err.Description, vbExclamation, "Pagina del informe"
End Sub
