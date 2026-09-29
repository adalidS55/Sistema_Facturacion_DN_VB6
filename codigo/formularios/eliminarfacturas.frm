VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rafactura 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisar/Anular Factura"
   ClientHeight    =   6285
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   15915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   1
      EndProperty
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
      Height          =   400
      Left            =   5040
      TabIndex        =   17
      Top             =   5698
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.TextBox Text6 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   1
      EndProperty
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
      Height          =   400
      Left            =   1680
      TabIndex        =   10
      Top             =   5698
      Width           =   1400
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """L""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   2
      EndProperty
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
      Height          =   400
      Left            =   14040
      TabIndex        =   9
      Top             =   5685
      Width           =   1400
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FF80&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """L""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   2
      EndProperty
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
      Height          =   400
      Left            =   11400
      TabIndex        =   8
      Top             =   5700
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """L""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   18442
         SubFormatType   =   2
      EndProperty
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
      Height          =   400
      Left            =   9120
      TabIndex        =   7
      Top             =   5700
      Width           =   800
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
      Height          =   400
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   15625
      _ExtentX        =   27570
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   9
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   405
      Left            =   11040
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   137363458
      CurrentDate     =   43596
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   8520
      TabIndex        =   2
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   137363457
      CurrentDate     =   43596
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3652
      TabIndex        =   19
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "T. Ganancia:"
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
      Left            =   3600
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   3720
      Picture         =   "eliminarfacturas.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "BUSCAR"
      Top             =   120
      Width           =   705
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   14595
      Picture         =   "eliminarfacturas.frx":BFF5
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   14520
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   13080
      Picture         =   "eliminarfacturas.frx":F4FC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ANULAR"
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
      Left            =   13065
      TabIndex        =   15
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "  =  Total  L."
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
      Left            =   12720
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cant. de Prod."
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
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "-- Desc."
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
      Left            =   10560
      TabIndex        =   12
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Total"
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
      Left            =   8040
      TabIndex        =   11
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Fact:"
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
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   435
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   7800
      TabIndex        =   4
      Top             =   435
      Width           =   615
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hora"
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
      Left            =   10440
      TabIndex        =   3
      Top             =   435
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   4215
      Left            =   0
      Picture         =   "eliminarfacturas.frx":12991
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "rafactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs As New ADODB.Recordset  'Creamos el objeto Recordset.

Dim cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset

Dim cn3 As New ADODB.Connection
Dim rs3 As New ADODB.Recordset

Dim cn4 As New ADODB.Connection
Dim rs4 As New ADODB.Recordset

Private Sub Command2_Click()
Image3_Click
Me.Hide
Menu.Show
End Sub

Private Sub Form_Activate()
Text1.SetFocus
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
    .Cols = 9
    For N = .FixedRows To .Rows - 1
      .TextMatrix(N, 0) = "Row " & N
    Next N
  End With
  Call WheelHook(Me)
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------

Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
    PrepararVinculaciones cn
rs.Source = "facturacion"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from facturacion", cn

Set rs2 = New ADODB.Recordset
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs2.Source = "facturacionnulas"
rs2.CursorType = adOpenKeyset
rs2.LockType = adLockOptimistic
rs2.Open "select * from facturacionnulas", cn2

Set rs3 = New ADODB.Recordset
cn3.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs3.Source = "facturareporte"
rs3.CursorType = adOpenKeyset
rs3.LockType = adLockOptimistic
rs3.Open "select * from facturareporte", cn3

Set rs4 = New ADODB.Recordset
cn4.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs4.Source = "productos"
rs4.CursorType = adOpenKeyset
rs4.LockType = adLockOptimistic
rs4.Open "select * from productos", cn4

cuadrofactnula

Image4.Left = 0
Image4.Top = 0
Image4.Height = rafactura.Height
Image4.Width = rafactura.Width

DTPicker1.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
cn2.Close
cn3.Close
cn4.Close

''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Call WheelUnHook(Me)

End Sub

Private Sub Image1_Click()
If (Len(Trim(Text1.Text)) = 0) Or (Not (IsNumeric(Text1.Text))) Then
    MsgBox "Debe especificar el Numero de Factura a Buscar, solo valores numericos", vbExclamation, "Anular/Ravisar Factura"
    Text1.SetFocus
    Exit Sub
Else

    ' busqueda en reportefact y facturacion
    rs3.Find "numfact = '" & Text1.Text & "'", , , 1
    rs.Find "numfact = '" & Text1.Text & "'", , , 1
    If (rs3.EOF = True Or rs3.BOF = True) And (rs.EOF = True Or rs.BOF = True) Then
        MsgBox "No hay registros con este numero de Factura", vbInformation, "Anular/Ravisar Factura"
        Exit Sub
    ElseIf (rs3.EOF = False And rs3.BOF = False) And (rs.EOF = False And rs.BOF = False) Then
        Text2.Text = rs3.Fields("subtotal")
        Text3.Text = rs3.Fields("totaldescuento")
        Text4.Text = rs3.Fields("totalneto")
        Text5.Text = rs3.Fields("totalganacia")
        DTPicker1.Value = rs3.Fields("fechafact")
        DTPicker2.Value = rs3.Fields("horafact")
        Text6.Text = rs.Fields("nprodfact")
        MSFlexGrid1.Rows = Text6.Text + 1
        'rs.MoveFirst
        For a = 1 To Text6.Text
            If rs.Fields("numfact") = Text1.Text Then
                MSFlexGrid1.TextMatrix(a, 1) = rs.Fields("cantidad")
                MSFlexGrid1.TextMatrix(a, 2) = rs.Fields("codprodfact")
                MSFlexGrid1.TextMatrix(a, 3) = rs.Fields("nombreprodfact")
                MSFlexGrid1.TextMatrix(a, 4) = rs.Fields("precio")
                MSFlexGrid1.TextMatrix(a, 5) = rs.Fields("descuentoi")
                'MSFlexGrid1.TextMatrix(a, 6) = rs.Fields("totalneto")
                MSFlexGrid1.TextMatrix(a, 6) = (Val(MSFlexGrid1.TextMatrix(a, 4)) - Val(MSFlexGrid1.TextMatrix(a, 5))) * Val(MSFlexGrid1.TextMatrix(a, 1))
                MSFlexGrid1.TextMatrix(a, 7) = rs.Fields("preciocosto")
                MSFlexGrid1.TextMatrix(a, 8) = rs.Fields("gananciai")
            End If
            rs.MoveNext
        Next
    End If
End If
End Sub

Private Sub Image2_Click()
    On Error GoTo Fallo
    If Not IsNumeric(Text1.Text) Then
        MsgBox "Busque una factura valida antes de anular.", vbExclamation
        Exit Sub
    End If
    If MsgBox("¿Esta seguro de ANULAR esta factura y devolver su inventario?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Anular factura") <> vbYes Then Exit Sub
    AnularVentaVinculada cn, CLng(Text1.Text)
    rs.Requery
    rs2.Requery
    rs3.Requery
    rs4.Requery
    Image3_Click
    MsgBox "Factura anulada. Se devolvieron las cantidades realmente descontadas.", vbInformation
    Exit Sub
Fallo:
    MsgBox "No se pudo anular la factura: " & Err.Description, vbExclamation
End Sub

Private Sub Image3_Click()
limpiarcuadrofactnula
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
cuadrofactnula
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Image1_Click
End If
End Sub

