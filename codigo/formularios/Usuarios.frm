VERSION 5.00
Begin VB.Form usuarios 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4875
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "DATOS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   4575
      Begin VB.TextBox Text1 
         DataField       =   "Id"
         DataSource      =   "Adodc1"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   285
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         DataField       =   "nombre"
         DataSource      =   "Adodc1"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   885
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         DataField       =   "contra"
         DataSource      =   "Adodc1"
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
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1485
         Width           =   2775
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1080
         Picture         =   "Usuarios.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "VER"
         Top             =   1485
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image Image9 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1080
         Picture         =   "Usuarios.frx":8E9F
         Stretch         =   -1  'True
         ToolTipText     =   "VER"
         Top             =   1485
         Width           =   405
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Empleado"
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
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
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
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
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
      Left            =   5880
      TabIndex        =   8
      Top             =   240
      Width           =   1335
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
      Left            =   5880
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "EDITAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   1155
      Picture         =   "Usuarios.frx":12B2C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "NUEVO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   195
      Picture         =   "Usuarios.frx":16729
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo Empleado:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   2160
      Picture         =   "Usuarios.frx":18F37
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   255
      Picture         =   "Usuarios.frx":24F2C
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   705
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   1455
      Picture         =   "Usuarios.frx":28599
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   705
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   2775
      Picture         =   "Usuarios.frx":3C762
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   705
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   705
      Left            =   3855
      Picture         =   "Usuarios.frx":3FBF7
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   705
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "BUSCAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2085
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "ELIMINAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "ACTUALIZAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "GUARDAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   975
   End
   Begin VB.Image Image6 
      Height          =   2775
      Left            =   0
      Picture         =   "Usuarios.frx":430FE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs As New ADODB.Recordset  'Creamos el objeto Recordset.

Private Sub Command1_Click()
Text1.Text = Empty
Text1.Enabled = True
Text2.Text = Empty
Text2.Enabled = True
Text3.Text = Empty
Text3.Enabled = True
Image2.Enabled = True
Image3.Enabled = False
Image4.Enabled = False

Text1.SetFocus
End Sub

Private Sub Command2_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Image2.Enabled = True
Image3.Enabled = True
Image4.Enabled = True
End Sub


Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs.Source = "usuarios"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from usuarios", cn

Image6.Height = usuarios.Height
Image6.Width = usuarios.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
End Sub

Private Sub Image1_Click()
If Len(Trim(Text4.Text)) = 0 Then
        MsgBox "Debe Especificar el texto a buscar y digitar correctamente", vbExclamation, "Usuarios"
        Text4.SetFocus
        Exit Sub
    Else
        rs.Find "codempleado = '" & Text4.Text & "'", , , 1
    End If
    If rs.BOF = False And rs.EOF = False Then
        Text1.Text = rs.Fields("codempleado")
        Text2.Text = rs.Fields("nombreempleado")
        Text3.Text = rs.Fields("contra")
    ElseIf rs.BOF = True Or rs.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Usuarios"
        Image5_Click
    End If
End Sub

Private Sub Image10_Click()
Text3.PasswordChar = "*"
Image10.Visible = False
Image9.Visible = True
End Sub

Private Sub Image2_Click()
If (Text1.Text = "") Or (Text2.Text = "") Or (Text3.Text = "") Then
    MsgBox "Todos la informacion debe ser llenada", vbExclamation, "Usuarios"
    Exit Sub
Else
    'Buscamos para no repetir datos
    rs.Find "codempleado = '" & Text1.Text & "'", , , 1
        If rs.BOF = False And rs.EOF = False Then
            MsgBox "Este registro ya existe el CODIGO debe ser un valor unico", vbCritical, "Usuarios"
            Text1.SetFocus
        Exit Sub
        ElseIf rs.BOF = True Or rs.EOF = True Then
            'guardar
            rs.AddNew
            rs("codempleado") = Text1.Text
            rs("nombreempleado") = Text2.Text
            rs("contra") = Text3.Text
            rs.Update
            MsgBox "Datos Guardados Exitosamente", vbInformation, "Usuarios"
        End If
End If
Image5_Click
End Sub

Private Sub Image3_Click()
If (Text1.Text = "") Or (Text2.Text = "") Or (Text3.Text = "") Then
    MsgBox "Datos en BLANCO. Para poder ACTUALIZAR debe realizar una buqueda primero", vbExclamation, "Usuarios"
    Exit Sub
Else
'buscamos si el codigo no existe (en caso que accidentalmente se oprima actualizar en lugar de guardar)
    rs.Find "codempleado = '" & Text4.Text & "'", , , 1
    If rs.BOF = False And rs.EOF = False Then
        'ACTUALIZAR
        If MsgBox("¿Desea ACTUALIZAR este Registro?", vbYesNo, "Usuarios") = vbYes Then
        rs.Update Array("codempleado", "nombreempleado", "contra"), Array(Text1.Text, Text2.Text, Text3.Text)
        rs.Update  'Actualizamos el registro.
        'Verificamos si no ocurrió ningún problema.
            If rs.State = 1 Or rs.State = 0 Then
                MsgBox "El registro se ha actualizado con éxito.", vbInformation, "Usuarios"
                Image5_Click
            Else
                MsgBox "Ha ocurrido un ERROR al actualizar el registro.", vbCritical, "Usuarios"
            End If
        Else
            MsgBox "NO se Completo la Actualizacion del Regisro", vbCritical, "Usuarios"
        End If
        '********************
    ElseIf rs.BOF = True Or rs.EOF = True Then
        MsgBox "El registro NO existe, digitar correctamente el codigo a buscar", vbExclamation, "Usuarios"
        Image5_Click
    End If
End If
End Sub

Private Sub Image4_Click()
If (Text1.Text = "") Or (Text2.Text = "") Or (Text3.Text = "") Then
    MsgBox "Primero Debe Buscar el Registro Para poder Eliminarlo", vbExclamation, "Usuarios"
Else
'**************************************************************************************
    If MsgBox("¿Esta seguro de Eliminar este Registro?, No se podran recuperra los datos.", vbYesNo, "Usuarios") = vbYes Then
        rs.Delete
        MsgBox "Registro Eliminado", vbInformation, "Usuarios"
        Image5_Click
        rs.MoveNext  'Nos movemos al siguiente registro para no provocar un error.
        If rs.EOF Then rs.MoveLast  'Si es el fin del archivo nos movemos al último registro.
    Else
        MsgBox "Los datos No se Eliminaron", vbCritical, "Usuarios"
    End If
'*************************************************************************************
End If
End Sub

Private Sub Image5_Click()
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
End Sub

Private Sub Image7_Click()
Command1_Click
End Sub

Private Sub Image8_Click()
Command2_Click
End Sub

Private Sub Image9_Click()
Text3.PasswordChar = Empty
Image10.Visible = True
Image9.Visible = False
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Image1_Click
End If
End Sub
