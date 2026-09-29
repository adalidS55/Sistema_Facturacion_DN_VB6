VERSION 5.00
Begin VB.Form claver 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   562
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "x"
      TabIndex        =   0
      Top             =   1042
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave:"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   2085
      Left            =   0
      Picture         =   "claver.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4305
   End
End
Attribute VB_Name = "claver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs As New ADODB.Recordset  'Creamos el objeto Recordset.
Dim cn2 As New ADODB.Connection 'Creamos el objeto Connection.
Dim rs2 As New ADODB.Recordset  'Creamos el objeto Recordset.


Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs.Source = "usuarios"
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
rs.Open "select * from usuarios", cn

Set rs2 = New ADODB.Recordset
cn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & rutaBD & ";Persist Security Info=False;JET OLEDB:DATABASE PASSWORD = Dnacho2024$@"
rs2.Source = "usuarios"
rs2.CursorType = adOpenKeyset
rs2.LockType = adLockOptimistic
rs2.Open "select * from usuarios", cn2

claver.Height = Image1.Height + 400
claver.Width = Image1.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Close
cn2.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text2.SetFocus
End If

If KeyCode = vbKeyDown Then
Text2.SetFocus
End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If (Len(Trim(Text1.Text)) = 0) Or (Len(Trim(Text2.Text)) = 0) Then
        MsgBox "Digite correctamente su Usuario y Clave", vbExclamation, "Clave"
        Text1.Text = Empty
        Text2.Text = Empty
        Text1.SetFocus
        Exit Sub
    Else
        rs.Find "codempleado = '" & Text1.Text & "'", , , 1
        rs2.Find "contra = '" & Text2.Text & "'", , , 1
        If (rs.BOF = False And rs.EOF = False) And (rs2.BOF = False And rs2.EOF = False) Then
            Text1.Text = Empty
            Text2.Text = Empty
            reporteventa.Show
            Unload Me
        ElseIf (rs.BOF = True Or rs.EOF = True) Or (rs2.BOF = True Or rs2.EOF = True) Then
            MsgBox "Usuario o Calve Incorrecto", vbExclamation, "Clave"
            Text1.Text = Empty
            Text2.Text = Empty
            Text1.SetFocus
            Exit Sub
'*********************************************
        End If
    End If
End If

If KeyCode = vbKeyUp Then
Text1.SetFocus
End If

End Sub

