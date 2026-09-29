VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   Caption         =   "Menú"
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   9915
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   8033.195
   ScaleMode       =   0  'User
   ScaleWidth      =   10030.35
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   4410
      Left            =   2160
      Picture         =   "Principal.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5175
   End
   Begin VB.Menu trans 
      Caption         =   "Transacciones"
      Begin VB.Menu fact 
         Caption         =   "Facturacion"
      End
      Begin VB.Menu ingprod 
         Caption         =   "Ingresar Productos"
      End
      Begin VB.Menu rafact 
         Caption         =   "Revisar/Anular Facturas"
      End
   End
   Begin VB.Menu inv 
      Caption         =   "Inventarios"
      Begin VB.Menu prod 
         Caption         =   "Productos"
      End
      Begin VB.Menu controlinv
         Caption         =   "Control de Inventario"
      End
      Begin VB.Menu repinv 
         Caption         =   "Reporte Inventarios"
      End
   End
   Begin VB.Menu user 
      Caption         =   "Usuarios"
   End
   Begin VB.Menu inf 
      Caption         =   "Informes"
      Begin VB.Menu resuvent 
         Caption         =   "Resumen de Venta Diario"
      End
      Begin VB.Menu rvd 
         Caption         =   "Reporte Venta Detallado"
      End
      Begin VB.Menu rvm 
         Caption         =   "Reporte Venta Mensual"
      End
      Begin VB.Menu reporprod 
         Caption         =   "Reporte de Producto"
      End
      Begin VB.Menu reporingre 
         Caption         =   "Reporte de Ingresos"
      End
      Begin VB.Menu rfa 
         Caption         =   "Reporte Facturas Anuladas"
      End
   End
   Begin VB.Menu sa 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fact_Click()
Facturacion.Show vbModeless
Facturacion.ZOrder 0
Facturacion.Text1.SetFocus
End Sub

Private Sub Form_Load()
InicializarRutaBD
End Sub

Private Sub Form_Resize()
Image1.Top = 0
Image1.Left = 0
Image1.Height = Menu.Height + 450
Image1.Width = Menu.Width + 150
End Sub

Private Sub ingprod_Click()
claveing.Show
'incioingresos
End Sub

Private Sub repinv_Click()
claveig.Show
'inventariosg.Show
End Sub

Private Sub prod_Click()
clavep.AbrirControlInventario = False
clavep.Show
End Sub

Private Sub rafact_Click()
clavenulas.Show
End Sub

Private Sub reporingre_Click()
reportingresos.Show
End Sub

Private Sub reporprod_Click()
reporteproducto.Show
End Sub

Private Sub resuvent_Click()
claver.Show
End Sub

Private Sub rfa_Click()
claverepnulas.Show
End Sub

Private Sub rvd_Click()
claverd.Show
End Sub

Private Sub rvm_Click()
claverm.Show
End Sub

Private Sub sa_Click()
Unload Me
End Sub

Private Sub user_Click()
claveu.Show
End Sub

Private Sub ventdet_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ventana As Form
    Dim abierta As Form
    Dim ventanas As New Collection
    'Recorrer una lista estable: descargar formularios modifica Forms.
    For Each ventana In Forms
        If Not ventana Is Me Then ventanas.Add ventana
    Next ventana
    For Each ventana In ventanas
        Unload ventana
        For Each abierta In Forms
            If abierta Is ventana Then
                Cancel = 1
                Exit Sub
            End If
        Next abierta
    Next ventana
End Sub

Private Sub controlinv_Click()
clavep.AbrirControlInventario = True
clavep.Show
End Sub
