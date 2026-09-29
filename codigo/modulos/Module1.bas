Attribute VB_Name = "Module1"
Public fila As Integer
Public fila2 As Integer
Public inv As Double
Dim pf As Integer
Public revisardesc As Double
Public ganancia As Double
Public descuento As Double
Public total As Double


Public rutaBD As String
Public rutaBD2 As String
Public rutaPendientes As String
Public Sub InicializarRutaBD()
    'El ejecutable y el proyecto VB6 se encuentran en la raiz del sistema.
    Dim carpetaDatos As String
    carpetaDatos = App.Path & "\datos"
    rutaBD = carpetaDatos & "\dondenacho1.mdb"
    rutaBD2 = carpetaDatos & "\Ingresos.mdb"
    rutaPendientes = carpetaDatos & "\pendientes"
End Sub



Public Sub borrartexto()
Ingresos.Text1 = Empty
Ingresos.Text2 = Empty
Ingresos.Text5 = Empty
Ingresos.Combo1 = Empty
'Ingresos.Text1.SetFocus
End Sub

Public Sub cuadroinventario()
inventariosg.MSFlexGrid1.ColWidth(0) = 1

inventariosg.MSFlexGrid1.Col = 1
inventariosg.MSFlexGrid1.Row = 0
inventariosg.MSFlexGrid1.Text = "CODIGO"
inventariosg.MSFlexGrid1.ColWidth(1) = 1800
inventariosg.MSFlexGrid1.ColAlignment(1) = 3

inventariosg.MSFlexGrid1.Col = 2
inventariosg.MSFlexGrid1.Row = 0
inventariosg.MSFlexGrid1.Text = "DESCRIPCION"
inventariosg.MSFlexGrid1.ColWidth(2) = 5200
inventariosg.MSFlexGrid1.ColAlignment(2) = 3

inventariosg.MSFlexGrid1.Col = 3
inventariosg.MSFlexGrid1.Row = 0
inventariosg.MSFlexGrid1.Text = "CANT."
inventariosg.MSFlexGrid1.ColWidth(3) = 1500
inventariosg.MSFlexGrid1.ColAlignment(3) = 3

inventariosg.MSFlexGrid1.Col = 4
inventariosg.MSFlexGrid1.Row = 0
inventariosg.MSFlexGrid1.Text = "PRECIO V."
inventariosg.MSFlexGrid1.ColWidth(4) = 1500
inventariosg.MSFlexGrid1.ColAlignment(4) = 3

inventariosg.MSFlexGrid1.Col = 5
inventariosg.MSFlexGrid1.Row = 0
inventariosg.MSFlexGrid1.Text = "PRECIO C."
inventariosg.MSFlexGrid1.ColWidth(5) = 1500
inventariosg.MSFlexGrid1.ColAlignment(5) = 3
End Sub

Public Sub limpiarcuadroinventario()
Dim v As Integer
Dim w As Integer
For v = 0 To inventariosg.MSFlexGrid1.Rows - 1
  For w = 0 To inventariosg.MSFlexGrid1.Cols - 1
  inventariosg.MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
inventariosg.MSFlexGrid1.Rows = 2
inventariosg.MSFlexGrid1.Cols = 6
End Sub

Public Sub cuadroreporte()
reporteventa.MSFlexGrid1.ColWidth(0) = 1

reporteventa.MSFlexGrid1.Col = 1
reporteventa.MSFlexGrid1.Row = 0
reporteventa.MSFlexGrid1.Text = "FACT."
reporteventa.MSFlexGrid1.ColWidth(1) = 1200
reporteventa.MSFlexGrid1.ColAlignment(1) = 3

reporteventa.MSFlexGrid1.Col = 2
reporteventa.MSFlexGrid1.Row = 0
reporteventa.MSFlexGrid1.Text = "FECHA"
reporteventa.MSFlexGrid1.ColWidth(2) = 1300
reporteventa.MSFlexGrid1.ColAlignment(2) = 3

reporteventa.MSFlexGrid1.Col = 3
reporteventa.MSFlexGrid1.Row = 0
reporteventa.MSFlexGrid1.Text = "HORA"
reporteventa.MSFlexGrid1.ColWidth(3) = 1300
reporteventa.MSFlexGrid1.ColAlignment(3) = 3

reporteventa.MSFlexGrid1.Col = 4
reporteventa.MSFlexGrid1.Row = 0
reporteventa.MSFlexGrid1.Text = "SUB-TOTAL"
reporteventa.MSFlexGrid1.ColWidth(4) = 1200
reporteventa.MSFlexGrid1.ColAlignment(4) = 3

reporteventa.MSFlexGrid1.Col = 5
reporteventa.MSFlexGrid1.Row = 0
reporteventa.MSFlexGrid1.Text = "T.DESC."
reporteventa.MSFlexGrid1.ColWidth(5) = 1200
reporteventa.MSFlexGrid1.ColAlignment(5) = 3

reporteventa.MSFlexGrid1.Col = 6
reporteventa.MSFlexGrid1.Row = 0
reporteventa.MSFlexGrid1.Text = "TOTAL NETO"
reporteventa.MSFlexGrid1.ColWidth(6) = 1500
reporteventa.MSFlexGrid1.ColAlignment(6) = 3

reporteventa.MSFlexGrid1.Col = 7
reporteventa.MSFlexGrid1.Row = 0
reporteventa.MSFlexGrid1.Text = "T.GANACIA"
reporteventa.MSFlexGrid1.ColWidth(7) = 1500
reporteventa.MSFlexGrid1.ColAlignment(7) = 3

End Sub

Public Sub cuadroreportemensual()
reporteventamensual.MSFlexGrid1.ColWidth(0) = 1

reporteventamensual.MSFlexGrid1.Col = 1
reporteventamensual.MSFlexGrid1.Row = 0
reporteventamensual.MSFlexGrid1.Text = "FACT."
reporteventamensual.MSFlexGrid1.ColWidth(1) = 1200
reporteventamensual.MSFlexGrid1.ColAlignment(1) = 3

reporteventamensual.MSFlexGrid1.Col = 2
reporteventamensual.MSFlexGrid1.Row = 0
reporteventamensual.MSFlexGrid1.Text = "FECHA"
reporteventamensual.MSFlexGrid1.ColWidth(2) = 1300
reporteventamensual.MSFlexGrid1.ColAlignment(2) = 3

reporteventamensual.MSFlexGrid1.Col = 3
reporteventamensual.MSFlexGrid1.Row = 0
reporteventamensual.MSFlexGrid1.Text = "HORA"
reporteventamensual.MSFlexGrid1.ColWidth(3) = 1300
reporteventamensual.MSFlexGrid1.ColAlignment(3) = 3

reporteventamensual.MSFlexGrid1.Col = 4
reporteventamensual.MSFlexGrid1.Row = 0
reporteventamensual.MSFlexGrid1.Text = "SUB-TOTAL"
reporteventamensual.MSFlexGrid1.ColWidth(4) = 1200
reporteventamensual.MSFlexGrid1.ColAlignment(4) = 3

reporteventamensual.MSFlexGrid1.Col = 5
reporteventamensual.MSFlexGrid1.Row = 0
reporteventamensual.MSFlexGrid1.Text = "T.DESC."
reporteventamensual.MSFlexGrid1.ColWidth(5) = 1200
reporteventamensual.MSFlexGrid1.ColAlignment(5) = 3

reporteventamensual.MSFlexGrid1.Col = 6
reporteventamensual.MSFlexGrid1.Row = 0
reporteventamensual.MSFlexGrid1.Text = "TOTAL NETO"
reporteventamensual.MSFlexGrid1.ColWidth(6) = 1500
reporteventamensual.MSFlexGrid1.ColAlignment(6) = 3

reporteventamensual.MSFlexGrid1.Col = 7
reporteventamensual.MSFlexGrid1.Row = 0
reporteventamensual.MSFlexGrid1.Text = "T.GANACIA"
reporteventamensual.MSFlexGrid1.ColWidth(7) = 1500
reporteventamensual.MSFlexGrid1.ColAlignment(7) = 3

End Sub

Public Sub limpiarcuadroreporte()
Dim v As Integer
Dim w As Integer
For v = 0 To reporteventa.MSFlexGrid1.Rows - 1
  For w = 0 To reporteventa.MSFlexGrid1.Cols - 1
  reporteventa.MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
reporteventa.MSFlexGrid1.Rows = 2
reporteventa.MSFlexGrid1.Cols = 8
reporteventa.Text1.Text = Empty
reporteventa.Text2.Text = Empty
reporteventa.Text3.Text = Empty
End Sub

Public Sub limpiarcuadroreportemensual()
Dim v As Integer
Dim w As Integer
For v = 0 To reporteventamensual.MSFlexGrid1.Rows - 1
  For w = 0 To reporteventamensual.MSFlexGrid1.Cols - 1
  reporteventamensual.MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
reporteventamensual.MSFlexGrid1.Rows = 2
reporteventamensual.MSFlexGrid1.Cols = 8
reporteventamensual.Text1.Text = Empty
reporteventamensual.Text2.Text = Empty
reporteventamensual.Text3.Text = Empty
End Sub

Public Sub cuadrodetallado()
reporteventadetallado.MSFlexGrid1.ColWidth(0) = 1

reporteventadetallado.MSFlexGrid1.Col = 1
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "FACT."
reporteventadetallado.MSFlexGrid1.ColWidth(1) = 1000
reporteventadetallado.MSFlexGrid1.ColAlignment(1) = 3

reporteventadetallado.MSFlexGrid1.Col = 2
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "FECHA"
reporteventadetallado.MSFlexGrid1.ColWidth(2) = 1300
reporteventadetallado.MSFlexGrid1.ColAlignment(2) = 3

reporteventadetallado.MSFlexGrid1.Col = 3
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "HORA"
reporteventadetallado.MSFlexGrid1.ColWidth(3) = 1300
reporteventadetallado.MSFlexGrid1.ColAlignment(3) = 3

reporteventadetallado.MSFlexGrid1.Col = 4
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "DETALLE"
reporteventadetallado.MSFlexGrid1.ColWidth(4) = 5000
reporteventadetallado.MSFlexGrid1.ColAlignment(4) = 3

reporteventadetallado.MSFlexGrid1.Col = 5
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "CANT."
reporteventadetallado.MSFlexGrid1.ColWidth(5) = 800
reporteventadetallado.MSFlexGrid1.ColAlignment(5) = 3

reporteventadetallado.MSFlexGrid1.Col = 6
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "PRECIO"
reporteventadetallado.MSFlexGrid1.ColWidth(6) = 1000
reporteventadetallado.MSFlexGrid1.ColAlignment(6) = 3
'areeglar
reporteventadetallado.MSFlexGrid1.Col = 7
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "PRECIO C."
reporteventadetallado.MSFlexGrid1.ColWidth(7) = 1000
reporteventadetallado.MSFlexGrid1.ColAlignment(7) = 3

reporteventadetallado.MSFlexGrid1.Col = 8
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "DESC."
reporteventadetallado.MSFlexGrid1.ColWidth(8) = 800
reporteventadetallado.MSFlexGrid1.ColAlignment(8) = 3

reporteventadetallado.MSFlexGrid1.Col = 9
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "SUB-TOTAL"
reporteventadetallado.MSFlexGrid1.ColWidth(9) = 1200
reporteventadetallado.MSFlexGrid1.ColAlignment(9) = 3

reporteventadetallado.MSFlexGrid1.Col = 10
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "T.DESC"
reporteventadetallado.MSFlexGrid1.ColWidth(10) = 800
reporteventadetallado.MSFlexGrid1.ColAlignment(10) = 3

reporteventadetallado.MSFlexGrid1.Col = 11
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "T.NETO"
reporteventadetallado.MSFlexGrid1.ColWidth(11) = 800
reporteventadetallado.MSFlexGrid1.ColAlignment(11) = 3

reporteventadetallado.MSFlexGrid1.Col = 12
reporteventadetallado.MSFlexGrid1.Row = 0
reporteventadetallado.MSFlexGrid1.Text = "T.GANANCIA"
reporteventadetallado.MSFlexGrid1.ColWidth(12) = 1300
reporteventadetallado.MSFlexGrid1.ColAlignment(12) = 3
End Sub

Public Sub cuadroproducto()
reporteproducto.MSFlexGrid1.ColWidth(0) = 1

reporteproducto.MSFlexGrid1.Col = 1
reporteproducto.MSFlexGrid1.Row = 0
reporteproducto.MSFlexGrid1.Text = "FACT."
reporteproducto.MSFlexGrid1.ColWidth(1) = 1000
reporteproducto.MSFlexGrid1.ColAlignment(1) = 3

reporteproducto.MSFlexGrid1.Col = 2
reporteproducto.MSFlexGrid1.Row = 0
reporteproducto.MSFlexGrid1.Text = "FECHA"
reporteproducto.MSFlexGrid1.ColWidth(2) = 1300
reporteproducto.MSFlexGrid1.ColAlignment(2) = 3

reporteproducto.MSFlexGrid1.Col = 3
reporteproducto.MSFlexGrid1.Row = 0
reporteproducto.MSFlexGrid1.Text = "HORA"
reporteproducto.MSFlexGrid1.ColWidth(3) = 1300
reporteproducto.MSFlexGrid1.ColAlignment(3) = 3

reporteproducto.MSFlexGrid1.Col = 4
reporteproducto.MSFlexGrid1.Row = 0
reporteproducto.MSFlexGrid1.Text = "DETALLE"
reporteproducto.MSFlexGrid1.ColWidth(4) = 5000
reporteproducto.MSFlexGrid1.ColAlignment(4) = 3

reporteproducto.MSFlexGrid1.Col = 5
reporteproducto.MSFlexGrid1.Row = 0
reporteproducto.MSFlexGrid1.Text = "CANT."
reporteproducto.MSFlexGrid1.ColWidth(5) = 800
reporteproducto.MSFlexGrid1.ColAlignment(5) = 3

reporteproducto.MSFlexGrid1.Col = 6
reporteproducto.MSFlexGrid1.Row = 0
reporteproducto.MSFlexGrid1.Text = "PRECIO"
reporteproducto.MSFlexGrid1.ColWidth(6) = 1000
reporteproducto.MSFlexGrid1.ColAlignment(6) = 3
'areeglar
reporteproducto.MSFlexGrid1.Col = 7
reporteproducto.MSFlexGrid1.Row = 0
reporteproducto.MSFlexGrid1.Text = "PRECIO C."
reporteproducto.MSFlexGrid1.ColWidth(7) = 1000
reporteproducto.MSFlexGrid1.ColAlignment(7) = 3

End Sub

Public Sub cuadroingresos()
reportingresos.MSFlexGrid1.ColWidth(0) = 1

reportingresos.MSFlexGrid1.Col = 1
reportingresos.MSFlexGrid1.Row = 0
reportingresos.MSFlexGrid1.Text = "CODIGO"
reportingresos.MSFlexGrid1.ColWidth(1) = 2000
reportingresos.MSFlexGrid1.ColAlignment(1) = 3

reportingresos.MSFlexGrid1.Col = 2
reportingresos.MSFlexGrid1.Row = 0
reportingresos.MSFlexGrid1.Text = "FECHA"
reportingresos.MSFlexGrid1.ColWidth(2) = 1300
reportingresos.MSFlexGrid1.ColAlignment(2) = 3

reportingresos.MSFlexGrid1.Col = 3
reportingresos.MSFlexGrid1.Row = 0
reportingresos.MSFlexGrid1.Text = "HORA"
reportingresos.MSFlexGrid1.ColWidth(3) = 1300
reportingresos.MSFlexGrid1.ColAlignment(3) = 3

reportingresos.MSFlexGrid1.Col = 4
reportingresos.MSFlexGrid1.Row = 0
reportingresos.MSFlexGrid1.Text = "DESCRIPCION"
reportingresos.MSFlexGrid1.ColWidth(4) = 5000
reportingresos.MSFlexGrid1.ColAlignment(4) = 3

reportingresos.MSFlexGrid1.Col = 5
reportingresos.MSFlexGrid1.Row = 0
reportingresos.MSFlexGrid1.Text = "CANT."
reportingresos.MSFlexGrid1.ColWidth(5) = 1000
reportingresos.MSFlexGrid1.ColAlignment(5) = 3

End Sub

Public Sub limpiarcuadrodetallado()
Dim v As Integer
Dim w As Integer
For v = 0 To reporteventadetallado.MSFlexGrid1.Rows - 1
  For w = 0 To reporteventadetallado.MSFlexGrid1.Cols - 1
  reporteventadetallado.MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
reporteventadetallado.MSFlexGrid1.Rows = 2
reporteventadetallado.MSFlexGrid1.Cols = 13
reporteventadetallado.Text1.Text = Empty
reporteventadetallado.Text2.Text = Empty
reporteventadetallado.Text3.Text = Empty
reporteventadetallado.Text4.Text = Empty
End Sub

Public Sub limpiarcuadroproducto()
Dim v As Integer
Dim w As Integer
For v = 0 To reporteproducto.MSFlexGrid1.Rows - 1
  For w = 0 To reporteproducto.MSFlexGrid1.Cols - 1
  reporteproducto.MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
reporteproducto.MSFlexGrid1.Rows = 2
reporteproducto.MSFlexGrid1.Cols = 8
reporteproducto.Text1.Text = Empty
reporteproducto.Text2.Text = Empty
reporteproducto.Text3.Text = Empty
reporteproducto.Text4.Text = Empty
reporteproducto.Text5.Text = Empty
End Sub

Public Sub limpiarcuadroingresos()
Dim v As Integer
Dim w As Integer
For v = 0 To reportingresos.MSFlexGrid1.Rows - 1
  For w = 0 To reportingresos.MSFlexGrid1.Cols - 1
  reportingresos.MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
reportingresos.MSFlexGrid1.Rows = 2
reportingresos.MSFlexGrid1.Cols = 6
reportingresos.Text1.Text = Empty
reportingresos.Text2.Text = Empty
reportingresos.Text4.Text = Empty
End Sub

Public Sub cuadrofactnula()
rafactura.MSFlexGrid1.ColWidth(0) = 5

rafactura.MSFlexGrid1.Col = 1
rafactura.MSFlexGrid1.Row = 0
rafactura.MSFlexGrid1.Text = "CANT."
rafactura.MSFlexGrid1.ColWidth(1) = 1000
rafactura.MSFlexGrid1.ColAlignment(1) = 3

rafactura.MSFlexGrid1.Col = 2
rafactura.MSFlexGrid1.Row = 0
rafactura.MSFlexGrid1.Text = "COD. PRODUCTO"
rafactura.MSFlexGrid1.ColWidth(2) = 1800
rafactura.MSFlexGrid1.ColAlignment(2) = 3

rafactura.MSFlexGrid1.Col = 3
rafactura.MSFlexGrid1.Row = 0
rafactura.MSFlexGrid1.Text = "D E T A L L E"
rafactura.MSFlexGrid1.ColWidth(3) = 6750
rafactura.MSFlexGrid1.ColAlignment(3) = 3

rafactura.MSFlexGrid1.Col = 4
rafactura.MSFlexGrid1.Row = 0
rafactura.MSFlexGrid1.Text = "PRECIO"
rafactura.MSFlexGrid1.ColWidth(4) = 1000
rafactura.MSFlexGrid1.ColAlignment(4) = 3

rafactura.MSFlexGrid1.Col = 5
rafactura.MSFlexGrid1.Row = 0
rafactura.MSFlexGrid1.Text = "DESC."
rafactura.MSFlexGrid1.ColWidth(5) = 1000
rafactura.MSFlexGrid1.ColAlignment(5) = 3

rafactura.MSFlexGrid1.Col = 6
rafactura.MSFlexGrid1.Row = 0
rafactura.MSFlexGrid1.Text = "TOTAL"
rafactura.MSFlexGrid1.ColWidth(6) = 1000
rafactura.MSFlexGrid1.ColAlignment(6) = 3

rafactura.MSFlexGrid1.Col = 7
rafactura.MSFlexGrid1.Row = 0
'despues cambiar a 1
rafactura.MSFlexGrid1.ColWidth(7) = 1000
rafactura.MSFlexGrid1.Text = "PRECIO C."
rafactura.MSFlexGrid1.ColAlignment(7) = 3
'despues cambiar a 1

rafactura.MSFlexGrid1.ColWidth(8) = 1700
rafactura.MSFlexGrid1.Col = 8
rafactura.MSFlexGrid1.Row = 0
rafactura.MSFlexGrid1.Text = "GANANCIA INDV."
rafactura.MSFlexGrid1.ColAlignment(8) = 3
End Sub

Public Sub limpiarcuadrofactnula()
Dim h As Integer
Dim I As Integer
rafactura.MSFlexGrid1.Rows = 2
For h = 0 To rafactura.MSFlexGrid1.Rows - 1
  For I = 0 To rafactura.MSFlexGrid1.Cols - 1
  rafactura.MSFlexGrid1.TextMatrix(h, I) = ""
  Next
Next
End Sub

Public Sub limpiarcuadrodetalladonulas()
Dim v As Integer
Dim w As Integer
For v = 0 To reporteventadetalladonulas.MSFlexGrid1.Rows - 1
  For w = 0 To reporteventadetalladonulas.MSFlexGrid1.Cols - 1
  reporteventadetalladonulas.MSFlexGrid1.TextMatrix(v, w) = ""
  Next
Next
reporteventadetalladonulas.MSFlexGrid1.Rows = 2
reporteventadetalladonulas.MSFlexGrid1.Cols = 13
reporteventadetalladonulas.Text1.Text = Empty
reporteventadetalladonulas.Text2.Text = Empty
reporteventadetalladonulas.Text3.Text = Empty
reporteventadetalladonulas.Text4.Text = Empty
End Sub

Public Sub cuadrodetalladonulas()
reporteventadetalladonulas.MSFlexGrid1.ColWidth(0) = 1

reporteventadetalladonulas.MSFlexGrid1.Col = 1
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "FACT."
reporteventadetalladonulas.MSFlexGrid1.ColWidth(1) = 1000
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(1) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 2
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "FECHA"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(2) = 1000
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(2) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 3
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "HORA"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(3) = 1000
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(3) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 4
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "DETALLE"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(4) = 5000
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(4) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 5
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "CANT."
reporteventadetalladonulas.MSFlexGrid1.ColWidth(5) = 800
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(5) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 6
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "PRECIO"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(6) = 800
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(6) = 3
'areeglar
reporteventadetalladonulas.MSFlexGrid1.Col = 7
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "PRECIO C."
reporteventadetalladonulas.MSFlexGrid1.ColWidth(7) = 1000
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(7) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 8
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "DESC."
reporteventadetalladonulas.MSFlexGrid1.ColWidth(8) = 800
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(8) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 9
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "SUB-TOTAL"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(9) = 1200
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(9) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 10
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "T.DESC"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(10) = 800
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(10) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 11
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "T.NETO"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(11) = 800
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(11) = 3

reporteventadetalladonulas.MSFlexGrid1.Col = 12
reporteventadetalladonulas.MSFlexGrid1.Row = 0
reporteventadetalladonulas.MSFlexGrid1.Text = "T.GANANCIA"
reporteventadetalladonulas.MSFlexGrid1.ColWidth(12) = 1300
reporteventadetalladonulas.MSFlexGrid1.ColAlignment(12) = 3
End Sub

Public Sub incioingresos()

             Ingresos.Show
            Ingresos.MSFlexGrid1.Rows = 2
            Ingresos.MSFlexGrid1.ColWidth(0) = 1
            Ingresos.MSFlexGrid1.ColWidth(1) = 2000
            Ingresos.MSFlexGrid1.Col = 1
            Ingresos.MSFlexGrid1.Row = 0
            Ingresos.MSFlexGrid1.Text = "CODIGO"
            Ingresos.MSFlexGrid1.Col = 2
            Ingresos.MSFlexGrid1.Row = 0
            Ingresos.MSFlexGrid1.Text = "DESCRIPCION"
            Ingresos.MSFlexGrid1.ColWidth(2) = 5000
            Ingresos.MSFlexGrid1.ColAlignment(2) = 5
            Ingresos.MSFlexGrid1.Col = 3
            Ingresos.MSFlexGrid1.Row = 0
            Ingresos.MSFlexGrid1.Text = "CANT."
            Ingresos.Text1.SetFocus

End Sub

Public Function NombreArchivoVisible(ByVal nombre As String) As String
    Dim patron As Object
    Set patron = CreateObject("VBScript.RegExp")
    patron.Pattern = "(\d{2}-\d{2}-\d{4} \d{2})-(\d{2})"
    NombreArchivoVisible = patron.Replace(nombre, "$1:$2")
End Function

Public Function NombreArchivoReal(ByVal nombre As String) As String
    NombreArchivoReal = Replace$(nombre, ":", "-")
End Function
