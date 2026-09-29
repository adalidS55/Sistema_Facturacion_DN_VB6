Attribute VB_Name = "ExportarInventario"
Option Explicit

Private Type DialogoArchivo
    tamano As Long
    propietario As Long
    instancia As Long
    filtro As String
    filtroPersonalizado As String
    maxFiltro As Long
    indiceFiltro As Long
    archivo As String
    maxArchivo As Long
    tituloArchivo As String
    maxTitulo As Long
    carpetaInicial As String
    titulo As String
    opciones As Long
    inicioNombre As Integer
    inicioExtension As Integer
    extensionPredeterminada As String
    datos As Long
    gancho As Long
    plantilla As String
End Type

Private Type EntradaZip
    nombre As String
    crc As Long
    tamano As Long
    posicion As Long
End Type

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef dialogo As DialogoArchivo) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExW" (ByVal origen As Long, ByVal destino As Long, ByVal opciones As Long) As Long
Private crcTabla(0 To 255) As Long
Private tablaLista As Boolean

Public Function ElegirExcelInventario(ByVal propietario As Long) As String
    Dim dialogo As DialogoArchivo, errorDialogo As Long
    With dialogo
        .tamano = Len(dialogo)
        .propietario = propietario
        .filtro = "Libro de Excel (*.xlsx)" & vbNullChar & "*.xlsx" & vbNullChar & vbNullChar
        .indiceFiltro = 1
        .archivo = "Control-Inventario-" & Format$(Now, "dd-mm-yyyy hh-nn") & ".xlsx" & String$(1024, vbNullChar)
        .maxArchivo = Len(.archivo)
        .carpetaInicial = App.Path
        .titulo = "Guardar control de inventario"
        .extensionPredeterminada = "xlsx"
        'Confirmar reemplazo, exigir carpeta existente y conservar directorio actual.
        .opciones = &H2 Or &H800 Or &H8 Or &H80000
    End With
    If GetSaveFileName(dialogo) <> 0 Then
        ElegirExcelInventario = Left$(dialogo.archivo, InStr(dialogo.archivo, vbNullChar) - 1)
        If LCase$(Right$(ElegirExcelInventario, 5)) <> ".xlsx" Then
            Err.Raise 5, , "El nombre del archivo debe terminar en .xlsx."
        End If
    Else
        errorDialogo = CommDlgExtendedError()
        If errorDialogo <> 0 Then Err.Raise 5, , "No se pudo abrir Guardar como. Código: " & errorDialogo
    End If
End Function

Private Function TextoXml(ByVal valor As String) As String
    Dim i As Long
    For i = 0 To 31
        If i <> 9 And i <> 10 And i <> 13 Then valor = Replace$(valor, Chr$(i), "")
    Next i
    valor = Replace$(valor, "&", "&amp;")
    valor = Replace$(valor, "<", "&lt;")
    valor = Replace$(valor, ">", "&gt;")
    valor = Replace$(valor, """", "&quot;")
    TextoXml = valor
End Function

Private Function NumeroXml(ByVal valor As Variant) As String
    NumeroXml = Trim$(Str$(CDbl(valor)))
End Function

Private Function TieneNumero(ByVal valor As Variant) As Boolean
    If IsNull(valor) Or IsEmpty(valor) Then Exit Function
    If VarType(valor) = vbString Then
        If Len(valor) = 0 Then Exit Function
    End If
    TieneNumero = True
End Function

Private Function CeldaTexto(ByVal referencia As String, ByVal valor As String, Optional ByVal estilo As Long = 0) As String
    CeldaTexto = "<c r=""" & referencia & """ s=""" & estilo & """ t=""inlineStr""><is><t xml:space=""preserve"">" & TextoXml(valor) & "</t></is></c>"
End Function

Private Function CeldaNumero(ByVal referencia As String, ByVal valor As Variant, ByVal estilo As Long) As String
    If Not TieneNumero(valor) Then Exit Function
    CeldaNumero = "<c r=""" & referencia & """ s=""" & estilo & """><v>" & NumeroXml(valor) & "</v></c>"
End Function

Private Function CeldaFormula(ByVal referencia As String, ByVal formula As String, ByVal valor As Variant, ByVal estilo As Long) As String
    Dim tipo As String, cache As String
    If TieneNumero(valor) Then
        cache = NumeroXml(valor)
    Else
        tipo = " t=""str"""
    End If
    CeldaFormula = "<c r=""" & referencia & """ s=""" & estilo & """" & tipo & "><f>" & TextoXml(formula) & "</f><v>" & cache & "</v></c>"
End Function

Private Function HojaInventario(ByRef filas As Variant) As String
    Dim xml As String, i As Long, r As Long, ultima As Long
    Dim diferencia As Variant, dinero As Variant, formula As String
    Dim totalDinero As Double
    ultima = UBound(filas, 1) + 1
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    xml = xml & "<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"
    xml = xml & "<dimension ref=""A1:H" & (ultima + 6) & """/><sheetViews><sheetView workbookViewId=""0"">"
    xml = xml & "<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/><selection pane=""bottomLeft"" activeCell=""A2"" sqref=""A2""/></sheetView></sheetViews>"
    xml = xml & "<sheetFormatPr defaultRowHeight=""15""/><cols>"
    xml = xml & "<col min=""1"" max=""1"" width=""21"" customWidth=""1""/><col min=""2"" max=""2"" width=""44"" customWidth=""1""/>"
    xml = xml & "<col min=""3"" max=""5"" width=""20"" customWidth=""1""/><col min=""6"" max=""6"" width=""4"" customWidth=""1""/>"
    xml = xml & "<col min=""7"" max=""8"" width=""20"" customWidth=""1""/></cols><sheetData>"
    xml = xml & "<row r=""1"" ht=""36"" customHeight=""1"">"
    xml = xml & CeldaTexto("A1", "Código Producto", 1) & CeldaTexto("B1", "Nombre", 1)
    xml = xml & CeldaTexto("C1", "Precio de venta", 1) & CeldaTexto("D1", "Cantidad antigua", 1)
    xml = xml & CeldaTexto("E1", "Cantidades Nuevas", 1)
    xml = xml & CeldaTexto("G1", "Diferencia", 1) & CeldaTexto("H1", "Dinero", 1) & "</row>"
    For i = 1 To UBound(filas, 1)
        r = i + 1
        diferencia = Empty
        dinero = Empty
        If TieneNumero(filas(i, 4)) And TieneNumero(filas(i, 5)) Then
            diferencia = CDbl(filas(i, 5)) - CDbl(filas(i, 4))
            If TieneNumero(filas(i, 3)) Then dinero = CDbl(filas(i, 3)) * CDbl(diferencia)
        End If
        xml = xml & "<row r=""" & r & """>"
        xml = xml & CeldaTexto("A" & r, CStr(filas(i, 1))) & CeldaTexto("B" & r, CStr(filas(i, 2)))
        xml = xml & CeldaNumero("C" & r, filas(i, 3), 2) & CeldaNumero("D" & r, filas(i, 4), 3) & CeldaNumero("E" & r, filas(i, 5), 3)
        formula = "IF(OR(D" & r & "="""",E" & r & "=""""),"""",E" & r & "-D" & r & ")"
        xml = xml & CeldaFormula("G" & r, formula, diferencia, 3)
        formula = "IF(OR(C" & r & "="""",G" & r & "=""""),"""",C" & r & "*G" & r & ")"
        xml = xml & CeldaFormula("H" & r, formula, dinero, 2) & "</row>"
        If TieneNumero(dinero) Then totalDinero = totalDinero + CDbl(dinero)
    Next i
    r = ultima + 2
    xml = xml & "<row r=""" & r & """>" & CeldaTexto("G" & r, "Total", 1)
    xml = xml & CeldaFormula("H" & r, "SUM(H2:H" & ultima & ")", totalDinero, 2) & "</row>"
    r = ultima + 4
    xml = xml & "<row r=""" & r & """>" & CeldaTexto("G" & r, "Sobrante", 1)
    xml = xml & "<c r=""H" & r & """ s=""2""/></row>"
    r = ultima + 6
    xml = xml & "<row r=""" & r & """>" & CeldaTexto("G" & r, "Diferencia", 1)
    xml = xml & CeldaFormula("H" & r, "SUM(H" & (ultima + 2) & ",H" & (ultima + 4) & ")", totalDinero, 2) & "</row>"
    xml = xml & "</sheetData><autoFilter ref=""A1:H" & ultima & """/>"
    xml = xml & "<conditionalFormatting sqref=""H2:H1048576"">"
    xml = xml & "<cfRule type=""cellIs"" dxfId=""0"" priority=""1"" operator=""greaterThan""><formula>0</formula></cfRule>"
    xml = xml & "<cfRule type=""cellIs"" dxfId=""1"" priority=""2"" operator=""lessThan""><formula>0</formula></cfRule>"
    HojaInventario = xml & "</conditionalFormatting></worksheet>"
End Function

Private Function EstilosInventario() As String
    Dim xml As String
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    xml = xml & "<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">"
    xml = xml & "<numFmts count=""2""><numFmt numFmtId=""164"" formatCode=""&quot;L&quot; #,##0.00""/><numFmt numFmtId=""165"" formatCode=""#,##0.####""/></numFmts>"
    xml = xml & "<fonts count=""2""><font><sz val=""11""/><name val=""Calibri""/></font><font><b/><sz val=""11""/><color rgb=""FFFFFFFF""/><name val=""Calibri""/></font></fonts>"
    xml = xml & "<fills count=""3""><fill><patternFill patternType=""none""/></fill><fill><patternFill patternType=""gray125""/></fill><fill><patternFill patternType=""solid""><fgColor rgb=""FF004040""/><bgColor indexed=""64""/></patternFill></fill></fills>"
    xml = xml & "<borders count=""1""><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
    xml = xml & "<cellStyleXfs count=""1""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0""/></cellStyleXfs>"
    xml = xml & "<cellXfs count=""4""><xf numFmtId=""0"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0""/>"
    xml = xml & "<xf numFmtId=""0"" fontId=""1"" fillId=""2"" borderId=""0"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyAlignment=""1""><alignment vertical=""center"" wrapText=""1""/></xf>"
    xml = xml & "<xf numFmtId=""164"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyNumberFormat=""1""/>"
    xml = xml & "<xf numFmtId=""165"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyNumberFormat=""1""/></cellXfs>"
    xml = xml & "<cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>"
    xml = xml & "<dxfs count=""2"">"
    xml = xml & "<dxf><font><color rgb=""FF006100""/></font><fill><patternFill patternType=""solid""><fgColor rgb=""FFC6EFCE""/><bgColor rgb=""FFC6EFCE""/></patternFill></fill></dxf>"
    xml = xml & "<dxf><font><color rgb=""FF9C0006""/></font><fill><patternFill patternType=""solid""><fgColor rgb=""FFFFC7CE""/><bgColor rgb=""FFFFC7CE""/></patternFill></fill></dxf>"
    EstilosInventario = xml & "</dxfs></styleSheet>"
End Function

Private Function Utf8(ByVal contenido As String) As Variant
    Dim flujo As New ADODB.Stream
    flujo.Type = adTypeText
    flujo.Charset = "utf-8"
    flujo.Open
    flujo.WriteText contenido
    flujo.Position = 0
    flujo.Type = adTypeBinary
    flujo.Position = 3
    Utf8 = flujo.Read
    flujo.Close
End Function

Private Function DesplazarBit(ByVal valor As Long) As Long
    DesplazarBit = (valor And &H7FFFFFFF) \ 2
    If valor < 0 Then DesplazarBit = DesplazarBit Or &H40000000
End Function

Private Function Crc32(ByRef bytes() As Byte) As Long
    Dim i As Long, bit As Long, valor As Long
    If Not tablaLista Then
        For i = 0 To 255
            valor = i
            For bit = 1 To 8
                If (valor And 1) <> 0 Then
                    valor = DesplazarBit(valor) Xor &HEDB88320
                Else
                    valor = DesplazarBit(valor)
                End If
            Next bit
            crcTabla(i) = valor
        Next i
        tablaLista = True
    End If
    valor = &HFFFFFFFF
    For i = LBound(bytes) To UBound(bytes)
        valor = DesplazarBit(DesplazarBit(DesplazarBit(DesplazarBit(DesplazarBit(DesplazarBit(DesplazarBit(DesplazarBit(valor)))))))) Xor crcTabla((valor Xor CLng(bytes(i))) And &HFF&)
    Next i
    Crc32 = Not valor
End Function

Private Sub EnteroZip(ByVal archivo As Integer, ByVal valor As Long)
    Dim corto As Integer
    If valor > 32767 Then valor = valor - 65536
    corto = CInt(valor)
    Put #archivo, , corto
End Sub

Private Sub LargoZip(ByVal archivo As Integer, ByVal valor As Long)
    Put #archivo, , valor
End Sub

Private Sub AgregarZip(ByVal archivo As Integer, ByRef entrada As EntradaZip, ByVal contenido As String)
    Dim bytes() As Byte, nombre() As Byte
    bytes = Utf8(contenido)
    nombre = Utf8(entrada.nombre)
    entrada.crc = Crc32(bytes)
    entrada.tamano = UBound(bytes) + 1
    entrada.posicion = Seek(archivo) - 1
    LargoZip archivo, &H4034B50
    EnteroZip archivo, 20
    EnteroZip archivo, &H800
    EnteroZip archivo, 0
    EnteroZip archivo, 0
    EnteroZip archivo, 33
    LargoZip archivo, entrada.crc
    LargoZip archivo, entrada.tamano
    LargoZip archivo, entrada.tamano
    EnteroZip archivo, UBound(nombre) + 1
    EnteroZip archivo, 0
    Put #archivo, , nombre
    Put #archivo, , bytes
End Sub

Private Sub CerrarZip(ByVal archivo As Integer, ByRef entradas() As EntradaZip)
    Dim i As Long, inicio As Long, tamano As Long, nombre() As Byte
    inicio = Seek(archivo) - 1
    For i = LBound(entradas) To UBound(entradas)
        nombre = Utf8(entradas(i).nombre)
        LargoZip archivo, &H2014B50
        EnteroZip archivo, 20
        EnteroZip archivo, 20
        EnteroZip archivo, &H800
        EnteroZip archivo, 0
        EnteroZip archivo, 0
        EnteroZip archivo, 33
        LargoZip archivo, entradas(i).crc
        LargoZip archivo, entradas(i).tamano
        LargoZip archivo, entradas(i).tamano
        EnteroZip archivo, UBound(nombre) + 1
        EnteroZip archivo, 0
        EnteroZip archivo, 0
        EnteroZip archivo, 0
        EnteroZip archivo, 0
        LargoZip archivo, 0
        LargoZip archivo, entradas(i).posicion
        Put #archivo, , nombre
    Next i
    tamano = Seek(archivo) - 1 - inicio
    LargoZip archivo, &H6054B50
    EnteroZip archivo, 0
    EnteroZip archivo, 0
    EnteroZip archivo, UBound(entradas) - LBound(entradas) + 1
    EnteroZip archivo, UBound(entradas) - LBound(entradas) + 1
    LargoZip archivo, tamano
    LargoZip archivo, inicio
    EnteroZip archivo, 0
End Sub

Public Sub GuardarExcelInventario(ByVal destino As String, ByRef filas As Variant)
    Dim archivo As Integer, abierto As Boolean, temporal As String, intento As Long
    Dim entradas(0 To 5) As EntradaZip, xml As String
    Dim numeroError As Long, mensaje As String
    On Error GoTo Fallo
    Do
        intento = intento + 1
        temporal = destino & "." & Format$(Now, "yyyymmddhhnnss") & "-" & intento & ".tmp"
    Loop While Len(Dir$(temporal)) > 0
    archivo = FreeFile
    Open temporal For Binary Access Write Lock Read Write As #archivo
    abierto = True
    entradas(0).nombre = "[Content_Types].xml"
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">"
    xml = xml & "<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/><Default Extension=""xml"" ContentType=""application/xml""/>"
    xml = xml & "<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>"
    xml = xml & "<Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>"
    xml = xml & "<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/></Types>"
    AgregarZip archivo, entradas(0), xml
    entradas(1).nombre = "_rels/.rels"
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/></Relationships>"
    AgregarZip archivo, entradas(1), xml
    entradas(2).nombre = "xl/workbook.xml"
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets><sheet name=""Control de Inventario"" sheetId=""1"" r:id=""rId1""/></sheets><calcPr calcId=""191029"" fullCalcOnLoad=""1"" calcMode=""auto""/></workbook>"
    AgregarZip archivo, entradas(2), xml
    entradas(3).nombre = "xl/_rels/workbook.xml.rels"
    xml = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""><Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/><Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""styles.xml""/></Relationships>"
    AgregarZip archivo, entradas(3), xml
    entradas(4).nombre = "xl/styles.xml"
    AgregarZip archivo, entradas(4), EstilosInventario()
    entradas(5).nombre = "xl/worksheets/sheet1.xml"
    AgregarZip archivo, entradas(5), HojaInventario(filas)
    CerrarZip archivo, entradas
    Close #archivo
    abierto = False
    If MoveFileEx(StrPtr(temporal), StrPtr(destino), 9) = 0 Then
        Err.Raise 75, , "No se pudo guardar el archivo. Compruebe que el destino no esté abierto y que tenga permiso para escribir en esa carpeta."
    End If
    Exit Sub
Fallo:
    numeroError = Err.Number
    mensaje = Err.Description
    On Error Resume Next
    If abierto Then Close #archivo
    If Len(temporal) > 0 Then
        If Len(Dir$(temporal)) > 0 Then Kill temporal
    End If
    On Error GoTo 0
    Err.Raise numeroError, "GuardarExcelInventario", mensaje
End Sub
