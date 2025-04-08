VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private xlApp1 As Excel.Application
Private xlApp2 As Excel.Application
Private xlBook As Excel.Workbook
Private xlSheet As Excel.Worksheet
Private xTipoRep As Boolean
Private xQuery As String
Private xNomCia As String

Private xCodTrab As String
Private xDniTrab As String
Private xNomTrab As String

Public Property Let CodigoTrbajador(ByVal vNewValue As Variant)
xCodTrab = vNewValue
End Property

Public Property Let DniTrbajador(ByVal vNewValue As Variant)
xCodTrab = vNewValue
End Property


Public Function p_AbrirExcel() As Boolean
    On Error GoTo Error
    
    Set xlApp1 = CreateObject("Excel.Application")
    xlApp1.Workbooks.Add
    Set xlApp2 = xlApp1.Application
    Set xlBook = xlApp2.Workbooks(1)
    Set xlSheet = xlBook.Worksheets("Hoja1")
    'xlApp2.Application.Visible = True
    p_AbrirExcel = True
    
    Exit Function
Error:
    MsgBox Err.Description, vbExclamation, "Mensaje de Error !!!"
End Function


Public Sub p_CerrarExcel()
'    Dim l_sCadena As String
'    Dim k As Integer
'
'    For k = 1 To UBound(m_aColumnas, 1)
'        l_sCadena = Chr$(k + 64) & ":" & Chr$(k + 64)
'        xlSheet.Range(l_sCadena).ColumnWidth = m_aColumnas(k).Ancho
'    Next
'    xlSheet.PageSetup.leftMargin = Application.InchesToPoints(0.196850393700787) 'Application.InchesToPoints(0.15748031496063)
'    xlSheet.PageSetup.rightMargin = 0 'Application.InchesToPoints(0.78740157480315)
'    xlSheet.PageSetup.topMargin = Application.InchesToPoints(0.196850393700787)
'    xlSheet.PageSetup.bottomMargin = Application.InchesToPoints(0.590551181102362)
'    xlSheet.PageSetup.HeaderMargin = Application.InchesToPoints(0.196850393700787) 'Application.InchesToPoints(0.15748031496063)
'    xlSheet.PageSetup.FooterMargin = 0
'    xlSheet.PageSetup.CenterHorizontally = True
'    xlSheet.PageSetup.CenterVertically = False
'    xlSheet.PageSetup.Orientation = xlPortrait
'    xlSheet.PageSetup.PrintTitleRows = m_sEncabezado
'    xlSheet.PageSetup.RightHeader = "Página &P de &N"

     With xlSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        

'        .LeftMargin = Application.InchesToPoints(0.5)
'        .RightMargin = Application.InchesToPoints(0.75)
'        .TopMargin = Application.InchesToPoints(1.5)
'        .BottomMargin = Application.InchesToPoints(1)
'        .HeaderMargin = Application.InchesToPoints(0.5)
'        .FooterMargin = Application.InchesToPoints(0.5)


        
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        '.PrintQuality = Array(120, 72)
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        '.PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 75 '80
    End With
    'xlSheet.PageSetup.CenterHeader = "Página &P de &N"
    
    xlApp1.ActiveSheet.Range("A1:A1").Select
    
'    xlApp2.Application.DisplayStatusBar = False
'    xlApp2.Application.DisplayFormulaBar = False
'    xlApp2.Application.CommandBars("Standard").Visible = False
'    xlApp2.Application.CommandBars("Formatting").Visible = False
'    xlApp2.Application.CommandBars("Stop Recording").Visible = False
'    xlApp2.Application.CommandBars("EuroValue").Visible = False
    
'    If (Trim(UCase(wuser)) <> "SA" And Trim(UCase(wuser)) <> "HANDIA" And Trim(UCase(wuser)) <> "RGUIA" And Trim(UCase(wuser)) <> "CECILIA" And Trim(UCase(wuser)) <> "JTITO") Then
'        xlSheet.Protect "ROSSELLO", True, True
'        xlBook.Protect "ROSSELLO", True, True
'    End If
    
    
    xlApp2.Application.ActiveWindow.DisplayGridLines = False
    'xlApp1.activelSheet.Range("A1:A1").Select
    xlApp2.Application.ActiveWindow.DisplayZeros = True
    xlApp2.Application.Caption = "Derechohabientes " & Trae_CIA(wcia)
    xlApp1.ActiveWindow.Zoom = 80
    xlApp2.Application.Visible = True
    Screen.MousePointer = 0
    
    If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    If Not xlBook Is Nothing Then Set xlBook = Nothing
    If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
End Sub


'Public Sub p_Reporte_StockvsDespacho(ByVal pRs As ADODB.Recordset)
'
'
'
'
'
'    With xlApp1.ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal)
'
'        '.Connection = _
'        "ODBC;DSN=MiConec;UID=Administrador;APP=Microsoft Office 2003;WSID=(local);DATABASE=bdmgallos;Trusted_Connection=Yes"
'        '.Connection = "ODBC;DRIVER=SQL Server;SERVER=(local);UID=SA;PAASWORD=MGALLOS;APP=Microsoft Office 2003;WSID=PC-ADMINJC;DATABASE=BDMGALLOS;Trusted_Connection=Yes"
'        .Connection = "ODBC;DRIVER=SQL Server;SERVER=" & gsSQL_Server & ";APP=Microsoft Office 2003;WSID=PC-ADMINJC;DATABASE=" & gsSQL_database & ";Trusted_Connection=Yes"
'        .CommandType = xlCmdSql
'        .CommandText = xQuery
''        .CommandText = Array( _
''        "SELECT RptTablePivot_StockvsDespacho.cod_prod, RptTablePivot_StockvsDespacho.familia, RptTablePivot_StockvsDespacho.material, RptTablePivot_StockvsDespacho.acabado, RptTablePivot_StockvsDespacho.borde" _
''        , _
''        ", RptTablePivot_StockvsDespacho.modelo, RptTablePivot_StockvsDespacho.grado, RptTablePivot_StockvsDespacho.dimension, RptTablePivot_StockvsDespacho.unidad, RptTablePivot_StockvsDespacho.cantidad, RptT" _
''        , _
''        "ablePivot_StockvsDespacho.despacho_prom, RptTablePivot_StockvsDespacho.ratio, RptTablePivot_StockvsDespacho.fecha_ult_ing, RptTablePivot_StockvsDespacho.fecha_ult_sal, RptTablePivot_StockvsDespacho.st" _
''      , _
''        "ring_dias_inmovil" & Chr(13) & "" & Chr(10) & "FROM bdmgallos.dbo.RptTablePivot_StockvsDespacho RptTablePivot_StockvsDespacho" & Chr(13) & "" & Chr(10) & "ORDER BY RptTablePivot_StockvsDespacho.ratio DESC" _
''        )
'
'        .CreatePivotTable TableDestination:=xlApp1.ActiveSheet.Range("A3"), TableName:="Tabla dinámica1", DefaultVersion:=xlPivotTableVersion10
'    End With
'
'    'ActiveSheet.PivotTables("Tabla dinámica1").Format xlTable2
'    Set xlSheet = xlApp1.ActiveSheet
'
'    'With xlApp1.ActiveSheet.PivotTables("Tabla dinámica1")
'    With xlSheet.PivotTables("Tabla dinámica1")
'        '.PivotTableWizard TableDestination:=.cells(3, 1)
'        '.cells(3, 1).Select
'        '.PivotTableWizard TableDestination:=.cells(3, 1)
'        '.cells(3, 1).Select
'        With .PivotFields("Familia")
'            .Orientation = xlPageField
'            .Position = 1
'            .Name = "Familia"
'        End With
'
'        With .PivotFields("cab_material")
'            .Orientation = xlPageField
'            .Position = 2
'            .Name = "Cab. Material"
'        End With
'        With .PivotFields("det_material")
'            .Orientation = xlPageField
'            .Position = 3
'            .Name = "Det. Material"
'        End With
'        With .PivotFields("Acabado")
'            .Orientation = xlPageField
'            .Position = 4
'            .Name = "Acabado"
'        End With
'
'        With .PivotFields("borde")
'            .Orientation = xlPageField
'            .Position = 5
'            .Name = "Borde"
'        End With
'
'        With .PivotFields("modelo")
'            .Orientation = xlPageField
'            .Position = 6
'            .Name = "Modelo"
'        End With
'        With .PivotFields("grado")
'            .Orientation = xlPageField
'            .Position = 7
'            .Name = "Grado"
'        End With
'        With .PivotFields("nom_prod")
'            .Orientation = xlRowField
'            .Position = 1
'            .Name = "Producto"
'        End With
'        With .PivotFields("dimension")
'            .Orientation = xlRowField
'            .Position = 2
'            .Name = "Formato"
'        End With
''        xlSheet.PivotTables ("tabla dinamica 1")
''        Dim x As PivotTable
''        Dim y As PivotField
''        y.Subtotals = False
''        y.
'
'        xlApp1.ActiveSheet.Range("A12").Select
'        xlApp1.ActiveSheet.Range("A12").Delete
'
'
'        With .PivotFields("stock_xm2")
'            .Orientation = xlRowField
'            .Position = 3
'            .Name = "Stock_xM2"
'            .NumberFormat = "#,##0.00"
'        End With
'        'xlApp1.ActiveSheet.Range("B10").Select
'        'xlApp1.ActiveSheet.Range("B10").Delete
'
''        xlSheet.Range("B12").Select
''        xlSheet.Range("B12").Delete
'
'        xlApp1.ActiveSheet.Range("B12").Select
'        xlApp1.ActiveSheet.Range("B12").Delete
'
'
'        With .PivotFields("despacho_prom")
'            .Orientation = xlRowField
'            .Position = 4
'            .Name = "Despacho Prom."
'            .NumberFormat = "#,##0.00"
'        End With
'        xlApp1.ActiveSheet.Range("C12").Select
'        xlApp1.ActiveSheet.Range("C12").Delete
'
'        With .PivotFields("ratio")
'            .Orientation = xlRowField
'            .Position = 5
'            .Name = "Ratio"
'            .NumberFormat = "#,##0.00"
'        End With
'        xlApp1.ActiveSheet.Range("D12").Select
'        xlApp1.ActiveSheet.Range("D12").Delete
'        With .PivotFields("fecha_ult_ing")
'            .Orientation = xlRowField
'            .Position = 6
'            .Name = "Ultimo Ingreso Parte Producción"
'        End With
'        xlApp1.ActiveSheet.Range("E12").Select
'        xlApp1.ActiveSheet.Range("E12").Delete
'        With .PivotFields("fecha_ult_sal")
'            .Orientation = xlRowField
'            .Position = 7
'            .Name = "Ultimo Desopacho"
'        End With
'        xlApp1.ActiveSheet.Range("F12").Select
'        xlApp1.ActiveSheet.Range("F12").Delete
'        With .PivotFields( _
'            "string_dias_inmovil")
'            .Orientation = xlRowField
'            .Position = 8
'            .Name = "Stock Inmovil (dias)"
'        End With
'        xlApp1.ActiveSheet.Range("G12").Select
'        xlApp1.ActiveSheet.Range("G12").Delete
'
'        With xlSheet
'                .Name = "Stock dinámico"
'               .Rows("1:1").Select
'               .Rows("1:1").Insert Shift:=xlDown
'               .Range("A1").FormulaR1C1 = "STOCK VS DESPACHO AL " & Format(Date, "DD/MM/YYYY")
'               '.Range("A2").Select
'               '.Range("A2").Font.Bold = True
'               '.Range("A2").FormulaR1C1 = "al 03/09/07"
'               .Range("A1").Select
'               .Range("A1").Font.Bold = True
'               With .Range("A1").Font
'                   .Name = "Arial"
'                   .Size = 14
'                   .Strikethrough = False
'                   .Superscript = False
'                   .Subscript = False
'                   .OutlineFont = False
'                   .Shadow = False
'                   .Underline = xlUnderlineStyleNone
'                   .ColorIndex = xlAutomatic
'               End With
'
'        End With
'
''        .AddDataField xlApp1.ActiveSheet.PivotTables _
''            ("Tabla dinámica1").PivotFields("unidad"), "Cuenta de unidad", xlCount
''        .PivotSelect "", xlDataAndLabel, True
''        xlApp1.ActiveSheet.Range("A8").Select
''        .PivotFields("Cuenta de unidad"). _
''            Orientation = xlHidden
''        .AddDataField xlApp1.ActiveSheet.PivotTables _
''            ("Tabla dinámica1").PivotFields("cantidad"), "Suma de cantidad", xlSum
''        .AddDataField xlApp1.ActiveSheet.PivotTables _
''            ("Tabla dinámica1").PivotFields("despacho_prom"), "Suma de despacho_prom", _
''            xlSum
''        .AddDataField xlApp1.ActiveSheet.PivotTables _
''            ("Tabla dinámica1").PivotFields("ratio"), "Suma de ratio", xlSum
''        With .PivotFields( _
''            "string_dias_inmovil")
''            .Orientation = xlRowField
''            .Position = 3
''        End With
''        .AddDataField xlApp1.ActiveSheet.PivotTables _
''            ("Tabla dinámica1").PivotFields("string_dias_inmovil"), _
''            "Cuenta de string_dias_inmovil", xlCount
''
'      '  .Format xlTable2
'    End With
'End Sub

Public Property Let TipoPakingList(ByVal vNewValue As Variant)
xTipoRep = vNewValue
End Property

Public Property Let NombreCia(ByVal vNewValue As Variant)
xNomCia = vNewValue
End Property


Public Property Let Sqlquery(ByVal vNewValue As Variant)
xQuery = vNewValue
End Property


Public Sub Reporte_Trabajadores(ByVal pCia As String, ByVal pEstado As Integer)

    Dim Sql As String
    Dim sDir1 As String
    Dim Rq As ADODB.Recordset
    Sql = " usp_pla_consultar_Trabajadores '" & pCia & "'," & pEstado & ""
    If Not fAbrRst(Rq, Sql) Then
        MsgBox "No existen registros para el criterio especificado", vbExclamation, "Reporte Derechohabientes"
        GoTo Termina:
    End If
    Dim Columnas(21) As String
    Dim Ancho(21) As Double
    Columnas(0) = "CODIGO": Ancho(0) = 8
    Columnas(1) = "DNI": Ancho(1) = 8
    Columnas(2) = "TRABAJADOR": Ancho(2) = 25
    Columnas(3) = "APE.PATERNO": Ancho(3) = 10
    Columnas(4) = "APE.MATERNO": Ancho(4) = 10
    Columnas(5) = "NOMBRES": Ancho(5) = 15
    Columnas(6) = "VINCULO": Ancho(6) = 9
    Columnas(7) = "SEXO": Ancho(7) = 4
    Columnas(8) = "TIP. DOC.": Ancho(8) = 15
    Columnas(9) = "NRO.DOC": Ancho(9) = 10
    Columnas(10) = "FECHA NACIMIENTO": Ancho(10) = 10
    Columnas(11) = "FECHA ALTA": Ancho(11) = 10
    Columnas(12) = "FECHA BAJA": Ancho(12) = 10
    Columnas(13) = "SITUACION": Ancho(13) = 9
    Columnas(14) = "DIRECCION 1": Ancho(14) = 35
    Columnas(15) = "UBIGEO 1": Ancho(15) = 35
    Columnas(16) = "OBSV.": Ancho(16) = 9
    Columnas(17) = "FECHA CESE.": Ancho(17) = 9
    Columnas(18) = "TIPO DE DOC. QUE ACREDITA EL VINCULO": Ancho(18) = 35
    Columnas(19) = "NUMERO DE DOC. QUE ACREDITA EL VINCULO": Ancho(19) = 25
    Columnas(20) = "MES/AÑO CONCEPCION": Ancho(20) = 10
        
       Dim nFil As Integer
       nFil = 1
  With xlSheet
        .Cells(nFil, 1).Value = xNomCia
        .Cells(nFil, 1).Font.Size = 10
        
        
        nFil = nFil + 2
        .Cells(nFil, 1).Value = "REPORTE DE TRABAJADORES  "
        .Cells(nFil, 1).Font.Size = 10
        .Cells(nFil, 1).Font.Bold = True
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Merge
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).HorizontalAlignment = xlCenter
        
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Borders.LineStyle = xlContinuous
        
        nFil = nFil + 2
        
                
        Dim I As Integer
        nCol = 1
        For I = 1 To UBound(Columnas) + 1
            .Range(.Cells(nFil, nCol), .Cells(nFil, nCol)).ColumnWidth = Ancho(I - 1)
            .Cells(nFil, nCol).Value = Columnas(I - 1)
            .Cells(nFil, nCol).Font.Bold = True
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).Merge
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).HorizontalAlignment = xlCenter
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).VerticalAlignment = xlCenter
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).WrapText = True

            nCol = nCol + 1
        Next
        
        .Range(.Cells(nFil, 1), .Cells(nFil + 1, UBound(Columnas))).Borders.LineStyle = xlContinuous
        nFil = nFil + 1
        Dim J As Integer
        J = 0
        Dim F As Integer
        F = 0
        Dim xTip As String
        xTip = ""
      
     
        Dim xItem As Integer
        xItem = 0
        'With Rq
            If Rq.RecordCount > 0 Then
                Rq.MoveFirst
                Do While Not Rq.EOF
                        'Barra.Value = Rq.AbsolutePosition
                        nFil = nFil + 1
                       
                        xItem = xItem + 1
                         .Cells(nFil, 1).Value = "'" & Trim(Rq!PlaCod)
                         .Cells(nFil, 2).Value = "'" & Trim(Rq!nro_doc)
                         .Cells(nFil, 3).Value = Trim(Rq!nom_trab)
                         .Cells(nFil, 4).Value = Trim(Rq!ap_pat)
                         .Cells(nFil, 5).Value = Trim(Rq!ap_mat)
                         .Cells(nFil, 6).Value = Trim(Rq!nombres)
                         .Cells(nFil, 7).Value = Trim(Rq!vinculodh)
                         .Cells(nFil, 8).Value = Trim(Rq!sexo)
                         
                         .Cells(nFil, 9).Value = Trim(Rq!tipo_docdh)
                         .Cells(nFil, 10).Value = "'" & Trim(Rq!numero)
                                                                 
                         .Cells(nFil, 11).NumberFormat = "dd/mm/yyyy"
                         .Cells(nFil, 11).Value = CVDate(Rq!fec_nacdh)
                            
                         If Trim(Rq!ACTIVO_BAJA) = "ACTIVO" Then
                            .Cells(nFil, 12).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 12).Value = CVDate(Rq!fecha_alta)
                            .Cells(nFil, 13).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 13).Value = ""
                         Else
                            .Cells(nFil, 12).NumberFormat = "dd/mm/yyyy"
                         
                            If Rq!fecha_alta = Null Then
                               .Cells(nFil, 12).Value = CVDate("01/01/1900")
                            Else
                                .Cells(nFil, 12).Value = ""
                            End If
                            
                            .Cells(nFil, 13).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 13).Value = CVDate(Rq!fecha_baja)
                            
                         End If
                         
                         .Cells(nFil, 14).Value = "'" & Trim(Rq!ACTIVO_BAJA)
                         
                         sDir1 = ""
                         
                         sDir1 = Trim(Rq!dire) & ""
                         .Cells(nFil, 15).Value = "'" & sDir1
                         .Cells(nFil, 16).Value = "'" & Trim(Rq!nom_ubigeo)
                         
                         
                         If Not IsNull(Rq!fcese) Then
                            .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Interior.ColorIndex = 37
                            .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Interior.Pattern = xlSolid
                            .Cells(nFil, 17).Value = "'CESADO"
                            .Cells(nFil, 18).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 18).Value = CVDate(Rq!fcese)
                         End If
                         
                         .Cells(nFil, 19).Value = "'" & Trim(Rq!tipdoc_acredita_paternidad)
                         .Cells(nFil, 20).Value = "'" & Trim(Rq!nrodoc_acredita_paternidad)
                         If Trim(Rq!mes_concepcion & "") <> "" Then
                            .Cells(nFil, 21).Value = "'" & Format(Left(Trim(Rq!mes_concepcion), 2), "00") & "/" & Right(Trim(Rq!mes_concepcion), 4)
                         End If
                    
                    Rq.MoveNext
                Loop
            End If
                     
                            
            nFil = nFil + 1
            .Cells(nFil, 1).Value = "'TOTAL DTRABAJADORES: "
            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).Merge
            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).HorizontalAlignment = xlLeft
            .Cells(nFil, 1).Font.Bold = True

            .Cells(nFil, 4).FormulaR1C1 = xItem
            .Cells(nFil, 4).Font.Bold = True
'
            .Range(.Cells(nFil, 8), .Cells(nFil, 8)).NumberFormat = "#,###,##0.000;[Red](#,###,##0.000)"
            
            nFil = nFil + 2
         
         
         End With
        

Termina:
Rq.Close
Set Rq = Nothing

End Sub

Public Sub Reporte_DerechoHabientes(ByVal pCia As String, ByVal pDni As String, ByVal pCodTrab As String, _
ByVal pCodVinculo As String, ByVal pSituacion As String, ByVal pFecIni As String, ByVal pFecfin As String, ByVal pEstado As Integer)

Dim Sql As String
Dim sDir1 As String
Dim Rq As ADODB.Recordset
Sql = " usp_pla_consultar_derechohabientes '" & pCia & "','" & pCodTrab & "','" & pDni & "','" & pCodVinculo & "'," & pSituacion & "," & pFecIni & "," & pFecfin & "," & pEstado & ""
If Not fAbrRst(Rq, Sql) Then
    MsgBox "No existen registros para el criterio especificado", vbExclamation, "Reporte Derechohabientes"
    GoTo Termina:
End If

        Dim Columnas(23) As String
        Dim Ancho(23) As Double
        Columnas(0) = "CODIGO": Ancho(0) = 8
        Columnas(1) = "DNI": Ancho(1) = 8
        Columnas(2) = "TRABAJADOR": Ancho(2) = 25
        Columnas(3) = "APE.PATERNO": Ancho(3) = 10
        Columnas(4) = "APE.MATERNO": Ancho(4) = 10
        Columnas(5) = "NOMBRES": Ancho(5) = 15
        Columnas(6) = "VINCULO": Ancho(6) = 9
        Columnas(7) = "SEXO": Ancho(7) = 4
        Columnas(8) = "TIP. DOC.": Ancho(8) = 15
        Columnas(9) = "NRO.DOC": Ancho(9) = 10
        Columnas(10) = "FECHA NACIMIENTO": Ancho(10) = 10
        Columnas(11) = "FECHA ALTA": Ancho(11) = 10
        Columnas(12) = "FECHA BAJA": Ancho(12) = 10
        Columnas(13) = "SITUACION": Ancho(13) = 9
        Columnas(14) = "DIRECCION 1": Ancho(14) = 35
        Columnas(15) = "UBIGEO 1": Ancho(15) = 35
        Columnas(16) = "OBSV.": Ancho(16) = 9
        Columnas(17) = "FECHA CESE.": Ancho(17) = 9
        Columnas(18) = "TIPO DE DOC. QUE ACREDITA EL VINCULO": Ancho(18) = 35
        Columnas(19) = "NUMERO DE DOC. QUE ACREDITA EL VINCULO": Ancho(19) = 25
        Columnas(20) = "MES/AÑO CONCEPCION": Ancho(20) = 10
        Columnas(21) = "FECHA INGRESO TRAB": Ancho(20) = 10
        Columnas(22) = "FECHA CESE TRAB": Ancho(20) = 10
        
       Dim nFil As Integer
       nFil = 1
  With xlSheet
        .Cells(nFil, 1).Value = xNomCia
        .Cells(nFil, 1).Font.Size = 10
        
        
        nFil = nFil + 2
        .Cells(nFil, 1).Value = "REPORTE DERECHOHABIENTES  "
        .Cells(nFil, 1).Font.Size = 10
        .Cells(nFil, 1).Font.Bold = True
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Merge
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).HorizontalAlignment = xlCenter
        
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Borders.LineStyle = xlContinuous
        
        nFil = nFil + 2
        
                
        Dim I As Integer
        nCol = 1
        For I = 1 To UBound(Columnas) + 1
            .Range(.Cells(nFil, nCol), .Cells(nFil, nCol)).ColumnWidth = Ancho(I - 1)
            .Cells(nFil, nCol).Value = Columnas(I - 1)
            .Cells(nFil, nCol).Font.Bold = True
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).Merge
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).HorizontalAlignment = xlCenter
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).VerticalAlignment = xlCenter
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).WrapText = True
                       
            
            nCol = nCol + 1
        Next
        
        .Range(.Cells(nFil, 1), .Cells(nFil + 1, UBound(Columnas))).Borders.LineStyle = xlContinuous
        nFil = nFil + 1
        Dim J As Integer
        J = 0
        Dim F As Integer
        F = 0
        Dim xTip As String
        xTip = ""
      
     
        Dim xItem As Integer
        xItem = 0
        'With Rq
            If Rq.RecordCount > 0 Then
                Rq.MoveFirst
                Do While Not Rq.EOF
                        'Barra.Value = Rq.AbsolutePosition
                        nFil = nFil + 1
                       
                        xItem = xItem + 1
                         .Cells(nFil, 1).Value = "'" & Trim(Rq!PlaCod)
                         .Cells(nFil, 2).Value = "'" & Trim(Rq!nro_doc)
                         .Cells(nFil, 3).Value = Trim(Rq!nom_trab)
                         .Cells(nFil, 4).Value = Trim(Rq!ap_pat)
                         .Cells(nFil, 5).Value = Trim(Rq!ap_mat)
                         .Cells(nFil, 6).Value = Trim(Rq!nombres)
                         .Cells(nFil, 7).Value = Trim(Rq!vinculodh)
                         .Cells(nFil, 8).Value = Trim(Rq!sexo)
                         
                         .Cells(nFil, 9).Value = Trim(Rq!tipo_docdh)
                         .Cells(nFil, 10).Value = "'" & Trim(Rq!numero)
                                                                 
                         .Cells(nFil, 11).NumberFormat = "dd/mm/yyyy"
                         .Cells(nFil, 11).Value = CVDate(Rq!fec_nacdh)
                            
                         If Trim(Rq!ACTIVO_BAJA) = "ACTIVO" Then
                            .Cells(nFil, 12).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 12).Value = CVDate(Rq!fecha_alta)
                            .Cells(nFil, 13).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 13).Value = ""
                         Else
                            .Cells(nFil, 12).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 12).Value = CVDate(Rq!fecha_alta)
                            
                            .Cells(nFil, 13).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 13).Value = CVDate(Rq!fecha_baja)
                            
                         End If
                         
                         .Cells(nFil, 14).Value = "'" & Trim(Rq!ACTIVO_BAJA)
                         
                         sDir1 = ""
                         If Trim(Rq!NOM_VIA & "") <> "" Then sDir1 = sDir1 & Trim(Rq!des_via & "") & " " & Trim(Rq!NOM_VIA & "")
                         If Trim(Rq!NOM_VIA & "") <> "" Then sDir1 = sDir1 & " " & Trim(Rq!nro_via1 & "")
                         If Trim(Rq!NRO & "") <> "" Then sDir1 = sDir1 & " DPTO. " & Trim(Rq!NRO & "")
                         If Trim(Rq!Interior & "") <> "" Then sDir1 = sDir1 & " INT. " & Trim(Rq!Interior & "")
                         If Trim(Rq!nro_manzana1 & "") <> "" Then sDir1 = sDir1 & " MZ. " & Trim(Rq!nro_manzana1 & "")
                         If Trim(Rq!nro_lote1 & "") <> "" Then sDir1 = sDir1 & " LT. " & Trim(Rq!nro_lote1 & "")
                         If Trim(Rq!nro_kilometro1 & "") <> "" Then sDir1 = sDir1 & " KM. " & Trim(Rq!nro_kilometro1 & "")
                         If Trim(Rq!nro_block1 & "") <> "" Then sDir1 = sDir1 & " BLOCK. " & Trim(Rq!nro_block1 & "")
                         If Trim(Rq!nro_etapa1 & "") <> "" Then sDir1 = sDir1 & " ETAPA. " & Trim(Rq!nro_etapa1 & "")
                         
                         If Trim(Rq!NOM_ZONA & "") <> "" Then sDir1 = sDir1 & " " & Trim(Rq!des_ZONA & "") & " " & Trim(Rq!NOM_ZONA & "")
                         
                         .Cells(nFil, 15).Value = "'" & sDir1
                         .Cells(nFil, 16).Value = "'" & Trim(Rq!nom_ubigeo)
                         
                         
                         If Not IsNull(Rq!fcese) Then
                            .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Interior.ColorIndex = 37
                            .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Interior.Pattern = xlSolid
                            .Cells(nFil, 17).Value = "'CESADO"
                            .Cells(nFil, 18).NumberFormat = "dd/mm/yyyy"
                            .Cells(nFil, 18).Value = CVDate(Rq!fcese)
                         End If
                         
                         .Cells(nFil, 19).Value = "'" & Trim(Rq!tipdoc_acredita_paternidad)
                         .Cells(nFil, 20).Value = "'" & Trim(Rq!nrodoc_acredita_paternidad)
                         If Trim(Rq!mes_concepcion & "") <> "" Then
                            .Cells(nFil, 20).Value = "'" & Format(Left(Trim(Rq!mes_concepcion), 2), "00") & "/" & Right(Trim(Rq!mes_concepcion), 4)
                         End If
                        .Cells(nFil, 21).NumberFormat = "dd/mm/yyyy"
                        .Cells(nFil, 21).Value = "'" & Trim(Rq!FechaIngresoTrab)
                        .Cells(nFil, 22).NumberFormat = "dd/mm/yyyy"
                        .Cells(nFil, 22).Value = "'" & Trim(Rq!FechaCeseTrab)
                    Rq.MoveNext
                Loop
            End If
                     
                            
            nFil = nFil + 1
            .Cells(nFil, 1).Value = "'TOTAL DERECHOHABIENTES: "
            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).Merge
            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).HorizontalAlignment = xlLeft
            .Cells(nFil, 1).Font.Bold = True

            .Cells(nFil, 4).FormulaR1C1 = xItem
            .Cells(nFil, 4).Font.Bold = True
'
            .Range(.Cells(nFil, 8), .Cells(nFil, 8)).NumberFormat = "#,###,##0.000;[Red](#,###,##0.000)"
            
            nFil = nFil + 2
         
         
         End With
        

Termina:
Rq.Close
Set Rq = Nothing

End Sub
