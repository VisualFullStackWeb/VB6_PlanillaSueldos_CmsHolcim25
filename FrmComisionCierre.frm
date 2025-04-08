VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmComisionCierre 
   Caption         =   "Cierre de Comision Ventas"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Cierre de Comision Ventas"
   MDIChild        =   -1  'True
   ScaleHeight     =   2355
   ScaleWidth      =   4605
   Begin VB.CommandButton Cmdpar 
      Caption         =   "Listar parametros objetivo"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Txtano 
      Enabled         =   0   'False
      Height          =   285
      Left            =   735
      TabIndex        =   2
      Top             =   120
      Width           =   705
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "FrmComisionCierre.frx":0000
      Left            =   1740
      List            =   "FrmComisionCierre.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2595
   End
   Begin VB.CommandButton BtnCierre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   540
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   300
      Left            =   1455
      TabIndex        =   4
      Top             =   120
      Width           =   255
      Size            =   "450;529"
   End
   Begin VB.Label LblCierre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "FrmComisionCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SubTot
    ComisionesAcumuladas As Currency
    BasicoAsigFamiliar As Currency
    Factor_Ton_Alcance_Objetivo  As Currency
    Comision_Toneladas  As Currency
    Factor_Soles_Alcance_Objetivo    As Currency
    Comision_Soles  As Currency
    
    TonVendida As Currency
    TonObjetivo As Currency
    FactorTon As Currency
    SolesVendida As Currency
    SolesObjetivo As Currency
    FactorSoles As Currency
    ImpComisionInicial  As Currency
    ImpComisionFinal  As Currency
    
    PorComision  As Double
End Type

Private Sub Carga_Cierre()
Dim Rq As ADODB.Recordset
Sql$ = "SELECT Fec_Crea fROM Comi_Cierre where cia='" & wcia & "' and Ano=" & Txtano.Text & " and Mes=" & Cmbmes.ListIndex + 1 & " and Status<>'*'"
If fAbrRst(Rq, Sql) Then
   LblCierre.Caption = "Periodo Cerrado"
   BtnCierre.Caption = "Abrir Periodo"
   LblCierre.ForeColor = vbRed
Else
   LblCierre.Caption = "Periodo Abierto"
   BtnCierre.Caption = "Cerrar Periodo"
   LblCierre.ForeColor = vbBlue
End If
End Sub

Private Sub BtnCierre_Click()

If BtnCierre.Caption = "Abrir Periodo" Then
   Sql$ = "Update Comi_Cierre set status='*' where Cia='" & wcia & "' and ano=" & Txtano.Text & " and mes=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
   cn.Execute Sql, 64
   LblCierre.Caption = "Periodo Abierto"
   BtnCierre.Caption = "Cerrar Periodo"
   Sql$ = "delete from Vta_ComisionMovimientos where añoComision =" & Txtano.Text & " and MesComision='" & Cmbmes.ListIndex + 1 & "' and cia='" & wcia & "'"
   cn.Execute Sql, 64
   Sql$ = "delete from VTA_COMISIONVENDEDOR_MENSUAL where año =" & Txtano.Text & " and Mes='" & Cmbmes.ListIndex + 1 & "' and cia='" & wcia & "'"
   cn.Execute Sql, 64
   LblCierre.ForeColor = vbBlue
Else
   Sql$ = "Insert Into Comi_Cierre values ('" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & wuser & "',getdate(),'')"
   cn.Execute Sql, 64
   LblCierre.Caption = "Periodo Cerrado"
   BtnCierre.Caption = "Abrir Periodo"
   
   Sql$ = "Comision_Cierre '" & wcia & "','" & (Cmbmes.ListIndex + 1) & "'," & Txtano.Text
   cn.Execute Sql, 64
   Dim n_año As Integer
   Dim n_mes As Integer
   n_año = Val(Txtano.Text)
   n_mes = Cmbmes.ListIndex + 1
   If n_mes = 1 Then n_año = n_año - 1: n_mes = 12 'obs
   If n_mes <> 1 Then n_mes = n_mes - 1
   
   'Sql$ = "select Ficha ,Representante,sum(TotalToneladas) as Toneladas,sum(ValorVtaSoles) as Soles  from  Vta_ComisionMovimientos where añoComision =" & n_año & " and MesComision='" & n_mes & "' and Mes='" & n_mes & "' and cia='" & wcia & "' and Ficha in (select codigo from Vta_ComisionVendedor where Status<>'*' and codigo<>'E0540') group by Ficha,Representante union select 'E0540' as Ficha ,'RODRIGUEZ VALDERRAMA LUIS EDGARDO' as Trabajador,sum(TotalToneladas) as Toneladas,sum(ValorVtaSoles) as Soles  from  Vta_ComisionMovimientos where añoComision = " & n_año & " and MesComision='" & n_mes & "' and cia='" & wcia & "'"
   'add jcms 261121, a. jimenes comision total nacional indicado por L. Rodriguez
   ' and ficha in ('E0553','E0580')
   Sql$ = "select Ficha ,Representante,sum(TotalToneladas) as Toneladas,sum(ValorVtaSoles) as Soles  "
   Sql$ = Sql$ & " from  Vta_ComisionMovimientos where añoComision =" & n_año & " and MesComision='" & n_mes & "' and Mes='" & n_mes & "' and cia='" & wcia & "' "
   Sql$ = Sql$ & " and Ficha in (select codigo from Vta_ComisionVendedor where Status<>'*' "
   Sql$ = Sql$ & " and codigo<>'E0540' "
   Sql$ = Sql$ & " and codigo<>'E0547' " '--J. Mesias calculo manual GC"
   Sql$ = Sql$ & " and codigo<>'E0546' " ' --M. Sovia calculo manual GC"
   Sql$ = Sql$ & " and codigo<>'E0414' " '--A. Jimenez vta nacional ->  GC"
   Sql$ = Sql$ & " )   group by Ficha,Representante"
   Sql$ = Sql$ & " union select 'E0414' as Ficha ,'JIMENEZ CORONADO VIOLETA VIOLETA' as Trabajador,sum(TotalToneladas) as Toneladas,sum(ValorVtaSoles) as Soles  from  Vta_ComisionMovimientos where añoComision = " & n_año & " and MesComision='" & n_mes & "' and cia='" & wcia & "'  and cod_cli in (select cod_cli from clientes where status<>'*' and left(cod_ubi,3) in ('000','001'))"
   Sql$ = Sql$ & " union select 'E0540' as Ficha ,'RODRIGUEZ VALDERRAMA LUIS EDGARDO' as Trabajador,sum(TotalToneladas) as Toneladas,sum(ValorVtaSoles) as Soles  from  Vta_ComisionMovimientos where añoComision = " & n_año & " and MesComision='" & n_mes & "' and cia='" & wcia & "'"
   
   'add jcms 220323
   'Sql$ = " exec usp_vta_listar_calculo_comision_vta_mensual '" & wcia & "'," & n_año & "," & n_mes
   Sql$ = " exec usp_vta_listar_calculo_comision_vta_mensual_2024 '" & wcia & "'," & n_año & "," & n_mes
        
   Dim ld_ValTon As Double
   Dim ld_ValSol As Double

   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$, 64)
   If Not Rs.RecordCount > 0 Then MsgBox "No Existen Vendedores que Comisionan , No llegaron al Objetivo", vbCritical, TitMsg: Exit Sub
   Rs.MoveFirst
   Do While Not Rs.EOF
   
        ld_ValTon = 0
        ld_ValSol = 0
        If Txtano.Text = "2022" And Cmbmes.ListIndex + 1 = 2 And Trim(Rs!Ficha) = "E0515" Then 'add cjms 230222 excepcion indicada por LR Feb22
            ld_ValTon = 3160
            ld_ValSol = 717904
        Else
            ld_ValTon = Rs!TONELADAS
            ld_ValSol = Rs!SOLES
        End If
   
      'Sql$ = "listar_Factor_comision_mod_2018 '" & Cmbmes.ListIndex + 1 & "'," & Txtano.Text & ",'" & wcia & "','" & Trim(rs!Ficha) & "','" & Trim(rs!REPRESENTANTE) & "'," & Replace(Trim(rs!Toneladas), ",", "") & "," & Replace(Trim(rs!SOLES), ",", "") & "," & n_mes & "," & n_año
'      Sql$ = "listar_Factor_comision_mod_2018 '" & Cmbmes.ListIndex + 1 & "'," & Txtano.Text & ",'" & wcia & "','" & Trim(Rs!Ficha) & "','" & Trim(Rs!REPRESENTANTE) & "'," & Replace(Trim(ld_ValTon), ",", "") & "," & Replace(Trim(ld_ValSol), ",", "") & "," & n_mes & "," & n_año
'      Sql$ = Sql$ & "," & Replace(Rs!PrecioProm_VtaMesxComisionar, ",", "") & "," & Replace(Rs!Imp_Soles_Meta_Precio_Promedio_Por_Cartera, ",", "") & "," & Replace(Rs!Porc_Meta, ",", "") & "," & Replace(Rs!Porc_aplicable_a_comision_previa, ",", "")
'
     If Rs!Ficha <> "E0000" Then
      Sql$ = "listar_Factor_comision_mod_2024 '" & Cmbmes.ListIndex + 1 & "'," & Txtano.Text & ",'" & wcia & "','" & Trim(Rs!Ficha) & "','" & Trim(Rs!REPRESENTANTE) & "'," & Replace(Trim(ld_ValTon), ",", "") & "," & Replace(Trim(ld_ValSol), ",", "") & "," & n_mes & "," & n_año
      Sql$ = Sql$ & "," & Replace(Rs!PrecioProm_VtaMesxComisionar, ",", "") & "," & Replace(Rs!Imp_Soles_Meta_Precio_Promedio_Por_Cartera, ",", "") & "," & Replace(Rs!Porc_Meta, ",", "") & "," & Replace(Rs!Porc_aplicable_a_comision_previa, ",", "")
      cn.Execute Sql, 64
      Debug.Print Sql
    End If
      
      
'      If Trim(rs!Ficha) = "E0483" Then
'         Sql$ = "listar_Factor_comision_mod_2018 '" & Cmbmes.ListIndex + 1 & "'," & Txtano.Text & ",'" & wcia & "','" & Trim("E0553") & "','" & Trim("ANDRE DELGADO FOPPIANI") & "'," & Replace(Trim(rs!Toneladas), ",", "") & "," & Replace(Trim(rs!SOLES), ",", "") & "," & n_mes & "," & n_año
'         cn.Execute Sql, 64
'      End If

'      If Trim(Rs!Ficha) = "E0580" Then 'add jcms 250122 indicado por L.Rodriguez
'
'         Sql$ = "listar_Factor_comision_mod_2024 '" & Cmbmes.ListIndex + 1 & "'," & Txtano.Text & ",'" & wcia & "','" & Trim("E0590") & "','" & Trim("MACHUCA BARDALES CYNTHIA MAGALLY") & "'," & Replace(Trim(Rs!TONELADAS), ",", "") & "," & Replace(Trim(Rs!SOLES), ",", "") & "," & n_mes & "," & n_año
'         Sql$ = Sql$ & "," & Replace(Rs!PrecioProm_VtaMesxComisionar, ",", "") & "," & Replace(Rs!Imp_Soles_Meta_Precio_Promedio_Por_Cartera, ",", "") & "," & Replace(Rs!Porc_Meta, ",", "") & "," & Replace(Rs!Porc_aplicable_a_comision_previa, ",", "")
'
'         cn.Execute Sql, 64
'      End If
      
      Rs.MoveNext
   Loop
   If Rs.State = 1 Then Rs.Close
   
   LblCierre.ForeColor = vbRed
   
   'Call Exp_Excel
   Call Exp_Excel_2024
   
End If
End Sub
'Private Sub Exp_Excel()
'Dim Rs As ADODB.Recordset
'Dim nFil As Integer
'Dim Sql As String
'
'Sql = "usp_pla_listar_comisiones '" & wcia & "', " & CInt(Txtano.Text) & ", " & Cmbmes.ListIndex + 1 & ""
'If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub
'Set xlApp1 = CreateObject("Excel.Application")
'xlApp1.Workbooks.Add
'Set xlApp2 = xlApp1.Application
'Set xlBook = xlApp2.Workbooks(1)
'Set xlSheet = xlApp2.Worksheets("HOJA1")
'xlSheet.Name = "comisiones"
'xlApp2.Sheets("comisiones").Select
'
'
'If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub
'
''xlSheet.Range("A:A").ColumnWidth = 4
''xlSheet.Range("C:C").ColumnWidth = 35
'xlSheet.Range("D:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
'
'nFil = 1
'
'Dim nColFin As Integer
'nColFin = 18
'
'xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
'xlSheet.Cells(nFil, 2).Font.Bold = True
'nFil = nFil + 1
'xlSheet.Cells(nFil, 2).Value = "COMISIONES DEL MES  " & Cmbmes.Text & " " & CInt(Txtano.Text) & " ( CONSIDERANDO LAS VENTAS DEL MES ANTERIOR)"
''xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
'
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).HorizontalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).VerticalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Merge
'
'xlSheet.Cells(nFil, 2).Font.Bold = True
'nFil = nFil + 1
'
'xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
'xlSheet.Cells(nFil, 2).Font.Bold = True
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).HorizontalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).VerticalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Merge
'
'
'nFil = nFil + 2
'
'xlSheet.Cells(nFil, 2).Value = "AÑO"
'xlSheet.Cells(nFil, 2).ColumnWidth = 6
'xlSheet.Cells(nFil, 3).Value = "MES"
'xlSheet.Cells(nFil, 3).ColumnWidth = 6
'xlSheet.Cells(nFil, 4).Value = "CODIGO"
'xlSheet.Cells(nFil, 4).ColumnWidth = 8
'xlSheet.Cells(nFil, 5).Value = "REPRESENTANTE"
'xlSheet.Cells(nFil, 5).ColumnWidth = 45
'
'xlSheet.Cells(nFil, 6).Value = "TONELADA"
'xlSheet.Cells(nFil, 7).Value = "OBJETIVO TON."
'xlSheet.Cells(nFil, 8).Value = "% ALCANCE TON." ' "FACTORTON"
'xlSheet.Cells(nFil, 9).Value = "SOLES"
'xlSheet.Cells(nFil, 9).ColumnWidth = 15
'xlSheet.Cells(nFil, 10).Value = "OBJETIVO SOL."
'xlSheet.Cells(nFil, 10).ColumnWidth = 15
'
'xlSheet.Cells(nFil, 11).Value = "% ALCANCE SOL." ' "FACTORSOL"
'xlSheet.Cells(nFil, 12).Value = "VALOR ACTUAL COMISIÓN" '"VALORCOMISION"
'
''ADD JCMS 220323
'xlSheet.Cells(nFil, 13).Value = "PRECIO PROMEDIO SOLES/TM DE LA VENTA"
'xlSheet.Cells(nFil, 14).Value = "META DE PRECIO PROMEDIO POR CARTERA"
'xlSheet.Cells(nFil, 15).Value = UCase("(%) ALCANCE DE PRECIO PROMEDIO META")
'xlSheet.Cells(nFil, 16).Value = UCase("FACTOR APLICABLE A LA COMISIÓN PREVIA")
'xlSheet.Cells(nFil, 17).Value = "VALOR COMISIÓN TOTAL"
'xlSheet.Cells(nFil, 18).Value = "% COMISIÓN VS. SOLES (*)"
'
'
'
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).HorizontalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).VerticalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Borders.LineStyle = xlContinuous
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Font.Bold = True
'
''.HorizontalAlignment = xlCenter
''.VerticalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).WrapText = True
'        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.Pattern = xlSolid
'        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.PatternColorIndex = xlAutomatic
'        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.ThemeColor = xlThemeColorAccent5
'        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.TintAndShade = 0.799981688894314
'        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.PatternTintAndShade = 0
'
'nFil = 6
'
'Dim lcTot As SubTot
'
'lcTot.ImpComision = 0
'lcTot.PorComision = 0
'
'
'Dim xFilIni As Integer
'xFilIni = nFil
'
'If Rs.RecordCount > 0 Then Rs.MoveFirst
'Do While Not Rs.EOF
'
'    xlSheet.Cells(nFil, 2).Value = "'" & Trim(Rs!año & "")
'    xlSheet.Cells(nFil, 3).Value = "'" & Format(Trim(Rs!Mes & ""), "00")
'    xlSheet.Cells(nFil, 4).Value = Trim(Rs!Codigo & "")
'    xlSheet.Cells(nFil, 5).Value = Trim(Rs!REPRESENTANTE & "")
'    xlSheet.Cells(nFil, 6).Value = Trim(Rs!TONELADA & "")
'    xlSheet.Cells(nFil, 7).Value = Trim(Rs!Objetivo & "")
'    xlSheet.Cells(nFil, 8).Value = Trim(Rs!FactorTon & "")
'    xlSheet.Cells(nFil, 9).Value = Trim(Rs!SOLES & "")
'    xlSheet.Cells(nFil, 10).Value = Trim(Rs!OBJETIVOSOL & "")
'    xlSheet.Cells(nFil, 11).Value = Trim(Rs!FACTORSOL & "")
'    xlSheet.Cells(nFil, 12).Value = Trim(Rs!VALORCOMISION & "")
'    xlSheet.Cells(nFil, 13).Value = Trim(Rs!PrecioProm_VtaMesxComisionar & "")
'    xlSheet.Cells(nFil, 14).Value = Trim(Rs!Imp_Soles_Meta_Precio_Promedio_Por_Cartera & "")
'    xlSheet.Cells(nFil, 15).Value = Trim(Rs!Porc_Meta / 100 & "")
'    xlSheet.Cells(nFil, 15).NumberFormat = "0.00%"
'
'    xlSheet.Cells(nFil, 16).Value = Trim(Rs!Porc_aplicable_a_comision_previa / 100 & "")
'    xlSheet.Cells(nFil, 16).NumberFormat = "0.00%"
'
'    xlSheet.Cells(nFil, 17).Value = Trim(Rs!ImpSoles_ComisionVta_Total & "")
'
'    xlSheet.Cells(nFil, 18).Value = Trim(Rs!COMISION_VS_SOLES & "")
'    xlSheet.Cells(nFil, 18).NumberFormat = "0.000%"
'
'
'    lcTot.ImpComision = lcTot.ImpComision + Rs!ImpSoles_ComisionVta_Total
'    lcTot.PorComision = lcTot.PorComision + Rs!COMISION_VS_SOLES
'
'
'
'    If Trim(Rs!Codigo & "") = "E0540" Then
'        lcTot.TonVendida = Rs!TONELADA
'        lcTot.TonObjetivo = Rs!Objetivo
'        lcTot.FactorTon = Rs!FactorTon
'        lcTot.SolesVendida = Rs!SOLES
'        lcTot.SolesObjetivo = Rs!OBJETIVOSOL
'        lcTot.FactorSoles = Rs!FACTORSOL
'    End If
'
'
'    nFil = nFil + 1
'    msum = msum + 1
'
'   Rs.MoveNext
'Loop
'
'    'nFil = nFil + 1
'    xlSheet.Cells(nFil, 2).Value = "TOTALES"
'    xlSheet.Cells(nFil, 6).Value = lcTot.TonVendida
'    xlSheet.Cells(nFil, 7).Value = lcTot.TonObjetivo
'    xlSheet.Cells(nFil, 8).Value = lcTot.FactorTon
'    xlSheet.Cells(nFil, 9).Value = lcTot.SolesVendida
'    xlSheet.Cells(nFil, 10).Value = lcTot.SolesObjetivo
'    xlSheet.Cells(nFil, 11).Value = lcTot.FactorSoles
'
'    xlSheet.Cells(nFil, 17).Value = lcTot.ImpComision
'
'    'xlSheet.Cells(nFil, 18).Value = lcTot.ImpComision / lcTot.SolesVendida
'    'corrige error
'    If lcTot.SolesVendida <> 0 Then
'        xlSheet.Cells(nFil, 18).Value = lcTot.ImpComision / lcTot.SolesVendida
'    Else
'        xlSheet.Cells(nFil, 18).Value = 0
'    End If
'    xlSheet.Range(xlSheet.Cells(nFil, 18), xlSheet.Cells(nFil, 18)).NumberFormat = "##0.000%"
'
'    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).HorizontalAlignment = xlCenter
'    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).VerticalAlignment = xlCenter
'    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).Merge
'    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Font.Bold = True
'
'    xlSheet.Range(xlSheet.Cells(xFilIni, 2), xlSheet.Cells(nFil, nColFin)).Borders.LineStyle = xlContinuous
'
'nFil = nFil + 2
'
'xlSheet.Cells(nFil, 2).Value = "(*) INCLUYE COMISIONES DE VENTAS SEGUN PLAN 2023"
'nFil = nFil + 2
'xlSheet.Cells(nFil, 2).Value = "'" & Now
'
'
''NOTA
'Rs.Close: Set Rs = Nothing
'
'xlApp2.Application.ActiveWindow.DisplayGridlines = False
'
'xlApp2.Application.Visible = True
'
'If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
'If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
'If Not xlBook Is Nothing Then Set xlBook = Nothing
'If Not xlSheet Is Nothing Then Set xlSheet = Nothing
'
'
'Screen.MousePointer = vbDefault
'End Sub
Private Sub CmbMes_Click()
Carga_Cierre
End Sub

Private Sub Cmdpar_Click()
'FrmComisionCierre.Txtano

FrmSetComi.PeriodoAno = Me.Txtano
Load FrmSetComi
FrmSetComi.Show


End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 4845 ' 4635
Me.Height = 2940 ' 2610
Txtano.Text = Year(Date)
Cmbmes.ListIndex = Month(Date) - 1
End Sub

Private Sub SpinButton1_SpinDown()
Txtano.Text = Txtano.Text - 1
Carga_Cierre
End Sub

Private Sub SpinButton1_SpinUp()
Txtano.Text = Txtano.Text + 1
Carga_Cierre
End Sub




Private Sub Exp_Excel_2024()
Dim Rs As ADODB.Recordset
Dim nFil As Integer
Dim Sql As String

'Sql = "usp_pla_listar_comisiones '" & wcia & "', " & CInt(Txtano.Text) & ", " & Cmbmes.ListIndex + 1 & ""
Sql = "usp_pla_listar_comisiones_2024 '" & wcia & "', " & CInt(Txtano.Text) & ", " & Cmbmes.ListIndex + 1 & ""

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "comisiones"
xlApp2.Sheets("comisiones").Select


If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

'xlSheet.Range("A:A").ColumnWidth = 4
'xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

Dim nColFin As Integer
nColFin = 24 '18

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
xlSheet.Cells(nFil, 2).Value = "COMISIONES DEL MES  " & Cmbmes.Text & " " & CInt(Txtano.Text) & " ( CONSIDERANDO LAS VENTAS DEL MES ANTERIOR)"
'xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1

xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
xlSheet.Cells(nFil, 2).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Merge


nFil = nFil + 2
Dim nCol As Integer
nCol = 2

xlSheet.Cells(nFil, nCol).Value = "AÑO"
xlSheet.Cells(nFil, nCol).ColumnWidth = 6
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "MES"
xlSheet.Cells(nFil, nCol).ColumnWidth = 6
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "CODIGO"
xlSheet.Cells(nFil, nCol).ColumnWidth = 8
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "REPRESENTANTE"
xlSheet.Cells(nFil, nCol).ColumnWidth = 45

nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "COMISIONES ACUMULADAS"
xlSheet.Cells(nFil, nCol).ColumnWidth = 15

nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "BASICO+ASIGNACION FAMILIAR VENDEDOR"
xlSheet.Cells(nFil, nCol).ColumnWidth = 15


nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "TONELADA"
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "OBJETIVO TON."
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "% ALCANCE TON." ' "FACTORTON"
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "SOLES"
xlSheet.Cells(nFil, nCol).ColumnWidth = 15
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "OBJETIVO SOL."
xlSheet.Cells(nFil, nCol).ColumnWidth = 15
nCol = nCol + 1

xlSheet.Cells(nFil, nCol).Value = "% ALCANCE SOL." ' "FACTORSOL"

nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "FACTOR TN. (ALCANCE OBJETIVO)"

nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "COMISION EN TONELADAS."

nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "FACTOR SOLES. (ALCANCE OBJETIVO)"

nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "COMISION EN SOLES"


nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "VALOR ACTUAL COMISIÓN" '"VALORCOMISION"
nCol = nCol + 1

'ADD JCMS 220323
xlSheet.Cells(nFil, nCol).Value = "PRECIO PROMEDIO SOLES/TM DE LA VENTA"
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "META DE PRECIO PROMEDIO POR CARTERA"
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = UCase("(%) ALCANCE DE PRECIO PROMEDIO META")
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = UCase("FACTOR APLICABLE A LA COMISIÓN PREVIA")
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "VALOR COMISIÓN TOTAL"
nCol = nCol + 1
xlSheet.Cells(nFil, nCol).Value = "% COMISIÓN VS. SOLES (*)"


xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Font.Bold = True

'.HorizontalAlignment = xlCenter
'.VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).WrapText = True
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.Pattern = xlSolid
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.PatternColorIndex = xlAutomatic
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.ThemeColor = xlThemeColorAccent5
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.TintAndShade = 0.799981688894314
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Interior.PatternTintAndShade = 0

nFil = 6

Dim lcTot As SubTot


    lcTot.ComisionesAcumuladas = 0
    lcTot.BasicoAsigFamiliar = 0
    lcTot.Factor_Ton_Alcance_Objetivo = 0
    lcTot.Comision_Toneladas = 0
    lcTot.Factor_Soles_Alcance_Objetivo = 0
    lcTot.Comision_Soles = 0
    
    lcTot.TonVendida = 0
    lcTot.TonObjetivo = 0
    lcTot.FactorTon = 0
    lcTot.SolesVendida = 0
    lcTot.SolesObjetivo = 0
    lcTot.FactorSoles = 0
    lcTot.ImpComisionInicial = 0
    lcTot.ImpComisionFinal = 0
    lcTot.PorComision = 0
    
    
Dim li_ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM As Currency
li_ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM = 0


Dim xFilIni As Integer
xFilIni = nFil

nCol = 2

If Rs.RecordCount > 0 Then Rs.MoveFirst
Do While Not Rs.EOF

    If li_ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM = 0 Then li_ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM = Rs!ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM
    
    nCol = 2
    xlSheet.Cells(nFil, nCol).Value = "'" & Trim(Rs!año & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = "'" & Format(Trim(Rs!Mes & ""), "00")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!Codigo & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!REPRESENTANTE & "")
    
    
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!imp_comision_acumulada & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!imp_basico_asigfamiliar_periodo & "")
    
    lcTot.ComisionesAcumuladas = lcTot.ComisionesAcumuladas + Rs!imp_comision_acumulada
    lcTot.BasicoAsigFamiliar = lcTot.BasicoAsigFamiliar + Rs!imp_basico_asigfamiliar_periodo
    
    
   
    
    
    
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!TONELADA & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!Objetivo & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!FactorTon & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!SOLES & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!OBJETIVOSOL & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!FACTORSOL & "")
    
    
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!factor_ton_escala_objetivo_periodo & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = (Rs!factor_ton_escala_objetivo_periodo * Rs!imp_basico_asigfamiliar_periodo) / 100
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!factor_soles_escala_objetivo_periodo & "")
    nCol = nCol + 1
    'xlSheet.Cells(nFil, nCol).Value = (Rs!factor_soles_escala_objetivo_periodo * Rs!imp_basico_asigfamiliar_periodo) / 100
    
    xlSheet.Cells(nFil, nCol).Value = (Rs!factor_soles_escala_objetivo_periodo * Rs!imp_basico_asigfamiliar_periodo) / 100
    
    'lcTot.Factor_Ton_Alcance_Objetivo = 0
    lcTot.Comision_Toneladas = lcTot.Comision_Toneladas + (Rs!factor_ton_escala_objetivo_periodo * Rs!imp_basico_asigfamiliar_periodo) / 100
    'lcTot.Factor_Soles_Alcance_Objetivo = 0
    lcTot.Comision_Soles = lcTot.Comision_Soles + (Rs!factor_soles_escala_objetivo_periodo * Rs!imp_basico_asigfamiliar_periodo) / 100
    
    
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!VALORCOMISION & "")
    
    lcTot.ImpComisionInicial = lcTot.ImpComisionInicial + Rs!VALORCOMISION
    
    
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!PrecioProm_VtaMesxComisionar & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!Imp_Soles_Meta_Precio_Promedio_Por_Cartera & "")
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!Porc_Meta / 100 & "")
    xlSheet.Cells(nFil, nCol).NumberFormat = "0.00%"
    nCol = nCol + 1
    
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!Porc_aplicable_a_comision_previa / 100 & "")
    xlSheet.Cells(nFil, nCol).NumberFormat = "0.00%"
    nCol = nCol + 1
    
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!ImpSoles_ComisionVta_Total & "")
    
    nCol = nCol + 1
    xlSheet.Cells(nFil, nCol).Value = Trim(Rs!COMISION_VS_SOLES & "")
    xlSheet.Cells(nFil, nCol).NumberFormat = "0.000%"
    
    
    
    lcTot.ImpComisionFinal = lcTot.ImpComisionFinal + Rs!ImpSoles_ComisionVta_Total
    lcTot.PorComision = lcTot.PorComision + Rs!COMISION_VS_SOLES


    
    If Trim(Rs!Codigo & "") = "E0414" Or Trim(Rs!Codigo & "") = "E0638" Then
        lcTot.TonVendida = lcTot.TonVendida + Rs!TONELADA
        lcTot.TonObjetivo = lcTot.TonObjetivo + Rs!Objetivo
        lcTot.FactorTon = lcTot.FactorTon + Rs!FactorTon
        lcTot.SolesVendida = lcTot.SolesVendida + Rs!SOLES
        lcTot.SolesObjetivo = lcTot.SolesObjetivo + Rs!OBJETIVOSOL
        lcTot.FactorSoles = lcTot.FactorSoles + Rs!FACTORSOL
    End If
    
                                        
    nFil = nFil + 1
    msum = msum + 1

   Rs.MoveNext
Loop

    'nFil = nFil + 1
    xlSheet.Cells(nFil, 2).Value = "TOTALES"
    
    xlSheet.Cells(nFil, 6).Value = lcTot.ComisionesAcumuladas
    xlSheet.Cells(nFil, 7).Value = lcTot.BasicoAsigFamiliar
    xlSheet.Cells(nFil, 8).Value = lcTot.TonVendida
    xlSheet.Cells(nFil, 9).Value = lcTot.TonObjetivo
    xlSheet.Cells(nFil, 10).Value = (lcTot.TonVendida / lcTot.TonObjetivo) * 100
    
    
    xlSheet.Cells(nFil, 11).Value = lcTot.SolesVendida
    xlSheet.Cells(nFil, 12).Value = lcTot.SolesObjetivo
    xlSheet.Cells(nFil, 13).Value = (lcTot.SolesVendida / lcTot.SolesObjetivo) * 100
    
        
    xlSheet.Cells(nFil, 15).Value = lcTot.Comision_Toneladas
    xlSheet.Cells(nFil, 17).Value = lcTot.Comision_Soles
    xlSheet.Cells(nFil, 18).Value = lcTot.ImpComisionInicial
    
    xlSheet.Cells(nFil, 19).Value = lcTot.SolesVendida / lcTot.TonVendida
    
    xlSheet.Cells(nFil, 20).Value = li_ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM 'valor asignado al Gerente de ventas indicado x MP
    
    If li_ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM <> 0 Then
        xlSheet.Cells(nFil, 21).Value = ((lcTot.SolesVendida / lcTot.TonVendida) / li_ImpTotal_OBJETIVO_PRECIO_PROMEDIO_SOLES_TM) * 100
    Else
        xlSheet.Cells(nFil, 21).Value = 0
    End If
    
    xlSheet.Cells(nFil, 23).Value = lcTot.ImpComisionFinal
    
    xlSheet.Cells(nFil, 24).Value = (lcTot.ImpComisionFinal / lcTot.SolesVendida) '* 100
    xlSheet.Range(xlSheet.Cells(nFil, 24), xlSheet.Cells(nFil, 24)).NumberFormat = "##0.000%"
    
    
    'xlSheet.Cells(nFil, 9).Value = lcTot.TonObjetivo
    'xlSheet.Cells(nFil, 10).Value = lcTot.FactorTon
    'xlSheet.Cells(nFil, 11).Value = lcTot.SolesVendida
    'xlSheet.Cells(nFil, 12).Value = lcTot.SolesObjetivo
    'xlSheet.Cells(nFil, 13).Value = lcTot.FactorSoles
    
    
    
    'xlSheet.Cells(nFil, 18).Value = lcTot.ImpComision / lcTot.SolesVendida
    'corrige error
'    If lcTot.SolesVendida <> 0 Then
'        xlSheet.Cells(nFil, 20).Value = lcTot.ImpComision / lcTot.SolesVendida
'    Else
'        xlSheet.Cells(nFil, 20).Value = 0
'    End If
'    xlSheet.Range(xlSheet.Cells(nFil, 20), xlSheet.Cells(nFil, 18)).NumberFormat = "##0.000%"
    
    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).HorizontalAlignment = xlCenter
    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).VerticalAlignment = xlCenter
    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).Merge
    xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, nColFin)).Font.Bold = True
    
    xlSheet.Range(xlSheet.Cells(xFilIni, 2), xlSheet.Cells(nFil, nColFin)).Borders.LineStyle = xlContinuous

nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "(*) INCLUYE COMISIONES DE VENTAS SEGUN PLAN 2024"
nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "'" & Now


'NOTA
Rs.Close: Set Rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridlines = False

xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub



