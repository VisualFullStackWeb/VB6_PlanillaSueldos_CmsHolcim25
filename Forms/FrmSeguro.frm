VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSeguro 
   Caption         =   "Seguros del Mes"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleMode       =   0  'User
   ScaleWidth      =   8425.984
   Begin VB.ComboBox Cbo 
      Height          =   315
      ItemData        =   "FrmSeguro.frx":0000
      Left            =   2025
      List            =   "FrmSeguro.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2385
   End
   Begin VB.TextBox txt_Year 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6641
      _Version        =   393216
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton CmdProcesa 
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   0
      Width           =   1185
      Caption         =   "Calculo"
      PicturePosition =   327683
      Size            =   "2090;873"
      Picture         =   "FrmSeguro.frx":0090
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Index           =   2
      Left            =   2130
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1905
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   645
   End
   Begin MSForms.SpinButton sp_Year 
      Height          =   315
      Left            =   1695
      TabIndex        =   4
      Top             =   120
      Width           =   255
      Size            =   "450;556"
   End
   Begin MSForms.CommandButton btn_Load 
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Width           =   1185
      Caption         =   "     Excel"
      PicturePosition =   327683
      Size            =   "2090;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FrmSeguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsTemporal  As ADODB.Recordset
Dim Cadena As String
Dim mMonth As String
Dim Cn_dbf As ADODB.Connection
Dim nFil As Integer
Dim nCol As Integer
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim xlApp  As Object
Dim Reg As Boolean
Dim mFactorEmp As Double
Dim mFactorObr As Double


Private Sub btn_Load_Click()
If MsgBox("Desea Generar Excel ?", vbQuestion + vbYesNo, "Sistema") = vbYes Then

    If Not IsNumeric(lbl(2).Caption) Or lbl(2).Caption * 1 < 1 Then
       MsgBox "No hay Informacion Generada... Verifique", vbInformation
       Exit Sub
    End If
    
    rsTemporal.MoveFirst
    'rsTemporal!tipoCambio
    'Do Until rsTemporal.EOF
    '   MsgBox "" & rsTemporal!tipoCambio
    '   rsTemporal.MoveNext
    'Loop
    If Not IsNumeric(rsTemporal!tipoCambio) Then
       MsgBox "No existe tipo de cambio... Verifique", vbInformation
       Exit Sub
    End If
    Call Menos_3("01", rsTemporal!tipoCambio)
    Call Menos_3("02", rsTemporal!tipoCambio)
    
    Detalle_seg_Vida ("01")
    Detalle_seg_Vida ("02")
    Formato_seg_Vida
    Formato_seg_VidaII
    
    Dim x As Boolean
    x = Exportar_Excel2
End If
End Sub

Private Sub Cbo_Click()
    mMonth = Format(Cbo.ListIndex + 1, "00")
    CargaData
End Sub
Sub CargaData()
    Cadena = "sp_Trae_Pla_Seguro_Vida '" & wcia & "'," & Txt_Year.Text & "," & Cbo.ListIndex + 1 & ""
    Set rsTemporal = OpenRecordset(Cadena, cn)
    Set dg.DataSource = rsTemporal
    With dg
        .Columns(0).Caption = "Codigo"
        .Columns(1).Caption = "Nombre"
        .Columns(2).Caption = "Tipo"
        .Columns(3).Caption = "F.Ingreso"
        .Columns(4).Caption = "Total"
        .Columns(5).Caption = "Año"
        .Columns(6).Caption = "Mes"
        .Columns(7).Caption = "Factor"
        .Columns(8).Caption = "Seguro"
        .Columns(9).Caption = "Tc"
        .Columns(0).Width = 600
        .Columns(1).Width = 2800
        .Columns(2).Width = 400
        .Columns(3).Width = 1000
        .Columns(4).Width = 800
        .Columns(5).Width = 400
        .Columns(6).Width = 400
        .Columns(7).Width = 400
        .Columns(8).Width = 400
        .Columns(9).Width = 400
        .Columns(4).Alignment = dbgRight
        .Columns(7).Alignment = dbgRight
        .Columns(8).Alignment = dbgRight
        .Columns(9).Alignment = dbgRight
        .Columns(4).NumberFormat = "####,##0.00"
        .Columns(7).NumberFormat = "#,##0.00"
        .Columns(8).NumberFormat = "#,##0.00"
        .Columns(9).NumberFormat = "#,##0.00"
    End With
    lbl(2).Caption = rsTemporal.RecordCount
    Reg = IIf(dg.Row >= 0, True, False)
End Sub
Private Sub cmdprocesa_Click()
If Reg = True Then
   If MsgBox("Desea Eliminar el calculo Previo", vbYesNo + vbQuestion) = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
    Elimina
    dg.Refresh
    Procesa
Else
    Procesa
End If
CargaData
dg.Refresh
End Sub

Private Sub Form_Load()
 Call Init_Form
 mFactorEmp = 0.45
 mFactorObr = 0.56
End Sub
Sub Elimina()
    Sql = "sp_Elimina_Pla_Seguro_Vida '" & wcia & "'," & Txt_Year.Text & "," & Cbo.ListIndex + 1 & ",'" & wuser & "'" & ""
    cn.Execute Sql
    wuser = ""
End Sub
Sub Procesa()
    Dim FecFin As String
    Dim FecIni As String
    Dim mtc As Double
    mtc = InputBox("Ingrese Tipo de Cambio", "Cuadro IV")
    
    If Not IsNumeric(mtc) Then
       MsgBox "Ingrese Correctamente el tipo de cambio", vbInformation
       Exit Sub
    End If
    
    FecFin = Format(Format(Ultimo_Dia(Cbo.ListIndex + 1, Val(Txt_Year.Text)), "00") & "/" & Format(Cbo.ListIndex + 1, "00") & "/" & Txt_Year.Text, "DD/MM/YYYY")
    FecIni = DateAdd("m", -6, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))
    
    Dim rs As ADODB.Recordset
    Dim nFil As Integer
       
    Sql = "usp_Pla_Seguro_Vida '" & wcia & "','01','" & FecIni & "','" & FecFin & "'," & mtc & "," & mFactorEmp & "," & mtc & ",'" & wuser & "'" & ""
    cn.Execute Sql
    Sql = "usp_Pla_Seguro_Vida '" & wcia & "','02','" & FecIni & "','" & FecFin & "', " & mtc & "," & mFactorObr & "," & mtc & ",'" & wuser & "'" & ""
    cn.Execute Sql

End Sub
Private Sub Init_Form()
    Txt_Year.Text = Year(Now)
    Me.Top = 50
    Me.Left = 50
    Me.Height = 5300
    Me.Width = 12420
End Sub

Private Sub Menos_3(lTipoTrab As String, mtc As Double)

Dim Sql As String

Dim mFactorE As Double
Dim mIgv As Double


If lTipoTrab = "01" Then mFactorE = mFactorEmp Else mFactorE = mFactorObr
mIgv = 18
Sql = "Select S.*,fnacimiento,"
Sql = Sql & "Case tipo_doc when '01' then nro_doc else '' End as Dni,"
Sql = Sql & "(select cuenta from maestros_31 where ciamaestro='01055' and cod_maestro3=p.cargo) as Cargo,"
Sql = Sql & "(select descripcion from pla_ccostos where cia='01' and codigo=s.area and status<>'*') as CCosto "
Sql = Sql & "From Pla_Calculo_Seguro_Vida S,Planillas P "
Sql = Sql & "Where s.status<>'*' and s.TipoTrab='" & lTipoTrab & "' and p.cia='01' and p.status<>'*' and p.placod=s.placod and S.año= " & Txt_Year.Text & " and S.Mes = " & Cbo.ListIndex + 1 & " order by s.placod"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

Dim lTotPlanilla As Double
Dim lTotPasaTope As Double
Dim lTotMenos As Double
Dim lNumTrabPasaTope As Integer

lTotPlanilla = 0: lTotPasaTope = 0: lTotMenos = 0: lNumTrabPasaTope = 0
Do While Not rs.EOF
   lTotPlanilla = lTotPlanilla + rs!Total
   If Trim(rs!tope & "") = "S" Then lTotPasaTope = lTotPasaTope + rs!Total: lNumTrabPasaTope = lNumTrabPasaTope + 1
   If Trim(rs!Pasa & "") <> "S" Then lTotMenos = lTotMenos + rs!Total
   rs.MoveNext
Loop

If lTipoTrab = "02" Then
   Set xlSheet = xlApp2.Worksheets("HOJA2")
   xlSheet.Name = "SEGO"
Else
   Set xlApp1 = CreateObject("Excel.Application")
   xlApp1.Workbooks.Add
   Set xlApp2 = xlApp1.Application
   xlApp2.Sheets.Add
   xlApp2.Sheets.Add
   xlApp2.Sheets.Add
   xlApp2.Sheets.Add
   xlApp2.Sheets.Add
    
   xlApp2.Sheets("Hoja1").Select
   xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
   xlApp2.Sheets("Hoja2").Select
   xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)
   xlApp2.Sheets("Hoja3").Select
   xlApp2.Sheets("Hoja3").Move Before:=xlApp2.Sheets(3)
   xlApp2.Sheets("Hoja4").Select
   xlApp2.Sheets("Hoja4").Move Before:=xlApp2.Sheets(4)
   xlApp2.Sheets("Hoja5").Select
   xlApp2.Sheets("Hoja5").Move Before:=xlApp2.Sheets(5)
   xlApp2.Sheets("Hoja6").Select
   xlApp2.Sheets("Hoja6").Move Before:=xlApp2.Sheets(5)
   xlApp2.Sheets("Hoja7").Select
   xlApp2.Sheets("Hoja7").Move Before:=xlApp2.Sheets(5)

   Set xlBook = xlApp2.Workbooks(1)
   Set xlSheet = xlApp2.Worksheets("HOJA1")
   xlSheet.Name = "SEGE"
End If

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 7
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:D").ColumnWidth = 11


xlSheet.Cells(1, 1).Value = Trae_CIA(wcia)
xlSheet.Cells(2, 1).Value = "CALCULO DE SEGURO DE VIDA"
xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 8)).Merge
xlSheet.Range(xlSheet.Cells(2, 1), xlSheet.Cells(2, 8)).HorizontalAlignment = xlCenter

If lTipoTrab = "01" Then
   xlSheet.Cells(4, 2).Value = "TOTAL PLANILLA SUELDOS"
Else
   xlSheet.Cells(4, 2).Value = "TOTAL PLANILLA SALARIOS"
End If
xlSheet.Cells(4, 5).Value = lTotPlanilla
xlSheet.Cells(4, 5).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(6, 2).Value = "PERSONAL CON MENOS DE 3 MESES"

xlSheet.Cells(8, 2).Value = "CODIGO"
xlSheet.Cells(8, 3).Value = "NOMBRE"
xlSheet.Cells(8, 4).Value = "DNI"
xlSheet.Cells(8, 5).Value = "FEC. NAC."
xlSheet.Cells(8, 6).Value = "FEC. ING."
xlSheet.Cells(8, 7).Value = "OCUPACION"
xlSheet.Cells(8, 8).Value = "SUELDO"

xlSheet.Range(xlSheet.Cells(8, 2), xlSheet.Cells(8, 8)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(8, 2), xlSheet.Cells(8, 8)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(8, 2), xlSheet.Cells(8, 8)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(8, 2), xlSheet.Cells(8, 8)).Font.Bold = True

nFil = 9
Dim msum As Integer
Dim mTotTrab As Integer
msum = 1
mTotTrab = 0
Do While Not rs.EOF
   If Trim(rs!Pasa & "") <> "S" Then
      xlSheet.Cells(nFil, 1).Value = msum
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = "'" & Trim(rs!DNI & "")
      xlSheet.Cells(nFil, 5).Value = rs!fnacimiento
      xlSheet.Cells(nFil, 5).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 6).Value = rs!fIngreso
      xlSheet.Cells(nFil, 6).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 7).Value = Trim(rs!Cargo & "")
      xlSheet.Cells(nFil, 8).Value = rs!Total * -1
      xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
      xlSheet.Cells(nFil, 9).Value = Trim(rs!ccosto & "")
      nFil = nFil + 1
      msum = msum + 1
   End If
   mTotTrab = mTotTrab + 1
   rs.MoveNext
Loop
msum = msum * -1
xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Cells(nFil, 8).Borders.LineStyle = xlContinuous

xlSheet.Cells(nFil + 2, 8).Value = lTotPlanilla - lTotMenos
xlSheet.Cells(nFil + 2, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Cells(nFil + 2, 8).Borders.LineStyle = xlContinuous

nFil = nFil + 4

xlSheet.Cells(nFil, 6).Value = mFactorE
xlSheet.Cells(nFil, 7).Value = "%"
xlSheet.Cells(nFil, 8).Value = Round((lTotPlanilla - lTotMenos) * (mFactorE / 100), 2)
xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(nFil + 1, 6).Value = mIgv
xlSheet.Cells(nFil + 1, 7).Value = "%"
xlSheet.Cells(nFil + 1, 8).Value = Round((Round((lTotPlanilla - lTotMenos) * (mFactorE / 100), 2)) * (mIgv / 100), 2)
xlSheet.Cells(nFil + 1, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(nFil + 2, 7).Value = "TOTAL"
xlSheet.Cells(nFil + 2, 8).Value = (Round((lTotPlanilla - lTotMenos) * (mFactorE / 100), 2)) + (Round((Round((lTotPlanilla - lTotMenos) * (mFactorE / 100), 2)) * (mIgv / 100), 2))

If lTipoTrab = "01" Then
   xlSheet.Cells(nFil + 4, 6).Value = "TOTAL EMPLEADOS"
Else
   xlSheet.Cells(nFil + 4, 6).Value = "TOTAL OBREROS"
End If
xlSheet.Cells(nFil + 4, 8).Value = mTotTrab

nFil = nFil + 8
xlSheet.Cells(nFil, 2).Value = "Trabajadores Con Remuneracion  Mayor a "
xlSheet.Cells(nFil, 4).Value = 3125
xlSheet.Cells(nFil, 4).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Cells(nFil, 5).Value = "US$"
xlSheet.Cells(nFil, 6).Value = mtc
xlSheet.Cells(nFil, 7).Value = "S/."
xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
xlSheet.Cells(nFil, 8).Value = Round(3125 * mtc, 2)
xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "Numero de Trabajadores"
xlSheet.Cells(nFil, 4).Value = lNumTrabPasaTope
xlSheet.Cells(nFil, 7).Value = "S/."
xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
xlSheet.Cells(nFil, 8).Value = lTotPasaTope
xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

nFil = nFil + 2
msum = 1
If rs.RecordCount > 1 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tope & "") = "S" Then
      xlSheet.Cells(nFil, 1).Value = msum
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = "'" & Trim(rs!DNI & "")
      xlSheet.Cells(nFil, 5).Value = rs!fnacimiento
      xlSheet.Cells(nFil, 5).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 6).Value = rs!fIngreso
      xlSheet.Cells(nFil, 6).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 7).Value = Trim(rs!Cargo & "")
      xlSheet.Cells(nFil, 8).Value = rs!Total * -1
      xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
      xlSheet.Cells(nFil, 9).Value = Trim(rs!ccosto & "")
      nFil = nFil + 1
      msum = msum + 1
   End If
   mTotTrab = mTotTrab + 1
   rs.MoveNext
Loop

rs.Close: Set rs = Nothing

Dim lFecha As String
lFecha = Format(Format(Ultimo_Dia(Cbo.ListIndex + 1, Val(Txt_Year.Text)), "00") & "/" & Format(Cbo.ListIndex + 1, "00") & "/" & Txt_Year.Text, "DD/MM/YYYY")

Sql = "usp_Pla_Seguro_Vida_Cuadra '" & wcia & "','" & lTipoTrab & "','" & lFecha & "'"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst
nFil = nFil + 2
Do While Not rs.EOF
   xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
   xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
   xlSheet.Cells(nFil, 4).Value = "F. Ingreso"
   xlSheet.Cells(nFil, 5).Value = rs!fIngreso
   xlSheet.Cells(nFil, 5).NumberFormat = "dd/mm/yyyy;@"
   xlSheet.Cells(nFil, 6).Value = rs!fIngreso
   xlSheet.Cells(nFil, 7).Value = rs!fcese
   xlSheet.Cells(nFil, 7).NumberFormat = "dd/mm/yyyy;@"
   Select Case Trim(rs!tipo & "")
          Case "A": xlSheet.Cells(nFil, 8).Value = "Trabajador cesado en periodo anterior"
          Case "C": xlSheet.Cells(nFil, 8).Value = "Trabajador cesado en periodo actual"
          Case "S": xlSheet.Cells(nFil, 8).Value = "Trabajador sin boleta en periodo actual"
   End Select
   nFil = nFil + 1
   rs.MoveNext
Loop
rs.Close: Set rs = Nothing
Screen.MousePointer = vbDefault

End Sub

Private Sub Detalle_seg_Vida(lTipoTrab As String)

Dim Sql As String

Sql = "Select placod,nombre,(select descripcion from pla_ccostos where cia='01' and codigo=Pla_Calculo_Seguro_Vida.area and status<>'*') as CCosto,"
Sql = Sql & "Basico , PromExt, PromOtros, Total "
Sql = Sql & "From Pla_Calculo_Seguro_Vida where status<>'*'  and tipotrab='" & lTipoTrab & "' and año= " & Txt_Year.Text & " and Mes = " & Cbo.ListIndex + 1 & " order by ccosto,placod"


If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

Dim lTotBasico As Double
Dim lTotExtras As Double
Dim lTotOtros As Double
Dim lTotTot As Double

lTotBasico = 0: lTotExtras = 0: lTotOtros = 0: lTotTot = 0

If lTipoTrab = "02" Then
   Set xlSheet = xlApp2.Worksheets("HOJA4")
   xlSheet.Name = "DETOBR"
Else
   Set xlSheet = xlApp2.Worksheets("HOJA3")
   xlSheet.Name = "DETEMP"
End If

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 7
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:D").ColumnWidth = 50
xlSheet.Range("E:H").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(2, 2).Value = "CALCULO DE SEGURO DE VIDA"


If lTipoTrab = "01" Then
   xlSheet.Cells(3, 2).Value = "TOTAL PLANILLA SUELDOS MES DE " & Cbo.Text & " - " & Txt_Year.Text
Else
   xlSheet.Cells(3, 2).Value = "TOTAL PLANILLA SALARIOS MES DE " & Cbo.Text & " - " & Txt_Year.Text
End If
xlSheet.Range(xlSheet.Cells(1, 2), xlSheet.Cells(3, 2)).Font.Bold = True

xlSheet.Cells(5, 2).Value = "CODIGO"
xlSheet.Cells(5, 3).Value = "NOMBRE"
xlSheet.Cells(5, 4).Value = "CENTRO DE COSTO"
xlSheet.Cells(5, 5).Value = "SUELDO"
xlSheet.Cells(5, 6).Value = "PROM.EXT"
xlSheet.Cells(5, 7).Value = "PROM.OTROS"
xlSheet.Cells(5, 8).Value = "TOTAL"

xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 8)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 8)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 8)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(5, 2), xlSheet.Cells(5, 8)).Font.Bold = True

nFil = 6
Dim msum As Integer
Dim mTotTrab As Integer
msum = 1
mTotTrab = 0
Dim mCCosto As String
mCCosto = Trim(rs!ccosto & "")
Do While Not rs.EOF
   If rs!Total <> 0 Then
      If mCCosto <> "" And mCCosto <> Trim(rs!ccosto & "") Then
         msum = msum * -1
         nFil = nFil + 1
         xlSheet.Cells(nFil, 4).Value = "TOTAL " & mCCosto
         xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 8)).Borders.LineStyle = xlContinuous
         msum = 1: nFil = nFil + 2
         mCCosto = Trim(rs!ccosto & "")
      End If
      
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = "'" & Trim(rs!ccosto & "")
      xlSheet.Cells(nFil, 5).Value = rs!Basico
      xlSheet.Cells(nFil, 6).Value = rs!PROMEXT
      xlSheet.Cells(nFil, 7).Value = rs!PromOtros
      xlSheet.Cells(nFil, 8).Value = rs!Total
      lTotBasico = lTotBasico + rs!Total
      lTotExtras = lTotExtras + rs!Total
      lTotOtros = lTotOtros + rs!Total
      lTotTot = lTotTot + rs!Total
      nFil = nFil + 1
      msum = msum + 1
   End If
   mTotTrab = mTotTrab + 1
   rs.MoveNext
Loop

msum = msum * -1
nFil = nFil + 1
xlSheet.Cells(nFil, 4).Value = "TOTAL " & mCCosto
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 8)).Borders.LineStyle = xlContinuous
nFil = nFil + 2

xlSheet.Cells(nFil, 4).Value = "TOTAL GENERAL"
xlSheet.Cells(nFil, 5).Value = lTotBasico
xlSheet.Cells(nFil, 6).Value = lTotExtras
xlSheet.Cells(nFil, 7).Value = lTotOtros
xlSheet.Cells(nFil, 8).Value = lTotTot
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 8)).Borders.LineStyle = xlContinuous

Screen.MousePointer = vbDefault

End Sub
Private Sub Formato_seg_Vida()

Dim RqF As ADODB.Recordset
Sql = "usp_pla_Seguro_vida_Formato2 " & Txt_Year.Text & "," & Cbo.ListIndex + 1 & ""
If (fAbrRst(RqF, Sql)) Then RqF.MoveFirst Else RqF.Close: Set RqF = Nothing: Exit Sub

Set xlSheet = xlApp2.Worksheets("HOJA5")
xlSheet.Name = "FORM"
   
xlSheet.Range("Q:Q").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Range("G:G").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("R:R").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("S:S").NumberFormat = "dd/mm/yyyy;@"



xlSheet.Cells(1, 1).Value = "tipo Documento"
xlSheet.Cells(1, 2).Value = "Documento de Identidad"
xlSheet.Cells(1, 3).Value = "Apellido Paterno"
xlSheet.Cells(1, 4).Value = "Apellido Materno"
xlSheet.Cells(1, 5).Value = "Primer nombre"
xlSheet.Cells(1, 6).Value = "Segundo nombre"
xlSheet.Cells(1, 7).Value = "fecha Nacimiento"
xlSheet.Cells(1, 8).Value = "Sexo"
xlSheet.Cells(1, 9).Value = "Nacionalidad"
xlSheet.Cells(1, 10).Value = "Ocupacion"
xlSheet.Cells(1, 11).Value = "Departamento"
xlSheet.Cells(1, 12).Value = "Provincia"
xlSheet.Cells(1, 13).Value = "Distrito"
xlSheet.Cells(1, 14).Value = "Direccion"
xlSheet.Cells(1, 15).Value = "Tipo de Trabajador"
xlSheet.Cells(1, 16).Value = "moneda Sueldo"
xlSheet.Cells(1, 17).Value = "importe Sueldo"
xlSheet.Cells(1, 18).Value = "Fecha Ing Seguro"
xlSheet.Cells(1, 19).Value = "Fecha Ing Empresa"
xlSheet.Cells(1, 20).Value = "Telefono"
xlSheet.Cells(1, 21).Value = "edad"


xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 30)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 30)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 30)).Borders.LineStyle = xlContinuous

nFil = 2

Do While Not RqF.EOF
   xlSheet.Cells(nFil, 1).Value = Trim(RqF!tipo_doc & "")
   xlSheet.Cells(nFil, 2).Value = "'" & Trim(RqF!DniTrab & "")
   xlSheet.Cells(nFil, 3).Value = Trim(RqF!Ape_pat & "")
   xlSheet.Cells(nFil, 4).Value = Trim(RqF!Ape_mat & "")
   xlSheet.Cells(nFil, 5).Value = Trim(RqF!Pri_Nombre & "")
   xlSheet.Cells(nFil, 6).Value = Trim(RqF!Seg_Nombre & "")
   xlSheet.Cells(nFil, 7).Value = RqF!fnacimiento
   xlSheet.Cells(nFil, 8).Value = RqF!sexo
   xlSheet.Cells(nFil, 9).Value = Trim(RqF!nacionalidad & "")
   xlSheet.Cells(nFil, 10).Value = Trim(RqF!Cargo & "")
   xlSheet.Cells(nFil, 15).Value = Trim(RqF!TipoTrab & "")
   xlSheet.Cells(nFil, 16).Value = Trim(RqF!Monedasueldo & "")
   xlSheet.Cells(nFil, 17).Value = RqF!sueldo
   xlSheet.Cells(nFil, 19).Value = RqF!fIngreso
   
   nFil = nFil + 1
   RqF.MoveNext
Loop

RqF.Close: Set RqF = Nothing
xlSheet.Range("A:AD").EntireColumn.AutoFit

End Sub

Private Sub Formato_seg_VidaII()

Dim RqF As ADODB.Recordset
Sql = "usp_pla_Seguro_vida_Formato2 " & Txt_Year.Text & "," & Cbo.ListIndex + 1 & ""
If (fAbrRst(RqF, Sql)) Then RqF.MoveFirst Else RqF.Close: Set RqF = Nothing: Exit Sub


Set xlSheet = xlApp2.Worksheets("HOJA6")
xlSheet.Name = "FORM2E"
   
xlSheet.Range("K:K").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Range("F:G").NumberFormat = "dd/mm/yyyy;@"


xlSheet.Cells(3, 1).Value = "CALCULO DE SEGURO DE VIDA - " & Cbo.Text & " " & Txt_Year.Text
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).Merge
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).HorizontalAlignment = xlCenter

xlSheet.Cells(5, 1).Value = "Cod Trabajador"
xlSheet.Cells(5, 2).Value = "Apellido Paterno"
xlSheet.Cells(5, 3).Value = "Apellido Materno"
xlSheet.Cells(5, 4).Value = "Nombres"
xlSheet.Cells(5, 5).Value = "DNI / CE"
xlSheet.Cells(5, 6).Value = "Fecha Nac"
xlSheet.Cells(5, 7).Value = "Fecha Ingreso"
xlSheet.Cells(5, 8).Value = "Fech Ing Seguro(*)"
xlSheet.Cells(5, 9).Value = "Nacionalidad"
xlSheet.Cells(5, 10).Value = "Ocupación , Profesion"
xlSheet.Cells(5, 11).Value = "Sueldo"


xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).Borders.LineStyle = xlContinuous

nFil = 6

Do While Not RqF.EOF
   'If RqF!TipoTrab = "01" Then
   'add jcms 020616
   If RqF!TipoTrab = "E" Then
      xlSheet.Cells(nFil, 1).Value = Trim(RqF!PlaCod & "")
      xlSheet.Cells(nFil, 2).Value = Trim(RqF!Ape_pat & "")
      xlSheet.Cells(nFil, 3).Value = Trim(RqF!Ape_mat & "")
      xlSheet.Cells(nFil, 4).Value = Trim(RqF!nombres & "")
      xlSheet.Cells(nFil, 5).Value = "'" & Trim(RqF!DniTrab & "")
      xlSheet.Cells(nFil, 6).Value = RqF!fnacimiento
      xlSheet.Cells(nFil, 7).Value = RqF!fIngreso
      xlSheet.Cells(nFil, 8).Value = ""
      xlSheet.Cells(nFil, 9).Value = Trim(RqF!nacionalidad & "")
      xlSheet.Cells(nFil, 10).Value = Trim(RqF!Cargo & "")
      xlSheet.Cells(nFil, 11).Value = RqF!sueldo
   nFil = nFil + 1
   End If
   RqF.MoveNext
Loop
xlSheet.Range("A:K").EntireColumn.AutoFit
xlSheet.Cells(1, 1).Value = "CIA. MINERA AGREGADOS CALCAREOS S.A."
'OBREROS
If RqF.RecordCount > 0 Then RqF.MoveFirst
Set xlSheet = xlApp2.Worksheets("HOJA7")
xlSheet.Name = "FORM2O"
   
xlSheet.Range("K:K").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Range("F:G").NumberFormat = "dd/mm/yyyy;@"


xlSheet.Cells(3, 1).Value = "CALCULO DE SEGURO DE VIDA - " & Cbo.Text & " " & Txt_Year.Text
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).Merge
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, 11)).HorizontalAlignment = xlCenter

xlSheet.Cells(5, 1).Value = "Cod Trabajador"
xlSheet.Cells(5, 2).Value = "Apellido Paterno"
xlSheet.Cells(5, 3).Value = "Apellido Materno"
xlSheet.Cells(5, 4).Value = "Nombres"
xlSheet.Cells(5, 5).Value = "DNI / CE"
xlSheet.Cells(5, 6).Value = "Fecha Nac"
xlSheet.Cells(5, 7).Value = "Fecha Ingreso"
xlSheet.Cells(5, 8).Value = "Fech Ing Seguro(*)"
xlSheet.Cells(5, 9).Value = "Nacionalidad"
xlSheet.Cells(5, 10).Value = "Ocupación , Profesion"
xlSheet.Cells(5, 11).Value = "Sueldo"


xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, 11)).Borders.LineStyle = xlContinuous

nFil = 6
If RqF.RecordCount > 0 Then RqF.MoveFirst
Do While Not RqF.EOF
   'If RqF!TipoTrab = "02" Then
   'add jcms 020616
   If RqF!TipoTrab = "O" Then
      xlSheet.Cells(nFil, 1).Value = Trim(RqF!PlaCod & "")
      xlSheet.Cells(nFil, 2).Value = Trim(RqF!Ape_pat & "")
      xlSheet.Cells(nFil, 3).Value = Trim(RqF!Ape_mat & "")
      xlSheet.Cells(nFil, 4).Value = Trim(RqF!nombres & "")
      xlSheet.Cells(nFil, 5).Value = "'" & Trim(RqF!DniTrab & "")
      xlSheet.Cells(nFil, 6).Value = RqF!fnacimiento
      xlSheet.Cells(nFil, 7).Value = RqF!fIngreso
      xlSheet.Cells(nFil, 8).Value = ""
      xlSheet.Cells(nFil, 9).Value = Trim(RqF!nacionalidad & "")
      xlSheet.Cells(nFil, 10).Value = Trim(RqF!Cargo & "")
      xlSheet.Cells(nFil, 11).Value = RqF!sueldo
   nFil = nFil + 1
   End If
   RqF.MoveNext
Loop

RqF.Close: Set RqF = Nothing

xlSheet.Range("A:K").EntireColumn.AutoFit
xlSheet.Cells(1, 1).Value = "CIA. MINERA AGREGADOS CALCAREOS S.A."


For I = 1 To 6
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridLines = False
   'xlApp2.ActiveWindow.Zoom = 80
Next

xlApp2.Sheets(1).Select
xlApp2.Application.Visible = True
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

Screen.MousePointer = vbDefault

End Sub

Public Function Exportar_Excel2() As Boolean
    
    On Error GoTo errSub
    
    Dim cn          As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim mRecordSet As ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
    
    Set mRecordSet = rsTemporal
    
   Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
    Set hoja = Libro.Worksheets(1)
    Excel.Visible = True: Excel.UserControl = True
    iCol = mRecordSet.Fields.count
    For iCol = 1 To mRecordSet.Fields.count
        hoja.Cells(1, iCol).Value = mRecordSet.Fields(iCol - 1).Name
    Next
    
    hoja.Cells(2, 1).CopyFromRecordset mRecordSet
    
    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.Rows.AutoFit
    
    
    Excel.Range("E:E").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
    Excel.Range("I:I").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
   ' xlSheet.Range("F:G").NumberFormat = "dd/mm/yyyy;@"

    If Not mRecordSet Is Nothing Then
        If mRecordSet.State = adStateOpen Then mRecordSet.Close
        Set mRecordSet = Nothing
    End If
   
    Set hoja = Nothing
    Set Libro = Nothing
    Excel.Visible = True
    
    Exportar_Excel2 = True
    
    Exit Function
    
    
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_Excel2 = False

End Function

