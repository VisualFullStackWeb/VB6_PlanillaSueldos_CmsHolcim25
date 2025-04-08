VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmGenAPFNET 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Reportes de Afp «"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "FrmGenAPFNET.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      ItemData        =   "FrmGenAPFNET.frx":030A
      Left            =   1320
      List            =   "FrmGenAPFNET.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtTasa 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   750
      MaxLength       =   4
      TabIndex        =   5
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "FrmGenAPFNET.frx":030E
      Left            =   1680
      List            =   "FrmGenAPFNET.frx":0336
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   80
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   75
         Width           =   825
      End
   End
   Begin MSComctlLib.ProgressBar P1 
      Height          =   135
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label lblTasa 
      Caption         =   "Tasa(%):"
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   285
      Left            =   1350
      TabIndex        =   6
      Top             =   600
      Width           =   255
      Size            =   "450;503"
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "FrmGenAPFNET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mmes As String
Dim vTipoTra As String
Dim VTipo As String
Dim vAfp As String
Dim mlinea As Integer
Dim rs1 As Recordset
Public TipoReporte As String

Private Sub Cmbafp_Click()
If Cmbafp.Text = "TOTAL" Then vAfp = "": mCodAfp = "" Else vAfp = fc_CodigoComboBox(Cmbafp, 2): mCodAfp = fc_CodigoComboBox(Cmbafp, 2)
End Sub

Private Sub Cmbcia_Click()
Cmbmes.ListIndex = Month(fecha) - 2
Txtano.Text = Format(Year(Date), "0000")
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
'Call fc_Descrip_Maestros2("01055", "", Cmbtipotra)
'Cmbtipotra.AddItem "TOTAL"
'Call fc_Descrip_Maestros2("01069", "", Cmbafp)
'Cmbafp.AddItem "TOTAL"
End Sub

Private Sub Cmbmes_Click()
mmes = Format(Cmbmes.ListIndex + 1, "00")
End Sub

Private Sub CmbTipo_Click()
VTipo = Format(Cmbtipo.ListIndex + 1, "00")
End Sub

Private Sub Cmbtipotra_Click()
If Cmbtipotra.Text = "TOTAL" Then vTipoTra = "" Else vTipoTra = fc_CodigoComboBox(Cmbtipotra, 2)
End Sub

Private Sub Command1_Click()
 Call Procesa_RepAfpNet
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5655
Me.Height = 2505
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Cmbmes.ListIndex = Month(Date) - 1
'Cmbtipo.ListIndex = 0

Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Cmbtipo.AddItem "TOTAL"

lblTasa.Visible = IIf(TipoReporte = "S", True, False)
txtTasa.Visible = IIf(TipoReporte = "S", True, False)
Me.Caption = IIf(TipoReporte = "A", "REPORTE AFP", "REPORTE SCTR")
End Sub

Sub Procesa_RepAfpNet()
Dim mCodAfp As String
Dim mNombArchivo As String
Dim mcad As String
Dim tot0 As Currency
Dim tot1 As Currency
Dim tot2 As Currency
Dim tot3 As Currency
Dim tot4 As Currency
Dim tot5 As Currency
Dim totP1 As Currency
Dim totP2 As Currency
Dim totP3 As Currency
Dim totP4 As Currency
Dim totP5 As Currency
Dim mnumtra As Integer
Dim mnumemple As Integer
Dim minicio As Boolean
Dim I As Integer
Dim mtotret As Currency
Dim FILTRAAFP As String
Dim ultfila As Integer
Dim sql1 As String
Dim sql2 As String
Dim mcadQta As String
Dim mcadsnp As String
Dim TrbajadorTipo As String

Dim mTipTrab As String
mTipTrab = ""
If Cmbtipo.ListIndex >= 0 Then If Cmbtipo.Text <> "TOTAL" Then mTipTrab = fc_CodigoComboBox(Cmbtipo, 2)

minicio = True
mcad = ""
tot0 = 0: tot1 = 0: tot2 = 0: tot3 = 0: tot4 = 0: tot5 = 0
totP1 = 0: totP2 = 0: totP3 = 0: totP4 = 0: totP5 = 0
'mCodAfp = fc_CodigoComboBox(Cmbafp, 2)

Sql$ = "SELECT DISTINCT cod_remu AS cod_remu  FROM plaafectos where cia='" & wcia & "' and tipo='D' and tboleta IN('01','04')  and  codigo='11' and status<>'*'"

If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   Do While Not rs.EOF
      mcad = mcad & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop

   mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
   If rs.State = 1 Then rs.Close
   
End If

Select Case TipoReporte
    
    Case Is = "A"
        TrbajadorTipo = mTipTrab
        If mTipTrab = "" Then TrbajadorTipo = "*"
        Sql$ = "usp_Reporte_Afpnet '" & wcia & "'," & Val(Txtano.Text) & "," & Val(mmes) & ",'" & TrbajadorTipo & "'"
     
    Case Is = "S"

        mcad = "sum" & mcad & " as remu,"
        Sql$ = "Select b.placod,rtrim(ap_pat) ap_pat,rtrim(ap_mat) ap_mat,rtrim(nom_1)+' '+rtrim(nom_2) as nombre, b.nro_doc, b.sexo, b.fnacimiento,"
        Sql$ = Sql$ & mcad
        
         mcad = "in ('01','05','02')"
             
         sql1$ = Sql$ & "b.fcese,b.fingreso, COALESCE(RTRIM(m.flag1)+ ' '+ m.descrip,'')  AS ccosto  from plahistorico a,planillas b "
         sql1$ = sql1$ & "LEFT JOIN maestros_2 m ON m.cod_maestro2=b.area AND m.status<>'*' AND ciamaestro='" & wcia & "044' "
         sql1$ = sql1$ & " where a.cia='" & wcia & "' and proceso "
         sql1$ = sql1$ & mcad & " and year(fechaproceso)= " & Val(Txtano.Text) & " "
         sql1$ = sql1$ & "and month(fechaproceso)=" & Val(mmes) & " "
         sql1$ = sql1$ & "and  a.status<>'*' and b.cia=a.cia and a.placod=b.placod and b.status<>'*' "
         If mTipTrab <> "" Then
            sql1$ = sql1$ & "and  a.tipotrab='" & mTipTrab & "' "
         End If
         sql1$ = sql1$ & "AND a.placod NOT IN(SELECT placod FROM plahistorico p WHERE p.cia='" & wcia & "' AND p.proceso = '04'  AND YEAR(fechaproceso)= " & Val(Txtano.Text) & " " & "AND Month(FechaProceso)=" & Val(mmes) & " AND status<>'*' )"
         sql1$ = sql1$ & " group by a.placod,b.ap_pat,b.ap_mat,b.nom_1,"
         sql1$ = sql1$ & "b.nom_2,b.sexo,b.fnacimiento, "
         sql1$ = sql1$ & "B.FCESE,b.nro_doc, b.placod,b.fingreso, COALESCE(RTRIM(m.flag1)+ ' '+ m.descrip,'') "
         
         Sql$ = sql1$ & " ORDER BY b.placod"

    Case Is = "P"
        
        mcadQta = ""
        
        Sql$ = "SELECT DISTINCT cod_remu AS cod_remu  FROM plaafectos WHERE cia='" & wcia & "' and tipo='D' and tboleta IN('01','04')  and  codigo='13' and status<>'*'"
        
        If (fAbrRst(rs, Sql$)) Then
           rs.MoveFirst
           Do While Not rs.EOF
              mcadQta = mcadQta & "i" & Trim(rs!cod_remu) & "+"
              rs.MoveNext
           Loop
        
           mcadQta = "(" & Mid(mcadQta, 1, Len(Trim(mcadQta)) - 1) & ")"
           If rs.State = 1 Then rs.Close
           
        End If
    
         mcadsnp = mcad
        
         mcad = "SUM" & mcad & " AS remu,"
         Sql$ = "SELECT a.placod,rtrim(ap_pat) ap_pat,rtrim(ap_mat) ap_mat,rtrim(nom_1)+' '+rtrim(nom_2) as nombre, b.nro_doc, b.fingreso,"
         Sql$ = Sql$ & mcad & "SUM" & mcadQta & " AS remuQta, SUM(CASE WHEN a.d04<>0 THEN " & mcadsnp & " ELSE 0 END) AS remuSNP, "
        
         mcad = "IN ('01','04','05','02')"
         
         sql1$ = Sql$ & "SUM(d13) AS d13,SUM(h14) AS h14,SUM(d04) AS d04, SUM(a01) AS a01 FROM plahistorico a,planillas b "
         sql1$ = sql1$ & " where a.cia='" & wcia & "' and proceso "
         sql1$ = sql1$ & mcad & " and year(fechaproceso)= " & Val(Txtano.Text) & " "
         If mTipTrab <> "" Then
            sql1$ = sql1$ & "and  a.tipotrab='" & mTipTrab & "' "
         End If

         sql1$ = sql1$ & "and month(fechaproceso)=" & Val(mmes) & " "
         sql1$ = sql1$ & "and  a.status<>'*' and b.cia=a.cia and a.placod=b.placod and b.status<>'*'"
         sql1$ = sql1$ & " group by a.placod,b.ap_pat,b.ap_mat,b.nom_1,"
         sql1$ = sql1$ & "b.nom_2,b.nro_doc,b.fingreso "
         
         Sql$ = sql1$ & " ORDER BY a.placod"
    
End Select



If Not (fAbrRst(rs1, Sql$)) Then
   Exit Sub
End If

Select Case TipoReporte

Case Is = "A"
    Call Excel_ReporteAFPNET
Case Is = "S"
    Call Excel_ReporteSCTR
Case Is = "P"
    Call Excel_ReportePE
    
End Select


End Sub

Private Sub Excel_ReporteAFPNET()
Dim numitem As Integer
numitem = 0
Dim TIPOMOV As String
Dim FECHAMOV As String
Dim fecha As String
fecha = "01/" & mmes & "/" & Txtano.Text


Fila = 1
     Set ObjExcel = CreateObject("Excel.Application")
        
        ObjExcel.Workbooks.Add
        
        'ObjExcel.Application.StandardFont = "arial"
        ObjExcel.Application.StandardFontSize = "9"
        ObjExcel.ActiveWorkbook.Sheets(1).Name = "ReporteAFP"
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
        ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
        With ObjExcel.ActiveSheet.PageSetup
            .LeftMargin = .Application.InchesToPoints(0)
            .RightMargin = .Application.InchesToPoints(0)
            .TopMargin = ObjExcel.Application.InchesToPoints(0.393700787401575)
            .BottomMargin = ObjExcel.Application.InchesToPoints(0.590551181102362)
            .PaperSize = xlPaperLetter
            .Zoom = 75
        End With
        
      
        If rs1.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        P1.Max = rs1.RecordCount
        P1.Min = 1
            'OBJEXCEL.Columns("A:A").ColumnWidth = 12
'            OBJEXCEL.Columns("B:B").ColumnWidth = 60
'            OBJEXCEL.Columns("C:C").ColumnWidth = 13
'            OBJEXCEL.Columns("D:D").ColumnWidth = 10
'            OBJEXCEL.Columns("E:E").ColumnWidth = 13
'            OBJEXCEL.Columns("F:F").ColumnWidth = 13
'            OBJEXCEL.Columns("G:G").ColumnWidth = 12
'            OBJEXCEL.Columns("H:H").ColumnWidth = 11
'            OBJEXCEL.Columns("I:I").ColumnWidth = 10
'            OBJEXCEL.Columns("J:J").ColumnWidth = 11
'            OBJEXCEL.Columns("K:K").ColumnWidth = 11
'            OBJEXCEL.Columns("L:L").ColumnWidth = 50
            ObjExcel.Columns("A").NumberFormat = "#####"
            ObjExcel.Columns("B").NumberFormat = "@"
            
            ObjExcel.Columns("D").NumberFormat = "@"
            ObjExcel.Columns("E").NumberFormat = "@"
            ObjExcel.Columns("F").NumberFormat = "@"
            ObjExcel.Columns("G").NumberFormat = "@"
            ObjExcel.Columns("H").NumberFormat = "@"
            ObjExcel.Columns("I").NumberFormat = "@"
            ObjExcel.Columns("J").NumberFormat = "@"
            ObjExcel.Columns("K").NumberFormat = "@"
            ObjExcel.Columns("L").NumberFormat = "######0.00"
            ObjExcel.Columns("M").NumberFormat = "######0.00"
            ObjExcel.Columns("N").NumberFormat = "######0.00"
            ObjExcel.Columns("O").NumberFormat = "######0.00"
            ObjExcel.Columns("P").NumberFormat = "@"
            ObjExcel.Columns("Q").NumberFormat = "@"
            
            rs1.MoveFirst
            Do While Not rs1.EOF
            
            
            TIPOMOV = "N"
            FECHAMOV = ""
                
            If Not IsNull(rs1!fcese) Then
                If Month(rs1!fcese) = Val(mmes) And Year(rs1!fcese) = Val(Txtano.Text) Then
                   TIPOMOV = "S"
                   FECHAMOV = rs1!fIngreso
                Else
                   If Year(rs1!fcese) > Val(Txtano.Text) Then
                      TIPOMOV = "S"
                   Else
                      If Year(rs1!fcese) = Val(Txtano.Text) Then
                         If Month(rs1!fcese) > Val(mmes) Then TIPOMOV = "S"
                      End If
                   End If
                End If
            Else
               TIPOMOV = "S"
            End If
            
               
'            If Not IsNull(rs1!fechavacai) Then
'                TIPOMOV = "5"
'                'FECHAMOV = UltimoDiaDelMes(CDate(fecha))
'                FECHAMOV = rs1!fechavacai
'            End If
            
             numitem = numitem + 1
             P1.Value = numitem
                ObjExcel.Range("A" & Fila & ":A" & Fila).Value = numitem
                ObjExcel.Range("B" & Fila & ":B" & Fila).Value = Trim(rs1!NUMAFP)
                If Trim(rs1!tipo_doc & "") = "01" Then ObjExcel.Range("C" & Fila & ":C" & Fila).Value = 0
                ObjExcel.Range("D" & Fila & ":D" & Fila).Value = Trim(rs1!nro_doc)
                ObjExcel.Range("E" & Fila & ":E" & Fila).Value = Trim(rs1!ap_pat)
                ObjExcel.Range("F" & Fila & ":F" & Fila).Value = Trim(rs1!ap_mat)
                ObjExcel.Range("G" & Fila & ":G" & Fila).Value = Trim(rs1!nombre)
                ObjExcel.Range("H" & Fila & ":H" & Fila).Value = TIPOMOV
                ObjExcel.Range("I" & Fila & ":I" & Fila).Value = Trim(rs1!ingresodevengue)
                ObjExcel.Range("J" & Fila & ":J" & Fila).Value = Trim(rs1!cesedevengue)
                ObjExcel.Range("K" & Fila & ":K" & Fila).Value = Trim(rs1!Excepcion)
                
                  
                ObjExcel.Range("L" & Fila & ":L" & Fila).Value = Format(rs1!remu, "######0.00")
                ObjExcel.Range("M" & Fila & ":M" & Fila).Value = 0
                ObjExcel.Range("N" & Fila & ":N" & Fila).Value = 0
                ObjExcel.Range("O" & Fila & ":O" & Fila).Value = 0
                ObjExcel.Range("P" & Fila & ":P" & Fila).Value = "N"
                ObjExcel.Range("Q" & Fila & ":Q" & Fila).Value = ""
                Fila = Fila + 1
                rs1.MoveNext
            Loop
            
                 ObjExcel.Visible = True
    
        
                Set ObjExcel = Nothing
                Screen.MousePointer = vbDefault
         Else
            MsgBox "No Existen Datos", vbInformation
         End If
         
    
End Sub
Private Sub Excel_ReporteAFPNET_antes()
Dim numitem As Integer
numitem = 0
Dim TIPOMOV As String
Dim FECHAMOV As String
Dim fecha As String
fecha = "01/" & mmes & "/" & Txtano.Text


Fila = 1
     Set ObjExcel = CreateObject("Excel.Application")
        
        ObjExcel.Workbooks.Add
        
        'ObjExcel.Application.StandardFont = "arial"
        ObjExcel.Application.StandardFontSize = "9"
        ObjExcel.ActiveWorkbook.Sheets(1).Name = "ReporteAFP"
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
        ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
        With ObjExcel.ActiveSheet.PageSetup
            .LeftMargin = .Application.InchesToPoints(0)
            .RightMargin = .Application.InchesToPoints(0)
            .TopMargin = ObjExcel.Application.InchesToPoints(0.393700787401575)
            .BottomMargin = ObjExcel.Application.InchesToPoints(0.590551181102362)
            .PaperSize = xlPaperLetter
            .Zoom = 75
        End With
        
      
        If rs1.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        P1.Max = rs1.RecordCount
        P1.Min = 1
            'OBJEXCEL.Columns("A:A").ColumnWidth = 12
'            OBJEXCEL.Columns("B:B").ColumnWidth = 60
'            OBJEXCEL.Columns("C:C").ColumnWidth = 13
'            OBJEXCEL.Columns("D:D").ColumnWidth = 10
'            OBJEXCEL.Columns("E:E").ColumnWidth = 13
'            OBJEXCEL.Columns("F:F").ColumnWidth = 13
'            OBJEXCEL.Columns("G:G").ColumnWidth = 12
'            OBJEXCEL.Columns("H:H").ColumnWidth = 11
'            OBJEXCEL.Columns("I:I").ColumnWidth = 10
'            OBJEXCEL.Columns("J:J").ColumnWidth = 11
'            OBJEXCEL.Columns("K:K").ColumnWidth = 11
'            OBJEXCEL.Columns("L:L").ColumnWidth = 50
            ObjExcel.Columns("A").NumberFormat = "#####"
            ObjExcel.Columns("B").NumberFormat = "@"
            
            ObjExcel.Columns("D").NumberFormat = "@"
            ObjExcel.Columns("E").NumberFormat = "@"
            ObjExcel.Columns("F").NumberFormat = "@"
            ObjExcel.Columns("G").NumberFormat = "@"
            ObjExcel.Columns("H").NumberFormat = "#"
            ObjExcel.Columns("I").NumberFormat = "dd/mm/yyyy;@"
            ObjExcel.Columns("J").NumberFormat = "######0.00"
            ObjExcel.Columns("K").NumberFormat = "######0.00"
            ObjExcel.Columns("L").NumberFormat = "######0.00"
            ObjExcel.Columns("M").NumberFormat = "######0.00"
            ObjExcel.Columns("N").NumberFormat = "@"
            ObjExcel.Columns("O").NumberFormat = "@"
            
            rs1.MoveFirst
            Do While Not rs1.EOF
            
            
            TIPOMOV = ""
            FECHAMOV = ""
                
            If Not IsNull(rs1!fIngreso) Then
                If Month(rs1!fIngreso) = Val(mmes) And Year(rs1!fIngreso) = Val(Txtano.Text) Then
                        TIPOMOV = "1"
                       FECHAMOV = rs1!fIngreso
                End If
            End If
            
            If Not IsNull(rs1!fcese) Then
                TIPOMOV = "2"
                FECHAMOV = rs1!fcese
            End If
                
'            If Not IsNull(rs1!fechavacai) Then
'                TIPOMOV = "5"
'                'FECHAMOV = UltimoDiaDelMes(CDate(fecha))
'                FECHAMOV = rs1!fechavacai
'            End If
            
             numitem = numitem + 1
             P1.Value = numitem
                ObjExcel.Range("A" & Fila & ":A" & Fila).Value = numitem
                ObjExcel.Range("B" & Fila & ":B" & Fila).Value = Trim(rs1!NUMAFP)
                If Trim(rs1!tipo_doc & "") = "01" Then ObjExcel.Range("C" & Fila & ":C" & Fila).Value = 0
                ObjExcel.Range("D" & Fila & ":D" & Fila).Value = Trim(rs1!nro_doc)
                ObjExcel.Range("E" & Fila & ":E" & Fila).Value = Trim(rs1!ap_pat)
                ObjExcel.Range("F" & Fila & ":F" & Fila).Value = Trim(rs1!ap_mat)
                ObjExcel.Range("G" & Fila & ":G" & Fila).Value = Trim(rs1!nombre)
                ObjExcel.Range("H" & Fila & ":H" & Fila).Value = TIPOMOV
                ObjExcel.Range("I" & Fila & ":I" & Fila).Value = IIf(TIPOMOV = "1", rs1!fIngreso, rs1!fcese)
                ObjExcel.Range("J" & Fila & ":J" & Fila).Value = Format(rs1!remu, "######0.00")
                ObjExcel.Range("K" & Fila & ":K" & Fila).Value = 0
                ObjExcel.Range("L" & Fila & ":L" & Fila).Value = 0
                ObjExcel.Range("M" & Fila & ":M" & Fila).Value = 0
                ObjExcel.Range("N" & Fila & ":N" & Fila).Value = "N"
                ObjExcel.Range("O" & Fila & ":O" & Fila).Value = ""
                Fila = Fila + 1
                rs1.MoveNext
            Loop
            
                 ObjExcel.Visible = True
    
        
                Set ObjExcel = Nothing
                Screen.MousePointer = vbDefault
         Else
            MsgBox "No Existen Datos", vbInformation
         End If
         
    
End Sub

Private Sub Excel_ReporteSCTR()
Dim numitem As Integer
numitem = 0
Dim TIPOMOV As String
Dim FECHAMOV As String
Dim fecha As String
fecha = "01/" & mmes & "/" & Txtano.Text


Fila = 4
     Set ObjExcel = CreateObject("Excel.Application")
        
        ObjExcel.Workbooks.Add
        
        'ObjExcel.Application.StandardFont = "arial"
        ObjExcel.Application.StandardFontSize = "9"
        ObjExcel.ActiveWorkbook.Sheets(1).Name = "ReporteAFP"
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
        ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
        With ObjExcel.ActiveSheet.PageSetup
            .LeftMargin = .Application.InchesToPoints(0)
            .RightMargin = .Application.InchesToPoints(0)
            .TopMargin = ObjExcel.Application.InchesToPoints(0.393700787401575)
            .BottomMargin = ObjExcel.Application.InchesToPoints(0.590551181102362)
            .PaperSize = xlPaperLetter
            .Zoom = 75
        End With
        
      
        If rs1.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        P1.Max = rs1.RecordCount
        P1.Min = IIf(P1.Max = 1, 0, 1)
        
        ObjExcel.Range("A1:A1").Value = Cmbcia.Text
        ObjExcel.Range("C2:C2").Value = "SEGURO COMPLEMENTARIO DE TRABAJO DE RIESGO PENSIONES " + Cmbmes.Text + " " + Txtano.Text
        ObjExcel.Range("A2:J2").Merge
        ObjExcel.Cells(2, 3).HorizontalAlignment = xlCenter
        
        ObjExcel.Range("A3:A3").Value = "CODIGO"
        ObjExcel.Range("B3:B3").Value = "FEC.INICIO"
        ObjExcel.Range("C3:C3").Value = "APELLIDOS Y NOMBRES"
        ObjExcel.Range("D3:D3").Value = "DNI"
        ObjExcel.Range("E3:E3").Value = "FEC. NAC."
        ObjExcel.Range("F3:F3").Value = "SEXO"
        ObjExcel.Range("G3:G3").Value = "TASA %"
        ObjExcel.Range("H3:H3").Value = "REM. SCTR"
        ObjExcel.Range("I3:I3").Value = "REM. ACUM."
        ObjExcel.Range("J3:J3").Value = "C.COSTO"
        
        ObjExcel.Range("A3:J3").Borders.LineStyle = xlContinuous
        
        ObjExcel.Columns("A:A").ColumnWidth = 12
        ObjExcel.Columns("C:C").ColumnWidth = 45
        ObjExcel.Columns("J:J").ColumnWidth = 35
        ObjExcel.Columns("B").NumberFormat = "dd/mm/yyyy;@"
        ObjExcel.Columns("D").NumberFormat = "@"
        ObjExcel.Columns("E").NumberFormat = "dd/mm/yyyy;@"
            
         rs1.MoveFirst
         Do While Not rs1.EOF
            
                numitem = numitem + 1
                P1.Value = numitem
                ObjExcel.Range("A" & Fila & ":A" & Fila).Value = Trim(rs1!PlaCod)
                ObjExcel.Range("B" & Fila & ":B" & Fila).Value = rs1!fIngreso
                ObjExcel.Range("C" & Fila & ":C" & Fila).Value = Trim(rs1!ap_pat) + " " + Trim(rs1!ap_mat) + " " + Trim(rs1!nombre)
                ObjExcel.Range("D" & Fila & ":D" & Fila).Value = Trim(rs1!nro_doc)
                ObjExcel.Range("E" & Fila & ":E" & Fila).Value = rs1!fnacimiento
                ObjExcel.Range("F" & Fila & ":F" & Fila).Value = rs1!sexo
                ObjExcel.Range("G" & Fila & ":G" & Fila).Value = txtTasa.Text
                ObjExcel.Range("H" & Fila & ":H" & Fila).Value = Format(Round((rs1!remu * txtTasa.Text) / 100, 2), "######0.00")
                ObjExcel.Range("I" & Fila & ":I" & Fila).Value = Format(rs1!remu, "#,###,##0.00")
                ObjExcel.Range("J" & Fila & ":J" & Fila).Value = rs1!ccosto
                Fila = Fila + 1
                rs1.MoveNext
            Loop
            
                 ObjExcel.Visible = True
    
        
                Set ObjExcel = Nothing
                Screen.MousePointer = vbDefault
         Else
            MsgBox "No Existen Datos", vbInformation
         End If
         
    
End Sub

Private Sub Excel_ReportePE()
Dim numitem As Integer
numitem = 0
Dim TIPOMOV As String
Dim FECHAMOV As String
Dim fecha As String
fecha = "01/" & mmes & "/" & Txtano.Text


Fila = 4
     Set ObjExcel = CreateObject("Excel.Application")
        
        ObjExcel.Workbooks.Add
        
        'ObjExcel.Application.StandardFont = "arial"
        ObjExcel.Application.StandardFontSize = "9"
        ObjExcel.ActiveWorkbook.Sheets(1).Name = "ReporteAFP"
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
        ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
        With ObjExcel.ActiveSheet.PageSetup
            .LeftMargin = .Application.InchesToPoints(0)
            .RightMargin = .Application.InchesToPoints(0)
            .TopMargin = ObjExcel.Application.InchesToPoints(0.393700787401575)
            .BottomMargin = ObjExcel.Application.InchesToPoints(0.590551181102362)
            .PaperSize = xlPaperLetter
            .Zoom = 75
        End With
        
      
        If rs1.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        P1.Max = rs1.RecordCount
        P1.Min = 1
        
        ObjExcel.Range("A1:A1").Value = Cmbcia.Text
        ObjExcel.Range("C2:C2").Value = "DATOS A TRANSFERIR A LA SUNAT P.E " + Cmbmes.Text + " " + Txtano.Text
        
        ObjExcel.Range("A3:A3").Value = "CODIGO"
        ObjExcel.Range("B3:B3").Value = "FEC.INICIO"
        ObjExcel.Range("C3:C3").Value = "APELLIDOS Y NOMBRES"
        ObjExcel.Range("D3:D3").Value = "DNI"
        ObjExcel.Range("E3:E3").Value = "DIAS. LAB"
        ObjExcel.Range("F3:F3").Value = "REM. SNP"
        ObjExcel.Range("G3:G3").Value = "SNP"
        ObjExcel.Range("H3:H3").Value = "REM. ESSALUD"
        ObjExcel.Range("I3:I3").Value = "ESSALUD"
        ObjExcel.Range("J3:J3").Value = "REM. 5TA."
        ObjExcel.Range("K3:K3").Value = "RTA. 5TA"
        
        ObjExcel.Range("A3:K3").Borders.LineStyle = xlContinuous
        
        ObjExcel.Columns("A:A").ColumnWidth = 12
        ObjExcel.Columns("B").NumberFormat = "dd/mm/yyyy;@"
                    
            
         rs1.MoveFirst
         Do While Not rs1.EOF
            
                numitem = numitem + 1
                P1.Value = numitem
                ObjExcel.Range("A" & Fila & ":A" & Fila).Value = Trim(rs1!PlaCod)
                ObjExcel.Range("B" & Fila & ":B" & Fila).Value = rs1!fIngreso
                ObjExcel.Range("C" & Fila & ":C" & Fila).Value = Trim(rs1!ap_pat) + " " + Trim(rs1!ap_mat) + " " + Trim(rs1!nombre)
                ObjExcel.Range("D" & Fila & ":D" & Fila).Value = Trim(rs1!nro_doc)
                ObjExcel.Range("E" & Fila & ":E" & Fila).Value = Format(rs1!h14, "##")
                ObjExcel.Range("F" & Fila & ":F" & Fila).Value = Format(IIf(rs1!d04 = 0, 0, rs1!remuSNP), "#,###,##0.00")
                ObjExcel.Range("G" & Fila & ":G" & Fila).Value = Format(rs1!d04, "#,###,##0.00")
                ObjExcel.Range("H" & Fila & ":H" & Fila).Value = Format(rs1!remu, "#,###,##0.00")
                ObjExcel.Range("I" & Fila & ":I" & Fila).Value = Format(rs1!a01, "#,###,##0.00")
                ObjExcel.Range("J" & Fila & ":J" & Fila).Value = Format(rs1!remuQta, "#,###,##0.00")
                ObjExcel.Range("K" & Fila & ":K" & Fila).Value = Format(rs1!d13, "#,###,##0.00")
                Fila = Fila + 1
                rs1.MoveNext
            Loop
            
                 ObjExcel.Visible = True
    
        
                Set ObjExcel = Nothing
                Screen.MousePointer = vbDefault
         Else
            MsgBox "No Existen Datos", vbInformation
         End If
         
    
End Sub


Function UltimoDiaDelMes(ByVal dtmFecha As Date) As Date
        UltimoDiaDelMes = DateSerial(Year(dtmFecha), Month(dtmFecha) + 1, 0)
End Function
