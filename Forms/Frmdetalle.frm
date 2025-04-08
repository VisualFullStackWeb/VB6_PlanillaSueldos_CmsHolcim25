VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frmdetalle 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "Frmdetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   300
      Left            =   5400
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Txtsemana 
      Height          =   285
      Left            =   4770
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "Frmdetalle.frx":030A
      Left            =   1200
      List            =   "Frmdetalle.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   720
      Width           =   2415
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   5400
      TabIndex        =   9
      Top             =   705
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Txtcodigo 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   80
         Width           =   4575
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
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Label Lblnombre 
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Lblsemana 
      AutoSize        =   -1  'True
      Caption         =   "Semana"
      Height          =   195
      Left            =   3840
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   720
      Width           =   300
   End
   Begin VB.Label Lblcodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T. Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   960
   End
End
Attribute VB_Name = "Frmdetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mTipo As String
Dim mlinea As Integer

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
End Sub

Private Sub CmbMes_Click()
Txtsemana.Text = ""
End Sub

Private Sub CmbTipo_Click()
If Cmbtipo.Text = "TOTAL" Then mTipo = "" Else mTipo = fc_CodigoComboBox(Cmbtipo, 2)
If mTipo = "01" Then
   If Me.Caption <> "Tabla Dinámica" Then
      Txtsemana.Text = ""
      Txtsemana.Visible = False
      Lblsemana.Visible = False
      UpDown2.Visible = False
   End If
Else
   
   Txtsemana.Visible = True
   Lblsemana.Visible = True
   UpDown2.Visible = True
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5865
Me.Height = 1995
Label3.Caption = "Año"
Lblsemana.Caption = "Semana"

Txtano.Text = Format(Year(Date), "0000")
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Select Case NameForm
       Case Is = "DETALLEQUINTA"
          Me.Caption = "Detalle de Quinta Categoria"
       Case Is = "DINAMICA"
          Me.Caption = "Tabla Dinámica"
          Label3.Caption = "Año Inicial"
          Label3.Left = 3840
          Lblsemana.Caption = "Año Final"
          Txtsemana.Visible = True
          Lblsemana.Visible = True
          UpDown2.Visible = True
          Cmbmes.Visible = False
          Label4.Visible = False
          Txtsemana.Text = Format(Year(Date), "0000")
End Select

If wTipoPla <> "99" And UCase(wuser) <> "SA" Then
   If wTipoPla = "" Then
      Call rUbiIndCmbBox(Cmbtipo, "02", "00")
   Else
      Call rUbiIndCmbBox(Cmbtipo, Trim(wTipoPla), "00")
   End If
   Cmbtipo.Enabled = False
End If
End Sub

Private Sub Txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Txtano.SetFocus
End Sub

Private Sub Txtcodigo_LostFocus()
Sql$ = nombre()
Sql$ = Sql$ + "placod from planillas where status<>'*' " _
     & "and cia='" & wcia & "' AND placod='" & Txtcodigo.Text & "'"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$)
If rs.RecordCount > 0 Then
   Lblnombre.Caption = rs!nombre
ElseIf Trim(Txtcodigo.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodigo.Text
   Txtcodigo.Text = ""
   Lblnombre.Caption = ""
End If
End Sub

Private Sub Txtsemana_LostFocus()
If Txtsemana.Text <> "" Then
   If Cmbmes.Visible Then
      Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Txtano.Text & "' and status<>'*' and semana='" & Format(Txtsemana.Text, "00") & "' and month(fechaf)=" & Cmbmes.ListIndex + 1 & ""
      If Not (fAbrRst(rs, Sql$)) Then MsgBox "Numero de Semana No corresponde al mes", vbInformation, "Detalle": Txtsemana.Text = ""
      If rs.State = 1 Then rs.Close
   End If
End If
End Sub

Private Sub UpDown1_DownClick()
If Txtano.Text = "" Then Txtano.Text = "0"
If Txtano.Text > 0 Then Txtano = Txtano - 1
End Sub

Private Sub UpDown1_UpClick()
If Txtano.Text = "" Then Txtano.Text = "0"
Txtano = Txtano + 1
End Sub
Public Sub Procesar_Detalle()
If Cmbmes.Visible = True Then If Cmbmes.ListIndex < 0 Then MsgBox "Debe Seleccionar Mes del Periodo", vbInformation, "Consultas y Reportes": Exit Sub
If Me.Caption <> "Tabla Dinámica" Then If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Seleccionar Tipo de Trabajador", vbInformation, "Consultas y Reportes": Exit Sub
If Cmbmes.Visible Then If mTipo <> "01" And (Val(Txtsemana) > 53 Or Val(Txtsemana) < 1) And Txtsemana.Text <> "" Then MsgBox "Indique correctamente el Numero de Semana", vbInformation, "Consultas y Reportes": Exit Sub
If Val(Txtano) < 1900 Or Val(Txtano) > 9999 Then MsgBox "Indique correctamente el Año del Periodo", vbInformation, "Consultas y Reportes": Exit Sub
Select Case NameForm
       Case Is = "DETALLEQUINTA"
            Procesa_Detalle_Quinta
       Case Is = "DINAMICA"
            If Not IsNumeric(Txtano.Text) Then MsgBox "Indique Año Inicial Correctamente", vbInformation
            If Not IsNumeric(Txtsemana.Text) Then MsgBox "Indique Año Final Correctamente", vbInformation: Exit Sub
            If Val(Txtano.Text) > Val(Txtsemana.Text) Then MsgBox "Año Final debe ser superior al Inicial", vbInformation: Exit Sub
            Dim mTT As String
            If Cmbtipo.ListIndex < 0 Then mTT = "*" Else mTT = fc_CodigoComboBox(Cmbtipo, 2)
            Screen.MousePointer = vbArrowHourglass
            Dim mTipoT As String
            mTipoT = "*"
            If Cmbtipo.ListIndex > 0 Then mTipoT = fc_CodigoComboBox(Cmbtipo, 2)
            Call Tabla_dinamica("*", Txtano.Text, Txtsemana.Text, mTipoT)
            Screen.MousePointer = vbDefault
End Select
End Sub
Private Sub Procesa_Detalle_Quinta()
  Dim IngresoOtrasEmpresas As String
  Dim IngresoOtraEmpresa As Currency
  Dim QuintaOtraEmpresa As Currency
  Dim rsF02 As ADODB.Recordset
           
'Excel

Dim mcad As String
Dim nFil As Integer
Dim nCol As Integer
'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet


Dim xlApp2 As Object
Dim xlApp1  As Object
Dim xlBook As Object
Dim xlSheet As Object

'Set rsClon = rsReport.Clone
'rsClon.MoveFirst
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Cells(1, 1).Value = Cmbcia
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(1, 1).Font.Size = 12
xlSheet.Cells(1, 1).HorizontalAlignment = xlCenter

xlSheet.Cells(3, 1).Value = "DETALLE DE QUINTA CATEGORIA"
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).Font.Size = 12
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter

If Txtsemana.Text = "" Then
   xlSheet.Cells(4, 1).Value = Space(5) & Cmbmes.Text & Space(5) & Txtano.Text
Else
   xlSheet.Cells(4, 1).Value = Space(5) & "SEMANA : " & Txtsemana.Text & Space(5) & Cmbmes.Text & Space(5) & Txtano.Text
End If
xlSheet.Cells(4, 1).Font.Bold = True
xlSheet.Cells(4, 1).Font.Size = 12
xlSheet.Cells(4, 1).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 1).Value = "MES"
xlSheet.Cells(6, 1).Font.Bold = True
xlSheet.Cells(6, 1).Font.Size = 12
xlSheet.Cells(6, 1).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 2).Value = "INGRESOS"
xlSheet.Cells(6, 2).Font.Bold = True
xlSheet.Cells(6, 2).Font.Size = 12
xlSheet.Cells(6, 2).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 3).Value = "QUINTA"
xlSheet.Cells(6, 3).Font.Bold = True
xlSheet.Cells(6, 3).Font.Size = 12
xlSheet.Cells(6, 3).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 4).Value = "INGRESO OTRA EMPRESA"
xlSheet.Cells(6, 4).Font.Bold = True
xlSheet.Cells(6, 4).Font.Size = 12
xlSheet.Cells(6, 4).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 5).Value = "QUINTA RETENIDA"
xlSheet.Cells(6, 5).Font.Bold = True
xlSheet.Cells(6, 5).Font.Size = 12
xlSheet.Cells(6, 5).HorizontalAlignment = xlCenter

nCol = 1
'xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, nCol - 1)).Merge
'xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, nCol - 1)).Merge

'xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, nCol - 1)).HorizontalAlignment = xlCenter
'xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, nCol - 1)).Font.Bold = True
'xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, nCol - 1)).Borders.LineStyle = xlContinuous
nFil = 8

'Reporte Normal

'Dim mcad As String
Dim mpertope As Integer
Dim msemano As Integer
Dim macui As Currency
Dim macus As Currency
Dim mnombre As String
Dim mnommes As String
Dim mtipotra As String
Dim mcodpla As String
Dim rs2 As ADODB.Recordset

RUTA$ = App.Path & "\REPORTS\" & "Dquinta.txt"
Open RUTA$ For Output As #1
Cabeza_Detalle_Quinta
If wGrupoPla = "01" Then
    Sql$ = "select DISTINCT cod_remu from plaafectos where tipo='D' and tboleta='01'  and  codigo='13' and status<>'*'"
Else
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='13' and status<>'*'"
End If
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcad = ""
   Do While Not rs.EOF
      mcad = mcad & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mTipo <> "01" Then
   Sql$ = "select max(semana) as ultima from plasemanas where cia='" & wcia & "' and ano='" & Txtano.Text & "' and status<>'*'"
   If (fAbrRst(rs, Sql$)) Then mpertope = rs(0): msemano = rs(0)
Else
   mpertope = 12
End If

If rs.State = 1 Then rs.Close

   
Sql$ = "select ph.fechaproceso,ph.placod,sum(" & mcad & ") as ing,sum(d13) as ded,ph.cia "
Sql$ = Sql$ & "from plahistorico ph INNER JOIN  planillas p ON (p.cia='" & wcia & "' and p.placod=ph.placod AND p.status!='*')"
Sql$ = Sql$ & "where ph.cia='" & wcia & "' and ph.status<>'*' and ph.placod LIKE '" & Trim(Txtcodigo.Text) + "%" & "' "

If Len(Trim(mTipo)) > 0 Then
    Sql$ = Sql$ & " and p.tipotrabajador='" & mTipo & "'"
End If

If mTipo <> "01" And Trim(Txtsemana.Text) <> "" Then
   Sql$ = Sql$ & "and year(ph.fechaproceso)=" & Val(Txtano) & _
   " and month(ph.fechaproceso)=" & Cmbmes.ListIndex + 1 & _
   " and ph.semana='" & Txtsemana.Text & _
   "' group by ph.placod,month(ph.fechaproceso),ph.fechaproceso,ph.cia " & _
   "order by ph.placod,month(ph.fechaproceso)"
Else
   Sql$ = Sql$ & "and year(ph.fechaproceso)=" & _
   Val(Txtano) & " and month(ph.fechaproceso)<=" & _
   Cmbmes.ListIndex + 1 & _
   " group by ph.placod,month(ph.fechaproceso),ph.fechaproceso,ph.cia " & _
   " order by ph.placod,month(ph.fechaproceso)"
End If

'Debug.Print SQL$

If (fAbrRst(rs, Sql$)) Then

   mnombre = ""
   mtipotra = ""
   Sql$ = nombre()
   Sql$ = Sql$ + "tipotrabajador from planillas where status<>'*' " _
   & "and cia='" & wcia & "' AND placod='" & rs!PlaCod & "'"
   If (fAbrRst(rs2, Sql$)) Then
  
      xlSheet.Cells(nFil, 1).Value = rs!PlaCod
      xlSheet.Cells(nFil, 2).Value = rs2!nombre
      
      If IngresoOtraEmpresa > 0 Then
         xlSheet.Cells(nFil, 4).Value = fCadNum(IngresoOtraEmpresa, "###,###,###.00")
         xlSheet.Cells(nFil, 5).Value = fCadNum(QuintaOtraEmpresa, "###,###,###.00")
         macui = macui + IngresoOtraEmpresa
         macus = macus + QuintaOtraEmpresa
      End If
      
      nFil = nFil + 1
      Print #1, Space(5) & rs!PlaCod & Space(3) & rs2!nombre & Space(3) & fCadNum(IngresoOtraEmpresa, "###,###,###.00") & Space(3) & fCadNum(QuintaOtraEmpresa, "###,###,###.00")
      Print #1,
      mlinea = mlinea + 2
      If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Detalle_Quinta
      mnombre = rs2!nombre
      mtipotra = rs2!TipoTrabajador
   End If
   If rs2.State = 1 Then rs2.Close
   rs.MoveFirst
   
   mcodpla = Trim(rs!PlaCod)
      
   Do While Not rs.EOF
       If mcodpla <> Trim(UCase(rs!PlaCod)) Then
          xlSheet.Cells(nFil, 1).Value = "TOTALES "
          xlSheet.Cells(nFil, 2).Value = fCadNum(macui, "###,###,###.00")
          xlSheet.Cells(nFil, 3).Value = fCadNum(macus, "###,###,###.00")
          nFil = nFil + 1
          
          Print #1, Space(23) & "--------------" & Space(1) & "--------------"
          Print #1, Space(23) & fCadNum(macui, "###,###,###.00") & Space(1) & fCadNum(macus, "###,###,###.00")
          Print #1,
          macui = 0: macus = 0
          mlinea = mlinea + 3
          If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Detalle_Quinta
  
          Sql$ = "select sum(i18) as util from plahistorico where cia='" & wcia & "' and status<>'*' and placod ='" & Trim(mcodpla) & "' "
          If mTipo <> "01" And Txtsemana.Text <> "" Then
             Sql$ = Sql$ & "and year(fechaproceso)=" & Val(Txtano) & " and month(fechaproceso)<=" & Cmbmes.ListIndex + 1 & " and semana<='" & Txtsemana.Text & "'"
          Else
             Sql$ = Sql$ & "and year(fechaproceso)=" & Val(Txtano) & " and month(fechaproceso)<=" & Cmbmes.ListIndex + 1 & ""
          End If
          If (fAbrRst(rs2, Sql$)) Then
             If Not IsNull(rs2!util) Then xlSheet.Cells(nFil, 1).Value = "     Participacion :": xlSheet.Cells(nFil, 2).Value = fCadNum(rs2!util, "###,###,###.00"): Print #1, "     Participacion :" & Space(3) & fCadNum(rs2!util, "###,###,###.00")
          End If
          If rs2.State = 1 Then rs2.Close
          Print #1,
          Print #1,
          mlinea = mlinea + 2
          nFil = nFil + 1
          If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Detalle_Quinta
          mnombre = ""
          mtipotra = ""
          Sql$ = nombre()
          Sql$ = Sql$ + "tipotrabajador from planillas where status<>'*' " _
          & "and cia='" & wcia & "' AND placod='" & rs!PlaCod & "'"
          If (fAbrRst(rs2, Sql$)) Then
            xlSheet.Cells(nFil, 1).Value = rs!PlaCod
            xlSheet.Cells(nFil, 2).Value = rs2!nombre
            
            mcodpla = Trim(rs!PlaCod)
           'añadido mgirao a solicitud kelly
            IngresoOtrasEmpresas = "N"
            QuintaOtraEmpresa = 0
            IngresoOtraEmpresa = 0
            Sql$ = "select Placod from Pla_Trab_Otras_Empresas where cia='" & wcia & "' and placod='" & mcodpla & "' and status<>'*'"
            If (fAbrRst(rsF02, Sql$)) Then IngresoOtrasEmpresas = "S"
            rsF02.Close: Set rsF02 = Nothing
            Sql$ = ""
           'Remuneraciones del mes de otras empresas
            If IngresoOtrasEmpresas = "S" Then
               Sql$ = "select isnull(SUM(ingreso),0) as basico,isnull(sum(QUINTA),0) AS quinta from Pla_Trab_Otras_Empresas_Meses where Cia='" & wcia & "' AND Ayo=" & Txtano.Text & " and Mes<=" & Cmbmes.ListIndex + 1 & " AND placod='" & mcodpla & "'  and Status<>'*'"
               If (fAbrRst(rsF02, Sql$)) Then IngresoOtraEmpresa = rsF02(0): QuintaOtraEmpresa = rsF02(1)
               If rsF02.State = 1 Then rsF02.Close
            End If
            If IngresoOtraEmpresa > 0 Then
               xlSheet.Cells(nFil, 4).Value = fCadNum(IngresoOtraEmpresa, "###,###,###.00")
               xlSheet.Cells(nFil, 5).Value = fCadNum(QuintaOtraEmpresa, "###,###,###.00")
               macui = macui + IngresoOtraEmpresa
               macus = macus + QuintaOtraEmpresa
            End If
            
            nFil = nFil + 1
             Print #1, Space(5) & rs!PlaCod & Space(3) & rs2!nombre & Space(3) & fCadNum(IngresoOtraEmpresa, "###,###,###.00") & Space(3) & fCadNum(QuintaOtraEmpresa, "###,###,###.00")
             Print #1,
             mlinea = mlinea + 2
             If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Detalle_Quinta
             mnombre = rs2!nombre
             mtipotra = rs2!TipoTrabajador
          End If
          
          If rs2.State = 1 Then rs2.Close
          mcodpla = Trim(rs!PlaCod)
       End If
       If mtipotra = mTipo Then
          mnommes = Name_Month(Format(Month(rs!FechaProceso), "00"))
          mnommes = lentexto(15, Left(mnommes, 15))
          If Not IsNull(rs!ing) Then macui = macui + rs!ing
          If Not IsNull(rs!ded) Then macus = macus + rs!ded
          xlSheet.Cells(nFil, 1).Value = mnommes
          xlSheet.Cells(nFil, 2).Value = fCadNum(rs!ing, "###,###,###.00")
          xlSheet.Cells(nFil, 3).Value = fCadNum(rs!ded, "###,###,###.00")
          nFil = nFil + 1
          Print #1, Space(5) & mnommes & Space(3) & fCadNum(rs!ing, "###,###,###.00") & Space(1) & fCadNum(rs!ded, "###,###,###.00")
          mlinea = mlinea + 1
          If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Detalle_Quinta
       End If
       rs.MoveNext
   Loop
   xlSheet.Cells(nFil, 1).Value = "TOTALES "
   xlSheet.Cells(nFil, 2).Value = fCadNum(macui, "###,###,###.00")
   xlSheet.Cells(nFil, 3).Value = fCadNum(macus, "###,###,###.00")
   nFil = nFil + 1
      
   Print #1, Space(23) & "--------------" & Space(1) & "--------------"
   Print #1, Space(23) & fCadNum(macui, "###,###,###.00") & Space(1) & fCadNum(macus, "###,###,###.00")
   Print #1,
   macui = 0: macus = 0
   'Utilidades
   Sql$ = "select sum(i18) as util from plahistorico where cia='" & wcia & "' and status<>'*' and placod ='" & Trim(mcodpla) & "' "
   If mTipo <> "01" And Txtsemana.Text <> "" Then
      Sql$ = Sql$ & "and year(fechaproceso)=" & Val(Txtano) & " and month(fechaproceso)<=" & Cmbmes.ListIndex + 1 & " and semana<='" & Txtsemana.Text & "'"
   Else
      Sql$ = Sql$ & "and year(fechaproceso)=" & Val(Txtano) & " and month(fechaproceso)<=" & Cmbmes.ListIndex + 1 & ""
   End If
   If (fAbrRst(rs2, Sql$)) Then
      If Not IsNull(rs2!util) Then Print #1, "     Participacion :" & Space(3) & fCadNum(rs2!util, "###,###,###.00")
   End If
   If rs2.State = 1 Then rs2.Close
End If
Close #1
xlSheet.Range("A:AZ").EntireColumn.AutoFit

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "DETALLE DE QUINTA CATEGORIA"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = vbDefault
'Panelprogress.Visible = False

Call Imprime_Txt("Dquinta.txt", RUTA$)
End Sub

Private Sub UpDown2_DownClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "1"
If Txtsemana.Text > 1 Then Txtsemana = Txtsemana - 1
End Sub

Private Sub UpDown2_UpClick()
If Txtsemana.Text = "" Then
   Txtsemana.Text = "1"
Else
   Txtsemana = Txtsemana + 1
End If
End Sub
Private Sub Cabeza_Detalle_Quinta()
Print #1, Space(5) & Cmbcia.Text
Print #1,
Print #1, Space(5) & "DETALLE DE QUINTA CATEGORIA"
Print #1,
If Txtsemana.Text = "" Then
   Print #1, Space(5) & Cmbmes.Text & Space(5) & Txtano.Text
Else
   Print #1, Space(5) & "SEMANA : " & Txtsemana.Text & Space(5) & Cmbmes.Text & Space(5) & Txtano.Text
End If
Print #1, Space(5) & String(70, "-")
Print #1, Space(5) & "MES                   INGRESOS          QUINTA      ING. OTRA EMPR. QUINTA RETENIDA"
Print #1, Space(5) & String(70, "-")
Print #1,
mlinea = 9
End Sub

