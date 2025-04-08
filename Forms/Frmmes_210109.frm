VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmmes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "Frmmes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frameconcepto 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   4695
      Begin VB.ComboBox Cmbconcepto 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   60
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Concepto"
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   80
         Width           =   690
      End
   End
   Begin VB.ComboBox Cmbtipbol 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Txtsemana 
      Height          =   285
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   80
         Width           =   4455
      End
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      ItemData        =   "Frmmes.frx":030A
      Left            =   1320
      List            =   "Frmmes.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "Frmmes.frx":030E
      Left            =   1320
      List            =   "Frmmes.frx":0336
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Data dat 
      Caption         =   "Dta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   705
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Threed.SSPanel Panelprogress 
      Height          =   795
      Left            =   45
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   1393
      _StockProps     =   15
      ForeColor       =   8388608
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.69
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Periodo"
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Lblsemana 
      AutoSize        =   -1  'True
      Caption         =   "Semana"
      Height          =   195
      Left            =   3240
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Lbltbol 
      AutoSize        =   -1  'True
      Caption         =   "T. Boleta"
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSForms.SpinButton SpinButton2 
      Height          =   300
      Left            =   4320
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
      Size            =   "450;529"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1125
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   300
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   255
      Size            =   "450;529"
   End
End
Attribute VB_Name = "Frmmes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdia As Integer
Dim mfecha As String
Dim mtipo As String
Dim mtipobol As String
Dim mlinea As Integer
Dim Vcadir As String
Dim Vcadinr As String
Dim Vcadd As String
Dim Vcada As String
Dim Vcadh As String
Dim Rcadir As String
Dim Rcadinr As String
Dim Rcadd As String
Dim Rcada As String
Dim Rcadh As String
Dim mpag As Integer
Dim mchartipo As String
Dim rs2 As ADODB.Recordset
Dim VConcepto As String
'Dim rpt As Variant

Dim nFil As Integer
Dim nCol As Integer
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
End Sub

Private Sub Cmbconcepto_Click()
VConcepto = fc_CodigoComboBox(Cmbconcepto, 2)
End Sub

Private Sub Cmbtipbol_Click()
If Cmbtipbol.Text = "TOTAL" Then mtipobol = "" Else mtipobol = fc_CodigoComboBox(Cmbtipbol, 2)
End Sub

Private Sub CmbTipo_Click()
If Cmbtipo.Text = "TOTAL" Then mtipo = "" Else mtipo = fc_CodigoComboBox(Cmbtipo, 2)
If InStr(1, Me.Caption, "DEDUCCIONES Y APORTACIONES") > 0 Then
Else
If mtipo <> "01" And mtipo <> "" Then
   Lblsemana.Visible = True
   TxtSemana.Visible = True
   SpinButton2.Visible = True
   TxtSemana.Text = ""
   Me.Height = 2370
Else
   If Me.Caption = "RESUMEN DE PLANILLAS" Then
     ' Me.Height = 2205
   Else
      Me.Height = 1845
   End If
   Lblsemana.Visible = False
   TxtSemana.Visible = False
   SpinButton2.Visible = False
   TxtSemana.Text = ""
End If
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 4860
Me.Height = 2370
Txtano.Text = Format(Year(Date), "0000")
CmbMes.ListIndex = Month(Date) - 2
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
Call fc_Descrip_Maestros2("01078", "", Cmbtipbol)
Cmbtipbol.ListIndex = 0
CmbMes.ListIndex = Month(Date) - 1
Cmbtipbol.Visible = False
Lbltbol.Visible = False
Lblsemana.Visible = False
TxtSemana.Visible = False
SpinButton2.Visible = False
Select Case NameForm
       Case Is = "SEGURODEVIDA"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Me.Caption = "CALCULO DE SEGURO DE VIDA"
            CmbMes.Visible = True
            Cmbtipo.AddItem "TOTAL"
        Case Is = "RESUMEN"
            Me.Height = 2370
            Frameconcepto.Visible = False
            Cmbtipbol.Visible = True
            Lbltbol.Visible = True
            Me.Caption = "RESUMEN DE PLANILLAS"
            CmbMes.Visible = True
            Cmbtipbol.AddItem "TOTAL"
        Case Is = "DEDUCAPOR"
            Me.Caption = "DEDUCCIONES Y APORTACIONES MENSUALES"
            Cmbtipo.AddItem "TOTAL"
            Frameconcepto.Visible = True
            Me.Height = 2370
            Cmbconcepto.Clear
            Sql$ = "select codinterno,descripcion from placonstante where tipomovimiento='03' and status <>'*' AND cia='" & wcia & "' order by descripcion"
            Call rCarCbo(Cmbconcepto, Sql$, "C", "00")
            CmbMes.Visible = True
        Case Is = "CUADROIVF"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Me.Caption = "ESTADISTICO - RESUMEN MENSUAL-ACUMULADO"
            CmbMes.Visible = True
        Case Is = "CUADROIV"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Me.Caption = "ESTADISTICO - HORAS TRABAJADAS EFECTIVAS"
            CmbMes.Visible = True
        Case Is = "SEGURO"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Me.Caption = "SEGURO RIESGO SALUD"
        Case Is = "REMUNERA"
            Me.Caption = "REMUNERACIONES"
            Cmbtipo.AddItem "TOTAL"
            Frameconcepto.Visible = True
            Me.Height = 2370
            Cmbconcepto.Clear
            Sql$ = "select codinterno,descripcion from placonstante where tipomovimiento='02' and status <>'*' AND CIA='" & wcia & "' order by descripcion"
            Call rCarCbo(Cmbconcepto, Sql$, "C", "00")
            CmbMes.Visible = True
        Case Is = "DEDUCAPORANUAL"
            Me.Caption = "DEDUCCIONES Y APORTACIONES ANUALES"
            Cmbtipo.AddItem "TOTAL"
            Frameconcepto.Visible = True
            Me.Height = 2370
            Cmbconcepto.Clear
            Sql$ = "select codinterno,descripcion from placonstante where tipomovimiento='03' and status <>'*' AND cia='" & wcia & "' order by descripcion"
            Call rCarCbo(Cmbconcepto, Sql$, "C", "00")
            CmbMes.Visible = False
        Case Is = "CERTIFICAQTA"
            Me.Caption = "CERTIFICADOS DE QTA CATEGORIA"
            CmbMes.ListIndex = 11
            CmbMes.Visible = False
End Select
Crea_Tablas
End Sub

Private Sub Form_Unload(Cancel As Integer)
dat.RecordSource = ""
dat.Refresh
dat.Database.TableDefs.Delete "Tmpdeduc"
dat.Database.Close
End Sub

Private Sub SpinButton1_SpinDown()
If Txtano.Text = "" Then Txtano.Text = "0"
If Txtano.Text > 0 Then Txtano = Txtano - 1
End Sub

Private Sub SpinButton1_SpinUp()
If Txtano.Text = "" Then Txtano.Text = "0"
Txtano = Txtano + 1
End Sub
Public Sub Procesar()

If CmbMes.ListIndex < 0 Then MsgBox "Debe Seleccionar Mes del Periodo", vbInformation, "Consultas y Reportes": Exit Sub
If Val(Txtano) < 1900 Or Val(Txtano) > 9999 Then MsgBox "Indique correctamente el Año del Periodo", vbInformation, "Consultas y Reportes": Exit Sub
mdia = Ultimo_Dia(CmbMes.ListIndex + 1, Val(Txtano.Text))
mfecha = Format(mdia, "00") & "/" & Format(CmbMes.ListIndex + 1, "00") & "/" & Format(Val(Txtano.Text), "0000")
Select Case NameForm
       Case Is = "SEGURODEVIDA"
            Procesa_Seguro_Vida
       Case Is = "RESUMEN"
            Resumen_Planilla
       Case Is = "DEDUCAPOR"
            If Cmbconcepto.ListIndex < 0 Then
               MsgBox "Seleccione Concepto", vbInformation, "Deducciones /Aportaciones"
            Else
               If VConcepto = "03" Then
                  Procesa_Aporte_Senati
               ElseIf VConcepto = "13" Then
                  Procesa_Lista_Quinta
               ElseIf VConcepto <> "" Then
                  Procesa_Deducciones_Aportaciones
               End If
            End If
       Case Is = "CUADROIVF"
            Procesa_CuadroIVF
       Case Is = "SEGURO"
            Procesa_Seguro
       Case Is = "REMUNERA"
            Procesa_Remunera
       Case Is = "DEDUCAPORANUAL"
            If Cmbconcepto.ListIndex < 0 Then
               MsgBox "Seleccione Concepto", vbInformation, "Deducciones /Aportaciones"
            Else
               If VConcepto = "13" Then
                  Procesa_Lista_Quinta_Anual
               ElseIf VConcepto <> "" Then
                  Procesa_Deduc_Apor_Anuales
               End If
            End If
       Case Is = "CERTIFICAQTA"
             Procesa_Certifica_Qta
 End Select
End Sub
Private Sub Procesa_Seguro_Vida()
Dim mcalc As Integer
Dim mseguro As Boolean
Dim mcad As String
Dim mtoting As Currency
Dim mtotplani As Currency
Dim mItem As Integer
mseguro = False
Sql$ = nombre()
Sql$ = Sql$ & "placod,tipotrabajador,fingreso,area from planillas where cia='" & wcia & "' and tipotrabajador LIKE '" & Trim(mtipo) + "%" & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
RUTA$ = App.path & "\REPORTS\" & "SegVida.txt"
Open RUTA$ For Output As #1
mtoting = 0
mItem = 1
mlinea = 60
Sql$ = "Select sum(totaling-i18) as ing from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
If (fAbrRst(rs2, Sql$)) Then
   If Not IsNull(rs2!ing) Then mtotplani = rs2!ing Else mtotplani = 0
End If
If rs2.State = 1 Then rs2.Close
If mlinea > 55 Then Cabeza_Seguro (mtotplani)
Do While Not rs.EOF
   If Year(rs!fingreso) = Val(Mid(mfecha, 7, 4)) Then
      mcalc = perendat(mfecha, Format(rs!fingreso, "dd/mm/yyyy"), "m")
      If mcalc < 3 Then mseguro = True
   ElseIf Year(rs!fingreso) > Val(Mid(mfecha, 7, 4)) Then
      mseguro = True
   Else
       If Year(rs!fingreso) < (Val(Mid(mfecha, 7, 4)) - 1) Then
          mseguro = False
       Else
          If Val(Mid(mfecha, 4, 2)) > 3 Or Month(rs!fingreso) < 10 Then
             mseguro = False
          Else
             If 12 - (Month(rs!fingreso) - Val(Mid(mfecha, 4, 2))) < 3 Then
                mseguro = True
             ElseIf 12 - (Month(rs!fingreso) - Val(Mid(mfecha, 4, 2))) = 3 Then
                If Day(rs!fingreso) < Val(Mid(mfecha, 1, 2)) Then
                   mseguro = True
                Else
                   mseguro = False
                End If
             Else
                mseguro = False
             End If
          End If
       End If
   End If
   If mseguro = True Then
      Sql$ = "Select sum(totaling-i18) as ing from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
      mcad = ""
      If (fAbrRst(rs2, Sql$)) Then
         If Not IsNull(rs2!ing) Then
            mcad = fCadNum(mItem, "##0") & ".-" & " " & Format(rs!fingreso, "dd/mm/yyyy") & " " & lentexto(40, Left(rs!nombre, 40)) & " (" & fCadNum(rs2!ing, "##,###,##0.00") & ")"
            mtoting = mtoting + rs2!ing
            If rs2.State = 1 Then rs2.Close
            wciamae = Determina_Maestro("01044")
            Sql$ = "Select descrip from maestros_2 where cod_maestro2='" & rs!Area & "' and status<>'*'"
            Sql$ = Sql$ & wciamae
            If (fAbrRst(rs2, Sql$)) Then mcad = mcad & "  " & rs2!descrip
            If rs2.State = 1 Then rs2.Close
            If mlinea > 55 Then Cabeza_Seguro (mtotplani)
            Print #1, mcad
            mlinea = mlinea + 1
         End If
      End If
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
Print #1, Space(58) & "---------------"
Print #1, Space(59) & fCadNum(mtotplani - mtoting, "##,###,##0.00")
Print #1,
Print #1,
Print #1, Space(42) & " 0.60 %" & Space(10) & fCadNum(Round((mtotplani - mtoting) * 0.006, 2), "##,###,##0.00")
Print #1, Space(42) & "18.00 %" & Space(10) & fCadNum(Round(Round((mtotplani - mtoting) * 0.006, 2) * 0.18, 2), "##,###,##0.00")
Print #1, Space(59) & "---------------"
Print #1, Space(59) & fCadNum((Round((mtotplani - mtoting) * 0.006, 2)) + (Round(Round((mtotplani - mtoting) * 0.006, 2) * 0.18, 2)), "##,###,##0.00")
Print #1, Space(59) & "==============="
Close #1
Call Imprime_Txt("SegVida.txt", RUTA$)
End Sub
Private Sub Cabeza_Seguro(Total As Currency)
Print #1, Trim(CmbCia.Text)
Print #1,
Print #1, Space(19) & "CALCULO DE SEGURO DE VIDA"
Print #1,
Print #1, Space(20) & "TOTAL PLANILLA SALARIOS " & Space(15) & fCadNum(Total, "##,###,##0.00")
Print #1,
Print #1, Space(17) & "PERSONAL CON MENOS DE 3 MESES"
Print #1,
Print #1, Space(6) & "FECHA ING.        NOMBRE"
Print #1, String(58, "-")
mlinea = 11
End Sub

Private Sub SpinButton2_SpinDown()
If TxtSemana.Text = "" Then TxtSemana.Text = "53": Exit Sub
If TxtSemana.Text > 1 Then TxtSemana = TxtSemana - 1
End Sub

Private Sub SpinButton2_SpinUp()
If TxtSemana.Text = "" Then TxtSemana.Text = "1": Exit Sub
If TxtSemana < 53 Then TxtSemana = TxtSemana + 1
End Sub
Private Sub Resumen_Planilla()
Dim rscargo As ADODB.Recordset
Dim mmes As Integer
Dim mano As Integer
Dim msem As String
Dim i As Integer
Dim mcad As String
Dim mcadI As String
Dim mcadd As String
Dim mcada As String
Dim mcadh As String
Dim mfor As Integer
Dim mc As Integer
Dim mcargo As String
Dim mremun As Currency
Dim numtra As Integer
Dim rs As New ADODB.Recordset
Dim Inicio As Boolean

If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Seleccionar Tipo de Trabajador", vbInformation, "Resumne de Planilla": Exit Sub
mmes = CmbMes.ListIndex + 1
mano = Val(Txtano.Text)
msem = TxtSemana.Text

If mtipobol <> "01" Then msem = ""
If mtipo = "01" Then msem = ""

'rpt = MsgBox("Desea Imprimir Titulo de Cabecera", vbYesNo)

'*****************desde aqui para la planilla************************************
mcadI = "": mcadd = "": mcada = ""

'*************esto puede ser*********************
For i = 1 To 50
   mcadI = mcadI & "sum(i" & Format(i, "00") & ") as i" & Format(i, "00") & ","
   MCADIT = MCADIT & "i" & Format(i, "00") & "+"
   If i <= 30 Then
      If i <= 20 Then
         mcadd = mcadd & "sum(d" & Format(i, "00") & ") as d" & Format(i, "00") & ","
         mcada = mcada & "sum(a" & Format(i, "00") & ") as a" & Format(i, "00") & ","
         mcaddt = mcaddt & "d" & Format(i, "00") & "+"
         mcadat = mcadat & "a" & Format(i, "00") & "+"
      End If
      If i = 14 Then
        Dim fecha As Date
        fecha = DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text))
        mcadh = mcadh & "CASE WHEN SUM(H14)>" & Day(fecha) & " THEN " & Day(fecha) & "  ELSE SUM(H14) END as h" & Format(i, "00") & ","
        mcadht = mcadht & "h" & Format(i, "00") & "+"
      Else
        mcadh = mcadh & "sum(h" & Format(i, "00") & ") as h" & Format(i, "00") & ","
        mcadht = mcadht & "h" & Format(i, "00") & "+"
      End If
   End If
Next

mcadI = Mid(mcadI, 1, Len(Trim(mcadI)) - 1)
mcadd = Mid(mcadd, 1, Len(Trim(mcadd)) - 1)
mcada = Mid(mcada, 1, Len(Trim(mcada)) - 1)
mcadh = Mid(mcadh, 1, Len(Trim(mcadh)) - 1)

Sql$ = "select " & mcadh & "," & mcadI & "," & mcadd & "," & mcada
Sql$ = Sql$ & ",sum(totaling) as totaling,sum(totalapo) as totalapo,sum(totalded) as totalded,sum(totneto) as totneto,placod "
Sql$ = Sql$ & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & mano & " and month(fechaproceso)=" & mmes & " and proceso LIKE '" & Trim(mtipobol) + "%" & "' and semana LIKE '" & Trim(msem) + "%" & "' " _
     & "and tipotrab='" & mtipo & "' and status<>'*' Group by placod order by placod"
'*****************************************************

If Not (Funciones.fAbrRst(rs, Sql$)) Then MsgBox "No Existen Boletas Registradas Segun Parametros", vbCritical, "Resumen de Planillas": Exit Sub
RUTA$ = App.path & "\REPORTS\" & "RESUMEN.txt"

Open RUTA$ For Output As #1

rs.MoveFirst
mlinea = 60
numtra = 0
Inicio = True
Do While Not rs.EOF
   If mlinea > 50 Then
   If Not Inicio Then
        Print #1, Chr(12) + Chr(13)
    Else
        Inicio = False
   End If

    Call Cabeza_Resumen
    
   End If
   Sql$ = nombre()
   Sql$ = Sql$ & "cargo,fingreso,fcese,ipss from planillas " & _
   "where cia='" & wcia & "' and placod='" & rs!PLACOD & _
   "' and status<>'*'"
  
   If Not (fAbrRst(rs2, Sql$)) Then
      MsgBox "Trabajador no Encontrado en el Maestro " & rs!PLACOD, vbCritical, "Resumen de Planillas"
      Exit Sub
   End If
   
   wciamae = Determina_Maestro_2("01055")
   Sql$ = "select cod_maestro3,descrip from maestros_3 where ciamaestro='" & wcia & "055" & "' and cod_maestro3='" & rs2!cargo & "'"
   'SQL$ = SQL$ & wciamae
 
   If (fAbrRst(rscargo, Sql$)) Then mcargo = rscargo!descrip Else mcargo = ""
   If rscargo.State = 1 Then rscargo.Close
   If IsNull(rs2!fcese) Then mcad = Space(10) Else mcad = Format(rs2!fcese, "dd/mm/yyyy")

   Print #1, rs!PLACOD & Space(5) & lentexto(40, Left(rs2!nombre, 40)) & "  " & lentexto(20, Left(mcargo, 20)) & "  " & Format(rs2!fingreso, "dd/mm/yyyy") & "  " & mcad & lentexto(15, Left(rs2!ipss, 15))
   If rs2.State = 1 Then rs2.Close
   numtra = numtra + 1
   'HORAS
   mfor = Len(Trim(Vcadh))
   mcad = ""
   For i = 1 To mfor Step 2
       mc = Val(Mid(Vcadh, i, 2))
       If rs(mc - 1) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(rs(mc - 1), "##,##0.00") & " "
       End If
   Next i
   Print #1, mcad
   mlinea = mlinea + 1
   
   'Ingresos Remunerativos
   mfor = Len(Trim(Vcadir))
   mcad = ""
   For i = 1 To mfor Step 2
       mc = Val(Mid(Vcadir, i, 2))
       If rs(mc + 29) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(rs(mc + 29), "##,##0.00") & " "
       End If
   Next i
   Print #1, mcad

   mlinea = mlinea + 1
   
   'Ingresos No Remunerativos
   mfor = Len(Trim(Vcadinr))
   mcad = ""
   For i = 1 To mfor Step 2
       mc = Val(Mid(Vcadinr, i, 2))
       If rs(mc + 29) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(rs(mc + 29), "##,##0.00") & " "
       End If
   Next i
   Print #1, mcad
   mlinea = mlinea + 1

   'DEDUCCIONES Y APORTACIONES
   mfor = Len(Trim(Vcadd))
   mcad = ""
   For i = 1 To mfor Step 2
       mc = Val(Mid(Vcadd, i, 2))
       If rs(mc + 79) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(rs(mc + 79), "##,##0.00") & " "
       End If
   Next i
   mfor = Len(Trim(Vcada))
   For i = 1 To mfor Step 2
       mc = Val(Mid(Vcada, i, 2))
       If rs(mc + 99) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(rs(mc + 99), "##,##0.00") & " "
       End If
   Next i
   mcad = mcad & Space(4) & fCadNum(rs!totaling, "####,##0.00") & " " & fCadNum(rs!totalded, "####,##0.00") & " " & fCadNum(rs!totneto, "####,##0.00")
   Print #1, mcad
   mlinea = mlinea + 1
   rs.MoveNext
Loop

Print #1,
Print #1, String(233, "=")
Print #1, Chr(12) + Chr(13)

Call Cabeza_Resumen

'Call Imprimir_Titulo

Print #1,
Print #1, Space(10) & "***** TOTAL PLANILLA *****"
Print #1,
Print #1, "                           H O R A S                             R E M U N E R A C I O N E S                        D E D U C C I O N E S                           A P O R T A C I O N E S"
Print #1, "                           ---------                             ---------------------------                        ---------------------                           -----------------------"
If rs.State = 1 Then rs.Close

Sql$ = "select " & mcadh & "," & mcadI & "," & mcadd & "," & mcada
Sql$ = Sql$ & ",sum(totaling) as totaling,sum(totalapo) as totalapo,sum(totalded) as totalded,sum(totneto) as totneto,placod "
Sql$ = Sql$ & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & mano & " and month(fechaproceso)=" & mmes & " and proceso LIKE '" & Trim(mtipobol) + "%" & "' and semana LIKE '" & Trim(msem) + "%" & "' " _
     & "and tipotrab='" & mtipo & "' and status<>'*' group by cia,placod"
     
     
Sql$ = "select " & mcadh & "," & mcadI & "," & mcadd & "," & mcada
Sql$ = Sql$ & ",sum(totaling) as totaling,sum(totalapo) as totalapo,sum(totalded) as totalded,sum(totneto) as totneto "
Sql$ = Sql$ & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & mano & " and month(fechaproceso)=" & mmes & " and proceso LIKE '" & Trim(mtipobol) + "%" & "' and semana LIKE '" & Trim(msem) + "%" & "' " _
     & "and tipotrab='" & mtipo & "' and status<>'*'"

If (fAbrRst(rs, Sql$)) Then

Dim m As Integer
m = 0
mremun = 0

For i = 1 To 100 Step 2
   m = m + 1
   
   'TOTAL HORAS
   mcad = Space(10)
   If i <= Len(Vcadh) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadh, 1, 20)
      Else
         mcad = mcad & Mid(Rcadh, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadh, i, 2))
      If rs(mc - 1) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc - 1), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(39)
   End If
   
   '=========================
   
   'TOTAL REMUNERACIONES AFECTAS
   mcad = mcad & Space(10)
   If i <= Len(Vcadir) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadir, 1, 20)
      Else
         mcad = mcad & Mid(Rcadir, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadir, i, 2))
      If rs(mc + 29) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc + 29), "###,###,##0.00")
         mremun = mremun + rs(mc + 29)
      End If
   Else
      mcad = mcad & Space(14)
   End If
   
   'TOTAL DEDUCCIONES
   mcad = mcad & Space(10)
   If i <= Len(Vcadd) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadd, 1, 20)
      Else
         mcad = mcad & Mid(Rcadd, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadd, i, 2))
      If rs(mc + 79) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc + 79), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(14)
   End If
   
   'TOTAL APORTACIONES
   mcad = mcad & Space(10)
   If i <= Len(Vcadir) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcada, 1, 20)
      Else
         mcad = mcad & Mid(Rcada, 20 * (m - 1) + 1, 20)
      End If
      mc = Val(Mid(Vcada, i, 2))
      If rs(mc + 99) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc + 99), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(14)
   End If

   If Trim(mcad) <> "" Then Print #1, mcad
'   Debug.Print mcad
Next
Print #1, Space(84) & "--------------"
Print #1, Space(59) & "Sub Total I          " & fCadNum(mremun, "###,###,###,##0.00")
Print #1, Space(65) & "NO REMUNERATIVOS"
Print #1, Space(65) & "----------------"
mremun = 0


For i = 1 To Len(Vcadinr) Step 2
   'TOTAL REMUNERACIONES NO AFECTAS
   mcad = ""
   mcad = mcad & Space(59)
   If i <= Len(Vcadinr) Then
      If i = 1 Then
         mcad = mcad & Mid(Rcadinr, 1, 20)
      Else
         mcad = mcad & Mid(Rcadinr, 20 * (i - 2) + 1, 20)
      End If
      
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadinr, i, 2))
      If rs(mc + 29) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(rs(mc + 29), "###,###,##0.00")
         mremun = mremun + rs(mc + 29)
      End If
   Else
      mcad = mcad & Space(14)
   End If
   If Trim(mcad) <> "" Then Print #1, mcad
Next i

Print #1, Space(84) & "--------------"
Print #1, Space(59) & "Sub Total II         " & fCadNum(mremun, "###,###,###,##0.00")
Print #1, Space(84) & "--------------                                   --------------                                 -----------"
Print #1, Space(59) & "* TOTAL REMUNERACION *" & fCadNum(rs!totaling, "##,###,###,##0.00") & Space(10) & "* TOTAL DEDUCCIONES * " & fCadNum(rs!totalded, "##,###,###,##0.00") & Space(10) & "* TOTAL APORTACIONES *" & fCadNum(rs!totalapo, "#,###,##0.00")
Print #1,
Print #1, Space(59) & "*** NETO PAGADO ***   " & fCadNum(rs!totneto, "##,###,###,##0.00")
Print #1, Space(59) & "======================================="
Print #1,
Print #1, "Total de Trabajadores    =>  " & Str(numtra)
Print #1, Chr(12) + Chr(13)
End If
Close #1
Call Funciones.Imprime_Txt("RESUMEN.txt", RUTA$)
End Sub
Private Sub Cabeza_Resumen()
Dim rsremu As ADODB.Recordset
Dim mcadir As String
Dim mcadinr As String
Dim mcadd As String
Dim mcada As String
Dim mcadh As String
Dim cad As String

mcadir = "": mcadd = "": mcada = "": mcadinr = "": mcadh = ""
Vcadir = "": Vcadd = "": Vcada = "": Vcadinr = "": Vcadh = ""
Rcadir = "": Rcadd = "": Rcada = "": Rcadinr = "": Rcadh = ""


If mtipo = "01" Then
   Rcadh = ""
   Vcadh = ""
Else
Rcadh = ""
   Vcadh = ""
'   mcadh = " NORMAL     DOMINICAL "
'   Rcadh = "HORAS NORMAL        HORAS DOMINICAL     "
   'Vcadh = "0102"
End If
Dim MCADENA As String

If Cmbtipbol.Text = "TOTAL" Then
    MCADENA = ""
Else
    MCADENA = " AND CHARINDEX('" & Left(Cmbtipbol.Text, 1) & "',flag1 )>0 "
End If

Sql$ = "select a.codigo,b.descrip from plaverhoras a," & _
       "maestros_2 b where b.ciamaestro='01077' " & _
       "and a.cia='" & wcia & "' and a.tipo_trab='" & _
       mtipo & "' and a.status<>'*' " & MCADENA _
       & "and a.codigo=b.cod_maestro2 order by codigo"
       
If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
Do While Not rs2.EOF
   mcadh = mcadh & " " & lentexto(9, Left(rs2!descrip, 9)) & " "
   Rcadh = Rcadh & lentexto(20, Left(rs2!descrip, 20))
   Vcadh = Vcadh & rs2!codigo
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close

'INGRESOS
Sql$ = "select a.tipo,a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
     & "where a.cia='" & wcia & "' and a.tipo='I' and a.status<>'*' " _
     & "and a.cia=b.cia and b.status<>'*' " & _
     "and b.tipomovimiento='02' and " & _
     "a.tipo_trab='" & mtipo & "' and a.codigo=b.codinterno " _
     & "order by a.tipo,a.codigo"

If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
Do While Not rs2.EOF
   Sql$ = "select cod_remu from plaafectos where cia='" & _
   wcia & "' and status<>'*' and cod_remu='" & rs2!codigo & _
   "' and tipo in ('A','D') AND CODIGO!='13' "
   If (fAbrRst(rsremu, Sql$)) Then
      mcadir = mcadir & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Rcadir = Rcadir & lentexto(20, Left(rs2!Descripcion, 20))
      Vcadir = Vcadir & rs2!codigo
   Else
      Rcadinr = Rcadinr & lentexto(20, Left(rs2!Descripcion, 20))
      mcadinr = mcadinr & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Vcadinr = Vcadinr & rs2!codigo
   End If
   If rsremu.State = 1 Then rsremu.Close
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close

'DEDUCCIONES Y APORTACIONES
Sql$ = "select a.tipo,a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
     & "where a.cia='" & wcia & "' and a.tipo<>'I' and a.status<>'*' " _
     & "and a.cia=b.cia and b.status<>'*' and b.tipomovimiento='03' and a.tipo_trab='" & mtipo & "' and a.codigo=b.codinterno " _
     & "order by a.tipo,a.codigo"
    
If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst

Do While Not rs2.EOF
   Select Case rs2!tipo
          Case Is = "D"
               mcadd = mcadd & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Rcadd = Rcadd & lentexto(20, Left(rs2!Descripcion, 20))
               Vcadd = Vcadd & rs2!codigo
          Case Is = "A"
               Rcada = Rcada & lentexto(20, Left(rs2!Descripcion, 20))
               mcada = mcada & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Vcada = Vcada & rs2!codigo
   End Select
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close


    Print #1, Chr(18) + Trim(CmbCia.Text) + Chr(14)
    'Print #1,
    'Print #1,
'MODIFICADO 11/10/2008 -  RICARDO HINOSTROZA
    If mtipo = "01" Then
        Print #1, Space(40) & Chr(14) & "PLANILLA " & IIf(Cmbtipbol.Text = "NORMAL", " DE PAGO ", Cmbtipbol.Text) & " EMPLEADOS" & Chr(20)
    Else
        Print #1, Space(40) & Chr(14) & "PLANILLA  " & IIf(Cmbtipbol.Text = "NORMAL", " DE PAGO ", Cmbtipbol.Text) & " OBREROS" & Chr(20)
    End If
    Print #1, Chr(18) + Chr(14)
    'Print #1, Chr(14)
'    Print #1, Chr(14)
 '   Print #1, Chr(14)
    Print #1, Chr(20)

    'Print #1, Chr(20)
    Print #1, Space(50) & "MES de " & CmbMes.Text & "  de " & Txtano.Text & Chr(15)
    If mtipo <> "01" And Val(TxtSemana.Text) > 0 Then
        Print #1, "F. Imp " & Format(Date, "dd/mm/yyyy") & Space(10) & "SEMANA No =>  " & TxtSemana.Text
    Else
        Print #1, "F. Imp " & Format(Date, "dd/mm/yyyy")
    End If
    Print #1, String(233, "-")
    Print #1, "Codigo    Apellidos y Nombres del Trabajador        Ocupacion             F. Ingreso   F.cese   I.P.S.S."
    
    'Resumen de Liquidación
    'Modificado el 05/09/2008 / Error cadena Vacia
    'Ricardo Hinostroza
    
    Print #1, String(IIf(Len(mcadh) = 0, 10, Len(mcadh)) / 2 - 5, "-") & "   HORAS  " & String(IIf(Len(mcadh) = 0, 10, Len(mcadh)) / 2 - 5, "-")
    Print #1, mcadh
    Print #1, String(Len(mcadir) / 2 - 12, "-") & " INGRESOS REMUNERATIVOS " & String(Len(mcadir) / 2 - 12, "-")
    Print #1, mcadir
    If Len(mcadinr) > 24 Then
        Print #1, String(Len(mcadinr) / 2 - 13, "-") & " INGRESOS NO REMUNERATIVOS " & String(Len(mcadinr) / 2 - 13, "-")
    Else
        Print #1, "- INGRESOS NO REMUNERATIVOS -"
    End If
    Print #1, mcadinr

    cad = String(Len(mcadd) / 2 - 8, "-") & "   DEDUCCIONES  " & String(Len(mcadd) / 2 - 8, "-") & "     "
    If Len(Trim(mcada)) > 25 Then
        cad = cad & String(Len(mcada) / 2 - 11, "-") & "  APORTACIONES  " & String(Len(mcada) / 2 - 11, "-")
    End If
    Print #1, cad
    Print #1, mcadd & Space(5) & mcada & "*** TOT.REM.    TOT. DED.   TOT. NETO"
    Print #1, String(233, "-")
    
mlinea = 16

End Sub
Private Sub Procesa_Lista_Quinta()
Dim mItem As Integer
Dim mcad As String
Dim mtotq As Currency
mItem = 0
msemana = ""
mpag = 0
If mtipo <> "01" And TxtSemana.Text <> "" Then
   Sql$ = "Select placod,sum(d13) as quinta from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and semana='" & Trim(TxtSemana.Text) & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod order by placod"
Else
   Sql$ = "Select placod,sum(d13) as quinta from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod order by placod"
End If
If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Hay Retencion de Quinta Segun Paramentros", vbInformation, "Lista de Quinta Categoria": Exit Sub
RUTA$ = App.path & "\REPORTS\" & "Lquinta.txt"
Open RUTA$ For Output As #1
rs.MoveFirst
Cabeza_Lista_Quinta
mtotq = 0
Do While Not rs.EOF
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Cabeza_Lista_Quinta
   Sql$ = nombre()
   Sql$ = Sql$ & "placod from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(45, Left(rs2!nombre, 45)) Else mcad = Space(40)
   mcad = rs!PLACOD & "   " & mcad & Space(5) & fCadNum(rs!quinta, "##,###,##0.00")
   Print #1, Space(2) & mcad
   mtotq = mtotq + rs!quinta
   mItem = mItem + 1
   rs.MoveNext
Loop
Print #1,
Print #1, Space(35) & "TOTAL :                 " & fCadNum(mtotq, "###,###,##0.00")
Print #1, Space(35) & "TOTAL TRABAJADORES      " & fCadNum(mItem, "##,###,###,###")
Close #1
Call Imprime_Txt("Lquinta.txt", RUTA$)
End Sub
Private Sub Cabeza_Lista_Quinta()
mpag = mpag + 1
Print #1, Chr(18) & Space(2) & Trim(CmbCia.Text) & Space(25) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(22) & "REPORTE DE QUINTA CATEGORIA"
Print #1, Space(23) & "PERIODO : " & CmbMes.Text & " - " & Format(Txtano.Text, "0000")
Print #1, Space(2) & Cmbtipo.Text
Print #1, Space(2) & String(72, "-")
Print #1, Space(2) & "CODIGO            NOMBRE                                         MONTO"
Print #1, Space(2) & String(72, "-")
mlinea = 10
End Sub
Private Sub Procesa_Aporte_Senati()
Dim mcad As String
Dim mtotsenati As Currency
Dim mtotafecto As Currency
Dim mtotsenatit As Currency
Dim mtotafectot As Currency
Dim mItem As Integer
Dim totaltrab As Integer
Dim mtipo As String
Dim marea As String
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='03' and status<>'*'"
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
mcad = "sum" & mcad & " as afecto "
Sql$ = nombre()
Sql = Sql$ & "p.placod,p.tipotrabajador,p.area,sum(a03) as senati," & mcad & "from planillas p,plahistorico h where h.cia='" & wcia & "' " _
    & "and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and h.status<>'*' and a03<>0 " _
    & "and p.cia=h.cia and p.placod=h.placod and p.status<>'*' " _
    & "group by h.placod order by p.tipotrabajador,p.area"
    
If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Hay Aportaciones al SENATI Segun Paramentros", vbInformation, "Aportaciones al SENATI": Exit Sub
RUTA$ = App.path & "\REPORTS\" & "Senati.txt"
Open RUTA$ For Output As #1
rs.MoveFirst
mtotsenati = 0
mtotafecto = 0
mtipo = ""
marea = ""
mItem = 1
mpag = 0
Call Cabeza_Senati(rs!Area, rs!TipoTrabajador)
totaltrab = 0
Do While Not rs.EOF
   If (mtipo <> rs!TipoTrabajador Or marea <> rs!Area) And mtipo <> "" Then
      Print #1, Space(5) & String(103, "-")
      Print #1, Space(30) & "T O T A L E S .... " & Space(13) & fCadNum(mtotafecto, "###,###,###.00") & Space(2) & fCadNum(mtotsenati, "###,###,###.00")
      Print #1, SaltaPag
      Call Cabeza_Senati(rs!Area, rs!TipoTrabajador)
      mtotsenati = 0
      mtotafecto = 0
      mtipo = rs!TipoTrabajador
      marea = rs!Area
      mItem = 1
   End If
   If mtipo = "" Then mtipo = rs!TipoTrabajador: marea = rs!Area
   Print #1, Space(5) & fCadNum(mItem, "##0") & ".-" & Space(5) & mchartipo & Space(3) & lentexto(40, Left(rs!nombre, 40)) & Space(3) & fCadNum(rs!AFECTO, "###,###,###.00") & Space(2) & fCadNum(rs!senati, "###,###,###.00") & Space(11) & rs!PLACOD
   mlinea = mlinea + 1
   totaltrab = totaltrab + 1
   If mlinea > 55 Then Print #1, SaltaPag: Call Cabeza_Senati(rs!Area, rs!TipoTrabajador)
   mtotsenati = mtotsenati + rs!senati
   mtotafecto = mtotafecto + rs!AFECTO
   mtotsenatit = mtotsenatit + rs!senati
   mtotafectot = mtotafectot + rs!AFECTO
   mItem = mItem + 1
   rs.MoveNext
Loop
Print #1, Space(5) & String(103, "-")
Print #1, Space(30) & "T O T A L E S .... " & Space(13) & fCadNum(mtotafecto, "###,###,###.00") & Space(2) & fCadNum(mtotsenati, "###,###,###.00")
Print #1,
Print #1,
Print #1, Space(5) & String(103, "-")
Print #1, Space(30) & "TOTALES GENERALES  " & Space(13) & fCadNum(mtotafectot, "###,###,###.00") & Space(2) & fCadNum(mtotsenatit, "###,###,###.00")
Print #1, Space(30) & "TOTAL TRABAJADORES AFECTOS : " & fCadNum(totaltrab, "#####")

Close #1
Call Imprime_Txt("Senati.txt", RUTA$)
End Sub
Private Sub Cabeza_Senati(ccosto, tipo)
Dim wciamae As String
Dim mdescosto As String
mdescosto = ""
mpag = mpag + 1
mchartipo = ""
wciamae = Determina_Maestro("01044")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where cod_maestro2='" & ccosto & "' and status<>'*'"
Sql$ = Sql$ & wciamae
If (fAbrRst(rs2, Sql$)) Then mdescosto = rs2!descrip
If rs2.State = 1 Then rs2.Close

Print #1, LetraChica & Space(5) & Trim(CmbCia.Text) & Space(2) & "( " & mdescosto & " )"
mdescosto = ""
wciamae = Determina_Maestro("01055")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where cod_maestro2='" & tipo & "' and status<>'*'"
Sql$ = Sql$ & wciamae
If (fAbrRst(rs2, Sql$)) Then mdescosto = rs2!descrip: mchartipo = Left(rs2!descrip, 1)
If rs2.State = 1 Then rs2.Close
Print #1,
Print #1, Space(20) & "REPORTE DE TRABAJADORES CON APORTACION AL SENATI - MES DE "; CmbMes.Text
Print #1, Space(45) & mdescosto
Print #1, Space(5) & Format(Date, "dd/mm/yyyy") & Space(81) & "Pagina  " & fCadNum(mpag, "###")
Print #1, Space(5) & String(103, "-")
Print #1, Space(5) & "  Orden  Tipo              Nombre                            Remun.Afecta       Senati          Codigo"
Print #1, Space(5) & String(103, "-")
mlinea = 7
End Sub
Private Sub Procesa_Deducciones_Aportaciones()
Dim mItem As Integer
Dim mcad As String
Dim mcadIA As String
Dim mcadID As String
Dim mfijo As Boolean
Dim totafeca As Currency
Dim totafecd As Currency
Dim totapo As Currency
Dim totded As Currency

msemana = ""
mpag = 0
mcad = "sum(d" & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo"
mfijo = False
Sql$ = "select status,adicional from placonstante where tipomovimiento='03' and codinterno='" & VConcepto & "' and status <>'*'"
If (fAbrRst(rs, Sql$)) Then If rs!status = "F" Or rs!adicional = "S" Then mfijo = True
If rs.State = 1 Then rs.Close

If mfijo = False Then
    'Ingresos Afectos para Aportacion
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='" & VConcepto & "' and status<>'*'"
    If (fAbrRst(rs, Sql$)) Then
       rs.MoveFirst
       mcadIA = ""
       Do While Not rs.EOF
          mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
          rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If rs.State = 1 Then rs.Close
    
    'Ingresos Afectos para Deduccion
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='" & VConcepto & "' and status<>'*'"
    If (fAbrRst(rs, Sql$)) Then
       rs.MoveFirst
       mcadID = ""
       Do While Not rs.EOF
          mcadID = mcadID & "i" & Trim(rs!cod_remu) & "+"
          rs.MoveNext
       Loop
       mcadID = "(" & Mid(mcadID, 1, Len(Trim(mcadID)) - 1) & ")"
    End If
    If rs.State = 1 Then rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    If mcadID <> "" Then mcad = mcad & ",sum" & mcadID & " as afectod"
    Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Hay " & Cmbconcepto.Text & " Segun Paramentros", vbInformation, "Deducciones y Aportaciones": Exit Sub
RUTA$ = App.path & "\REPORTS\" & "DeducApor.txt"
Open RUTA$ For Output As #1
rs.MoveFirst
Cabeza_Lista_DeducApor (rs!tipotrab)
mItem = 1
totafeca = 0: totafecd = 0: totapo = 0: totded = 0
Do While Not rs.EOF
   If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Lista_DeducApor (rs!tipotrab)
   Sql$ = nombre()
   Sql$ = Sql$ & "placod from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   mcad = fCadNum(mItem, "##0") & ".-" & "  " & mchartipo & "  " & mcad
   If mfijo = False Then
      If mcadIA <> "" Then mcad = mcad & Space(2) & fCadNum(rs!afectoa, "#,###,###.00") & Space(2) & fCadNum(rs!APO, "###,###.00") & Space(5): totafeca = totafeca + rs!afectoa Else mcad = mcad & Space(31)
      If mcadID <> "" Then mcad = mcad & Space(2) & fCadNum(rs!afectod, "#,###,###.00") & Space(2) & fCadNum(rs!ded, "###,###.00"): totafecd = totafecd + rs!afectod Else mcad = mcad & Space(26)
   Else
      mcad = mcad & Space(16) & fCadNum(rs!APO, "###,###.00") & Space(21) & fCadNum(rs!ded, "###,###.00")
   End If
   mcad = mcad & Space(2) & fCadNum(rs!APO + rs!ded, "#,###,###.00")
   Print #1, Space(5) & mcad
   totapo = totapo + rs!APO
   totded = totded + rs!ded
   mItem = mItem + 1
   rs.MoveNext
Loop
Print #1, Space(2) & String(126, "-")
Print #1, Space(17) & "T O T A L E S .... " & Space(21) & fCadNum(totafeca, "#,###,###.00") & Space(2) & fCadNum(totapo, "###,###.00") & Space(7) & fCadNum(totafecd, "#,###,###.00") & Space(2) & fCadNum(totded, "###,###.00"); Space(2) & fCadNum(totapo + totded, "#,###,###.00")
Close #1
Call Imprime_Txt("DeducApor.txt", RUTA$)
End Sub
Private Sub Cabeza_Lista_DeducApor(tipo)
mpag = mpag + 1
wciamae = Determina_Maestro("01055")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where cod_maestro2='" & tipo & "' and status<>'*'"
Sql$ = Sql$ & wciamae
If (fAbrRst(rs2, Sql$)) Then mchartipo = Left(rs2!descrip, 1)
If rs2.State = 1 Then rs2.Close

Print #1, Chr(15) & Space(2) & Trim(CmbCia.Text) & Space(80) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(2) & "REPORTE DE " & Cmbconcepto.Text
Print #1, Space(50) & "PERIODO : " & CmbMes.Text & " - " & Format(Txtano.Text, "0000")
Print #1, Space(2) & Cmbtipo.Text
Print #1, Space(2) & String(126, "-")
Print #1, Space(2) & "Orden Tipo         Nombre                                Rem. Afecta     Empr.          Rem. Afecta     Trab.          Total"
Print #1, Space(2) & String(126, "-")
mlinea = 10
End Sub
Private Sub Procesa_CuadroIVF()
Dim mcad As String
Dim mcadprint As String
Dim totalano As Currency
Dim totalmes As Currency
Dim totalanon As Currency
Dim totalmesn As Currency
Dim tano As Currency
Dim tmes As Currency
Dim tanoporc As Currency
Dim tmesporc As Currency
Dim msum As Integer
Dim ArrgConceptos() As Variant

Dim printyn As Boolean
If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Ingresar Tipo de Trabajador", vbInformation, "Estadisticos": Exit Sub
totalano = 0: totalmes = 0
mcad = "select sum(totaling)as ingreso,sum(totneto) as neto "
Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
If Not IsNull(rs!ingreso) Then totalanon = rs!neto: totalano = rs!ingreso
If rs.State = 1 Then rs.Close

Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
If Not IsNull(rs!ingreso) Then totalmes = rs!ingreso: totalmesn = rs!neto
If rs.State = 1 Then rs.Close

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 1
xlSheet.Range("B:B").ColumnWidth = 34
xlSheet.Range("C:C").ColumnWidth = 17
xlSheet.Range("E:E").ColumnWidth = 17
xlSheet.Range("G:G").ColumnWidth = 17
xlSheet.Range("D:D").ColumnWidth = 8
xlSheet.Range("F:F").ColumnWidth = 8
xlSheet.Range("H:H").ColumnWidth = 8
xlSheet.Range("C:H").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 2).Value = CmbCia.Text
xlSheet.Cells(1, 2).Font.Size = 12
xlSheet.Cells(1, 2).Font.Bold = True

If mtipo = "01" Then
   xlSheet.Cells(3, 2).Value = "RESUMEN MENSUAL DE PLANILLAS DE SUELDOS "
Else
   xlSheet.Cells(3, 2).Value = "RESUMEN MENSUAL DE PLANILLAS DE JORNALES "
End If
xlSheet.Cells(3, 2).Font.Bold = True
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 8)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 8)).HorizontalAlignment = xlCenter
xlSheet.Cells(4, 2).Value = CmbMes.Text & " " & Txtano.Text
xlSheet.Cells(4, 2).Font.Bold = True
xlSheet.Range(xlSheet.Cells(4, 2), xlSheet.Cells(4, 8)).Merge
xlSheet.Range(xlSheet.Cells(4, 2), xlSheet.Cells(4, 8)).HorizontalAlignment = xlCenter

xlSheet.Cells(6, 2).Value = "** Expresado en Nuevos Soles **"
xlSheet.Cells(7, 2).Value = "CONCEPTO"
xlSheet.Cells(7, 3).Value = "MENSUAL"
xlSheet.Cells(7, 4).Value = "%"
xlSheet.Cells(7, 5).Value = "ACUMULADO"
xlSheet.Cells(7, 6).Value = "%"
xlSheet.Cells(7, 7).Value = "PROMEDIO"
xlSheet.Cells(7, 8).Value = "%"
xlSheet.Range(xlSheet.Cells(7, 2), xlSheet.Cells(7, 8)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(7, 2), xlSheet.Cells(7, 8)).Borders.LineStyle = xlContinuous

nFil = 8
Panelprogress.Caption = "Generando Estadistico"
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Sql$ = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' order by codinterno"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
tano = 0: tmes = 0
Barra.Value = 0
Barra.Max = rs.RecordCount
mcad = ""
MAXROW = 0
Erase ArrgConceptos
Do While Not rs.EOF
   ReDim Preserve ArrgConceptos(0 To MAXROW)
   mcad = mcad & "sum(I" & rs!codinterno & "),"
   ArrgConceptos(MAXROW) = rs!Descripcion
   MAXROW = MAXROW + 1
   rs.MoveNext
Loop
If Len(Trim(mcad)) <> 0 Then
   mcad = "select " & Mid(mcad, 1, Len(Trim(mcad)) - 1)
   printyn = False
   'Acumulado
   mcadprint = ""
   Sql$ = mcad & " from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst

   For i = 0 To rs.RecordCount - 1
      Barra.Value = i
      If Not IsNull(rs2(i)) Then
        xlSheet.Cells(nFil, 2).Value = ArrgConceptos(i)
         If rs2(i) <> 0 Then
            xlSheet.Cells(nFil, 5).Value = rs2(i)
            xlSheet.Cells(nFil, 6).Value = rs2(i) * 100 / totalano
            xlSheet.Cells(nFil, 7).Value = rs2(i) / (CmbMes.ListIndex + 1)
            xlSheet.Cells(nFil, 8).Value = rs2(i) * 100 / totalano
            tano = tano + rs2(i)
            tanoporc = tanoporc + (rs2(i) * 100 / totalano)
            printyn = True
         End If
      End If
      mcadprint = mcad
      'If rs2.State = 1 Then rs2.Close
  
      'Mensual
      If printyn = True Then
         Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
         If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
         If Not IsNull(rs2(i)) Then
            tmes = tmes + rs2(i)
            tmesporc = tmesporc + (rs2(i) * 100 / totalmes)
            xlSheet.Cells(nFil, 3).Value = rs2(i)
            xlSheet.Cells(nFil, 4).Value = rs2(i) * 100 / totalmes
         Else
         End If
         'xlSheet.Cells(nFil, 2).Value = rs!descripcion
         nFil = nFil + 1
         'If rs2.State = 1 Then rs2.Close
      End If
   Next
End If
msum = (9) * -1
nFil = nFil + 1
xlSheet.Cells(nFil, 2).Value = "T O T A L"
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
For i = 3 To 8
   xlSheet.Cells(nFil, i).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next i
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 8)).Font.Bold = True
nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "Cuota Patronal"
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

If rs.State = 1 Then rs.Close

'Aportaciones

Sql$ = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and status<>'*' order by codinterno"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
tano = 0: tmes = 0
rs.MoveFirst
Barra.Value = 0
Barra.Max = rs.RecordCount
Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(A" & rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!ingreso) Then
      If rs2!ingreso <> 0 Then
         xlSheet.Cells(nFil, 5).Value = rs2!ingreso
         xlSheet.Cells(nFil, 7).Value = rs2!ingreso / (CmbMes.ListIndex + 1)
         tano = tano + rs2!ingreso
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(A" & rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!ingreso) Then
         xlSheet.Cells(nFil, 3).Value = rs2!ingreso
         tmes = tmes + rs2!ingreso
      End If
      xlSheet.Cells(nFil, 2).Value = rs!Descripcion
      nFil = nFil + 1
      If rs2.State = 1 Then rs2.Close
   End If
   rs.MoveNext
Loop

nFil = nFil + 1

xlSheet.Cells(nFil, 2).Value = "Descuento Personal"
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

'Deducciones
rs.MoveFirst
Barra.Value = 0
Barra.Max = rs.RecordCount
Do While Not rs.EOF
   Barra.Value = rs.AbsolutePosition
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(D" & rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!ingreso) Then
      If rs2!ingreso <> 0 Then
         xlSheet.Cells(nFil, 5).Value = rs2!ingreso
         xlSheet.Cells(nFil, 7).Value = rs2!ingreso / (CmbMes.ListIndex + 1)
         tano = tano + rs2!ingreso
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(D" & rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!ingreso) Then
         xlSheet.Cells(nFil, 3).Value = rs2!ingreso
         tmes = tmes + rs2!ingreso
      End If
      xlSheet.Cells(nFil, 2).Value = rs!Descripcion
      nFil = nFil + 1
      If rs2.State = 1 Then rs2.Close
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
Panelprogress.Visible = False
tmes = tmes + totalmesn
tano = tano + totalanon
nFil = nFil + 1
xlSheet.Cells(nFil, 2).Value = "Neto Percibido Personal"
xlSheet.Cells(nFil, 3).Value = totalmesn
xlSheet.Cells(nFil, 5).Value = totalanon
xlSheet.Cells(nFil, 7).Value = totalanon / (CmbMes.ListIndex + 1)
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 8)).Font.Bold = True
nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "Gasto Total Empresa"
xlSheet.Cells(nFil, 3).Value = tmes
xlSheet.Cells(nFil, 5).Value = tano
xlSheet.Cells(nFil, 7).Value = tano / (CmbMes.ListIndex + 1)
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 8)).Font.Bold = True

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "RESUMEN DE PLANILLA"
xlApp2.ActiveWindow.Zoom = 80
xlSheet.PageSetup.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

End Sub
Private Sub Procesa_CuadroIV()

End Sub
Private Sub Procesa_Seguro()
Dim mtope As Currency
Dim mperiodo As String
Dim mtiposeg As String
Dim marea As String
Dim wciamae1 As String
Dim wciamae2 As String
Dim mItem As Integer
Dim tottrab As Integer
Dim mtingarea As Currency
Dim mtingplanta As Currency
Dim mtingtotal As Currency
Dim mtsaludarea As Currency
Dim mtsaludplanta As Currency
Dim mtsaludtotal As Currency
Dim mtpensionarea As Currency
Dim mtpensionplanta As Currency
Dim mtpensiontotal As Currency
Dim mtasa1 As Currency
Dim mtasa2 As Currency
Dim M1 As Integer
Dim m2 As Integer
Dim m3 As Integer
Dim m1ss As Currency
Dim m2ss As Currency
Dim m3ss As Currency
Dim m1sp As Currency
Dim m2sp As Currency
Dim m3sp As Currency


mperiodo = Txtano.Text & Format(CmbMes.ListIndex + 1, "00")
mtope = 0
Sql$ = "select tope from plaafp where cia='" & wcia & "' and status<>'*' and periodo='" & mperiodo & "'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   If Not IsNull(rs!TOPE) Then mtope = rs!TOPE
Else
   MsgBox "No se Ha Registrado el Tope Para el Periodo", vbInformation, "Calculo de Seguro (SCRT)"
   Exit Sub
End If
If mtope = 0 Then MsgBox "No se Ha Registrado el Tope Para el Periodo", vbInformation, "Calculo de Seguro (SCRT)": Exit Sub
If rs.State = 1 Then rs.Close
Sql$ = "select distinct(h.placod),p.ap_pat,p.ap_mat,p.nom_1,p.nom_2,p.area,p.tipcalcseguro from planillas p,plahistorico h " _
     & "where h.cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " " _
     & "and h.status<>'*' and p.status<>'*' and p.tipcalcseguro<>'' and tipotrab LIKE '" & Trim(mtipo) + "%" & "'" _
     & "and h.cia=p.cia and h.placod=p.placod order by p.tipcalcseguro,p.area,h.placod"
     
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   marea = rs!Area
   mtiposeg = rs!tipcalcseguro
End If
mtiposeg = ""
marea = ""

RUTA$ = App.path & "\REPORTS\" & "SeguroSCRT.txt"
Open RUTA$ For Output As #1
Cabecera_SCRT (mtope)
Call Quiebre(marea, mtiposeg, True)
mItem = 1
tottrab = 0

mtingarea = 0: mtsaludarea = 0: mtpensionarea = 0
mtingplanta = 0: mtsaludplanta = 0: mtpensionplanta = 0
mtingtotal = 0: mtsaludtotal = 0: mtpensiontotal = 0
M1 = 0: m2 = 0: m3 = 0
m1ss = 0: m2ss = 0: m3ss = 0: m1sp = 0: m2sp = 0: m3sp = 0
mpag = 0
Do While Not rs.EOF
   If mtiposeg <> rs!tipcalcseguro Then
      Print #1, Space(46) & " " & "------------" & " " & "---------" & " " & "---------"
      Print #1, Space(46) & " " & fCadNum(mtingarea, "#,###,###.00") & " " & fCadNum(mtsaludarea, "##,###.00") & " " & fCadNum(mtpensionarea, "##,###.00")
      Print #1, Space(46) & " " & "------------" & " " & "---------" & " " & "---------"
      Print #1, Space(46) & " " & fCadNum(mtingplanta, "#,###,###.00") & " " & fCadNum(mtsaludplanta, "##,###.00") & " " & fCadNum(mtpensionplanta, "##,###.00")
      Print #1, Space(46) & " " & "------------" & " " & "---------" & " " & "---------"
      mlinea = mlinea + 5
      If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
      marea = rs!Area
      mtiposeg = rs!tipcalcseguro
      mtingarea = 0: mtsaludarea = 0: mtpensionarea = 0
      mtingplanta = 0: mtsaludplanta = 0: mtpensionplanta = 0
      Call Quiebre(marea, mtiposeg, True)
      mItem = 1
   ElseIf marea <> rs!Area Then
      Print #1, Space(46) & " " & "------------" & " " & "---------" & " " & "---------"
      Print #1, Space(46) & " " & fCadNum(mtingarea, "#,###,###.00") & " " & fCadNum(mtsaludarea, "##,###.00") & " " & fCadNum(mtpensionarea, "##,###.00")
      mlinea = mlinea + 5
      If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
      marea = rs!Area
      mtiposeg = rs!tipcalcseguro
      mtingarea = 0: mtsaludarea = 0: mtpensionarea = 0
      Call Quiebre(marea, mtiposeg, False)
      mItem = 1
   End If
   mtasa1 = 0: mtasa2 = 0
   Sql$ = "select tipomov,tasa1 from platasaplanilla where cia='" & wcia & "' and codinterno='" & rs!tipcalcseguro & "' and status<>'*' order by tipomov"
   If (fAbrRst(rs2, Sql$)) Then
      rs2.MoveFirst
      Do While Not rs2.EOF
         If rs2!tipomov = "01" Then mtasa1 = rs2!tasa1
         If rs2!tipomov = "02" Then mtasa2 = rs2!tasa1
         rs2.MoveNext
      Loop
   End If
   If rs2.State = 1 Then rs2.Close
   If mtasa1 = 0 Or mtasa2 = 0 Then
      MsgBox "No se han registrado las tasas correspondientes", vbCritical, "SCRT"
      Close #1
      Exit Sub
   End If
   Sql$ = "select sum(totaling-i18) as ing from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and placod='" & rs!PLACOD & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then
      mtasa1 = Round(rs2!ing * mtasa1 / 100, 2)
      If rs2!ing > mtope Then mtasa2 = Round(mtope * mtasa2 / 100, 2) Else mtasa2 = Round(rs2!ing * mtasa2 / 100, 2)
      Print #1, fCadNum(mItem, "###") & ".-" & rs!PLACOD & " " & lentexto(35, Left(RTrim(rs!ap_pat) & " " & RTrim(rs!ap_mat) & " " & RTrim(rs!nom_1) & " " & RTrim(rs!nom_2), 40)) & " " & fCadNum(rs2!ing, "#,###,###.00") & " " & fCadNum(mtasa1, "##,###.00") & " " & fCadNum(mtasa2, "##,###.00")
      mlinea = mlinea + 1
      If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
      Select Case rs!tipcalcseguro
             Case Is = "01"
                  m1ss = m1ss + mtasa1
                  m1sp = m1sp + mtasa2
                  M1 = M1 + 1
             Case Is = "02"
                  m2ss = m2ss + mtasa1
                  m2sp = m2sp + mtasa2
                  m2 = m2 + 1
             Case Is = "02"
                  m3ss = m3ss + mtasa1
                  m3sp = m3sp + mtasa2
                  m3 = m3 + 1
      End Select
      mtingarea = mtingarea + rs2!ing: mtsaludarea = mtsaludarea + mtasa1: mtpensionarea = mtpensionarea + mtasa2
      mtingplanta = mtingplanta + rs2!ing: mtsaludplanta = mtsaludplanta + mtasa1: mtpensionplanta = mtpensionplanta + mtasa2
      mtingtotal = mtingtotal + rs2!ing: mtsaludtotal = mtsaludtotal + mtasa1: mtpensiontotal = mtpensiontotal + mtasa2
      
      
      totrab = tottrab + 1
      mItem = mItem + 1
   End If
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
Loop
Print #1, Space(46) & " " & "------------" & " " & "---------" & " " & "---------"
Print #1, Space(46) & " " & fCadNum(mtingarea, "#,###,###.00") & " " & fCadNum(mtsaludarea, "##,###.00") & " " & fCadNum(mtpensionarea, "##,###.00")
Print #1, Space(46) & " " & "------------" & " " & "---------" & " " & "---------"
Print #1, Space(46) & " " & fCadNum(mtingplanta, "#,###,###.00") & " " & fCadNum(mtsaludplanta, "##,###.00") & " " & fCadNum(mtpensionplanta, "##,###.00")
Print #1, Space(46) & " " & "------------" & " " & "---------" & " " & "---------"

mlinea = mlinea + 5
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)

Print #1, Space(46) & " " & "============" & " " & "=========" & " " & "========="
Print #1, Space(46) & " " & fCadNum(mtingtotal, "#,###,###.00") & " " & fCadNum(mtsaludtotal, "##,###.00") & " " & fCadNum(mtpensiontotal, "##,###.00")
Print #1, Space(46) & " " & "============" & " " & "=========" & " " & "========="

mlinea = mlinea + 3
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)

Print #1,
Print #1,
mlinea = mlinea + 2
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, Space(2) & "T.Trabajado                Resumen " & Cmbtipo.Text
Print #1, Space(29) & "--------------------------"
mlinea = mlinea + 2
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1,
Print #1, fCadNum(M1, "####") & Space(8) & "PERSONAL DE PLANTA                             " & fCadNum(m1ss, "###,###.00") & " " & fCadNum(m1sp, "##,###.00")
Print #1, fCadNum(m2, "####") & Space(8) & "PERSONAL ADMINISTRATIVO EN PLANTA              " & fCadNum(m2ss, "###,###.00") & " " & fCadNum(m2sp, "##,###.00")
Print #1, fCadNum(m3, "####") & Space(8) & "PERSONAL ADMINISTRATIVO SOLO OFICINA           " & fCadNum(m3ss, "###,###.00") & " " & fCadNum(m3sp, "##,###.00")
mlinea = mlinea + 4
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, Space(58) & " " & " " & "=========" & " " & "========="
Print #1, fCadNum(m3 + m2 + M1, "####") & Space(55) & fCadNum(m3ss + m2ss + m1ss, "###,###.00") & " " & fCadNum(m3sp + m2sp + m1sp, "##,###.00")
Print #1, Space(58) & " " & " " & "=========" & " " & "========="
Close #1
Call Imprime_Txt("SeguroSCRT.txt", RUTA$)
End Sub
Private Sub Quiebre(Area As String, planta As String, otra As Boolean)
Print #1,
If otra = True Then
   wciamae2 = Determina_Maestro("01074")
   Sql$ = "Select descrip from maestros_2 where cod_maestro2='" & planta & "' and status<>'*'"
   Sql$ = Sql$ & wciamae2
   If (fAbrRst(rs2, Sql$)) Then
      mtiposeg = Trim(cod_maestro2)
      Print #1, rs2!descrip
   End If
   Print #1,
   If rs2.State = 1 Then rs2.Close
End If

wciamae1 = Determina_Maestro("01044")
Sql$ = "Select descrip from maestros_2 where cod_maestro2='" & Area & "' and status<>'*'"
Sql$ = Sql$ & wciamae1
If (fAbrRst(rs2, Sql$)) Then
   mtiposeg = Trim(cod_maestro2)
   Print #1, rs2!descrip
End If
Print #1,
If rs2.State = 1 Then rs2.Close
End Sub
Private Sub Cabecera_SCRT(TOPE As Currency)
mpag = mpga + 1
Print #1, CmbCia.Text & "    " & Cmbtipo.Text & "  MES DE " & CmbMes.Text & " DEL " & Txtano.Text
Print #1, "TOPE : " & fCadNum(TOPE, "###,###,###.00") & Space(48) & "Pag : " & fCadNum(mpag, "####")
Print #1,
Print #1, "CODIGO         NOMBRE                          TOT, INGRESO  S.RIESGO  S.COMPL."
Print #1, "                                                              SALUD   PENSIONES"
Print #1, String(79, "-")
mlinea = 7
End Sub
Private Sub Procesa_Remunera()
Dim mnumh As Integer
Dim mItem As Integer
Dim mcad As String
Dim totali As Currency
Dim totalh As Currency
If Cmbconcepto.ListIndex < 0 Then MsgBox "Debe Seleccionar Concepto", vbInformation, "Remuneraciones": Exit Sub
msemana = ""
mpag = 0
mcad = "sum(i" & Format(VConcepto, "00") & ") as ing"
mnumh = Remun_Horas(VConcepto)
If mnumh > 0 Then mcad = mcad & ",sum(h" & Format(mnumh, "00") & ") as horas"

Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
If Not (fAbrRst(rs, Sql$)) Then MsgBox "No Hay " & Cmbconcepto.Text & " Segun Paramentros", vbInformation, "Deducciones y Aportaciones": Exit Sub
RUTA$ = App.path & "\REPORTS\" & "Remunera.txt"
Open RUTA$ For Output As #1
rs.MoveFirst
Cabeza_Remunera (rs!tipotrab)
mItem = 1
totali = 0: totalh = 0
Do While Not rs.EOF
   If rs!ing <> 0 Then
      If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Remunera (rs!tipotrab)
      Sql$ = nombre()
      Sql$ = Sql$ & "placod from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      mcad = fCadNum(mItem, "##0") & ".-" & "  " & mchartipo & "  " & mcad
      If mnumh > 0 Then
         mcad = mcad & Space(2) & fCadNum(rs!horas, "#####.00") & Space(3) & fCadNum(rs!ing, "###,###.00")
      Else
         mcad = mcad & Space(13) & fCadNum(rs!ing, "###,###.00")
      End If
      Print #1, Space(5) & mcad
      totali = totali + rs!ing
      If mnumh > 0 Then
         totalh = totalh + rs!horas
      End If
      mItem = mItem + 1
   End If
   rs.MoveNext
Loop
Print #1, Space(2) & String(76, "-")
Print #1, Space(57) & fCadNum(totalh, "#####.00") & Space(3) & fCadNum(totali, "###,###.00")
Close #1
Call Imprime_Txt("Remunera.txt", RUTA$)
End Sub
Private Sub Cabeza_Remunera(tipo)
mpag = mpag + 1
wciamae = Determina_Maestro("01055")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where cod_maestro2='" & tipo & "' and status<>'*'"
Sql$ = Sql$ & wciamae
If (fAbrRst(rs2, Sql$)) Then mchartipo = Left(rs2!descrip, 1)
If rs2.State = 1 Then rs2.Close

Print #1, Chr(15) & Space(2) & Trim(CmbCia.Text) & Space(30) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(2) & "REPORTE DE " & Cmbconcepto.Text
Print #1, Space(30) & "PERIODO : " & CmbMes.Text & " - " & Format(Txtano.Text, "0000")
Print #1, Space(2) & Cmbtipo.Text
Print #1, Space(2) & String(76, "-")
Print #1, Space(2) & "Orden  Tipo        Nombre                                HORAS      IMPORTE"
Print #1, Space(2) & String(76, "-")
mlinea = 10
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Procesa_Deduc_Apor_Anuales()
Dim mItem As Integer
Dim mcad As String
Dim mcadIA As String
Dim mcadID As String
Dim mfijo As Boolean
Dim cadnombre As String
Dim mtot As Currency
Dim mtipoB As String
Dim XPLACOD As String

cadnombre = nombre()
msemana = ""
mpag = 0
mtipoB = ""

mcad = "sum(d" & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo "

Sql$ = "select " & mcad & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*'"

If (fAbrRst(rs, Sql$)) Then
   If rs!ded <> 0 And rs!APO <> 0 Then
      If MsgBox("Desea Reporte de Aportacion : SI= Aportacion    NO= Deduccion", vbQuestion + vbYesNo + vbDefaultButton1, "Reportes Anuales") = vbNo Then mtipoB = "A" Else mtipoB = "D"
   ElseIf rs!ded <> 0 Then
      mtipoB = "D"
   ElseIf rs!APO <> 0 Then
      mtipoB = "A"
   End If
End If

If rs.State = 1 Then rs.Close

mcad = ""


mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False
Sql$ = "select status,adicional from placonstante where tipomovimiento='03' and codinterno='" & VConcepto & "' and status <>'*'"
If (fAbrRst(rs, Sql$)) Then If rs!status = "F" Or rs!adicional = "S" Then mfijo = True
If rs.State = 1 Then rs.Close

If mfijo = False Then
    'Ingresos Afectos para Aportacion Normal
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(rs, Sql$)) Then
       rs.MoveFirst
       mcadIA = ""
       Do While Not rs.EOF
          mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
          rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If rs.State = 1 Then rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
    If VConcepto = "04" Then 'ONP
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
        'Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then 'AFP
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*'  Group by placod,tipotrab order by placod"
    End If
Else
    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
        ' Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
        ' Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND (i05=0 OR i06 =0 ) Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    Else
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*'  Group by placod,tipotrab order by placod"
    End If
End If

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
'Cabeza_Lista_DeducApor (RS!tipotrab)
mItem = 1
Do While Not rs.EOF
   Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PLACOD = Trim(rs!PLACOD)
   
   'NUEVO CODIGO
   If XPLACOD = "" Then
    XPLACOD = "'" & Trim(rs!PLACOD) & "'"
   Else
    XPLACOD = XPLACOD & ",'" & Trim(rs!PLACOD) & "'"
   End If
   
   dat.Recordset!nombre = Trim(mcad)
   If mfijo = False Then
      If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
   End If
   dat.Recordset!importe = rs!APO
   dat.Recordset!AFP = rs!AFP
   dat.Recordset.Update
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Vacaciones
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False

If mfijo = False Then
    'Ingresos Afectos para Aportacion Vacaciones
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='02'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(rs, Sql$)) Then
       rs.MoveFirst
       mcadIA = ""
       Do While Not rs.EOF
          mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
          rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If rs.State = 1 Then rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
       If VConcepto = "04" Then
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
            'Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' and codafp='01' and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       ElseIf VConcepto = "11" Then
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' and codafp<>'01' and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       Else
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*'  and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       End If
Else
    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*'  Group by placod,tipotrab order by placod"
    End If
End If

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!vaca = rs!afectoa
      End If
      dat.Recordset!importe = rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!vaca = rs!afectoa
      End If
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset.Update
      
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Gratificacion
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False

If mfijo = False Then
    'Ingresos Afectos para Aportacion Gratificacion
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='03'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(rs, Sql$)) Then
       rs.MoveFirst
       mcadIA = ""
       Do While Not rs.EOF
          mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
          rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If rs.State = 1 Then rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*'  Group by placod,tipotrab order by placod"
        'Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND codafp='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
    End If
    
Else

    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
        'Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND codafp ='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
    End If
    
End If

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Trim(rs!PLACOD) + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      End If
      dat.Recordset!importe = rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      End If
      
      If VConcepto = "04" Then
        dat.Recordset!importe = dat.Recordset!importe + rs!ded 'apo -DED
      ElseIf VConcepto = "11" Then
        dat.Recordset!importe = dat.Recordset!importe + rs!ded  'apo -DED
      End If
      
     ' dat.Recordset!importe = dat.Recordset!importe + rs!APO 'apo -DED
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset.Update
      
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

'Liquidaciones y Subsidios
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo "
If VConcepto = "04" Then
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND CODAFP='01' Group by placod,tipotrab order by placod"
ElseIf VConcepto = "11" Then
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND CODAFP<>'01' Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!liquid = rs!afectoa
      End If
      dat.Recordset!importe = rs!APO
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!liquid = rs!afectoa
      End If
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset.Update
      
   End If
   rs.MoveNext
Loop

'MODIFICACION DE RICARDO HINOSTROZA PARA AFP
If VConcepto = "11" Then
    Sql$ = "SELECT PLACOD," & _
            "SUM(D11) As DED " & _
            "From PLAHISTORICO" & _
            " where cia='" & wcia & "'" & _
            " and year(fechaproceso)=" & Val(Txtano.Text) & _
            " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' " & _
            " AND status<>'*' " & _
            " AND codafp<>'01' " & _
            " Group by placod,tipotrab " & _
            " order by placod"
        
    If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
        Do While Not rs.EOF
        dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
        If Not dat.Recordset.NoMatch Then
              dat.Recordset.Edit
              dat.Recordset!importe = rs!ded
              dat.Recordset.Update
           rs.MoveNext
        End If
        Loop
End If

'MODIFICACION DE RICARDO HINOSTROZA PARA ESSALUD (NO FIGURA CORRECTAMENTE, DE CORRIJE CON EL SGTE CODIGO
If VConcepto = "01" Then

   Sql$ = "SELECT PLACOD, " & _
        "SUM(a01) As DED " & _
        "From PLAHISTORICO " & _
        "where cia='" & wcia & "'" & _
        "and year(fechaproceso)=" & Val(Txtano.Text) & _
        " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' " & _
        "AND status<>'*' " & _
        "Group by placod,tipotrab " & _
        "order by placod"
        
        If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
        Do While Not rs.EOF
  '      If Trim(rs!PLACOD) = "PO067" Then Stop
        dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
        
        If Not dat.Recordset.NoMatch Then
              dat.Recordset.Edit
              dat.Recordset!importe = rs!ded
              dat.Recordset.Update
           rs.MoveNext
        End If
        Loop
        
 End If

If rs.State = 1 Then rs.Close
RUTA$ = App.path & "\REPORTS\" & "BenAnual.txt"
Open RUTA$ For Output As #1
dat.Refresh
With dat.Recordset
If Not .EOF Then .MoveFirst
miten = 1
Cabeza_Deduc_Anual (mtipoB)
mtot = 0
.MoveFirst
Do While Not .EOF

NOM$ = Trim(!nombre)

If Len(NOM) = 26 Then NOM = NOM & Space(2)
 'If Trim(!PLACOD) = "PO067" Then Stop
   If !sueldo <> 0 Or !AFP <> 0 Or !grati <> 0 Or !vaca <> 0 Or !liquid <> 0 Or !importe <> 0 Then
    mcad = fCadNum(mItem, "####") & ".- " & lentexto(36, Trim(NOM)) & Space(1) & fCadNum(!sueldo, "####,###.00") & Space(2) & fCadNum(!AFP, "###,###.00") & Space(2)
    mcad = mcad & fCadNum(!grati, "####,###.00") & Space(2) & fCadNum(!vaca, "####,###.00") & Space(2) & fCadNum(!liquid, "####,###.00") & Space(2)
    mcad = mcad & fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!vaca) + Val(!liquid), "####,###.00") & Space(1) & fCadNum(!importe, "####,###.00")
    Print #1, mcad
    mtot = mtot + !importe
    mItem = mItem + 1
    mlinea = mlinea + 1
    If mlinea > 55 Then Print #1, SaltaPag: Call Cabeza_Deduc_Anual(mtipoB)
  End If
  .MoveNext
Loop
Print #1, Space(120) & "----------"
Print #1, Space(118) & fCadNum(mtot, "#,###,###.00")
End With
Close #1
Call Imprime_Txt("BenAnual.txt", RUTA$)
dat.Database.Execute "delete from Tmpdeduc"
End Sub
Private Sub Crea_Tablas()
Dim tdf0 As TableDef
dat.DatabaseName = nom_BD
dat.Refresh
Set tdf0 = dat.Database.CreateTableDef("Tmpdeduc")
With tdf0
        .Fields.Append .CreateField("placod", dbText)
        .Fields(0).AllowZeroLength = True
        .Fields.Append .CreateField("nombre", dbText)
        .Fields(1).AllowZeroLength = True
        .Fields.Append .CreateField("sueldo", dbText)
        .Fields(2).AllowZeroLength = True
        .Fields(2).DefaultValue = "0.00"
        .Fields.Append .CreateField("afp", dbText)
        .Fields(3).AllowZeroLength = True
        .Fields(3).DefaultValue = "0.00"
        .Fields.Append .CreateField("grati", dbText)
        .Fields(4).AllowZeroLength = True
        .Fields(4).DefaultValue = "0.00"
        .Fields.Append .CreateField("vaca", dbText)
        .Fields(5).AllowZeroLength = True
        .Fields(5).DefaultValue = "0.00"
        .Fields.Append .CreateField("liquid", dbText)
        .Fields(6).AllowZeroLength = True
        .Fields(6).DefaultValue = "0.00"
        .Fields.Append .CreateField("total", dbText)
        .Fields(7).AllowZeroLength = True
        .Fields(7).DefaultValue = "0.00"
        .Fields.Append .CreateField("importe", dbText)
        .Fields(8).AllowZeroLength = True
        .Fields(8).DefaultValue = "0.00"
        .Fields.Append .CreateField("util", dbText)
        .Fields(9).AllowZeroLength = True
        .Fields(9).DefaultValue = "0.00"
        .Fields.Append .CreateField("afp03", dbText)
        .Fields(10).AllowZeroLength = True
        .Fields(10).DefaultValue = "0.00"
        .Fields.Append .CreateField("cese", dbText)
        .Fields(10).AllowZeroLength = True
        dat.Database.TableDefs.Append tdf0
    End With
    dat.RecordSource = "Tmpdeduc"
    dat.Refresh
    Set tdf0 = Nothing
End Sub
Private Sub Cabeza_Deduc_Anual(tipo)

'CAMBIO AQUI
Print #1, LetraChica

Select Case VConcepto
Case VConcepto
Case Is = "01"
    mtipoB = "A"
End Select

Select Case VConcepto
Case Is = "04"
    Print #1, CmbCia.Text & Space(10) & "SISTEMA NACION DE PENSIONES " & Txtano.Text
    Print #1, "ORDEN           NOMBRE                            SUELDO     INC. AFP    GRATIF.     VACACIO.     Liq/Subs  SNP.REM TOTAL     IMPORTE"
Case Is = "11"
    Print #1, CmbCia.Text & Space(10) & IIf(mtipoB = "A", "APORTES A     ", "DEDUCCIONES A ") & Trim(Cmbconcepto.Text) & " " & Txtano.Text
    Print #1, "ORDEN           NOMBRE                            SUELDO     INC. AFP    GRATIF.     VACACIO.     Liq/Subs  AFP.REM TOTAL     IMPORTE"
Case Is = "01"
    Print #1, CmbCia.Text & Space(10) & IIf(mtipoB = "A", "APORTES A     ", "DEDUCCIONES A ") & Trim(Cmbconcepto.Text) & " " & Txtano.Text
    Print #1, "ORDEN           NOMBRE                            SUELDO     INC. AFP    GRATIF.     VACACIO.     Liq/Subs  ESS.REM TOTAL     IMPORTE"
Case Else
    Print #1, CmbCia.Text & Space(10) & IIf(mtipoB = "A", "APORTES A     ", "DEDUCCIONES A ") & Trim(Cmbconcepto.Text) & " " & Txtano.Text
    Print #1, "ORDEN           NOMBRE                            SUELDO     INC. AFP    GRATIF.     VACACIO.     Liq/Subs  REM.TOTAL     IMPORTE"
End Select

Print #1, String(130, "-")
mlinea = 4
End Sub
Private Sub Procesa_Lista_Quinta_Anual()
Dim mItem As Integer
Dim mcad As String
Dim mcadIA As String
Dim mcadID As String
Dim cadnombre As String
Dim mtot As Currency
Dim mtipoB As String
Dim MUIT As Currency
Dim mtope As Currency
Dim mtopquin As Currency
Dim mFactor As Currency
Dim t1 As Currency: Dim t2 As Currency: Dim t3 As Currency: Dim t4 As Currency
Dim t5 As Currency: Dim t6 As Currency: Dim t7 As Currency: Dim t8 As Currency

cadnombre = nombre()
msemana = ""
mpag = 0
mtipoB = "D"
MUIT = 0
Sql$ = "select uit from plauit where cia='" & wcia & "' and ano='" & Txtano.Text & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then MUIT = rs!uit
If rs.State = 1 Then rs.Close

mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"


'Ingresos Afectos para Aportacion Normal
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso IN ('01') and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1

Do While Not rs.EOF
   Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PLACOD = rs!PLACOD
   dat.Recordset!nombre = mcad
   If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
   dat.Recordset!importe = rs!APO
   dat.Recordset!AFP = rs!AFP
   dat.Recordset!util = rs!util
   dat.Recordset!afp03 = rs!afp03
   If IsDate(rs2!fcese) Then
    dat.Recordset!CESE = Format(rs2!fcese, "dd/mm/yyyy")
   Else
    dat.Recordset!CESE = " "
   End If
'   If IsNull(rs2!FCESE) Then
'      dat.Recordset!cese = ""
'   Else
'      If IsDate(rs2!FCESE) Then dat.Recordset!cese = Format(RS!FCESE, "dd/mm/yyyy") Else dat.Recordset!cese = ""
'   End If
   dat.Recordset.Update
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Vacaciones
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"

'Ingresos Afectos para Aportacion Vacaciones
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='02'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*' "
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso IN ('02') and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*' "
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
      dat.Recordset!importe = rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset!afp03 = rs!afp03
      dat.Recordset!util = rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!CESE = Format(rs!fcese, "dd/mm/yyyy") Else dat.Recordset!CESE = ""
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!sueldo = dat.Recordset!sueldo + rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + rs!afp03
      dat.Recordset!util = dat.Recordset!util + rs!util
      dat.Recordset.Update
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Gratificacion
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"


'Ingresos Afectos para Aportacion Gratificacion
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='03'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab, " & mcad & " from plahistorico where cia='" & wcia & "' and proceso IN ('03') and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      dat.Recordset!importe = rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset!afp03 = rs!afp03
      dat.Recordset!util = rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!CESE = Format(rs!fcese, "dd/mm/yyyy") Else dat.Recordset!CESE = ""
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + rs!afp03
      'dat.Recordset!util = dat.Recordset!util + rs!AFP
      dat.Recordset!util = dat.Recordset!util
      dat.Recordset.Update
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

'Liquidaciones y Subsidios
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i18) as util "
Sql$ = "Select placod,tipotrab,sum(totaling-i18) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05','06') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"
If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*' "
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
      dat.Recordset!importe = rs!APO
      dat.Recordset!util = rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!CESE = Format(rs!fcese, "dd/mm/yyyy") Else dat.Recordset!CESE = ""
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!sueldo = dat.Recordset!sueldo + rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!util = dat.Recordset!util + rs!util
      dat.Recordset.Update
      
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
RUTA$ = App.path & "\REPORTS\" & "QtaAnual.txt"
Open RUTA$ For Output As #1
dat.Refresh
With dat.Recordset
If Not .EOF Then .MoveFirst
miten = 1
mtot = 0
Cabeza_QuInta_Anual
t1 = 0: t2 = 0: t3 = 0: t4 = 0
t5 = 0: t6 = 0: t7 = 0: t8 = 0

Do While Not .EOF
   'If Trim(!PLACOD) = "PO045" Then Stop
   mtope = Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) - Val(!afp03)
   mtopquin = Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) - Val(!afp03) - (7 * MUIT)
   mFactor = 0
   'Select Case mtope
   Select Case mtopquin
          Case Is < (Round(MUIT * 27, 2) + 1)
               'mFactor = Round(mtope * 0.15, 2)
               mFactor = Round(mtopquin * 0.15, 2)
          Case Is < (Round(MUIT * 54, 2) + 1)
               'mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
               mFactor = Round(((mtopquin - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
          Case Else
               'mFactor = Round(((mtope - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
               mFactor = Round(((mtopquin - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
    End Select
    If mFactor < 0 Then mFactor = 0
   mcad = !PLACOD & " " & !nombre & Space(1) & fCadNum(!sueldo, "####,###.00") & Space(1) & fCadNum(!util, "###,###.00") & Space(1) & fCadNum(!AFP, "###,###.00") & Space(1)
   mcad = mcad & fCadNum(!grati, "####,###.00") & Space(1) & fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util), "####,###.00") & Space(1) & fCadNum(Val(!afp03) * -1, "###,###.00") & Space(1)
   mcad = mcad & fCadNum(mtope, "####,###.00") & Space(1)
   mcad = mcad & fCadNum(MUIT * 7, "###,###.00") & Space(1) & fCadNum(mtope - (MUIT * 7), "###,###.00") & Space(1) & fCadNum(mFactor, "###,###.00") & Space(1) & fCadNum(!importe, "###,###.00") & Space(1) & fCadNum(!importe - mFactor, "##,###.00") & Space(2)
   mcad = mcad & IIf(!CESE = "01/01/1900", "", !CESE)
   Print #1, mcad
   t1 = t1 + !sueldo: t2 = t2 + !util: t3 = t3 + !AFP: t4 = t4 + !grati
   t5 = t5 + (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util)): t6 = t6 + (Val(!afp03) * -1): t7 = t7 + mtope
   t8 = t8 + !importe
   mItem = mItem + 1
   mlinea = mlinea + 1
   If mlinea > 55 Then Print #1, SaltaPag: Call Cabeza_QuInta_Anual
   .MoveNext
Loop
Print #1,
Print #1, Space(47) & fCadNum(t1, "####,###.00") & Space(1) & fCadNum(t2, "###,###.00") & Space(1) & fCadNum(t3, "###,###.00") & Space(1) & fCadNum(t4, "####,###.00") & Space(1) & fCadNum(t5, "####,###.00") & Space(1) & fCadNum(t6, "###,###.00") & Space(1) & fCadNum(t7, "####,###.00") & Space(34) & fCadNum(t8, "###,###.00")
End With
Close #1
Call Imprime_Txt("QtaAnual.txt", RUTA$)
dat.Database.Execute "delete from Tmpdeduc"
End Sub
Private Sub Cabeza_QuInta_Anual()
Print #1, LetraChica
Print #1, CmbCia.Text
Print #1, "REPORTE DE REMUNERACION ACUMULADA AL MES DE " & CmbMes.Text & " " & Txtano.Text
Print #1, String(195, "-")
Print #1, "CODIGO        NOMBRE                              REMUN.        UTIL.   INC. AFP   GRATIFIC.  ING. TOTAL     AFP 3%    REM.QTA.    7UIT    REM.AFECTA   QTA.CAT.   QTA.RET.     DIF.     F.CESE"
Print #1, String(195, "-")
mlinea = 4
End Sub
Private Sub Procesa_Certifica_Qta()
Dim mItem As Integer
Dim mcad As String
Dim mcadIA As String
Dim mcadID As String
Dim cadnombre As String
Dim mtot As Currency
Dim mtipoB As String
Dim MUIT As Currency
Dim mtope As Currency
Dim mFactor As Currency
Dim t1 As Currency: Dim t2 As Currency: Dim t3 As Currency: Dim t4 As Currency
Dim t5 As Currency: Dim t6 As Currency: Dim t7 As Currency: Dim t8 As Currency

cadnombre = nombre()
msemana = ""
mpag = 0
mtipoB = "D"
VConcepto = "13"
MUIT = 0
Sql$ = "select uit from plauit where cia='" & wcia & "' and ano='" & Txtano.Text & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then MUIT = rs!uit
If rs.State = 1 Then rs.Close

mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"


'Ingresos Afectos para Aportacion Normal
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1

Do While Not rs.EOF
   Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PLACOD = rs!PLACOD
   dat.Recordset!nombre = mcad
   If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
   dat.Recordset!importe = rs!APO
   dat.Recordset!AFP = rs!AFP
   dat.Recordset!util = rs!util
   dat.Recordset!afp03 = rs!afp03
   If IsDate(rs2!fcese) Then dat.Recordset!CESE = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!CESE = " "
   dat.Recordset.Update
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Vacaciones
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"

'Ingresos Afectos para Aportacion Vacaciones
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='02'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
      dat.Recordset!importe = rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset!afp03 = rs!afp03
      dat.Recordset!util = rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!CESE = Format(rs!fcese, "dd/mm/yyyy") Else dat.Recordset!CESE = ""
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!sueldo = dat.Recordset!sueldo + rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + rs!afp03
      dat.Recordset!util = dat.Recordset!util + rs!util
      dat.Recordset.Update
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
'Gratificacion
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"


'Ingresos Afectos para Aportacion Gratificacion
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='03'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   rs.MoveFirst
   mcadIA = ""
   Do While Not rs.EOF
      mcadIA = mcadIA & "i" & Trim(rs!cod_remu) & "+"
      rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If rs.State = 1 Then rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab, " & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
mItem = 1
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      dat.Recordset!importe = rs!APO
      dat.Recordset!AFP = rs!AFP
      dat.Recordset!afp03 = rs!afp03
      dat.Recordset!util = rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!CESE = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!CESE = Null
      dat.Recordset.Update
   Else
   '***************************************************************
   ' MODIFICA   : 10/01/2008
   ' MOTIVO     : ERROR EN LA ASIGNACION DE SUMA SE COMENTA EL CODIGO
   ' MODIFICADO POR : RICARDO HINOSTROZA
   '****************************************************************
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!grati = rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + rs!afp03
      dat.Recordset!util = dat.Recordset!util ' + rs!AFP
      dat.Recordset.Update
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

'Liquidaciones y Subsidios
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i18) as util "
Sql$ = "Select placod,tipotrab,sum(totaling-i18) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','06') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod" ','05'
If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst
Do While Not rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + rs!PLACOD + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PLACOD = rs!PLACOD
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = rs!afectoa
      dat.Recordset!importe = rs!APO
      dat.Recordset!util = rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!CESE = Format(rs!fcese, "dd/mm/yyyy") Else dat.Recordset!CESE = ""
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!sueldo = dat.Recordset!sueldo + rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + rs!APO
      dat.Recordset!util = dat.Recordset!util + rs!util
      dat.Recordset.Update
      
   End If
   rs.MoveNext
Loop
Dim IREM_NETA As Double

IREM_NETA = 0

If rs.State = 1 Then rs.Close
RUTA$ = App.path & "\REPORTS\" & "QtaAnual.txt"
Open RUTA$ For Output As #1
dat.Refresh
With dat.Recordset
If Not .EOF Then .MoveFirst
miten = 1
mtot = 0
t1 = 0: t2 = 0: t3 = 0: t4 = 0
t5 = 0: t6 = 0: t7 = 0: t8 = 0
IREM_NETA = 0
Do While Not .EOF

   '***************************************************************
   ' MODIFICA   : 10/01/2008
   ' MOTIVO     : ERROR EN LA RESTA DEL TOPE DE SUMA SE COMENTA EL CODIGO
   ' MODIFICADO POR : RICARDO HINOSTROZA
   '****************************************************************
   
   mtope = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) - Val(!afp03)) - ((MUIT * 7))
   IREM_NETA = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util)) - Val(!afp03)
   
   mFactor = 0
   Select Case mtope
          Case Is < (Round(MUIT * 27, 2) + 1)
               mFactor = Round((mtope) * 0.15, 2)  '- (muit * 7)
          Case Is < (Round(MUIT * 54, 2) + 1)
               mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
          Case Else
               mFactor = Round(((mtope - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
    End Select
    If mFactor < 0 Then mFactor = 0
    Print #1, Chr(218) & String(46, Chr(196)) & Chr(194) & String(31, Chr(196)) & Chr(191)
    Print #1, Chr(179) & "CERTIFICADO DE RETENCIONES SOBRE RENTAS DE    " & Chr(179) & "      EJERCICIO GRAVABLE       " & Chr(179)
    Print #1, Chr(179) & "5ta. CATEGORIA (Art. 45 del Reglamento del    " & Chr(179) & Space(31) & Chr(179)
    Print #1, Chr(179) & "Impuesto a la Renta (D.S No 122-94-EF)        " & Chr(179) & Space(13) & Txtano.Text & Space(14) & Chr(179)
    Print #1, Chr(192) & String(46, Chr(196)) & Chr(193) & String(31, Chr(196)) & Chr(217)
    Print #1,
    Print #1, Chr(218) & String(78, Chr(196)) & Chr(191)
    Print #1, Chr(179) & "Razon Social del Empleador: " & lentexto(32, Left(CmbCia.Text, 32)) & "   RUC " & wruc & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(34) & "CERTIFICA" & Space(35) & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "Que a Don (Doña) : " & lentexto(59, Left(!nombre, 59)) & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "se le ha retenido por concepto de Impuesto a la renta el importe (S/ " & fCadNum(!importe, "##,###.00") & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(194) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "1.- RENTAS BRUTAS" & Space(48) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Sueldo o Jornal Basico" & Space(39) & Chr(179) & fCadNum(!sueldo, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Comisiones            " & Space(39) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Clausula de Salvaguarda" & Space(38) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Participacion en las Utilidades" & Space(30) & Chr(179) & fCadNum(!util, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Incremento AFP" & Space(47) & Chr(179) & fCadNum(!AFP, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Gratificaciones" & Space(46) & Chr(179) & fCadNum(!grati, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "REMUNERACION BRUTA TOTAL    " & Space(37) & Chr(179) & fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util), "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "(-)AFP 3% (Art.71 DL 25897 6/12/92 Art.74 DS 054 97-ef 14/05/97) " & Chr(179) & fCadNum(!afp03, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    'Print #1, Chr(179) & "REMUNERACION NETA PARA QUINTA CATEGORIA" & Space(26) & Chr(179) & fCadNum(mtope, "#,###,###.00") & Chr(179)
    'IREM_NETA
    Print #1, Chr(179) & "REMUNERACION NETA PARA QUINTA CATEGORIA" & Space(26) & Chr(179) & fCadNum(IREM_NETA, "#,###,###.00") & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "2.- DEDUCCIONES SOBRE LA RENTA DE 5TA CATEGORIA" & Space(18) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    7 UIT (S/. " & fCadNum(MUIT, "###,###.00") & ")" & Space(39) & Chr(179) & fCadNum(MUIT * 7, "#,###,###.00") & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    'Print #1, Chr(179) & "3.RENTA NETA IMPONIBLE (1-2)" & Space(37) & Chr(179) & fCadNum(mtope - (MUIT * 7), "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & "3.RENTA NETA IMPONIBLE (1-2)" & Space(37) & Chr(179) & fCadNum(mtope, "#,###,###.00") & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "4.IMPUESTO A LA RENTA" & Space(44) & Chr(179) & fCadNum(mFactor, "#,###,###.00") & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "5.CREDITOS CONTRA EL IMPUESTO" & Space(36) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Credito por dividendos" & Space(39) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Credito por donaciones (tasa media sobre el monto computable)" & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "6.TOTAL DEL IMPUESTO RETENIDO" & Space(36) & Chr(179) & fCadNum(!importe, "#,###,###.00") & Chr(179)
    Print #1, Chr(192) & String(65, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(217)
    Print #1,
    Print #1,
    Print #1,
    Print #1, "  OBSERVACIONES" & Space(20) & String(25, Chr(196)) & Space(2) & String(15, Chr(196))
    Print #1, Space(36) & "FIRMA DEL REPRESENTANTE" & Space(7) & "RECIBIDO"
    Print #1, Space(46) & "LEGAL"
    Print #1,
  '  Print #1, "  * APLICABLES SOLO PARA TRABAJADORES CON RENTAS DE OTRAS CATEGORIA"
    Print #1, SaltaPag
   .MoveNext
Loop
End With
Close #1
Call Imprime_Txt("QtaAnual.txt", RUTA$)
dat.Database.Execute "delete from Tmpdeduc"
End Sub
Private Sub Procesa_CuadroIVF_TXT()
Dim mcad As String
Dim mcadprint As String
Dim totalano As Currency
Dim totalmes As Currency
Dim totalanon As Currency
Dim totalmesn As Currency
Dim tano As Currency
Dim tmes As Currency
Dim tanoporc As Currency
Dim tmesporc As Currency

Dim printyn As Boolean
If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Ingresar Tipo de Trabajador", vbInformation, "Estadisticos": Exit Sub
totalano = 0: totalmes = 0
mcad = "select sum(totaling)as ingreso,sum(totneto) as neto "
Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
If Not IsNull(rs!ingreso) Then totalanon = rs!neto: totalano = rs!ingreso
If rs.State = 1 Then rs.Close

Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
If Not IsNull(rs!ingreso) Then totalmes = rs!ingreso: totalmesn = rs!neto
If rs.State = 1 Then rs.Close


RUTA$ = App.path & "\REPORTS\" & "CuadroIVF.txt"
Open RUTA$ For Output As #1
Print #1, LetraChica
If mtipo = "01" Then
   Print #1, Space(41) & "CUADRO IV F - " & CmbMes.Text & " " & Txtano.Text & Space(42) & "Pag.  17"
   Print #1, Space(41) & "--------------------------------"
   Print #1, Space(36) & "RESUMEN MENSUAL DE PLANILLAS DE SUELDOS" & Space(33) & "No." & Format(CmbMes.ListIndex + 1, "00") & Txtano.Text
   Print #1, Space(36) & "--------------------------------------"
Else
   Print #1, Space(41) & "CUADRO IV A - " & CmbMes.Text & " " & Txtano.Text & Space(42) & "Pag.  12"
   Print #1, Space(41) & "--------------------------------------"
   Print #1, Space(36) & "RESUMEN MENSUAL DE PLANILLAS DE JORNALES" & Space(33) & "No." & Format(CmbMes.ListIndex + 1, "00") & Txtano.Text
   Print #1, Space(36) & "----------------------------------------"
End If
Print #1, "** Expresado en Nuevos Soles **"
Print #1, String(117, "-")
Print #1,
Print #1, "        CONCEPTO                 MENSUAL          %            ACUMULADO         %              PROMEDIO         %"
Print #1,
Print #1, String(117, "-")
Sql$ = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' order by codinterno"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
tano = 0: tmes = 0
Do While Not rs.EOF
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(I" & rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!ingreso) Then
      If rs2!ingreso <> 0 Then
         mcad = fCadNum(rs2!ingreso, "#,###,###,###.00") & "     " & fCadNum(rs2!ingreso * 100 / totalano, "###.00") & "     "
         mcad = mcad & fCadNum(rs2!ingreso / (CmbMes.ListIndex + 1), "#,###,###,###.00") & "     " & fCadNum(rs2!ingreso * 100 / totalano, "###.00") & " "
         tano = tano + rs2!ingreso
         tanoporc = tanoporc + (rs2!ingreso * 100 / totalano)
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(I" & rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!ingreso) Then
         mcad = fCadNum(rs2!ingreso, "#,###,###,###.00") & "     " & fCadNum(rs2!ingreso * 100 / totalmes, "###.00") & "     " & mcadprint
         tmes = tmes + rs2!ingreso
         tmesporc = tmesporc + (rs2!ingreso * 100 / totalmes)
      Else
        mcad = fCadNum(0, "#,###,###,###.00") & "     " & fCadNum(0, "###.00") & "     " & mcad
      End If
      mcad = lentexto(26, Left(rs!Descripcion, 26)) & mcad
   
      Print #1, mcad
      If rs2.State = 1 Then rs2.Close
   End If
   rs.MoveNext
Loop
Print #1,
Print #1, "                              ------------     ------         ------------     ------         ------------     -----"
Print #1, "  T O T A L" & Space(15) & fCadNum(tmes, "#,###,###,###.00") & "     " & fCadNum(tmesporc, "###.00") & "     " & fCadNum(tano, "#,###,###,###.00") & "     " & fCadNum(tmesporc, "###.00") & "     " & fCadNum(tano / (CmbMes.ListIndex + 1), "#,###,###,###.00") & "     " & fCadNum(tmesporc, "###.00")
Print #1,
Print #1,
Print #1, "Cuota Patronal"
Print #1, "--------------"
Print #1,
If rs.State = 1 Then rs.Close

'Aportaciones

Sql$ = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and status<>'*' order by codinterno"
If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
tano = 0: tmes = 0

rs.MoveFirst
Do While Not rs.EOF
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(A" & rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!ingreso) Then
      If rs2!ingreso <> 0 Then
         mcad = fCadNum(rs2!ingreso, "#,###,###,###.00") & "                "
         mcad = mcad & fCadNum(rs2!ingreso / (CmbMes.ListIndex + 1), "#,###,###,###.00") & "            "
         tano = tano + rs2!ingreso
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(A" & rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!ingreso) Then
         mcad = fCadNum(rs2!ingreso, "#,###,###,###.00") & "                 " & mcadprint
         tmes = tmes + rs2!ingreso
      Else
        mcad = fCadNum(0, "#,###,###,###.00") & "                " & mcad
      End If
      mcad = lentexto(26, Left(rs!Descripcion, 26)) & mcad
   
      Print #1, mcad
      If rs2.State = 1 Then rs2.Close
   End If
   rs.MoveNext
Loop

Print #1,
Print #1, "Descuento Personal"
Print #1, "------------------"
Print #1,

'Deducciones
rs.MoveFirst
Do While Not rs.EOF
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(D" & rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!ingreso) Then
      If rs2!ingreso <> 0 Then
         mcad = fCadNum(rs2!ingreso, "#,###,###,###.00") & "                "
         mcad = mcad & fCadNum(rs2!ingreso / (CmbMes.ListIndex + 1), "#,###,###,###.00") & "            "
         tano = tano + rs2!ingreso
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(D" & rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & CmbMes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!ingreso) Then
         mcad = fCadNum(rs2!ingreso, "#,###,###,###.00") & "                 " & mcadprint
         tmes = tmes + rs2!ingreso
      Else
        mcad = fCadNum(0, "#,###,###,###.00") & "                " & mcad
      End If
      mcad = lentexto(26, Left(rs!Descripcion, 26)) & mcad
   
      Print #1, mcad
      If rs2.State = 1 Then rs2.Close
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
tmes = tmes + totalmesn
tano = tano + totalanon
Print #1,
Print #1, "Neto Percibido Personal   " & fCadNum(totalmesn, "#,###,###,###.00") & Space(17) & fCadNum(totalanon, "#,###,###,###.00") & Space(16) & fCadNum(totalanon / (CmbMes.ListIndex + 1), "#,###,###,###.00")
Print #1, "                              ------------                    ------------                    ------------"
Print #1, "Gasto Total Empresa" & Space(7) & fCadNum(tmes, "#,###,###,###.00") & "                " & fCadNum(tano, "#,###,###,###.00") & "                " & fCadNum(tano / (CmbMes.ListIndex + 1), "#,###,###,###.00")
Print #1,
Print #1,
Print #1,
Print #1,
Print #1, Space(40) & "----------------" & Space(20) & "----------------"
Print #1, Space(40) & "   Hecho Por    " & Space(20) & "  Revisado Por  "
Close #1
Call Imprime_Txt("CuadroIVF.txt", RUTA$)
End Sub
Sub Imprimir_Titulo()
    Dim i_Contador_Titulo As Integer
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Space(70) & "Mes de " & CmbMes.Text & " del " & Txtano
    Print #1, ""
    For i_Contador_Titulo = 0 To 8
        Print #1, ""
    Next
End Sub
