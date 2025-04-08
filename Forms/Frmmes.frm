VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmmes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2250
   ClientLeft      =   2760
   ClientTop       =   795
   ClientWidth     =   6165
   Icon            =   "Frmmes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkOriginal 
      Caption         =   "Original"
      Height          =   195
      Left            =   2160
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox ChkResumen 
      Caption         =   "Resumen"
      Height          =   195
      Left            =   3720
      TabIndex        =   31
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CmbPlanta 
      Height          =   315
      ItemData        =   "Frmmes.frx":030A
      Left            =   1320
      List            =   "Frmmes.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ComboBox cbo_Cuenta 
      Height          =   315
      ItemData        =   "Frmmes.frx":030E
      Left            =   1200
      List            =   "Frmmes.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   2640
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.ComboBox CmbBcoPago 
      Height          =   315
      ItemData        =   "Frmmes.frx":0312
      Left            =   1200
      List            =   "Frmmes.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2280
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSDataListLib.DataCombo CboTrabajador 
      Height          =   315
      Left            =   1320
      TabIndex        =   18
      Top             =   1830
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.ComboBox Cmbtipbol 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1380
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Txtsemana 
      Height          =   285
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1380
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   60
         Width           =   4455
      End
   End
   Begin VB.ComboBox Cmbtipo 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Frmmes.frx":0316
      Left            =   1320
      List            =   "Frmmes.frx":0318
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   990
      Width           =   3255
   End
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   1335
      TabIndex        =   2
      Top             =   600
      Width           =   705
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "Frmmes.frx":031A
      Left            =   2340
      List            =   "Frmmes.frx":0342
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2235
   End
   Begin VB.Data dat 
      Caption         =   "Dta"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker FecProceso 
      Height          =   315
      Left            =   4680
      TabIndex        =   23
      Top             =   990
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   12615808
      CalendarTitleForeColor=   16777215
      Format          =   118751233
      CurrentDate     =   37616
   End
   Begin VB.Frame FrameEmple 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   6135
      Begin VB.OptionButton OptQuincena 
         BackColor       =   &H8000000C&
         Caption         =   "Quincena"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton OptBoleta 
         BackColor       =   &H8000000C&
         Caption         =   "Boleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frameconcepto 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   1320
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
   Begin Threed.SSPanel Panelprogress 
      Height          =   660
      Left            =   960
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   4005
      _Version        =   65536
      _ExtentX        =   7064
      _ExtentY        =   1164
      _StockProps     =   15
      ForeColor       =   4210752
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Label LblPlanta 
      AutoSize        =   -1  'True
      Caption         =   "Planta"
      Height          =   195
      Left            =   600
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   28
      Top             =   2685
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   2325
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      Caption         =   "Fecha a Procesar"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4680
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSForms.CheckBox Chk 
      Height          =   345
      Left            =   180
      TabIndex        =   19
      Top             =   1815
      Width           =   1050
      VariousPropertyBits=   1015031827
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1852;609"
      Value           =   "0"
      Caption         =   "Filtrar por"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Top             =   1380
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Lbltbol 
      AutoSize        =   -1  'True
      Caption         =   "T. Boleta"
      Height          =   195
      Left            =   600
      TabIndex        =   11
      Top             =   1380
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSForms.SpinButton SpinButton2 
      Height          =   300
      Left            =   4320
      TabIndex        =   10
      Top             =   1380
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
      Top             =   990
      Width           =   1125
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   300
      Left            =   2055
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
Dim mFecha As String
Dim mTipo As String
Dim mtipobol As String
Dim mlinea As Integer
Dim Vcadir As String
Dim Vcadir2 As String
Dim Vcadd As String
Dim Vcada As String
Dim Vcadh As String
Dim Rcadir As String
Dim Rcadir2 As String
Dim Rcadd As String
Dim Rcada As String
Dim Rcadh As String
Dim mpag As Integer
Dim mchartipo As String
Dim rs2 As ADODB.Recordset
Dim VConcepto As String
Dim PagEfect As Boolean
Dim VBcoPago As String
Dim mPlanta As String
Dim mano As Integer
Dim mmes As Integer
Dim msem As String
Dim numtra As Integer
Dim numtraCosto As Integer
Dim lBol As String
Dim lSem As String

'Dim rpt As Variant

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


Dim xlApp  As Object



Private Sub CmbBcoPago_Click()
Call Cargar_Cuenta_Banco
End Sub


Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
End Sub

Private Sub Cmbconcepto_Click()
VConcepto = fc_CodigoComboBox(CmbConcepto, 2)
End Sub

Private Sub Cmbtipbol_Click()
If Cmbtipbol.Text = "TOTAL" Then mtipobol = "" Else mtipobol = fc_CodigoComboBox(Cmbtipbol, 2)
FrameEmple.Visible = False
If Me.Caption = "BANCOS" And mTipo = "01" And mtipobol = "01" Then
   FrameEmple.Visible = True
   Me.Height = 2700
End If

End Sub

Private Sub CmbTipo_Click()
If Cmbtipo.Text = "TOTAL" Then mTipo = "" Else mTipo = fc_CodigoComboBox(Cmbtipo, 2)
If InStr(1, Me.Caption, "DEDUCCIONES Y APORTACIONES") > 0 Then
Else
If mTipo <> "01" And mTipo <> "" Then
   Lblsemana.Visible = True
   Txtsemana.Visible = True
   SpinButton2.Visible = True
   Txtsemana.Text = ""
   'Me.Height = 2370
   If CboTrabajador.Visible = True Then
        Me.Height = 2580
        CboTrabajador.Top = 1815
        Chk.Top = 1815
   End If
Else
   If Me.Caption = "RESUMEN DE PLANILLAS" Then
     ' Me.Height = 2205
   Else
      'Me.Height = 1845
   End If
   Lblsemana.Visible = False
   Txtsemana.Visible = False
   SpinButton2.Visible = False
   Txtsemana.Text = ""
   If CboTrabajador.Visible = True Then
        Me.Height = 2280
        CboTrabajador.Top = 1425
        Chk.Top = 1425
   End If
End If
End If

If CboTrabajador.Visible = True Then
    Call Trae_Trabajador
End If
FrameEmple.Visible = False
If Me.Caption = "BANCOS" And mTipo = "01" And mtipobol = "01" Then
   FrameEmple.Visible = True
   Me.Height = 2700
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 4860
Me.Height = 2370
Txtano.Text = Format(Year(Date), "0000")
Cmbmes.ListIndex = Month(Date) - 2
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Call fc_Descrip_Maestros2("01078", "", Cmbtipbol)
Call fc_Descrip_Maestros2("01007", "", CmbBcoPago, False)
Call fc_Descrip_Maestros2("01070", "", CmbPlanta)

Cmbtipbol.ListIndex = 0
Cmbmes.ListIndex = Month(Date) - 1
CboTrabajador.Visible = False
Chk.Visible = False
Cmbtipbol.Visible = False
Lbltbol.Visible = False
Lblsemana.Visible = False
Txtsemana.Visible = False
SpinButton2.Visible = False
Label4.Visible = False
CmbBcoPago.Visible = False
Label5.Visible = False
cbo_Cuenta.Visible = False
ChkOriginal.Visible = False
Select Case NameForm
       Case Is = "SEGURODEVIDA"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Me.Caption = "CALCULO DE SEGURO DE VIDA"
            Label2.Visible = False
            Cmbtipo.Visible = False
            Cmbmes.Visible = True
            Cmbtipo.AddItem "TOTAL"
        Case Is = "CUADROIV"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Me.Caption = "CUADRO IV"
            Label2.Visible = False
            Cmbtipo.Visible = False
            Cmbmes.Visible = True
            Cmbtipo.AddItem "TOTAL"
        Case Is = "RESUMEN"
            Me.Height = 2690
            Frameconcepto.Visible = False
            Cmbtipbol.Visible = True
            Lbltbol.Visible = True
            Me.Caption = "RESUMEN DE PLANILLAS"
            Cmbmes.Visible = True
            Cmbtipbol.AddItem "TOTAL"
            LblPlanta.Visible = True
            CmbPlanta.Visible = True
            CmbPlanta.AddItem "TOTAL"
            ChkResumen.Visible = True
        Case Is = "DEDUCAPOR"
            Me.Caption = "DEDUCCIONES Y APORTACIONES MENSUALES"
            Cmbtipo.AddItem "TOTAL"
            Frameconcepto.Visible = True
            Me.Height = 2370
            CmbConcepto.Clear
            Sql$ = "select codinterno,descripcion from placonstante where tipomovimiento='03' and status <>'*' AND cia='" & wcia & "' order by descripcion"
            Call rCarCbo(CmbConcepto, Sql$, "C", "00")
            Cmbmes.Visible = True
        Case Is = "CUADROIV"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Me.Caption = "ESTADISTICO - HORAS TRABAJADAS EFECTIVAS"
            Cmbmes.Visible = True
        Case Is = "SEGURO"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Cmbtipo.Visible = False
            Label2.Visible = False
            ChkOriginal.Visible = True
           ' ChkOriginal.Enable = True
            Me.Caption = "SEGURO RIESGO SALUD"
        Case Is = "SNPESSALUD"
            Me.Height = 1845
            Frameconcepto.Visible = False
            Cmbtipo.Visible = False
            Label2.Visible = False
            Me.Caption = "SNP - ESSALUD"
        Case Is = "REMUNERA"
            Me.Caption = "REMUNERACIONES"
            Cmbtipo.AddItem "TOTAL"
            Frameconcepto.Visible = True
            Me.Height = 2370
            CmbConcepto.Clear
            Sql$ = "select codinterno,descripcion from placonstante where tipomovimiento='02' and status <>'*' AND CIA='" & wcia & "' order by descripcion"
            Call rCarCbo(CmbConcepto, Sql$, "C", "00")
            Cmbmes.Visible = True
        Case Is = "DEDUCAPORANUAL"
            Me.Caption = "DEDUCCIONES Y APORTACIONES ANUALES"
            Cmbtipo.AddItem "TOTAL"
            Frameconcepto.Visible = True
            Me.Height = 2370
            CmbConcepto.Clear
            Sql$ = "select codinterno,descripcion from placonstante where tipomovimiento='03' and status <>'*' AND cia='" & wcia & "' order by descripcion"
            Call rCarCbo(CmbConcepto, Sql$, "C", "00")
            Cmbmes.Visible = False
        Case Is = "CERTIFICAQTA"
            Me.Caption = "CERTIFICADOS DE QTA CATEGORIA"
            Cmbmes.ListIndex = 11
            Cmbmes.Visible = False
            Me.Height = 2280
            CboTrabajador.Top = 1425
            CboTrabajador.Visible = True
            Chk.Top = 1425
            Chk.Visible = True
        Case Is = "BANCOS"
            Me.Caption = "BANCOS"
            'Label4.Visible = True
            'CmbBcoPago.Visible = True
            'Label5.Visible = True
            'cbo_Cuenta.Visible = True
            'Me.Height = 3600
            Me.Width = 6255
            Frameconcepto.Visible = False
            Cmbtipbol.Visible = True
            Lbltbol.Visible = True
            Cmbmes.Visible = True
            FecProceso.Visible = True
            Lblfecha.Visible = True
            FecProceso.Value = Date
End Select
Crea_Tablas
End Sub

Private Sub Trae_Trabajador()
    Dim mTipo As String
    mTipo = Trim(fc_CodigoComboBox(Me.Cmbtipo, 2))
    If mTipo = "99" Then mTipo = Empty
    Cadena = "SP_TRAE_TRABAJADOR '" & wcia & "'," & _
            "'" & Trim(mTipo) & "'"
    Set Ors = New ADODB.Recordset
    Ors.CursorLocation = adUseClient
    Ors.Open Cadena, cn, adOpenStatic, adLockReadOnly
    If Not Ors.EOF Then
        With Me.CboTrabajador
            Set .RowSource = Ors
            .ListField = "NOMBRE"
            .DataField = "NOMBRE"
            .BoundColumn = "CODIGO"
        End With
    End If
    Me.CboTrabajador.Refresh
    Me.CboTrabajador.BoundText = "99999"
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

If Cmbmes.ListIndex < 0 Then MsgBox "Debe Seleccionar Mes del Periodo", vbInformation, "Consultas y Reportes": Exit Sub
If Val(Txtano) < 1900 Or Val(Txtano) > 9999 Then MsgBox "Indique correctamente el Año del Periodo", vbInformation, "Consultas y Reportes": Exit Sub
mdia = Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text))
mFecha = Format(mdia, "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Format(Val(Txtano.Text), "0000")
Select Case NameForm
       Case Is = "SEGURODEVIDA"
            Procesa_Seguro_Vida
       Case Is = "CUADROIV"
            Procesa_Cuadro_IV
       Case Is = "RESUMEN"
            Resumen_Planilla
       Case Is = "DEDUCAPOR"
            If CmbConcepto.ListIndex < 0 Then
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
       Case Is = "SEGURO"
            If ChkOriginal.Value = 0 Then
               Procesa_Seguro (1)
            Else
               Procesa_Seguro (2)
            End If
       Case Is = "SNPESSALUD"
            procesa_snpessalud
       Case Is = "REMUNERA"
            Procesa_Remunera
       Case Is = "DEDUCAPORANUAL"
            If CmbConcepto.ListIndex < 0 Then
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
       Case Is = "BANCOS"
             Archivo_Bancos
 End Select
End Sub
Private Sub Archivo_Bancos_TXT()

mmes = Cmbmes.ListIndex + 1
mano = Val(Txtano.Text)
msem = Txtsemana.Text
If Me.CmbBcoPago.ListIndex < 0 Then
    MsgBox "Seleccione el Banco", vbCritical, Me.Caption
    Exit Sub
End If
If Me.cbo_Cuenta.ListIndex < 0 Then
    MsgBox "Seleccione la cuenta del banco", vbCritical, Me.Caption
    Exit Sub
End If

PagEfect = False
Sql$ = "select distinct(p.pagobanco) "
If FrameEmple.Visible And OptQuincena.Value Then Sql$ = Sql$ & "from plaquincena q,planillas p " Else Sql$ = Sql$ & "from plahistorico q,planillas p "
Sql$ = Sql$ & "where q.cia='" & wcia & "' and year(q.fechaproceso)=" & mano & " and month(q.fechaproceso)=" & mmes & " "
If mtipobol = "01" Then Sql$ = Sql$ & "and q.proceso IN('01','05') " Else Sql$ = Sql$ & "and proceso LIKE '" & Trim(mtipobol) + "%" & "' "
If mTipo <> "01" And mTipo <> "" Then Sql$ = Sql$ & "and semana LIKE '" & Trim(msem) + "%" & "' "
Sql$ = Sql$ & "and q.status<>'*' and p.status<>'*' and p.cia=q.cia and p.placod=q.placod and p.pagobanco='" & VBcoPago & "'"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
Do While Not Rs.EOF
   If Trim(Rs(0) & "") <> "" Then Call Procesa_Archivo_Banco(Rs(0), mano, mmes, msem)
   Rs.MoveNext
Loop
Rs.Close: Set Rs = Nothing
If PagEfect Then Call Procesa_Archivo_Banco("**", mano, mmes, msem)
MsgBox "Se Generaron los Archivos", vbInformation
End Sub
Private Sub Procesa_Archivo_Banco_Txt(banco As String, mano As Integer, mmes As Integer, msem As String)
Dim March As String
Dim nFil As Integer
Dim lDesTipoBol As String
Dim rs2 As ADODB.Recordset
Dim lCta As String

If FrameEmple.Visible Then
   If OptBoleta.Value = False And OptQuincena.Value = False Then
      MsgBox "Indique Si es Quincena o Fin de Mes", vbInformation
      Exit Sub
   End If
End If
If banco = "**" Then
   March = "PHEFCT.XLS"
Else
   March = "PH" & wcia & banco & ".XLS"
   
   'Archivo de Texto
   Dim mArchPla As String
   Dim lCad As String
   Dim lCheckSum As Double
   lCheckSum = 0
   lCta = ""
   
   mArchPla = "PH" & wcia & Format(mano, "0000") & Format(mmes, "00")
   If Txtsemana.Visible = True Then
      mArchPla = mArchPla & Txtsemana.Text
   ElseIf FrameEmple.Visible Then
      If OptQuincena.Value = True Then mArchPla = mArchPla & "QE" Else mArchPla = mArchPla & "BE"
   End If
   mArchPla = mArchPla & ".txt"
   
   RUTA$ = App.Path & "\REPORTS\" & mArchPla
   Open RUTA$ For Output As #1
   
   
   Sql$ = "select cuentabco from PlaBcoCta where cia='" & wcia & "' and status<>'*'"
   If fAbrRst(rs2, Sql$) Then lCta = Trim(rs2(0) & "")
   rs2.Close
   If Len(lCta) < 14 Then lCta = Mid(lCta, 1, 3) & Llenar_Ceros(Mid(lCta, 4, Len(lCta) - 6), 8) & Right(lCta, 3)
   lCad = "#1HC" & lCta & Space(6) & "S/"
   
   lCheckSum = Val(Mid(lCta, 4, 15))
   
End If

mOrigen$ = Path_Reports & March

Set fso = CreateObject("Scripting.FileSystemObject")

If Not (fso.FileExists(mOrigen$)) Then
    If banco <> "**" Then Close #1
    MsgBox "El archivo " & March & Chr(13) & "No se encuentra"
    Exit Sub
End If

Mdestino$ = App.Path & "\Reports\" & March
CopyFile mOrigen, Mdestino, FILE_NOTIFY_CHANGE_LAST_WRITE

Sql$ = "select distinct(p.pagobanco) "

Sql$ = "select q.semana,q.proceso,q.placod,q.fechaproceso," _
     & "(select tipo_doc from planillas where placod=q.placod and status<>'*') as td," _
     & "(select nro_doc from planillas where placod=q.placod and status<>'*') as nd," _
     & "(select AP_PAT from planillas where placod=q.placod and status<>'*') as APEP," _
     & "(select AP_MAT from planillas where placod=q.placod and status<>'*') as APEM," _
     & "(select LTRIM(RTRIM(NOM_1))+' '+LTRIM(RTRIM(NOM_2)) from planillas where placod=q.placod and status<>'*') as NOM, " _
     & "(select pagonumcta from planillas where placod=q.placod and status<>'*') as Cuenta," _
     & "q.totneto,q.Proceso,q.PlaCod "

If FrameEmple.Visible And OptQuincena.Value Then Sql$ = Sql$ & "from plaquincena q,planillas p " Else Sql$ = Sql$ & "from plahistorico q,planillas p "
Sql$ = Sql$ & "where q.cia='" & wcia & "' and year(q.fechaproceso)=" & mano & " and month(q.fechaproceso)=" & mmes & " "

If mtipobol = "03" Then
   If banco = "**" Then Sql$ = Sql$ & "and q.proceso='03' " Else Sql$ = Sql$ & "and proceso LIKE '" & Trim(mtipobol) + "%" & "' "
Else
   If banco = "**" Then Sql$ = Sql$ & "and q.proceso IN('01','05') " Else Sql$ = Sql$ & "and proceso LIKE '" & Trim(mtipobol) + "%" & "' "
End If


If mTipo <> "01" And mTipo <> "" Then Sql$ = Sql$ & "and semana LIKE '" & Trim(msem) + "%" & "' "
If banco <> "**" Then Sql$ = Sql$ & "and pagobanco='" & banco & "' "
Sql$ = Sql$ & "and p.tipotrabajador='" & fc_CodigoComboBox(Cmbtipo, 2) & "' and q.status<>'*' and p.status<>'*' and p.cia=q.cia and p.placod=q.placod"

If Not (fAbrRst(rs2, Sql$)) Then rs2.Close: Set rs2 = Nothing: Exit Sub
rs2.MoveFirst

Set xlApp = GetObject(App.Path & "\Reports\" & March)

Set xlApp2 = xlApp.Application
For I = 1 To xlApp2.Workbooks.count
    If xlApp2.Workbooks(I).Name = March Then
       Set xlBook = xlApp2.Workbooks(I)
       xlApp2.Workbooks(I).Activate
       numwindows = I
       GoTo Continua
    End If
Next I

Continua:

xlApp2.Application.Visible = True
xlApp2.Parent.Windows(March).Visible = True

If banco = "**" Then
   Set xlSheet = xlApp2.Worksheets("Efectivo")
Else
   Set xlSheet = xlApp2.Worksheets("Deposito")
End If
xlSheet.Activate

If banco = "**" Then
   xlSheet.Cells(1, 4).Value = Cmbcia.Text
   If Trim(rs2!semana & "") <> "" Then xlSheet.Cells(2, 2).Value = "SEMANA No. ": xlSheet.Cells(2, 4).Value = "'" & Trim(rs2!semana & "")
   xlSheet.Cells(1, 8).Value = "'" & Format(Day(rs2!FechaProceso), "00") & "/" & Format(Month(rs2!FechaProceso), "00") & "/" & Format(Year(rs2!FechaProceso), "0000")
End If
If banco = "**" Then nFil = 8 Else nFil = 37

'Total Planilla
Dim lCount As Integer
lCount = 0
If banco <> "**" Then
   Dim lTotPla As Double
   lTotPla = 0
   Do While Not rs2.EOF
      If Trim(rs2!Cuenta & "") <> "" And rs2!Proceso <> "05" Then
         lTotPla = lTotPla + rs2!totneto
         lCta = Replace(Trim(rs2!Cuenta & ""), "-", "")
         lCheckSum = lCheckSum + Val(Mid(lCta, 4, 15))
         lCount = lCount + 1
       End If
      rs2.MoveNext
   Loop
   lCad = lCad & Format(Int(lTotPla), "0000000000000") & Format((lTotPla - Int(lTotPla)) * 100, "00")
   lCad = lCad & Format(FecProceso.Day, "00") & Format(FecProceso.Month, "00") & Format(FecProceso.Year, "0000")
   lCad = lCad & Space(20)
   lCad = lCad & Llenar_Ceros(Trim(Str(lCheckSum)), 15)
   lCad = lCad & Format(lCount, "000000") & "1" & Space(15) & "1"
   Print #1, lCad
End If
'Fin Total Planilla


If rs2.RecordCount > 0 Then rs2.MoveFirst
Do While Not rs2.EOF
   If banco = "**" Then
      lDesTipoBol = ""
      If rs2!Proceso = "01" Then lDesTipoBol = "Normal"
      If rs2!Proceso = "09" Then lDesTipoBol = "VACACIONES PROVISIONADAS"
      If rs2!Proceso = "10" Then lDesTipoBol = "TRANSFERENCIA"
      If rs2!Proceso = "04" Then lDesTipoBol = "LIQUIDACION"
      If rs2!Proceso = "05" Then lDesTipoBol = "SUBSIDIO"
      If rs2!Proceso = "07" Then lDesTipoBol = "DEPOSITO CTS"
      If rs2!Proceso = "08" Then lDesTipoBol = "VACACIONES PAGADAS"
      If rs2!Proceso = "02" Then lDesTipoBol = "VACACIONES"
      If rs2!Proceso = "03" Then lDesTipoBol = "GRATIFICACION"
   
      If Trim(rs2!Cuenta & "") = "" Or rs2!Proceso = "05" Then
         xlSheet.Cells(nFil, 2).Value = nFil - 7
         If Trim(rs2!td & "") = "01" Then xlSheet.Cells(nFil, 3).Value = "'" & Trim(rs2!nd & "")
         xlSheet.Cells(nFil, 4).Value = Trim(rs2!PlaCod & "")
         xlSheet.Cells(nFil, 5).Value = Trim(rs2!apep & "")
         xlSheet.Cells(nFil, 6).Value = Trim(rs2!apem & "")
         xlSheet.Cells(nFil, 7).Value = Trim(rs2!NOM & "")
         xlSheet.Cells(nFil, 8).Value = rs2!totneto
         xlSheet.Cells(nFil, 9).Value = lDesTipoBol
         nFil = nFil + 1
      End If
   Else
      If Trim(rs2!Cuenta & "") <> "" And rs2!Proceso <> "05" Then
         If Trim(rs2!td & "") = "01" Then
            xlSheet.Cells(nFil, 5).Value = "1"
            xlSheet.Cells(nFil, 6).Value = Trim(rs2!nd & "")
         End If
         xlSheet.Cells(nFil, 8).Value = Trim(rs2!apep & "")
         xlSheet.Cells(nFil, 9).Value = Trim(rs2!apem & "")
         xlSheet.Cells(nFil, 10).Value = Trim(rs2!NOM & "")
         xlSheet.Cells(nFil, 11).Value = Trim(rs2!Cuenta & "")
         xlSheet.Cells(nFil, 13).Value = rs2!totneto
         lCad = " 2A"
         lCta = Replace(Trim(rs2!Cuenta & ""), "-", "")
         If Len(lCta) < 14 Then lCta = Mid(lCta, 1, 3) & Llenar_Ceros(Mid(lCta, 4, Len(lCta) - 6), 8) & Right(lCta, 3)
         lCad = lCad & lCta & Space(6)
         lCta = Trim(rs2!apep & "") & " " & Trim(rs2!apem & "") & " " & Trim(rs2!NOM & "")
         lCta = Replace(Trim(lCta & ""), "Ñ", "N")
         lCad = lCad & lentexto(40, lCta) & "S/"
         lCad = lCad & Format(Int(rs2!totneto), "0000000000000") & Format((rs2!totneto - Int(rs2!totneto)) * 100, "00")
         lCad = lCad & Space(40) & "0"
         
         If Trim(rs2!td & "") = "01" Then
            lCad = lCad & "DNI"
            lCad = lCad & lentexto(12, Trim(rs2!nd & ""))
         Else
            lCad = lCad & Space(3 + 12)
         End If
         lCad = lCad & "0"
         Print #1, lCad
         nFil = nFil + 1
      Else
         PagEfect = True
      End If
   End If
   rs2.MoveNext
Loop
xlApp2.Parent.Windows(numwindows).WindowState = xlMaximized
xlApp2.Parent.Windows(numwindows).Visible = True

If banco <> "**" Then Close #1

End Sub
Private Sub Procesa_Seguro_Vida()

Dim FecFin As String
Dim FecIni As String
Dim mtc As Double
mtc = InputBox("Ingrese Tipo de Cambio", "Cuadro IV")

If Not IsNumeric(mtc) Then
   MsgBox "Ingrese Correctamente el tipo de cambio", vbInformation
   Exit Sub
End If

FecFin = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "DD/MM/YYYY")
FecIni = DateAdd("m", -5, CDate("01/" & Month(FecFin) & "/" & Year(FecFin)))

Dim Rs As ADODB.Recordset
Dim nFil As Integer

'Sql = "Delete from Pla_Calculo_Seguro_Vida where "
Sql = "update Pla_Calculo_Seguro_Vida set Status='*' where Año=" & Year(FecFin) & " and Mes='" & Month(FecFin) & "'"
cn.Execute Sql


Sql = "usp_Pla_Seguro_Vida '" & wcia & "','01','" & FecIni & "','" & FecFin & "'," & mtc & ""
cn.Execute Sql
Sql = "usp_Pla_Seguro_Vida '" & wcia & "','02','" & FecIni & "','" & FecFin & "', " & mtc & ""
cn.Execute Sql

Call Menos_3("01", mtc)
Call Menos_3("02", mtc)
Detalle_seg_Vida ("01")
Detalle_seg_Vida ("02")
Formato_seg_Vida
Formato_seg_VidaII
End Sub
Private Sub Formato_seg_VidaII()

Dim RqF As ADODB.Recordset
Sql = "usp_pla_Seguro_vida_Formato"
If (fAbrRst(RqF, Sql)) Then RqF.MoveFirst Else RqF.Close: Set RqF = Nothing: Exit Sub


Set xlSheet = xlApp2.Worksheets("HOJA6")
xlSheet.Name = "FORM2E"
   
xlSheet.Range("K:K").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Range("F:G").NumberFormat = "dd/mm/yyyy;@"


xlSheet.Cells(3, 1).Value = "CALCULO DE SEGURO DE VIDA - " & Cmbmes.Text & " " & Txtano.Text
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


xlSheet.Cells(3, 1).Value = "CALCULO DE SEGURO DE VIDA - " & Cmbmes.Text & " " & Txtano.Text
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
   xlApp2.Application.ActiveWindow.DisplayGridlines = False
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

Private Sub Formato_seg_Vida()

Dim RqF As ADODB.Recordset
Sql = "usp_pla_Seguro_vida_Formato"
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
Private Sub Formato_seg_Vida_Hasta_MArzo_2016()

Dim RqF As ADODB.Recordset
Sql = "usp_pla_Seguro_vida_Formato"
If (fAbrRst(RqF, Sql)) Then RqF.MoveFirst Else RqF.Close: Set RqF = Nothing: Exit Sub

Set xlSheet = xlApp2.Worksheets("HOJA5")
xlSheet.Name = "FORM"
   
xlSheet.Range("AA:AA").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
xlSheet.Range("N:N").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("U:U").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("V:V").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("AC:AC").NumberFormat = "dd/mm/yyyy;@"

xlSheet.Cells(1, 1).Value = "Razon Social"
xlSheet.Cells(1, 2).Value = "RUC"
xlSheet.Cells(1, 3).Value = "Giro/Actividad"
xlSheet.Cells(1, 4).Value = "Email Contacto"
xlSheet.Cells(1, 5).Value = "Telefono"
xlSheet.Cells(1, 6).Value = "Departamento"
xlSheet.Cells(1, 7).Value = "Provincia"
xlSheet.Cells(1, 8).Value = "Distrito"
xlSheet.Cells(1, 9).Value = "Dirección"
xlSheet.Cells(1, 10).Value = "Proyecto/Obra"
xlSheet.Cells(1, 11).Value = "DNI-Representante Legal"
xlSheet.Cells(1, 12).Value = "REPRES. LEGAL"
xlSheet.Cells(1, 13).Value = "REPRES. CARGO"
xlSheet.Cells(1, 14).Value = "REPRES. FEC NACIMIENTO"
xlSheet.Cells(1, 15).Value = "Nivel de Riesgo"
xlSheet.Cells(1, 16).Value = "DNI / CE"
xlSheet.Cells(1, 17).Value = "Apellido Paterno"
xlSheet.Cells(1, 18).Value = "Apellido Materno"
xlSheet.Cells(1, 19).Value = "Nombres"
xlSheet.Cells(1, 20).Value = "cod Trabajador"
xlSheet.Cells(1, 21).Value = "fecha Nac"
xlSheet.Cells(1, 22).Value = "Fecha Ingreso a la Empresa"
xlSheet.Cells(1, 23).Value = "Fech Ingreso al Seguro"
xlSheet.Cells(1, 24).Value = "Nacionalidad"
xlSheet.Cells(1, 25).Value = "Ocupación , Profesión"
xlSheet.Cells(1, 26).Value = "Moneda del Sueldo"
xlSheet.Cells(1, 27).Value = "Sueldo"
xlSheet.Cells(1, 28).Value = "Tipo de Movimiento"
xlSheet.Cells(1, 29).Value = "Fecha de Inicio de Vigencia"
xlSheet.Cells(1, 30).Value = "Moneda de Prima"

xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 30)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 30)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 30)).Borders.LineStyle = xlContinuous

nFil = 2

Do While Not RqF.EOF
   xlSheet.Cells(nFil, 1).Value = Trim(RqF!empresa & "")
   xlSheet.Cells(nFil, 2).Value = Trim(RqF!RUC & "")
   xlSheet.Cells(nFil, 3).Value = Trim(RqF!giro & "")
   xlSheet.Cells(nFil, 4).Value = Trim(RqF!Email & "")
   xlSheet.Cells(nFil, 5).Value = Trim(RqF!telefono & "")
   xlSheet.Cells(nFil, 6).Value = Trim(RqF!departamento & "")
   xlSheet.Cells(nFil, 7).Value = Trim(RqF!provincia & "")
   xlSheet.Cells(nFil, 8).Value = Trim(RqF!DISTRITO & "")
   xlSheet.Cells(nFil, 9).Value = Trim(RqF!Direccion & "")
   xlSheet.Cells(nFil, 10).Value = ""
   xlSheet.Cells(nFil, 11).Value = "'" & Trim(RqF!DNIRep & "")
   xlSheet.Cells(nFil, 12).Value = Trim(RqF!RepLegal & "")
   xlSheet.Cells(nFil, 13).Value = Trim(RqF!Cargo & "")
   xlSheet.Cells(nFil, 14).Value = RqF!FecNacRep
   xlSheet.Cells(nFil, 15).Value = ""
   xlSheet.Cells(nFil, 16).Value = "'" & Trim(RqF!DniTrab & "")
   xlSheet.Cells(nFil, 17).Value = Trim(RqF!Ape_pat & "")
   xlSheet.Cells(nFil, 18).Value = Trim(RqF!Ape_mat & "")
   xlSheet.Cells(nFil, 19).Value = Trim(RqF!nombres & "")
   xlSheet.Cells(nFil, 20).Value = Trim(RqF!PlaCod & "")
   xlSheet.Cells(nFil, 21).Value = RqF!fnacimiento
   xlSheet.Cells(nFil, 22).Value = RqF!fIngreso
   xlSheet.Cells(nFil, 23).Value = ""
   xlSheet.Cells(nFil, 24).Value = Trim(RqF!nacionalidad & "")
   xlSheet.Cells(nFil, 25).Value = Trim(RqF!Cargo & "")
   xlSheet.Cells(nFil, 26).Value = Trim(RqF!Monedasueldo & "")
   xlSheet.Cells(nFil, 27).Value = RqF!sueldo
   xlSheet.Cells(nFil, 28).Value = Trim(RqF!tipo & "")
   xlSheet.Cells(nFil, 29).Value = ""
   xlSheet.Cells(nFil, 30).Value = Trim(RqF!MonPrima & "")
   nFil = nFil + 1
   RqF.MoveNext
Loop

RqF.Close: Set RqF = Nothing
xlSheet.Range("A:AD").EntireColumn.AutoFit

End Sub

Private Sub Detalle_seg_Vida(lTipoTrab As String)

Dim Sql As String
Dim FecFin As String
FecFin = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "mm/dd/YYYY")
Sql = "Select placod,nombre,(select descripcion from pla_ccostos where cia='01' and codigo=Pla_Calculo_Seguro_Vida.area and status<>'*') as CCosto,"
Sql = Sql & "Basico , PromExt, PromOtros, Total "
''Sql = Sql & "From Pla_Calculo_Seguro_Vida where ( year(fingreso) <> " & Val(Txtano.Text) & "And month(fingreso) <> (" & Cmbmes.ListIndex + 1 & ")) and tipotrab='" & lTipoTrab & "' order by ccosto,placod"
Sql = Sql & "From Pla_Calculo_Seguro_Vida where tipotrab='" & lTipoTrab & "' order by ccosto,placod"

'Sql = "usp_Pla_Seguro_Vida '" & wcia & "','02','" & FecIni & "','" & FecFin & "', " & mtc & ""
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

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
   xlSheet.Cells(3, 2).Value = "TOTAL PLANILLA SUELDOS MES DE " & Cmbmes.Text & " - " & Txtano.Text
Else
   xlSheet.Cells(3, 2).Value = "TOTAL PLANILLA SALARIOS MES DE " & Cmbmes.Text & " - " & Txtano.Text
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
mCCosto = Trim(Rs!ccosto & "")
Do While Not Rs.EOF
   If Rs!Total <> 0 Then
      If mCCosto <> "" And mCCosto <> Trim(Rs!ccosto & "") Then
         msum = msum * -1
         nFil = nFil + 1
         xlSheet.Cells(nFil, 4).Value = "TOTAL " & mCCosto
         xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
         xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 8)).Borders.LineStyle = xlContinuous
         msum = 1: nFil = nFil + 2
         mCCosto = Trim(Rs!ccosto & "")
      End If
      
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = "'" & Trim(Rs!ccosto & "")
      xlSheet.Cells(nFil, 5).Value = Rs!Basico
      xlSheet.Cells(nFil, 6).Value = Rs!PROMEXT
      xlSheet.Cells(nFil, 7).Value = Rs!PromOtros
      xlSheet.Cells(nFil, 8).Value = Rs!Total
      lTotBasico = lTotBasico + Rs!Basico
      lTotExtras = lTotExtras + Rs!PROMEXT
      lTotOtros = lTotOtros + Rs!PromOtros
      lTotTot = lTotTot + Rs!Total
      nFil = nFil + 1
      msum = msum + 1
   End If
   mTotTrab = mTotTrab + 1
   Rs.MoveNext
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

Private Sub Menos_3(lTipoTrab As String, mtc As Double)

Dim Sql As String

Dim mFactorE As Double
Dim mIgv As Double


If lTipoTrab = "01" Then mFactorE = 0.33 Else mFactorE = 0.41
mIgv = 18
Sql = "Select S.*,fnacimiento,"
Sql = Sql & "Case tipo_doc when '01' then nro_doc else '' End as Dni,"
Sql = Sql & "(select cuenta from maestros_31 where ciamaestro='01055' and cod_maestro3=p.cargo) as Cargo,"
Sql = Sql & "(select descripcion from pla_ccostos where cia='01' and codigo=s.area and status<>'*') as CCosto "
Sql = Sql & "From Pla_Calculo_Seguro_Vida S,Planillas P "
'Sql = Sql & "Where not ( year(s.fingreso) = " & Val(Txtano.Text) & "And month(s.fingreso) = (" & Cmbmes.ListIndex + 1 & ")) and "
Sql = Sql & "Where s.TipoTrab='" & lTipoTrab & "' and p.cia='01' and p.status<>'*' and p.placod=s.placod order by s.placod"

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

Dim lTotPlanilla As Double
Dim lTotPasaTope As Double
Dim lTotMenos As Double
Dim lNumTrabPasaTope As Integer

lTotPlanilla = 0: lTotPasaTope = 0: lTotMenos = 0: lNumTrabPasaTope = 0
Do While Not Rs.EOF
    'If (Year(rs!fIngreso) <> " & Val(Txtano.Text) & " And Month(rs!fIngreso) <> (" & Cmbmes.ListIndex + 1 & ")) Then
    'If Trim(rs!PlaCod & "") <> "E0520" Then
        lTotPlanilla = lTotPlanilla + Rs!Total
        If Trim(Rs!tope & "") = "S" Then lTotPasaTope = lTotPasaTope + Rs!Total: lNumTrabPasaTope = lNumTrabPasaTope + 1
        If Trim(Rs!Pasa & "") <> "S" Then lTotMenos = lTotMenos + Rs!Total
    'End If
   Rs.MoveNext
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

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

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
Do While Not Rs.EOF
   If Trim(Rs!Pasa & "") <> "S" Then
      xlSheet.Cells(nFil, 1).Value = msum
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = "'" & Trim(Rs!DNI & "")
      xlSheet.Cells(nFil, 5).Value = Rs!fnacimiento
      xlSheet.Cells(nFil, 5).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 6).Value = Rs!fIngreso
      xlSheet.Cells(nFil, 6).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 7).Value = Trim(Rs!Cargo & "")
      xlSheet.Cells(nFil, 8).Value = Rs!Total * -1
      xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
      xlSheet.Cells(nFil, 9).Value = Trim(Rs!ccosto & "")
      nFil = nFil + 1
      msum = msum + 1
   End If
   mTotTrab = mTotTrab + 1
   Rs.MoveNext
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
If Rs.RecordCount > 1 Then Rs.MoveFirst
Do While Not Rs.EOF
   If Trim(Rs!tope & "") = "S" And (Month(Rs!fIngreso) <> (Cmbmes.ListIndex + 1)) Then
   'And Year(rs!fIngreso) <> Val(Txtano.Text)) Then
     'If Trim(rs!tope & "") = "S" And Trim(rs!PlaCod) <> "E0520" Then
       xlSheet.Cells(nFil, 1).Value = msum
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = "'" & Trim(Rs!DNI & "")
      xlSheet.Cells(nFil, 5).Value = Rs!fnacimiento
      xlSheet.Cells(nFil, 5).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 6).Value = Rs!fIngreso
      xlSheet.Cells(nFil, 6).NumberFormat = "dd/mm/yyyy;@"
      xlSheet.Cells(nFil, 7).Value = Trim(Rs!Cargo & "")
      xlSheet.Cells(nFil, 8).Value = Rs!Total * -1
      xlSheet.Cells(nFil, 8).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
      xlSheet.Cells(nFil, 9).Value = Trim(Rs!ccosto & "")
      nFil = nFil + 1
      msum = msum + 1
   End If
   mTotTrab = mTotTrab + 1
   Rs.MoveNext
Loop

Rs.Close: Set Rs = Nothing

Dim lFecha As String
lFecha = Format(Format(Ultimo_Dia(Cmbmes.ListIndex + 1, Val(Txtano.Text)), "00") & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text, "DD/MM/YYYY")

Sql = "usp_Pla_Seguro_Vida_Cuadra '" & wcia & "','" & lTipoTrab & "','" & lFecha & "'"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
nFil = nFil + 2
Do While Not Rs.EOF
   xlSheet.Cells(nFil, 2).Value = Trim(Rs!PlaCod & "")
   xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
   xlSheet.Cells(nFil, 4).Value = "F. Ingreso"
   xlSheet.Cells(nFil, 5).Value = Rs!fIngreso
   xlSheet.Cells(nFil, 5).NumberFormat = "dd/mm/yyyy;@"
   xlSheet.Cells(nFil, 6).Value = Rs!fIngreso
   xlSheet.Cells(nFil, 7).Value = Rs!fcese
   xlSheet.Cells(nFil, 7).NumberFormat = "dd/mm/yyyy;@"
   Select Case Trim(Rs!tipo & "")
          Case "A": xlSheet.Cells(nFil, 8).Value = "Trabajador cesado en periodo anterior"
          Case "C": xlSheet.Cells(nFil, 8).Value = "Trabajador cesado en periodo actual"
          Case "S": xlSheet.Cells(nFil, 8).Value = "Trabajador sin boleta en periodo actual"
   End Select
   nFil = nFil + 1
   Rs.MoveNext
Loop
Rs.Close: Set Rs = Nothing
'For I = 1 To 3
'   xlApp2.Sheets(I).Select
'   xlApp2.Sheets(I).Range("A1:A1").Select
'   xlApp2.Application.ActiveWindow.DisplayGridLines = False
'   'xlApp2.ActiveWindow.Zoom = 80
'Next
'xlApp2.Sheets(1).Select
'xlApp2.Application.Visible = True

'If lTipoTrab = "02" Then
'   If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
'   If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
'   If Not xlBook Is Nothing Then Set xlBook = Nothing
'   If Not xlSheet Is Nothing Then Set xlSheet = Nothing
'End If

Screen.MousePointer = vbDefault

End Sub
Private Sub Procesa_Seguro_Vida_Antes()
Dim mcalc As Integer
Dim mseguro As Boolean
Dim mcad As String
Dim mtoting As Currency
Dim mtotplani As Currency
Dim mItem As Integer
mseguro = False
Sql$ = nombre()
Sql$ = Sql$ & "placod,tipotrabajador,fingreso,area from planillas where cia='" & wcia & "' and tipotrabajador LIKE '" & Trim(mTipo) + "%" & "' and cat_trab!='04' and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
RUTA$ = App.Path & "\REPORTS\" & "SegVida.txt"
Open RUTA$ For Output As #1
mtoting = 0
mItem = 1
mlinea = 60
Sql$ = "Select sum(totaling-i18) as ing from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
If (fAbrRst(rs2, Sql$)) Then
   If Not IsNull(rs2!ing) Then mtotplani = rs2!ing Else mtotplani = 0
End If
If rs2.State = 1 Then rs2.Close
If mlinea > 55 Then Cabeza_Seguro (mtotplani)
Do While Not Rs.EOF
   If Year(Rs!fIngreso) = Val(Mid(mFecha, 7, 4)) Then
      mcalc = perendat(mFecha, Format(Rs!fIngreso, "dd/mm/yyyy"), "m")
      If mcalc < 3 Then mseguro = True
   ElseIf Year(Rs!fIngreso) > Val(Mid(mFecha, 7, 4)) Then
      mseguro = True
   Else
       If Year(Rs!fIngreso) < (Val(Mid(mFecha, 7, 4)) - 1) Then
          mseguro = False
       Else
          If Val(Mid(mFecha, 4, 2)) > 3 Or Month(Rs!fIngreso) < 10 Then
             mseguro = False
          Else
             If 12 - (Month(Rs!fIngreso) - Val(Mid(mFecha, 4, 2))) < 3 Then
                mseguro = True
             ElseIf 12 - (Month(Rs!fIngreso) - Val(Mid(mFecha, 4, 2))) = 3 Then
                If Day(Rs!fIngreso) < Val(Mid(mFecha, 1, 2)) Then
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
      Sql$ = "Select sum(totaling-i18) as ing from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and placod='" & Trim(Rs!PlaCod) & "' and status<>'*'"
      mcad = ""
      If (fAbrRst(rs2, Sql$)) Then
         If Not IsNull(rs2!ing) Then
            mcad = fCadNum(mItem, "##0") & ".-" & " " & Format(Rs!fIngreso, "dd/mm/yyyy") & " " & lentexto(40, Left(Rs!nombre, 40)) & " (" & fCadNum(rs2!ing, "##,###,##0.00") & ")"
            mtoting = mtoting + rs2!ing
            If rs2.State = 1 Then rs2.Close
            wciamae = Determina_Maestro("01044")
            Sql$ = "Select descrip from maestros_2 where cod_maestro2='" & Rs!Area & "' and status<>'*'"
            Sql$ = Sql$ & wciamae
            If (fAbrRst(rs2, Sql$)) Then mcad = mcad & "  " & rs2!DESCRIP
            If rs2.State = 1 Then rs2.Close
            If mlinea > 55 Then Cabeza_Seguro (mtotplani)
            Print #1, mcad
            mlinea = mlinea + 1
         End If
      End If
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
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
Print #1, Trim(Cmbcia.Text)
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
If Txtsemana.Text = "" Then Txtsemana.Text = "53": Exit Sub
If Txtsemana.Text > 1 Then Txtsemana = Txtsemana - 1
End Sub

Private Sub SpinButton2_SpinUp()
If Txtsemana.Text = "" Then Txtsemana.Text = "1": Exit Sub
If Txtsemana < 53 Then Txtsemana = Txtsemana + 1
End Sub
Private Sub Resumen_Planilla_antes()
Dim rscargo As ADODB.Recordset
Dim mmes As Integer
Dim mano As Integer
Dim msem As String
Dim I As Integer
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
Dim Rs As New ADODB.Recordset
Dim Inicio As Boolean


If CmbPlanta.Text = "TOTAL" Or CmbPlanta.ListIndex < 0 Then mPlanta = "" Else mPlanta = fc_CodigoComboBox(CmbPlanta, 2)

If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Seleccionar Tipo de Trabajador", vbInformation, "Resumne de Planilla": Exit Sub
mmes = Cmbmes.ListIndex + 1
mano = Val(Txtano.Text)
msem = Txtsemana.Text

If mtipobol <> "01" And mtipobol <> "05" Then msem = ""
If mTipo = "01" Then msem = ""

'rpt = MsgBox("Desea Imprimir Titulo de Cabecera", vbYesNo)

'*****************desde aqui para la planilla************************************
mcadI = "": mcadd = "": mcada = ""

'*************esto puede ser*********************
For I = 1 To 50
   mcadI = mcadI & "sum(i" & Format(I, "00") & ") as i" & Format(I, "00") & ","
   MCADIT = MCADIT & "i" & Format(I, "00") & "+"
   If I <= 30 Then
      If I <= 21 Then
         mcadd = mcadd & "sum(d" & Format(I, "00") & ") as d" & Format(I, "00") & ","
         mcada = mcada & "sum(a" & Format(I, "00") & ") as a" & Format(I, "00") & ","
         mcaddt = mcaddt & "d" & Format(I, "00") & "+"
         mcadat = mcadat & "a" & Format(I, "00") & "+"
      End If
      If I = 14 Then
        Dim fecha As Date
        fecha = DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text))
        mcadh = mcadh & "CASE WHEN SUM(H14)>" & Day(fecha) & " THEN " & Day(fecha) & "  ELSE SUM(H14) END as h" & Format(I, "00") & ","
        mcadht = mcadht & "h" & Format(I, "00") & "+"
      Else
        mcadh = mcadh & "sum(h" & Format(I, "00") & ") as h" & Format(I, "00") & ","
        mcadht = mcadht & "h" & Format(I, "00") & "+"
      End If
   End If
Next

mcadI = Mid(mcadI, 1, Len(Trim(mcadI)) - 1)
mcadd = Mid(mcadd, 1, Len(Trim(mcadd)) - 1)
mcada = Mid(mcada, 1, Len(Trim(mcada)) - 1)
mcadh = Mid(mcadh, 1, Len(Trim(mcadh)) - 1)

Sql$ = "select " & mcadh & "," & mcadI & "," & mcadd & "," & mcada
Sql$ = Sql$ & ",sum(totaling) as totaling,sum(totalapo) as totalapo,sum(totalded) as totalded,sum(totneto) as totneto,placod "
Sql$ = Sql$ & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & mano & " and month(fechaproceso)=" & mmes & " and proceso LIKE '" & Trim(mtipobol) + "%" & "' and semana LIKE '" & Trim(msem) + "%" & "' "
If mPlanta <> "" Then Sql$ = Sql$ & " and obra='" & mPlanta & "' "
Sql$ = Sql$ & " and tipotrab='" & mTipo & "' and status<>'*' Group by placod order by placod"
'*****************************************************

If Not (Funciones.fAbrRst(Rs, Sql$)) Then MsgBox "No Existen Boletas Registradas Segun Parametros", vbCritical, "Resumen de Planillas": Exit Sub

Dim mArchRes As String
mArchRes = "RE" & wcia & mTipo & mtipobol & ".txt"
RUTA$ = App.Path & "\REPORTS\" & mArchRes

Open RUTA$ For Output As #1

Rs.MoveFirst
mlinea = 60
numtra = 0
Inicio = True
Do While Not Rs.EOF
   If mlinea > 50 Then
   If Not Inicio Then
        Print #1, Chr(12) + Chr(13)
    Else
        Inicio = False
   End If

    Call Cabeza_Resumen("")
    
   End If
   Sql$ = nombre()
   Sql$ = Sql$ & "cargo,fingreso,fcese,ipss from planillas " & _
   "where cia='" & wcia & "' and placod='" & Rs!PlaCod & _
   "' and status<>'*'"
  
   If Not (fAbrRst(rs2, Sql$)) Then
      MsgBox "Trabajador no Encontrado en el Maestro " & Rs!PlaCod, vbCritical, "Resumen de Planillas"
      Exit Sub
   End If
   
   wciamae = Determina_Maestro_2("01055")
   Sql$ = "select cod_maestro3,descrip from maestros_3 where ciamaestro='" & wcia & "055" & "' and cod_maestro3='" & rs2!Cargo & "'"
   'SQL$ = SQL$ & wciamae
 
   If (fAbrRst(rscargo, Sql$)) Then mcargo = rscargo!DESCRIP Else mcargo = ""
   If rscargo.State = 1 Then rscargo.Close
   If IsNull(rs2!fcese) Then mcad = Space(10) Else mcad = Format(rs2!fcese, "dd/mm/yyyy")

   Print #1, Rs!PlaCod & Space(5) & lentexto(40, Left(rs2!nombre, 40)) & "  " & lentexto(20, Left(mcargo, 20)) & "  " & Format(rs2!fIngreso, "dd/mm/yyyy") & "  " & mcad & lentexto(15, Left(rs2!ipss, 15))
   If rs2.State = 1 Then rs2.Close
   numtra = numtra + 1
   'HORAS
   mfor = Len(Trim(Vcadh))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadh, I, 2))
       If Rs(mc - 1) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc - 1), "##,##0.00") & " "
       End If
   Next I
   Print #1, mcad
   mlinea = mlinea + 1
   
   'Ingresos Remunerativos
   mfor = Len(Trim(Vcadir))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadir, I, 2))
       If Rs(mc + 29) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc + 29), "##,##0.00") & " "
       End If
   Next I
   Print #1, mcad

   mlinea = mlinea + 1
   
   'Ingresos No Remunerativos
   mfor = Len(Trim(Vcadinr))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadinr, I, 2))
       If Rs(mc + 29) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc + 29), "##,##0.00") & " "
       End If
   Next I
   Print #1, mcad
   mlinea = mlinea + 1

   'DEDUCCIONES Y APORTACIONES
   mfor = Len(Trim(Vcadd))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadd, I, 2))
       If Rs(mc + 79) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc + 79), "##,##0.00") & " "
       End If
   Next I
   mfor = Len(Trim(Vcada))
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcada, I, 2))
       If Rs(mc + 99) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc + 99), "##,##0.00") & " "
       End If
   Next I
   mcad = mcad & Space(4) & fCadNum(Rs!totaling, "####,##0.00") & " " & fCadNum(Rs!totalded, "####,##0.00") & " " & fCadNum(Rs!totneto, "####,##0.00")
   Print #1, mcad
   mlinea = mlinea + 1
   Rs.MoveNext
Loop

Print #1,
Print #1, String(233, "=")
Print #1, Chr(12) + Chr(13)

Call Cabeza_Resumen("")

'Call Imprimir_Titulo

Print #1,
Print #1, Space(10) & "***** TOTAL PLANILLA *****"

Print #1,
Print #1, "                           H O R A S                             R E M U N E R A C I O N E S                        D E D U C C I O N E S                           A P O R T A C I O N E S"
Print #1, "                           ---------                             ---------------------------                        ---------------------                           -----------------------"
If Rs.State = 1 Then Rs.Close

Sql$ = "select " & mcadh & "," & mcadI & "," & mcadd & "," & mcada
Sql$ = Sql$ & ",sum(totaling) as totaling,sum(totalapo) as totalapo,sum(totalded) as totalded,sum(totneto) as totneto,placod "
Sql$ = Sql$ & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & mano & " and month(fechaproceso)=" & mmes & " and proceso LIKE '" & Trim(mtipobol) + "%" & "' and semana LIKE '" & Trim(msem) + "%" & "' " _
     & "and tipotrab='" & mTipo & "' and status<>'*' group by cia,placod"
     
     
Sql$ = "select " & mcadh & "," & mcadI & "," & mcadd & "," & mcada
Sql$ = Sql$ & ",sum(totaling) as totaling,sum(totalapo) as totalapo,sum(totalded) as totalded,sum(totneto) as totneto "
Sql$ = Sql$ & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & mano & " and month(fechaproceso)=" & mmes & " and proceso LIKE '" & Trim(mtipobol) + "%" & "' and semana LIKE '" & Trim(msem) + "%" & "' "
If mPlanta <> "" Then Sql$ = Sql$ & " and obra='" & mPlanta & "' "
Sql$ = Sql$ & " and tipotrab='" & mTipo & "' and status<>'*'"

If (fAbrRst(Rs, Sql$)) Then

Dim m As Integer
m = 0
mremun = 0

For I = 1 To 100 Step 2
   m = m + 1
   
   'TOTAL HORAS
   mcad = Space(10)
   If I <= Len(Vcadh) Then
      If I = 1 Then
         mcad = mcad & Mid(Rcadh, 1, 20)
      Else
         mcad = mcad & Mid(Rcadh, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadh, I, 2))
      If Rs(mc - 1) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(Rs(mc - 1), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(39)
   End If
   
   '=========================
   
   'TOTAL REMUNERACIONES AFECTAS
   mcad = mcad & Space(10)
   If I <= Len(Vcadir) Then
      If I = 1 Then
         mcad = mcad & Mid(Rcadir, 1, 20)
      Else
         mcad = mcad & Mid(Rcadir, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadir, I, 2))
      If Rs(mc + 29) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(Rs(mc + 29), "###,###,##0.00")
         mremun = mremun + Rs(mc + 29)
      End If
   Else
      mcad = mcad & Space(39)
   End If
   
   'TOTAL DEDUCCIONES
   mcad = mcad & Space(10)
   If I <= Len(Vcadd) Then
      If I = 1 Then
         mcad = mcad & Mid(Rcadd, 1, 20)
      Else
         mcad = mcad & Mid(Rcadd, 20 * (m - 1) + 1, 20)
      End If
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadd, I, 2))
      If Rs(mc + 79) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(Rs(mc + 79), "###,###,##0.00")
      End If
   Else
      mcad = mcad & Space(24)
   End If
   
   'TOTAL APORTACIONES
   mcad = mcad & Space(10)
   If I <= Len(Vcada) Then
      If I = 1 Then
         mcad = mcad & Mid(Rcada, 1, 20)
      Else
         mcad = mcad & Mid(Rcada, 20 * (m - 1) + 1, 20)
      End If
      mc = Val(Mid(Vcada, I, 2))
      If Rs(mc + 99) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(Rs(mc + 99), "###,###,##0.00")
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

Dim Z As Integer
Z = 0
For I = 1 To Len(Vcadinr) Step 2

   'TOTAL REMUNERACIONES NO AFECTAS
   'MODIFICADO 05/08/2008 RICARDO HINOSTROZA
   'MODIFICAR EL ACCESSO A LOS CONCEPTOS NO REMUNERATIVOS
   
   mcad = ""
   mcad = mcad & Space(59)
   If I <= Len(Vcadinr) Then
      If I = 1 Then
         mcad = mcad & Mid(Rcadinr, 1, 20)
      Else
          Z = Z + 1
          mcad = mcad & Mid(Rcadinr, 20 * (Z) + 1, 20)
          'mcad = mcad & Mid(Rcadinr, 20 * (i - 2) + 1, 20)
      End If
      
      mcad = mcad & Space(5)
      mc = Val(Mid(Vcadinr, I, 2))
      If Rs(mc + 29) = 0 Then
         mcad = mcad & Space(14)
      Else
         mcad = mcad & fCadNum(Rs(mc + 29), "###,###,##0.00")
         mremun = mremun + Rs(mc + 29)
      End If
   Else
      mcad = mcad & Space(14)
   End If
   If Trim(mcad) <> "" Then Print #1, mcad
Next I

Print #1, Space(84) & "--------------"
Print #1, Space(59) & "Sub Total II         " & fCadNum(mremun, "###,###,###,##0.00")
Print #1, Space(84) & "--------------                                   --------------                                 -----------"
Print #1, Space(59) & "* TOTAL REMUNERACION *" & fCadNum(Rs!totaling, "##,###,###,##0.00") & Space(10) & "* TOTAL DEDUCCIONES * " & fCadNum(Rs!totalded, "##,###,###,##0.00") & Space(10) & "* TOTAL APORTACIONES *" & fCadNum(Rs!totalapo, "#,###,##0.00")
Print #1,
Print #1, Space(59) & "*** NETO PAGADO ***   " & fCadNum(Rs!totneto, "##,###,###,##0.00")
Print #1, Space(59) & "======================================="
Print #1,
Print #1, "Total de Trabajadores    =>  " & Str(numtra)
Print #1, Chr(12) + Chr(13)
End If
Close #1
Call Funciones.Imprime_Txt(mArchRes, RUTA$)
End Sub
Private Sub Cabeza_Resumen_Antes(lDesCosto As String)
Dim rsremu As ADODB.Recordset
Dim mcadir As String
Dim mcadinr As String
Dim mcadd As String
Dim mcada As String
Dim mcadh As String
Dim cad As String
Dim f1 As String
Dim f2 As String
f1 = "": f2 = ""
mcadir = "": mcadd = "": mcada = "": mcadinr = "": mcadh = ""
Vcadir = "": Vcadd = "": Vcada = "": Vcadinr = "": Vcadh = ""
Rcadir = "": Rcadd = "": Rcada = "": Rcadinr = "": Rcadh = ""

'Fecha de semana
Sql$ = "select fechai,fechaf from plasemanas where cia='" & wcia & "' and ano='" & Format(Txtano.Text, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then f1 = Format(Rs!fechai, "dd/mm/yyyy"):     f2 = Format(Rs!fechaf, "dd/mm/yyyy")
If Rs.State = 1 Then Rs.Close
'Fin de Fecha de semana

If mTipo = "01" Then
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

If wGrupoPla = "01" Then 'En gallos no Pintar Dias en Resumen
   Sql$ = "select a.codigo,b.descrip from plaverhoras a," & _
       "maestros_2 b where b.ciamaestro='01077' " & _
       "and a.cia='" & wcia & "' and a.tipo_trab='" & _
       mTipo & "' and a.status<>'*' " & MCADENA _
       & "and a.codigo=b.cod_maestro2 and codigo not in('14','25','07','08','22') order by codigo"
Else
   Sql$ = "select a.codigo,b.descrip from plaverhoras a," & _
       "maestros_2 b where b.ciamaestro='01077' " & _
       "and a.cia='" & wcia & "' and a.tipo_trab='" & _
       mTipo & "' and a.status<>'*' " & MCADENA _
       & "and a.codigo=b.cod_maestro2 order by codigo"
End If
If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
Do While Not rs2.EOF
   mcadh = mcadh & " " & lentexto(9, Left(rs2!DESCRIP, 9)) & " "
   Rcadh = Rcadh & lentexto(20, Left(rs2!DESCRIP, 20))
   Vcadh = Vcadh & rs2!Codigo
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close

'INGRESOS

'MODIFICADO 05/08/2008 RICARDO HINOSTROZA
'MODIFICAR LOS CONCEPTOS QUE SE APLICARON EN PLANILLA

'Sql$ = "select a.tipo,a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
'    & "where a.cia='" & wcia & "' and a.tipo='I' and a.status<>'*' " _
'    & "and a.cia=b.cia and b.status<>'*' " & _
'  "and b.tipomovimiento='02' and " & _
   '  "a.tipo_trab='" & mtipo & "' and a.codigo=b.codinterno " _
   '  & "order by a.tipo,a.codigo"
   
Sql$ = "PLASS_INGRESO_RESUMENPLANILLA '" & wcia & "','" & Trim(Txtano.Text) & "','" & Trim(Cmbmes.ListIndex + 1) & "'"

If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
Do While Not rs2.EOF
   Sql$ = "select cod_remu from plaafectos where cia='" & _
   wcia & "' and status<>'*' and cod_remu='" & rs2!Codigo & _
   "' and tipo in ('A','D') AND CODIGO!='13' "
   If (fAbrRst(rsremu, Sql$)) Then
      mcadir = mcadir & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Rcadir = Rcadir & lentexto(20, Left(rs2!Descripcion, 20))
      Vcadir = Vcadir & rs2!Codigo
   Else
      Rcadinr = Rcadinr & lentexto(20, Left(rs2!Descripcion, 20))
      mcadinr = mcadinr & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Vcadinr = Vcadinr & rs2!Codigo
      'mcadinr =mcadinr & " " & lentexto(9, Left(rs2!Descripcion, 9))
   End If
   If rsremu.State = 1 Then rsremu.Close
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close

'DEDUCCIONES Y APORTACIONES
Sql$ = "select a.tipo,a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
     & "where a.cia='" & wcia & "' and a.tipo<>'I' and a.status<>'*' " _
     & "and a.cia=b.cia and b.status<>'*' and b.tipomovimiento='03' and a.tipo_trab='" & mTipo & "' and a.codigo=b.codinterno " _
     & "order by a.tipo,a.codigo"
    
If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst

Do While Not rs2.EOF
   Select Case rs2!tipo
          Case Is = "D"
               mcadd = mcadd & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Rcadd = Rcadd & lentexto(20, Left(rs2!Descripcion, 20))
               Vcadd = Vcadd & rs2!Codigo
          Case Is = "A"
               Rcada = Rcada & lentexto(20, Left(rs2!Descripcion, 20))
               mcada = mcada & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Vcada = Vcada & rs2!Codigo
   End Select
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close


    Print #1, Chr(18) + Trim(Cmbcia.Text) + Chr(14)
    'Print #1,
    'Print #1,
    Dim mPlanta As String
    If CmbPlanta.Text = "TOTAL" Or CmbPlanta.ListIndex < 0 Then mPlanta = "" Else mPlanta = " ( PLANTA : " & CmbPlanta.Text & " )"
    
    If mTipo = "01" Then
        Print #1, Space(40) & Chr(14) & "PLANILLA DE PAGO EMPLEADOS " & Cmbtipbol.Text & mPlanta & Chr(20)
    Else
        Print #1, Space(40) & Chr(14) & "PLANILLA DE PAGO OBREROS " & Cmbtipbol.Text & mPlanta & Chr(20)
    End If
    Print #1, Chr(18) + Chr(14)
    'Print #1, Chr(14)
'    Print #1, Chr(14)
 '   Print #1, Chr(14)
    Print #1, Chr(20)

    'Print #1, Chr(20)
    Print #1, Space(50) & "MES de " & Cmbmes.Text & "  de " & Txtano.Text & Chr(15)
    If mTipo <> "01" And Val(Txtsemana.Text) > 0 Then
        Print #1, "F. Imp " & Format(Date, "dd/mm/yyyy") & Space(10) & "SEMANA No =>  " & Txtsemana.Text & "  Del " & f1 & " Al " & f2 & "  " & lDesCosto
    Else
        Print #1, "F. Imp " & Format(Date, "dd/mm/yyyy") & "  " & lDesCosto
    End If
    Print #1, String(233, "-")
    Print #1, "Codigo    Apellidos y Nombres del Trabajador        Ocupacion             F. Ingreso   F.cese   I.P.S.S."
    'Resumen de Liquidación
    'Modificado el 05/09/2008 / Error cadena Vacia
    'Ricardo Hinostroza
    
    Print #1, String(IIf(Len(mcadh) = 0, 10, Len(mcadh)) / 2 - 5, "-") & "   HORAS  " & String(IIf(Len(mcadh) = 0, 10, Len(mcadh)) / 2 - 5, "-")

    Print #1, mcadh
    Print #1, String(Abs(Len(mcadir) / 2 - 12), "-") & " INGRESOS REMUNERATIVOS " & String(Abs(Len(mcadir) / 2 - 12), "-")
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
mSemana = ""
mpag = 0
If mTipo <> "01" And Txtsemana.Text <> "" Then
   Sql$ = "Select placod,sum(d13) as quinta from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and semana='" & Trim(Txtsemana.Text) & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' Group by placod order by placod"
Else
   Sql$ = "Select placod,sum(d13) as quinta from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' Group by placod order by placod"
End If
If Not (fAbrRst(Rs, Sql$)) Then MsgBox "No Hay Retencion de Quinta Segun Paramentros", vbInformation, "Lista de Quinta Categoria": Exit Sub
RUTA$ = App.Path & "\REPORTS\" & "Lquinta.txt"
Open RUTA$ For Output As #1
Rs.MoveFirst
Cabeza_Lista_Quinta
mtotq = 0
Do While Not Rs.EOF
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Cabeza_Lista_Quinta
   Sql$ = nombre()
   Sql$ = Sql$ & "placod from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(45, Left(rs2!nombre, 45)) Else mcad = Space(40)
   mcad = Rs!PlaCod & "   " & mcad & Space(5) & fCadNum(Rs!quinta, "##,###,##0.00")
   Print #1, Space(2) & mcad
   mtotq = mtotq + Rs!quinta
   mItem = mItem + 1
   Rs.MoveNext
Loop
Print #1,
Print #1, Space(35) & "TOTAL :                 " & fCadNum(mtotq, "###,###,##0.00")
Print #1, Space(35) & "TOTAL TRABAJADORES      " & fCadNum(mItem, "##,###,###,###")
Close #1
Call Imprime_Txt("Lquinta.txt", RUTA$)
End Sub
Private Sub Cabeza_Lista_Quinta()
mpag = mpag + 1
Print #1, Chr(18) & Space(2) & Trim(Cmbcia.Text) & Space(25) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(22) & "REPORTE DE QUINTA CATEGORIA"
Print #1, Space(23) & "PERIODO : " & Cmbmes.Text & " - " & Format(Txtano.Text, "0000")
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
Dim mTipo As String
Dim MArea As String
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='03' and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
   Rs.MoveFirst
   mcad = ""
   Do While Not Rs.EOF
      mcad = mcad & "i" & Trim(Rs!cod_remu) & "+"
      Rs.MoveNext
   Loop
   mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
End If
If Rs.State = 1 Then Rs.Close
mcad = "sum" & mcad & " as afecto "
Sql$ = nombre()
Sql = Sql$ & "p.placod,p.tipotrabajador,p.area,sum(a03) as senati," & mcad & "from planillas p,plahistorico h where h.cia='" & wcia & "' " _
    & "and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and h.status<>'*' and a03<>0 " _
    & "and p.cia=h.cia and p.placod=h.placod and p.status<>'*' " _
    & "group by h.placod order by p.tipotrabajador,p.area"
    
If Not (fAbrRst(Rs, Sql$)) Then MsgBox "No Hay Aportaciones al SENATI Segun Paramentros", vbInformation, "Aportaciones al SENATI": Exit Sub
RUTA$ = App.Path & "\REPORTS\" & "Senati.txt"
Open RUTA$ For Output As #1
Rs.MoveFirst
mtotsenati = 0
mtotafecto = 0
mTipo = ""
MArea = ""
mItem = 1
mpag = 0
Call Cabeza_Senati(Rs!Area, Rs!TipoTrabajador)
totaltrab = 0
Do While Not Rs.EOF
   If (mTipo <> Rs!TipoTrabajador Or MArea <> Rs!Area) And mTipo <> "" Then
      Print #1, Space(5) & String(103, "-")
      Print #1, Space(30) & "T O T A L E S .... " & Space(13) & fCadNum(mtotafecto, "###,###,###.00") & Space(2) & fCadNum(mtotsenati, "###,###,###.00")
      Print #1, SaltaPag
      Call Cabeza_Senati(Rs!Area, Rs!TipoTrabajador)
      mtotsenati = 0
      mtotafecto = 0
      mTipo = Rs!TipoTrabajador
      MArea = Rs!Area
      mItem = 1
   End If
   If mTipo = "" Then mTipo = Rs!TipoTrabajador: MArea = Rs!Area
   Print #1, Space(5) & fCadNum(mItem, "##0") & ".-" & Space(5) & mchartipo & Space(3) & lentexto(40, Left(Rs!nombre, 40)) & Space(3) & fCadNum(Rs!AFECTO, "###,###,###.00") & Space(2) & fCadNum(Rs!senati, "###,###,###.00") & Space(11) & Rs!PlaCod
   mlinea = mlinea + 1
   totaltrab = totaltrab + 1
   If mlinea > 55 Then Print #1, SaltaPag: Call Cabeza_Senati(Rs!Area, Rs!TipoTrabajador)
   mtotsenati = mtotsenati + Rs!senati
   mtotafecto = mtotafecto + Rs!AFECTO
   mtotsenatit = mtotsenatit + Rs!senati
   mtotafectot = mtotafectot + Rs!AFECTO
   mItem = mItem + 1
   Rs.MoveNext
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
If (fAbrRst(rs2, Sql$)) Then mdescosto = rs2!DESCRIP
If rs2.State = 1 Then rs2.Close

Print #1, LetraChica & Space(5) & Trim(Cmbcia.Text) & Space(2) & "( " & mdescosto & " )"
mdescosto = ""
wciamae = Determina_Maestro("01055")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where cod_maestro2='" & tipo & "' and status<>'*'"
Sql$ = Sql$ & wciamae
If (fAbrRst(rs2, Sql$)) Then mdescosto = rs2!DESCRIP: mchartipo = Left(rs2!DESCRIP, 1)
If rs2.State = 1 Then rs2.Close
Print #1,
Print #1, Space(20) & "REPORTE DE TRABAJADORES CON APORTACION AL SENATI - MES DE "; Cmbmes.Text
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

mSemana = ""
mpag = 0
mcad = "sum(d" & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo"
mfijo = False
Sql$ = "select status,adicional from placonstante where tipomovimiento='03' and codinterno='" & VConcepto & "' and status <>'*'"
If (fAbrRst(Rs, Sql$)) Then If Rs!status = "F" Or Rs!adicional = "S" Then mfijo = True
If Rs.State = 1 Then Rs.Close

If mfijo = False Then
    'Ingresos Afectos para Aportacion
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='01'  and  codigo='" & VConcepto & "' and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadIA = ""
       Do While Not Rs.EOF
          mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    'Ingresos Afectos para Deduccion
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='01'  and  codigo='" & VConcepto & "' and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadID = ""
       Do While Not Rs.EOF
          mcadID = mcadID & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadID = "(" & Mid(mcadID, 1, Len(Trim(mcadID)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    If mcadID <> "" Then mcad = mcad & ",sum" & mcadID & " as afectod"
    Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(Rs, Sql$)) Then MsgBox "No Hay " & CmbConcepto.Text & " Segun Paramentros", vbInformation, "Deducciones y Aportaciones": Exit Sub
RUTA$ = App.Path & "\REPORTS\" & "DeducApor.txt"
Open RUTA$ For Output As #1
Rs.MoveFirst
Cabeza_Lista_DeducApor (Rs!TipoTrab)
mItem = 1
totafeca = 0: totafecd = 0: totapo = 0: totded = 0
Do While Not Rs.EOF
   If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Lista_DeducApor (Rs!TipoTrab)
   Sql$ = nombre()
   Sql$ = Sql$ & "placod from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   mcad = fCadNum(mItem, "##0") & ".-" & "  " & mchartipo & "  " & mcad
   If mfijo = False Then
      If mcadIA <> "" Then mcad = mcad & Space(2) & fCadNum(Rs!afectoa, "#,###,###.00") & Space(2) & fCadNum(Rs!APO, "###,###.00") & Space(5): totafeca = totafeca + Rs!afectoa Else mcad = mcad & Space(31)
      If mcadID <> "" Then mcad = mcad & Space(2) & fCadNum(Rs!afectod, "#,###,###.00") & Space(2) & fCadNum(Rs!ded, "###,###.00"): totafecd = totafecd + Rs!afectod Else mcad = mcad & Space(26)
   Else
      mcad = mcad & Space(16) & fCadNum(Rs!APO, "###,###.00") & Space(21) & fCadNum(Rs!ded, "###,###.00")
   End If
   mcad = mcad & Space(2) & fCadNum(Rs!APO + Rs!ded, "#,###,###.00")
   Print #1, Space(5) & mcad
   totapo = totapo + Rs!APO
   totded = totded + Rs!ded
   mItem = mItem + 1
   Rs.MoveNext
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
If (fAbrRst(rs2, Sql$)) Then mchartipo = Left(rs2!DESCRIP, 1)
If rs2.State = 1 Then rs2.Close

Print #1, Chr(15) & Space(2) & Trim(Cmbcia.Text) & Space(80) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(2) & "REPORTE DE " & CmbConcepto.Text
Print #1, Space(50) & "PERIODO : " & Cmbmes.Text & " - " & Format(Txtano.Text, "0000")
Print #1, Space(2) & Cmbtipo.Text
Print #1, Space(2) & String(126, "-")
Print #1, Space(2) & "Orden Tipo         Nombre                                Rem. Afecta     Empr.          Rem. Afecta     Trab.          Total"
Print #1, Space(2) & String(126, "-")
mlinea = 10
End Sub

Private Sub Procesa_Seguro_txt(mtope As Double, lTipoTrab As String)
Dim mperiodo As String
Dim mtiposeg As String
Dim MArea As String
Dim wciamae1 As String
Dim wciamae2 As String
Dim mItem As Integer
Dim tottrab As Integer
Dim mtingarea As Currency
Dim mtingareaP As Currency
Dim mtingplanta As Currency
Dim mtingplantaP As Currency
Dim mtingtotal As Currency
Dim mtingtotalP As Currency
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
Dim m4 As Integer
Dim m1ss As Currency
Dim m2ss As Currency
Dim m3ss As Currency
Dim m1sp As Currency
Dim m2sp As Currency
Dim m3sp As Currency

Dim m4ss As Currency
Dim m4sp As Currency
mpag = 0
If Rs.RecordCount > 0 Then
   Rs.MoveFirst
   Do While Not Rs.EOF
      If Rs!cesado = "N" Then
         mtiposeg = Trim(Rs!tipcalcseguro & "")
         MArea = Trim(Rs!Area & "")
         Rs.MoveFirst
         Exit Do
      End If
      Rs.MoveNext
   Loop
End If


RUTA$ = App.Path & "\REPORTS\" & "SeguroSCRT" & lTipoTrab & ".txt"
Open RUTA$ For Output As #1
Cabecera_SCRT (mtope)
Call Quiebre(MArea, mtiposeg)
mItem = 1
tottrab = 0

mtingarea = 0: mtsaludarea = 0: mtpensionarea = 0: mtingareaP = 0
mtingplanta = 0: mtsaludplanta = 0: mtpensionplanta = 0: mtingplantaP = 0
mtingtotal = 0: mtsaludtotal = 0: mtpensiontotal = 0: mtingtotalP = 0
M1 = 0: m2 = 0: m3 = 0: m4 = 0
m1ss = 0: m2ss = 0: m3ss = 0: m1sp = 0: m2sp = 0: m3sp = 0: m4sp = 0: m4ss = 0
Dim mSuledoTope As Double
Do While Not Rs.EOF
   'If rs!TipoTrabajador = lTipoTrab And Trim(rs!cesado) = "N" Then
   'add jcms 011221 se quita del detalle los trabajadores sin boletas , indidcado por KP+JA
   If Rs!TipoTrabajador = lTipoTrab And Trim(Rs!cesado) = "N" And Trim(Rs!SINBOLETA) <> "S" Then
       If mtiposeg <> Trim(Rs!tipcalcseguro & "") Then
          Print #1, Space(49) & " " & "------------" & " " & "---------" & " " & "------------" & " " & "---------"
          Print #1, Space(49) & " " & fCadNum(mtingarea, "#,###,###.00") & " " & fCadNum(mtsaludarea, "##,###.00") & " " & fCadNum(mtingareaP, "#,###,###.00") & "  " & fCadNum(mtpensionarea, "##,###.00")
          Print #1, Space(49) & " " & "------------" & " " & "---------" & " " & "------------" & "  " & "---------"
          Print #1, Space(49) & " " & fCadNum(mtingplanta, "#,###,###.00") & " " & fCadNum(mtsaludplanta, "##,###.00") & " " & fCadNum(mtingplantaP, "#,###,###.00") & "  " & fCadNum(mtpensionplanta, "##,###.00")
          Print #1, Space(49) & " " & "------------" & " " & "---------" & " " & "------------" & "  " & "---------"
          mlinea = mlinea + 5
          If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
          MArea = Trim(Rs!Area & "")
          mtiposeg = Trim(Rs!tipcalcseguro & "")
          mtingarea = 0: mtsaludarea = 0: mtpensionarea = 0: mtingareaP = 0
          mtingplanta = 0: mtsaludplanta = 0: mtpensionplanta = 0: mtingplantaP = 0
          
          Call Quiebre(MArea, mtiposeg)
          mItem = 1
       ElseIf MArea <> Rs!Area Then
          Print #1, Space(49) & " " & "------------" & " " & "---------" & " " & "------------" & "  " & "---------"
          Print #1, Space(49) & " " & fCadNum(mtingarea, "#,###,###.00") & " " & fCadNum(mtsaludarea, "##,###.00") & " " & fCadNum(mtingareaP, "#,###,###.00") & "  " & fCadNum(mtpensionarea, "##,###.00")
          mlinea = mlinea + 5
          If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
          MArea = Trim(Rs!Area & "")
          mtiposeg = Rs!tipcalcseguro
          mtingarea = 0: mtsaludarea = 0: mtpensionarea = 0: mtingareaP = 0
          Call Quiebre(MArea, "")
          mItem = 1
       End If
       mtasa1 = Rs!porsalud: mtasa2 = Rs!porpension
   
      mtasa1 = Round(Rs!sueldo * mtasa1 / 100, 2)
      If Rs!sueldo > mtope Then mSuledoTope = mtope Else mSuledoTope = Rs!sueldo
      If Rs!sueldo > mtope Then mtasa2 = Round(mtope * mtasa2 / 100, 2) Else mtasa2 = Round(Rs!sueldo * mtasa2 / 100, 2)
      Print #1, fCadNum(mItem, "###") & ".-" & Rs!PlaCod & " " & lentexto(35, Left(RTrim(Rs!ap_pat) & " " & RTrim(Rs!ap_mat) & " " & RTrim(Rs!nom_1) & " " & RTrim(Rs!nom_2), 40)) & " " & fCadNum(Rs!sueldo, "#,###,###.00") & " " & fCadNum(mtasa1, "##,###.00") & " " & fCadNum(mSuledoTope, "#,###,###.00") & "  " & fCadNum(mtasa2, "##,###.00")
      mlinea = mlinea + 1
      If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
      Select Case Rs!codsctr
             Case Is = "01"
                  m1ss = m1ss + mtasa1
                  m1sp = m1sp + mtasa2
                  M1 = M1 + 1
             Case Is = "02"
                  m2ss = m2ss + mtasa1
                  m2sp = m2sp + mtasa2
                  m2 = m2 + 1
             Case Is = "03"
                  m3ss = m3ss + mtasa1
                  m3sp = m3sp + mtasa2
                  m3 = m3 + 1
             Case Is = "04"
                  m4ss = m4ss + mtasa1
                  m4sp = m4sp + mtasa2
                  m4 = m4 + 1
      End Select
      mtingarea = mtingarea + Rs!sueldo: mtsaludarea = mtsaludarea + mtasa1: mtpensionarea = mtpensionarea + mtasa2
      mtingareaP = mtingareaP + mSuledoTope
      mtingplanta = mtingplanta + Rs!sueldo: mtsaludplanta = mtsaludplanta + mtasa1: mtpensionplanta = mtpensionplanta + mtasa2
      mtingplantaP = mtingplantaP + mSuledoTope
      mtingtotal = mtingtotal + Rs!sueldo: mtsaludtotal = mtsaludtotal + mtasa1: mtpensiontotal = mtpensiontotal + mtasa2
      mtingtotalP = mtingtotalP + mSuledoTope
      
      totrab = tottrab + 1
      mItem = mItem + 1
   End If
   Rs.MoveNext
Loop
Print #1, Space(49) & " " & "------------" & " " & "---------" & " " & "------------" & "  " & "---------"
Print #1, Space(49) & " " & fCadNum(mtingarea, "#,###,###.00") & " " & fCadNum(mtsaludarea, "##,###.00") & " " & fCadNum(mtingareaP, "#,###,###.00") & "  " & fCadNum(mtpensionarea, "##,###.00")
Print #1, Space(49) & " " & "------------" & " " & "---------" & " " & "------------" & "  " & "---------"
Print #1, Space(49) & " " & fCadNum(mtingplanta, "#,###,###.00") & " " & fCadNum(mtsaludplanta, "##,###.00") & " " & fCadNum(mtingplantaP, "#,###,###.00") & "  " & fCadNum(mtpensionplanta, "##,###.00")
Print #1, Space(49) & " " & "------------" & " " & "---------" & " " & "------------" & "  " & "---------"

mlinea = mlinea + 5
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)

Print #1, Space(49) & " " & "============" & " " & "=========" & " " & "============" & "  " & "========="
Print #1, Space(49) & " " & fCadNum(mtingtotal, "#,###,###.00") & " " & fCadNum(mtsaludtotal, "##,###.00") & " " & fCadNum(mtingtotalP, "#,###,###.00") & "  " & fCadNum(mtpensiontotal, "##,###.00")
Print #1, Space(49) & " " & "============" & " " & "=========" & " " & "============" & "  " & "========="

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
Print #1, fCadNum(m4, "####") & Space(8) & "PERSONAL ADMINISTRATIVO EN SOCAVON             " & fCadNum(m4ss, "###,###.00") & " " & fCadNum(m4sp, "##,###.00")
mlinea = mlinea + 5
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, Space(58) & " " & " " & "=========" & " " & "========="
Print #1, fCadNum(m4 + m3 + m2 + M1, "####") & Space(55) & fCadNum(m4ss + m3ss + m2ss + m1ss, "###,###.00") & " " & fCadNum(m4sp + m3sp + m2sp + m1sp, "##,###.00")
Print #1, Space(58) & " " & " " & "=========" & " " & "========="
Print #1,
Print #1,
mlinea = mlinea + 3
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, "TRABAJADORES CESAD"
Print #1,
mlinea = mlinea + 2
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
If Rs.RecordCount > 0 Then Rs.MoveFirst
mItem = 1
mtingtotal = 0
Do While Not Rs.EOF
   If Rs!TipoTrabajador = lTipoTrab And Trim(Rs!cesado) = "S" Then
      Print #1, fCadNum(mItem, "###") & ".-" & Rs!PlaCod & " " & lentexto(35, Left(RTrim(Rs!ap_pat) & " " & RTrim(Rs!ap_mat) & " " & RTrim(Rs!nom_1) & " " & RTrim(Rs!nom_2), 40)) & " " & fCadNum(Rs!sueldo, "#,###,###.00")
      mtingtotal = mtingtotal + Rs!sueldo
      mItem = mItem + 1
      If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
      mlinea = mlinea + 1
   End If
   Rs.MoveNext
Loop
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, Space(49) & " " & "============"
Print #1, Space(49) & " " & fCadNum(mtingtotal, "#,###,###.00")
Print #1, Space(49) & " " & "============"

'Ingresos no afectos
mlinea = mlinea + 3
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, "NO SE CONSIDERAN LOS SIGUIENTES CONCEPTOS"
Print #1,
mlinea = mlinea + 2
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
If Rs.RecordCount > 0 Then Rs.MoveFirst
mItem = 1
mtingtotal = 0
Do While Not Rs.EOF
   If Rs!TipoTrabajador = lTipoTrab And Trim(Rs!cesado) = "*" Then
      Print #1, fCadNum(mItem, "###") & ".-" & Rs!PlaCod & " " & lentexto(35, Left(RTrim(Rs!ap_pat) & " " & RTrim(Rs!ap_mat) & " " & RTrim(Rs!nom_1) & " " & RTrim(Rs!nom_2), 40)) & " " & fCadNum(Rs!sueldo, "#,###,###.00")
      mtingtotal = mtingtotal + Rs!sueldo
      mItem = mItem + 1
      If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
      mlinea = mlinea + 1
   End If
   Rs.MoveNext
Loop
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, Space(49) & " " & "============"
Print #1, Space(49) & " " & fCadNum(mtingtotal, "#,###,###.00")
Print #1, Space(49) & " " & "============"
mlinea = mlinea + 3

'TRABAJADORES SIN BOLETAS
Print #1,
Print #1,
Print #1,
mlinea = mlinea + 3

Print #1, "TRABAJADORES QUE SE INCLUYEN QUE NO TIENEN BOLETA REGISTRADA EN EL PERIODO"
Print #1,
mlinea = mlinea + 2
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
If Rs.RecordCount > 0 Then Rs.MoveFirst
mItem = 1
mtingtotal = 0
Do While Not Rs.EOF
   If Rs!TipoTrabajador = lTipoTrab And Rs!SINBOLETA = "S" Then
      Print #1, fCadNum(mItem, "###") & ".-" & Rs!PlaCod & " " & lentexto(35, Left(RTrim(Rs!ap_pat) & " " & RTrim(Rs!ap_mat) & " " & RTrim(Rs!nom_1) & " " & RTrim(Rs!nom_2), 40)) & " " & fCadNum(Rs!sueldo, "#,###,###.00")
      mtingtotal = mtingtotal + Rs!sueldo
      mItem = mItem + 1
      If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
      mlinea = mlinea + 1
   End If
   Rs.MoveNext
Loop
If mlinea > 55 Then Print #1, SaltaPag: Cabecera_SCRT (mtope)
Print #1, Space(49) & " " & "============"
Print #1, Space(49) & " " & fCadNum(mtingtotal, "#,###,###.00")
Print #1, Space(49) & " " & "============"
Close #1

If lTipoTrab = "02" Then
   MsgBox "Se Generarón los reportes en " & App.Path & "\REPORTS\"
End If
End Sub
Private Sub Quiebre(Area As String, mtiposeg As String)
If mtiposeg <> "" Then
   Print #1, mtiposeg
   Print #1,
End If
Print #1, Area
Print #1,
End Sub
Private Sub Cabecera_SCRT(tope As Currency)
mpag = mpag + 1
Print #1, Cmbcia.Text & "    " & Cmbtipo.Text & "  MES DE " & Cmbmes.Text & " DEL " & Txtano.Text
Print #1, "TOPE : " & fCadNum(tope, "###,###,###.00") & Space(48) & "Pag : " & fCadNum(mpag, "####")
Print #1,
Print #1, "CODIGO         NOMBRE                              TOT, INGRESO  S.RIESGO   INGRESO     SEGURO COMP. "
Print #1, "                                                                  SALUD     CON TOPE     PENSIONES"
Print #1, String(100, "-")
mlinea = 7
End Sub
Private Sub Procesa_Remunera()
Dim mnumh As Integer
Dim mItem As Integer
Dim mcad As String
Dim totali As Currency
Dim totalh As Currency
If CmbConcepto.ListIndex < 0 Then MsgBox "Debe Seleccionar Concepto", vbInformation, "Remuneraciones": Exit Sub
mSemana = ""
mpag = 0
mcad = "sum(i" & Format(VConcepto, "00") & ") as ing"
mnumh = Remun_Horas(VConcepto)
If mnumh > 0 Then mcad = mcad & ",sum(h" & Format(mnumh, "00") & ") as horas"

Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
If Not (fAbrRst(Rs, Sql$)) Then MsgBox "No Hay " & CmbConcepto.Text & " Segun Paramentros", vbInformation, "Deducciones y Aportaciones": Exit Sub
RUTA$ = App.Path & "\REPORTS\" & "Remunera.txt"
Open RUTA$ For Output As #1
Rs.MoveFirst
Cabeza_Remunera (Rs!TipoTrab)
mItem = 1
totali = 0: totalh = 0
Do While Not Rs.EOF
   If Rs!ing <> 0 Then
      If mlinea > 55 Then Print #1, SaltaPag: Cabeza_Remunera (Rs!TipoTrab)
      Sql$ = nombre()
      Sql$ = Sql$ & "placod from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      mcad = fCadNum(mItem, "##0") & ".-" & "  " & mchartipo & "  " & mcad
      If mnumh > 0 Then
         mcad = mcad & Space(2) & fCadNum(Rs!horas, "#####.00") & Space(3) & fCadNum(Rs!ing, "###,###.00")
      Else
         mcad = mcad & Space(13) & fCadNum(Rs!ing, "###,###.00")
      End If
      Print #1, Space(5) & mcad
      totali = totali + Rs!ing
      If mnumh > 0 Then
         totalh = totalh + Rs!horas
      End If
      mItem = mItem + 1
   End If
   Rs.MoveNext
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
If (fAbrRst(rs2, Sql$)) Then mchartipo = Left(rs2!DESCRIP, 1)
If rs2.State = 1 Then rs2.Close

Print #1, Chr(15) & Space(2) & Trim(Cmbcia.Text) & Space(30) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(2) & "REPORTE DE " & CmbConcepto.Text
Print #1, Space(30) & "PERIODO : " & Cmbmes.Text & " - " & Format(Txtano.Text, "0000")
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
mSemana = ""
mpag = 0
mtipoB = ""

mcad = "sum(d" & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo "

Sql$ = "select " & mcad & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'"

If (fAbrRst(Rs, Sql$)) Then
   If Rs!ded <> 0 And Rs!APO <> 0 Then
      If MsgBox("Desea Reporte de Aportacion : SI= Aportacion    NO= Deduccion", vbQuestion + vbYesNo + vbDefaultButton1, "Reportes Anuales") = vbNo Then mtipoB = "A" Else mtipoB = "D"
   ElseIf Rs!ded <> 0 Then
      mtipoB = "D"
   ElseIf Rs!APO <> 0 Then
      mtipoB = "A"
   End If
End If

If Rs.State = 1 Then Rs.Close

mcad = ""


mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False
Sql$ = "select status,adicional from placonstante where tipomovimiento='03' and codinterno='" & VConcepto & "' and status <>'*'"
If (fAbrRst(Rs, Sql$)) Then If Rs!status = "F" Or Rs!adicional = "S" Then mfijo = True
If Rs.State = 1 Then Rs.Close

If mfijo = False Then
    'Ingresos Afectos para Aportacion Normal
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadIA = ""
       Do While Not Rs.EOF
          mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
    If VConcepto = "04" Then 'ONP
        'Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND (i05=0 OR i06 =0 ) Group by placod,tipotrab order by placod"
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then 'AFP
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'  Group by placod,tipotrab order by placod"
    End If
Else
    If VConcepto = "04" Then
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
        ' Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND (i05=0 OR i06 =0 ) Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    Else
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'  Group by placod,tipotrab order by placod"
    End If
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
'Cabeza_Lista_DeducApor (RS!tipotrab)
mItem = 1
Do While Not Rs.EOF
   Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Trim(Rs!PlaCod) & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PlaCod = Trim(Rs!PlaCod)
   
   'NUEVO CODIGO
   If XPLACOD = "" Then
    XPLACOD = "'" & Trim(Rs!PlaCod) & "'"
   Else
    XPLACOD = XPLACOD & ",'" & Trim(Rs!PlaCod) & "'"
   End If
   
   dat.Recordset!nombre = Trim(mcad)
   If mfijo = False Then
      If mcadIA <> "" Then dat.Recordset!sueldo = Rs!afectoa
   End If
   dat.Recordset!importe = Rs!APO
   dat.Recordset!AFP = Rs!AFP
   dat.Recordset.Update
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
'Vacaciones
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False

If mfijo = False Then
    'Ingresos Afectos para Aportacion Vacaciones
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='02'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadIA = ""
       Do While Not Rs.EOF
          mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
       If VConcepto = "04" Then
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp='01' and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       ElseIf VConcepto = "11" Then
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp<>'01' and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       Else
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'  and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       End If
Else
    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*'  Group by placod,tipotrab order by placod"
    End If
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!vaca = Rs!afectoa
      End If
      dat.Recordset!importe = Rs!APO
      dat.Recordset!AFP = Rs!AFP
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!vaca = Rs!afectoa
      End If
      dat.Recordset!importe = dat.Recordset!importe + Rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + Rs!AFP
      dat.Recordset.Update
      
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
'Gratificacion
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False

If mfijo = False Then
    'Ingresos Afectos para Aportacion Gratificacion
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='03'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadIA = ""
       Do While Not Rs.EOF
          mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND codafp='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND codafp<>'01' and d11<>0 Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and d04<>0 Group by placod,tipotrab order by placod"
    End If
    
Else

    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND codafp ='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
    End If
    
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Trim(Rs!PlaCod) + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Trim(Rs!PlaCod) & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!grati = Rs!afectoa
      End If
      dat.Recordset!importe = Rs!APO
      dat.Recordset!AFP = Rs!AFP
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!grati = Rs!afectoa
      End If
      
      If VConcepto = "04" Then
        dat.Recordset!importe = dat.Recordset!importe + Rs!ded 'apo -DED
      ElseIf VConcepto = "11" Then
        dat.Recordset!importe = dat.Recordset!importe + Rs!ded  'apo -DED
      End If
      
     ' dat.Recordset!importe = dat.Recordset!importe + rs!APO 'apo -DED
      dat.Recordset!AFP = dat.Recordset!AFP + Rs!AFP
      dat.Recordset.Update
      
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'Liquidaciones y Subsidios
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo "
If VConcepto = "04" Then
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND CODAFP='01' Group by placod,tipotrab order by placod"
ElseIf VConcepto = "11" Then
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND CODAFP<>'01' Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!liquid = Rs!afectoa
      End If
      dat.Recordset!importe = Rs!APO
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!liquid = Rs!afectoa
      End If
      dat.Recordset!importe = dat.Recordset!importe + Rs!APO
      dat.Recordset.Update
      
   End If
   Rs.MoveNext
Loop

'MODIFICACION DE RICARDO HINOSTROZA PARA AFP
If VConcepto = "11" Then
    Sql$ = "SELECT PLACOD," & _
            "SUM(D11) As DED " & _
            "From PLAHISTORICO" & _
            " where cia='" & wcia & "'" & _
            " and year(fechaproceso)=" & Val(Txtano.Text) & _
            " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' " & _
            " AND status<>'*' " & _
            " AND codafp<>'01' " & _
            " Group by placod,tipotrab " & _
            " order by placod"
        
    If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
        Do While Not Rs.EOF
        dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
        If Not dat.Recordset.NoMatch Then
              dat.Recordset.Edit
              dat.Recordset!importe = Rs!ded
              dat.Recordset.Update
           Rs.MoveNext
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
        " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' " & _
        "AND status<>'*' " & _
        "Group by placod,tipotrab " & _
        "order by placod"
        
        If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
        Do While Not Rs.EOF
        dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
        If Not dat.Recordset.NoMatch Then
              dat.Recordset.Edit
              dat.Recordset!importe = Rs!ded
              dat.Recordset.Update
           Rs.MoveNext
        End If
        Loop
        
 End If

If Rs.State = 1 Then Rs.Close
RUTA$ = App.Path & "\REPORTS\" & "BenAnual.txt"
Open RUTA$ For Output As #1
dat.Refresh
With dat.Recordset
If Not .EOF Then .MoveFirst
miten = 1
mtot = 0
Cabeza_Deduc_Anual (mtipoB)
Do While Not .EOF

NOM$ = Trim(!nombre)

If Len(NOM) = 26 Then NOM = NOM & Space(2)

   mcad = fCadNum(mItem, "####") & ".- " & lentexto(36, Trim(NOM)) & Space(1) & fCadNum(!sueldo, "####,###.00") & Space(2) & fCadNum(!AFP, "###,###.00") & Space(2)
   mcad = mcad & fCadNum(!grati, "####,###.00") & Space(2) & fCadNum(!vaca, "####,###.00") & Space(2) & fCadNum(!liquid, "####,###.00") & Space(2)
   mcad = mcad & fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!vaca) + Val(!liquid), "####,###.00") & Space(1) & fCadNum(!importe, "####,###.00")
   Print #1, mcad
   mtot = mtot + !importe
   mItem = mItem + 1
   mlinea = mlinea + 1
   If mlinea > 55 Then Print #1, SaltaPag: Call Cabeza_Deduc_Anual(mtipoB)
   .MoveNext
Loop
Print #1, Space(120) & "----------"
Print #1, Space(118) & fCadNum(mtot, "#,###,###.00")
End With
Close #1
Call Imprime_Txt("BenAnual.txt", RUTA$)
dat.Database.Execute "delete from Tmpdeduc"
End Sub
Private Sub Procesa_Deduc_Apor_Anuales1()
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
mSemana = ""
mpag = 0
mtipoB = ""

mcad = "sum(d" & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo "

Sql$ = "select " & mcad & "from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'"

If (fAbrRst(Rs, Sql$)) Then
   If Rs!ded <> 0 And Rs!APO <> 0 Then
      If MsgBox("Desea Reporte de Aportacion : SI= Aportacion    NO= Deduccion", vbQuestion + vbYesNo + vbDefaultButton1, "Reportes Anuales") = vbNo Then mtipoB = "A" Else mtipoB = "D"
   ElseIf Rs!ded <> 0 Then
      mtipoB = "D"
   ElseIf Rs!APO <> 0 Then
      mtipoB = "A"
   End If
End If

If Rs.State = 1 Then Rs.Close

mcad = ""


mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False
Sql$ = "select status,adicional from placonstante where tipomovimiento='03' and codinterno='" & VConcepto & "' and status <>'*'"
If (fAbrRst(Rs, Sql$)) Then If Rs!status = "F" Or Rs!adicional = "S" Then mfijo = True
If Rs.State = 1 Then Rs.Close

If mfijo = False Then
    'Ingresos Afectos para Aportacion Normal
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadIA = ""
       Do While Not Rs.EOF
          mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
    If VConcepto = "04" Then 'ONP
        'Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mtipo) + "%" & "' AND status<>'*' AND (i05=0 OR i06 =0 ) Group by placod,tipotrab order by placod"
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then 'AFP
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'  Group by placod,tipotrab order by placod"
    End If
Else
    If VConcepto = "04" Then
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
        ' Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mtipo) + "%" & "' and status<>'*' AND (i05=0 OR i06 =0 ) Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    Else
         Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'  Group by placod,tipotrab order by placod"
    End If
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
'Cabeza_Lista_DeducApor (RS!tipotrab)
mItem = 1
Do While Not Rs.EOF
   Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Trim(Rs!PlaCod) & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PlaCod = Trim(Rs!PlaCod)
   
   'NUEVO CODIGO
   If XPLACOD = "" Then
    XPLACOD = "'" & Trim(Rs!PlaCod) & "'"
   Else
    XPLACOD = XPLACOD & ",'" & Trim(Rs!PlaCod) & "'"
   End If
   
   dat.Recordset!nombre = Trim(mcad)
   If mfijo = False Then
      If mcadIA <> "" Then dat.Recordset!sueldo = Rs!afectoa
   End If
   dat.Recordset!importe = Rs!APO
   dat.Recordset!AFP = Rs!AFP
   dat.Recordset.Update
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
'Vacaciones
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False

If mfijo = False Then
    'Ingresos Afectos para Aportacion Vacaciones
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='02'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadIA = ""
       Do While Not Rs.EOF
          mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
       If VConcepto = "04" Then
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp='01' and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       ElseIf VConcepto = "11" Then
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' and codafp<>'01' and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       Else
            Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*'  and placod in(" & XPLACOD & ") Group by placod,tipotrab order by placod"
       End If
Else
    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp<>'01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' and codafp='01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*'  Group by placod,tipotrab order by placod"
    End If
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!vaca = Rs!afectoa
      End If
      dat.Recordset!importe = Rs!APO
      dat.Recordset!AFP = Rs!AFP
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!vaca = Rs!afectoa
      End If
      dat.Recordset!importe = dat.Recordset!importe + Rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + Rs!AFP
      dat.Recordset.Update
      
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
'Gratificacion
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as ded,sum(a" & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp"
mfijo = False

If mfijo = False Then
    'Ingresos Afectos para Aportacion Gratificacion
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='03'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06') and status<>'*'"
    If (fAbrRst(Rs, Sql$)) Then
       Rs.MoveFirst
       mcadIA = ""
       Do While Not Rs.EOF
          mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
          Rs.MoveNext
       Loop
       mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
    End If
    If Rs.State = 1 Then Rs.Close
    
    If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
    
    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND codafp='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
    End If
    
Else

    If VConcepto = "04" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND codafp ='01' Group by placod,tipotrab order by placod"
    ElseIf VConcepto = "11" Then
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND codafp<>'01' Group by placod,tipotrab order by placod"
    Else
        Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' Group by placod,tipotrab order by placod"
    End If
    
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Trim(Rs!PlaCod) + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Trim(Rs!PlaCod) & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!grati = Rs!afectoa
      End If
      dat.Recordset!importe = Rs!APO
      dat.Recordset!AFP = Rs!AFP
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!grati = Rs!afectoa
      End If
      
      If VConcepto = "04" Then
        dat.Recordset!importe = dat.Recordset!importe + Rs!ded 'apo -DED
      ElseIf VConcepto = "11" Then
        dat.Recordset!importe = dat.Recordset!importe + Rs!ded  'apo -DED
      End If
      
     ' dat.Recordset!importe = dat.Recordset!importe + rs!APO 'apo -DED
      dat.Recordset!AFP = dat.Recordset!AFP + Rs!AFP
      dat.Recordset.Update
      
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'Liquidaciones y Subsidios
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo "
If VConcepto = "04" Then
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND CODAFP='01' Group by placod,tipotrab order by placod"
ElseIf VConcepto = "11" Then
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' AND CODAFP<>'01' Group by placod,tipotrab order by placod"
Else
    Sql$ = "Select placod,tipotrab,sum(totaling) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso in('04','05') and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' Group by placod,tipotrab order by placod"
End If

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!liquid = Rs!afectoa
      End If
      dat.Recordset!importe = Rs!APO
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mfijo = False Then
         If mcadIA <> "" Then dat.Recordset!liquid = Rs!afectoa
      End If
      dat.Recordset!importe = dat.Recordset!importe + Rs!APO
      dat.Recordset.Update
      
   End If
   Rs.MoveNext
Loop

'MODIFICACION DE RICARDO HINOSTROZA PARA AFP
If VConcepto = "11" Then
    Sql$ = "SELECT PLACOD," & _
            "SUM(D11) As DED " & _
            "From PLAHISTORICO" & _
            " where cia='" & wcia & "'" & _
            " and year(fechaproceso)=" & Val(Txtano.Text) & _
            " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' " & _
            " AND status<>'*' " & _
            " AND codafp<>'01' " & _
            " Group by placod,tipotrab " & _
            " order by placod"
        
    If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
        Do While Not Rs.EOF
        dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
        If Not dat.Recordset.NoMatch Then
              dat.Recordset.Edit
              dat.Recordset!importe = Rs!ded
              dat.Recordset.Update
           Rs.MoveNext
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
        " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' " & _
        "AND status<>'*' " & _
        "Group by placod,tipotrab " & _
        "order by placod"
        
        If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
        Do While Not Rs.EOF
        dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
        If Not dat.Recordset.NoMatch Then
              dat.Recordset.Edit
              dat.Recordset!importe = Rs!ded
              dat.Recordset.Update
           Rs.MoveNext
        End If
        Loop
        
 End If

If Rs.State = 1 Then Rs.Close
RUTA$ = App.Path & "\REPORTS\" & "BenAnual.txt"
Open RUTA$ For Output As #1
dat.Refresh
With dat.Recordset
If Not .EOF Then .MoveFirst
miten = 1
mtot = 0
Cabeza_Deduc_Anual (mtipoB)
Do While Not .EOF

NOM$ = Trim(!nombre)

If Len(NOM) = 26 Then NOM = NOM & Space(2)

   mcad = fCadNum(mItem, "####") & ".- " & lentexto(36, Trim(NOM)) & Space(1) & fCadNum(!sueldo, "####,###.00") & Space(2) & fCadNum(!AFP, "###,###.00") & Space(2)
   mcad = mcad & fCadNum(!grati, "####,###.00") & Space(2) & fCadNum(!vaca, "####,###.00") & Space(2) & fCadNum(!liquid, "####,###.00") & Space(2)
   mcad = mcad & fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!vaca) + Val(!liquid), "####,###.00") & Space(1) & fCadNum(!importe, "####,###.00")
   Print #1, mcad
   mtot = mtot + !importe
   mItem = mItem + 1
   mlinea = mlinea + 1
   If mlinea > 55 Then Print #1, SaltaPag: Call Cabeza_Deduc_Anual(mtipoB)
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
        .Fields(11).AllowZeroLength = True
        .Fields.Append .CreateField("fingreso", dbText)
        .Fields(12).AllowZeroLength = True
        .Fields.Append .CreateField("tipo_doc", dbText)
        .Fields(0).AllowZeroLength = True
        .Fields.Append .CreateField("nro_doc", dbText)
        .Fields(0).AllowZeroLength = True
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
    Print #1, Cmbcia.Text & Space(10) & "SISTEMA NACION DE PENSIONES " & Txtano.Text
    Print #1, "ORDEN           NOMBRE                            SUELDO     INC. AFP    GRATIF.     VACACIO.     Liq/Subs  SNP.REM TOTAL     IMPORTE"
Case Is = "11"
    Print #1, Cmbcia.Text & Space(10) & IIf(mtipoB = "A", "APORTES A     ", "DEDUCCIONES A ") & Trim(CmbConcepto.Text) & " " & Txtano.Text
    Print #1, "ORDEN           NOMBRE                            SUELDO     INC. AFP    GRATIF.     VACACIO.     Liq/Subs  AFP.REM TOTAL     IMPORTE"
Case Is = "01"
    Print #1, Cmbcia.Text & Space(10) & IIf(mtipoB = "A", "APORTES A     ", "DEDUCCIONES A ") & Trim(CmbConcepto.Text) & " " & Txtano.Text
    Print #1, "ORDEN           NOMBRE                            SUELDO     INC. AFP    GRATIF.     VACACIO.     Liq/Subs  ESS.REM TOTAL     IMPORTE"
Case Else
    Print #1, Cmbcia.Text & Space(10) & IIf(mtipoB = "A", "APORTES A     ", "DEDUCCIONES A ") & Trim(CmbConcepto.Text) & " " & Txtano.Text
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
Dim mFactor As Currency
Dim t1 As Currency: Dim t2 As Currency: Dim t3 As Currency: Dim t4 As Currency
Dim t5 As Currency: Dim t6 As Currency: Dim t7 As Currency: Dim t8 As Currency

cadnombre = nombre()
mSemana = ""
mpag = 0
mtipoB = "D"
MUIT = 0
Sql$ = "select uit from plauit where cia='" & wcia & "' and ano='" & Txtano.Text & "' and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then MUIT = Rs!uit
If Rs.State = 1 Then Rs.Close

mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"

'Ingresos Afectos para Aportacion Normal
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
   Rs.MoveFirst
   mcadIA = ""
   Do While Not Rs.EOF
      mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
      Rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If Rs.State = 1 Then Rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
'Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='01' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"
Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' Group by placod,tipotrab order by placod"

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1

Do While Not Rs.EOF
   Sql$ = cadnombre & "placod,fcese from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PlaCod = Rs!PlaCod
   dat.Recordset!nombre = mcad
   If mcadIA <> "" Then dat.Recordset!sueldo = Rs!afectoa
   dat.Recordset!importe = Rs!APO
   dat.Recordset!AFP = Rs!AFP
   dat.Recordset!util = Rs!util
   dat.Recordset!afp03 = Rs!afp03
   If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = " "
   dat.Recordset.Update
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

RUTA$ = App.Path & "\REPORTS\" & "QtaAnual.txt"
Open RUTA$ For Output As #1
dat.Refresh
With dat.Recordset
If Not .EOF Then .MoveFirst
miten = 1
mtot = 0
Cabeza_QuInta_Anual
t1 = 0: t2 = 0: t3 = 0: t4 = 0
t5 = 0: t6 = 0: t7 = 0: t8 = 0

Dim Mi_tope As Double

Do While Not .EOF
   mtope = Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) - Val(!afp03)
   mFactor = 0
   Mi_tope = mtope - (MUIT * 7)
   Select Case Mi_tope '(mtope - (MUIT * 7))
          Case Is < (Round(MUIT * 27, 2) + 1)
               mFactor = Round(Mi_tope * 0.15, 2) 'Round((mtope - (MUIT * 7)) * 0.15, 2)
          Case Is < (Round(MUIT * 54, 2) + 1)
               mFactor = Round(((Mi_tope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2) 'Round((((mtope - (MUIT * 7)) - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
          Case Else
               mFactor = Round(((Mi_tope - (MUIT * 54)) * 0.3) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15) 'Round((((mtope - (MUIT * 7)) - (MUIT * 54)) * 0.3) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
    End Select
    If mFactor < 0 Then mFactor = 0
   mcad = !PlaCod & " " & !nombre & Space(1) & fCadNum(!sueldo, "####,###.00") & Space(1) & fCadNum(!util, "###,###.00") & Space(1) & fCadNum(!AFP, "###,###.00") & Space(1)
   mcad = mcad & fCadNum(!grati, "####,###.00") & Space(1) & fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util), "####,###.00") & Space(1) & fCadNum(Val(!afp03) * -1, "###,###.00") & Space(1)
   mcad = mcad & fCadNum(mtope, "####,###.00") & Space(1)
   'mcad = mcad & fCadNum(MUIT * 7, "###,###.00") & Space(1) & fCadNum(mtope - (MUIT * 7), "###,###.00") & Space(1) & fCadNum(mFactor, "###,###.00") & Space(1) & fCadNum(!importe, "###,###.00") & Space(1) & fCadNum(!importe - mFactor, "##,###.00") & Space(2)
   mcad = mcad & fCadNum(MUIT * 7, "###,###.00") & Space(1) & fCadNum(mtope - (MUIT * 7), "###,###.00") & Space(1) & fCadNum(mFactor, "###,###.00") & Space(1) & fCadNum(!importe, "###,###.00") & Space(1) & fCadNum(mFactor - !importe, "##,###.00") & Space(2)
   mcad = mcad & !cese
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
Print #1, Cmbcia.Text
Print #1, "REPORTE DE REMUNERACION ACUMULADA AL MES DE " & Cmbmes.Text & " " & Txtano.Text
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
Dim mTransfer As Currency
mTransfer = 0
cadnombre = nombre()
mSemana = ""
mpag = 0
mtipoB = "D"
VConcepto = "13"
MUIT = 0
Sql$ = "select uit from plauit where cia='" & wcia & "' and ano='" & Txtano.Text & "' and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then MUIT = Rs!uit
If Rs.State = 1 Then Rs.Close
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"
'mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05) as afp,sum(i18) as util,sum(i06) as afp03"



'Ingresos Afectos para Aportacion Normal
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='01'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
   Rs.MoveFirst
   mcadIA = ""
   Do While Not Rs.EOF
      mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
      Rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If Rs.State = 1 Then Rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso in('01','11') and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' "

Dim Id_Trab As String

If Chk.Value = True Then
    
    Id_Trab = Trim(CboTrabajador.BoundText)
    
    If Len(Id_Trab) > 6 Then MsgBox "Error de Usuario: Debe de seleccion un Trabajador.": Exit Sub
    
    Sql$ = Sql$ & "AND PLACOD = '" & Id_Trab & "' "
End If

Sql$ = Sql$ & "Group by placod,tipotrab order by placod"

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1

Do While Not Rs.EOF
   Sql$ = cadnombre & "tipo_doc,nro_doc,placod,fcese,fingreso from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
   
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
   dat.Recordset.AddNew
   dat.Recordset!PlaCod = Rs!PlaCod
   dat.Recordset!nro_doc = Trim(rs2!nro_doc & "")
   dat.Recordset!tipo_doc = Trim(rs2!tipo_doc & "")
   dat.Recordset!nombre = mcad
   If mcadIA <> "" Then dat.Recordset!sueldo = Rs!afectoa
   dat.Recordset!importe = Rs!APO
   dat.Recordset!AFP = Rs!AFP
   dat.Recordset!util = Rs!util
   dat.Recordset!afp03 = Rs!afp03
   If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = " "
   dat.Recordset!fIngreso = rs2!fIngreso
   
   dat.Recordset.Update
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
'Vacaciones
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"

'Ingresos Afectos para Aportacion Vacaciones
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='02'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
   Rs.MoveFirst
   mcadIA = ""
   Do While Not Rs.EOF
      mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
      Rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If Rs.State = 1 Then Rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab," & mcad & " from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' "

If Chk.Value = True Then
    
    Id_Trab = Trim(CboTrabajador.BoundText)
    
    If Len(Id_Trab) > 6 Then MsgBox "Error de Usuario: Debe de seleccion un Trabajador.": Exit Sub
    
    Sql$ = Sql$ & "AND PLACOD = '" & Id_Trab & "' "
End If

Sql$ = Sql$ & "Group by placod,tipotrab order by placod"

'"Group by placod,tipotrab order by placod"

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "placod,fcese,fingreso from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nro_doc = Trim(Rs!nro_doc & "")
      dat.Recordset!tipo_doc = Trim(Rs!tipo_doc & "")
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = Rs!afectoa
      dat.Recordset!importe = Rs!APO
      dat.Recordset!AFP = Rs!AFP
      dat.Recordset!afp03 = Rs!afp03
      dat.Recordset!util = Rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = ""
      dat.Recordset!fIngreso = rs2!fIngreso
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!sueldo = dat.Recordset!sueldo + Rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + Rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + Rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + Rs!afp03
      dat.Recordset!util = dat.Recordset!util + Rs!util
      dat.Recordset.Update
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
'Gratificacion
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i05+i06) as afp,sum(i18) as util,sum(i06) as afp03"


'Ingresos Afectos para Aportacion Gratificacion
Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & mtipoB & "' and tboleta='03'  and  codigo='" & VConcepto & "' and cod_remu not in('05','06','18') and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
   Rs.MoveFirst
   mcadIA = ""
   Do While Not Rs.EOF
      mcadIA = mcadIA & "i" & Trim(Rs!cod_remu) & "+"
      Rs.MoveNext
   Loop
   mcadIA = "(" & Mid(mcadIA, 1, Len(Trim(mcadIA)) - 1) & ")"
End If
If Rs.State = 1 Then Rs.Close

If mcadIA <> "" Then mcad = mcad & ",sum" & mcadIA & " as afectoa"
Sql$ = "Select placod,tipotrab, " & mcad & " from plahistorico where cia='" & wcia & "' and proceso='03' and year(fechaproceso)=" & Val(Txtano.Text) & "  and tipotrab LIKE '" & Trim(mTipo) + "%" & "' AND status<>'*' AND PROCESO <>'05' "


If Chk.Value = True Then
    
    Id_Trab = Trim(CboTrabajador.BoundText)
    
    If Len(Id_Trab) > 6 Then MsgBox "Error de Usuario: Debe de seleccion un Trabajador.": Exit Sub
    
    Sql$ = Sql$ & "AND PLACOD = '" & Id_Trab & "' "
End If

Sql$ = Sql$ & "Group by placod,tipotrab order by placod"

If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
mItem = 1
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "tipo_doc,nro_doc,placod,fcese,fingreso from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!grati = Rs!afectoa
      dat.Recordset!importe = Rs!APO
      dat.Recordset!AFP = Rs!AFP
      dat.Recordset!afp03 = Rs!afp03
      dat.Recordset!util = Rs!util
      If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = Null
      dat.Recordset!fIngreso = rs2!fIngreso
      dat.Recordset.Update
   Else
   '***************************************************************
   ' MODIFICA   : 10/01/2008
   ' MOTIVO     : ERROR EN LA ASIGNACION DE SUMA SE COMENTA EL CODIGO
   ' MODIFICADO POR : RICARDO HINOSTROZA
   '****************************************************************
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!grati = Rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + Rs!APO
      dat.Recordset!AFP = dat.Recordset!AFP + Rs!AFP
      dat.Recordset!afp03 = dat.Recordset!afp03 + Rs!afp03
      dat.Recordset!util = dat.Recordset!util ' + rs!AFP
      dat.Recordset.Update
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'Liquidaciones y Subsidios
mcad = "sum(" & mtipoB & Format(VConcepto, "00") & ") as apo,sum(i18) as util "
Sql$ = "Select placod,tipotrab,sum(totaling-(i18+i41+i45)) as afectoa," & mcad & "from plahistorico where cia='" & wcia & "' and proceso ='04' and year(fechaproceso)=" & Val(Txtano.Text) & " and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and status<>'*' " ','05'

If Chk.Value = True Then
    
    Id_Trab = Trim(CboTrabajador.BoundText)
    
    If Len(Id_Trab) > 6 Then MsgBox "Error de Usuario: Debe de seleccion un Trabajador.": Exit Sub
    
    Sql$ = Sql$ & "AND PLACOD = '" & Id_Trab & "' "
End If

Sql$ = Sql$ & "Group by placod,tipotrab order by placod"


If Not (fAbrRst(Rs, Sql$)) Then Else Rs.MoveFirst
Do While Not Rs.EOF
   If Not dat.Recordset.BOF Then dat.Recordset.MoveFirst
   dat.Recordset.FindFirst "placod = """ + Rs!PlaCod + """"
   If dat.Recordset.NoMatch Then
      Sql$ = cadnombre & "tipo_doc,nro_doc,placod,fcese,fingreso from planillas where cia='" & wcia & "' and placod='" & Rs!PlaCod & "' and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(40, Left(rs2!nombre, 40)) Else mcad = Space(40)
      dat.Recordset.AddNew
      dat.Recordset!PlaCod = Rs!PlaCod
      dat.Recordset!nombre = mcad
      If mcadIA <> "" Then dat.Recordset!sueldo = Rs!afectoa
      dat.Recordset!importe = Rs!APO
      dat.Recordset!util = Rs!util
      If Trim(rs2!fcese & "") <> "" Then If IsDate(rs2!fcese) Then dat.Recordset!cese = Format(rs2!fcese, "dd/mm/yyyy") Else dat.Recordset!cese = ""
      dat.Recordset!fIngreso = rs2!fIngreso
      dat.Recordset.Update
   Else
      dat.Recordset.Edit
      If mcadIA <> "" Then dat.Recordset!sueldo = dat.Recordset!sueldo + Rs!afectoa
      dat.Recordset!importe = dat.Recordset!importe + Rs!APO
      dat.Recordset!util = dat.Recordset!util + Rs!util
      dat.Recordset.Update
      
   End If
   Rs.MoveNext
Loop
Dim IREM_NETA As Double

IREM_NETA = 0

If Rs.State = 1 Then Rs.Close
RUTA$ = App.Path & "\REPORTS\" & "QtaAnual.txt"
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
   
   Dim mQtaTransfer As Currency
   mTransfer = 0: mQtaTransfer = 0
   'Transferencia de Otra Empresa
   Sql$ = "select sum(totaling) as ing,sum(d13) as Qta from plahistorico where cia='" & wcia & "' and proceso='10' and year(fechaproceso)=" & Val(Txtano.Text) & "  and placod='" & Trim(!PlaCod) & "' and status<>'*' "
   If (fAbrRst(Rs, Sql$)) Then
      If Trim(Rs(0) & "") <> "" Then mTransfer = Rs(0)
      If Trim(Rs(1) & "") <> "" Then mQtaTransfer = Rs(1)
   End If
   Rs.Close
   
   'CAMBIAR para 3%
   mtope = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) + mTransfer - Val(!afp03)) - ((MUIT * 7))
   'mtope = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) + mTransfer) - ((MUIT * 7))
   
   'CAMBIAR para 3%
   IREM_NETA = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util)) - Val(!afp03) + mTransfer
   'IREM_NETA = (Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util)) + mTransfer
   
   mFactor = 0
   If Val(Txtano.Text) > 2014 Then
      Select Case mtope
          Case Is < (Round(MUIT * 5, 2) + 1)
               mFactor = Round((mtope) * 0.08, 2)
          Case Is < (Round(MUIT * 20, 2) + 1)
               mFactor = Round(((mtope - (MUIT * 5)) * 0.14) + (MUIT * 5) * 0.08, 2)
          Case Is < (Round(MUIT * 35, 2) + 1)
               mFactor = Round(((mtope - (MUIT * 20)) * 0.17) + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08)
          Case Is < (Round(MUIT * 45, 2) + 1)
               mFactor = Round(((mtope - (MUIT * 35)) * 0.2) + ((MUIT * 35) - (MUIT * 20)) * 0.17 + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08)
          Case Else
               mFactor = Round(((mtope - (MUIT * 45)) * 0.3) + ((MUIT * 45) - (MUIT * 35)) * 0.2 + ((MUIT * 35) - (MUIT * 20)) * 0.17 + ((MUIT * 20) - (MUIT * 5)) * 0.14, 2) + ((MUIT * 5) * 0.08)
      End Select
   Else
      Select Case mtope
          Case Is < (Round(MUIT * 27, 2) + 1)
               mFactor = Round((mtope) * 0.15, 2)  '- (muit * 7)
          Case Is < (Round(MUIT * 54, 2) + 1)
               mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
          Case Else
               mFactor = Round(((mtope - (MUIT * 54)) * 0.3) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
      End Select
   End If
    
    If mFactor < 0 Then mFactor = 0
    Print #1, Chr(218) & String(46, Chr(196)) & Chr(194) & String(31, Chr(196)) & Chr(191)
    Print #1, Chr(179) & "CERTIFICADO DE RETENCIONES SOBRE RENTAS DE    " & Chr(179) & "      EJERCICIO GRAVABLE       " & Chr(179)
    Print #1, Chr(179) & "5ta. CATEGORIA (Art. 45 del Reglamento del    " & Chr(179) & Space(31) & Chr(179)
    Print #1, Chr(179) & "Impuesto a la Renta (D.S No 122-94-EF)        " & Chr(179) & Space(13) & Txtano.Text & Space(14) & Chr(179)
    Print #1, Chr(192) & String(46, Chr(196)) & Chr(193) & String(31, Chr(196)) & Chr(217)
    Print #1,
    Print #1, Chr(218) & String(78, Chr(196)) & Chr(191)
    Print #1, Chr(179) & "Razon Social del Empleador: " & lentexto(32, Left(Cmbcia.Text, 32)) & "   RUC " & wruc & Chr(179)
    
    Print #1, Chr(179) & "Domicilio    :  AV. UNIVERSITARIA No. 6330-Z.I. INFANTAS" & Space(22) & Chr(179)
    Print #1, Chr(179) & "Representante Legal : INFANTES POMAR ANGEL FEDERICO     DNI  25535997" & Space(9) & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & Space(34) & "CERTIFICA" & Space(35) & Chr(179)
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    
    If Trim(!tipo_doc & "") = "01" Then
       Print #1, Chr(179) & "Que al Sr(a): " & lentexto(46, Left(!nombre, 46)) & RTrim(!PlaCod) & " DNI " & Trim(!nro_doc & "") & Chr(179)
    Else
       Print #1, Chr(179) & "Que al Sr(a): " & lentexto(55, Left(!nombre, 55)) & RTrim(!PlaCod) & " DNI " & Chr(179)
    End If
    
    Print #1, Chr(195) & String(78, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "se le ha retenido por concepto de Impuesto a la renta el importe(S/ " & IIf(RTrim(!PlaCod) = "E0508", fCadNum(2231.18, "#,###,###.00"), fCadNum(CCur(!importe) + mQtaTransfer, "#,###,###.00")) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(194) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "1.- RENTAS BRUTAS" & Space(48) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Transferencia         " & Space(39) & Chr(179) & fCadNum(mTransfer, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    
    Print #1, Chr(179) & "    Sueldo o Jornal Basico" & Space(39) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(43451.19, "#,###,###.00"), fCadNum(!sueldo, "#,###,###.00")) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    
    Print #1, Chr(179) & "    Comisiones            " & Space(39) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Clausula de Salvaguarda" & Space(38) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Participacion en las Utilidades" & Space(30) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(2804.9, "#,###,###.00"), fCadNum(!util, "#,###,###.00")) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Incremento AFP" & Space(47) & Chr(179) & fCadNum(!AFP, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Gratificaciones" & Space(46) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(7623.77, "#,###,###.00"), fCadNum(!grati, "#,###,###.00")) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "REMUNERACION BRUTA TOTAL    " & Space(37) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(53879.86, "#,###,###.00"), fCadNum(Val(!sueldo) + Val(!AFP) + Val(!grati) + Val(!util) + mTransfer, "#,###,###.00")) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    'Cambiar para 3%
    'Print #1, Chr(179) & "(-)AFP 3% (Art.71 DL 25897 6/12/92 Art.74 DS 054 97-ef 14/05/97) " & Chr(179) & fCadNum(!afp03, "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & "                                                                 " & Chr(179) & Space(12) & Chr(179)
    
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    'Print #1, Chr(179) & "REMUNERACION NETA PARA QUINTA CATEGORIA" & Space(26) & Chr(179) & fCadNum(mtope, "#,###,###.00") & Chr(179)
    'IREM_NETA
    Print #1, Chr(179) & "REMUNERACION NETA PARA QUINTA CATEGORIA" & Space(26) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(53879.86, "#,###,###.00"), fCadNum(IREM_NETA, "#,###,###.00")) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "2.- DEDUCCIONES SOBRE LA RENTA DE 5TA CATEGORIA" & Space(18) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    7 UIT (S/. " & fCadNum(MUIT, "###,###.00") & ")" & Space(39) & Chr(179) & fCadNum(MUIT * 7, "#,###,###.00") & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    'Print #1, Chr(179) & "3.RENTA NETA IMPONIBLE (1-2)" & Space(37) & Chr(179) & fCadNum(mtope - (MUIT * 7), "#,###,###.00") & Chr(179)
    Print #1, Chr(179) & "3.RENTA NETA IMPONIBLE (1-2)" & Space(37) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(24829.86, "#,###,###.00"), fCadNum(mtope, "#,###,###.00")) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "4.IMPUESTO A LA RENTA" & Space(44) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(2231.18, "#,###,###.00"), fCadNum(mFactor, "#,###,###.00")) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "5.CREDITOS CONTRA EL IMPUESTO" & Space(36) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Credito por dividendos" & Space(39) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & Space(65) & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(179) & "    Credito por donaciones (tasa media sobre el monto computable)" & Chr(179) & Space(12) & Chr(179)
    Print #1, Chr(195) & String(65, Chr(196)) & Chr(197) & String(12, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "6.TOTAL DEL IMPUESTO RETENIDO" & Space(36) & Chr(179) & IIf(RTrim(!PlaCod) = "E0508", fCadNum(2231.18, "#,###,###.00"), fCadNum(CCur(!importe) + mQtaTransfer, "#,###,###.00")) & Chr(179)
    Print #1, Chr(192) & String(65, Chr(196)) & Chr(193) & String(12, Chr(196)) & Chr(217)
    Print #1, "Lima, 31 de Diciembre " & Str(Val(Txtano.Text))
    Print #1,
    Print #1,
    Print #1, Space(36) & String(24, Chr(196)) & Space(2) & String(15, Chr(196))
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
Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
If Not IsNull(Rs!Ingreso) Then totalanon = Rs!neto: totalano = Rs!Ingreso
If Rs.State = 1 Then Rs.Close

Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
If Not IsNull(Rs!Ingreso) Then totalmes = Rs!Ingreso: totalmesn = Rs!neto
If Rs.State = 1 Then Rs.Close


RUTA$ = App.Path & "\REPORTS\" & "CuadroIVF.txt"
Open RUTA$ For Output As #1
Print #1, LetraChica
If mTipo = "01" Then
   Print #1, Space(41) & "CUADRO IV F - " & Cmbmes.Text & " " & Txtano.Text & Space(42) & "Pag.  17"
   Print #1, Space(41) & "--------------------------------"
   Print #1, Space(36) & "RESUMEN MENSUAL DE PLANILLAS DE SUELDOS" & Space(33) & "No." & Format(Cmbmes.ListIndex + 1, "00") & Txtano.Text
   Print #1, Space(36) & "--------------------------------------"
Else
   Print #1, Space(41) & "CUADRO IV A - " & Cmbmes.Text & " " & Txtano.Text & Space(42) & "Pag.  12"
   Print #1, Space(41) & "--------------------------------------"
   Print #1, Space(36) & "RESUMEN MENSUAL DE PLANILLAS DE JORNALES" & Space(33) & "No." & Format(Cmbmes.ListIndex + 1, "00") & Txtano.Text
   Print #1, Space(36) & "----------------------------------------"
End If
Print #1, "** Expresado en SOLES **"
Print #1, String(117, "-")
Print #1,
Print #1, "        CONCEPTO                 MENSUAL          %            ACUMULADO         %              PROMEDIO         %"
Print #1,
Print #1, String(117, "-")
Sql$ = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' order by codinterno"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
tano = 0: tmes = 0
Do While Not Rs.EOF
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(I" & Rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!Ingreso) Then
      If rs2!Ingreso <> 0 Then
         mcad = fCadNum(rs2!Ingreso, "#,###,###,###.00") & "     " & fCadNum(rs2!Ingreso * 100 / totalano, "###.00") & "     "
         mcad = mcad & fCadNum(rs2!Ingreso / (Cmbmes.ListIndex + 1), "#,###,###,###.00") & "     " & fCadNum(rs2!Ingreso * 100 / totalano, "###.00") & " "
         tano = tano + rs2!Ingreso
         tanoporc = tanoporc + (rs2!Ingreso * 100 / totalano)
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(I" & Rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!Ingreso) Then
         mcad = fCadNum(rs2!Ingreso, "#,###,###,###.00") & "     " & fCadNum(rs2!Ingreso * 100 / totalmes, "###.00") & "     " & mcadprint
         tmes = tmes + rs2!Ingreso
         tmesporc = tmesporc + (rs2!Ingreso * 100 / totalmes)
      Else
        mcad = fCadNum(0, "#,###,###,###.00") & "     " & fCadNum(0, "###.00") & "     " & mcad
      End If
      mcad = lentexto(26, Left(Rs!Descripcion, 26)) & mcad
   
      Print #1, mcad
      If rs2.State = 1 Then rs2.Close
   End If
   Rs.MoveNext
Loop
Print #1,
Print #1, "                              ------------     ------         ------------     ------         ------------     -----"
Print #1, "  T O T A L" & Space(15) & fCadNum(tmes, "#,###,###,###.00") & "     " & fCadNum(tmesporc, "###.00") & "     " & fCadNum(tano, "#,###,###,###.00") & "     " & fCadNum(tmesporc, "###.00") & "     " & fCadNum(tano / (Cmbmes.ListIndex + 1), "#,###,###,###.00") & "     " & fCadNum(tmesporc, "###.00")
Print #1,
Print #1,
Print #1, "Cuota Patronal"
Print #1, "--------------"
Print #1,
If Rs.State = 1 Then Rs.Close

'Aportaciones

Sql$ = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and status<>'*' order by codinterno"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
tano = 0: tmes = 0

Rs.MoveFirst
Do While Not Rs.EOF
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(A" & Rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!Ingreso) Then
      If rs2!Ingreso <> 0 Then
         mcad = fCadNum(rs2!Ingreso, "#,###,###,###.00") & "                "
         mcad = mcad & fCadNum(rs2!Ingreso / (Cmbmes.ListIndex + 1), "#,###,###,###.00") & "            "
         tano = tano + rs2!Ingreso
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(A" & Rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!Ingreso) Then
         mcad = fCadNum(rs2!Ingreso, "#,###,###,###.00") & "                 " & mcadprint
         tmes = tmes + rs2!Ingreso
      Else
        mcad = fCadNum(0, "#,###,###,###.00") & "                " & mcad
      End If
      mcad = lentexto(26, Left(Rs!Descripcion, 26)) & mcad
   
      Print #1, mcad
      If rs2.State = 1 Then rs2.Close
   End If
   Rs.MoveNext
Loop

Print #1,
Print #1, "Descuento Personal"
Print #1, "------------------"
Print #1,

'Deducciones
Rs.MoveFirst
Do While Not Rs.EOF
   printyn = False
   'Acumulado
   mcad = ""
   mcadprint = ""
   mcad = "select sum(D" & Rs!codinterno & ") as ingreso "
   Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
   If Not IsNull(rs2!Ingreso) Then
      If rs2!Ingreso <> 0 Then
         mcad = fCadNum(rs2!Ingreso, "#,###,###,###.00") & "                "
         mcad = mcad & fCadNum(rs2!Ingreso / (Cmbmes.ListIndex + 1), "#,###,###,###.00") & "            "
         tano = tano + rs2!Ingreso
         printyn = True
      End If
   End If
   mcadprint = mcad
   If rs2.State = 1 Then rs2.Close
   'Mensual
   If printyn = True Then
      mcad = "select sum(D" & Rs!codinterno & ") as ingreso "
      Sql$ = mcad & "from plahistorico where cia='" & wcia & "' and tipotrab LIKE '" & Trim(mTipo) + "%" & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
      If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
      If Not IsNull(rs2!Ingreso) Then
         mcad = fCadNum(rs2!Ingreso, "#,###,###,###.00") & "                 " & mcadprint
         tmes = tmes + rs2!Ingreso
      Else
        mcad = fCadNum(0, "#,###,###,###.00") & "                " & mcad
      End If
      mcad = lentexto(26, Left(Rs!Descripcion, 26)) & mcad
   
      Print #1, mcad
      If rs2.State = 1 Then rs2.Close
   End If
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
tmes = tmes + totalmesn
tano = tano + totalanon
Print #1,
Print #1, "Neto Percibido Personal   " & fCadNum(totalmesn, "#,###,###,###.00") & Space(17) & fCadNum(totalanon, "#,###,###,###.00") & Space(16) & fCadNum(totalanon / (Cmbmes.ListIndex + 1), "#,###,###,###.00")
Print #1, "                              ------------                    ------------                    ------------"
Print #1, "Gasto Total Empresa" & Space(7) & fCadNum(tmes, "#,###,###,###.00") & "                " & fCadNum(tano, "#,###,###,###.00") & "                " & fCadNum(tano / (Cmbmes.ListIndex + 1), "#,###,###,###.00")
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
    Print #1, Space(70) & "Mes de " & Cmbmes.Text & " del " & Txtano
    Print #1, ""
    For i_Contador_Titulo = 0 To 8
        Print #1, ""
    Next
End Sub
Private Sub Cargar_Cuenta_Banco()
Dim Cuenta As String

VBcoPago = fc_CodigoComboBox(CmbBcoPago, 2)
Sql$ = "usp_pla_PlaBcoCta '" & wcia & "','" & VBcoPago & "','','','','" & wuser & "',4"
Me.cbo_Cuenta.Clear
If fAbrRst(Rs, Sql$) Then
    Rs.MoveFirst
    Do While Not Rs.EOF
        Cuenta = Trim(Mid(Trim(Rs!moneda), 1, 3)) & Space(3 - Len(Trim(Mid(Trim(Rs!moneda), 1, 3))))
        Cuenta = Cuenta & Space(1) & Trim(Rs!cuentabco)
        cbo_Cuenta.AddItem Trim(Cuenta)
        Rs.MoveNext
    Loop
End If
Rs.Close
Set Rs = Nothing

End Sub

Private Sub Procesa_Archivo_Banco(banco As String, mano As Integer, mmes As Integer, msem As String)
Dim March As String
Dim nFil As Integer
Dim lDesTipoBol As String
Dim rs2 As ADODB.Recordset
Dim lCta As String

VBcoPago = fc_CodigoComboBox(CmbBcoPago, 2)


If FrameEmple.Visible Then
   If OptBoleta.Value = False And OptQuincena.Value = False Then
      MsgBox "Indique Si es Quincena o Fin de Mes", vbInformation
      Exit Sub
   End If
End If
If banco = "**" Then
   March = "PHEFCT.XLS"
Else
   'March = "PH" & wcia & banco & ".XLS"
   March = "BCOCREDITO.XLS"
   
   'Archivo de Texto
   Dim lCad As String
   Dim lCheckSum As Double
   lCheckSum = 0
   lCta = ""
   
   Sql$ = "select cuentabco from PlaBcoCta where cia='" & wcia & "' and status<>'*'"
   If fAbrRst(rs2, Sql$) Then lCta = Trim(rs2(0) & "")
   rs2.Close
   If Len(lCta) < 14 Then lCta = Mid(lCta, 1, 3) & Llenar_Ceros(Mid(lCta, 4, Len(lCta) - 6), 8) & Right(lCta, 3)
   lCad = "#1HC" & lCta & Space(6) & "S/"
   lCheckSum = Val(Mid(lCta, 4, 15))
End If

'mOrigen$ = Path_Reports & March
'mOrigen$ = App.Path & "\REPORTS\BCOCREDITO.XLS"
mOrigen$ = App.Path & "\REPORTS\" & March

Set fso = CreateObject("Scripting.FileSystemObject")

If Not (fso.FileExists(mOrigen$)) Then MsgBox "El archivo " & March & Chr(13) & "No se encuentra": Exit Sub

Mdestino$ = App.Path & "\Reports\" & March
CopyFile mOrigen, Mdestino, FILE_NOTIFY_CHANGE_LAST_WRITE

Sql$ = "select distinct(p.pagobanco) "

Sql$ = "select q.semana,q.proceso,q.placod,q.fechaproceso," _
     & "(select tipo_doc from planillas where placod=q.placod and status<>'*') as td," _
     & "(select nro_doc from planillas where placod=q.placod and status<>'*') as nd," _
     & "(select AP_PAT from planillas where placod=q.placod and status<>'*') as APEP," _
     & "(select AP_MAT from planillas where placod=q.placod and status<>'*') as APEM," _
     & "(select LTRIM(RTRIM(NOM_1))+' '+LTRIM(RTRIM(NOM_2)) from planillas where placod=q.placod and status<>'*') as NOM, " _
     & "(select pagonumcta from planillas where placod=q.placod and status<>'*') as Cuenta," _
     & "q.totneto,q.Proceso,q.PlaCod "

If FrameEmple.Visible And OptQuincena.Value Then Sql$ = Sql$ & "from plaquincena q,planillas p " Else Sql$ = Sql$ & "from plahistorico q,planillas p "
Sql$ = Sql$ & "where q.cia='" & wcia & "' and year(q.fechaproceso)=" & mano & " and month(q.fechaproceso)=" & mmes & " "

If mtipobol = "03" Then
   If banco = "**" Then Sql$ = Sql$ & "and q.proceso='03' " Else Sql$ = Sql$ & "and proceso LIKE '" & Trim(mtipobol) + "%" & "' "
Else
   If banco = "**" Then Sql$ = Sql$ & "and q.proceso IN('01','05') " Else Sql$ = Sql$ & "and proceso LIKE '" & Trim(mtipobol) + "%" & "' "
End If


If mTipo <> "01" And mTipo <> "" Then Sql$ = Sql$ & "and semana LIKE '" & Trim(msem) + "%" & "' "
If banco <> "**" Then Sql$ = Sql$ & "and pagobanco='" & banco & "' "
Sql$ = Sql$ & "and p.tipotrabajador='" & fc_CodigoComboBox(Cmbtipo, 2) & "' and q.status<>'*' and p.status<>'*' and p.cia=q.cia and p.placod=q.placod"

If Not (fAbrRst(rs2, Sql$)) Then rs2.Close: Set rs2 = Nothing: Exit Sub
rs2.MoveFirst

Set xlApp = GetObject(App.Path & "\Reports\" & March)

Set xlApp2 = xlApp.Application
For I = 1 To xlApp2.Workbooks.count
    If xlApp2.Workbooks(I).Name = March Then
       Set xlBook = xlApp2.Workbooks(I)
       xlApp2.Workbooks(I).Activate
       numwindows = I
       GoTo Continua
    End If
Next I

Continua:

xlApp2.Application.Visible = True
xlApp2.Parent.Windows(March).Visible = True

If banco = "**" Then
   Set xlSheet = xlApp2.Worksheets("Efectivo")
Else
   Set xlSheet = xlApp2.Worksheets("Deposito")
End If
xlSheet.Activate

If banco = "**" Then
   xlSheet.Cells(1, 4).Value = Cmbcia.Text
   If Trim(rs2!semana & "") <> "" Then xlSheet.Cells(2, 2).Value = "SEMANA No. ": xlSheet.Cells(2, 4).Value = "'" & Trim(rs2!semana & "")
   xlSheet.Cells(1, 8).Value = "'" & Format(Day(rs2!FechaProceso), "00") & "/" & Format(Month(rs2!FechaProceso), "00") & "/" & Format(Year(rs2!FechaProceso), "0000")
End If
If banco = "**" Then nFil = 8 Else nFil = 37

'Total Planilla
Dim lCount As Integer
lCount = 0
If banco <> "**" Then
   Dim lTotPla As Double
   lTotPla = 0
   Do While Not rs2.EOF
      If Trim(rs2!Cuenta & "") <> "" And rs2!Proceso <> "05" Then
         lTotPla = lTotPla + rs2!totneto
         lCta = Replace(Trim(rs2!Cuenta & ""), "-", "")
         lCheckSum = lCheckSum + Val(Mid(lCta, 4, 15))
         lCount = lCount + 1
       End If
      rs2.MoveNext
   Loop
End If
'Fin Total Planilla


If rs2.RecordCount > 0 Then rs2.MoveFirst
Do While Not rs2.EOF
   If banco = "**" Then
      lDesTipoBol = ""
      If rs2!Proceso = "01" Then lDesTipoBol = "Normal"
      If rs2!Proceso = "09" Then lDesTipoBol = "VACACIONES PROVISIONADAS"
      If rs2!Proceso = "10" Then lDesTipoBol = "TRANSFERENCIA"
      If rs2!Proceso = "04" Then lDesTipoBol = "LIQUIDACION"
      If rs2!Proceso = "05" Then lDesTipoBol = "SUBSIDIO"
      If rs2!Proceso = "07" Then lDesTipoBol = "DEPOSITO CTS"
      If rs2!Proceso = "08" Then lDesTipoBol = "VACACIONES PAGADAS"
      If rs2!Proceso = "02" Then lDesTipoBol = "VACACIONES"
      If rs2!Proceso = "03" Then lDesTipoBol = "GRATIFICACION"
   
      If Trim(rs2!Cuenta & "") = "" Or rs2!Proceso = "05" Then
         xlSheet.Cells(nFil, 2).Value = nFil - 7
         If Trim(rs2!td & "") = "01" Then xlSheet.Cells(nFil, 3).Value = "'" & Trim(rs2!nd & "")
         xlSheet.Cells(nFil, 4).Value = Trim(rs2!PlaCod & "")
         xlSheet.Cells(nFil, 5).Value = Trim(rs2!apep & "")
         xlSheet.Cells(nFil, 6).Value = Trim(rs2!apem & "")
         xlSheet.Cells(nFil, 7).Value = Trim(rs2!NOM & "")
         xlSheet.Cells(nFil, 8).Value = rs2!totneto
         xlSheet.Cells(nFil, 9).Value = lDesTipoBol
         nFil = nFil + 1
      End If
   Else
      If Trim(rs2!Cuenta & "") <> "" And rs2!Proceso <> "05" Then
         If Trim(rs2!td & "") = "01" Then
            xlSheet.Cells(nFil, 5).Value = "1"
            xlSheet.Cells(nFil, 6).Value = Trim(rs2!nd & "")
         End If
         xlSheet.Cells(nFil, 8).Value = Trim(rs2!apep & "")
         xlSheet.Cells(nFil, 9).Value = Trim(rs2!apem & "")
         xlSheet.Cells(nFil, 10).Value = Trim(rs2!NOM & "")
         xlSheet.Cells(nFil, 11).Value = Trim(rs2!Cuenta & "")
         xlSheet.Cells(nFil, 13).Value = rs2!totneto
         lCad = " 2A"
         lCta = Replace(Trim(rs2!Cuenta & ""), "-", "")
         If Len(lCta) < 14 And Len(lCta) > 6 Then lCta = Mid(lCta, 1, 3) & Llenar_Ceros(Mid(lCta, 4, Len(lCta) - 6), 8) & Right(lCta, 3)
         lCad = lCad & lCta & Space(6)
         lCta = Trim(rs2!apep & "") & " " & Trim(rs2!apem & "") & " " & Trim(rs2!NOM & "")
         lCta = Replace(Trim(lCta & ""), "Ñ", "N")
         lCad = lCad & lentexto(40, lCta) & "S/"
         lCad = lCad & Format(Int(rs2!totneto), "0000000000000") & Format((rs2!totneto - Int(rs2!totneto)) * 100, "00")
         lCad = lCad & Space(40) & "0"
         
         If Trim(rs2!td & "") = "01" Then
            lCad = lCad & "DNI"
            lCad = lCad & lentexto(12, Trim(rs2!nd & ""))
         Else
            lCad = lCad & Space(3 + 12)
         End If
         nFil = nFil + 1
      Else
         PagEfect = True
      End If
   End If
   rs2.MoveNext
Loop
End Sub
Private Sub Archivo_Bancos()
Dim mano As Integer
Dim mmes As Integer
Dim msem As String

mmes = Cmbmes.ListIndex + 1
mano = Val(Txtano.Text)
msem = Txtsemana.Text
PagEfect = False
Sql$ = "select distinct(p.pagobanco) "
If FrameEmple.Visible And OptQuincena.Value Then Sql$ = Sql$ & "from plaquincena q,planillas p " Else Sql$ = Sql$ & "from plahistorico q,planillas p "
Sql$ = Sql$ & "where q.cia='" & wcia & "' and year(q.fechaproceso)=" & mano & " and month(q.fechaproceso)=" & mmes & " "
If mtipobol = "01" Then Sql$ = Sql$ & "and q.proceso IN('01','05') " Else Sql$ = Sql$ & "and proceso LIKE '" & Trim(mtipobol) + "%" & "' "
If mTipo <> "01" And mTipo <> "" Then Sql$ = Sql$ & "and semana LIKE '" & Trim(msem) + "%" & "' "
Sql$ = Sql$ & "and q.status<>'*' and p.status<>'*' and p.cia=q.cia and p.placod=q.placod"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
Do While Not Rs.EOF
   If Trim(Rs(0) & "") <> "" Then Call Procesa_Archivo_Banco(Rs(0), mano, mmes, msem)
   Rs.MoveNext
Loop
Rs.Close: Set Rs = Nothing
If PagEfect Then Call Procesa_Archivo_Banco("**", mano, mmes, msem)
MsgBox "Se Generaron los Archivos", vbInformation
End Sub
Private Sub Procesa_Cuadro_IV()
Dim vDias As Integer
vDias = InputBox("Ingrese Número de días", "Cuadro IV")
CuadroIV ("01")
CuadroIV ("02")
CuadroIVH (vDias)
End Sub
Private Sub CuadroIVH(mdias As Integer)
Dim Sql As String

Sql = "Usp_Pla_CuadroIV_Horas '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & "," & mdias & ""

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

Set xlSheet = xlApp2.Worksheets("HOJA3")
xlSheet.Name = "IVOH"

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 1
xlSheet.Range("B:B").ColumnWidth = 14
xlSheet.Range("C:C").ColumnWidth = 6
xlSheet.Range("D:M").ColumnWidth = 15

xlSheet.Range("C:C").HorizontalAlignment = xlCenter

xlSheet.Range("D:O").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(2, 2).Value = "CUADRO IV"
xlSheet.Cells(3, 2).Value = "MES DE " & Cmbmes.Text & " DE " & Txtano.Text

xlSheet.Range(xlSheet.Cells(2, 2), xlSheet.Cells(2, 13)).Merge
xlSheet.Range(xlSheet.Cells(2, 2), xlSheet.Cells(2, 13)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 13)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(1, 13)).HorizontalAlignment = xlCenter

xlSheet.Cells(5, 2).Value = "** Expresado en SOLES **"
xlSheet.Cells(5, 13).Value = "Pag.11"
xlSheet.Cells(5, 13).HorizontalAlignment = xlRight

xlSheet.Cells(6, 2).Value = "Sección"
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(8, 2)).Merge
xlSheet.Cells(6, 3).Value = "Infor."
xlSheet.Range(xlSheet.Cells(6, 3), xlSheet.Cells(8, 3)).Merge
xlSheet.Cells(6, 4).Value = "# de Obre."
xlSheet.Cells(7, 4).Value = "por"
xlSheet.Cells(8, 4).Value = "Semana"
xlSheet.Cells(6, 5).Value = "Indice"
xlSheet.Cells(7, 5).Value = "de"
xlSheet.Cells(8, 5).Value = "Requeri."
xlSheet.Cells(6, 6).Value = "Horas Trabajo Neto"
xlSheet.Range(xlSheet.Cells(6, 6), xlSheet.Cells(6, 7)).Merge
xlSheet.Cells(8, 6).Value = "Jor.Nor"
xlSheet.Cells(8, 7).Value = "Sob. T."
xlSheet.Cells(6, 8).Value = "Total Horas"
xlSheet.Cells(7, 8).Value = "Efectiva"
xlSheet.Cells(8, 8).Value = "Trabajadas"
xlSheet.Cells(6, 9).Value = "Horas Efecti."
xlSheet.Cells(7, 9).Value = "Trabajadas"
xlSheet.Cells(8, 9).Value = "%"
xlSheet.Cells(6, 10).Value = "SOBRE-"
xlSheet.Cells(8, 10).Value = "TASA"
xlSheet.Cells(6, 11).Value = "Comp.por"
xlSheet.Cells(7, 11).Value = "Sobre-"
xlSheet.Cells(8, 11).Value = "tiempo"
xlSheet.Cells(6, 12).Value = "Total N.Soles"
xlSheet.Cells(7, 12).Value = "sin (Dominical,"
xlSheet.Cells(8, 12).Value = "Subsidio,Feriado)"
xlSheet.Cells(6, 13).Value = "Promedio"
xlSheet.Cells(8, 13).Value = "N.Soles/horas"

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(6, 13)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(8, 2), xlSheet.Cells(8, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(36, 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 13), xlSheet.Cells(36, 13)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(36, 2), xlSheet.Cells(36, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(36, 2)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(36, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 3), xlSheet.Cells(36, 3)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 4), xlSheet.Cells(36, 4)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 5), xlSheet.Cells(36, 5)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 6), xlSheet.Cells(36, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 7), xlSheet.Cells(36, 7)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 8), xlSheet.Cells(36, 8)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 9), xlSheet.Cells(36, 9)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 10), xlSheet.Cells(36, 10)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 11), xlSheet.Cells(36, 11)).Borders(xlEdgeRight).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 12), xlSheet.Cells(36, 12)).Borders(xlEdgeRight).LineStyle = xlContinuous

xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(8, 13)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(8, 13)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(1, 2), xlSheet.Cells(8, 13)).Font.Bold = True

nFil = 9
Dim lHomDia As Double
Dim lJorPrmM As Double
Dim lJorPrmA As Double

Do While Not Rs.EOF
   If Mid(Trim(Rs!tipo & ""), 2, 1) = "A" Then xlSheet.Cells(nFil, 2).Value = Mid(Trim(Rs!DesCosto & ""), 4, 20)
   xlSheet.Cells(nFil, 3).Value = Mid(Trim(Rs!tipo & ""), 2, 1)
   If Rs!NumObr <> 0 Then xlSheet.Cells(nFil, 4).Value = Rs!NumObr
   If Rs!IndReq <> 0 Then xlSheet.Cells(nFil, 5).Value = Rs!IndReq
   If Rs!JorHor <> 0 Then xlSheet.Cells(nFil, 6).Value = Rs!JorHor
   If Rs!HExtra <> 0 Then xlSheet.Cells(nFil, 7).Value = Rs!HExtra
   If Rs!totalh <> 0 Then xlSheet.Cells(nFil, 8).Value = Rs!totalh
   If Rs!PorcHr <> 0 Then xlSheet.Cells(nFil, 9).Value = Rs!PorcHr
   If Rs!Sobret <> 0 Then xlSheet.Cells(nFil, 10).Value = Rs!Sobret
   If Rs!ComSob <> 0 Then xlSheet.Cells(nFil, 11).Value = Rs!ComSob
   If Rs!totsol <> 0 Then xlSheet.Cells(nFil, 12).Value = Rs!totsol
   If Rs!PromSo <> 0 Then xlSheet.Cells(nFil, 13).Value = Rs!PromSo
   nFil = nFil + 1
   If Mid(Trim(Rs!tipo & ""), 2, 1) = "M" And Rs!ccosto = "99" Then lHomDia = Rs!JorHor
   If Mid(Trim(Rs!tipo & ""), 2, 1) = "M" And Rs!ccosto = "99" Then lJorPrmM = Round((Rs!totsol / Rs!totalh) * 8, 2)
   If Mid(Trim(Rs!tipo & ""), 2, 1) = "A" And Rs!ccosto = "99" Then lJorPrmA = Round((Rs!totsol / Rs!totalh) * 8, 2)
   Rs.MoveNext
Loop
nFil = nFil + 2
xlSheet.Cells(nFil, 4).Value = "HOMBRE DIA"
xlSheet.Cells(nFil, 5).Value = Round(lHomDia / (mdias * 8), 2)
xlSheet.Cells(nFil, 12).Value = "JORNAL PROMEDIO"
nFil = nFil + 1
xlSheet.Cells(nFil, 11).Value = "MENSUAL"
xlSheet.Cells(nFil + 1, 11).Value = "ANUAL"

xlSheet.Cells(nFil, 12).Value = lJorPrmM
xlSheet.Cells(nFil + 1, 12).Value = lJorPrmA

Rs.Close: Set Rs = Nothing

nFil = nFil + 3
xlSheet.Cells(nFil, 7).Value = "Hecho Por"
xlSheet.Cells(nFil, 7).HorizontalAlignment = xlCenter
xlSheet.Cells(nFil, 7).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Cells(nFil, 10).Value = "Revisado Por"
xlSheet.Cells(nFil, 10).HorizontalAlignment = xlCenter
xlSheet.Cells(nFil, 10).Borders(xlEdgeTop).LineStyle = xlContinuous

'Trabajadores que no tienen horas trabajadas pero tiene ingresos

Sql = "Usp_Pla_Cuadra_CuadroIV '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ""
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
   nFil = nFil + 3
   xlSheet.Cells(nFil, 4).Value = "Trabajadores Sin horas trabajadas con remuneraciones"
   nFil = nFil + 2
   xlSheet.Cells(nFil, 4).Value = "CODIGO"
   xlSheet.Cells(nFil, 5).Value = "INGRESOS"
   xlSheet.Cells(nFil, 6).Value = "DOMINICAL"
   xlSheet.Cells(nFil, 7).Value = "FERIADO"
   xlSheet.Cells(nFil, 8).Value = "SUBS.MAT"
   xlSheet.Cells(nFil, 9).Value = "SUBS.ENF"
   xlSheet.Cells(nFil, 10).Value = "ENF.PAG"
   xlSheet.Cells(nFil, 11).Value = "TOTAL"
   nFil = nFil + 1
Do While Not Rs.EOF
   xlSheet.Cells(nFil, 4).Value = Trim(Rs!PlaCod & "")
   xlSheet.Cells(nFil, 5).Value = Rs!INGRESOS
   xlSheet.Cells(nFil, 6).Value = Rs!dominical * -1
   xlSheet.Cells(nFil, 7).Value = Rs!feriado * -1
   xlSheet.Cells(nFil, 8).Value = Rs!subsmat * -1
   xlSheet.Cells(nFil, 9).Value = Rs!SubsEnf * -1
   xlSheet.Cells(nFil, 10).Value = Rs!EnfPag * -1
   xlSheet.Cells(nFil, 11).Value = (Rs!INGRESOS - Rs!dominical - Rs!feriado - Rs!subsmat - Rs!SubsEnf)
   Rs.MoveNext
   nFil = nFil + 1
Loop
Rs.Close: Set Rs = Nothing

Dim I As Integer
For I = 1 To 3
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridlines = False
Next

xlApp2.Sheets(1).Select
xlApp2.Application.Visible = True

Screen.MousePointer = vbDefault
End Sub

Private Sub CuadroIV(TipoTrab As String)
Dim Sql As String

Sql = "Usp_Pla_CuadroIV '" & wcia & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & TipoTrab & "'"

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

If TipoTrab = "02" Then
   Set xlSheet = xlApp2.Worksheets("HOJA2")
   xlSheet.Name = "IVO"
Else
   Set xlApp1 = CreateObject("Excel.Application")
   xlApp1.Workbooks.Add
   Set xlApp2 = xlApp1.Application
   xlApp2.Sheets.Add
    
   xlApp2.Sheets("Hoja1").Select
   xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
   xlApp2.Sheets("Hoja2").Select
   xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)
   xlApp2.Sheets("Hoja3").Select
   xlApp2.Sheets("Hoja3").Move Before:=xlApp2.Sheets(3)

   Set xlBook = xlApp2.Workbooks(1)
   Set xlSheet = xlApp2.Worksheets("HOJA1")
   xlSheet.Name = "IVE"
End If

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 2
xlSheet.Range("B:B").ColumnWidth = 36
xlSheet.Range("C:H").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

If TipoTrab = "01" Then
   xlSheet.Cells(1, 2).Value = "CUADRO IV F - " & Cmbmes.Text & " " & Txtano.Text
   xlSheet.Cells(1, 8).Value = "Pag. 17"
   xlSheet.Cells(3, 2).Value = "RESUMEN  MENSUAL  DE  PLANILLAS DE SUELDOS"
   xlSheet.Cells(3, 8).Value = "No." & Format(Cmbmes.ListIndex + 1, "00") & "." & Txtano.Text
Else
   xlSheet.Cells(1, 2).Value = "CUADRO IV A - " & Cmbmes.Text & " " & Txtano.Text
   xlSheet.Cells(1, 8).Value = "Pag. 12"
   xlSheet.Cells(3, 2).Value = "RESUMEN  MENSUAL  DE  PLANILLAS DE SALARIOS"
   xlSheet.Cells(3, 8).Value = "No." & Format(Cmbmes.ListIndex + 1, "00") & "." & Txtano.Text
End If
xlSheet.Range(xlSheet.Cells(1, 2), xlSheet.Cells(1, 7)).Merge
xlSheet.Range(xlSheet.Cells(1, 2), xlSheet.Cells(1, 7)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(3, 7)).Merge
xlSheet.Range(xlSheet.Cells(3, 2), xlSheet.Cells(1, 7)).HorizontalAlignment = xlCenter

xlSheet.Cells(5, 2).Value = "** Expresado en SOLES **"

xlSheet.Range(xlSheet.Cells(7, 2), xlSheet.Cells(7, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(7, 2), xlSheet.Cells(7, 8)).Borders(xlEdgeTop).Weight = xlMedium

xlSheet.Cells(8, 2).Value = "CONCEPTO"
xlSheet.Cells(8, 3).Value = "MENSUAL"
xlSheet.Cells(8, 4).Value = "%"
xlSheet.Cells(8, 5).Value = "ACUMULADO"
xlSheet.Cells(8, 6).Value = "%"
xlSheet.Cells(8, 7).Value = "PROMEDIO"
xlSheet.Cells(8, 8).Value = "%"

xlSheet.Range(xlSheet.Cells(8, 2), xlSheet.Cells(8, 8)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(1, 2), xlSheet.Cells(8, 8)).Font.Bold = True

xlSheet.Range(xlSheet.Cells(9, 2), xlSheet.Cells(9, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(9, 2), xlSheet.Cells(9, 8)).Borders(xlEdgeBottom).Weight = xlMedium

Dim lTotMesGto As Double
Dim lTotAcumGto As Double
Dim lTotMesNeto As Double
Dim lTotAcumNeto As Double

lTotMesGto = 0: lTotAcumGto = 0: lTotMesNeto = 0: lTotAcumNeto = 0

nFil = 11
Dim msum As Integer
msum = 1
Do While Not Rs.EOF
   If Trim(Rs!tipo & "") = "I" Then
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!descon & "")
      If Rs!Mes <> 0 Then xlSheet.Cells(nFil, 3).Value = Rs!Mes
      If Rs!porcm <> 0 Then xlSheet.Cells(nFil, 4).Value = Rs!porcm
      If Rs!acum <> 0 Then xlSheet.Cells(nFil, 5).Value = Rs!acum
      If Rs!porcA <> 0 Then xlSheet.Cells(nFil, 6).Value = Rs!porcA
      If Rs!Promedio <> 0 Then xlSheet.Cells(nFil, 7).Value = Rs!Promedio
      If Rs!porcP <> 0 Then xlSheet.Cells(nFil, 8).Value = Rs!porcP
      nFil = nFil + 1
      msum = msum + 1
      lTotMesGto = lTotMesGto + Rs!Mes
      lTotAcumGto = lTotAcumGto + Rs!acum
      lTotMesNeto = lTotMesNeto + Rs!Mes
      lTotAcumNeto = lTotAcumNeto + Rs!acum
   End If
   Rs.MoveNext
Loop

xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 8)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 8)).Borders(xlEdgeBottom).Weight = xlMedium

nFil = nFil + 1
msum = msum * -1

xlSheet.Cells(nFil, 2) = "TOTAL REMUNERACIONES"
xlSheet.Cells(nFil, 2).Font.Bold = True
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Cells(nFil, 3).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 4).Value = 100
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 6).Value = 100
xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).Value = 100

nFil = nFil + 2
xlSheet.Cells(nFil, 2) = "APORTACIONES DEL EMPLEADOR"
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2
If Rs.RecordCount > 0 Then Rs.MoveFirst
Do While Not Rs.EOF
   If Trim(Rs!tipo & "") = "A" Then
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!descon & "")
      If Rs!Mes <> 0 Then xlSheet.Cells(nFil, 3).Value = Rs!Mes
      If Rs!porcm <> 0 Then xlSheet.Cells(nFil, 4).Value = Rs!porcm
      If Rs!acum <> 0 Then xlSheet.Cells(nFil, 5).Value = Rs!acum
      If Rs!porcA <> 0 Then xlSheet.Cells(nFil, 6).Value = Rs!porcA
      If Rs!Promedio <> 0 Then xlSheet.Cells(nFil, 7).Value = Rs!Promedio
      If Rs!porcP <> 0 Then xlSheet.Cells(nFil, 8).Value = Rs!porcP
      nFil = nFil + 1
      msum = msum + 1
      lTotMesGto = lTotMesGto + Rs!Mes
      lTotAcumGto = lTotAcumGto + Rs!acum
   End If
   Rs.MoveNext
Loop

nFil = nFil + 1
xlSheet.Cells(nFil, 2) = "DEDUCCIONES AL TRABAJADOR"
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2
If Rs.RecordCount > 0 Then Rs.MoveFirst
Do While Not Rs.EOF
   If Trim(Rs!tipo & "") = "D" Then
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!descon & "")
      If Rs!Mes <> 0 Then xlSheet.Cells(nFil, 3).Value = Rs!Mes
      If Rs!porcm <> 0 Then xlSheet.Cells(nFil, 4).Value = Rs!porcm
      If Rs!acum <> 0 Then xlSheet.Cells(nFil, 5).Value = Rs!acum
      If Rs!porcA <> 0 Then xlSheet.Cells(nFil, 6).Value = Rs!porcA
      If Rs!Promedio <> 0 Then xlSheet.Cells(nFil, 7).Value = Rs!Promedio
      If Rs!porcP <> 0 Then xlSheet.Cells(nFil, 8).Value = Rs!porcP
      nFil = nFil + 1
      msum = msum + 1
      lTotMesNeto = lTotMesNeto - Rs!Mes
      lTotAcumNeto = lTotAcumNeto - Rs!acum
     
   End If
   Rs.MoveNext
Loop

nFil = nFil + 1
If TipoTrab = "01" Then
   xlSheet.Cells(nFil, 2) = "NETO PERCIBIDO EMPLEADO"
Else
   xlSheet.Cells(nFil, 2) = "NETO PERCIBIDO OBRERO"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True

If lTotMesNeto <> 0 Then xlSheet.Cells(nFil, 3).Value = lTotMesNeto
If lTotAcumNeto <> 0 Then xlSheet.Cells(nFil, 5).Value = lTotAcumNeto
If lTotAcumNeto <> 0 Then xlSheet.Cells(nFil, 7).Value = Round(lTotAcumNeto / (Cmbmes.ListIndex + 1), 2)

nFil = nFil + 2
xlSheet.Cells(nFil, 2) = "GASTO TOTAL EMPRESA"
xlSheet.Cells(nFil, 2).Font.Bold = True

If lTotMesGto <> 0 Then xlSheet.Cells(nFil, 3).Value = lTotMesGto
If lTotAcumGto <> 0 Then xlSheet.Cells(nFil, 5).Value = lTotAcumGto
If lTotAcumGto <> 0 Then xlSheet.Cells(nFil, 7).Value = Round(lTotAcumGto / (Cmbmes.ListIndex + 1), 2)

If TipoTrab = "02" Then
   nFil = nFil + 2
   Sql = "Select 'M' as Mes,Count(1) as sem From plasemanas where cia='" & wcia & "' and year(fechaf)=" & Txtano.Text & " and month(fechaf)=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
   Sql = Sql & " Union All Select 'A' as Mes,Count(1) as sem From plasemanas where cia='" & wcia & "' and year(fechaf)=" & Txtano.Text & " and month(fechaf)<=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
   If (fAbrRst(Rs, Sql)) Then
      xlSheet.Cells(nFil, 2).Value = "No. Semanas Registradas :"
      Rs.MoveFirst
      Do While Not Rs.EOF
         If Rs!Mes = "M" Then
            xlSheet.Cells(nFil, 3).Value = Rs!sem
            xlSheet.Cells(nFil, 3).NumberFormat = "#,###,##0;[Red](#,###,##0)"
         Else
            xlSheet.Cells(nFil, 5).Value = Rs!sem
            xlSheet.Cells(nFil, 5).NumberFormat = "#,###,##0;[Red](#,###,##0)"
         End If
         Rs.MoveNext
      Loop
   End If
End If

Rs.Close: Set Rs = Nothing

nFil = nFil + 6
xlSheet.Cells(nFil, 4).Value = "Hecho Por"
xlSheet.Cells(nFil, 4).HorizontalAlignment = xlCenter
xlSheet.Cells(nFil, 4).Borders(xlEdgeTop).LineStyle = xlContinuous
xlSheet.Cells(nFil, 6).Value = "Revisado Por"
xlSheet.Cells(nFil, 6).HorizontalAlignment = xlCenter
xlSheet.Cells(nFil, 6).Borders(xlEdgeTop).LineStyle = xlContinuous

Screen.MousePointer = vbDefault
End Sub
Private Sub Resumen_Planilla()
Dim rscargo As ADODB.Recordset
Dim I As Integer
Dim mcad As String
Dim mcadI As String
Dim mcadd As String
Dim mcada As String
Dim mcadh As String
Dim mfor As Integer
Dim mc As Integer
Dim mcargo As String
Dim mremun As Currency

Dim Rs As New ADODB.Recordset
Dim Inicio As Boolean

If CmbPlanta.Text = "TOTAL" Or CmbPlanta.ListIndex < 0 Then mPlanta = "*" Else mPlanta = fc_CodigoComboBox(CmbPlanta, 2)

If Cmbtipo.ListIndex < 0 Then MsgBox "Debe Seleccionar Tipo de Trabajador", vbInformation, "Resumne de Planilla": Exit Sub
mmes = Cmbmes.ListIndex + 1
mano = Val(Txtano.Text)
msem = Txtsemana.Text

If mtipobol <> "01" And mtipobol <> "05" Then msem = ""
If mTipo = "01" Then msem = ""

'rpt = MsgBox("Desea Imprimir Titulo de Cabecera", vbYesNo)

'*****************desde aqui para la planilla************************************
lBol = mtipobol
lSem = msem
If mtipobol = "" Then lBol = "*"
If msem = "" Then lSem = "*"
Sql$ = "usp_Pla_Resumen_Planilla '" & wcia & "'," & mano & "," & mmes & ",'" & lBol & "','" & lSem & "','" & mTipo & "','" & mPlanta & "','*','N'"
'*****************************************************


If Not (Funciones.fAbrRst(Rs, Sql$)) Then MsgBox "No Existen Boletas Registradas Segun Parametros", vbCritical, "Resumen de Planillas": Exit Sub

Dim mArchRes As String
If mtipobol = "*" Then
   mArchRes = "RE" & wcia & mTipo & ".txt"
Else
   mArchRes = "RE" & wcia & mTipo & mtipobol & ".txt"
End If

RUTA$ = App.Path & "\REPORTS\" & mArchRes

Open RUTA$ For Output As #1

Rs.MoveFirst
mlinea = 60
numtra = 0
Inicio = True
Dim lCcosto As String
Dim lDesCosto As String
lCcosto = Trim(Rs!ccosto & "")
lDesCosto = Trim(Rs!DesCosto & "")
Do While Not Rs.EOF
   If lCcosto <> Trim(Rs!ccosto & "") Then
      Call Resumen_Total(lCcosto, lDesCosto)
      lCcosto = Trim(Rs!ccosto & "")
      lDesCosto = Trim(Rs!DesCosto & "")
      Call Cabeza_Resumen(lDesCosto)
      numtraCosto = 0
   End If
   If mlinea > 50 Then
   If Not Inicio Then
        Print #1, Chr(12) + Chr(13)
    Else
        Inicio = False
   End If

    Call Cabeza_Resumen(lDesCosto)
    
   End If

   mcargo = Trim(Rs!descargo & "")
   If IsNull(Rs!fcese) Then mcad = Space(10) Else mcad = Format(Rs!fcese, "dd/mm/yyyy")

   Print #1, Rs!PlaCod & Space(5) & lentexto(40, Left(Rs!nombre, 40)) & "  " & lentexto(20, Left(mcargo, 20)) & "  " & Format(Rs!fIngreso, "dd/mm/yyyy") & "  " & mcad & lentexto(15, Left(Rs!ipss, 15))

   numtra = numtra + 1
   numtraCosto = numtraCosto + 1
   'HORAS
   mfor = Len(Trim(Vcadh))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadh, I, 2))
       If Rs(mc - 1) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc - 1), "##,##0.00") & " "
       End If
   Next I
   Print #1, mcad
   mlinea = mlinea + 1
   
   'Ingresos Primera Linea
   mfor = Len(Trim(Vcadir))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadir, I, 2))
       If Rs(mc + 29) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc + 29), "##,##0.00") & " "
       End If
   Next I
   Print #1, mcad

   mlinea = mlinea + 1
   
   'Ingresos segunda Linea
   mfor = Len(Trim(Vcadir2))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadir2, I, 2))
       If Rs(mc + 29) = 0 Then
          mcad = mcad & Space(11)
       Else
          'mcad = mcad & " " & fCadNum(Rs(mc + 29), "###,##0.00") & " "
          'add 060924
          mcad = mcad & " " & fCadNum(Rs(mc + 29), "###,###,##0.00") & " "
       End If
   Next I
   Print #1, mcad
   mlinea = mlinea + 1

   'DEDUCCIONES Y APORTACIONES
   mfor = Len(Trim(Vcadd))
   mcad = ""
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcadd, I, 2))
       If Rs(mc + 79) = 0 Then
          mcad = mcad & Space(11)
       Else
          'mcad = mcad & " " & fCadNum(Rs(mc + 79), "##,##0.00") & " "
           'add 060924
          mcad = mcad & " " & fCadNum(Rs(mc + 79), "###,##0.00") & " "
          
       End If
   Next I
   mfor = Len(Trim(Vcada))
   For I = 1 To mfor Step 2
       mc = Val(Mid(Vcada, I, 2))
       If Rs(mc + 99) = 0 Then
          mcad = mcad & Space(11)
       Else
          mcad = mcad & " " & fCadNum(Rs(mc + 99), "##,##0.00") & " "
       End If
   Next I
   mcad = mcad & Space(4) & fCadNum(Rs!totaling, "###,###,##0.00") & " " & fCadNum(Rs!totalded, "###,###,##0.00") & " " & fCadNum(Rs!totneto, "###,###,##0.00")
   Print #1, mcad
   mlinea = mlinea + 1
   Print #1,
   mlinea = mlinea + 1
   Rs.MoveNext
Loop
Call Resumen_Total(lCcosto, lDesCosto)
Call Resumen_Total("", "")

Close #1
Call Funciones.Imprime_Txt(mArchRes, RUTA$)
End Sub

Private Sub Resumen_Total(ccosto As String, DesCosto As String)
Print #1,
Print #1, String(233, "=")
Print #1, Chr(12) + Chr(13)

Call Cabeza_Resumen(DesCosto)

'Call Imprimir_Titulo

Print #1,
Print #1, Space(10) & "***** TOTAL PLANILLA *****"

Print #1,
Print #1, "                           H O R A S                             R E M U N E R A C I O N E S                        D E D U C C I O N E S                           A P O R T A C I O N E S"
Print #1, "                           ---------                             ---------------------------                        ---------------------                           -----------------------"
If Rs.State = 1 Then Rs.Close

Dim mTotal As String
If ccosto = "" Then mTotal = "S" Else mTotal = "N"
If ccosto = "" Then ccosto = "*"
Sql$ = "usp_Pla_Resumen_Planilla '" & wcia & "'," & mano & "," & mmes & ",'" & lBol & "','" & lSem & "','" & mTipo & "','" & mPlanta & "','" & ccosto & "','" & mTotal & "'"
If (fAbrRst(Rs, Sql$)) Then
    Dim m As Integer
    m = 0
    mremun = 0
    
    For I = 1 To 102 Step 2
       m = m + 1
       
       'TOTAL HORAS
       mcad = Space(10)
       If I <= Len(Vcadh) Then
          If I = 1 Then
             mcad = mcad & Mid(Rcadh, 1, 20)
          Else
             mcad = mcad & Mid(Rcadh, 20 * (m - 1) + 1, 20)
          End If
          mcad = mcad & Space(5)
          mc = Val(Mid(Vcadh, I, 2))
          If Rs(mc - 1) = 0 Then
             mcad = mcad & Space(14)
          Else
             mcad = mcad & fCadNum(Rs(mc - 1), "###,###,##0.00")
          End If
       Else
          mcad = mcad & Space(39)
       End If
       
       '=========================
       
       'TOTAL REMUNERACIONES AFECTAS
       mcad = mcad & Space(10)
       If I <= Len(Vcadir + Vcadir2) Then
          If I = 1 Then
             mcad = mcad & Mid(Rcadir + Rcadir2, 1, 20)
          Else
             mcad = mcad & Mid(Rcadir + Rcadir2, 20 * (m - 1) + 1, 20)
          End If
          mcad = mcad & Space(5)
          mc = Val(Mid(Vcadir + Vcadir2, I, 2))
          If Rs(mc + 29) = 0 Then
             mcad = mcad & Space(14)
          Else
             mcad = mcad & fCadNum(Rs(mc + 29), "###,###,##0.00")
             mremun = mremun + Rs(mc + 29)
          End If
       Else
          mcad = mcad & Space(39)
       End If
       
       'TOTAL DEDUCCIONES
       mcad = mcad & Space(10)
       If I <= Len(Vcadd) Then
          If I = 1 Then
             mcad = mcad & Mid(Rcadd, 1, 20)
          Else
             mcad = mcad & Mid(Rcadd, 20 * (m - 1) + 1, 20)
          End If
          mcad = mcad & Space(5)
          mc = Val(Mid(Vcadd, I, 2))
          If Rs(mc + 79) = 0 Then
             mcad = mcad & Space(14)
          Else
             mcad = mcad & fCadNum(Rs(mc + 79), "###,###,##0.00")
          End If
       Else
          mcad = mcad & Space(24)
       End If
       
       'TOTAL APORTACIONES
       mcad = mcad & Space(10)
       If I <= Len(Vcada) Then
          If I = 1 Then
             mcad = mcad & Mid(Rcada, 1, 20)
          Else
             mcad = mcad & Mid(Rcada, 20 * (m - 1) + 1, 20)
          End If
          mc = Val(Mid(Vcada, I, 2))
          If Rs(mc + 100) = 0 Then
             mcad = mcad & Space(14)
          Else
             mcad = mcad & fCadNum(Rs(mc + 100), "###,###,##0.00")
          End If
       Else
          mcad = mcad & Space(14)
       End If
    
       If Trim(mcad) <> "" Then Print #1, mcad
    '   Debug.Print mcad
    Next
'    Print #1, Space(84) & "--------------"
'    Print #1, Space(59) & "Sub Total I          " & fCadNum(mremun, "###,###,###,##0.00")
'    Print #1, Space(65) & "NO REMUNERATIVOS"
'    Print #1, Space(65) & "----------------"
    mremun = 0
    
'    Dim Z As Integer
'    Z = 0
'    For I = 1 To Len(Vcadinr) Step 2
'
'       'TOTAL REMUNERACIONES NO AFECTAS
'       'MODIFICADO 05/08/2008 RICARDO HINOSTROZA
'       'MODIFICAR EL ACCESSO A LOS CONCEPTOS NO REMUNERATIVOS
'
'       mcad = ""
'       mcad = mcad & Space(59)
'       If I <= Len(Vcadinr) Then
'          If I = 1 Then
'             mcad = mcad & Mid(Rcadinr, 1, 20)
'          Else
'              Z = Z + 1
'              mcad = mcad & Mid(Rcadinr, 20 * (Z) + 1, 20)
'              'mcad = mcad & Mid(Rcadinr, 20 * (i - 2) + 1, 20)
'          End If
'
'          mcad = mcad & Space(5)
'          mc = Val(Mid(Vcadinr, I, 2))
'          If rs(mc + 29) = 0 Then
'             mcad = mcad & Space(14)
'          Else
'             mcad = mcad & fCadNum(rs(mc + 29), "###,###,##0.00")
'             mremun = mremun + rs(mc + 29)
'          End If
'       Else
'          mcad = mcad & Space(14)
'       End If
'       If Trim(mcad) <> "" Then Print #1, mcad
'    Next I
    
'    Print #1, Space(84) & "--------------"
'    Print #1, Space(59) & "Sub Total            " & fCadNum(mremun, "###,###,###,##0.00")
'    Print #1, Space(84) & "--------------                                   --------------                                 -----------"

    Print #1, Space(84) & "--------------                                   --------------                                 -----------"
    Print #1, Space(59) & "* TOTAL REMUNERACION *" & fCadNum(Rs!totaling, "##,###,###,##0.00") & Space(10) & "* TOTAL DEDUCCIONES * " & fCadNum(Rs!totalded, "##,###,###,##0.00") & Space(10) & "* TOTAL APORTACIONES *" & fCadNum(Rs!totalapo, "#,###,##0.00")
    Print #1,
    Print #1, Space(59) & "*** NETO PAGADO ***   " & fCadNum(Rs!totneto, "##,###,###,##0.00")
    Print #1, Space(59) & "======================================="
    Print #1,
    If mTotal = "S" Then
       Print #1, "Total de Trabajadores    =>  " & Str(numtra)
    Else
       Print #1, "Total de Trabajadores    =>  " & Str(numtraCosto)
    End If
    
    Print #1, Chr(12) + Chr(13)
End If

End Sub
Private Sub Cabeza_Resumen(lDesCosto As String)
Dim rsremu As ADODB.Recordset
Dim mcadir As String
Dim mcadinr As String
Dim mcadd As String
Dim mcada As String
Dim mcadh As String
Dim cad As String
Dim f1 As String
Dim f2 As String
f1 = "": f2 = ""
mcadir = "": mcadd = "": mcada = "": mcadir2 = "": mcadh = ""
Vcadir = "": Vcadd = "": Vcada = "": Vcadir2 = "": Vcadh = ""
Rcadir = "": Rcadd = "": Rcada = "": Rcadir2 = "": Rcadh = ""

'Fecha de semana
Sql$ = "select fechai,fechaf from plasemanas where cia='" & wcia & "' and ano='" & Format(Txtano.Text, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then f1 = Format(Rs!fechai, "dd/mm/yyyy"):     f2 = Format(Rs!fechaf, "dd/mm/yyyy")
If Rs.State = 1 Then Rs.Close
'Fin de Fecha de semana

If mTipo = "01" Then
   Rcadh = ""
   Vcadh = ""
Else
   Rcadh = ""
   Vcadh = ""
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
       mTipo & "' and a.status<>'*' " & MCADENA _
       & "and a.codigo=b.cod_maestro2 and codigo not in('22','24','08','21','09','26') order by codigo"

If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
Do While Not rs2.EOF
   mcadh = mcadh & " " & lentexto(9, Left(rs2!DESCRIP, 9)) & " "
   Rcadh = Rcadh & lentexto(20, Left(rs2!DESCRIP, 20))
   Vcadh = Vcadh & rs2!Codigo
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close

Sql$ = "PLASS_INGRESO_RESUMENPLANILLA '" & wcia & "','" & Trim(Txtano.Text) & "','" & Trim(Cmbmes.ListIndex + 1) & "'"

If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst
Dim lCount As Integer
lCount = 0
Do While Not rs2.EOF
   If lCount <= 22 Then
      mcadir = mcadir & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Rcadir = Rcadir & lentexto(20, Left(rs2!Descripcion, 20))
      Vcadir = Vcadir & rs2!Codigo
   Else
      mcadir2 = mcadir2 & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
      Rcadir2 = Rcadir2 & lentexto(20, Left(rs2!Descripcion, 20))
      Vcadir2 = Vcadir2 & rs2!Codigo
   End If
   lCount = lCount + 1
   rs2.MoveNext
Loop
If rs2.State = 1 Then rs2.Close

'DEDUCCIONES Y APORTACIONES
Sql$ = "select a.tipo,a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
     & "where a.cia='" & wcia & "' and a.tipo<>'I' and a.status<>'*' " _
     & "and a.cia=b.cia and b.status<>'*' and b.tipomovimiento='03' and a.tipo_trab='" & mTipo & "' and a.codigo=b.codinterno " _
     & "order by a.tipo,a.codigo"
    
If (fAbrRst(rs2, Sql$)) Then rs2.MoveFirst

Do While Not rs2.EOF
   Select Case rs2!tipo
          Case Is = "D"
               mcadd = mcadd & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Rcadd = Rcadd & lentexto(20, Left(rs2!Descripcion, 20))
               Vcadd = Vcadd & rs2!Codigo
          Case Is = "A"
               Rcada = Rcada & lentexto(20, Left(rs2!Descripcion, 20))
               mcada = mcada & " " & lentexto(9, Left(rs2!Descripcion, 9)) & " "
               Vcada = Vcada & rs2!Codigo
   End Select
   rs2.MoveNext
Loop

If rs2.State = 1 Then rs2.Close


    Print #1, Chr(18) + Trim(Cmbcia.Text) + Chr(14)
    Dim mPlanta As String
    If CmbPlanta.Text = "TOTAL" Or CmbPlanta.ListIndex < 0 Then mPlanta = "" Else mPlanta = " ( PLANTA : " & CmbPlanta.Text & " )"
    
    If mTipo = "01" Then
        Print #1, Space(40) & Chr(14) & "PLANILLA DE PAGO EMPLEADOS " & Cmbtipbol.Text & mPlanta & Chr(20)
    Else
        Print #1, Space(40) & Chr(14) & "PLANILLA DE PAGO OBREROS " & Cmbtipbol.Text & mPlanta & Chr(20)
    End If
    Print #1, Chr(18) + Chr(14)
    Print #1, Chr(20)

    'Print #1, Chr(20)
    Print #1, Space(50) & "MES de " & Cmbmes.Text & "  de " & Txtano.Text & Chr(15)
    If mTipo <> "01" And Val(Txtsemana.Text) > 0 Then
        Print #1, "F. Imp " & Format(Date, "dd/mm/yyyy") & Space(10) & "SEMANA No =>  " & Txtsemana.Text & "  Del " & f1 & " Al " & f2 & "  " & lDesCosto
    Else
        Print #1, "F. Imp " & Format(Date, "dd/mm/yyyy") & "  " & lDesCosto
    End If
    Print #1, String(253, "-")
    Print #1, "Codigo    Apellidos y Nombres del Trabajador        Ocupacion             F. Ingreso   F.cese   I.P.S.S."
    
    'Print #1, String(IIf(Len(mcadh) = 0, 10, Len(mcadh)) / 2 - 5, "-") & "   HORAS  " & String(IIf(Len(mcadh) = 0, 10, Len(mcadh)) / 2 - 5, "-")
    Print #1, String(121, "-") & " H O R A S " & String(121, "-")

    Print #1, mcadh
    
    Print #1, String(112, "-") & " R E M U N E R A C I O N E S " & String(112, "-")
    Print #1, mcadir
    If mcadir2 <> "" Then
       Print #1, mcadir2
    End If
    
    cad = String(Len(mcadd) / 2 - 8, "-") & "   DEDUCCIONES  " & String(Len(mcadd) / 2 - 8, "-") & "     "
    If Len(Trim(mcada)) > 25 Then
        cad = cad & String(Len(mcada) / 2 - 11, "-") & "  APORTACIONES  " & String(Len(mcada) / 2 - 11, "-")
    End If
    Print #1, cad
    Print #1, mcadd & Space(5) & mcada & "*** TOT.REM.    TOT. DED.   TOT. NETO"
    Print #1, String(253, "-")
    
mlinea = 16

End Sub
Private Sub Procesa_Seguro(ByVal tipoS As Integer)
Dim mtope As Double
Dim manotope As Integer
Dim mmmestope As Integer
Dim mperiodoTope As String

manotope = Val(Txtano.Text)
mmmestope = Cmbmes.ListIndex + 1

If mmmestope = 12 Then
   manotope = manotope + 1
   mmmestope = 1
Else
   mmmestope = mmmestope + 1
End If

mperiodoTope = Format(manotope, "0000") & Format(mmmestope, "00")

mtope = 0
Sql$ = "select top 1 Tope from plaSCTR where cia='" & wcia & "' and periodo='" & mperiodoTope & "' and status<>'*'"
If (fAbrRst(Rs, Sql$)) Then
   Rs.MoveFirst
   If Not IsNull(Rs!tope) Then mtope = Rs!tope
Else
   MsgBox "No se Ha Registrado el Tope Para el Periodo " & mperiodoTope, vbInformation, "Calculo de Seguro (SCRT)"
   Rs.Close: Set Rs = Nothing
   Exit Sub
End If

Sql$ = "select porsalud,porpension from plaSCTR where cia='" & wcia & "' and codSCTR='" & wcia & "' and periodo='" & Txtano.Text + Format(Cmbmes.ListIndex + 1, "00") & "' and status<>'*'"
If Not (fAbrRst(Rs, Sql$)) Then
   MsgBox "No se Han Registrado porcentajes para el Periodo", vbInformation, "Calculo de Seguro (SCRT)"
   Rs.Close: Set Rs = Nothing
   Exit Sub
End If

If mtope = 0 Then MsgBox "No se Ha Registrado el Tope Para el Periodo " & mperiodoTope, vbInformation, "Calculo de Seguro (SCRT)": Rs.Close: Set Rs = Nothing: Exit Sub

If tipoS = 1 Then
    Sql$ = "Usp_Pla_Seguro_Riesgo '" & wcia & "'," & Val(Txtano.Text) & "," & Cmbmes.ListIndex + 1 & ""
Else
    Sql$ = "Usp_Pla_Seguro_Riesgo2 '" & wcia & "'," & Val(Txtano.Text) & "," & Cmbmes.ListIndex + 1 & ""
End If
     
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
xlApp2.Sheets.Add
 
xlApp2.Sheets("Hoja1").Select
xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
xlApp2.Sheets("Hoja2").Select
xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)

Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "SALUD"

xlSheet.Range("G:G").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("Y:Y").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("A:G").NumberFormat = "@"
xlSheet.Range("P:S").NumberFormat = "@"
'xlSheet.Range("K:K").NumberFormat = "####,##0.00;[Red](####,##0.00)"
xlSheet.Range("G:G").HorizontalAlignment = xlLeft

xlSheet.Cells(1, 1).Value = "Tipo Documento"
xlSheet.Cells(1, 2).Value = "Documento de Identidad"
xlSheet.Cells(1, 3).Value = "Apellido Paterno"
xlSheet.Cells(1, 4).Value = "Apellido Materno"
xlSheet.Cells(1, 5).Value = "Primer nombre"
xlSheet.Cells(1, 6).Value = "Segundo nombre"
xlSheet.Cells(1, 7).Value = "fecha Nacimiento"
xlSheet.Cells(1, 8).Value = "Sexo"
xlSheet.Cells(1, 9).Value = "Nacionalidad"
xlSheet.Cells(1, 10).Value = "codigo Asegurado"
xlSheet.Cells(1, 11).Value = "OCUPACION"
xlSheet.Cells(1, 12).Value = "Departamento"
xlSheet.Cells(1, 13).Value = "Provincia"
xlSheet.Cells(1, 14).Value = "Distrito"
xlSheet.Cells(1, 15).Value = "Direccion"
xlSheet.Cells(1, 16).Value = "RUC"
xlSheet.Cells(1, 17).Value = "Nivel Riesgo"
xlSheet.Cells(1, 18).Value = "Mes de Planilla"
xlSheet.Cells(1, 19).Value = "moneda Sueldo"
xlSheet.Cells(1, 20).Value = "importe Sueldo"
xlSheet.Cells(1, 21).Value = "Condicion"
xlSheet.Cells(1, 22).Value = "Proy/Obra"
xlSheet.Cells(1, 23).Value = "Tipo Producto"
xlSheet.Cells(1, 24).Value = "Tipo Movimiento"
xlSheet.Cells(1, 25).Value = "Fecha Inicio Vigencia"
xlSheet.Cells(1, 26).Value = "moneda Prima"
xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 26)).Borders.LineStyle = xlContinuous
nFil = 2
Do While Not Rs.EOF
   If Trim(Rs!cesado & "") = "N" Then
  ' And (rs!sueldo) <> 0 Then
      xlSheet.Cells(nFil, 1).Value = Trim(Rs!td)
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!DNI)
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!ap_pat)
      xlSheet.Cells(nFil, 4).Value = Trim(Rs!ap_mat)
      xlSheet.Cells(nFil, 5).Value = Trim(Rs!nom_1)
      xlSheet.Cells(nFil, 6).Value = Trim(Rs!nom_2)
      xlSheet.Cells(nFil, 7).Value = Format(Day(Rs!fnacimiento), "00") & Format(Month(Rs!fnacimiento), "00") & Format(Year(Rs!fnacimiento), "0000")
      xlSheet.Cells(nFil, 8).Value = Trim(Rs!sexo)
      xlSheet.Cells(nFil, 9).Value = Trim(Rs!Nacion)
      xlSheet.Cells(nFil, 10).Value = Trim(Rs!PlaCod)
      xlSheet.Cells(nFil, 11).Value = Trim(Rs!Cargo)
      xlSheet.Cells(nFil, 12).Value = Trim(Rs!departamento)
      xlSheet.Cells(nFil, 13).Value = Trim(Rs!provincia)
      xlSheet.Cells(nFil, 14).Value = Trim(Rs!DISTRITO)
      xlSheet.Cells(nFil, 15).Value = Trim(Rs!Direccion)
      xlSheet.Cells(nFil, 16).Value = Trim(Rs!RUC)
      xlSheet.Cells(nFil, 17).Value = Trim(Rs!NivelRiesgo)
      xlSheet.Cells(nFil, 18).Value = Trim(Rs!Periodo)
      xlSheet.Cells(nFil, 19).Value = Trim(Rs!moneda)
      xlSheet.Cells(nFil, 20).Value = Rs!sueldo
      xlSheet.Cells(nFil, 21).Value = Trim(Rs!Condicion)
      xlSheet.Cells(nFil, 22).Value = ""
      xlSheet.Cells(nFil, 23).Value = "S"
      xlSheet.Cells(nFil, 24).Value = Trim(Rs!movi)
      xlSheet.Cells(nFil, 25).Value = Trim(Rs!FecIni)
      xlSheet.Cells(nFil, 26).Value = Trim(Rs!MonPrima)
      nFil = nFil + 1
   End If
   Rs.MoveNext
Loop

xlSheet.Range("A:Z").EntireColumn.AutoFit

'PENSION
If Rs.RecordCount > 0 Then Rs.MoveFirst
Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "PENSION"
If Rs.RecordCount > 0 Then Rs.MoveFirst

xlSheet.Range("G:G").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("Y:Y").NumberFormat = "dd/mm/yyyy;@"
xlSheet.Range("A:G").NumberFormat = "@"
xlSheet.Range("P:S").NumberFormat = "@"
xlSheet.Range("G:G").HorizontalAlignment = xlLeft

xlSheet.Cells(1, 1).Value = "Tipo Documento"
xlSheet.Cells(1, 2).Value = "Documento de Identidad"
xlSheet.Cells(1, 3).Value = "Apellido Paterno"
xlSheet.Cells(1, 4).Value = "Apellido Materno"
xlSheet.Cells(1, 5).Value = "Primer nombre"
xlSheet.Cells(1, 6).Value = "Segundo nombre"
xlSheet.Cells(1, 7).Value = "fecha Nacimiento"
xlSheet.Cells(1, 8).Value = "Sexo"
xlSheet.Cells(1, 9).Value = "Nacionalidad"
xlSheet.Cells(1, 10).Value = "codigo Asegurado"
xlSheet.Cells(1, 11).Value = "OCUPACION"
xlSheet.Cells(1, 12).Value = "Departamento"
xlSheet.Cells(1, 13).Value = "Provincia"
xlSheet.Cells(1, 14).Value = "Distrito"
xlSheet.Cells(1, 15).Value = "Direccion"
xlSheet.Cells(1, 16).Value = "RUC"
xlSheet.Cells(1, 17).Value = "Nivel Riesgo"
xlSheet.Cells(1, 18).Value = "Mes de Planilla"
xlSheet.Cells(1, 19).Value = "moneda Sueldo"
xlSheet.Cells(1, 20).Value = "importe Sueldo"
xlSheet.Cells(1, 21).Value = "Condicion"
xlSheet.Cells(1, 22).Value = "Proy/Obra"
xlSheet.Cells(1, 23).Value = "Tipo Producto"
xlSheet.Cells(1, 24).Value = "Tipo Movimiento"
xlSheet.Cells(1, 25).Value = "Fecha Inicio Vigencia"
xlSheet.Cells(1, 26).Value = "moneda Prima"
xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 26)).Borders.LineStyle = xlContinuous
nFil = 2
Do While Not Rs.EOF
   If Trim(Rs!cesado & "") = "N" Then
   'And (rs!sueldo) <> 0 Then
      xlSheet.Cells(nFil, 1).Value = Trim(Rs!td)
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!DNI)
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!ap_pat)
      xlSheet.Cells(nFil, 4).Value = Trim(Rs!ap_mat)
      xlSheet.Cells(nFil, 5).Value = Trim(Rs!nom_1)
      xlSheet.Cells(nFil, 6).Value = Trim(Rs!nom_2)
      xlSheet.Cells(nFil, 7).Value = Format(Day(Rs!fnacimiento), "00") & Format(Month(Rs!fnacimiento), "00") & Format(Year(Rs!fnacimiento), "0000")
      xlSheet.Cells(nFil, 8).Value = Trim(Rs!sexo)
      xlSheet.Cells(nFil, 9).Value = Trim(Rs!Nacion)
      xlSheet.Cells(nFil, 10).Value = Trim(Rs!PlaCod)
      xlSheet.Cells(nFil, 11).Value = Trim(Rs!Cargo)
      xlSheet.Cells(nFil, 12).Value = Trim(Rs!departamento)
      xlSheet.Cells(nFil, 13).Value = Trim(Rs!provincia)
      xlSheet.Cells(nFil, 14).Value = Trim(Rs!DISTRITO)
      xlSheet.Cells(nFil, 15).Value = Trim(Rs!Direccion)
      xlSheet.Cells(nFil, 16).Value = Trim(Rs!RUC)
      xlSheet.Cells(nFil, 17).Value = Trim(Rs!NivelRiesgo)
      xlSheet.Cells(nFil, 18).Value = Trim(Rs!Periodo)
      xlSheet.Cells(nFil, 19).Value = Trim(Rs!moneda)
      If Rs!sueldo > mtope Then
         xlSheet.Cells(nFil, 20).Value = mtope
      Else
         xlSheet.Cells(nFil, 20).Value = Rs!sueldo
      End If
      xlSheet.Cells(nFil, 21).Value = Trim(Rs!Condicion)
      xlSheet.Cells(nFil, 22).Value = ""
      xlSheet.Cells(nFil, 23).Value = "P"
      xlSheet.Cells(nFil, 24).Value = Trim(Rs!movi)
      xlSheet.Cells(nFil, 25).Value = Trim(Rs!FecIni)
      xlSheet.Cells(nFil, 26).Value = Trim(Rs!MonPrima)
      nFil = nFil + 1
   End If
   Rs.MoveNext
Loop

xlSheet.Range("A:Z").EntireColumn.AutoFit

Dim I As Integer
For I = 1 To 2
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridlines = False
Next

xlApp2.Sheets(1).Select
xlApp2.Application.Visible = True
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing



Call Procesa_Seguro_txt(mtope, "01")
Call Procesa_Seguro_txt(mtope, "02")

Screen.MousePointer = vbDefault
End Sub
Private Sub procesa_snpessalud()

Sql = "Usp_Pla_CCC '" & wcia & "'," & Val(Txtano.Text) & "," & Cmbmes.ListIndex + 1 & ",'**'"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst Else Rs.Close: Set Rs = Nothing: Exit Sub
CCC_Empleados ("01")
CCC_Empleados ("02")
End Sub
Private Sub CCC_Empleados(lTipoTrab As String)

Dim lIpssVida As Integer
Dim lSenati As Integer
Dim lTrabajadores As Integer
Dim lCuenta As Integer
Dim lBaseSenati As Double
Dim lDedSenati As Double
Dim lIpssVidaApor As Double

lIpssVida = 0: lSenati = 0: lTrabajadores = 0: lBaseSenati = 0
lCuenta = 1

If lTipoTrab = "02" Then
   Set xlSheet = xlApp2.Worksheets("HOJA2")
   xlSheet.Name = "OBREROS"
Else
   Set xlApp1 = CreateObject("Excel.Application")
   xlApp1.Workbooks.Add
   Set xlApp2 = xlApp1.Application
   xlApp2.Sheets.Add
   xlApp2.Sheets("Hoja1").Select
   xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)

   Set xlBook = xlApp2.Workbooks(1)
   Set xlSheet = xlApp2.Worksheets("HOJA1")
   xlSheet.Name = "EMPLEADOS"
End If

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("B:B").ColumnWidth = 8
xlSheet.Range("C:C").ColumnWidth = 40
xlSheet.Range("D:K").ColumnWidth = 12

xlSheet.Range("D:K").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(2, 2).Value = "APORTACIONES Y DEDUCCIONES  ( " & Cmbmes.Text & " - " & Txtano.Text & " )"
xlSheet.Range(xlSheet.Cells(2, 2), xlSheet.Cells(2, 11)).Merge
xlSheet.Range(xlSheet.Cells(2, 2), xlSheet.Cells(2, 11)).HorizontalAlignment = xlCenter

If lTipoTrab = "01" Then
   xlSheet.Cells(4, 2).Value = "PLANILLA SUELDOS"
Else
   xlSheet.Cells(4, 2).Value = "PLANILLA SALARIOS"
End If

nFil = 6
xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil + 1, 2)).Merge
xlSheet.Cells(nFil, 3).Value = "NOMBRE"
xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil + 1, 3)).Merge
xlSheet.Cells(nFil, 4).Value = "SNP"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 5)).Merge
xlSheet.Cells(nFil + 1, 4).Value = "BASE"
xlSheet.Cells(nFil + 1, 5).Value = "CALCULO"
xlSheet.Cells(nFil, 6).Value = "ESSALUD"
xlSheet.Range(xlSheet.Cells(nFil, 6), xlSheet.Cells(nFil, 7)).Merge
xlSheet.Cells(nFil + 1, 6).Value = "BASE"
xlSheet.Cells(nFil + 1, 7).Value = "CALCULO"
xlSheet.Cells(nFil, 8).Value = "SENATI"
xlSheet.Range(xlSheet.Cells(nFil, 8), xlSheet.Cells(nFil, 9)).Merge
xlSheet.Cells(nFil + 1, 8).Value = "BASE"
xlSheet.Cells(nFil + 1, 9).Value = "CALCULO"
xlSheet.Cells(nFil, 10).Value = "QUINTA"
xlSheet.Range(xlSheet.Cells(nFil, 10), xlSheet.Cells(nFil + 1, 10)).Merge
xlSheet.Cells(nFil, 11).Value = "ESS.VIDA"
xlSheet.Range(xlSheet.Cells(nFil, 11), xlSheet.Cells(nFil + 1, 11)).Merge
      
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil + 1, 11)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil + 1, 11)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil + 1, 11)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil + 1, 11)).Font.Bold = True

nFil = 8
Dim msum As Integer
Dim mTotTrab As Integer

mTotTrab = 0
Dim mCosto As String
mCosto = ""
msum = 1
Do While Not Rs.EOF
   If lTipoTrab = Rs!TipoTrab And Trim(Rs!PlaCod & "") <> "*****" And Trim(Rs!Planta & "") <> "**" And Trim(Rs!ccosto & "") <> "**" Then
      If mCosto <> Trim(Rs!DesCosto & "") Then
         If mCosto <> "" Then
            lCuenta = 1
            msum = msum * -1
            nFil = nFil + 1
            xlSheet.Cells(nFil, 2).Value = "TOTAL " & Trim(mCosto & "")
            xlSheet.Cells(nFil, 2).Font.Bold = True
               
            xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Cells(nFil, 9).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Cells(nFil, 10).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Cells(nFil, 11).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 11)).Borders.LineStyle = xlContinuous
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 11)).Font.Bold = True

            nFil = nFil + 1
            msum = 1
         End If
         
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(Rs!DesCosto & "")
         xlSheet.Cells(nFil, 2).Font.Bold = True
         nFil = nFil + 1
         mCosto = Trim(Rs!DesCosto & "")
      End If
      xlSheet.Cells(nFil, 1).Value = lCuenta
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = Rs!BSNP
      xlSheet.Cells(nFil, 5).Value = Rs!DSNP
      xlSheet.Cells(nFil, 6).Value = Rs!BESSALUD
      xlSheet.Cells(nFil, 7).Value = Rs!DESSALUD
      xlSheet.Cells(nFil, 8).Value = Rs!BSENATI
      xlSheet.Cells(nFil, 9).Value = Rs!DSENATI
      xlSheet.Cells(nFil, 10).Value = Rs!DQUINTA
      xlSheet.Cells(nFil, 11).Value = Rs!DVIDA
      xlSheet.Cells(nFil, 12).Value = Trim(Rs!desplanta & "")
      lTrabajadores = lTrabajadores + 1
      If Rs!DSENATI > 0 Then
         lSenati = lSenati + 1
         lBaseSenati = lBaseSenati + Rs!BSENATI
         lDedSenati = lDedSenati + Rs!DSENATI
      End If
      If Rs!DVIDA > 0 Then lIpssVida = lIpssVida + 1: lIpssVidaApor = lIpssVidaApor + Rs!DVIDA
      lCuenta = lCuenta + 1
      nFil = nFil + 1
      msum = msum + 1
   End If
   Rs.MoveNext
Loop

msum = msum * -1
nFil = nFil + 1
xlSheet.Cells(nFil, 2).Value = "TOTAL " & Trim(mCosto & "")
xlSheet.Cells(nFil, 2).Font.Bold = True
   
xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 9).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 10).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 11).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 11)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 11)).Font.Bold = True

nFil = nFil + 3

If Rs.RecordCount > 0 Then Rs.MoveFirst
msum = 1
Do While Not Rs.EOF
   If lTipoTrab = Rs!TipoTrab And Trim(Rs!PlaCod & "") = "*****" And Trim(Rs!Planta & "") <> "**" Then
      xlSheet.Cells(nFil, 2).Value = ""
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!desplanta & "")
      xlSheet.Cells(nFil, 4).Value = Rs!BSNP
      xlSheet.Cells(nFil, 5).Value = Rs!DSNP
      xlSheet.Cells(nFil, 6).Value = Rs!BESSALUD
      xlSheet.Cells(nFil, 7).Value = Rs!DESSALUD
      xlSheet.Cells(nFil, 8).Value = Rs!BSENATI
      xlSheet.Cells(nFil, 9).Value = Rs!DSENATI
      xlSheet.Cells(nFil, 10).Value = Rs!DQUINTA
      xlSheet.Cells(nFil, 11).Value = Rs!DVIDA
      xlSheet.Cells(nFil, 12).Value = Trim(Rs!DesCosto & "")
      nFil = nFil + 1
      msum = msum + 1
   End If
   Rs.MoveNext
Loop

nFil = nFil + 1

xlSheet.Cells(nFil, 3).Value = "TOTAL GENERAL"
xlSheet.Cells(nFil, 3).Font.Bold = True
msum = msum * -1
xlSheet.Cells(nFil, 4).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 5).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 6).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 7).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 8).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 9).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 10).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 11).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Cells(nFil, 12).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 11)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 11)).Font.Bold = True

nFil = nFil + 3

xlSheet.Cells(nFil, 2).Value = "NUMERO DE TRABAJADORES"
xlSheet.Cells(nFil, 4).Value = lTrabajadores
xlSheet.Cells(nFil, 4).NumberFormat = "#,###,##0;[Red](#,###,##0)"


nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "TRABAJADORES CON INGRESOS MENORES AL SUELDO MINIMO"
nFil = nFil + 2
If Rs.RecordCount > 0 Then Rs.MoveFirst
lCuenta = 1
Do While Not Rs.EOF
   If lTipoTrab = Rs!TipoTrab And Trim(Rs!PlaCod & "") <> "*****" And Trim(Rs!Planta & "") = "**" And Trim(Rs!ccosto & "") = "**" Then
      xlSheet.Cells(nFil, 1).Value = lCuenta
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
      xlSheet.Cells(nFil, 6).Value = Rs!BESSALUD
      xlSheet.Cells(nFil, 7).Value = Rs!DESSALUD
      If Rs!BSNP <> 0 Then xlSheet.Cells(nFil, 8).Value = "SUBSIDIO"
      If Rs!BSNP <> 0 Then xlSheet.Cells(nFil, 9).Value = Rs!BSNP
      If Trim(Rs!desplanta & "") <> "" Then xlSheet.Cells(nFil, 10).Value = "F. CESE"
      xlSheet.Cells(nFil, 11).Value = Trim(Rs!desplanta & "")
      xlSheet.Cells(nFil, 11).HorizontalAlignment = xlRight
      xlSheet.Cells(nFil, 11).NumberFormat = "dd/mm/yyyy;@"
      nFil = nFil + 1
   End If
   Rs.MoveNext
Loop

nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "TRABAJADORES QUE NO APORTARON ESSALUD VIDA"
If Rs.RecordCount > 0 Then Rs.MoveFirst
lCuenta = 1
Do While Not Rs.EOF
   If lTipoTrab = Rs!TipoTrab And Trim(Rs!PlaCod & "") <> "*****" And Trim(Rs!Planta & "") = "**" And Trim(Rs!ccosto & "") <> "**" Then
      xlSheet.Cells(nFil, 1).Value = lCuenta
      xlSheet.Cells(nFil, 2).Value = Trim(Rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(Rs!nombre & "")
      nFil = nFil + 1
      lCuenta = lCuenta + 1
   End If
   Rs.MoveNext
Loop

nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "NUMERO DE TRABAJADORES CON APORTACION A SENATI"
xlSheet.Cells(nFil, 4).Value = lSenati
xlSheet.Cells(nFil, 4).NumberFormat = "#,###,##0;[Red](#,###,##0)"
xlSheet.Cells(nFil, 5).Value = "BASE"
xlSheet.Cells(nFil, 5).HorizontalAlignment = xlRight
xlSheet.Cells(nFil, 6).Value = lBaseSenati
xlSheet.Cells(nFil, 7).Value = "APORTE"
xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
xlSheet.Cells(nFil, 8).Value = lDedSenati

nFil = nFil + 2
xlSheet.Cells(nFil, 2).Value = "NUMERO DE TRABAJADORES CON APORTACION A ESSALUD VIDA"
xlSheet.Cells(nFil, 4).Value = lIpssVida
xlSheet.Cells(nFil, 4).NumberFormat = "#,###,##0;[Red](#,###,##0)"
xlSheet.Cells(nFil, 7).Value = "APORTE"
xlSheet.Cells(nFil, 7).HorizontalAlignment = xlRight
xlSheet.Cells(nFil, 8).Value = lIpssVidaApor


If lTipoTrab = "02" Then
   For I = 1 To 3
      xlApp2.Sheets(I).Select
      xlApp2.Sheets(I).Range("A1:A1").Select
      xlApp2.Application.ActiveWindow.DisplayGridlines = False
      'xlApp2.ActiveWindow.Zoom = 80
   Next

   xlApp2.Sheets(1).Select
   xlApp2.Application.Visible = True
End If

End Sub


