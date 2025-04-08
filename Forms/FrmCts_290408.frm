VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmCts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depositos de CTS"
   ClientHeight    =   7305
   ClientLeft      =   4395
   ClientTop       =   2400
   ClientWidth     =   7455
   Icon            =   "FrmCts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdProvision 
      Caption         =   "Generar Asientos Provision"
      Height          =   315
      Left            =   75
      TabIndex        =   14
      Top             =   6915
      Width           =   2160
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcular"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5985
      TabIndex        =   12
      Top             =   6915
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Certificados"
      Height          =   315
      Left            =   3435
      TabIndex        =   11
      Top             =   6915
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reporte por Bancos"
      Height          =   315
      Left            =   4410
      TabIndex        =   10
      Top             =   6915
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reporte CTS"
      Height          =   315
      Left            =   2235
      TabIndex        =   9
      Top             =   6915
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   7455
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         ItemData        =   "FrmCts.frx":030A
         Left            =   5400
         List            =   "FrmCts.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   120
         Width           =   1935
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "FrmCts.frx":0334
         Left            =   840
         List            =   "FrmCts.frx":035C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   2055
      End
      Begin VB.TextBox Txtano 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
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
         Left            =   3960
         TabIndex        =   18
         Top             =   120
         Width           =   1365
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   100
         Width           =   255
         Size            =   "450;661"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
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
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5775
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   7455
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "FrmCts.frx":03C4
         Height          =   5295
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   9340
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "nombre"
            Caption         =   "Nombre"
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
            DataField       =   "monto_total"
            Caption         =   "Rem. Afecta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "provision_actual"
            Caption         =   "Deposito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "placod"
            Caption         =   "codigo"
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
         BeginProperty Column04 
            DataField       =   "fingreso"
            Caption         =   "fingreso"
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
               ColumnWidth     =   4649.953
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   6240
         TabIndex        =   16
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deposito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   4800
         TabIndex        =   15
         Top             =   5520
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   6135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   165
         Width           =   840
      End
   End
   Begin MSAdodcLib.Adodc AdoCabeza 
      Height          =   1410
      Left            =   2520
      Top             =   4080
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   2487
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mperiodo As String
Dim Sql As String
Dim mpag As Integer
Dim RUTA As String
Dim mlinea As Integer
'***************codigo nuevo giovanni 17082007*************************
Dim s_MesSeleccion As String
Dim rs_Liquidacion As ADODB.Recordset
Dim rs_Liquidacion2 As ADODB.Recordset
Dim s_CodEmpresa_Starsoft As String
Dim i_Numero_Voucher As Integer
Dim i_Numero_VoucherG As String
Dim s_TipoTrabajador_Cts As String
Dim s_Dia_Proceso_Prov As String
'**********************************************************************

Const COL_PRIMERMES = 5
Const COL_SEGUNDOMES = 11
Dim ArrReporte() As Variant

Const MAXCOL = 19
Const CLM_CODIGO = 1
Const CLM_TRABAJADOR = 2
Const CLM_FECINGRESO = 3
Const CLM_FECCESE = 4
Const CLM_TIEMPOSERV = 5
Const CLM_JORNAL = 6
Const CLM_AFP3 = 7
Const CLM_BONOCOSTO = 8
Const CLM_BONOTSERV = 9
Const CLM_ASIGFAM = 10
Const CLM_PROMGRATI = 11
Const CLM_PROMHEXTRAS = 12
Const CLM_PROMOTROSPAG = 13
Const CLM_PROMHVERANO = 14
Const CLM_PROMTURNO = 15
Const CLM_PROMPRODUCC = 16
Const CLM_PROMBONOPROD = 17
Const CLM_PROMREMDI = 18
Const CLM_JORNALINDEN = 19
Const CLM_TIEMPOSERVACT = 3
Const CLM_MONTOINDEN1 = 4
Const CLM_MONTOINDEN2 = 5
Const CLM_MONTOINDENSINTP = 6
Const CLM_MONTOTOTINDEN = 7
Const CLM_PROVANOANT = 8
Const CLM_AJUSTEPROVI = 9
Const CLM_PROVIANOACTUAL = 10
Const CLM_PROVMESACTUAL = 11
Const CLM_SALDOPEND = 12

Private Sub Cmbmes_Click()
Carga_Cts
End Sub
Private Sub CmbTipo_Click()
    Call Carga_Cts
End Sub
Private Sub CmdProvision_Click()
    Call Generar_Asientos_Contables_Provision
End Sub
Private Sub Command1_Click()
'Reporte_Cts
ReporteCts
End Sub

Private Sub Command2_Click()
Reporte_Cts_Bancos
End Sub

Private Sub Command3_Click()
Procesa_Certifica_Cts
End Sub

Private Sub Command4_Click()
'**********codigo agregado giovanni 10092007**************************
If Cmbtipo.Text = "" Then
    MsgBox "Debe Seleccionar Tipo Trabajador", vbInformation
Else
    PROVICIONES_CTS
End If
'*********************************************************************
    'PROVICIONES_CTS
End Sub
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
'Me.Height = 8065
'Me.Width = 7575
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
If Month(Date) = 1 Then
   Txtano.Text = Format(Year(Date) - 1, "0000")
   Cmbmes.ListIndex = 11
Else
  Txtano.Text = Format(Year(Date), "0000")
  Cmbmes.ListIndex = Month(Date) - 2
End If
   
End Sub

Private Sub Calculo_CTS()
Dim mcad As String
Dim mcadIns As String
Dim mCadVal As String
Dim mCadFields As String
Dim mFactor As Currency
Dim mTotal As Currency
Dim mDepo As Currency
Dim rs2 As ADODB.Recordset
Dim rsAfectos As ADODB.Recordset
Dim i As Integer, X As Integer
Dim INSCAD As String, VALUESCAD As String
Dim NRO As Integer, cad As String
Dim VALORES(50) As Integer

Screen.MousePointer = vbArrowHourglass

mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
Sql = "select factorcts from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql)) Then
   If IsNull(rs!factorcts) Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   If rs!factorcts = 0 Then
      MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
Else
   MsgBox "No se ha registrado el Factor para calculo de CTS", vbInformation, "Calculo de CTS"
   Screen.MousePointer = vbDefault
   Exit Sub
End If

mFactor = rs!factorcts
rs.Close

Sql = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='C' and status<>'*' order by cod_remu"
If (fAbrRst(rsAfectos, Sql)) Then
   rsAfectos.MoveFirst
Else
  MsgBox "No Se Registran Remuneraciones Afectas para el Calculo", vbCritical, "Calculo de CTS"
  Screen.MousePointer = vbDefault
  Exit Sub
End If

mcad = ""
mCadFields = ""

NRO = 0
Do While Not rsAfectos.EOF
   NRO = NRO + 1
   mcad = mcad & "SUM(I" & Format(rsAfectos!cod_remu, "00") & "),"
   mCadFields = mCadFields & "I" & Format(rsAfectos!cod_remu, "00") & ","
   rsAfectos.MoveNext
Loop

mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
Sql = "select " & mcad
Sql = Sql & ",placod from plahistorico where cia='" & _
wcia & "' and year(fechaproceso)=" & Trim(Txtano.Text) & _
" and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & _
" and status<>'*' Group by placod"

  If (fAbrRst(rs, Sql)) Then
     rs.MoveFirst
  Else
     MsgBox "No Se Registran Boletas para el Calculo", vbCritical, "Calculo de CTS"
     Screen.MousePointer = vbDefault
     Exit Sub
  End If

Sql = wInicioTrans
cn.Execute Sql

Sql = "update platserdep set status='*' where cia='" & wcia & "' and periodo='" & mperiodo & "' and status<>'*'"
cn.Execute Sql

Do While Not rs.EOF
   Sql = "select fingreso,cargo,ctsbanco,ctstipcta,ctsmoneda,ctsnumcta from planillas where cia='" & wcia & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
   If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
   rsAfectos.MoveFirst
   mcadIns = ""
   i = 0
   mTotal = 0: mDepo = 0
   Do While Not rsAfectos.EOF
      mCadVal = "0"
      mCadVal = rs(i)
      
      If Not IsNull(rs(i)) Then mTotal = mTotal + rs(i)
      VALORES(Val(rsAfectos(0))) = rs(i)
      mcadIns = mcadIns & mCadVal & ","
      rsAfectos.MoveNext
      i = i + 1
   Loop
   mDepo = Round((mTotal * mFactor) / 100, 2)
  
   INSCAD = "i05,i07,i08," & _
   "i16,i17,i18,i19,i20," & _
   "i21,i22,i23,i24,i25,i26,i27,i28,i29,i30,i31," & _
   "i32,i33,i34,i35,i36,i37,i38,i39,i40,i41,i42," & _
   "i43,i44,i45,i46,i47,i48,i49,i50"
   
   INSCAD = ""
   VALUESCAD = ""
   For X = 1 To 50
       INSCAD = INSCAD & "I" & Format(X, "00") & ","
       VALUESCAD = VALUESCAD & VALORES(X) & ","
   Next
   
   'VALUESCAD = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
  
   mcadIns = mcadIns & Str(mTotal) & ","
   Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & "  "
   Sql = Sql & "insert into platserdep(cia,placod,fechaingreso,cargo," & INSCAD
   Sql = Sql & "total,factor,fecha,periodo,banco,cta,moneda,nro_cta,status,interes,TOTLIQUID) "
   Sql = Sql & "values('" & wcia & "','" & Trim(rs!PLACOD) & "','" & Format(rs2!fingreso, FormatFecha) & "','" & Trim(rs2!cargo) & "',"
   Sql = Sql & VALUESCAD & mTotal & ","
   Sql = Sql & "" & mFactor & "," & FechaSys & ",'" & mperiodo & "','" & rs2!ctsbanco & "','" & rs2!ctstipcta & "',"
   Sql = Sql & "'" & rs2!ctsmoneda & "','" & _
   Trim(rs2!ctsnumcta) & "','',0,0" & ")"
   

   
   cn.Execute Sql
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
Loop

Sql = wFinTrans
cn.Execute Sql

If rs.State = 1 Then rs.Close
If rsAfectos.State = 1 Then rsAfectos.Close
Screen.MousePointer = vbDefault
Carga_Cts
End Sub

Private Sub SpinButton1_SpinDown()
If Trim(Txtano.Text) = "" Then
   Txtano.Text = Format(Year(Date), "0000")
Else
   Txtano.Text = Txtano.Text - 1
End If
End Sub

Private Sub SpinButton1_SpinUp()
If Trim(Txtano.Text) = "" Then
   Txtano.Text = Format(Year(Date), "0000")
Else
   Txtano.Text = Txtano.Text + 1
End If
End Sub

Private Sub Txtano_Change()
    Carga_Cts
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Carga_Cts()

If Cmbtipo.Text = "" Then
    Exit Sub
End If

Call Captura_Tipo_Trabajador

If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
mperiodo = Trim(Txtano.Text) & Format(Cmbmes.ListIndex + 1, "00")
Sql = nombre()

Select Case s_TipoTrabajador_Cts
    Case "01"
        Sql = Sql & "a.placod,a.monto_total,a.provision_actual,b.fingreso " _
        & "from plaprovcts a,planillas b " _
        & "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
        & "and a.placod=b.placod and a.cia=b.cia and b.tipotrabajador='01' and b.status<>'*' order by nombre"
    Case "02"
        Sql = Sql & "a.placod,a.monto_total,a.provision_actual,b.fingreso " _
        & "from plaprovcts a,planillas b " _
        & "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
        & "and a.placod=b.placod and a.cia=b.cia and b.tipotrabajador='02' and b.status<>'*' order by nombre"
    Case "03"
        Sql = Sql & "a.placod,a.monto_total,a.provision_actual,b.fingreso " _
        & "from plaprovcts a,planillas b " _
        & "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
        & "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by nombre"
End Select

'sql = sql & "a.placod,a.monto_total,a.provision_actual,b.fingreso " _
'& "from plaprovcts a,planillas b " _
'& "where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and a.status<>'*' " _
'& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by nombre"

cn.CursorLocation = adUseClient
Set AdoCabeza.Recordset = cn.Execute(Sql$, 64)
If AdoCabeza.Recordset.RecordCount > 0 Then
   AdoCabeza.Recordset.MoveFirst
   Command4.Enabled = False
   Sql = "select SUM(provision_actual) from plaprovcts where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and status<>'*'"
   If (fAbrRst(rs, Sql)) Then Lbltotal.Caption = Format(rs(0), "###,###.00")
   rs.Close
Else
   Command4.Enabled = True
   Lbltotal.Caption = "0.00"
End If
Dgrdcabeza.Refresh
End Sub
Public Sub Elimina_Cts()

mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")

If MsgBox("Desea Eliminar Calculo de Cts ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
Else
    Sql = wInicioTrans
    cn.Execute Sql
    
    Sql = "update plaprovcts set status='*' where cia='" & wcia & "' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and year(fechaproceso)=" & Txtano.Text & " and status<>'*'"
    cn.Execute Sql
    
    Sql = wFinTrans
    cn.Execute Sql
    
    Carga_Cts
End If

Screen.MousePointer = vbDefault

End Sub
Private Sub Reporte_Cts()
Dim mcad As String
Dim mtotd As Currency
Dim mtott As Currency
mpag = 0
If AdoCabeza.Recordset.RecordCount <= 0 Then
   MsgBox "No Hay Depositos", vbInformation, "Depositos de CTS": Exit Sub
   Exit Sub
End If
RUTA = App.path & "\REPORTS\" & "RepCts.txt"
Open RUTA For Output As #1
Cabeza_Lista_CTS
mtotd = 0: mtott = 0
AdoCabeza.Recordset.MoveFirst
Do While Not AdoCabeza.Recordset.EOF
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Cabeza_Lista_CTS
   mcad = lentexto(55, Left(AdoCabeza.Recordset!nombre, 55))
   mcad = AdoCabeza.Recordset!PLACOD & "   " & mcad & Space(5) & fCadNum(AdoCabeza.Recordset!Total, "##,###,##0.00") & Space(5) & fCadNum(AdoCabeza.Recordset!totliquid, "##,###,##0.00")
   Print #1, Space(2) & mcad
   mtotd = mtotd + AdoCabeza.Recordset!totliquid
   mtott = mtott + AdoCabeza.Recordset!Total
   AdoCabeza.Recordset.MoveNext
Loop
Print #1,
Print #1, Space(35) & "TOTAL :                           " & fCadNum(mtott, "###,###,##0.00") & Space(4) & fCadNum(mtotd, "###,###,##0.00")
Close #1
Call Imprime_Txt("RepCts.txt", RUTA)
End Sub
Private Sub Cabeza_Lista_CTS()
mpag = mpag + 1
Print #1, Chr(18) & Space(2) & Trim(Cmbcia.Text) & Space(25) & "Pag. " & fCadNum(mpag, "###")
Print #1,
Print #1, Space(22) & "DEPOSITO POR TIEMPO DE SERVICIO"
Print #1, Space(23) & "PERIODO : " & Cmbmes.Text & " - " & Format(Txtano.Text, "0000") & Chr(15)
Print #1, Space(2) & String(107, "-")
Print #1, Space(2) & "CODIGO            NOMBRE                                               REMUNERACION          MONTO A "
Print #1, Space(2) & "                                                                        COMPUTABLE          DEPOSITAR"
Print #1, Space(2) & String(107, "-")
mlinea = 10
End Sub
Private Sub Reporte_Cts_Bancos()
Dim mcad As String
Dim mcod As String
Dim mItem As Integer
Dim mTotDep As Currency
Dim rs2 As ADODB.Recordset
Dim wciamae As String
Dim mtotsbco As Currency
Dim mtotnbco As Integer
Dim mtotsbcoTS As Currency
Dim mtotnbcoTS As Integer
Dim mtotsbcoTD As Currency
Dim mtotnbcoTD As Integer

If Trim(Txtano.Text) = "" Then Exit Sub
If Cmbmes.ListIndex < 0 Then Exit Sub
mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
Sql$ = nombre


Sql = "select ap_pat,ap_mat,rtrim(nom_1)+' '+rtrim(nom_2) AS nombre,"
Sql = Sql & "a.placod,a.totalremun AS total,a.monto_total AS totliquid,B.ctsbanco AS banco,b.ctsnumcta as nro_cta,b.ctsmoneda as moneda,b.fingreso " _
& "from plaprovcts a,planillas b " _
& "where a.cia='" & wcia & "' and MONTH(a.FECHAPROCESO)='" & Format(Cmbmes.ListIndex + 1, "00") & "' and a.status<>'*' " _
& "and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by banco,B.moneda,nombre"

If (fAbrRst(rs, Sql)) Then
   rs.MoveFirst
Else
   MsgBox "No Hay Depositos", vbInformation, "Depositos de CTS": Exit Sub
   Exit Sub
End If

RUTA = App.path & "\REPORTS\" & "BcoCts.txt"
Open RUTA For Output As #1
Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
mcod = rs!banco & rs!moneda
mItem = 0: mTotDep = 0
Do While Not rs.EOF
   If mcod <> rs!banco & rs!moneda Then
      Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
      Print #1, Space(37) & "TOTAL INSCRITOS < " & fCadNum(mItem, "#####") & " >                            TOTAL ACUM.       < " & fCadNum(mTotDep, "##,###,##0.00") & " >"
      mItem = 0: mTotDep = 0
     Print #1, Chr(12) + Chr(13): Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
     mcod = rs!banco & rs!moneda
   End If
   If mlinea > 55 Then Print #1, Chr(12) + Chr(13): Call Cabeza_Banco_CTS(rs!banco, rs!moneda)
   mcad = lentexto(20, Left(rs!nro_cta, 20)) & Space(5)
   mcad = mcad & lentexto(20, Left(rs!ap_pat, 20)) & Space(2) & lentexto(20, Left(rs!ap_mat, 20)) & Space(2) & Space(2) & lentexto(25, Left(rs!nombre, 25))
   mcad = mcad & Space(2) & fCadNum(rs!totliquid, "##,###,##0.00")
   Print #1, Space(12) & mcad
   mlinea = mlinea + 1
   mItem = mItem + 1
   mTotDep = mTotDep + rs!totliquid
   rs.MoveNext
Loop
Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
Print #1, Space(37) & "TOTAL INSCRITOS < " & fCadNum(mItem, "#####") & " >                            TOTAL ACUM.       < " & fCadNum(mTotDep, "##,###,##0.00") & " >"
If rs.State = 1 Then rs.Close


Sql = "select distinct(banco) from platserdep where cia='" & wcia & "' and periodo='" & mperiodo & "' and  status<>'*'"

Sql = "select distinct(ctsbanco) as banco from plAprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and  a.status<>'*'"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Print #1, Chr(12) + Chr(13): Cabeza_Banco_Resumen
mtotnbcoTS = 0: mtotsbcoTS = 0
mtotnbcoTD = 0: mtotsbcoTD = 0
Do While Not rs.EOF
   mcad = ""
   wciamae = Determina_Maestro("01007")
   mtotnbco = 0: mtotsbco = 0
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs!banco & "'" & Space(5)
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(33, Left(rs2!descrip, 23)) Else mcad = Space(33)
   If rs2.State = 1 Then rs2.Close
   
   'Dolares
   Sql = "select count(a.placod) as num from plAprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
   Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda<>'" & wmoncont & "' and a.status<>'*'"
   
   If (fAbrRst(rs2, Sql)) Then
      If rs2!Num = 0 Then
         mcad = mcad & "     0"
      Else
         mcad = mcad & " " & fCadNum(rs2!Num, "#####")
      End If
      mtotnbco = mtotnbco + rs2!Num
      mtotnbcoTD = mtotnbcoTD + rs2!Num
      rs2.Close
    
      Sql = "select sum(a.monto_total) as depo from plAprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
      Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda<>'" & wmoncont & "' and a.status<>'*'"
      
      
      If (fAbrRst(rs2, Sql)) Then
         If IsNull(rs2!depo) Then
            mcad = mcad & " " & fCadNum(0, "#,###,##0.00")
         Else
            mcad = mcad & " " & fCadNum(rs2!depo, "#,###,##0.00")
            mtotsbco = mtotsbco + rs2!depo
            mtotsbcoTD = mtotsbcoTD + rs2!depo
         End If
      End If
      rs2.Close
      mcad = mcad & Space(2)
   End If
   
   'Soles
    Sql = "select count(a.placod) as num from plAprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
    Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda='" & wmoncont & "' and a.status<>'*'"
   
   If (fAbrRst(rs2, Sql)) Then
      mcad = mcad & Space(8)
      If rs2!Num = 0 Then
         mcad = mcad & "     0"
      Else
         mcad = mcad & " " & fCadNum(rs2!Num, "#####")
      End If
      mtotnbco = mtotnbco + rs2!Num
      mtotnbcoTS = mtotnbcoTS + rs2!Num
      rs2.Close
            
      Sql = "select sum(monto_total) as depo from plAprovcts a inner join planillas b ON (b.cia=a.cia and b.placod=a.placod and b.status!='*' )"
        Sql = Sql & " where a.cia='" & wcia & "' and month(a.fechaproceso)=" & Cmbmes.ListIndex + 1 & " and  b.ctsbanco='" & rs!banco & "' and b.ctsmoneda='" & wmoncont & "' and a.status<>'*'"
      
      If (fAbrRst(rs2, Sql)) Then
         If IsNull(rs2!depo) Then
            mcad = mcad & " " & fCadNum(0, "#,###,##0.00")
         Else
            mcad = mcad & " " & fCadNum(rs2!depo, "#,###,##0.00")
            mtotsbco = mtotsbco + rs2!depo
            mtotsbcoTS = mtotsbcoTS + rs2!depo
         End If
      End If
      rs2.Close
      'Total
      mcad = mcad & Space(10) & " " & fCadNum(mtotnbco, "#####") & " " & fCadNum(mtotsbco, "#,###,##0.00")
      mcad = mcad & Space(2)
   End If
   Print #1, Space(12) & mcad
   rs.MoveNext
Loop
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
'Total General
mcad = "  ** T O T A L ** " & Space(16) & fCadNum(mtotnbcoTD, "#####") & " " & fCadNum(mtotsbcoTD, "#,###,##0.00")
mcad = mcad & Space(11) & fCadNum(mtotnbcoTS, "#####") & " " & fCadNum(mtotsbcoTS, "#,###,##0.00")
mcad = mcad & Space(11) & fCadNum(mtotnbcoTS + mtotnbcoTD, "#####") & " " & fCadNum(mtotsbcoTS + mtotsbcoTD, "#,###,##0.00")
Print #1, Space(12) & mcad
Close #1
Call Imprime_Txt("BcoCts.txt", RUTA)
End Sub
Private Sub Cabeza_Banco_CTS(banco, moneda)
Dim wciamae As String
Dim rs2 As ADODB.Recordset
Dim mBanc As String
Dim MMON As String
If Trim(banco) = "" Then
   mBanc = "SIN BANCO"
Else
   mBanc = ""
   wciamae = Determina_Maestro("01007")
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & banco & "'" & Space(5)
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rs2, Sql$)) Then mBanc = lentexto(30, Left(rs2!descrip, 30))
   If rs2.State = 1 Then rs2.Close
End If
If Trim(moneda) = "" Then
   MMON = ""
ElseIf moneda = wmoncont Then
   MMON = "Moneda Nacional (" & moneda & ")"
Else
   MMON = "Moneda Extranjera (" & moneda & ")"
End If
Print #1, Chr(18) & Space(10) & Trim(Cmbcia.Text)
Print #1,
Print #1, Space(27) & mBanc
Print #1,
Print #1, Space(27) & MMON & Chr(15)
Print #1,
Print #1, Space(16) & "CUENTA                              APELLIDOS Y NOMBRES                                              MONTO"
Print #1, Space(12) & String(20, "_") & Space(5) & String(71, "_") & Space(2) & String(13, "_")
Print #1,
mlinea = 9
End Sub
Private Sub Cabeza_Banco_Resumen()
Print #1, Chr(18) & Space(10) & Trim(Cmbcia.Text)
Print #1,
Print #1, Space(27) & "R E S U M E N" & Chr(15)
Print #1,
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
Print #1,
Print #1, Space(46) & "DEP.MON.EXT (DOLAR)              DEP. SOLES                   T O T A L"
Print #1, Space(16) & "BANCOS                        -------------------          -------------------          -------------------"
Print #1, Space(46) & "# TRAB    MONTO S/.          # TRAB    MONTO S/.          # TRAB    MONTO S/."
Print #1, Space(12) & String(30, "_") & Space(5) & String(19, "_") & Space(10) & String(19, "_") & Space(10) & String(19, "_")
Print #1,
mlinea = 9
End Sub
Private Sub Procesa_Certifica_Cts()
Dim mItem As Integer
Dim mcad As String
Dim mnombre As String
Dim cadnombre As String
Dim mcadrem As String
Dim mtot As Currency
Dim mFactor As Currency
Dim mdir As String
Dim mmoneda As String
Dim cadefectivo As String
Dim mdpto As String
Dim rs2 As ADODB.Recordset
Dim rsRem As ADODB.Recordset
Dim wciamae As String
Dim mempleado As String
Dim mobrero As String
Dim mfec As String
Dim i As Integer

mfec = fMaxDay(Cmbmes.ListIndex + 1, Val(Txtano)) & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano
cadnombre = nombre()
mperiodo = Txtano.Text & Format(Cmbmes.ListIndex + 1, "00")
mdir = ""
mdpto = ""
mFactor = 6

Sql$ = "select distinct(direcc),nro,u.dpto  from cia c,ubigeos u where c.cod_cia='" & wcia & "' and c.status<>'*' and left(u.cod_ubi,5)=left(c.cod_ubi,5)"

Sql$ = "SELECT A.*,v.flag1 AS via,z.dpto as zona FROM CIA A LEFT OUTER JOIN ubigeos z ON (Left(z.cod_ubi,5)=left(a.cod_ubi,5))"
Sql$ = Sql$ & "    LEFT OUTER JOIN maestros_2 v ON ( v.ciamaestro='01036'  and v.cod_maestro2 =  a.cod_via) WHERE a.cod_cia='" & wcia & "'"

If (fAbrRst(rs, Sql$)) Then mdir = Left(Trim(rs!via) & " " & rs!direcc, 56) & " No " & Trim(rs!NRO) & " - " & Trim(rs!DPTO): mdpto = Trim(rs!zona)

rs.Close

Sql = "SELECT p.ctsbanco,p.ctsmoneda,p.ctstipcta,p.ctsnumcta,ppc.* FROM PLAPROVCTS ppc INNER JOIN planillas p ON (p.cia=ppc.cia and p.placod=ppc.placod and p.status!='*') "
Sql = Sql & "where ppc.cia='" & wcia & "' and MONTH(ppc.fechaproceso)='" & Format(Cmbmes.ListIndex + 1, "00") & "' and ppc.status<>'*' order by PPC.placod"

If Not (fAbrRst(rs, Sql$)) Then Else rs.MoveFirst

RUTA = App.path & "\REPORTS\" & "CertCts.txt"
Open RUTA For Output As #1

Do While Not rs.EOF
   wciamae = Determina_Maestro("01007")
   Sql$ = "Select descrip from maestros_2 where status<>'*' and cod_maestro2='" & rs!ctsbanco & "'" & Space(5)
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rs2, Sql$)) Then mcad = lentexto(48, Left(rs2!descrip, 48)) Else mcad = Space(48)
   If rs2.State = 1 Then rs2.Close

   If rs!ctsmoneda = wmoncont Then
      mmoneda = "     CUENTA A PLAZO/MONEDA NACIONAL (SOLES)       "
   Else
      mmoneda = "     CUENTA A PLAZO/EXTRANJERA NACIONAL (DOLARES) "
   End If
   Sql$ = cadnombre & "placod,tipotrabajador,fingreso,fcese from planillas where cia='" & wcia & "' and placod='" & rs!PLACOD & "' and status<>'*'"
   If (fAbrRst(rs2, Sql$)) Then mnombre = lentexto(41, Left(rs2!nombre, 41)) Else mnombre = Space(40)
   If rs2!TipoTrabajador = "01" Then
      mempleado = " EMPLEADO [X] "
      mobrero = " OBRERO   [ ] "
   Else
      mempleado = " EMPLEADO [ ] "
      mobrero = " OBRERO   [X] "
   End If
   
    Print #1, Chr(18) & "Liquidacion de Compensacion por Tiempo de Servicios(CTS)" & Chr(15)
    
    Print #1, "Ley de Compensacion por Tiempo de Servicio TUO del D.Leg 650 (DS-1-97-TR de 27-02-97 Y DS 4-97-TR)"
    
    Print #1, Chr(218) & String(87, Chr(196)) & Chr(194) & String(44, Chr(196)) & Chr(191)
    
    Print #1, Chr(179) & "Nombre o Razon Social del Empleador :  " & lentexto(48, Left(Cmbcia.Text, 48)) & Chr(179) & "Ciudad y Fecha  " & lentexto(16, Left(mdpto, 16)) & Format(Date, "dd/mm/yyyy") & "  " & Chr(179)
    
    Print #1, Chr(195) & String(77, Chr(196)) & Chr(194) & String(9, Chr(196)) & Chr(193) & String(16, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & "Direccion del Empleador :   " & lentexto(49, Left(mdir, 49)) & Chr(179) & "  Deposito  [X]           " & Chr(179) & "     Pago Directo  [ ]     " & Chr(179)
    
    Print #1, Chr(195) & String(53, Chr(196)) & Chr(194) & String(23, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & Space(15) & "ENTIDAD DEPOSITARIA" & Space(19) & Chr(179) & "Tipo de Cuenta  " & Space(34) & Chr(179) & "No de Cuenta" & Space(15) & Chr(179)
    
    Print #1, Chr(179) & Space(5) & mcad & Chr(179) & mmoneda & Chr(179) & Space(1) & lentexto(26, Left(rs!ctsnumcta, 26)) & Chr(179)
    
    Print #1, Chr(195) & String(53, Chr(196)) & Chr(193) & String(23, Chr(196)) & Chr(194) & String(26, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & Space(28) & "DATOS DEL TRABAJADOR" & Space(29) & Chr(179) & "    CONCEPTOS REMUNERATIVOS PARA DETERMINAR LA        " & Chr(179)
    
    Print #1, Chr(179) & Space(77) & Chr(179) & "               REMUNERACION COMPUTABLE                " & Chr(179)
    
    Print #1, Chr(195) & String(17, Chr(196)) & Chr(194) & String(59, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & "CODIGO: " & lentexto(9, Left(rs!PLACOD, 9)) & Chr(179) & "Apell.y Nombres : " & mnombre & Chr(179) & "         CONCEPTO         " & Chr(179) & "            MONTO          " & Chr(179)
    
    Print #1, Chr(195) & String(17, Chr(196)) & Chr(197) & String(14, Chr(196)) & Chr(194) & String(17, Chr(196)) & Chr(194) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    
    Print #1, Chr(179) & "Fecha de Ingreso " & Chr(179) & mempleado & Chr(179) & "  Fecha de Cese  " & Chr(179) & "      Motivo de Cese      " & Chr(179) & "                          " & Chr(179) & "                           " & Chr(179)
    
    mcad = ""
    
    mItem = 0
    For i = 9 To 108
        If rs(i) <> 0 Then
           mcadrem = Chr(179)
           Sql = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' and codinterno='" & Right(rs(i).Name, 2) & "'"
           If (fAbrRst(rsRem, Sql$)) Then
                Select Case Trim(rsRem!Descripcion)
                Case Is = "BASICO": mcadrem = mcadrem & lentexto(26, Left("SUELDO BASICO", 26))
                Case Is = "GRATIFICACION": mcadrem = mcadrem & lentexto(26, Left("PROM.GRATIFICACION", 26))
                Case Else
                    mcadrem = mcadrem & lentexto(26, Left(rsRem!Descripcion, 26))
                End Select
           Else
                mcadrem = mcadrem & Space(26)
           End If
           mcadrem = mcadrem & Chr(179) & Space(12) & fCadNum(rs(i), "##,###,##0.00") & Space(2) & Chr(179)
           rsRem.Close
           mItem = mItem + 1
           Select Case mItem
                  Case Is = 1
                       mcad = Chr(179) & "   " & Format(rs2!fingreso, "dd/mm/yyyy") & "    " & Chr(179) & mobrero & Chr(179) & Space(3)
                       If IsNull(rs2!fcese) Then mcad = mcad & Space(10) Else mcad = mcad & Format(rs2!fcese, "dd/mm/yyyy")
                       mcad = mcad & Space(4) & Chr(179) & Space(26) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(194) & String(13, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                       Print #1, Chr(179) & "   " & Chr(179) & Space(19) & "PERIODO DE SERVICIOS QUE SE CANCELA" & Space(19) & Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
                  Case Is = 2
                       If Month(rs2!fingreso) = Cmbmes.ListIndex + 1 And Year(rs2!fingreso) = Val(Txtano.Text) Then
                          mcad = "Del : " & Format(rs2!fingreso, "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & " No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                          cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
                       Else
                          mcad = "Del : " & Format(DateAdd("m", -6, DateAdd("d", 1, mfec)), "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & "  No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                          cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
                       End If
                       mcad = Chr(179) & " 1 " & Chr(179) & mcad
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(197) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 3
                       mcad = Chr(179) & " 2 " & Chr(179) & "Periodo no Comp.(desc.en Obser.)  " & Chr(179) & Space(38) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(3, Chr(196)) & Chr(193) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 4
                       mcad = Chr(179) & "TIEMPO EFECTIVO A LIQUIDAR (1-2)      " & Chr(179) & Space(38) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 5
                       mcad = Chr(179) & "REM.COMPUT.MENSUAL PARA LIQUID.EL PER." & Chr(179) & Space(18) & fCadNum(rs!totalremun, "##,###,##0.00") & Space(20 - Len(fCadNum(rs!totalremun, "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 6
                       mcad = Chr(179) & "Un dozavo de la Rem.compt.mensual     " & Chr(179) & Space(18) & fCadNum((rs!totalremun / 12), "##,###,##0.00") & Space(20 - Len(fCadNum((rs!totalremun / 12), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 7
                       mcad = Chr(179) & "Un 30avo del 12avo de la Rem. Compt.  " & Chr(179) & Space(18) & fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00") & Space(20 - Len(fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 8
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 9
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                  Case Is = 10
                       mcad = Chr(179) & Space(77) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
           End Select
        End If
    Next
    For i = mItem To 10
        mItem = mItem + 1
        mcadrem = Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
        Select Case mItem
               Case Is = 1
                    mcad = Chr(179) & "   " & Format(rs2!fingreso, "dd/mm/yyyy") & "    " & Chr(179) & mobrero & Chr(179) & Space(3)
                    If IsNull(rs2!fcese) Then mcad = mcad & Space(10) Else mcad = mcad & Format(rs2!fcese, "dd/mm/yyyy")
                    mcad = mcad & Space(4) & Chr(179) & Space(26) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(194) & String(13, Chr(196)) & Chr(193) & String(14, Chr(196)) & Chr(193) & String(17, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
                    Print #1, Chr(179) & "   " & Chr(179) & Space(19) & "PERIODO DE SERVICIOS QUE SE CANCELA" & Space(19) & Chr(179) & Space(26) & Chr(179) & Space(27) & Chr(179)
               Case Is = 2
                    If Month(rs2!fingreso) = Cmbmes.ListIndex + 1 And Year(rs2!fingreso) = Val(Txtano.Text) Then
                       mcad = "Del : " & Format(rs2!fingreso, "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & " No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                          cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
                    Else
                       mcad = "Del : " & Format(DateAdd("m", -6, DateAdd("d", 1, mfec)), "dd/mm/yyyy") & "  Al : " & mfec & Space(3) & "No De Meses : " & Mid(rs!recordact, 4, 2) & "  No de Dias : " & Mid(rs!recordact, 6, 3) & Space(3) & mcadrem
                       cadefectivo = Chr(179) & "   POR" & Space(3) & Mid(rs!recordact, 4, 2) & " MES " & Mid(rs!recordact, 6, 3) & " Dias" & Space(8) & Chr(179)
                    End If
                    mcad = Chr(179) & " 1 " & Chr(179) & mcad
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(197) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 3
                    mcad = Chr(179) & " 2 " & Chr(179) & "Periodo no Comp.(desc.en Obser.)" & Space(41) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(3, Chr(196)) & Chr(193) & String(73, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 4
                    mcad = Chr(179) & "TIEMPO EFECTIVO A LIQUIDAR (1-2)" & Space(45) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 5
                        mcad = Chr(179) & "REM.COMPUT.MENSUAL PARA LIQUID.EL PER." & Chr(179) & Space(18) & fCadNum(rs!totalremun, "##,###,##0.00") & Space(20 - Len(fCadNum(rs!totalremun, "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 6
                        mcad = Chr(179) & "Un dozavo de la Rem.compt.mensual     " & Chr(179) & Space(18) & fCadNum((rs!totalremun / 12), "##,###,##0.00") & Space(20 - Len(fCadNum((rs!totalremun / 12), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 7
                       mcad = Chr(179) & "Un 30avo del 12avo de la Rem. Compt.  " & Chr(179) & Space(18) & fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00") & Space(20 - Len(fCadNum(((rs!totalremun / 12) / 30), "##,###,##0.00"))) & mcadrem
                       Print #1, mcad
                       Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 8
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 9
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
               Case Is = 10
                    mcad = Chr(179) & Space(77) & mcadrem
                    Print #1, mcad
                    Print #1, Chr(195) & String(77, Chr(196)) & Chr(197) & String(26, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
        End Select
    Next
    mcadrem = Chr(179)
    Sql = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' and codinterno='" & Right(rs(i).Name, 2) & "'"
    mcadrem = mcadrem & "TOTAL" & Space(21)
    mcadrem = mcadrem & Chr(179) & Space(12) & fCadNum(rs!totalremun, "##,###,##0.00") & Space(2) & Chr(179)
    mcad = Chr(179) & Space(77) & mcadrem
    Print #1, mcad
    Print #1, Chr(195) & String(77, Chr(196)) & Chr(193) & String(26, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(43) & "LIQUIDACION DE LAS CTS CON EFECTO CANCELATORIO" & Space(43) & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(194) & String(71, Chr(196)) & Chr(194) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & Space(3) & "TIEMPO EFECTIVO A LIQUIDAR" & Space(3) & Chr(179) & Space(27) & "CALCULO DE LA CTS" & Space(27) & Chr(179) & "            MONTO          " & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    Print #1, cadefectivo & Space(20) & "TOTAL CTS DEPOSITADA O PAGADA :" & Space(20) & Chr(179) & Space(12) & fCadNum(rs!monto_total, "##,###,##0.00") & Space(2) & Chr(179)
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(197) & String(71, Chr(196)) & Chr(197) & String(27, Chr(196)) & Chr(180)
    mcad = Chr(179) & Space(32) & Chr(179) & Space(20) & Space(27) & Space(24) & Chr(179) & Space(27) & Chr(179)
    Print #1, mcad
    Print #1, Chr(195) & String(32, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(194) & String(43, Chr(196)) & Chr(193) & String(27, Chr(196)) & Chr(180)
    Print #1, Chr(179) & "OBSERVACIONES :" & Space(45) & Chr(179) & " CONSTANCIA DE RECEPCION :" & Space(45) & Chr(179)
    Print #1, Chr(179) & Space(60) & Chr(179) & " 1)De la Presente Liquidacion" & Space(42) & Chr(179)
    Print #1, Chr(179) & Space(60) & Chr(179) & " 2)Del Documento que Acredita el Deposito de la CTS antes detalleda" & Space(4) & Chr(179)
    Print #1, Chr(195) & String(11, Chr(196)) & Chr(194) & String(48, Chr(196)) & Chr(180) & Space(71) & Chr(179)
    Print #1, Chr(179) & "VoB." & Space(7) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & Space(48) & Chr(179) & Space(71) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & String(48, "-") & Chr(179) & String(71, "-") & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & "Nombres,Apellidos y Cargo del Representante del " & Chr(179) & Space(17) & "Firma del Trabajador (Huella Digital)" & Space(17) & Chr(179)
    Print #1, Chr(179) & Space(11) & Chr(179) & "            Empleador(Sello y Firma)            " & Chr(179) & Space(23) & mnombre & Space(7) & Chr(179)
    Print #1, Chr(192) & String(11, Chr(196)) & Chr(193) & String(48, Chr(196)) & Chr(193) & String(71, Chr(196)) & Chr(217)
    
    Print #1, SaltaPag
   rs.MoveNext
Loop
Close #1
Call Imprime_Txt("CertCts.txt", RUTA$)

End Sub

Private Sub PROVICIONES_CTS()
Dim sSQL As String
Dim MAXROW As Long, MAXCOL As Integer, MaxColInicial As Integer
Dim rs As ADODB.Recordset, rsAUX As ADODB.Recordset
Dim CantMes As String, Campo As String
Dim FecIni As String, FecFin As String, FecProceso As String
Dim i As Integer, MaxColTemp As Integer
Dim dblFactor As Currency, CADENA As String
Dim factor_essalud As Currency, totaportes As Currency
Dim sCol As Integer, curfactor As Currency
Dim sSQLI As String, sSQLP As String

Const COL_CODIGO = 0
Const COL_FECING = 1
Const COL_AREA = 2

Call Captura_Tipo_Trabajador

MAXCOL = 2
i = MAXCOL + 1
FecProceso = Format(DateAdd("d", -1, Format(DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text), "dd/mm/yyyy")), "dd/mm/yyyy")

'*************codigo ingresado giovanni 11092007**********************************
Select Case Month(FecProceso)
    Case 1: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 2: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 3: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 4: FecIni = "01/11/" & Val(Txtano.Text) - 1
    Case 5: FecIni = "01/05/" & Val(Txtano.Text)
    Case 6: FecIni = "01/05/" & Val(Txtano.Text)
    Case 7: FecIni = "01/05/" & Val(Txtano.Text)
    Case 8: FecIni = "01/05/" & Val(Txtano.Text)
    Case 9: FecIni = "01/05/" & Val(Txtano.Text)
    Case 10: FecIni = "01/05/" & Val(Txtano.Text)
    Case 11: FecIni = "01/11/" & Val(Txtano.Text)
    Case 12: FecIni = "01/11/" & Val(Txtano.Text)
End Select
'*********************************************************************************

'*************codigo modificado giovanni 11092007*********************************
'If Cmbmes.ListIndex + 1 < 10 Then
 '   FecIni = "01/05/" & Txtano.Text
    'FecFin = Format(DateAdd("d", -1, "01/10/" & Txtano.Text), "DD/MM/YYYY")
    FecFin = FecProceso
'Else
  '  FecIni = "01/10/" & Txtano.Text
    'FecFin = Format(DateAdd("d", -1, CDate("01/04/" & Txtano.Text + 1)), "DD/MM/YYYY")
   ' FecFin = FecProceso
'End If
'*********************************************************************************

Erase ArrReporte

sSQL = "select distinct factor from platasaanexo where status!='*' and tipomovimiento='02' and basecalculo=16 and cia='" & wcia & "'"
If (fAbrRst(rs, sSQL)) Then
    dblFactor = rs(0)
    rs.Close
End If

'**********codigo modificado giovanni 10092007*****************************
'sSQL = "SELECT concepto,campo,sn_promedio,0 AS factor,sn_carga FROM estructura_provisiones WHERE tipo='C' and "
'sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
'sSQL = sSQL & " UNION ALL SELECT concepto,campo,sn_promedio,b.factor,sn_carga FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
'sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.status <> '*' and b.codinterno=a.campo) WHERE a.tipo='C'"
sSQL = "SELECT 'I'+campo as CampUnion,sn_promedio,0 AS factor,sn_carga FROM estructura_provisiones WHERE tipo='C' and "
sSQL = sSQL & " campo NOT IN (SELECT codinterno FROM platasaanexo WHERE cia='" & wcia & "' and tipomovimiento='02' and basecalculo='16')"
sSQL = sSQL & " UNION ALL SELECT 'P'+campo as CampUnion,sn_promedio,b.factor,sn_carga FROM estructura_provisiones a INNER JOIN platasaanexo b ON"
sSQL = sSQL & " (b.cia='" & wcia & "' and b.tipomovimiento='02' and b.basecalculo='16' and b.status <> '*' and b.codinterno=a.campo) WHERE a.tipo='C'"
'**************************************************************************

If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
        MAXCOL = MAXCOL + 1
        rs.MoveNext
    Loop

    rs.MoveFirst
    MaxColTemp = MAXCOL + 1
    MAXCOL = MAXCOL + 9
    
    ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
    CADENA = ""
    Do While Not rs.EOF
        'ArrReporte(I, MAXROW) = Trim(rs!concepto) & rs!Campo
        ArrReporte(i, MAXROW) = Trim(rs!CAMPUNION)
        If CInt(rs!sn_carga) <> 0 Then
            '****************codigo modificado giovanni 08092007************************
            'CADENA = CADENA & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & rs!Campo & "=CASE WHEN dbo.fc_cargapagos('" & wcia & "',a.placod,'" & FecIni & "','" & FecFin & "','" & "I" & Trim(rs!Campo) & "')=1 THEN SUM(COALESCE(" & Trim(rs!concepto) & Trim(rs!Campo) & ",0))"
            If Mid(rs!CAMPUNION, 2, 2) = "29" Then
                CADENA = CADENA & " CASE WHEN (SELECT COUNT(ISNULL(I29,0)) FROM PLAHISTORICO PH WHERE PH.CIA ='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01') >=3 THEN " & _
                         "ISNULL((SELECT SUM(COALESCE(I29,0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='01/11/2007' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01')/6,0) ELSE 0 END as P29"
                
              'CADENA = CADENA & "ISNULL((SELECT SUM(COALESCE(I29,0)) FROM plahistorico ph WHERE ph.cia='06' and ph.status!='*'  and  ph.FECHAPROCESO>='01/11/2007' AND ph.FECHAPROCESO<='30/04/2008' and ph.placod=a.placod AND PH.PROCESO='01')/6,0) as P29 "
            Else
                CADENA = CADENA & IIf(CInt(rs!sn_promedio) = 0, "I", "P") & Mid(rs!CAMPUNION, 2, 2) & "=CASE WHEN dbo.fc_cargapagos2('" & wcia & "',a.placod,'" & Val(Month(FecIni)) & "','" & Val(Month(FecFin)) & "','" & Year(FecFin) & "')=3 THEN (SELECT SUM(COALESCE(I" & Trim(Mid(rs!CAMPUNION, 2, 2)) & ",0)) FROM plahistorico ph WHERE ph.cia='" & wcia & "' and ph.status!='*'  and  ph.FECHAPROCESO>='" & FecIni & "' " & _
                "AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod AND PH.PROCESO='01')"
            End If
            
            
            '***************************************************************************
            If CInt(rs!sn_promedio) = -1 And Trim(Mid(rs!CAMPUNION, 1, 1)) = "A" Then
                CADENA = CADENA & "/(SELECT aportacion FROM PLACONSTANTE WHERE tipomovimiento='03' AND cia='" & wcia & "' AND status!='*' AND codinterno='" & Mid(rs!CAMPUNION, 2, 2) & "'),"
            '********codigo modificado giovanni 10092007***************
            'ElseIf CInt(rs!sn_promedio) = -1 And Trim(Mid(rs!campunion, 1, 1)) = "I" Then
            ElseIf CInt(rs!sn_promedio) = -1 And (Trim(Mid(rs!CAMPUNION, 1, 1)) = "I" Or Trim(Mid(rs!CAMPUNION, 1, 1)) = "P") Then
            '**********************************************************
                                If Mid(rs!CAMPUNION, 2, 2) = "29" Then
                                        CADENA = CADENA & ","
                                Else
                                        If Val(rs!factor) <> 0 Then
                                            CADENA = CADENA & "/" & rs!factor & " ELSE 0 END,"
                                        Else
                                            CADENA = CADENA & "/" & dblFactor & " ELSE 0 END,"
                                        End If
                    
                                End If
                            Else
                     CADENA = CADENA & ","
                End If
        Else
            'CADENA = CADENA & "COALESCE((SELECT prb.importe/(factor_horas/" & hORAS_X_DIA & ") FROM plaremunbase prb WHERE cia='" & wcia & "' and placod=a.placod AND status!='*' and concepto='" & Trim(rs!Campo) & "'),0) as '" & Trim(rs!concepto) & Trim(rs!Campo) & "',"
            CADENA = CADENA & "COALESCE((SELECT prb.importe/(factor_horas/" & hORAS_X_DIA & ") FROM plaremunbase prb WHERE cia='" & wcia & "' and placod=a.placod AND status!='*' and concepto='" & Trim(Mid(rs!CAMPUNION, 2, 2)) & "'),0) as '" & Trim(rs!CAMPUNION) & "',"
        End If
        
        i = i + 1
        rs.MoveNext
    Loop
  
    CADENA = Mid(CADENA, 1, Len(Trim(CADENA)) - 1)
    rs.Close
End If

sSQL = "SET DATEFORMAT DMY SELECT"
sSQL = sSQL & " a.placod,LTRIM(RTRIM(a.ap_pat))+' '+LTRIM(RTRIM(a.ap_mat))+' '+LTRIM(RTRIM(a.ap_cas))+' '+LTRIM(RTRIM(a.nom_1))+' '+LTRIM(RTRIM(a.nom_2)),"
sSQL = sSQL & " fingreso , fcese, a.area,a.tipotrabajador,"
sSQL = sSQL & CADENA
sSQL = sSQL & " FROM planillas a LEFT OUTER JOIN plahistorico ph ON (ph.cia='" & wcia & "' and ph.status!='*' "
sSQL = sSQL & " and ph.FECHAPROCESO>='" & FecIni & "' AND ph.FECHAPROCESO<='" & FecFin & "' and ph.placod=a.placod) "

'************codigo agregado giovanni 05092007**********************
Select Case Format(Cmbmes.ListIndex + 2, "00")
    Case "02": s_Dia_Proceso_Prov = "28"
    Case Else: s_Dia_Proceso_Prov = "30"
End Select
'*******************************************************************

'sSQL = sSQL & " WHERE a.status!='*' and a.cia='" & wcia & "' and (a.fcese >'" & "30/" & Format(CmbMes.ListIndex + 2, "00") & "/" & txtano.Text & "' or a.fcese is null)"
Select Case s_TipoTrabajador_Cts
    Case "01": sSQL = sSQL & " WHERE a.status!='*' and a.cia='" & wcia & "' and (a.fcese >'" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' or a.fcese is null) and a.tipotrabajador='01'"
    Case "02": sSQL = sSQL & " WHERE a.status!='*' and a.cia='" & wcia & "' and (a.fcese >'" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' or a.fcese is null) and a.tipotrabajador='02'"
    Case "03": sSQL = sSQL & " WHERE a.status!='*' and a.cia='" & wcia & "' and (a.fcese >'" & DateAdd("d", -1, DateAdd("m", 1, "01" & "/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Txtano.Text)) & "' or a.fcese is null)"
End Select
sSQL = sSQL & " GROUP BY a.PLACOD , a.ap_pat, a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.fingreso, a.fcese,a.area,a.tipotrabajador"

MAXROW = MAXROW + 1
  Dim PROMEDIO_HRS As Integer
If (fAbrRst(rs, sSQL)) Then
    Do While Not rs.EOF
    PROMEDIO_HRS = 0
     '*******************codigo agregado giovanni 25082007***************
        If Year(rs!fingreso) < Txtano Then GoTo procesa_ProV_CTS_Nueva
        If Month(rs!fingreso) <= Cmbmes.ListIndex + 1 Then
procesa_ProV_CTS_Nueva:
'*******************************************************************
          
            ReDim Preserve ArrReporte(0 To MAXCOL, 0 To MAXROW)
            CantMes = ""
            CantMes = CalcularMeses(rs!fingreso)

            ArrReporte(COL_CODIGO, MAXROW) = Trim(rs!PLACOD)
            ArrReporte(COL_FECING, MAXROW) = rs!fingreso
            ArrReporte(COL_AREA, MAXROW) = rs!Area
            For i = 6 To rs.Fields.count - 1
                      
                sCol = BuscaColumna(rs.Fields(i).Name, MAXCOL)
                If sCol > 0 Then
                    If Trim(rs!TipoTrabajador) = "01" Then
                        Select Case Mid(rs.Fields(i).Name, 1, 1)
                            Case "I"
                                ArrReporte(sCol, MAXROW) = Round(rs.Fields(i).Value * DIAS_TRABAJO, 2)
                            Case "P"
                                If PROMEDIO_HRS = 0 Then
                                    If Mid(rs.Fields(11).Name, 1, 3) = "P29" Then
                                     '  MsgBox ("LO INGRESA A VARIABLE")
                                       ArrReporte(sCol, MAXROW) = Round(rs.Fields(11).Value, 2)
                                       PROMEDIO_HRS = PROMEDIO_HRS + 1
                                       Else
                                            ArrReporte(sCol, MAXROW) = Round(rs.Fields(i).Value, 2)
                                    End If
                                Else
                                        ArrReporte(sCol, MAXROW) = Round(rs.Fields(i).Value, 2)
                                End If
                                
                               
                                'Else
                                    'ArrReporte(sCol, MAXROW) = Round(Rs.Fields(i).Value, 2)
                                'End If
                                
                        End Select
'                        ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value * DIAS_TRABAJO, 2)
                    Else
                        Select Case Mid(rs.Fields(i).Name, 1, 1)
                            Case "I"
                                ArrReporte(sCol, MAXROW) = Round(rs.Fields(i).Value, 2)
                            Case "P"
                                ArrReporte(sCol, MAXROW) = Round(rs.Fields(i).Value / DIAS_TRABAJO, 2)
                        End Select
                        'ArrReporte(sCol, MAXROW) = Round(rs.Fields(I).Value, 2)
                    End If
                    totaportes = totaportes + ArrReporte(sCol, MAXROW)
                End If
            Next

            i = MaxColTemp
        
            ArrReporte(i, MAXROW) = CantMes
        
            ' PROMEDIO GRATIFICACION
            i = i + 1
            ArrReporte(i, 0) = "P15"

            If Cmbmes.ListIndex + 1 > 6 Then
                'sql = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=7 and YEAR(fechaproceso)=" & txtano.Text & " and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
                Sql = "select ISNULL(totaling,0) from plahistorico where cia='" & wcia & "' and proceso='03' and placod='" & Trim(rs!PLACOD) & "' and status!='*' and month(fechaproceso)=7 and YEAR(fechaproceso)=" & Txtano.Text & " "
            Else
                'sql = "select gratmes from plaprovgrati where cia='" & wcia & "' and month(fechaproceso)=12 and YEAR(fechaproceso)=" & txtano.Text - 1 & " and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
                Sql = "select ISNULL(totaling,0) from plahistorico where cia='" & wcia & "' and proceso='03' and placod='" & Trim(rs!PLACOD) & "' and status!='*' and month(fechaproceso)=12 and YEAR(fechaproceso)=" & Txtano.Text - 1 & " "
            End If

        
            If (fAbrRst(rsAUX, Sql)) Then
            '    ArrReporte(I, MAXROW) = Round(rsAUX(0) / 6, 2)
            
                If Trim(rs!TipoTrabajador) = "01" Then
                    ArrReporte(i, MAXROW) = Round(rsAUX(0) / 6, 2)
                Else
                    ArrReporte(i, MAXROW) = Round(rsAUX(0) / 180, 2)
                End If
            
                rsAUX.Close
            Else
                ArrReporte(i, MAXROW) = 0
            End If
            totaportes = totaportes + ArrReporte(i, MAXROW)

            'TOTAL DE JORNAL
            i = i + 1
            ArrReporte(i, MAXROW) = totaportes
        
            'TOTAL DE INDENNISATORIO
            i = i + 1
            ArrReporte(i, MAXROW) = MontoIndecnizado(CantMes, totaportes, rs!TipoTrabajador)
         
            'PROVISION DEL AÑO PASADO
            i = i + 1
            If Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
                Sql = "select monto_total from plaprovcts where cia='" & wcia & "' and month(fechaproceso)=12 and year(fechaproceso)=" & Txtano.Text - 1 & " and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
                If (fAbrRst(rsAUX, Sql)) Then
                    ArrReporte(i, MAXROW) = rsAUX(0)
                    rsAUX.Close
                Else
                    ArrReporte(i, MAXROW) = 0
                End If
            Else
                ArrReporte(i, MAXROW) = 0
            End If
        
            'AJUSTE DE PROVISION
            i = i + 1
            ArrReporte(i, MAXROW) = ArrReporte(i - 2, MAXROW) - ArrReporte(i - 1, MAXROW)
        
            ' PROVISION DE ESTE AÑO
            i = i + 1
            If Cmbmes.ListIndex = 0 Then
                Sql = "select 0"
            Else
                If Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
                    Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>='01/01/" & Txtano.Text & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
                ElseIf Cmbmes.ListIndex + 1 = COL_PRIMERMES Or Cmbmes.ListIndex + 1 = COL_SEGUNDOMES Then
                    Sql = "select 0"
                Else
                    If Cmbmes.ListIndex + 1 >= COL_PRIMERMES And Cmbmes.ListIndex + 1 < COL_SEGUNDOMES Then
                        Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>='01/" & COL_PRIMERMES & "/" & Txtano.Text & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
                    Else
                        Sql = "select COALESCE(SUM(COALESCE(provision_actual,0)),0) from plaprovcts where cia='" & wcia & "' and fechaproceso>='01/" & COL_SEGUNDOMES & "/" & Txtano.Text & "' AND fechaproceso<'01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text & "' and placod='" & Trim(rs!PLACOD) & "' and status<>'*'"
                    End If
                End If
                
            End If
            If (fAbrRst(rsAUX, Sql)) Then
                ArrReporte(i, MAXROW) = rsAUX(0)
                rsAUX.Close
            Else
                ArrReporte(i, MAXROW) = 0
            End If
        
            'PROVISION DEL MES
            i = i + 1
            
            If ArrReporte(i - 2, MAXROW) < ArrReporte(i - 1, MAXROW) Then
                ArrReporte(i, MAXROW) = 0
               ' ArrReporte(i, MAXROW) = ArrReporte(i - 2, MAXROW)
            Else
                ArrReporte(i, MAXROW) = ArrReporte(i - 2, MAXROW) - ArrReporte(i - 1, MAXROW)
            End If
            
            MAXROW = MAXROW + 1
        
            'RECORREMOS LOS CAMPOS DE INGRESO DE LA TABLA X(
            sSQLI = ""
            For i = 1 To 50
                Campo = "I" & Format(i, "00")
                sCol = BuscaColumna(Campo, MAXCOL)
                If sCol > 0 Then
                    sSQLI = sSQLI & IIf(Len(Trim(ArrReporte(sCol, MAXROW - 1))) = 0, "0", ArrReporte(sCol, MAXROW - 1)) & ","
                Else
                    sSQLI = sSQLI & "0,"
                End If
            Next

            'RECORREMOS LOS CAMPOS DE PROMEDIO DE LA TABLA X(
            sSQLP = ""
            For i = 1 To 50
                Campo = "P" & Format(i, "00")
                sCol = BuscaColumna(Campo, MAXCOL)
                If sCol > 0 Then
                    sSQLP = sSQLP & ArrReporte(sCol, MAXROW - 1) & ","
                Else

                    sSQLP = sSQLP & "0,"
                End If
            Next

            sSQL = ""
            sSQL = "INSERT plaprovcts VALUES ('" & wcia & "','" & ArrReporte(COL_CODIGO, MAXROW - 1) & "','" & Trim(rs!TipoTrabajador) & "','" & ArrReporte(MaxColTemp, MAXROW - 1) & "','" & ArrReporte(COL_AREA, MAXROW - 1) & "',"
            sSQL = sSQL & sSQLI & sSQLP & ArrReporte(MaxColTemp + 2, MAXROW - 1) & ",0,0,0," & ArrReporte(MaxColTemp + 3, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 4, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 5, MAXROW - 1) & ","
            sSQL = sSQL & ArrReporte(MaxColTemp + 6, MAXROW - 1) & "," & ArrReporte(MaxColTemp + 7, MAXROW - 1) & ",0,'" & Format(FecProceso, "DD/MM/YYYY") & "',' ')"
        
            cn.Execute (sSQL)
            totaportes = 0
        
        End If
'        End If
        rs.MoveNext

    Loop
    
End If
Carga_Cts
End Sub


Private Function CalcularMeses(ByVal pFecIngreso As String) As String
Dim mesestmp As String
Dim año As String, mes As String, Dia As String
Dim FecIngTmp As String
Dim pFecproc As String
Dim pFecValida As String


año = 0: mes = 0: Dia = 0
If Year(CDate(pFecIngreso)) > Val(Txtano.Text) Then GoTo salir

    If Cmbmes.ListIndex + 1 >= COL_SEGUNDOMES Or Cmbmes.ListIndex + 1 < COL_PRIMERMES Then
        pFecValida = DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)
        pFecValida = DateAdd("d", -1, "01/" & Month(pFecValida) & "/" & Year(pFecValida))
        
        If Year(CDate(pFecIngreso)) < Year(CDate(pFecValida)) Then
            
            If Month(CDate(pFecValida)) >= 1 And Month(CDate(pFecValida)) < COL_PRIMERMES Then
            'ESTE CAMBIO ES DE EMERGENCIA SE DEBERA DE MODIFICAR DESPUES
               If Day(CDate(pFecIngreso)) > 1 And Month(CDate(pFecIngreso)) = COL_SEGUNDOMES And (Year(pFecValida) - Year(CDate(pFecIngreso))) = 1 Then
                    mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES + 1 & "/" & Txtano.Text - 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
                    If Day(pFecIngreso) = 1 Then
                        Dia = 0
                        pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
                    Else
                        pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
                        Dia = DateDiff("d", pFecIngreso, pFecproc)
                        pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
                    End If
                Else
                    mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text - 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
               End If
                
                'mes = Abs(DateDiff("m", CDate(pFecIngreso), "01/" & CmbMes.ListIndex + 1 & "/" & Txtano.Text)) + 1
            Else
                mes = Abs(DateDiff("m", "01/" & COL_SEGUNDOMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
            End If
            
        Else
            año = 0
            'Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
            
'            pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
'            Dia = DateDiff("d", pFecIngreso, pFecproc)
'            pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
'            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
            
            If Day(pFecIngreso) = 1 Then
                Dia = 0
                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
            Else
                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
                Dia = DateDiff("d", pFecIngreso, pFecproc)
                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
            End If
            
            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
            
            If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
            
        End If
    Else

        '************codigo modificado giovanni 23082007******************************
        'If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & txtano.Text) Or pFecIngreso < CDate("01/" & COL_SEGUNDOMES & "/" & txtano.Text) Then
        If pFecIngreso < CDate("01/" & COL_PRIMERMES & "/" & Txtano.Text) Then
        '*****************************************************************************

            mes = Abs(DateDiff("m", "01/" & COL_PRIMERMES & "/" & Txtano.Text, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)) + 1
                 
        Else
'            año = 0
'            mes = Abs(DateDiff("m", pFecIngreso, DateAdd("d", -1, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
'            Dia = Abs(DateDiff("d", DateAdd("m", mes, pFecIngreso), DateAdd("d", -2, DateAdd("m", 1, "01/" & CmbMes.ListIndex + 1 & "/" & txtano.Text))))
            
            año = 0
            If Day(pFecIngreso) = 1 Then
                Dia = 0
                pFecproc = DateAdd("d", -1, DateAdd("m", 1, pFecIngreso))
            Else
                pFecproc = DateAdd("d", -1, "01/" & Month(DateAdd("m", 1, pFecIngreso)) & "/" & Year(DateAdd("m", 1, pFecIngreso)))
                Dia = DateDiff("d", pFecIngreso, pFecproc)
                pFecproc = DateAdd("d", Val(Dia), pFecIngreso)
            End If
            
            mes = DateDiff("m", pFecproc, DateAdd("d", -1, DateAdd("m", 1, "01/" & Cmbmes.ListIndex + 1 & "/" & Txtano.Text)))
            
            If Day(pFecIngreso) = 1 Then mes = CStr(Val(mes) + 1)
        End If
                
    End If
    
salir:
'If mes > 6 Then mes = mes - 6
CalcularMeses = Space(2 - Len(Trim(año))) & año & " " & Space(2 - Len(Trim(mes))) & mes & " " & Space(2 - Len(Trim(Dia))) & Dia & " "
End Function

Private Function MontoIndecnizado(ByVal pValor As String, ByVal pImporte As String, ByVal pTipoTrab As String) As String
Dim montotmp As Currency

If pTipoTrab = "02" Then
    montotmp = Val(Mid(pValor, 4, 2)) * ((pImporte * DIAS_TRABAJO) / 12)
    montotmp = Round(montotmp + Val(Mid(pValor, 7, 2)) * (pImporte / 12), 2)
Else
    montotmp = Val(Mid(pValor, 4, 2)) * (pImporte / 12)
    montotmp = Round(montotmp + Val(Mid(pValor, 7, 2)) * ((pImporte / 12) / DIAS_TRABAJO), 2)
End If

MontoIndecnizado = montotmp

End Function

Private Function BuscaColumna(ByVal pCampo As String, ByVal pMaxcol As Integer) As Integer
Dim iRow As Integer
BuscaColumna = 0
For iRow = 3 To pMaxcol
    If ArrReporte(iRow, 0) = pCampo Then
        BuscaColumna = iRow
        Exit Function
    End If
Next
End Function


Private Sub ReporteCts()
Dim rs As Object
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim i As Long
Dim FILA As Integer
Dim COLUMNA As Integer
Dim ArrTotales() As Variant

ReDim Preserve ArrTotales(0 To 23)

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 7
xlSheet.Range("B:B").ColumnWidth = 23.29
xlSheet.Range("C:C").ColumnWidth = 9.71
xlSheet.Range("D:S").ColumnWidth = 10.3
xlSheet.Range("C:S").HorizontalAlignment = xlCenter

xlSheet.Cells(1, 1).Value = Trim(Cmbcia.Text)
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(2, 2).Value = " REPORTE DE PROVISION POR TIEMPO DE SERVICIO "
xlSheet.Cells(2, 2).Font.Bold = True
xlSheet.Cells(2, 2).Font.Size = 12
xlSheet.Cells(2, 2).HorizontalAlignment = xlCenter
xlSheet.Range("B2:S2").Merge

xlSheet.Range("A6:S6").Merge
xlSheet.Range("A6:S6") = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
FILA = 7
xlSheet.Cells(FILA, 1).Value = "Codigo"
xlSheet.Cells(FILA, 2).Value = "Nombre Trabajador"
xlSheet.Cells(FILA, 3).Value = "F.Ing"
xlSheet.Cells(FILA, 4).Value = "F.Cese"
xlSheet.Cells(FILA, 5).Value = "Tiempo Servicio"
xlSheet.Cells(FILA, 6).Value = "Jornal"
xlSheet.Cells(FILA + 1, 6).Value = "Basico"
xlSheet.Cells(FILA, 7).Value = "Bonif"
xlSheet.Cells(FILA + 1, 7).Value = "Afp 3%"
xlSheet.Cells(FILA, 8).Value = "Bonif"
xlSheet.Cells(FILA + 1, 8).Value = "Costo Vida"
xlSheet.Cells(FILA, 9).Value = "Bonif"
xlSheet.Cells(FILA + 1, 9).Value = "T.Servicio"
xlSheet.Cells(FILA, 10).Value = "Asignacion"
xlSheet.Cells(FILA + 1, 10).Value = "Familiar"
xlSheet.Cells(FILA, 11).Value = "Promedio"
xlSheet.Cells(FILA + 1, 11).Value = "Gratific"
xlSheet.Cells(FILA, 12).Value = "Promedio "
xlSheet.Cells(FILA + 1, 12).Value = "H.Extras"
xlSheet.Cells(FILA, 13).Value = "Promedio"
xlSheet.Cells(FILA + 1, 13).Value = "Otros Pagos"
xlSheet.Cells(FILA, 14).Value = "Promedio"
xlSheet.Cells(FILA + 1, 14).Value = "H. Verano"
xlSheet.Cells(FILA, 15).Value = "Promedio"
xlSheet.Cells(FILA + 1, 15).Value = "P.xTurno"
xlSheet.Cells(FILA, 16).Value = "Promedio"
xlSheet.Cells(FILA + 1, 16).Value = "P.xProducc"
xlSheet.Cells(FILA, 17).Value = "Promedio"
xlSheet.Cells(FILA + 1, 17).Value = "Bo.Prod."
xlSheet.Cells(FILA, 18).Value = "Promedio"
xlSheet.Cells(FILA + 1, 18).Value = "Rendim."
xlSheet.Cells(FILA, 19).Value = "Jornal"
xlSheet.Cells(FILA + 1, 19).Value = "Indennizat"
FILA = FILA + 2
xlSheet.Cells(FILA, 3).Value = "Tiempo"
xlSheet.Cells(FILA + 1, 3).Value = "T.Serv"
xlSheet.Cells(FILA, 4).Value = "Monto Inden."
xlSheet.Cells(FILA + 1, 4).Value = "Periodo 1"
xlSheet.Cells(FILA, 5).Value = "Monto Inden."
xlSheet.Cells(FILA + 1, 5).Value = "Periodo 2"
xlSheet.Cells(FILA, 6).Value = "Monto Inden"
xlSheet.Cells(FILA + 1, 6).Value = "Sin Topes"
xlSheet.Cells(FILA, 7).Value = "Monto Total"
xlSheet.Cells(FILA + 1, 7).Value = "Indennizatorio"
xlSheet.Cells(FILA, 8).Value = "Provisionado"
xlSheet.Cells(FILA + 1, 8).Value = "Año Anterior"
xlSheet.Cells(FILA, 9).Value = "Ajuste de"
xlSheet.Cells(FILA + 1, 9).Value = "Provision"
xlSheet.Cells(FILA, 10).Value = "Provisionado"
xlSheet.Cells(FILA + 1, 10).Value = "Este Año"
xlSheet.Cells(FILA, 11).Value = "Provision "
xlSheet.Cells(FILA + 1, 11).Value = "Mes Actual"
xlSheet.Cells(FILA, 12).Value = "Saldo Pend."
xlSheet.Cells(FILA + 1, 12).Value = "De Provision"
FILA = FILA + 2
xlSheet.Range("A" & FILA & ":S" & FILA).Merge
xlSheet.Range("A" & FILA & ":S" & FILA) = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
FILA = FILA + 4

Captura_Tipo_Trabajador

Select Case s_TipoTrabajador_Cts
    Case "01"
        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
        "CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='01' ORDER BY A.PLACOD"
    Case "02"
        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
        "CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='02' ORDER BY A.PLACOD"
    Case "03"
        Sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
        "CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
        "monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
        "a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " ORDER BY A.PLACOD"
End Select
    

'sql = "SELECT a.placod,ltrim(rtrim(b.ap_pat))+' '+ltrim(rtrim(b.ap_mat))," & _
'"CONVERT(VARCHAR(10),b.fingreso,103),CONVERT(VARCHAR(10),b.fcese,103),'',i01,i06,i07,i04,i02,p15,p10+p11+p21,p16,0,0, p24+p25,p26,0,totalremun,a.recordact,monto_inde1,monto_inde2,monto_indetope," & _
'"monto_total,prov_anoante,ajuste_prov,provision_ano,provision_actual,saldo_prov FROM PLAPROVCTS a LEFT OUTER JOIN PLANILLAS b on (b.cia=a.cia and b.placod=a.placod and b.status!='*') WHERE " & _
'"a.cia='" & wcia & "' and year(fechaproceso)=" & Txtano.Text & " and a.status!='*' and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " and b.tipotrabajador='01' ORDER BY A.PLACOD"

If (fAbrRst(rs, Sql)) Then
    Do While Not rs.EOF
        
        FILA = FILA + 1
        For i = 0 To rs.Fields.count - 1
            If i = 19 Then
                COLUMNA = 2
                FILA = FILA + 1
            End If
            COLUMNA = COLUMNA + 1
            If i > 4 Then
                If i = 19 Then
                    xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(i))
                Else
                    xlSheet.Cells(FILA, COLUMNA).Value = Format(Trim(rs(i)), "###,###,##0.00")
                    xlSheet.Cells(FILA, COLUMNA).NumberFormat = "#,###,##0.00"
                End If
                'xlSheet.Cells(FILA, COLUMNA).Value = Format(Trim(rs(I)), "###,###,##0.00")
                'xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(I))
                'xlSheet.Cells(FILA, COLUMNA).NumberFormat = "#,###,##0.00"
                If i <> 19 Then ArrTotales(i - 5) = ArrTotales(i - 5) + Val(rs(i)) Else ArrTotales(i - 5) = ""
            Else
                If i = 2 Then
                    xlSheet.Cells(FILA, COLUMNA).Value = Format(Trim(rs(i)), "MM/DD/YYYY")
                Else
                    xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(i))
                End If
                'xlSheet.Cells(FILA, COLUMNA).Value = Trim(rs(I))
                
            End If
        Next
        COLUMNA = 0
        
        rs.MoveNext
    Loop
End If

COLUMNA = 5
FILA = FILA + 3
For i = 0 To UBound(ArrTotales)
    If i = 14 Then COLUMNA = 2: FILA = FILA + 1
    COLUMNA = COLUMNA + 1
    If i <> 14 Then
        xlSheet.Cells(FILA, COLUMNA).Value = Format(Trim(ArrTotales(i)), "###,###,##0.00")
        xlSheet.Cells(FILA, COLUMNA).NumberFormat = "#,###,##0.00"
    Else
        xlSheet.Cells(FILA, COLUMNA).Value = ArrTotales(i)
    End If
Next i

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE CTS"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

End Sub
'**************codigo nuevo giovanni 17082007*********************************
Sub Generar_Asientos_Contables_Provision()
    Dim s_Centro_Costo As String
    Dim s_Cuenta9 As String
    Dim s_CuentaProvision As String
    Dim s_Cuenta6 As String
    Dim s_Cuenta7 As String
    Dim i_Contador_Vueltas As Integer
    Dim s_Tipo_Trabajador_Prov As String

    Call Captura_Mes_Seleccionado
    Call Elimina_Registros_Existentes(wcia, "14")
    
    Call Recupera_Provision_Liquidacion(wcia, Txtano, s_MesSeleccion)
    Set rs_Liquidacion = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    If rs_Liquidacion.EOF = False Then
        Do While Not rs_Liquidacion.EOF
            If Recupera_Centro_Costo(wcia, rs_Liquidacion!ccosto) = True Then
                s_Centro_Costo = RTrim(Crear_Plan_Contable.s_CentroCostoPub)
                If s_Centro_Costo <> "" Then
                    Call Recupera_Informacion_Parametros_Asiento_Provision(wcia, rs_Liquidacion!tipo, _
                    "TIEMPO SERVICIO")
                    Set rs_Liquidacion2 = Reportes_Centrales.rs_RptCentrales_pub
                    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
                    s_Cuenta9 = s_Centro_Costo & rs_Liquidacion2!ctacentrocosto
                    s_CuentaProvision = rs_Liquidacion2!CtaProvision
                    Set rs_Liquidacion2 = Nothing
                    Call Codigo_Empresa_Starsoft
                    Call Conectar_Base_Datos_Access(s_CodEmpresa_Starsoft)
                    Call Recuperar_Cuentas_Naturaleza_Transferencia(s_Cuenta9)
                    Set rs_Liquidacion2 = CompRetenciones.rs_compRetenciones_Pub
                    Set CompRetenciones.rs_compRetenciones_Pub = Nothing
                    s_Cuenta6 = rs_Liquidacion2!plancta_Cargo1
                    s_Cuenta7 = rs_Liquidacion2!plancta_abono1
                    Call Cerrar_Conexion_Base_Datos_Access
                    Call Genera_Numero_Voucher
                    Select Case rs_Liquidacion!tipo
                        Case "01": s_Tipo_Trabajador_Prov = "1"
                        Case "02": s_Tipo_Trabajador_Prov = "0"
                    End Select
                    For i_Contador_Vueltas = 1 To 4
                        Select Case i_Contador_Vueltas
                            Case 1
                                Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!codigo, Txtano, s_MesSeleccion, _
                                "00", i_Numero_VoucherG, s_Cuenta9, "", rs_Liquidacion!provmes, 1, 0, 14, s_Tipo_Trabajador_Prov, 14)
                            Case 2
                                Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!codigo, Txtano, s_MesSeleccion, _
                                "00", i_Numero_VoucherG, s_Cuenta6, "", rs_Liquidacion!provmes, 1, 0, 14, s_Tipo_Trabajador_Prov, 14)
                            Case 3
                                Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!codigo, Txtano, s_MesSeleccion, _
                                "00", i_Numero_VoucherG, s_Cuenta7, "", rs_Liquidacion!provmes, 2, 1, 14, s_Tipo_Trabajador_Prov, 14)
                            Case 4
                                Call Ejecuta_Asientos_Contables(wcia, rs_Liquidacion!codigo, Txtano, s_MesSeleccion, _
                                "00", i_Numero_VoucherG, s_CuentaProvision, "", rs_Liquidacion!provmes, 2, 1, 14, s_Tipo_Trabajador_Prov, 14)
                        End Select
                    Next i_Contador_Vueltas
                    rs_Liquidacion.MoveNext
                End If
            End If
        Loop
    End If
    MsgBox "Asientos Generados Satisfactoriamente"
End Sub
Sub Captura_Mes_Seleccionado()
    Select Case Cmbmes.Text
        Case "ENERO": s_MesSeleccion = "01": Case "FEBRERO": s_MesSeleccion = "02"
        Case "MARZO": s_MesSeleccion = "03": Case "ABRIL": s_MesSeleccion = "04"
        Case "MAYO": s_MesSeleccion = "05": Case "JUNIO": s_MesSeleccion = "06"
        Case "JULIO": s_MesSeleccion = "07": Case "AGOSTO": s_MesSeleccion = "08"
        Case "SETIEMBRE": s_MesSeleccion = "09": Case "OCTUBRE": s_MesSeleccion = "10"
        Case "NOVIEMBRE": s_MesSeleccion = "11": Case "DICIEMBRE": s_MesSeleccion = "12"
    End Select
End Sub
Sub Ejecuta_Asientos_Contables(CodCompañia As String, CodTrabajador As String, Año_Proceso As _
String, MesSeleccion As String, SemProceso As String, Voucher As String, CtaContable As String, _
DescCta As String, MontoInt As Double, Opcion As Integer, TipoAsiento As String, TipoBoleta As _
String, TipoTrabajador As String, Tipo_Boleta As String)
    If Verifica_Existencia_Registro(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
    SemProceso, Voucher, CtaContable, Tipo_Boleta) = False Then
        Select Case Opcion
            Case 1
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, 0, MontoInt, TipoBoleta, _
                TipoTrabajador)
            Case 2
                Call Graba_Asiento_Contable(CodCompañia, CodTrabajador, Año_Proceso, MesSeleccion, _
                SemProceso, Voucher, CtaContable, TipoAsiento, DescCta, MontoInt, 0, TipoBoleta, _
                TipoTrabajador)
        End Select
    End If
End Sub
Sub Codigo_Empresa_Starsoft()
    Call Recuperar_Codigo_Empresa_Starsoft(wcia)
    Set rs_Liquidacion2 = Reportes_Centrales.rs_RptCentrales_pub
    Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
    s_CodEmpresa_Starsoft = rs_Liquidacion2!ciastar
    Set rs_Liquidacion2 = Nothing
End Sub
Sub Genera_Numero_Voucher()
    i_Numero_Voucher = i_Numero_Voucher + 1
    Select Case i_Numero_Voucher
        Case Is < 10: i_Numero_VoucherG = "000" & i_Numero_Voucher
        Case Is < 100: i_Numero_VoucherG = "00" & i_Numero_Voucher
        Case Is < 1000: i_Numero_VoucherG = "0" & i_Numero_Voucher
        Case Is < 10000: i_Numero_VoucherG = i_Numero_Voucher
    End Select
End Sub
Sub Captura_Tipo_Trabajador()
    'Captura_Tipo_Trabajador
    Select Case Cmbtipo.Text
        Case "EMPLEADO": s_TipoTrabajador_Cts = "01"
        Case "OBRERO": s_TipoTrabajador_Cts = "02"
        Case "TOTAL": s_TipoTrabajador_Cts = "03"
    End Select
End Sub

