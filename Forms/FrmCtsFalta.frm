VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmCtsFalta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTS DIAS FALTA"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8100
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   4320
      TabIndex        =   18
      Top             =   440
      Width           =   3615
      Begin VB.OptionButton OptFaltas 
         Caption         =   "Faltas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptSubsidios 
         Caption         =   "Subsidios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin Threed.SSPanel SpnlRegistra 
      Height          =   2175
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   6975
      _Version        =   65536
      _ExtentX        =   12303
      _ExtentY        =   3836
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.TextBox TxtDias 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtplacod 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtpersonal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   840
         Width           =   5655
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   720
         Left            =   4800
         TabIndex        =   14
         ToolTipText     =   "Salir"
         Top             =   1320
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   1270
         _StockProps     =   78
         Caption         =   "Aceptar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Picture         =   "FrmCtsFalta.frx":0000
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   720
         Left            =   5880
         TabIndex        =   15
         ToolTipText     =   "Salir"
         Top             =   1320
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   1270
         _StockProps     =   78
         Caption         =   "Eliminar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Picture         =   "FrmCtsFalta.frx":059A
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   225
         Left            =   6560
         TabIndex        =   16
         ToolTipText     =   "Salir"
         Top             =   60
         Width           =   345
         _Version        =   65536
         _ExtentX        =   609
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "X"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Picture         =   "FrmCtsFalta.frx":0C14
      End
      Begin VB.Label LblTit 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   6735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dias"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Personal"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.TextBox Txtano 
      Height          =   315
      Left            =   3150
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "FrmCtsFalta.frx":0C30
      Left            =   1080
      List            =   "FrmCtsFalta.frx":0C58
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid dgctsfalta 
      Bindings        =   "FrmCtsFalta.frx":0CC0
      Height          =   6735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11880
      _Version        =   393216
      Appearance      =   0
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "placod"
         Caption         =   "Codigo"
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
         DataField       =   "empleado"
         Caption         =   "Empleado"
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
      BeginProperty Column02 
         DataField       =   "dias"
         Caption         =   "Dias"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5625.071
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   524.976
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoctsfalta 
      Height          =   330
      Left            =   240
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.ComboBox Cmbcia 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000001&
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
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   315
      Left            =   3750
      TabIndex        =   5
      Top             =   600
      Width           =   255
      Size            =   "450;556"
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCtsFalta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmbmes_Click()
carga_cts_falta
End Sub

Private Sub dgctsfalta_DblClick()
Dim lTipo As String
If OptSubsidios.Value = True Then lTipo = "S" Else lTipo = "F"
Sql$ = "Select dias From pla_cts_faltas where cia='" & wcia & "' and tipo='" & lTipo & "' and placod='" & Trim(dgctsfalta.Columns(0) & "") & "' "
Sql$ = Sql$ & " and ayo=" & Txtano.Text & " and mes=" & Cmbmes.ListIndex + 1 & " and status<>'*'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then
   Nuevo_Registro (False)
   txtplacod.Text = Trim(dgctsfalta.Columns(0) & "")
   TxtDias.Text = rs!dias
   txtplacod.Enabled = False
   Busca_Trabajador
End If
rs.Close: Set rs = Nothing
End Sub

Private Sub Form_Load()
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Txtano.Text = Year(Date)

End Sub
Public Sub carga_cts_falta()
Dim lTipo As String
If OptSubsidios.Value = True Then lTipo = "S" Else lTipo = "F"
    If Trim(Cmbmes.Text & "") = "" Then Exit Sub
    Sql$ = "select fcts.placod, rtrim(ap_pat) + ' ' + rtrim(ap_mat) + ' ' + rtrim(nom_1) + ' ' + rtrim(nom_2) as empleado ,dias " & _
    " from pla_cts_faltas fcts inner join planillaS pl " & _
    " on fcts.placod=pl.placod and fcts.cia=pl.cia and pl.status<>'*' and fcts.status<>'*' " & _
    " where fcts.cia='" & wcia & "' and tipo='" & lTipo & "' and fcts.ayo=" & Txtano.Text & " and fcts.mes=" & Cmbmes.ListIndex + 1 & " ORDER BY fcts.placod"
    
    cn.CursorLocation = adUseClient
    Set adoctsfalta.Recordset = cn.Execute(Sql$, 64)
    If adoctsfalta.Recordset.RecordCount > 0 Then adoctsfalta.Recordset.MoveFirst
    dgctsfalta.Refresh
    Screen.MousePointer = vbDefault
End Sub
Public Sub Eliminar()
If MsgBox("Desea Eliminar DiaFalta ? ", vbQuestion + vbYesNo + vbDefaultButton1, "Seteo de Usuario") = vbNo Then
    Exit Sub
Else
    cod_per = Me.dgctsfalta.Columns(0)
    fechafalta = Me.dgctsfalta.Columns(2)
    Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    Sql$ = Sql$ & "update cts_falta  set status='*' where cia=  '" & wcia & "' and placod= '" & cod_per & "' and fecha_falta= convert(datetime,'" & fechafalta & "',103)"
    cn.Execute Sql$
    carga_cts_falta
    dgctsfalta.Refresh
End If
End Sub
Public Sub Nuevo_Registro(lNew As Boolean)
If Trim(Cmbmes.Text & "") = "" Then MsgBox "Seleccione Mes", vbInformation: Exit Sub
If Not IsNumeric(Txtano.Text) Then MsgBox "Ingrese Año Correctamente": Exit Sub

If Cmbmes.ListIndex + 1 <> 4 And Cmbmes.ListIndex + 1 <> 10 Then
   MsgBox "Solo se pueden registrar dias para los meses de OCTUBRE y ABRIL", vbInformation
   Exit Sub
End If

If OptFaltas.Value = True Then
   LblTit.Caption = "Dias Faltas"
Else
   LblTit.Caption = "Dias Subsidiados"
End If
txtplacod.Text = ""
txtpersonal.Text = ""
TxtDias.Text = ""
dgctsfalta.Enabled = False
Txtano.Enabled = False
Cmbmes.Enabled = False
Frame1.Enabled = False
SpinButton1.Enabled = False
txtplacod.Enabled = True
If lNew Then SSCommand1.Visible = False Else SSCommand1.Visible = True
SpnlRegistra.Visible = True
If txtplacod.Enabled = True Then txtplacod.SetFocus
End Sub

Private Sub OptFaltas_Click()
carga_cts_falta
End Sub

Private Sub OptSubsidios_Click()
carga_cts_falta
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

Private Sub SSCommand1_Click()
Graba_Dias (3)
Habilita_Reg
End Sub

Private Sub SSCommand2_Click()
Habilita_Reg
End Sub
Private Sub Habilita_Reg()
dgctsfalta.Enabled = True
Txtano.Enabled = True
Cmbmes.Enabled = True
Frame1.Enabled = True
SpinButton1.Enabled = True
SpnlRegistra.Visible = False

End Sub

Private Sub SSCommand3_Click()
Busca_Trabajador
If Trim(txtplacod.Text & "") = "" Then MsgBox "Ingrese Codigo de trabajador", vbInformation: Exit Sub
If Not IsNumeric(TxtDias.Text) Then MsgBox "Ingrese Correctamente Dias", vbInformation: Exit Sub
If CCur(TxtDias.Text) = 0 Then MsgBox "Ingrese Correctamente Dias", vbInformation: Exit Sub
Graba_Dias (1)
Nuevo_Registro (True)
End Sub
Private Sub Graba_Dias(lAccion As Integer)
Dim lTipo As String
If OptSubsidios.Value = True Then lTipo = "S" Else lTipo = "F"

If lExcluye = "S" Then lExcluye = "" Else lExcluye = "S"
Sql$ = "Usp_Pla_Cts_Faltas '" & wcia & "','" & lTipo & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & ",'" & Trim(txtplacod.Text) & "'," & TxtDias.Text & ",'" & wuser & "'," & lAccion & ""
cn.Execute Sql$
carga_cts_falta
End Sub

Private Sub Txtano_Change()
carga_cts_falta
End Sub

Private Sub TxtDias_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Or KeyAscii = 44 Then KeyAscii = 0
End Sub

Private Sub txtplacod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Frmgrdpla.Show
    Frmgrdpla.ZOrder 0
End If

End Sub

Private Sub txtplacod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtDias.SetFocus
End Sub

Private Sub txtplacod_LostFocus()
Busca_Trabajador
End Sub
Private Sub Busca_Trabajador()
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = nombre()
Sql$ = Sql$ & "placod,tipotrabajador from planillas where status<>'*' " _
     & "and cia='" & wcia & "' AND placod='" & Trim(txtplacod.Text) & "'"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$)
If rs.RecordCount > 0 Then
   txtpersonal.Text = Space(5) & Trim(rs!nombre)
   TxtDias.SetFocus
ElseIf Trim(txtplacod.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & txtplacod.Text
   txtplacod.Text = ""
   txtpersonal.Text = ""
   txtplacod.SetFocus
End If
End Sub
