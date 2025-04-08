VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSelCia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Selección de Compañia «"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ControlBox      =   0   'False
   Icon            =   "FrmSelCia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "FrmSelCia.frx":030A
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   0
      TabIndex        =   2
      Top             =   -90
      Width           =   5160
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "FrmSelCia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SSCommand1_Click
End If
End Sub

Private Sub Form_Load()
DoEvents
Me.Top = 4000
Me.Left = 2900
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmSelCia = Nothing
End Sub

Public Sub SSCommand1_Click()
Sql$ = "SELECT ctemonedac,senati,ruc,minimo,NOMBREBD,RAZSOC,FactAsigFamiliar from cia where cod_cia='" & wcia & "' and status<>'*'"

If (fAbrRst(rs, Sql$)) Then
   wmoncont = rs!ctemonedac
   sueldominimo = IIf(IsNull(rs!Minimo), 0, rs!Minimo)
   porcasigfamiliar = IIf(IsNull(rs!FactAsigFamiliar), 0, rs!FactAsigFamiliar / 100)
   If IsNull(rs!senati) Then wSenati = "" Else wSenati = rs!senati
   wNomBd = Trim(rs!NOMBREBD & "")
   NOMBREEMPRESA = Trim(rs!razsoc & "")
End If
If rs.State = 1 Then rs.Close

'verificar usuario con autorizacion
Sql$ = "select * from users where cod_cia='" & wcia & "' and  status<>'*' and name_user='" & wuser & "' and sistema='" & wCodSystem & "'"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount <= 0 And StrConv(wuser, 1) <> wAdmin Then
   MsgBox "Usuario sin Autorizacion", vbCritical, "Acceso Al sistema de Bancos"
   Exit Sub
End If
If rs.State = 1 Then rs.Close

wTipoPla = Tipo_Planilla

'' Tipo de Cambio
'SQL$ = "SELECT compra,contable,factor from tipo_cambio where fec_crea " _
' & " BETWEEN '" & Format(Date, FormatFecha) + FormatTimei & "' AND '" & Format(Date, FormatFecha) + FormatTimef & "' " _
' & " AND cod_cia='" & wcia & "'"
'
'cn.CursorLocation = adUseClient
'Set RS = New ADODB.Recordset
'Set RS = cn.Execute(SQL$, 64)
'If RS.RecordCount > 0 Then
'   mtipo_cambio = RS!compra
'   mtc_contable = RS!contable
'   findmon = True
'Else
'   MsgBox "No hay Tipo de Cambio del Día,Por Favor Ingrese", vbInformation
'   Load Frmtipcamb
'   Frmtipcamb.Show
'   Frmtipcamb.ZOrder 0
'End If


MDIplared.Enabled = True
'MDIplared.SetFocus
Unload FrmSelCia

MDIplared.Caption = "SISTEMA DE PLANILLAS - " & CmbCia.Text & " Versión 2016 " & "[ " & App.Major & "." & App.Minor & "." & App.Revision & " ]"
MDIplared.Activa_Menu
Set rs = New ADODB.Recordset
Set rs = cn.Execute("Select cteigv,ruc from cia where cod_cia='" & wcia & "' and status<>'*'")
If rs.RecordCount <= 0 Then
   MsgBox "No Existen Datos de Compañia", vbCritical, "Planillas"
   Load Frmcia
   Frmcia.Show
   Frmcia.ZOrder 0
Else
   wigv = rs!cteigv * 100
   wruc = Trim(rs!RUC)
End If
End Sub
