VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frmgrdsubsidio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relacion de Subsidios y Liquidaciones"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "Frmgrdsubsidio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frame3 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   8415
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "Frmgrdsubsidio.frx":030A
         Height          =   5175
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   9128
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
            DataField       =   "fechaproceso"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "totaling"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """S/."" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "user_crea"
            Caption         =   "user_crea"
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
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3960
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adocabeza 
         Height          =   330
         Left            =   4680
         Top             =   1440
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
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   720
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker Cmbal 
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   63504385
      CurrentDate     =   37665
   End
   Begin MSComCtl2.DTPicker Cmbdel 
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   63504385
      CurrentDate     =   37665
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   4215
      Begin VB.OptionButton Opcliquid 
         Caption         =   "Liquididacion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   165
         Width           =   1440
      End
      Begin VB.OptionButton Opcsubs 
         Caption         =   "Subsidio x enfermedad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   6735
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   5880
      TabIndex        =   12
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Inicio"
      Height          =   195
      Left            =   3120
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1125
   End
End
Attribute VB_Name = "Frmgrdsubsidio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VTipotrab As String

Private Sub Cmbal_Change()
Procesa_Cabeza_Subsidio
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Procesa_Cabeza_Subsidio
End Sub

Private Sub Cmbdel_Change()
Procesa_Cabeza_Subsidio
End Sub

Private Sub CmbTipo_Click()
If Cmbtipo.ListIndex = -1 Then
   VTipotrab = ""
Else
   VTipotrab = fc_CodigoComboBox(Cmbtipo, 2)
End If
Procesa_Cabeza_Subsidio
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 8250
Me.Height = 7350
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Cmbdel.Value = Date
Cmbal.Value = Date
End Sub
Public Sub Procesa_Cabeza_Subsidio()
Dim mano As Integer
Dim mmes As Integer
On Error GoTo CORRIGE
'If Opcsubs.Value = True Then VTipo = "05" Else VTipo = "04"
'---RODA If Opcsubs.Value = True Then VTipo = "04" Else VTipo = "05"
If Opcsubs.Value = True Then VTipo = "04" Else VTipo = "05"

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & nombre()
Sql$ = Sql$ & "a.placod,a.totaling,a.fechaproceso,user_crea " _
     & "from plahistorico a,planillas b " _
     & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and a.tipotrab LIKE '" & Trim(VTipotrab) + "%" & "' " _
     & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' " _
     & "and fechaproceso between '" & Format(Cmbdel.Value, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(Cmbal.Value, FormatFecha) & Space(1) & FormatTimef & "'"
     
cn.CursorLocation = adUseClient

Set Adocabeza.Recordset = cn.Execute(Sql$, 64)
If Adocabeza.Recordset.RecordCount > 0 Then Adocabeza.Recordset.MoveFirst
Dgrdcabeza.Refresh
Screen.MousePointer = vbDefault
Exit Sub
CORRIGE:
     MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Sub
Private Sub Opcliquid_Click()
Procesa_Cabeza_Subsidio
End Sub
Private Sub Opcsubs_Click()
Procesa_Cabeza_Subsidio
End Sub
Public Sub Elimina_Subsidio()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
Dim Vano As Integer
Dim Vmes As Integer
Dim Vdia As Integer
NroTrans = 0
Vano = Val(Mid(Dgrdcabeza.Columns(0), 7, 4))
Vmes = Val(Mid(Dgrdcabeza.Columns(0), 4, 2))
Vdia = Val(Mid(Dgrdcabeza.Columns(0), 1, 2))
Mgrab = MsgBox("Seguro de Eliminar ", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub
If Opcsubs.Value = True Then VTipo = "04" Else VTipo = "05"

cn.BeginTrans
NroTrans = 1

Sql$ = "set dateformat " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
& "where cia='" & wcia & "' and proceso='" & VTipo & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and " & wFuncdia & "(fechaproceso)=" & Vdia & " " _
& "and status<>'*' and placod='" & Dgrdcabeza.Columns(1) & "'"
cn.Execute Sql$
cn.CommitTrans
MsgBox "Eliminación Satisfactoria", vbInformation, Me.Caption
Procesa_Cabeza_Subsidio
Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox Err.Description, vbCritical, Me.Caption
Procesa_Cabeza_Subsidio
End Sub


