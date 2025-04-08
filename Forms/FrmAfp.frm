VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmAfp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» AFP «"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   Icon            =   "FrmAfp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Constantes de Calculo"
      TabPicture(0)   =   "FrmAfp.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtTopeMax"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Cmbmes"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Txtano"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Persona Responsable"
      TabPicture(1)   =   "FrmAfp.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.TextBox Txtano 
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   7800
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox Cmbmes 
         Height          =   315
         ItemData        =   "FrmAfp.frx":0342
         Left            =   4920
         List            =   "FrmAfp.frx":036A
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   480
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Height          =   3135
         Left            =   -74640
         TabIndex        =   9
         Top             =   720
         Width           =   8295
         Begin VB.ComboBox Cmbcta 
            Height          =   315
            ItemData        =   "FrmAfp.frx":03D2
            Left            =   3600
            List            =   "FrmAfp.frx":03D4
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox TxtCtaAfp 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6000
            TabIndex        =   19
            Text            =   "CUENTA CORRIENTE"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox TxtTlfAfp 
            Height          =   285
            Left            =   3480
            TabIndex        =   17
            Top             =   2280
            Width           =   1335
         End
         Begin VB.ComboBox CmbAreaAfp 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6000
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox TxtRespAfp 
            Height          =   285
            Left            =   600
            TabIndex        =   13
            Top             =   1680
            Width           =   4935
         End
         Begin VB.ComboBox CmbBcoAfp 
            Height          =   315
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No de Cuenta"
            Height          =   195
            Left            =   3600
            TabIndex        =   18
            Top             =   480
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono de Persona Responsable"
            Height          =   195
            Left            =   600
            TabIndex        =   16
            Top             =   2280
            Width           =   2460
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
            Height          =   195
            Left            =   6000
            TabIndex        =   14
            Top             =   1440
            Width           =   330
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsable de AFP"
            Height          =   195
            Left            =   600
            TabIndex        =   12
            Top             =   1440
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Institucion Finaciera"
            Height          =   195
            Left            =   480
            TabIndex        =   10
            Top             =   480
            Width           =   1410
         End
      End
      Begin VB.TextBox TxtTopeMax 
         Alignment       =   1  'Right Justify
         Height          =   305
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   10935
         Begin MSDataGridLib.DataGrid DgrdAfp 
            Height          =   3015
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   5318
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "descripcion"
               Caption         =   "Descripcion"
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
               DataField       =   "afp01"
               Caption         =   "Apor. Oblig."
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
               DataField       =   "afp02"
               Caption         =   "seg. Inval."
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
               DataField       =   "afp03"
               Caption         =   "Cont. Ipss"
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
            BeginProperty Column04 
               DataField       =   "afp04"
               Caption         =   "Com. Flujo"
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
            BeginProperty Column05 
               DataField       =   "fijo"
               Caption         =   "Fijo"
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
            BeginProperty Column06 
               DataField       =   "tope"
               Caption         =   "tope"
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
            BeginProperty Column07 
               DataField       =   "codafp"
               Caption         =   "codafp"
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
            BeginProperty Column08 
               DataField       =   "afp05"
               Caption         =   "Com. Mixta"
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
            BeginProperty Column09 
               DataField       =   "TOTAL"
               Caption         =   "Total Flujo"
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
            BeginProperty Column10 
               DataField       =   "TotalMixta"
               Caption         =   "Total Mixta"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   3089.764
               EndProperty
               BeginProperty Column01 
                  Alignment       =   1
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   840.189
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  Object.Visible         =   0   'False
                  ColumnWidth     =   720
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column07 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column08 
                  Alignment       =   1
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column09 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column10 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
         Height          =   195
         Left            =   7320
         TabIndex        =   23
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
         Height          =   195
         Left            =   4440
         TabIndex        =   21
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tope Maximo"
         Height          =   195
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6975
         TabIndex        =   4
         Top             =   150
         Width           =   1815
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
         TabIndex        =   2
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmAfp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsafp As New Recordset
Dim Vmes As String
Dim VBcoAfp As String
Dim VAreaAfp As String

Private Sub CmbAreaAfp_Click()
VAreaAfp = fc_CodigoComboBox(CmbAreaAfp, 2)
End Sub

Private Sub CmbBcoAfp_Click()
VBcoAfp = fc_CodigoComboBox(CmbBcoAfp, 2)
Call rCarCbo(Cmbcta, "select moneda,cuenta,cheque,bcoasiento from bancocta where cia='" & wcia & "' and banco='" & VBcoAfp & "' and status<>'*'", "C", "00")
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Cmbmes.ListIndex = Val(Month(Date)) - 1
Txtano.Text = Right(Lblfecha.Caption, 4)
Call fc_Descrip_Maestros2(wcia & "007", "", CmbBcoAfp)
Call fc_Descrip_Maestros2("01044", "", CmbAreaAfp)
Procesa
End Sub

Private Sub Cmbmes_Click()
Vmes = Format(Cmbmes.ListIndex + 1, "00")
Procesa

End Sub

Private Sub DgrdAfp_AfterColEdit(ByVal ColIndex As Integer)
If rsafp.RecordCount > 0 Then
      rsafp("TOTAL") = rsafp("AFP01") + rsafp("AFP02") + rsafp("AFP03") + rsafp("AFP04")
      rsafp("TotalMixta") = rsafp("AFP01") + rsafp("AFP02") + rsafp("AFP03") + rsafp("AFP05")
End If
End Sub

Private Sub DgrdAfp_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
    rsafp("TOTAL") = rsafp("AFP01") + rsafp("AFP02") + rsafp("AFP03") + rsafp("AFP04")
    rsafp("TotalMixta") = rsafp("AFP01") + rsafp("AFP02") + rsafp("AFP03") + rsafp("AFP05")
 End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 5460
Me.Width = 11355
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (rs Is Nothing) Then
    If rs.State = 1 Then rs.Close
End If
Set DgrdAfp.DataSource = Nothing
If rsafp.State = 1 Then rsafp.Close
Set rsafp = Nothing
End Sub
Private Sub Txtano_Change()
Procesa
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Crea_Rs()
    If rsafp.State = 1 Then rsafp.Close
    rsafp.Fields.Append "codafp", adChar, 2, adFldIsNullable
    rsafp.Fields.Append "descripcion", adChar, 90, adFldIsNullable
    rsafp.Fields.Append "afp01", adCurrency, 18, adFldIsNullable
    rsafp.Fields.Append "afp02", adCurrency, 18, adFldIsNullable
    rsafp.Fields.Append "afp03", adCurrency, 18, adFldIsNullable
    rsafp.Fields.Append "afp04", adCurrency, 18, adFldIsNullable
    rsafp.Fields.Append "afp05", adCurrency, 18, adFldIsNullable
    rsafp.Fields.Append "tope", adCurrency, 18, adFldIsNullable
    rsafp.Fields.Append "periodo", adChar, 4, adFldIsNullable
    rsafp.Fields.Append "TOTAL", adCurrency, 18, adFldIsNullable
    rsafp.Fields.Append "TotalMixta", adCurrency, 18, adFldIsNullable
    rsafp.Open
    Set DgrdAfp.DataSource = rsafp
End Sub
Private Sub Procesa()
Dim wciamae As String
Dim mperiodo As String
Dim rs2 As ADODB.Recordset
Dim I As Integer

mperiodo = Txtano.Text & Vmes
TxtTopeMax.Text = "0.00"
wciamae = Determina_Maestro("01069")
Sql$ = "Select * from maestros_2 where status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsafp.RecordCount > 0 Then
   rsafp.MoveFirst
   Do While Not rsafp.EOF
      rsafp.Delete
      rsafp.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then MsgBox "No Existen AFP's Registradas", vbCritical, TitMsg: Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsafp.AddNew
   rsafp!CodAfp = Trim(rs!COD_MAESTRO2)
   rsafp!Descripcion = Trim(rs!DESCRIP)
   
   Sql$ = "Select * from plaafp where cia='" & wcia & "' and codafp='" & rs!COD_MAESTRO2 & "' and status<>'*' and periodo='" & mperiodo & "'"
   cn.CursorLocation = adUseClient
   Set rs2 = New ADODB.Recordset
   Set rs2 = cn.Execute(Sql$, 64)
   If rs2.RecordCount <= 0 Then
      If rs2.State = 1 Then rs2.Close
      If Vmes = "01" Then mperiodo = Format(Val(Txtano.Text) - 1, "0000") & "12" Else: mperiodo = Txtano.Text & Format(Val(Vmes) - 1, "00")
      Sql$ = "Select * from plaafp where cia='" & wcia & "' and codafp='" & Trim(rs!COD_MAESTRO2) & "' and status<>'*' and periodo='" & mperiodo & "'"
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
   End If
   If rs2.RecordCount > 0 Then
      TxtTopeMax.Text = Format(rs2!tope, "###,###.00")
      rsafp!afp01 = rs2!afp01
      rsafp!afp02 = rs2!afp02
      rsafp!afp03 = rs2!afp03
      rsafp!afp04 = rs2!afp04
      rsafp!AFP05 = rs2!AFP05
   Else
      rsafp!afp01 = "0.00"
      rsafp!afp02 = "0.00"
      rsafp!afp03 = "0.00"
      rsafp!afp04 = "0.00"
      rsafp!AFP05 = "0.00"
   End If
   If rs2.State = 1 Then rs2.Close
   rs.MoveNext
   Call DgrdAfp_KeyUp(13, 10)
Loop

If rs.State = 1 Then rs.Close

Sql$ = "Select * from cia where cod_cia='" & wcia & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then
   TxtRespAfp.Text = rs!afpresponsable
   TxtTlfAfp.Text = rs!afpresptlf
   Call rUbiIndCmbBox(CmbBcoAfp, rs!afpbanco, "00")
   Call rUbiIndCmbBox(CmbAreaAfp, rs!afparearesp, "00")
   For I = 0 To Cmbcta.ListCount - 1
       If Left(Cmbcta.List(I), 15) = Left(rs!afpnrocta, 15) Then Cmbcta.ListIndex = I: Exit For
   Next
End If

End Sub
Public Function GrabarAfp()
Dim mperiodo As String
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
Mgrab = MsgBox("Seguro de Grabar Seteo de AFP", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
mperiodo = Txtano.Text & Vmes
If TxtTopeMax.Text = "" Then MsgBox "Debe Registrar Tope Maximo", vbCritical, TitMsg: Exit Function
If CCur(TxtTopeMax.Text) <= 0 Then MsgBox "Debe Registrar Tope Maximo", vbCritical, TitMsg: Exit Function
If CCur(Txtano.Text) <= 0 Then MsgBox "Debe Registrar Año", vbCritical, TitMsg: Exit Function

Screen.MousePointer = vbArrowHourglass
cn.BeginTrans
NroTrans = 1
   
Dim Rq As ADODB.Recordset
Sql = "select cod_cia from cia where status<>'*'"
If fAbrRst(Rq, Sql) Then Rq.MoveFirst
Do While Not Rq.EOF
   Sql$ = "Update plaafp set status='*' where cia='" & Rq!cod_cia & "' and periodo='" & mperiodo & "' and status<>'*'"
   cn.Execute Sql$
   If rsafp.RecordCount > 0 Then rsafp.MoveFirst
   Do While Not rsafp.EOF
      Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      Sql$ = Sql$ & "INSERT INTO plaafp values('" & Rq!cod_cia & "','" & rsafp!CodAfp & "','" & Trim(rsafp!Descripcion) & "'," & CCur(rsafp!afp01) & ", " _
           & "" & CCur(rsafp!afp02) & "," & CCur(rsafp!afp03) & "," & CCur(rsafp!afp04) & "," & CCur(rsafp!AFP05) & "," & CCur(TxtTopeMax.Text) & ", " _
           & "''," & FechaSys & ",'" & mperiodo & "')"
      cn.Execute Sql$
      rsafp.MoveNext
   Loop
   Rq.MoveNext
Loop
Rq.Close: Set Rq = Nothing

cn.CommitTrans
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Screen.MousePointer = vbDefault
Exit Function
ErrorTrans:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault
'Exit Sub
'FUNKA:
 '  MsgBox "Error : " & Err.Description, vbCritical, "Planillas"
End Function

Private Sub TxtTopeMax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Cmbmes.SetFocus
End Sub
Private Sub TxtTopeMax_LostFocus()
TxtTopeMax.Text = Format(TxtTopeMax.Text, "###,###.00")
End Sub
