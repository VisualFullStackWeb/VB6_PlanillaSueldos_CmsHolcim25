VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmHoras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Seteo de Horas de Planilla «"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   Icon            =   "FrmHoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9855
      Begin VB.ListBox LstHoras 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "FrmHoras.frx":030A
         Left            =   6240
         List            =   "FrmHoras.frx":0317
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   275
         Left            =   4990
         TabIndex        =   9
         Top             =   640
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   467
         _StockProps     =   15
         Caption         =   "G"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   270
         Left            =   4500
         TabIndex        =   8
         Top             =   640
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   485
         _StockProps     =   15
         Caption         =   "V"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   270
         Left            =   4000
         TabIndex        =   7
         Top             =   640
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   485
         _StockProps     =   15
         Caption         =   "N"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   495
         Left            =   4005
         TabIndex        =   6
         Top             =   130
         Width           =   1480
         _Version        =   65536
         _ExtentX        =   2611
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "Tipo de Boleta"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid Dgrdhoras 
         Height          =   6375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   11245
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   4
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
            Caption         =   "                     Descripcion"
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
            DataField       =   "normal"
            Caption         =   "normal"
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
            DataField       =   "vaca"
            Caption         =   "vacaciones"
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
            DataField       =   "grati"
            Caption         =   "gratificacion"
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
            DataField       =   "signo"
            Caption         =   "Signo"
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
         BeginProperty Column05 
            DataField       =   "tareo"
            Caption         =   "Tareo"
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
         BeginProperty Column06 
            DataField       =   "defecto"
            Caption         =   "Por Defecto"
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
         BeginProperty Column07 
            DataField       =   "fijo"
            Caption         =   "Fijo"
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
            DataField       =   "codigo"
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
         BeginProperty Column09 
            DataField       =   "valor"
            Caption         =   "    Valor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "orden"
            Caption         =   "Orden"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   3555.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column06 
               Button          =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   645.165
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   8535
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
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
         Left            =   7560
         TabIndex        =   5
         Top             =   120
         Width           =   1335
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
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rshoras As New Recordset

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Crea_Rs
End Sub

Private Sub DgrdHoras_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 1, 2, 3, 5, 7
            If Trim(UCase(Dgrdhoras.Columns(ColIndex))) <> "S" And Trim(Dgrdhoras.Columns(ColIndex)) <> "" Then
               MsgBox "Solo Puede ser [S]i", vbCritical, TitMsg
               Dgrdhoras.Columns(ColIndex) = ""
            End If
       Case Is = 4
            If Trim(Dgrdhoras.Columns(ColIndex)) <> "-" And Trim(Dgrdhoras.Columns(ColIndex)) <> "" Then
               MsgBox "Solo Puede ser [-]", vbCritical, TitMsg
               Dgrdhoras.Columns(ColIndex) = ""
            End If
End Select
End Sub

Private Sub Dgrdhoras_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Dgrdhoras.Col = 6 Then
        KeyAscii = 0
        Cancel = True
        Dgrdhoras_ButtonClick (ColIndex)
End If
End Sub

Private Sub Dgrdhoras_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer
Y = Dgrdhoras.Row
xtop = Dgrdhoras.Top + Dgrdhoras.RowTop(Y) + Dgrdhoras.RowHeight
Select Case ColIndex
Case 6:
       xleft = Dgrdhoras.Left + Dgrdhoras.Columns(6).Left
       With LstHoras
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = Dgrdhoras.Top + Dgrdhoras.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7620
Me.Width = 10035
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
LblFecha.Caption = Format(Date, "dd/mm/yyyy")
Procesa_Horas
End Sub
Private Sub Crea_Rs()
    If rshoras.State = 1 Then rshoras.Close
    rshoras.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rshoras.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rshoras.Fields.Append "normal", adChar, 1, adFldIsNullable
    rshoras.Fields.Append "vaca", adChar, 1, adFldIsNullable
    rshoras.Fields.Append "grati", adChar, 1, adFldIsNullable
    rshoras.Fields.Append "signo", adChar, 1, adFldIsNullable
    rshoras.Fields.Append "tareo", adChar, 1, adFldIsNullable
    rshoras.Fields.Append "defecto", adChar, 10, adFldIsNullable
    rshoras.Fields.Append "fijo", adChar, 1, adFldIsNullable
    rshoras.Fields.Append "valor", adDecimal, 20, adFldIsNullable
    rshoras.Fields.Append "orden", adVarChar, 2, adFldIsNullable
    rshoras.Open
    Set Dgrdhoras.DataSource = rshoras
End Sub
Private Sub Procesa_Horas()
Dim wciamae As String
wciamae = Determina_Maestro("01077")
Sql$ = "Select * from maestros_2 where status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rshoras.RecordCount > 0 Then
   rshoras.MoveFirst
   Do While Not rshoras.EOF
      rshoras.Delete
      rshoras.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rshoras.AddNew
   rshoras!codigo = Trim(rs!COD_MAESTRO2)
   rshoras!Descripcion = rs!DESCRIP
   If InStr(1, Trim(rs!flag1), "N") > 0 Then rshoras!Normal = "S" Else rshoras!Normal = ""
   If InStr(1, Trim(rs!flag1), "V") > 0 Then rshoras!vaca = "S" Else rshoras!vaca = ""
   If InStr(1, Trim(rs!flag1), "G") > 0 Then rshoras!grati = "S" Else rshoras!grati = ""
   rshoras!signo = Trim(rs!flag2)
   rshoras!ORDEN = Trim(rs!flag7 & "")
   If Left(rs!flag3, 2) = "DI" Then rshoras!defecto = "DIAS"
   If Left(rs!flag3, 2) = "HO" Then rshoras!defecto = "HORAS"
   If Left(rs!flag3, 2) = "MI" Then rshoras!defecto = "MINUTOS"
   
   'If Trim(RS!FLAG4) <> "" Then
   rshoras!tareo = Trim(rs!FLAG4)
   'End If
   
   rshoras!fijo = Trim(rs!flag5)
   If Not IsNull(rs!flag6) Then
      If Val(rs!flag6) > 0 Then rshoras!VALOR = Val(rs!flag6)
   End If
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
End Sub

Private Sub LstHoras_Click()
If LstHoras.ListIndex > -1 Then
    Dgrdhoras.Columns(6) = Trim(LstHoras.Text)
    Dgrdhoras.Col = 6
    Dgrdhoras.SetFocus
    LstHoras.ZOrder 1
    LstHoras.Visible = False
End If
End Sub
Private Sub LstHoras_LostFocus()
LstHoras.Visible = False
End Sub
Public Function GrabarHorasPla()
Dim mm As String
Dim mTipo As String
Dim mvalor As String
On Error GoTo ErrorTrans
Dim NroTrans As Integer
NroTrans = 0
Mgrab = MsgBox("Seguro de Grabar Seteo de Horas de Planilla", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Function
Screen.MousePointer = vbArrowHourglass
cn.BeginTrans
NroTrans = 1
If rshoras.RecordCount > 0 Then rshoras.MoveFirst

Dim xciamae As String
Dim cod As String
cod = "01077"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Sql$ = "SELECT GENERAL FROM maestros where right(ciamaestro,3)='" & Right(cod, 3) & "' and status<>'*' "
If (fAbrRst(rs, Sql$)) Then
   If rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If
If rs.State = 1 Then rs.Close

Do While Not rshoras.EOF
   If IsNull(rshoras!defecto) Then
      mm = ""
   Else
      mm = Left(rshoras!defecto, 2)
   End If
   mTipo = ""
   If rshoras!Normal = "S" Then mTipo = "N"
   If rshoras!vaca = "S" Then mTipo = mTipo & "V"
   If rshoras!grati = "S" Then mTipo = mTipo & "G"
   If IsNull(rshoras!VALOR) Then mvalor = "" Else mvalor = Str(rshoras!VALOR)
   Sql$ = "update maestros_2 set flag1='" & mTipo & "',flag2='" & rshoras!signo & "',flag3='" & mm & "',flag4='" & rshoras!tareo & "',flag5='" & rshoras!fijo & "',flag6='" & mvalor & "',flag7='" & Trim(rshoras!ORDEN & "") & "' where cod_maestro2='" & rshoras!codigo & "' and status<>'*'"
   Sql$ = Sql$ & xciamae
   cn.Execute Sql$
   rshoras.MoveNext
Loop

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

End Function

