VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmDeduccion 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Deducciones y Aportaciones «"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "Frmdeduccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4110
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   8970
      Begin VB.PictureBox pct_conceptos 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2700
         Left            =   5085
         ScaleHeight     =   2670
         ScaleWidth      =   2715
         TabIndex        =   14
         Top             =   540
         Visible         =   0   'False
         Width           =   2745
         Begin VB.CommandButton cmdsalir 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   16
            Top             =   0
            Width           =   330
         End
         Begin VB.ListBox lstconceptos 
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
            Height          =   2280
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   15
            Top             =   315
            Width           =   2625
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Caption         =   "Conceptos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   2670
         End
      End
      Begin VB.PictureBox pctempresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   2190
         ScaleHeight     =   795
         ScaleWidth      =   4305
         TabIndex        =   8
         Top             =   1335
         Visible         =   0   'False
         Width           =   4335
         Begin VB.ComboBox cbocia 
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
            Left            =   45
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   405
            Width           =   3840
         End
         Begin Threed.SSCommand ssc_importa 
            Height          =   330
            Left            =   3915
            TabIndex        =   11
            Top             =   405
            Width           =   330
            _Version        =   65536
            _ExtentX        =   582
            _ExtentY        =   582
            _StockProps     =   78
            BevelWidth      =   1
            AutoSize        =   2
            Picture         =   "Frmdeduccion.frx":030A
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Selecione Cia. de Donde Importar Información"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   195
            TabIndex        =   10
            Top             =   45
            Width           =   3885
         End
      End
      Begin VB.CheckBox CheckPersona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3840
         TabIndex        =   6
         Top             =   345
         Visible         =   0   'False
         Width           =   225
      End
      Begin MSDataGridLib.DataGrid DgrdConcepto 
         Height          =   3375
         Left            =   120
         TabIndex        =   4
         Top             =   90
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         ColumnCount     =   9
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
            DataField       =   "personal"
            Caption         =   "Personal"
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
            DataField       =   "deduccion"
            Caption         =   "Deduccion"
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
            DataField       =   "aportacion"
            Caption         =   "Aportacion"
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
            DataField       =   "status"
            Caption         =   "Status"
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
         BeginProperty Column06 
            DataField       =   "adicional"
            Caption         =   "Manual"
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
            DataField       =   "codconcepto"
            Caption         =   "codconcepto"
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
            DataField       =   "codsunat"
            Caption         =   "CodSunat"
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
               Locked          =   -1  'True
               ColumnWidth     =   3105.071
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   524.976
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1035.213
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   7080
         Picture         =   "Frmdeduccion.frx":08A4
         Top             =   3600
         Width           =   240
      End
      Begin VB.Label lblconceptos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione Conceptos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   525
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Importar Informacion de Otra Cia."
         Top             =   3495
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importar Movimientos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   525
         Left            =   7020
         TabIndex        =   12
         ToolTipText     =   "Importar Informacion de Otra Cia."
         Top             =   3495
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status: [P]orcentaje  [F]ijo         Adicional: [S]i  [N]o"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   525
         Left            =   1965
         TabIndex        =   7
         Top             =   3495
         Width           =   5025
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8985
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   60
         Width           =   4455
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
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
         Left            =   7275
         TabIndex        =   5
         Top             =   120
         Width           =   1575
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
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmDeduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VConcepto As String
Dim rsconcepto As New Recordset
Dim str_cptocargados As String

Private Sub CheckPersona_Click()
If CheckPersona.Value = 1 Then
   rsconcepto!personal = "S"
   Dgrdconcepto.Columns(2) = "0.00"
   Dgrdconcepto.Columns(3) = "0.00"
Else
   rsconcepto!personal = ""
End If
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
VConcepto = "03"
Procesa
End Sub

Private Sub CmbTipo_Click()
VConcepto = Left(Cmbtipo.Text, 2)
Procesa
End Sub

Private Sub cmdsalir_Click()
Dim I As Long

pct_conceptos.Visible = False

For I = 0 To Me.LstConceptos.ListCount - 1
    If LstConceptos.Selected(I) Then
        rsconcepto.AddNew
        rsconcepto!Descripcion = LstConceptos.List(I)
        rsconcepto!personal = ""
        rsconcepto!aportacion = 0
        rsconcepto!deduccion = 0
        rsconcepto!status = " "
        rsconcepto!codigo = ""
        rsconcepto!adicional = ""
        rsconcepto!codconcepto = Format(LstConceptos.ItemData(I), "000")
        str_cptocargados = str_cptocargados & "'" & Format(LstConceptos.ItemData(I), "000") & "',"
    End If
Next I

End Sub

Private Sub DgrdConcepto_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 4
            If UCase(Dgrdconcepto.Columns(4)) <> "P" And Dgrdconcepto.Columns(4) <> "F" Then
               MsgBox "Status Solo Puede ser [P]orcentaje o [F]ijo", vbCritical, TitMsg
               Dgrdconcepto.Columns(4) = ""
            End If
       Case Is = 6
            If UCase(Trim(Dgrdconcepto.Columns(6))) <> "S" And Trim(Dgrdconcepto.Columns(6)) <> "N" Then
               MsgBox "Adicional Solo Puede ser [S]i o [N]o", vbCritical, TitMsg
               Dgrdconcepto.Columns(6) = ""
            End If
       Case Is = 2
            If UCase(Trim(Dgrdconcepto.Columns(1))) <> "S" Then
               If CCur(Dgrdconcepto.Columns(2)) <> 0 Then Dgrdconcepto.Columns(3) = "0.00"
            Else
               'Dgrdconcepto.Columns(2) = "0.00"
               'Dgrdconcepto.Columns(3) = "0.00"
            End If
       Case Is = 3
            If UCase(Trim(Dgrdconcepto.Columns(1))) <> "S" Then
               If CCur(Dgrdconcepto.Columns(3)) <> 0 Then Dgrdconcepto.Columns(2) = "0.00"
            Else
               'Dgrdconcepto.Columns(2) = "0.00"
               'Dgrdconcepto.Columns(3) = "0.00"
            End If
End Select
End Sub

Private Sub DgrdConcepto_OnAddNew()
rsconcepto.AddNew
rsconcepto!deduccion = 0
rsconcepto!aportacion = 0
End Sub

Private Sub DgrdConcepto_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rsconcepto.AbsolutePage > 0 Then
   If UCase(rsconcepto!personal) = "S" Then CheckPersona.Value = 1 Else CheckPersona.Value = 0
   CheckPersona.Left = Dgrdconcepto.Left + Dgrdconcepto.Columns(1).Left + 250
   CheckPersona.Top = Dgrdconcepto.Top + Dgrdconcepto.RowTop(Dgrdconcepto.Row) + 5
   CheckPersona.Visible = True
   CheckPersona.ZOrder 0
End If
End Sub

Private Sub DgrdConcepto_Scroll(Cancel As Integer)
CheckPersona.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If pctempresa.Visible Then
        pctempresa.Visible = False
    End If
End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 4965
    Me.Width = 9075
    LblFecha.Caption = Format(Date, "dd/mm/yyyy")
    Crea_Rs
    Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
    Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Private Sub Crea_Rs()
    If rsconcepto.State = 1 Then rsconcepto.Close
    rsconcepto.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsconcepto.Fields.Append "personal", adChar, 1, adFldIsNullable
    rsconcepto.Fields.Append "deduccion", adCurrency, 18, adFldIsNullable
    rsconcepto.Fields.Append "aportacion", adCurrency, 18, adFldIsNullable
    rsconcepto.Fields.Append "status", adChar, 1, adFldIsNullable
    rsconcepto.Fields.Append "adicional", adChar, 1, adFldIsNullable
    rsconcepto.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsconcepto.Fields.Append "codconcepto", adChar, 3, adFldIsNullable
    rsconcepto.Fields.Append "semana", adChar, 3, adFldIsNullable
    rsconcepto.Fields.Append "codsunat", adChar, 4, adFldIsNullable
    rsconcepto.Open
    Set Dgrdconcepto.DataSource = rsconcepto
End Sub
Private Sub Procesa()
str_cptocargados = ""
Sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='" & VConcepto & "' and status<>'*' order by codinterno"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsconcepto.RecordCount > 0 Then
   rsconcepto.MoveFirst
   Do While Not rsconcepto.EOF
      rsconcepto.Delete
      rsconcepto.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF

   rsconcepto.AddNew
   rsconcepto!Descripcion = rs!Descripcion
   rsconcepto!personal = rs!personal
   rsconcepto!aportacion = rs!aportacion
   rsconcepto!deduccion = rs!deduccion
   rsconcepto!status = rs!status
   rsconcepto!codigo = rs!codinterno
   rsconcepto!adicional = rs!adicional
   rsconcepto!codconcepto = rs!cod_concepto
   rsconcepto!semana = rs!semana_pago
   rsconcepto!CODSUNAT = Trim(rs!CODSUNAT & "")
   
   str_cptocargados = str_cptocargados & "'" & rs!cod_concepto & "',"
   
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close
End Sub
Public Sub Graba_Concepto()
On Error GoTo Grabacion
Dim NroError As Integer
NroError = 0

With rsconcepto
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                If Trim(!CODSUNAT & "") = "" Then
                    MsgBox "Ingrese Código de Sunat, dato obligatorio", vbExclamation, Me.Caption
                    Me.Dgrdconcepto.Col = 8
                    Me.Dgrdconcepto.SetFocus
                    Exit Sub
                ElseIf Len(Trim(!CODSUNAT & "")) < 4 Or Len(Trim(!CODSUNAT & "")) > 4 Then
                    MsgBox "Ingrese Código de Sunat (4 digiros), dato obligatorio", vbExclamation, Me.Caption
                    Me.Dgrdconcepto.Col = 8
                    Me.Dgrdconcepto.SetFocus
                    Exit Sub
                End If
                
            .MoveNext
        Loop
    End If
End With
Dim Mcodigo As String
If MsgBox("Desea Grabar Conceptos", vbYesNo + vbQuestion) = vbNo Then Screen.MousePointer = vbDefault: Exit Sub

 

cn.BeginTrans
NroError = 1

   Dim Rq As ADODB.Recordset
   Sql = "select cod_cia from cia where status<>'*'"
   If fAbrRst(Rq, Sql) Then Rq.MoveFirst
   Do While Not Rq.EOF
      Sql = "update placonstante set status='*' where cia='" & Rq!cod_cia & "' and tipomovimiento='" & VConcepto & "' and status<>'*'"
      cn.Execute Sql
      If rsconcepto.RecordCount > 0 Then
         rsconcepto.MoveFirst
         Do While Not rsconcepto.EOF
            If IsNull(rsconcepto!codigo) Then rsconcepto!codigo = ""
            If Val(rsconcepto!codigo) <> 0 Then
               Mcodigo = Trim(rsconcepto!codigo)
            Else
               Sql = "select max(codinterno) as codigo from placonstante where cia='" & Rq!cod_cia & "' and tipomovimiento='" & VConcepto & "'"
               cn.CursorLocation = adUseClient
               Set rs = New ADODB.Recordset
               Set rs = cn.Execute(Sql, 64)
               If rs.RecordCount > 0 And Not IsNull(rs!codigo) Then Mcodigo = Format(Val(rs!codigo) + 1, "00") Else Mcodigo = "01"
               If rs.State = 1 Then rs.Close
            End If
      
            If Trim(rsconcepto!Descripcion) <> "" Then
               Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
               Sql = Sql & "INSERT INTO placonstante values('" & Rq!cod_cia & "','" & VConcepto & "','" & Mcodigo & "','" & Trim(rsconcepto!Descripcion) & "'," & CCur(rsconcepto!deduccion) & "," & CCur(rsconcepto!aportacion) & ",'" & rsconcepto!personal & "','" & rsconcepto!status & "','','" & rsconcepto!adicional & "'," & FechaSys & ",'','','','" & rsconcepto!codconcepto & "'," & IIf(IsNull(rsconcepto!semana), "NULL", rsconcepto!semana) & ",'" & Trim(rsconcepto!CODSUNAT) & "','','')"
       
               cn.Execute Sql
            End If
            rsconcepto.MoveNext
         Loop
      End If
      Rq.MoveNext
   Loop
cn.CommitTrans

MsgBox "Se guardaron los datos correctamente", vbInformation, Me.Caption

Screen.MousePointer = vbDefault
Procesa

Exit Sub
Grabacion:
If NroError = 1 Then
    cn.RollbackTrans
End If

MsgBox "Error " & ERR.Description, vbCritical, TitMsg
Screen.MousePointer = vbDefault
End Sub

Private Sub Label4_Click()
Call rCarCbo(cbocia, Carga_Cia, "C", "00")
pctempresa.Visible = True
cbocia.ListIndex = 0
End Sub

Private Sub lblconceptos_Click()
Dim sSQL As String
Dim rs As Object

If Len(Trim(str_cptocargados)) > 0 Then
    sSQL = "select cod_concepto,desc_concepto from tconceptos where cod_concepto not in (" & Mid(str_cptocargados, 1, Len(Trim(str_cptocargados)) - 1) & ") AND status!='*'"
Else
    sSQL = "select cod_concepto,desc_concepto from tconceptos where status!='*'"
End If

Set rs = cn.Execute(sSQL)

LstConceptos.Clear

If Not rs.EOF Then
    Do While Not rs.EOF
        With LstConceptos
            .AddItem rs!desc_concepto
            .ItemData(.NewIndex) = rs!cod_concepto
        End With
        rs.MoveNext
    Loop
    rs.Close
End If

Set rs = Nothing

pct_conceptos.Visible = True
End Sub

Private Sub ssc_importa_Click()
Dim sSQL As String
Dim NroError As Integer
On Error GoTo Importar
NroError = 0
cn.BeginTrans
NroError = 1
sSQL$ = "EXEC sp_i_importa_data '" & VConcepto & "','" & Format(cbocia.ItemData(cbocia.ListIndex), "00") & "','" & wcia & "'"

cn.Execute sSQL

cn.CommitTrans

MsgBox "Se Realizo la Importación Correctamente", vbQuestion, TitMsg

pctempresa.Visible = False
Procesa
Exit Sub

Importar:
    If NroError = 1 Then
        cn.RollbackTrans
    End If
    MsgBox "Error " & ERR.Description, vbCritical, TitMsg

End Sub
