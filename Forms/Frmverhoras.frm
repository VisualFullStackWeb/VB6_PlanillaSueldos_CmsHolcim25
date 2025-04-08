VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form Frmverhoras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seteo de Horas por Compañia"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "Frmverhoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   3825
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6495
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      Begin vbAcceleratorSGrid6.vbalGrid vbgcts 
         Height          =   6225
         Left            =   90
         TabIndex        =   7
         Top             =   135
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   10980
         NoVerticalGridLines=   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   14737632
         GridLineColor   =   15466236
         HighlightBackColor=   15466236
         HighlightForeColor=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HeaderButtons   =   0   'False
         HeaderFlat      =   -1  'True
         BorderStyle     =   2
         DisableIcons    =   -1  'True
         HighlightSelectedIcons=   0   'False
      End
      Begin vbalIml6.vbalImageList vbalImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   14924
         Images          =   "Frmverhoras.frx":030A
         Version         =   131072
         KeyCount        =   13
         Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿ"
      End
      Begin MSDataGridLib.DataGrid DgrdHoras 
         Height          =   6255
         Left            =   75
         TabIndex        =   4
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   11033
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "concepto"
            Caption         =   "Concepto"
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
            DataField       =   "ver"
            Caption         =   "Ver"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   4350.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   285.165
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
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
      Width           =   5475
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   80
         Width           =   4290
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
         Top             =   75
         Width           =   825
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1350
   End
End
Attribute VB_Name = "Frmverhoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rshoras As New Recordset
Dim wciamae As String
Dim VTipo As String

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
'Crea_Rs
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
procesa_VerHoras
End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
procesa_VerHoras
End Sub

'Private Sub DgrdHoras_AfterColEdit(ByVal ColIndex As Integer)
'Select Case ColIndex
'       Case Is = 1
'       If rshoras!ver = "s" Then rshoras!ver = "S"
'       If rshoras!ver <> "S" Then
'          MsgBox "Solo se puede Ingresar [S]", vbInformation, "Horas"
'          rshoras!ver = ""
'       End If
'End Select
'End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5595
Me.Height = 7875
InicializaGrilla
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
'Private Sub Crea_Rs()
'    If rshoras.State = 1 Then rshoras.Close
'    rshoras.Fields.Append "codigo", adChar, 2, adFldIsNullable
'    rshoras.Fields.Append "concepto", adChar, 45, adFldIsNullable
'    rshoras.Fields.Append "ver", adChar, 1, adFldIsNullable
'    rshoras.Open
'    rshoras.AddNew
'    Set DgrdHoras.DataSource = rshoras
'End Sub
Private Sub procesa_VerHoras()
Dim rs2 As ADODB.Recordset
'If rshoras.RecordCount > 0 Then rshoras.MoveFirst
'Do While Not rshoras.EOF
'   rshoras.Delete
'   rshoras.MoveNext
'Loop
If Cmbtipo.ListIndex < 0 Then Exit Sub

wciamae = Determina_Maestro("01077")
'SQL$ = "Select cod_maestro2,descrip from maestros_2 where status<>'*'"
'SQL$ = SQL$ & wciamae

Sql$ = "SELECT m.cod_maestro2,m.descrip,CASE WHEN NOT  ph.codigo IS NULL THEN ph.codigo ELSE '' END as codigo" & _
    " FROM maestros_2 m LEFT OUTER JOIN plaverhoras ph ON (ph.codigo=m.cod_maestro2 AND ph.cia='" & wcia & "' " & _
    " and ph.status<>'*' and ph.tipo_trab='" & VTipo & "') WHERE m.status<>'*' " & wciamae

'cn.CursorLocation = adUseClient
'Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)

vbgcts.Redraw = False
vbgcts.Clear
If Not rs.EOF Then

    Do Until rs.EOF
        vbgcts.AddRow
        vbgcts.CellDetails vbgcts.Rows, 1, rs!DESCRIP
        vbgcts.CellDetails vbgcts.Rows, 2, , DT_CENTER, IIf(Len(Trim(rs!codigo)) = 0, iCHCKINAC, iCHCKACT), , , , 10
        vbgcts.CellDetails vbgcts.Rows, 3, rs!COD_MAESTRO2
        
        rs.MoveNext
    Loop
rs.Close
End If
vbgcts.Redraw = True

'Do Until rs.EOF
'   rshoras.AddNew
'   rshoras!concepto = Trim(rs!descrip)
'   rshoras!codigo = Trim(rs!cod_maestro2)
'   SQL$ = "select * from plaverhoras where cia='" & wcia & "' and status<>'*' and tipo_trab='" & VTipo & "' and codigo='" & rs!cod_maestro2 & "'"
'   If (fAbrRst(rs2, SQL$)) Then rshoras!ver = "S" Else rshoras!ver = ""
'   If rs2.State = 1 Then rs2.Close
'   rs.MoveNext
'Loop

Set rs = Nothing
End Sub
Public Sub Grabar_VerHoras()
Dim iFila As Long
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
If Trim(Cmbtipo.Text) = "" Then MsgBox "Debe Seleccionar Tipo de Trabajador", vbInformation, "Tareo": Cmbtipo.SetFocus: Exit Sub

If MsgBox("Desea Grabar", vbYesNo + vbQuestion) = vbNo Then Exit Sub

cn.BeginTrans
NroTrans = 1

'rshoras.MoveFirst
Screen.MousePointer = vbHourglass

If wGrupoPla = "01" Then
   Dim Rq As ADODB.Recordset
   Sql = "select cod_cia from cia where status<>'*' and prefijo<>'' and not prefijo is null"
   If fAbrRst(Rq, Sql) Then Rq.MoveFirst
   Do While Not Rq.EOF
      Sql$ = "delete from plaverhoras where cia='" & Rq!cod_cia & "' and tipo_trab='" & VTipo & "'"
      cn.Execute Sql
      For iFila = 1 To vbgcts.Rows
         If vbgcts.CellIcon(iFila, 2) = iCHCKACT Then
            Sql$ = "INSERT INTO plaverhoras values('" & Rq!cod_cia & "','" & VTipo & "','" & Trim(vbgcts.CellText(iFila, 3)) & "','','" & wuser & "'," & FechaSys & ")"
            cn.Execute Sql$
         End If
      Next
      Rq.MoveNext
   Loop
Else
   Sql$ = "delete from plaverhoras where cia='" & wcia & "' and tipo_trab='" & VTipo & "'"
   cn.Execute Sql
   For iFila = 1 To vbgcts.Rows
      If vbgcts.CellIcon(iFila, 2) = iCHCKACT Then
         Sql$ = "INSERT INTO plaverhoras values('" & wcia & "','" & VTipo & "','" & Trim(vbgcts.CellText(iFila, 3)) & "','','" & wuser & "'," & FechaSys & ")"
         cn.Execute Sql$
      End If
   Next
End If

cn.CommitTrans
Screen.MousePointer = vbDefault
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault
End Sub

Private Sub InicializaGrilla()
With vbgcts
    .Redraw = False
    
      
      ' Set grid lines
      .GridLines = True
      .GridLineMode = ecgGridFillControl
      
      ' Various display and behaviour settings
      .HighlightSelectedIcons = False
      .RowMode = True
      .Editable = True
      .SingleClickEdit = True
      ' Currently there's a problem if you set StretchLastColumnToFit = true
      ' when the grid's redraw style is set to true, as the first column
      ' ends up the wrong width.
      .StretchLastColumnToFit = True
    
    .ScrollBarStyle = ecgSbrFlat
    .ImageList = vbalImageList1
    
    .AddColumn "concepto", "Concepto", ecgHdrTextALignCentre, , 290
    .AddColumn "chk", "", ecgHdrTextALignCentre, , 40
    .AddColumn "codigo", "", ecgHdrTextALignCentre, , , False

    .SetHeaders
    
    .Redraw = True
End With
End Sub

Private Sub vbgcts_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
vbgcts.Redraw = False
If Me.vbgcts.Rows > 0 Then
    If lCol = 2 Then
        bCancel = True
        If vbgcts.CellIcon(lRow, lCol) = iCHCKINAC Then
            vbgcts.CellIcon(lRow, lCol) = iCHCKACT
        Else
            vbgcts.CellIcon(lRow, lCol) = iCHCKINAC
        End If
    End If
End If
vbgcts.Redraw = True

End Sub


