VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form FrmSeteoCts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remuneraciones Afectas a CTS"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "FrmSeteoCts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   6735
      Begin VB.TextBox TxtFactor 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   30
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Factor de Calculo Mensual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   6
         Top             =   45
         Width           =   2160
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5775
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   6735
      Begin vbAcceleratorSGrid6.vbalGrid vbgcts 
         Height          =   5505
         Left            =   135
         TabIndex        =   8
         Top             =   135
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   9710
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
      Begin MSDataGridLib.DataGrid Dgrdafectos 
         Height          =   5535
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9763
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
            DataField       =   "ingresos"
            Caption         =   "Ingresos"
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
         BeginProperty Column02 
            DataField       =   "afecto"
            Caption         =   "Afecto"
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
               ColumnWidth     =   5325.166
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   540.284
            EndProperty
         EndProperty
      End
      Begin vbalIml6.vbalImageList vbalImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   953
         _ExtentY        =   953
         Size            =   14924
         Images          =   "FrmSeteoCts.frx":030A
         Version         =   131072
         KeyCount        =   13
         Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿ"
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   5655
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
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmSeteoCts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsremunera As New Recordset
Dim Sql As String

'Private Sub Dgrdafectos_AfterColEdit(ByVal ColIndex As Integer)
'Select Case ColIndex
'       Case Is = 2
'            If Dgrdafectos.Columns(ColIndex) <> "S" And Dgrdafectos.Columns(ColIndex) <> "" Then
'               MsgBox "Solo Puede ser [S]i", vbCritical, "Remuneraciones Afectas"
'               Dgrdafectos.Columns(ColIndex) = ""
'            End If
'End Select
'End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7170
Me.Width = 6855
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
'Crea_Rs
InicializaGrilla
'SQL = "select * from placonstante where cia='" & wcia & "' and tipomovimiento='02' and status<>'*' order by codinterno"
'
'cn.CursorLocation = adUseClient
'Set rs = New ADODB.Recordset
'Set rs = cn.Execute(SQL$, 64)
'If rs.RecordCount > 0 Then rs.MoveFirst
'Do While Not rs.EOF
'   rsremunera.AddNew
'   rsremunera!INGRESOS = rs!descripcion
'   rsremunera!codigo = rs!codinterno
'   rs.MoveNext
'Loop
'If rs.State = 1 Then rs.Close
'
'If rsremunera.RecordCount > 0 Then rsremunera.MoveFirst
'Do While Not rsremunera.EOF
'   SQL$ = "select tipo from plaafectos where cia='" & wcia & "' and tipo='C' and status<>'*'  and cod_remu='" & rsremunera!codigo & "'"
'   cn.CursorLocation = adUseClient
'   Set rs = New ADODB.Recordset
'   Set rs = cn.Execute(SQL$, 64)
'   If rs.RecordCount > 0 Then rsremunera!afecto = "S" Else rsremunera!afecto = ""
'   rsremunera.MoveNext
'Loop
'If rs.State = 1 Then rs.Close

'{<MA>}  01/02/07

Sql = "SELECT pc.codinterno,pc.descripcion,pc.cod_concepto,CASE WHEN NOT pa.tipo IS NULL  then 'S' ELSE '' END AS afecto" & _
        " FROM placonstante pc LEFT OUTER JOIN plaafectos pa ON (pa.cia='" & wcia & "' and pa.tipo='C' and pa.status<>'*'  and pa.cod_remu=pc.codinterno)" & _
        " WHERE pc.cia='" & wcia & "' and pc.tipomovimiento='02' and pc.status<>'*' ORDER BY pc.codinterno"
        
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then rs.MoveFirst

If Not rs.EOF Then
    Do While Not rs.EOF
'        rsremunera.AddNew
'        rsremunera!INGRESOS = rs!descripcion
'        rsremunera!codigo = rs!codinterno
'        rsremunera!afecto = rs!afecto
'        rs.MoveNext

        With vbgcts
            .AddRow
            .CellDetails .Rows, 1, rs!Descripcion, DT_LEFT
            .CellDetails .Rows, 2, , DT_CENTER, IIf(rs!AFECTO = "S", iCHCKACT, iCHCKINAC), , , , 20
            .CellDetails .Rows, 3, rs!codinterno
        End With
        rs.MoveNext
    Loop
    rs.Close
End If


'If rsremunera.RecordCount > 0 Then rsremunera.MoveFirst

Sql$ = "select factorcts from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then TxtFactor.Text = IIf(IsNull(rs!factorcts), 0, rs!factorcts)
rs.Close
End Sub
'Private Sub Crea_Rs()
'    If rsremunera.State = 1 Then rsremunera.Close
'    rsremunera.Fields.Append "ingresos", adChar, 35, adFldIsNullable
'    rsremunera.Fields.Append "codigo", adChar, 2, adFldIsNullable
'    rsremunera.Fields.Append "afecto", adChar, 1, adFldIsNullable
'    rsremunera.Open
'    Set Dgrdafectos.DataSource = rsremunera
'End Sub

Private Sub TxtFactor_KeyPress(KeyAscii As Integer)
TxtFactor.Text = TxtFactor.Text + fc_ValDecimal(KeyAscii)
End Sub

Private Sub TxtFactor_LostFocus()
If Trim(TxtFactor.Text) <> "" Then
   If IsNull(TxtFactor.Text) Then
      MsgBox "Ingrese Correctamente el Porcentaje", vbCritical, "CTS"
      TxtFactor.SetFocus
   End If
End If
End Sub
Public Sub Grabar_Seteo_Cts()
Dim Mgrab As Integer
Dim iFila As Long
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0

If Not IsNumeric(TxtFactor.Text) Then MsgBox "Ingrese Correctamente Factor", vbCritical, "CTS": Screen.MousePointer = vbDefault: Exit Sub
Mgrab = MsgBox("Seguro de Grabar Remuneraciones Afectas a CTS", vbYesNo + vbQuestion, "Parametros de CTS")
If Mgrab <> 6 Then Exit Sub
Screen.MousePointer = vbArrowHourglass
'If rsremunera.RecordCount > 0 Then rsremunera.MoveFirst

cn.BeginTrans
NroTrans = 1

Sql$ = "update plaafectos set status='*' where cia='" & wcia & "' and tipo='C' and status<>'*'"
cn.Execute Sql$

'Do While Not rsremunera.EOF
'   If rsremunera!afecto = "S" Then
'      SQL$ = "INSERT INTO plaafectos values('" & wcia & "','C','','', " _
'          & "'" & rsremunera!codigo & "','','" & wuser & "'," & FechaSys & ")"
'
'      cn.Execute SQL$
'   End If
'   rsremunera.MoveNext
'Loop

vbgcts.Redraw = False
For iFila = 1 To vbgcts.Rows
    If vbgcts.CellIcon(iFila, 2) = iCHCKACT Then
                
              Sql$ = "INSERT INTO plaafectos values('" & wcia & "','C','','', " & "'" & vbgcts.CellText(iFila, 3) & "','','" & wuser & "'," & FechaSys & ")"

      cn.Execute Sql$
    End If
Next iFila
vbgcts.Redraw = True

Sql$ = "update cia set factorcts=" & CCur(TxtFactor.Text) & " where cod_cia='" & wcia & "' and status<>'*'"
cn.Execute Sql$

cn.CommitTrans
MsgBox "Se guardarón los datos correctamente", vbInformation, Me.Caption
Screen.MousePointer = vbDefault
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
    
    .AddColumn "concepto", "Concepto", ecgHdrTextALignCentre, , 340
    .AddColumn "chk", "Afecto", ecgHdrTextALignCentre, , 65
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

