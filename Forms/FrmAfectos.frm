VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form FrmAfectos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Configuración de Conceptos Afectos «"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "FrmAfectos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6015
      Left            =   -15
      TabIndex        =   6
      Top             =   1515
      Width           =   8415
      Begin vbAcceleratorSGrid6.vbalGrid vbgcts 
         Height          =   5775
         Left            =   135
         TabIndex        =   14
         Top             =   135
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   10186
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
         Height          =   5820
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10266
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            DataField       =   "descripcion"
            Caption         =   "Concepto Remunerativo"
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
               ColumnWidth     =   6809.953
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   645.165
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   75
      TabIndex        =   3
      Top             =   540
      Width           =   2010
      Begin VB.OptionButton Opcdeduc 
         Caption         =   "Deduccion"
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   540
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Opcapor 
         Caption         =   "Aportacion"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox Cmbdedapo 
      Height          =   315
      Left            =   3285
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   2775
   End
   Begin VB.ComboBox Cmbtipbol 
      Height          =   315
      Left            =   3285
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   80
         Width           =   7170
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   8
         Top             =   120
         Width           =   825
      End
   End
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   7755
      Top             =   975
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   14924
      Images          =   "FrmAfectos.frx":030A
      Version         =   131072
      KeyCount        =   13
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
      Height          =   195
      Index           =   1
      Left            =   2445
      TabIndex        =   13
      Top             =   630
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
      Height          =   195
      Left            =   6795
      TabIndex        =   12
      Top             =   675
      Width           =   450
   End
   Begin VB.Label Lblfecha 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   6795
      TabIndex        =   11
      Top             =   945
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "T. Boleta"
      Height          =   195
      Index           =   0
      Left            =   2430
      TabIndex        =   10
      Top             =   1125
      Width           =   645
   End
End
Attribute VB_Name = "FrmAfectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsremunera As New Recordset
Dim VTipo As String
Dim VConcepto As String
Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Sql$ = "select codinterno,descripcion from placonstante where tipomovimiento='03' and status not in('*','F') and (adicional<>'S' or codinterno='13') and cia='" & wcia & "' order by descripcion"
Call rCarCbo(Cmbdedapo, Sql$, "C", "00")
Call fc_Descrip_Maestros2("01078", "", Cmbtipbol)
Procesa
End Sub

Private Sub Cmbdedapo_Click()
VConcepto = fc_CodigoComboBox(Cmbdedapo, 2)
Procesa
End Sub

Private Sub Cmbtipbol_Click()
VTipo = fc_CodigoComboBox(Cmbtipbol, 2)
Procesa
End Sub

Private Sub Dgrdafectos_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 2
            If UCase(Trim(Dgrdafectos.Columns(ColIndex))) <> "S" And Trim(Dgrdafectos.Columns(ColIndex)) <> "" Then
               MsgBox "Solo Puede ser [S]i", vbCritical, "Remuneraciones Afectas"
               Dgrdafectos.Columns(ColIndex) = ""
            End If
End Select
End Sub

Private Sub Dgrdafectos_OnAddNew()
rsremunera.AddNew
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 8490
Me.Height = 7935
InicializaGrilla
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub

Public Sub Grabar_Afectos()
Dim mTipo As String
Dim iFila As Long
Dim NroTranas As Integer
On Error GoTo ErrorTrans
NroTrans = 0
If Opcapor.Value <> True And Opcdeduc.Value <> True Then MsgBox "Debe Indicar Aportacion o Deduciion", vbInformation, "Remuneraciones Afectas": Exit Sub
If Cmbdedapo.ListIndex < 0 Then MsgBox "Debe Seleccionar Concepto", vbInformation, "Remuneraciones Afectas": Cmbdedapo.SetFocus: Exit Sub
If Cmbtipbol.ListIndex < 0 Then MsgBox "Debe Seleccionar Tipo de Boleta", vbInformation, "Remuneraciones Afectas": Cmbtipbol.SetFocus: Exit Sub
Mgrab = MsgBox("Seguro de Grabar Remuneraciones Afectas", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub
Screen.MousePointer = vbArrowHourglass
If Opcapor.Value = True Then mTipo = "A" Else mTipo = "D"

cn.BeginTrans
NroTrans = 1
If wGrupoPla = "01" Then
   Dim Rq As ADODB.Recordset
   Sql = "select cod_cia from cia where status<>'*' and prefijo<>'' and not prefijo is null"
   If fAbrRst(Rq, Sql) Then Rq.MoveFirst
   Do While Not Rq.EOF
   
      Sql$ = "update plaafectos set status='*' where tipo='" & mTipo & "' and status<>'*' and codigo='" & VConcepto & "' " _
           & "and tboleta='" & VTipo & "' AND CIA='" & Rq!cod_cia & "'"
      cn.Execute Sql$

      

      For iFila = 1 To vbgcts.Rows
         If Me.vbgcts.CellIcon(iFila, 2) = iCHCKACT Then
            Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql
            Sql$ = Sql$ & " INSERT INTO plaafectos values('" & Rq!cod_cia & "','" & mTipo & "','" & Trim(VConcepto) & "','" & Trim(VTipo) & "', " _
                & "'" & Me.vbgcts.CellText(iFila, 3) & "','','" & wuser & "'," & FechaSys & ")"
            cn.Execute Sql$
         End If
      Next iFila
      Rq.MoveNext
   Loop
Else
   Sql$ = "update plaafectos set status='*' where tipo='" & mTipo & "' and status<>'*' and codigo='" & VConcepto & "' " _
        & "and tboleta='" & VTipo & "' AND CIA='" & wcia & "'"
   cn.Execute Sql$

   For iFila = 1 To vbgcts.Rows
      If Me.vbgcts.CellIcon(iFila, 2) = iCHCKACT Then
         Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql
         Sql$ = Sql$ & " INSERT INTO plaafectos values('" & wcia & "','" & mTipo & "','" & Trim(VConcepto) & "','" & Trim(VTipo) & "', " _
             & "'" & Me.vbgcts.CellText(iFila, 3) & "','','" & wuser & "'," & FechaSys & ")"
         cn.Execute Sql$
      End If
   Next iFila
End If

cn.CommitTrans

MsgBox "Grabacion Exitosa", vbInformation, Me.Caption
Screen.MousePointer = vbDefault

Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox ERR.Description, vbCritical, Me.Caption
Screen.MousePointer = vbDefault
End Sub
Private Sub Procesa()
Dim mTipo
vbgcts.Clear
If Cmbcia.ListIndex < 0 Then Exit Sub
If Cmbtipbol.ListIndex < 0 Then Exit Sub
If Cmbdedapo.ListIndex < 0 Then Exit Sub
If Opcapor.Value = True Then mTipo = "A" Else mTipo = "D"

Sql$ = "SELECT pc.codinterno,pc.descripcion,case when not tipo is null then tipo else '' end tipo FROM placonstante pc LEFT OUTER " & _
"JOIN plaafectos pa ON (pa.cod_remu=pc.codinterno AND pa.cia='" & wcia & "' and pa.tipo='" & mTipo & "' and pa.status<>'*' and" & _
" pa.codigo='" & VConcepto & "' and pa.tboleta='" & VTipo & "') WHERE pc.tipomovimiento='02' and pc.status!='*' and pc.cia='" & wcia & "'"

Set rs = cn.Execute(Sql)

If Not rs.EOF Then
vbgcts.Redraw = False
    Do While Not rs.EOF
        
        With vbgcts
            .AddRow
            .CellDetails .Rows, 1, rs!Descripcion
            .CellDetails .Rows, 2, , DT_CENTER, IIf(Len(Trim(rs!Tipo)) = 0, iCHCKINAC, iCHCKACT), , , , 15
            .CellDetails .Rows, 3, rs!codinterno
            
        End With
        
        rs.MoveNext
    Loop
vbgcts.Redraw = True
rs.Close
End If
Set rs = Nothing

End Sub

Private Sub Opcapor_Click()
Procesa
End Sub

Private Sub Opcdeduc_Click()
Procesa
End Sub

Private Sub InicializaGrilla()
With vbgcts
        .Redraw = False
    
      .GridLines = True
      .GridLineMode = ecgGridFillControl
      
      .HighlightSelectedIcons = False
      .RowMode = True
      .Editable = True
      .SingleClickEdit = True
      
      .StretchLastColumnToFit = True
    
    .ScrollBarStyle = ecgSbrFlat
    .ImageList = vbalImageList1
    
    .AddColumn "concepto", "Concepto Remunerativo", ecgHdrTextALignCentre, , 450
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
