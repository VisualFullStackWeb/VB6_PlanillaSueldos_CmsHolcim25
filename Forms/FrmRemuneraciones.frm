VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmRemuneraciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Conceptos Remunerativos (placonstante) «"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   Icon            =   "FrmRemuneraciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6330
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   10440
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
         Left            =   4455
         ScaleHeight     =   2670
         ScaleWidth      =   2715
         TabIndex        =   11
         Top             =   1830
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   14
            Top             =   0
            Width           =   2670
         End
      End
      Begin VB.PictureBox pctempresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   1620
         ScaleHeight     =   795
         ScaleWidth      =   4305
         TabIndex        =   7
         Top             =   2295
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
            TabIndex        =   8
            Top             =   405
            Width           =   3840
         End
         Begin Threed.SSCommand ssc_importa 
            Height          =   330
            Left            =   3915
            TabIndex        =   9
            Top             =   405
            Width           =   330
            _Version        =   65536
            _ExtentX        =   582
            _ExtentY        =   582
            _StockProps     =   78
            BevelWidth      =   1
            AutoSize        =   2
            Picture         =   "FrmRemuneraciones.frx":030A
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
      Begin MSDataGridLib.DataGrid DgrdRemun 
         Height          =   5820
         Left            =   75
         TabIndex        =   4
         Top             =   75
         Width           =   9570
         _ExtentX        =   16880
         _ExtentY        =   10266
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
         ColumnCount     =   10
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
            DataField       =   "calculo"
            Caption         =   "Calculo"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "basico"
            Caption         =   "Basico"
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
            DataField       =   "factor"
            Caption         =   "Factor"
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
            DataField       =   "codconcepto"
            Caption         =   "cod_concepto"
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
         BeginProperty Column08 
            DataField       =   "extraord"
            Caption         =   "Extraordinario"
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
            DataField       =   "promqta"
            Caption         =   "Prom. Qta."
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
               ColumnWidth     =   3644.788
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   5160
         Picture         =   "FrmRemuneraciones.frx":08A4
         Top             =   5985
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
         Height          =   375
         Left            =   75
         TabIndex        =   6
         ToolTipText     =   "Importar Informacion de Otra Cia."
         Top             =   5910
         Width           =   2310
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFEFC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importar Movimientos  "
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
         Height          =   375
         Left            =   5085
         TabIndex        =   5
         ToolTipText     =   "Importar Informacion de Otra Cia."
         Top             =   5925
         Width           =   2400
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
      Width           =   10440
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   150
         Width           =   8655
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
         Left            =   75
         TabIndex        =   1
         Top             =   165
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmRemuneraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VConcepto As String
Dim rsremun As New Recordset
Dim str_cptocargados As String


Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
VConcepto = "02"
Procesa
End Sub

Private Sub cmdsalir_Click()
Dim I As Long

pct_conceptos.Visible = False

For I = 0 To Me.lstconceptos.ListCount - 1
    If lstconceptos.Selected(I) Then
        rsremun.AddNew
        rsremun!concepto = lstconceptos.List(I)
        rsremun!AFECTO = " "
        rsremun!calculo = " "
        rsremun!Codigo = ""
        rsremun!Basico = ""
        rsremun!factor = ""
        rsremun!codconcepto = Format(lstconceptos.ItemData(I), "000")
        
        str_cptocargados = str_cptocargados & "'" & Format(lstconceptos.ItemData(I), "000") & "',"
    End If
Next I

End Sub

Private Sub DgrdRemun_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 1
            If Trim(UCase(DgrdRemun.Columns(1))) <> "S" And Trim(UCase(DgrdRemun.Columns(1))) <> "N" Then
               MsgBox "Calculo Solo Puede ser [S]i o [N]o", vbCritical, TitMsg
               DgrdRemun.Columns(1) = ""
            End If
       Case Is = 2
            If Trim(UCase(DgrdRemun.Columns(2))) <> "S" And Trim(UCase(DgrdRemun.Columns(2))) <> "N" Then
               MsgBox "Afecto Solo Puede ser [S]i o [N]o", vbCritical, TitMsg
               DgrdRemun.Columns(2) = ""
            End If
       Case Is = 4
            If Trim(UCase(DgrdRemun.Columns(4))) <> "S" And Trim(UCase(DgrdRemun.Columns(4))) <> "N" Then
               MsgBox "Basico Solo Puede ser [S]i o [N]o", vbCritical, TitMsg
               DgrdRemun.Columns(4) = ""
            End If
       Case Is = 5
            If Trim(UCase(DgrdRemun.Columns(5))) <> "S" And Trim(UCase(DgrdRemun.Columns(5))) <> "N" Then
               MsgBox "Factor Solo Puede ser [S]i o [N]o", vbCritical, TitMsg
               DgrdRemun.Columns(5) = ""
            End If
       Case Is = 8
            If Trim(UCase(DgrdRemun.Columns(8))) <> "S" And Trim(UCase(DgrdRemun.Columns(8))) <> "N" Then
               MsgBox "Factor Solo Puede ser [S]i o [N]o", vbCritical, TitMsg
               DgrdRemun.Columns(8) = ""
            End If
       Case Is = 9
            If Trim(UCase(DgrdRemun.Columns(9))) <> "S" And Trim(UCase(DgrdRemun.Columns(9))) <> "N" Then
               MsgBox "Factor Solo Puede ser [S]i o [N]o", vbCritical, TitMsg
               DgrdRemun.Columns(9) = ""
            End If
End Select
DgrdRemun.Columns(ColIndex) = Trim(DgrdRemun.Columns(ColIndex))
End Sub

Private Sub DgrdRemun_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub DgrdRemun_OnAddNew()
rsremun.AddNew
End Sub

Private Sub DgrdRemun_RowResize(Cancel As Integer)
Cancel = True
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
Me.Height = 7470
Me.Width = 9870
Crea_Rs
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
End Sub
Private Sub Crea_Rs()
    If rsremun.State = 1 Then rsremun.Close
    rsremun.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsremun.Fields.Append "concepto", adChar, 45, adFldIsNullable
    rsremun.Fields.Append "afecto", adChar, 1, adFldIsNullable
    rsremun.Fields.Append "calculo", adChar, 1, adFldIsNullable
    rsremun.Fields.Append "basico", adChar, 1, adFldIsNullable
    rsremun.Fields.Append "factor", adChar, 1, adFldIsNullable
    rsremun.Fields.Append "codconcepto", adChar, 3, adFldIsNullable
    rsremun.Fields.Append "semanapago", adChar, 2, adFldIsNullable
    rsremun.Fields.Append "codsunat", adChar, 4, adFldIsNullable
    rsremun.Fields.Append "extraord", adChar, 1, adFldIsNullable
    rsremun.Fields.Append "promqta", adChar, 1, adFldIsNullable
    
    rsremun.Open
    Set DgrdRemun.DataSource = rsremun
End Sub
Private Sub Procesa()

str_cptocargados = ""

Sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='" & VConcepto & "' and status<>'*' order by codinterno"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rsremun.RecordCount > 0 Then
   rsremun.MoveFirst
   Do While Not rsremun.EOF
      rsremun.Delete
      rsremun.MoveNext
   Loop
End If
If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsremun.AddNew
   rsremun!concepto = rs!Descripcion
   rsremun!AFECTO = rs!status
   rsremun!calculo = rs!calculo
   rsremun!Codigo = rs!codinterno
   rsremun!Basico = rs!Basico
   rsremun!factor = rs!factor
   rsremun!codconcepto = rs!cod_concepto
   rsremun!semanapago = CStr(IIf(IsNull(rs!semana_pago), "", rs!semana_pago))
   rsremun!CODSUNAT = Trim(rs!CODSUNAT & "")
   rsremun!extraord = Trim(rs!extraord & "")
   rsremun!promqta = Trim(rs!promqta & "")

   str_cptocargados = str_cptocargados & "'" & rs!cod_concepto & "',"
   
   rs.MoveNext
Loop
If rsremun.RecordCount > 0 Then rsremun.MoveFirst
If rs.State = 1 Then rs.Close
End Sub
Public Sub Graba_Remunera()

On Error GoTo Salir
Dim NroTrans As Integer
NroTrans = 0
Dim Mcodigo As String
If MsgBox("Desea Grabar Conceptos Remunerativos", vbYesNo + vbQuestion) = vbNo Then Screen.MousePointer = vbDefault: Exit Sub

cn.BeginTrans
NroTrans = 1
   Dim Rq As ADODB.Recordset
   Sql = "select cod_cia from cia where status<>'*'"
   
   If fAbrRst(Rq, Sql) Then Rq.MoveFirst
   Do While Not Rq.EOF
      Sql = "update placonstante set status='*' where cia='" & Rq!cod_cia & "' and tipomovimiento='" & VConcepto & "' and status<>'*'"
      cn.Execute Sql
      If rsremun.RecordCount > 0 Then
         rsremun.MoveFirst
         Do While Not rsremun.EOF
            If IsNull(rsremun!Codigo) Then rsremun!Codigo = ""
            If Val(rsremun!Codigo) <> 0 Then
               Mcodigo = Trim(rsremun!Codigo)
            Else
               Sql = "select max(codinterno) as codigo from placonstante where cia='" & Rq!cod_cia & "' and tipomovimiento='" & VConcepto & "'"
               cn.CursorLocation = adUseClient
               Set rs = New ADODB.Recordset
               Set rs = cn.Execute(Sql, 64)
               If rs.RecordCount > 0 And Not IsNull(rs!Codigo) Then Mcodigo = Format(Val(rs!Codigo) + 1, "00") Else Mcodigo = "01"
               If rs.State = 1 Then rs.Close
            End If
      
            If Trim(rsremun!concepto) <> "" Then
                Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & "  "
                Sql$ = Sql$ & "INSERT INTO placonstante values('" & Rq!cod_cia & "','" & VConcepto & "','" & Mcodigo & "','" & Trim(rsremun!concepto) & "',0,0,'','" & rsremun!AFECTO & "','" & rsremun!calculo & "',''," & FechaSys & ",'" & rsremun!Basico & "','" & VTipotrab & "','" & rsremun!factor & "','" & rsremun!codconcepto & "'," & IIf(Len(Trim(rsremun!semanapago & "")) = 0, "NULL", rsremun!semanapago) & ",'" & Trim(rsremun!CODSUNAT & "") & "','" & Trim(rsremun!extraord & "") & "','" & Trim(rsremun!promqta & "") & "')"
                cn.Execute Sql$
            End If
            rsremun.MoveNext
         Loop
      End If
      Rq.MoveNext
   Loop
   Rq.Close: Set Rq = Nothing

cn.CommitTrans

MsgBox "Se guardaron los datos correctamente", vbInformation, Me.Caption

Screen.MousePointer = vbDefault
Call Procesa

Exit Sub

Salir:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox "Error " & Err.Description, vbCritical, TitMsg
Screen.MousePointer = vbDefault
End Sub

Sub Eliminar()
    

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

lstconceptos.Clear

If Not rs.EOF Then
    Do While Not rs.EOF
        With lstconceptos
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
Dim NroTrans As Integer
NroTrans = 0
On Error GoTo Importar
cn.BeginTrans
NroTrans = 1
sSQL$ = "EXEC sp_i_importa_data '" & VConcepto & "','" & Format(cbocia.ItemData(cbocia.ListIndex), "00") & "','" & wcia & "'"

cn.Execute sSQL

cn.CommitTrans

MsgBox "Se Realizo la Importación Correctamente", vbQuestion, TitMsg

pctempresa.Visible = False
Procesa

Exit Sub
Importar:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    MsgBox "Error " & Err.Description, vbCritical, TitMsg
End Sub
