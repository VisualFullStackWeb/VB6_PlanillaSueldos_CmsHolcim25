VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmHorasPlanta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horas Planta"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   LinkTopic       =   "Horas Planta"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox CboTipo_Trab 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.TextBox Txt_Year 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      MaxLength       =   4
      TabIndex        =   1
      Top             =   45
      Width           =   645
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "FrmHorasPlanta.frx":0000
      Left            =   1440
      List            =   "FrmHorasPlanta.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   780
      Width           =   3735
   End
   Begin MSComctlLib.ProgressBar pbCadebecera 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   3
      Top             =   1695
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbDetalle 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSForms.CommandButton btn 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   1140
      Width           =   1515
      Caption         =   "     Procesar"
      PicturePosition =   327683
      Size            =   "2672;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1170
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   795
      Width           =   930
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   1365
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   45
      Width           =   660
   End
   Begin MSForms.SpinButton Sb_Year 
      Height          =   315
      Left            =   2055
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Width           =   255
      Size            =   "450;556"
   End
End
Attribute VB_Name = "FrmHorasPlanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cadena      As String
Dim mTipo       As String
Dim pAll        As Integer
Dim Id_Trab     As String
Dim Ors         As ADODB.Recordset
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook

Private Sub btn_Click()
    Screen.MousePointer = vbHourglass
    Me.pbCadebecera.Value = 0
    Me.pbDetalle.Value = 0
    Me.lbl(4).Caption = Empty
    Me.lbl(4).Caption = "Iniciando el Proceso..."
    Me.lbl(4).Refresh
    If Len(Me.Txt_Year.Text) = 0 Then MsgBox "Error de Usuario: Verificar la Fecha.": Exit Sub
       Call Exp_Excel
      
    Screen.MousePointer = vbDefault

End Sub
Private Sub Exp_Excel()
Dim rs As ADODB.Recordset
Dim nFil As Integer
Dim Sql As String

Sql = "usp_pla_listar_horas_planta '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "HORAS PLANTA"
xlApp2.Sheets("HORAS PLANTA").Select


If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE HORAS POR PLANTA MES " & Cmbmes.Text & " " & CInt(Txt_Year.Text)
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1

xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "TIPO REM."
xlSheet.Cells(nFil, 3).Value = "PLANTA"
xlSheet.Cells(nFil, 4).Value = "HORAS"
xlSheet.Cells(nFil, 5).Value = "NRO. TRABAJADORES"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).Font.Bold = True


nFil = 6
Dim lCount As Integer
lCount = 1

Dim ccosto, v_tipo As String
Dim v_total_h, v_total_t  As Long
ccosto = ""
v_tipo = ""
v_total_h = 0
v_total_t = 0
If rs.RecordCount > 0 Then rs.MoveFirst
Dim msum As Integer
msum = 0
Do While Not rs.EOF
      
      If v_tipo <> rs!tipo And v_tipo <> "" Then
        nFil = nFil + 1
        xlSheet.Cells(nFil, 3).Value = "TOTAL"
        xlSheet.Cells(nFil, 4).Value = v_total_h
        xlSheet.Cells(nFil, 5).Value = v_total_t
        v_total_h = 0
        v_total_t = 0
        nFil = nFil + 3
        xlSheet.Cells(nFil, 2).Value = "TIPO REM."
        xlSheet.Cells(nFil, 3).Value = "PLANTA"
        xlSheet.Cells(nFil, 4).Value = "HORAS"
        xlSheet.Cells(nFil, 5).Value = "NRO. TRABAJADORES"
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).HorizontalAlignment = xlCenter
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).VerticalAlignment = xlCenter
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).Borders.LineStyle = xlContinuous
        xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 5)).Font.Bold = True
        nFil = nFil + 1
    End If
     v_tipo = rs!tipo
     v_total_h = v_total_h + rs!horas
     v_total_t = v_total_t + rs!NroTrabjadores
    
      xlSheet.Cells(nFil, 2).Value = Trim(rs!tipo & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!Planta & "")
      xlSheet.Cells(nFil, 4).Value = Trim(rs!horas & "")
      xlSheet.Cells(nFil, 5).Value = Trim(rs!NroTrabjadores & "")
      nFil = nFil + 1
      msum = msum + 1

   rs.MoveNext
Loop
nFil = nFil + 1
xlSheet.Cells(nFil, 3).Value = "TOTAL"
xlSheet.Cells(nFil, 4).Value = v_total_h
xlSheet.Cells(nFil, 5).Value = v_total_t
rs.Close: Set rs = Nothing

nFil = nFil + 5

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE"
xlSheet.Cells(nFil, 4).Value = "HORAS"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 4)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 4)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 4)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 4)).Font.Bold = True

nFil = nFil + 1

Sql = "usp_pla_listar_horas_planta2 '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub
            
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      xlSheet.Cells(nFil, 4).Value = Trim(rs!horas & "")
      nFil = nFil + 1
   rs.MoveNext
Loop


nFil = 2

Set xlSheet = xlApp2.Worksheets("HOJA3")
xlSheet.Name = "HORAS PLANTA DET"
xlApp2.Sheets("HORAS PLANTA DET").Select

xlSheet.Cells(nFil, 2).Value = "TIPO TRAB"
xlSheet.Cells(nFil, 3).Value = "CODIGO"
xlSheet.Cells(nFil, 4).Value = "PLANTA"
xlSheet.Cells(nFil, 5).Value = "CCOSTO"
xlSheet.Cells(nFil, 6).Value = "HNORMAL"
xlSheet.Cells(nFil, 7).Value = "HDOMINICAL"
xlSheet.Cells(nFil, 8).Value = "HEXTRAS"

xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 8)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 8)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 8)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 8)).Font.Bold = True

nFil = nFil + 1

Sql = "usp_pla_listar_horas_planta3 '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub
            
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
      xlSheet.Cells(nFil, 2).Value = Trim(rs!TipoTrab & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 4).Value = Trim(rs!Planta & "")
      xlSheet.Cells(nFil, 5).Value = Trim(rs!ccosto & "")
      xlSheet.Cells(nFil, 6).Value = Trim(rs!Hnormal & "")
      xlSheet.Cells(nFil, 7).Value = Trim(rs!hdominical & "")
      xlSheet.Cells(nFil, 8).Value = Trim(rs!hextras & "")
      nFil = nFil + 1
   rs.MoveNext
Loop

rs.Close: Set rs = Nothing


xlApp2.Application.ActiveWindow.DisplayGridLines = False

xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  Call Load
End Sub
Private Sub Load()
    Call Trae_Tipo_Trab
    'Call Trae_Trabajador
    Me.Txt_Year.Text = Empty
    lbl(4).Caption = Empty
    Me.pbCadebecera.Value = 0
    Me.pbDetalle.Value = 0
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not Ors Is Nothing Then
        If Ors.State = adStateOpen Then Ors.Close
        Set Ors = Nothing
    End If
End Sub
Private Sub txt_Year_KeyPress(KeyAscii As Integer)
    Call NumberOnly(KeyAscii)
End Sub

Private Sub Sb_Year_SpinDown()
    Call Down(Me.Txt_Year)
End Sub

Private Sub Sb_Year_SpinUp()
    Call Up(Me.Txt_Year)
End Sub

Private Sub CboTipo_Trab_Click()
    mTipo = Empty
    mTipo = Trim(fc_CodigoComboBox(Me.CboTipo_Trab, 2))
End Sub

Private Sub Trae_Tipo_Trab()
    Cadena = "SP_TRAE_TIPO_TRAB"
    Call rCarCbo(Me.CboTipo_Trab, Cadena, "XX", "00")
    Me.CboTipo_Trab.ListIndex = 0
End Sub

