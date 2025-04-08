VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmHorasExtras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Detalle de Horas Extras"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "FrmHorasExtras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptHoras 
      Caption         =   "Horas"
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton OptSoles 
      Caption         =   "Soles"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   0
      Width           =   975
   End
   Begin VB.Frame FrameTrimestres 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CheckBox ChkOtrosPagos 
         Caption         =   "Otros Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1575
      End
      Begin VB.OptionButton Opt4 
         Caption         =   "Cuarto"
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "Tercer"
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "Segundo"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Primer"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CheckBox ChkTrimestral 
      Caption         =   "Trimestral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   60
      Width           =   1215
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "FrmHorasExtras.frx":030A
      Left            =   1440
      List            =   "FrmHorasExtras.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Txt_Year 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      MaxLength       =   4
      TabIndex        =   1
      Top             =   105
      Width           =   645
   End
   Begin VB.ComboBox CboTipo_Trab 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3750
   End
   Begin MSComctlLib.ProgressBar pbCadebecera 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   2
      Top             =   1770
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbDetalle 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   3
      Top             =   1665
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSForms.SpinButton Sb_Year 
      Height          =   315
      Left            =   2055
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   105
      Width           =   255
      Size            =   "450;556"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   105
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   855
      Visible         =   0   'False
      Width           =   765
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
      TabIndex        =   5
      Top             =   1230
      Width           =   1200
   End
   Begin MSForms.CommandButton btn 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   1515
      Caption         =   "     Procesar"
      PicturePosition =   327683
      Size            =   "2672;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FrmHorasExtras"
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
    'If mTipo = "99" Or mTipo = "999" Or mTipo = Empty Then MsgBox "Error de Usuario: Debe de seleccionar un tipo de empleado.": Exit Sub
    If ChkTrimestral.Value = 1 Then
       Call Exp_Excel_Trim
    Else
       Call Exp_Excel
       Call Exp_Excel_antes
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Exp_Excel_antes()
Dim rs As ADODB.Recordset
Dim nFil As Integer
Dim Sql As String
If OptSoles.Value = True Then
Sql = "usp_Pla_Horas_Extras_antes '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""
Else
Sql = "usp_Pla_Horas_Extras_antes_horas '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""
End If
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

'Set xlApp1 = CreateObject("Excel.Application")
'xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "H_ext_meses"
xlApp2.Sheets("H_ext_meses").Select


If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
    xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE HORAS EXTRAS (EMPLEADOS)"
Else
    xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS DE HORAS EXTRAS (EMPLEADOS)"
End If
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Font.Bold = True

Dim I As Integer
Dim lName As String
For I = 4 To Cmbmes.ListIndex + 4
    xlSheet.Cells(nFil, I).Value = rs(I).Name
Next
xlSheet.Cells(nFil, I).Value = "TOTAL"
xlSheet.Cells(nFil, I + 1).Value = "%"

nFil = 5
Dim lCount As Integer
lCount = 1

Dim ccosto As String
ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst
Dim lTotalE As Double
Dim lTotalO As Double
lTotalE = 0: lTotalO = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!PlaCod & "") = "*****" Then lTotalE = rs!Total
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") = "*****" Then lTotalO = rs!Total
   rs.MoveNext
Loop

If rs.RecordCount > 0 Then rs.MoveFirst
Dim msum As Integer
msum = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") <> "*****" Then
      If ccosto <> Trim(rs!ccosto & "") Then
         If ccosto <> "" Then
            msum = msum * -1
            For I = 4 To Cmbmes.ListIndex + 6
               xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            Next
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
            xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
         End If
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
         nFil = nFil + 2
         ccosto = Trim(rs!ccosto & "")
         msum = 0
      End If
        
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      For I = 4 To Cmbmes.ListIndex + 4
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalE
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      nFil = nFil + 1
      msum = msum + 1
   End If
   rs.MoveNext
Loop
msum = msum * -1
For I = 4 To Cmbmes.ListIndex + 6
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") = "*****" Then
     For I = 4 To Cmbmes.ListIndex + 4
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, 3).Value = "TOTAL EMPLEADOS"
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalE
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop

'OBREROS
nFil = nFil + 5

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE HORAS EXTRAS (OBREROS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS DE HORAS EXTRAS (OBREROS)"
End If
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Font.Bold = True

For I = 4 To Cmbmes.ListIndex + 4
    xlSheet.Cells(nFil, I).Value = rs(I).Name
Next
xlSheet.Cells(nFil, I).Value = "TOTAL"
xlSheet.Cells(nFil, I + 1).Value = "%"

If rs.RecordCount > 0 Then rs.MoveFirst
msum = 0: ccosto = ""
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") <> "*****" Then
      If ccosto <> Trim(rs!ccosto & "") Then
         If ccosto <> "" Then
            msum = msum * -1
            For I = 4 To Cmbmes.ListIndex + 6
               xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            Next
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
            xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
         End If
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
         nFil = nFil + 2
         ccosto = Trim(rs!ccosto & "")
         msum = 0
      End If
        
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      For I = 4 To Cmbmes.ListIndex + 4
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalO
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      nFil = nFil + 1
      msum = msum + 1
   End If
   rs.MoveNext
Loop
msum = msum * -1
For I = 4 To Cmbmes.ListIndex + 6
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") = "*****" Then
     For I = 4 To Cmbmes.ListIndex + 4
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, 3).Value = "TOTAL OBREROS"
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalO
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop

'SEMANAS
nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = ".." And Trim(rs!PlaCod & "") = "....." Then
     For I = 4 To Cmbmes.ListIndex + 4
          xlSheet.Cells(nFil, I).Value = rs(I)
          xlSheet.Cells(nFil, I).NumberFormat = "#,##0_ ;[Red]-#,##0 "
          xlSheet.Cells(nFil, I).HorizontalAlignment = xlCenter
          xlSheet.Cells(nFil, I).Font.Bold = True
      Next
      xlSheet.Cells(nFil, 3).Value = "NUMERO DE SEMANAS"
      xlSheet.Cells(nFil, 3).Font.Bold = True
      Exit Do
   End If
   rs.MoveNext
Loop

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "**" And Trim(rs!PlaCod & "") = "....." Then
     For I = 4 To Cmbmes.ListIndex + 4
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, 3).Value = "TOTAL EMPLEADOS y OBREROS"
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / (lTotalE + lTotalO)
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Borders.LineStyle = xlContinuous
      nFil = nFil + 2
      xlSheet.Cells(nFil, I - 4).Value = "PROMEDIO MENSUAL EMPLEADOS Y OBREROS"
      xlSheet.Cells(nFil, I).Value = Round(rs!Total / (Cmbmes.ListIndex + 1), 2)
      xlSheet.Cells(nFil, I).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop
rs.Close: Set rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridLines = False
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub

Private Sub ChkTrimestral_Click()
If ChkTrimestral.Value = 1 Then FrameTrimestres.Visible = True Else FrameTrimestres.Visible = False
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
    Me.OptSoles.Value = True
    Me.OptHoras.Value = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Ors Is Nothing Then
        If Ors.State = adStateOpen Then Ors.Close
        Set Ors = Nothing
    End If
End Sub

Private Function Trae_Informacion() As Boolean
On Error GoTo MyErr
    Screen.MousePointer = vbHourglass
    Trae_Informacion = False
    Me.lbl(4).Refresh
    Me.lbl(4).Caption = "Espere por favor...."
    Me.lbl(4).Refresh
    Cadena = Empty
    Cadena = "EXEC SP_TRAE_HORAS_EXTRAS " & _
            "'" & wcia & "'," & _
            "" & CInt(Me.Txt_Year.Text) & "," & _
            "'" & mTipo & "'," & _
            "" & pAll & "," & _
            "'" & Id_Trab & "'"
    If (Not EXEC_SQL(Cadena, cn)) Then Exit Function
    Trae_Informacion = True
    Screen.MousePointer = vbDefault
    Exit Function
MyErr:
    MsgBox ERR.Number & Space(1) & ERR.Description, vbCritical + vbOKOnly, "Error"
    ERR.Clear
End Function

Private Sub Option1_Click()

End Sub

Private Sub OptHoras_Click()
If OptHoras.Value = True Then OptSoles.Value = False
End Sub

Private Sub OptSoles_Click()
If OptSoles.Value = True Then OptHoras.Value = False
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
    'Call Trae_Trabajador
End Sub

Private Sub Trae_Tipo_Trab()
    Cadena = "SP_TRAE_TIPO_TRAB"
    Call rCarCbo(Me.CboTipo_Trab, Cadena, "XX", "00")
    Me.CboTipo_Trab.ListIndex = 0
End Sub
Private Sub Exp_Excel_Trim()
Dim rs As ADODB.Recordset
Dim nFil As Integer
Dim Sql As String
Dim lMEs As Integer
Dim lTipo As String
Dim ltit As String

If ChkOtrosPagos.Value = 1 Then lTipo = "O": ltit = "OTROS PAGOS" Else lTipo = "E": ltit = "HORAS EXTRAS"

If Opt1.Value Then lMEs = 4
If Opt2.Value Then lMEs = 7
If Opt3.Value Then lMEs = 10
If Opt4.Value Then lMEs = 13

If OptSoles.Value = True Then
Sql = "Usp_Pla_HExtras_Trim '" & wcia & "', " & CInt(Txt_Year.Text) & ",'" & lTipo & "'," & lMEs & ""
Else
Sql = "Usp_Pla_HExtras_Trim_Horas '" & wcia & "', " & CInt(Txt_Year.Text) & ",'" & lTipo & "'," & lMEs & ""
End If

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add

Set xlApp2 = xlApp1.Application

xlApp2.Sheets.Add
xlApp2.Sheets("Hoja1").Select
xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
xlApp2.Sheets("Hoja2").Select
xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)

Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Empleados"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:T").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE PROMEDIO DE " & ltit & " (OBREROS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS PROMEDIO DE " & ltit & " (OBREROS)"
End If

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Merge
xlSheet.Cells(nFil, 4).Value = "AÑO " & Val(Txt_Year.Text) - 1
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).HorizontalAlignment = xlCenter

xlSheet.Cells(nFil, 9).Value = "AÑO " & Val(Txt_Year.Text)
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).Merge
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).HorizontalAlignment = xlCenter

nFil = nFil + 1

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Cells(nFil, 4).Value = "1Trim."
xlSheet.Cells(nFil, 5).Value = "2Trim."
xlSheet.Cells(nFil, 6).Value = "3Trim."
xlSheet.Cells(nFil, 7).Value = "4Trim."
'xlSheet.Cells(nFil, 8).Value = "Total"

xlSheet.Cells(nFil, 9).Value = "1Trim."
xlSheet.Cells(nFil, 10).Value = "2Trim."
xlSheet.Cells(nFil, 11).Value = "3Trim."
xlSheet.Cells(nFil, 12).Value = "4Trim."
'xlSheet.Cells(nFil, 14).Value = "Total"



xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 12)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 12)).HorizontalAlignment = xlCenter

nFil = 5
Dim lCount As Integer
lCount = 1

Dim ccosto As String
ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "01" Then
      If ccosto <> Trim(rs!ccosto & "") Then
             nFil = nFil + 2
            xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
            nFil = nFil + 2
            ccosto = Trim(rs!ccosto & "")
      End If
        
      xlSheet.Cells(nFil, 4).Value = rs!AnteriorPrimert
      xlSheet.Cells(nFil, 5).Value = rs!AnteriorSegundot
      xlSheet.Cells(nFil, 6).Value = rs!AnteriorTercert
      xlSheet.Cells(nFil, 7).Value = rs!AnteriorCuartot
      'xlSheet.Cells(nFil, 8).Value = rs!AnteriorTotal

      xlSheet.Cells(nFil, 9).Value = rs!ActualPrimert
      xlSheet.Cells(nFil, 10).Value = rs!ActualSegundot
      xlSheet.Cells(nFil, 11).Value = rs!ActualTercert
      xlSheet.Cells(nFil, 12).Value = rs!ActualCuartot
      'xlSheet.Cells(nFil, 13).Value = rs!ActualTotal

      If Trim(rs!PlaCod & "") = "Z9***" Then
         xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
      Else
         xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
         xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      End If
      nFil = nFil + 1
   End If
   rs.MoveNext
Loop

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "**" And Trim(rs!PlaCod & "") = "EMPLE" Then
      
      xlSheet.Cells(nFil, 4).Value = rs!AnteriorPrimert
      xlSheet.Cells(nFil, 5).Value = rs!AnteriorSegundot
      xlSheet.Cells(nFil, 6).Value = rs!AnteriorTercert
      xlSheet.Cells(nFil, 7).Value = rs!AnteriorCuartot
      'xlSheet.Cells(nFil, 8).Value = rs!AnteriorTotal

      xlSheet.Cells(nFil, 9).Value = rs!ActualPrimert
      xlSheet.Cells(nFil, 10).Value = rs!ActualSegundot
      xlSheet.Cells(nFil, 11).Value = rs!ActualTercert
      xlSheet.Cells(nFil, 12).Value = rs!ActualCuartot
      'xlSheet.Cells(nFil, 13).Value = rs!ActualTotal
      
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 12)).Font.Bold = True
      Exit Do
   End If
   rs.MoveNext
Loop

xlApp2.Application.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
   
Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "Obreros"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:T").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE PROMEDIO DE " & ltit & " (OBREROS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS PROMEDIO DE " & ltit & " (OBREROS)"
End If


xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Merge
xlSheet.Cells(nFil, 4).Value = "AÑO " & Val(Txt_Year.Text) - 1
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).HorizontalAlignment = xlCenter

xlSheet.Cells(nFil, 9).Value = "AÑO " & Val(Txt_Year.Text)
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).Merge
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 9), xlSheet.Cells(nFil, 12)).HorizontalAlignment = xlCenter

nFil = nFil + 1

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Cells(nFil, 4).Value = "1Trim."
xlSheet.Cells(nFil, 4).Value = "1Trim."
xlSheet.Cells(nFil, 5).Value = "2Trim."
xlSheet.Cells(nFil, 6).Value = "3Trim."
xlSheet.Cells(nFil, 7).Value = "4Trim."
'xlSheet.Cells(nFil, 8).Value = "Total"

xlSheet.Cells(nFil, 9).Value = "1Trim."
xlSheet.Cells(nFil, 10).Value = "2Trim."
xlSheet.Cells(nFil, 11).Value = "3Trim."
xlSheet.Cells(nFil, 12).Value = "4Trim."
'xlSheet.Cells(nFil, 14).Value = "Total"

xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 12)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 12)).HorizontalAlignment = xlCenter

nFil = 5

ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "02" Then
      If ccosto <> Trim(rs!ccosto & "") Then
             nFil = nFil + 2
            xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
            nFil = nFil + 2
            ccosto = Trim(rs!ccosto & "")
      End If
        
      xlSheet.Cells(nFil, 4).Value = rs!AnteriorPrimert
      xlSheet.Cells(nFil, 5).Value = rs!AnteriorSegundot
      xlSheet.Cells(nFil, 6).Value = rs!AnteriorTercert
      xlSheet.Cells(nFil, 7).Value = rs!AnteriorCuartot
      'xlSheet.Cells(nFil, 8).Value = rs!AnteriorTotal

      xlSheet.Cells(nFil, 9).Value = rs!ActualPrimert
      xlSheet.Cells(nFil, 10).Value = rs!ActualSegundot
      xlSheet.Cells(nFil, 11).Value = rs!ActualTercert
      xlSheet.Cells(nFil, 12).Value = rs!ActualCuartot
      'xlSheet.Cells(nFil, 13).Value = rs!ActualTotal
      
      If Trim(rs!PlaCod & "") = "Z9***" Then
         xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
      Else
         xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
         xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      End If
      nFil = nFil + 1
   End If
   rs.MoveNext
Loop

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "**" And Trim(rs!PlaCod & "") = "OBRER" Then
   
      xlSheet.Cells(nFil, 4).Value = rs!AnteriorPrimert
      xlSheet.Cells(nFil, 5).Value = rs!AnteriorSegundot
      xlSheet.Cells(nFil, 6).Value = rs!AnteriorTercert
      xlSheet.Cells(nFil, 7).Value = rs!AnteriorCuartot
      'xlSheet.Cells(nFil, 8).Value = rs!AnteriorTotal

      xlSheet.Cells(nFil, 9).Value = rs!ActualPrimert
      xlSheet.Cells(nFil, 10).Value = rs!ActualSegundot
      xlSheet.Cells(nFil, 11).Value = rs!ActualTercert
      xlSheet.Cells(nFil, 12).Value = rs!ActualCuartot
      'xlSheet.Cells(nFil, 13).Value = rs!ActualTotal
      
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 12)).Borders.LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 12)).Font.Bold = True
      Exit Do
   End If
   rs.MoveNext
Loop

rs.Close: Set rs = Nothing
xlApp2.Application.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
For I = 1 To 2
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridLines = False
   'xlApp2.ActiveWindow.Zoom = 80
Next
xlApp2.Sheets(1).Select
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub
Private Sub Exp_Excel()
Dim rs As ADODB.Recordset
Dim nFil As Integer
Dim Sql As String

If OptSoles.Value = True Then
Sql = "usp_Pla_Horas_Extras '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""
Else
Sql = "usp_Pla_Horas_Extras_horas '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""
End If

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Horas_Ext"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE HORAS EXTRAS (EMPLEADOS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS DE HORAS EXTRAS (EMPLEADOS)"
End If

xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).Font.Bold = True

Dim I As Integer
Dim lName As String
For I = 4 To 5
   If I = 4 Then
      If Cmbmes.ListIndex + 1 = 1 Then
         xlSheet.Cells(nFil, I).Value = "DIC. " & Format(Val(Txt_Year.Text) - 1, "0000")
      Else
         xlSheet.Cells(nFil, I).Value = Name_Month(Format(Cmbmes.ListIndex, "00"))
      End If
   Else
      xlSheet.Cells(nFil, I).Value = Name_Month(Format(Cmbmes.ListIndex + 1, "00"))
   End If
Next

xlSheet.Cells(nFil, I).Value = "DIFERENCIA"
xlSheet.Cells(nFil, I + 1).Value = "Var. %"

nFil = 5
Dim lCount As Integer
lCount = 1

Dim ccosto As String
ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst
Dim lTotalE As Double
Dim lTotalO As Double
lTotalE = 0: lTotalO = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!PlaCod & "") = "*****" Then lTotalE = rs!Diferencia
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") = "*****" Then lTotalO = rs!Diferencia
   rs.MoveNext
Loop

If rs.RecordCount > 0 Then rs.MoveFirst
Dim msum As Integer
msum = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") <> "*****" Then
      If ccosto <> Trim(rs!ccosto & "") Then
         If ccosto <> "" Then
            msum = msum * -1
            For I = 4 To 6
               xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            Next
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
            If xlSheet.Cells(nFil, 4).Value <> 0 Then
               xlSheet.Cells(nFil, I).Value = xlSheet.Cells(nFil, 5).Value / xlSheet.Cells(nFil, 4).Value - 1
            End If
            If xlSheet.Cells(nFil, 4).Value <> 0 And xlSheet.Cells(nFil, 5).Value = 0 Then xlSheet.Cells(nFil, I).Value = -1
            If xlSheet.Cells(nFil, 4).Value = 0 And xlSheet.Cells(nFil, 5).Value <> 0 Then xlSheet.Cells(nFil, I).Value = 1

            xlSheet.Cells(nFil, I).NumberFormat = "0.00%"
         End If
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
         nFil = nFil + 2
         ccosto = Trim(rs!ccosto & "")
         msum = 0
      End If
        
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      For I = 4 To 7
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
      nFil = nFil + 1
      msum = msum + 1
   End If
   rs.MoveNext
Loop
msum = msum * -1
For I = 4 To 6
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
If xlSheet.Cells(nFil, 4).Value <> 0 Then
   xlSheet.Cells(nFil, I).Value = xlSheet.Cells(nFil, 5).Value / xlSheet.Cells(nFil, 4).Value - 1
End If
If xlSheet.Cells(nFil, 4).Value <> 0 And xlSheet.Cells(nFil, 5).Value = 0 Then xlSheet.Cells(nFil, I).Value = -1
If xlSheet.Cells(nFil, 4).Value = 0 And xlSheet.Cells(nFil, 5).Value <> 0 Then xlSheet.Cells(nFil, I).Value = 1

xlSheet.Cells(nFil, I).NumberFormat = "0.00%"

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") = "*****" Then
     For I = 4 To 7
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, 3).Value = "TOTAL EMPLEADOS"
      xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop

'OBREROS
nFil = nFil + 5

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE HORAS EXTRAS (OBREROS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE  HORAS EXTRAS (OBREROS)"
End If
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1

xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 7)).Font.Bold = True


For I = 4 To 5
   If I = 4 Then
      If Cmbmes.ListIndex + 1 = 1 Then
         xlSheet.Cells(nFil, I).Value = "DIC. " & Format(Val(Txt_Year.Text) - 1, "0000")
      Else
         xlSheet.Cells(nFil, I).Value = Name_Month(Format(Cmbmes.ListIndex, "00"))
      End If
   Else
      xlSheet.Cells(nFil, I).Value = Name_Month(Format(Cmbmes.ListIndex + 1, "00"))
   End If
Next

xlSheet.Cells(nFil, I).Value = "DIFERENCIA"
xlSheet.Cells(nFil, I + 1).Value = "Var. %"

lCount = 1
ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst

lTotalE = 0: lTotalO = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!PlaCod & "") = "*****" Then lTotalE = rs!Diferencia
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") = "*****" Then lTotalO = rs!Diferencia
   rs.MoveNext
Loop

If rs.RecordCount > 0 Then rs.MoveFirst

msum = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") <> "*****" Then
      If ccosto <> Trim(rs!ccosto & "") Then
         If ccosto <> "" Then
            msum = msum * -1
            For I = 4 To 6
               xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            Next
            If xlSheet.Cells(nFil, 4).Value <> 0 Then
               xlSheet.Cells(nFil, I).Value = xlSheet.Cells(nFil, 5).Value / xlSheet.Cells(nFil, 4).Value - 1
            End If
            If xlSheet.Cells(nFil, 4).Value <> 0 And xlSheet.Cells(nFil, 5).Value = 0 Then xlSheet.Cells(nFil, I).Value = -1
            If xlSheet.Cells(nFil, 4).Value = 0 And xlSheet.Cells(nFil, 5).Value <> 0 Then xlSheet.Cells(nFil, I).Value = 1
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
            xlSheet.Cells(nFil, I).NumberFormat = "0.00%"
         End If
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
         nFil = nFil + 2
         ccosto = Trim(rs!ccosto & "")
         msum = 0
      End If
        
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      For I = 4 To 7
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
      nFil = nFil + 1
      msum = msum + 1
   End If
   rs.MoveNext
Loop
msum = msum * -1
For I = 4 To 6
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
If xlSheet.Cells(nFil, 4).Value <> 0 Then
   xlSheet.Cells(nFil, I).Value = xlSheet.Cells(nFil, 5).Value / xlSheet.Cells(nFil, 4).Value - 1
End If
If xlSheet.Cells(nFil, 4).Value <> 0 And xlSheet.Cells(nFil, 5).Value = 0 Then xlSheet.Cells(nFil, I).Value = -1
If xlSheet.Cells(nFil, 4).Value = 0 And xlSheet.Cells(nFil, 5).Value <> 0 Then xlSheet.Cells(nFil, I).Value = 1

xlSheet.Cells(nFil, I).NumberFormat = "0.00%"

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") = "*****" Then
     For I = 4 To 7
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, 3).Value = "TOTAL EMPLEADOS"
      xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop
'SEMANAS
nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = ".." And Trim(rs!PlaCod & "") = "....." Then
     For I = 4 To 5
          xlSheet.Cells(nFil, I).Value = rs(I + 1)
          xlSheet.Cells(nFil, I).NumberFormat = "#,##0_ ;[Red]-#,##0 "
          xlSheet.Cells(nFil, I).HorizontalAlignment = xlCenter
          xlSheet.Cells(nFil, I).Font.Bold = True
      Next
      xlSheet.Cells(nFil, 3).Value = "NUMERO DE SEMANAS"
      xlSheet.Cells(nFil, 4).Value = ""
      xlSheet.Cells(nFil, 3).Font.Bold = True
      Exit Do
   End If
   rs.MoveNext
Loop

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "**" And Trim(rs!PlaCod & "") = "....." Then
     For I = 4 To 6
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, 3).Value = "TOTAL EMPLEADOS y OBREROS"
     If xlSheet.Cells(nFil, 4).Value <> 0 Then
        xlSheet.Cells(nFil, I).Value = xlSheet.Cells(nFil, 5).Value / xlSheet.Cells(nFil, 4).Value - 1
     End If
     If xlSheet.Cells(nFil, 4).Value <> 0 And xlSheet.Cells(nFil, 5).Value = 0 Then xlSheet.Cells(nFil, I).Value = -1
     If xlSheet.Cells(nFil, 4).Value = 0 And xlSheet.Cells(nFil, 5).Value <> 0 Then xlSheet.Cells(nFil, I).Value = 1

      xlSheet.Cells(nFil, I).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
      nFil = nFil + 2
      
'      xlSheet.Cells(nFil, I - 4).Value = "PROMEDIO MENSUAL EMPLEADOS Y OBREROS"
'      xlSheet.Cells(nFil, I).Value = Round(rs!Diferencia / (Cmbmes.ListIndex + 1), 2)
'      xlSheet.Cells(nFil, I).Borders.LineStyle = xlContinuous
'
'      xlSheet.Cells(nFil, 4).Value = Round(rs!anterior, 2)
'      xlSheet.Cells(nFil, 4).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop
rs.Close: Set rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridLines = False
'xlApp2.ActiveWindow.Zoom = 80
'xlApp2.Application.Visible = True

'If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
'If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
'If Not xlBook Is Nothing Then Set xlBook = Nothing
'If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub
Private Sub Exp_Excel_Antes_promedio()
Dim rs As ADODB.Recordset
Dim nFil As Integer
Dim Sql As String

Sql = "usp_Pla_Horas_Extras '" & wcia & "', " & CInt(Txt_Year.Text) & ", " & Cmbmes.ListIndex + 1 & ""
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Horas_Ext"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE HORAS EXTRAS (EMPLEADOS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS EXTRAS (EMPLEADOS)"
End If
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Font.Bold = True

Dim I As Integer
Dim lName As String
For I = 4 To Cmbmes.ListIndex + 5
    xlSheet.Cells(nFil, I).Value = rs(I).Name
Next
xlSheet.Cells(nFil, I).Value = "TOTAL"
xlSheet.Cells(nFil, I + 1).Value = "%"

nFil = 5
Dim lCount As Integer
lCount = 1

Dim ccosto As String
ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst
Dim lTotalE As Double
Dim lTotalO As Double
lTotalE = 0: lTotalO = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!PlaCod & "") = "*****" Then lTotalE = rs!Total
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") = "*****" Then lTotalO = rs!Total
   rs.MoveNext
Loop

If rs.RecordCount > 0 Then rs.MoveFirst
Dim msum As Integer
msum = 0
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") <> "*****" Then
      If ccosto <> Trim(rs!ccosto & "") Then
         If ccosto <> "" Then
            msum = msum * -1
            For I = 4 To Cmbmes.ListIndex + 7
               xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            Next
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
            xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
         End If
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
         nFil = nFil + 2
         ccosto = Trim(rs!ccosto & "")
         msum = 0
      End If
        
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      For I = 4 To Cmbmes.ListIndex + 5
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalE
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      nFil = nFil + 1
      msum = msum + 1
   End If
   rs.MoveNext
Loop
msum = msum * -1
For I = 4 To Cmbmes.ListIndex + 7
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "01" And Trim(rs!TipoTrab & "") <> "**" And Trim(rs!PlaCod & "") = "*****" Then
     For I = 4 To Cmbmes.ListIndex + 5
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, 3).Value = "TOTAL EMPLEADOS"
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalE
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop

'OBREROS
nFil = nFil + 5

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE HORAS EXTRAS (OBREROS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS EXTRAS (OBREROS)"
End If
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Merge

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).VerticalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Font.Bold = True

For I = 4 To Cmbmes.ListIndex + 5
    xlSheet.Cells(nFil, I).Value = rs(I).Name
Next
xlSheet.Cells(nFil, I).Value = "TOTAL"
xlSheet.Cells(nFil, I + 1).Value = "%"

If rs.RecordCount > 0 Then rs.MoveFirst
msum = 0: ccosto = ""
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") <> "*****" Then
      If ccosto <> Trim(rs!ccosto & "") Then
         If ccosto <> "" Then
            msum = msum * -1
            For I = 4 To Cmbmes.ListIndex + 7
               xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
            Next
            xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
            xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"
         End If
         nFil = nFil + 2
         xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
         nFil = nFil + 2
         ccosto = Trim(rs!ccosto & "")
         msum = 0
      End If
        
      xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      For I = 4 To Cmbmes.ListIndex + 5
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalO
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      nFil = nFil + 1
      msum = msum + 1
   End If
   rs.MoveNext
Loop
msum = msum * -1
For I = 4 To Cmbmes.ListIndex + 7
   xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
Next
xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
xlSheet.Cells(nFil, I - 1).NumberFormat = "0.00%"

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "02" And Trim(rs!PlaCod & "") = "*****" Then
     For I = 4 To Cmbmes.ListIndex + 5
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, 3).Value = "TOTAL OBREROS"
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / lTotalO
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop

'SEMANAS
nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = ".." And Trim(rs!PlaCod & "") = "....." Then
     For I = 4 To Cmbmes.ListIndex + 5
          xlSheet.Cells(nFil, I).Value = rs(I)
          xlSheet.Cells(nFil, I).NumberFormat = "#,##0_ ;[Red]-#,##0 "
          xlSheet.Cells(nFil, I).HorizontalAlignment = xlCenter
          xlSheet.Cells(nFil, I).Font.Bold = True
      Next
      xlSheet.Cells(nFil, 3).Value = "NUMERO DE SEMANAS"
      xlSheet.Cells(nFil, 4).Value = ""
      xlSheet.Cells(nFil, 3).Font.Bold = True
      Exit Do
   End If
   rs.MoveNext
Loop

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!TipoTrab & "") = "**" And Trim(rs!PlaCod & "") = "....." Then
     For I = 4 To Cmbmes.ListIndex + 5
          xlSheet.Cells(nFil, I).Value = rs(I)
      Next
      xlSheet.Cells(nFil, I).Value = rs!Total
      xlSheet.Cells(nFil, 3).Value = "TOTAL EMPLEADOS y OBREROS"
      xlSheet.Cells(nFil, I + 1).Value = rs!Total / (lTotalE + lTotalO)
      xlSheet.Cells(nFil, I + 1).NumberFormat = "0.00%"
      xlSheet.Range(xlSheet.Cells(nFil, 3), xlSheet.Cells(nFil, Cmbmes.ListIndex + 7)).Borders.LineStyle = xlContinuous
      nFil = nFil + 2
      xlSheet.Cells(nFil, I - 4).Value = "PROMEDIO MENSUAL EMPLEADOS Y OBREROS"
      xlSheet.Cells(nFil, I).Value = Round(rs!Total / (Cmbmes.ListIndex + 1), 2)
      xlSheet.Cells(nFil, I).Borders.LineStyle = xlContinuous
      
      xlSheet.Cells(nFil, 4).Value = Round(rs!promant, 2)
      xlSheet.Cells(nFil, 4).Borders.LineStyle = xlContinuous
      Exit Do
   End If
   rs.MoveNext
Loop
rs.Close: Set rs = Nothing

xlApp2.Application.ActiveWindow.DisplayGridLines = False
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub

Private Sub Exp_Excel_Trim_Antes()
Dim rs As ADODB.Recordset
Dim nFil As Integer
Dim Sql As String
Dim lMEs As Integer
Dim lTipo As String
Dim ltit As String

If ChkOtrosPagos.Value = 1 Then lTipo = "O": ltit = "OTROS PAGOS" Else lTipo = "E": ltit = "HORAS EXTRAS"

If Opt1.Value Then lMEs = 4
If Opt2.Value Then lMEs = 7
If Opt3.Value Then lMEs = 10
If Opt4.Value Then lMEs = 13

Sql = "Usp_Pla_HExtras_Trim '" & wcia & "', " & CInt(Txt_Year.Text) & ",'" & lTipo & "'," & lMEs & ""
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add

Set xlApp2 = xlApp1.Application

xlApp2.Sheets.Add
xlApp2.Sheets("Hoja1").Select
xlApp2.Sheets("Hoja1").Move Before:=xlApp2.Sheets(1)
xlApp2.Sheets("Hoja2").Select
xlApp2.Sheets("Hoja2").Move Before:=xlApp2.Sheets(2)

Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Empleados"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:T").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE " & ltit & " (OBREROS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS DE " & ltit & " (OBREROS)"
End If

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Cells(nFil, 4).Value = "Ene."
xlSheet.Cells(nFil, 5).Value = "Feb."
xlSheet.Cells(nFil, 6).Value = "Mar."
xlSheet.Cells(nFil, 7).Value = "1Trim."
xlSheet.Cells(nFil, 8).Value = "Abr."
xlSheet.Cells(nFil, 9).Value = "May."
xlSheet.Cells(nFil, 10).Value = "Jun."
xlSheet.Cells(nFil, 11).Value = "2Trim."
xlSheet.Cells(nFil, 12).Value = "Jul."
xlSheet.Cells(nFil, 13).Value = "Ago."
xlSheet.Cells(nFil, 14).Value = "Sep."
xlSheet.Cells(nFil, 15).Value = "3Trim."
xlSheet.Cells(nFil, 16).Value = "Oct."
xlSheet.Cells(nFil, 17).Value = "Nov."
xlSheet.Cells(nFil, 18).Value = "Dic."
xlSheet.Cells(nFil, 19).Value = "4Trim."
xlSheet.Cells(nFil, 20).Value = "Total"

xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 20)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 20)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 20)).HorizontalAlignment = xlCenter

nFil = 5
Dim lCount As Integer
lCount = 1

Dim ccosto As String
ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "01" Then
      If ccosto <> Trim(rs!ccosto & "") Then
             nFil = nFil + 2
            xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
            nFil = nFil + 2
            ccosto = Trim(rs!ccosto & "")
      End If
        
      xlSheet.Cells(nFil, 4).Value = rs!Extra01
      xlSheet.Cells(nFil, 5).Value = rs!Extra02
      xlSheet.Cells(nFil, 6).Value = rs!Extra03
      xlSheet.Cells(nFil, 7).Value = rs!Primert
      xlSheet.Cells(nFil, 8).Value = rs!Extra04
      xlSheet.Cells(nFil, 9).Value = rs!Extra05
      xlSheet.Cells(nFil, 10).Value = rs!Extra06
      xlSheet.Cells(nFil, 11).Value = rs!Segundot
      xlSheet.Cells(nFil, 12).Value = rs!Extra07
      xlSheet.Cells(nFil, 13).Value = rs!Extra08
      xlSheet.Cells(nFil, 14).Value = rs!Extra09
      xlSheet.Cells(nFil, 15).Value = rs!Tercert
      xlSheet.Cells(nFil, 16).Value = rs!Extra10
      xlSheet.Cells(nFil, 17).Value = rs!Extra11
      xlSheet.Cells(nFil, 18).Value = rs!Extra12
      xlSheet.Cells(nFil, 19).Value = rs!Cuartot
      xlSheet.Cells(nFil, 20).Value = rs!Total

      If Trim(rs!PlaCod & "") = "Z9***" Then
         xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 20)).Borders.LineStyle = xlContinuous
      Else
         xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
         xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      End If
      nFil = nFil + 1
   End If
   rs.MoveNext
Loop

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "**" And Trim(rs!PlaCod & "") = "EMPLE" Then
      xlSheet.Cells(nFil, 4).Value = rs!Extra01
      xlSheet.Cells(nFil, 5).Value = rs!Extra02
      xlSheet.Cells(nFil, 6).Value = rs!Extra03
      xlSheet.Cells(nFil, 7).Value = rs!Primert
      xlSheet.Cells(nFil, 8).Value = rs!Extra04
      xlSheet.Cells(nFil, 9).Value = rs!Extra05
      xlSheet.Cells(nFil, 10).Value = rs!Extra06
      xlSheet.Cells(nFil, 11).Value = rs!Segundot
      xlSheet.Cells(nFil, 12).Value = rs!Extra07
      xlSheet.Cells(nFil, 13).Value = rs!Extra08
      xlSheet.Cells(nFil, 14).Value = rs!Extra09
      xlSheet.Cells(nFil, 15).Value = rs!Tercert
      xlSheet.Cells(nFil, 16).Value = rs!Extra10
      xlSheet.Cells(nFil, 17).Value = rs!Extra11
      xlSheet.Cells(nFil, 18).Value = rs!Extra12
      xlSheet.Cells(nFil, 19).Value = rs!Cuartot
      xlSheet.Cells(nFil, 20).Value = rs!Total
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 20)).Borders.LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 20)).Font.Bold = True
      Exit Do
   End If
   rs.MoveNext
Loop

xlApp2.Application.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
   
Set xlSheet = xlApp2.Worksheets("HOJA2")
xlSheet.Name = "Obreros"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst Else rs.Close: Set rs = Nothing: Exit Sub

xlSheet.Range("A:A").ColumnWidth = 4
xlSheet.Range("C:C").ColumnWidth = 35
xlSheet.Range("D:T").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

nFil = 1

xlSheet.Cells(nFil, 2).Value = Trae_CIA(wcia)
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "REPORTE DE IMPORTE DE " & ltit & " (OBREROS)"
Else
xlSheet.Cells(nFil, 2).Value = "REPORTE DE HORAS " & ltit & " (OBREROS)"
End If

xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 1
If OptSoles.Value = True Then
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN SOLES"
Else
xlSheet.Cells(nFil, 2).Value = "EXPRESADO EN HORAS"
End If
xlSheet.Cells(nFil, 2).Font.Bold = True
nFil = nFil + 2

xlSheet.Cells(nFil, 2).Value = "CODIGO"
xlSheet.Cells(nFil, 3).Value = "NOMBRE DEL TRABAJADOR"
xlSheet.Cells(nFil, 4).Value = "Ene."
xlSheet.Cells(nFil, 5).Value = "Feb."
xlSheet.Cells(nFil, 6).Value = "Mar."
xlSheet.Cells(nFil, 7).Value = "1Trim."
xlSheet.Cells(nFil, 8).Value = "Abr."
xlSheet.Cells(nFil, 9).Value = "May."
xlSheet.Cells(nFil, 10).Value = "Jun."
xlSheet.Cells(nFil, 11).Value = "2Trim."
xlSheet.Cells(nFil, 12).Value = "Jul."
xlSheet.Cells(nFil, 13).Value = "Ago."
xlSheet.Cells(nFil, 14).Value = "Sep."
xlSheet.Cells(nFil, 15).Value = "3Trim."
xlSheet.Cells(nFil, 16).Value = "Oct."
xlSheet.Cells(nFil, 17).Value = "Nov."
xlSheet.Cells(nFil, 18).Value = "Dic."
xlSheet.Cells(nFil, 19).Value = "4Trim."
xlSheet.Cells(nFil, 20).Value = "Total"
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 20)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 20)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, 20)).HorizontalAlignment = xlCenter

nFil = 5

ccosto = ""

If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "02" Then
      If ccosto <> Trim(rs!ccosto & "") Then
             nFil = nFil + 2
            xlSheet.Cells(nFil, 2).Value = Trim(rs!ccosto & "")
            nFil = nFil + 2
            ccosto = Trim(rs!ccosto & "")
      End If
        
      xlSheet.Cells(nFil, 4).Value = rs!Extra01
      xlSheet.Cells(nFil, 5).Value = rs!Extra02
      xlSheet.Cells(nFil, 6).Value = rs!Extra03
      xlSheet.Cells(nFil, 7).Value = rs!Primert
      xlSheet.Cells(nFil, 8).Value = rs!Extra04
      xlSheet.Cells(nFil, 9).Value = rs!Extra05
      xlSheet.Cells(nFil, 10).Value = rs!Extra06
      xlSheet.Cells(nFil, 11).Value = rs!Segundot
      xlSheet.Cells(nFil, 12).Value = rs!Extra07
      xlSheet.Cells(nFil, 13).Value = rs!Extra08
      xlSheet.Cells(nFil, 14).Value = rs!Extra09
      xlSheet.Cells(nFil, 15).Value = rs!Tercert
      xlSheet.Cells(nFil, 16).Value = rs!Extra10
      xlSheet.Cells(nFil, 17).Value = rs!Extra11
      xlSheet.Cells(nFil, 18).Value = rs!Extra12
      xlSheet.Cells(nFil, 19).Value = rs!Cuartot
      xlSheet.Cells(nFil, 20).Value = rs!Total

      If Trim(rs!PlaCod & "") = "Z9***" Then
         xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 20)).Borders.LineStyle = xlContinuous
      Else
         xlSheet.Cells(nFil, 2).Value = Trim(rs!PlaCod & "")
         xlSheet.Cells(nFil, 3).Value = Trim(rs!nombre & "")
      End If
      nFil = nFil + 1
   End If
   rs.MoveNext
Loop

nFil = nFil + 2
If rs.RecordCount > 0 Then rs.MoveFirst
Do While Not rs.EOF
   If Trim(rs!tipo & "") = "**" And Trim(rs!PlaCod & "") = "OBRER" Then
      xlSheet.Cells(nFil, 4).Value = rs!Extra01
      xlSheet.Cells(nFil, 5).Value = rs!Extra02
      xlSheet.Cells(nFil, 6).Value = rs!Extra03
      xlSheet.Cells(nFil, 7).Value = rs!Primert
      xlSheet.Cells(nFil, 8).Value = rs!Extra04
      xlSheet.Cells(nFil, 9).Value = rs!Extra05
      xlSheet.Cells(nFil, 10).Value = rs!Extra06
      xlSheet.Cells(nFil, 11).Value = rs!Segundot
      xlSheet.Cells(nFil, 12).Value = rs!Extra07
      xlSheet.Cells(nFil, 13).Value = rs!Extra08
      xlSheet.Cells(nFil, 14).Value = rs!Extra09
      xlSheet.Cells(nFil, 15).Value = rs!Tercert
      xlSheet.Cells(nFil, 16).Value = rs!Extra10
      xlSheet.Cells(nFil, 17).Value = rs!Extra11
      xlSheet.Cells(nFil, 18).Value = rs!Extra12
      xlSheet.Cells(nFil, 19).Value = rs!Cuartot
      xlSheet.Cells(nFil, 20).Value = rs!Total
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 20)).Borders.LineStyle = xlContinuous
      xlSheet.Range(xlSheet.Cells(nFil, 4), xlSheet.Cells(nFil, 20)).Font.Bold = True
      Exit Do
   End If
   rs.MoveNext
Loop

rs.Close: Set rs = Nothing
xlApp2.Application.ActiveSheet.PageSetup.PrintTitleRows = "$1:$5"
For I = 1 To 2
   xlApp2.Sheets(I).Select
   xlApp2.Sheets(I).Range("A1:A1").Select
   xlApp2.Application.ActiveWindow.DisplayGridLines = False
   'xlApp2.ActiveWindow.Zoom = 80
Next
xlApp2.Sheets(1).Select
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
    

Screen.MousePointer = vbDefault
End Sub
