VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmRemuneracion_CTS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Remuneración - CTS «"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "FrmRemuneracion_CTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkVaca 
      Caption         =   "Provision Vacaciones"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox Cbo_Banco 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1755
      Width           =   4575
   End
   Begin VB.ComboBox Cbo_End 
      Height          =   315
      ItemData        =   "FrmRemuneracion_CTS.frx":030A
      Left            =   2100
      List            =   "FrmRemuneracion_CTS.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   540
      Width           =   1695
   End
   Begin VB.TextBox txt_End 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   540
      Width           =   615
   End
   Begin VB.TextBox txt_Start 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox Cbo_Start 
      Height          =   315
      ItemData        =   "FrmRemuneracion_CTS.frx":039A
      Left            =   2100
      List            =   "FrmRemuneracion_CTS.frx":03C2
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox Cbo_Concepto_Bol 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1365
      Width           =   4575
   End
   Begin VB.ComboBox CboTipo_Trab 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
   Begin MSForms.CommandButton btn 
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2145
      Width           =   1695
      Caption         =   "     Procesar"
      PicturePosition =   327683
      Size            =   "2990;661"
      Picture         =   "FrmRemuneracion_CTS.frx":042A
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ent. Bancaria"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   13
      Top             =   1755
      Width           =   960
   End
   Begin MSForms.SpinButton sp_End 
      Height          =   315
      Left            =   1815
      TabIndex        =   11
      Top             =   540
      Width           =   255
      Size            =   "450;556"
   End
   Begin MSForms.SpinButton sp_Start 
      Height          =   315
      Left            =   1815
      TabIndex        =   10
      Top             =   120
      Width           =   255
      Size            =   "450;556"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   8
      Top             =   540
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Boleta"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   1365
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   1125
   End
End
Attribute VB_Name = "FrmRemuneracion_CTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cadena      As String
Dim mTipo_Trab  As String
Dim mTipo_Banco As String
Dim mProceso    As String
Dim mAll        As Integer

Private Sub Trae_Tipo_Trab()
    Cadena = "SP_TRAE_TIPO_TRAB"
    Call rCarCbo(CboTipo_Trab, Cadena, "XX", "00")
    CboTipo_Trab.ListIndex = 0
End Sub

Private Sub Trae_Concepto_Boleta()
    Cadena = "SP_TRAE_CONCEPTO_BOL"
    Call rCarCbo(Cbo_Concepto_Bol, Cadena, "XX", "00")
    Cbo_Concepto_Bol.ListIndex = 0
End Sub

Private Sub Trae_Ent_Bancaria()
    Cadena = "SP_TRAE_ENT_BANCARIA"
    Call rCarCbo(Cbo_Banco, Cadena, "XX", "00")
    Cbo_Banco.ListIndex = 0
End Sub

Private Function Ver_Fecha() As Boolean
    Ver_Fecha = False
    If CDate(DateSerial(txt_Start.Text, (Cbo_Start.ListIndex + 1) + 0, 1)) > CDate(DateSerial(txt_End.Text, (Cbo_End.ListIndex + 1) + 1, 0)) Then
        MsgBox "Fecha incorrecta, Verifique.", vbExclamation + vbOKOnly, "Sistema"
        Exit Function
    End If
    Ver_Fecha = True
End Function

Private Sub btn_Click()
    If Ver_Fecha = False Then Exit Sub
    If mTipo_Trab = "99" Then MsgBox "No ha seleccionado el tipo de trabajador, Verifique", vbExclamation + vbOKOnly, "Sistema": Exit Sub
    If ChkVaca.Value = 1 Then
       Exp_Excel_Vaca
    Else
       Call Exp_Excel
    End If
End Sub

Private Sub Cbo_Banco_Click()
    mTipo_Banco = Empty
    mTipo_Banco = Trim(fc_CodigoComboBox(Cbo_Banco, 2))
End Sub

Private Sub Cbo_Concepto_Bol_Click()
    mAll = 0
    mProceso = Empty
    mProceso = Trim(fc_CodigoComboBox(Cbo_Concepto_Bol, 2))
    If mProceso = "99" Then mAll = 0 Else mAll = 1
End Sub

Private Sub CboTipo_Trab_Click()
    mTipo_Trab = Empty
    mTipo_Trab = Trim(fc_CodigoComboBox(CboTipo_Trab, 2))
End Sub

Private Sub ChkVaca_Click()
If ChkVaca.Value = 1 Then
   txt_Start.Visible = False
   sp_Start.Visible = False
   Cbo_Start.Visible = False
   lbl(3).Visible = False
   Cbo_Concepto_Bol.Visible = False
Else
   txt_Start.Visible = True
   sp_Start.Visible = True
   Cbo_Start.Visible = True
   lbl(3).Visible = True
   Cbo_Concepto_Bol.Visible = True
  
End If
End Sub

Private Sub Form_Load()
    Call Init
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub sp_End_SpinDown()
    Call Down(txt_End)
End Sub

Private Sub sp_End_SpinUp()
    Call Up(txt_End)
End Sub

Private Sub sp_Start_SpinDown()
    Call Down(txt_Start)
End Sub

Private Sub sp_Start_SpinUp()
    Call Up(txt_Start)
End Sub

Private Sub Init()
    Call Trae_Tipo_Trab
    Call Trae_Concepto_Boleta
    Call Trae_Ent_Bancaria
    txt_Start.Text = Format(Year(Now), "0000")
    txt_End.Text = Format(Year(Now), "0000")
    Cbo_Start.ListIndex = 0
    Cbo_End.ListIndex = Month(Now) - 1
End Sub

Private Sub Exp_Excel()
    Dim rsPersonal  As ADODB.Recordset
    Dim rsInfo      As ADODB.Recordset
    Dim rsdatos     As ADODB.Recordset
    Dim dtStart     As Date
    Dim dtEnd       As Date
    Dim Diferencia  As Integer
    Dim ObjExcel    As Object
    Dim Columna     As Integer
    Dim Fila        As Integer
    Const Formato   As String = "#,##0.00"
    Dim mEXCEL      As String
    'FORMATEAR LA FECHA DE INICIO Y FIN DE LA CONSULTA
    'DiasMes = Day(DateSerial(Year(fecha), Month(fecha) + 1, 0))
    dtStart = DateSerial(txt_Start.Text, (Cbo_Start.ListIndex + 1) + 0, 1)
    dtEnd = DateSerial(txt_End.Text, (Cbo_End.ListIndex + 1) + 1, 0)
    
    'ARMAMOS LA CABECERA DE LA HORA EN EXCEL
        If Not ObjExcel Is Nothing Then
            Set ObjExcel = Nothing
        End If
    mEXCEL = "REM_CTS" & "_" & Left(Cbo_Start.Text, 3) & "." & txt_Start.Text & "_" & Left(Cbo_End.Text, 3) & "." & txt_End.Text
    Set ObjExcel = CreateObject("Excel.Application")
    ObjExcel.Workbooks.Add
    ObjExcel.Application.StandardFont = "Arial"
    ObjExcel.Application.StandardFontSize = "8"
    ObjExcel.ActiveWorkbook.Sheets(1).Name = mEXCEL ' "REM_CTS" & "_" & Left(Cbo_Start.Text, 3) & "." & txt_Start.Text & "_" & Left(Cbo_End.Text, 3) & "." & txt_End.Text
    ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
    ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
    ObjExcel.ActiveSheet.PageSetup.Zoom = 100
    Fila = 1
    With ObjExcel
        .Range("A" & Fila & ":A" & Fila).Value = Trae_CIA(wcia)
        Fila = 2
        .Range("A" & Fila & ":A" & Fila).Value = "RUC N° " & Trae_RUC(wcia)
        Fila = 3
        .Range("A" & Fila & ":A" & Fila).Value = "DIRECCIÓN " & Trae_DIRECCION(wcia)
        Fila = 5
        .Range("A" & Fila & ":A" & Fila).Value = "DNI"
        .Range("B" & Fila & ":B" & Fila).Value = "APELLIDOS Y NOMBRES"
        .Range("C" & Fila & ":C" & Fila).Value = "ENTIDAD DEPOSITORIA - CTS"
        .Range("D" & Fila & ":D" & Fila).Value = "N° CTA CTE"
        .Range("E" & Fila & ":E" & Fila).Value = "IMPORTE DE REMUNERACIONES BRUTAS"
    End With
    
    Fila = 6
    
    'ARMAMOS LA CABECERA
    Diferencia = DateDiff("m", dtStart, dtEnd) + 1
    Dim I As Integer
    Columna = 4
    For I = 1 To Diferencia
'        If Format(dtStart, "mm/yyyy") = Format(dtEnd, "mm/yyyy") Then
'            Exit For
'        End If
        
        ObjExcel.Range(m_LetraColumna(Columna + I) & Fila & ":" & m_LetraColumna(Columna + I) & Fila).Value = mes_palabras2(Month(dtStart)) & Space(1) & Year(dtStart)
        dtStart = DateAdd("m", 1, dtStart)
    Next I
    
    ObjExcel.Range(m_LetraColumna(Columna + Diferencia + 1) & Fila & ":" & m_LetraColumna(Columna + Diferencia + 1) & Fila).Value = "TOTAL"
    
    'VOLVEMOS SETEAR LOS LA FECHA DE INICIO
    dtStart = DateSerial(txt_Start.Text, (Cbo_Start.ListIndex + 1) + 0, 1)
    
    'TRAEMOS EL CODIGO DEL TRABAJADOR CON REGISTROS EN ESAS FECHAS
    Cadena = "SP_TRAE_PERSONAL_BOLETA '" & wcia & "', '" & mTipo_Trab & "', '" & mTipo_Banco & "'"
    Set rsPersonal = OpenRecordset(Cadena, cn)
    If rsPersonal.EOF Then
        MsgBox "No existe informácion para mostrar, Verifique.", vbExclamation + vbOKOnly, "Sistema"
        Exit Sub
    End If
    Fila = 7
    
    If Not rsPersonal.EOF Then
        Do While Not rsPersonal.EOF
            Columna = 1
            Cadena = "SP_TRAE_INFO_TRAB '" & wcia & "', '" & Trim(rsPersonal!PlaCod) & "'"
            Set rsdatos = OpenRecordset(Cadena, cn)
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = rsdatos!DNI
            Columna = 2
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = rsdatos!nombre
            Columna = 3
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = rsdatos!banco
            Columna = 4
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = rsdatos!BCO_CTA
            
            For I = 1 To Diferencia
'                If Format(dtStart, "mm/yyyy") = Format(dtEnd, "mm/yyyy") Then
'                    Exit For
'                End If
                
                Cadena = "SP_TRAE_IMPORTE_BOLETA '" & wcia & "', " & Year(dtStart) & ", " & Month(dtStart) & ", '" & mProceso & "', " & mAll & ", '" & Trim(rsPersonal!PlaCod) & "', '" & mTipo_Trab & "'"
                Set rsInfo = OpenRecordset(Cadena, cn)
                ObjExcel.Range(m_LetraColumna(Columna + I) & Fila & ":" & m_LetraColumna(Columna + I) & Fila).Value = rsInfo!Total
            
                dtStart = DateAdd("m", 1, dtStart)
            Next I
            
            'TOTALISAMOS POR TRABAJADOR
            ObjExcel.Range(m_LetraColumna(Columna + Diferencia + 1) & Fila & ":" & m_LetraColumna(Columna + Diferencia + 1) & Fila).Value = "=SUM(RC[-" & Diferencia & "]:RC[-1])"
            
            'VOLVEMOS SETEAR LA FECHA DE INICIO
            dtStart = DateSerial(txt_Start.Text, (Cbo_Start.ListIndex + 1) + 0, 1)
            Fila = Fila + 1
            rsPersonal.MoveNext
        Loop
    End If
    
    'FORMATO PARA LA CABECERA
    ObjExcel.Range("A1:A1").Font.Bold = True
    ObjExcel.Range("A2:A2").Font.Bold = True
    ObjExcel.Range("A3:A3").Font.Bold = True
    ObjExcel.Range("A1:A1").Font.Size = 9
    ObjExcel.Range("A2:A2").Font.Size = 9
    ObjExcel.Range("A3:A3").Font.Size = 9
    ObjExcel.Range("A1:" & m_LetraColumna(Diferencia + 4) & "1").Merge
    ObjExcel.Range("A2:" & m_LetraColumna(Diferencia + 4) & "2").Merge
    ObjExcel.Range("A3:" & m_LetraColumna(Diferencia + 4) & "3").Merge
    
    'PARA LAS CELDAS DEL DETALLE
    ObjExcel.Range("A5:D5").Font.Bold = True
    ObjExcel.Range("E5:" & m_LetraColumna(Diferencia + 4) & "5").Font.Bold = True
    ObjExcel.Range("A5:A6,B5:B6,C5:C6,D5:D6" & "," & m_LetraColumna(Diferencia + 5) & "5:" & m_LetraColumna(Diferencia + 5) & "6").Select
    ObjExcel.Selection.VerticalAlignment = xlCenter
    ObjExcel.Selection.HorizontalAlignment = xlCenter
    ObjExcel.Selection.MergeCells = True
    ObjExcel.Selection.Font.Bold = True
    ObjExcel.Selection.Interior.ColorIndex = 41
    ObjExcel.Selection.Font.ColorIndex = 2
    ObjExcel.Columns("A:D").EntireColumn.AutoFit
    ObjExcel.Range("E6:" & m_LetraColumna(Diferencia + 4) & "6").Select
    ObjExcel.Selection.Font.Bold = True
    ObjExcel.Selection.Interior.ColorIndex = 41
    ObjExcel.Selection.Font.ColorIndex = 2
    ObjExcel.Range("E5:" & m_LetraColumna(Diferencia + 4) & "5").Select
    ObjExcel.Selection.Interior.ColorIndex = 41
    ObjExcel.Selection.Font.ColorIndex = 2
    'ObjExcel.Selection.MergeCells = True
    ObjExcel.Range("A5:" & m_LetraColumna(Diferencia + 4) & "5").HorizontalAlignment = xlCenter

    Columna = 4
    For I = 1 To Diferencia
        ObjExcel.Columns(m_LetraColumna(Columna + I) & ":" & m_LetraColumna(Columna + I)).NumberFormat = Formato '"##0.00"
        ObjExcel.Columns(m_LetraColumna(Columna + 1 + I) & ":" & m_LetraColumna(Columna + 1 + I)).EntireColumn.AutoFit
    Next I
    
    ObjExcel.Columns(m_LetraColumna(Columna + Diferencia + 1) & ":" & m_LetraColumna(Columna + Diferencia + 1)).NumberFormat = Formato '"##0.00"
    ObjExcel.Columns(m_LetraColumna(Columna + Diferencia + 1) & ":" & m_LetraColumna(Columna + Diferencia + 1)).EntireColumn.AutoFit
    ObjExcel.Columns(m_LetraColumna(Columna + Diferencia + 1) & ":" & m_LetraColumna(Columna + Diferencia + 1)).Font.Bold = True
    If MsgBox("Desea visualizar el archivo ahora?", vbQuestion + vbYesNo, "Sistema") = vbYes Then
        ObjExcel.Visible = True
    Else
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).SaveAs Path_Excel & mEXCEL & ".xls"
        MsgBox "Archivo grabado." & vbCrLf & "UBICACION : " & Path_Excel & mEXCEL, vbInformation + vbOKOnly, "Sistema"
        ObjExcel.Quit
    End If
    
    If Not ObjExcel Is Nothing Then
            Set ObjExcel = Nothing
    End If
    
End Sub
Private Sub Exp_Excel_Vaca()
    Dim rsPersonal  As ADODB.Recordset
    Dim rsInfo      As ADODB.Recordset
    Dim rsdatos     As ADODB.Recordset
    Dim dtStart     As Date
    Dim dtEnd       As Date
    Dim Diferencia  As Integer
    Dim ObjExcel    As Object
    Dim Columna     As Integer
    Dim Fila        As Integer
    Const Formato   As String = "#,##0.00"
    Dim mEXCEL      As String
    'FORMATEAR LA FECHA DE INICIO Y FIN DE LA CONSULTA
    'DiasMes = Day(DateSerial(Year(fecha), Month(fecha) + 1, 0))
    dtStart = DateSerial(txt_Start.Text, (Cbo_Start.ListIndex + 1) + 0, 1)
    dtEnd = DateSerial(txt_End.Text, (Cbo_End.ListIndex + 1) + 1, 0)
    
    'ARMAMOS LA CABECERA DE LA HORA EN EXCEL
        If Not ObjExcel Is Nothing Then
            Set ObjExcel = Nothing
        End If
    mEXCEL = "REM_CTS" & "_" & Left(Cbo_Start.Text, 3) & "." & txt_Start.Text & "_" & Left(Cbo_End.Text, 3) & "." & txt_End.Text
    Set ObjExcel = CreateObject("Excel.Application")
    ObjExcel.Workbooks.Add
    ObjExcel.Application.StandardFont = "Arial"
    ObjExcel.Application.StandardFontSize = "8"
    ObjExcel.ActiveWorkbook.Sheets(1).Name = mEXCEL ' "REM_CTS" & "_" & Left(Cbo_Start.Text, 3) & "." & txt_Start.Text & "_" & Left(Cbo_End.Text, 3) & "." & txt_End.Text
    ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
    ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
    ObjExcel.ActiveSheet.PageSetup.Zoom = 100
    Fila = 1
    With ObjExcel
        .Range("A" & Fila & ":A" & Fila).Value = Trae_CIA(wcia)
        Fila = 2
        .Range("A" & Fila & ":A" & Fila).Value = "RUC N° " & Trae_RUC(wcia)
        Fila = 3
        .Range("A" & Fila & ":A" & Fila).Value = "DIRECCIÓN " & Trae_DIRECCION(wcia)
        Fila = 5
        .Range("A" & Fila & ":A" & Fila).Value = "DNI"
        .Range("B" & Fila & ":B" & Fila).Value = "APELLIDO PATERNO"
        .Range("C" & Fila & ":C" & Fila).Value = "APELLIDO MATERNO"
        .Range("D" & Fila & ":D" & Fila).Value = "PRIMER NOMBRE"
        .Range("E" & Fila & ":E" & Fila).Value = "SEGUNDO NOMBRE"
        
        .Range("F" & Fila & ":F" & Fila).Value = "ENTIDAD DEPOSITORIA - CTS"
        .Range("G" & Fila & ":G" & Fila).Value = "N° CTA CTE"
        .Range("H" & Fila & ":H" & Fila).Value = "REMUN. DE PROV. DE VACACIONES"
        
    End With
    
    Fila = 6
    
    'ARMAMOS LA CABECERA
    'Diferencia = DateDiff("m", dtStart, dtEnd) + 1
    Diferencia = 1
    Dim I As Integer
    Columna = 8

    ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = mes_palabras2(Month(dtEnd)) & Space(1) & Year(dtEnd)
    dtStart = DateAdd("m", 1, dtStart)

    
    ObjExcel.Range(m_LetraColumna(9) & Fila - 1 & ":" & m_LetraColumna(9) & Fila - 1).Value = "X 4"
    
    ObjExcel.Range(m_LetraColumna(10) & Fila - 1 & ":" & m_LetraColumna(10) & Fila - 1).Value = "MONEDA CTA. CTS"
    
    'VOLVEMOS SETEAR LOS LA FECHA DE INICIO
    dtStart = DateSerial(txt_Start.Text, (Cbo_Start.ListIndex + 1) + 0, 1)
    
    'TRAEMOS EL CODIGO DEL TRABAJADOR CON REGISTROS EN ESAS FECHAS
    Cadena = "SP_TRAE_PERSONAL_BOLETA '" & wcia & "', '" & mTipo_Trab & "', '" & mTipo_Banco & "'"
    Set rsPersonal = OpenRecordset(Cadena, cn)
    If rsPersonal.EOF Then
        MsgBox "No existe informácion para mostrar, Verifique.", vbExclamation + vbOKOnly, "Sistema"
        Exit Sub
    End If
    Fila = 7
    
    If Not rsPersonal.EOF Then
        Do While Not rsPersonal.EOF
            Columna = 1
            Cadena = "SP_TRAE_INFO_TRAB '" & wcia & "', '" & Trim(rsPersonal!PlaCod) & "'"
            Set rsdatos = OpenRecordset(Cadena, cn)
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = rsdatos!DNI
            Columna = 2
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = RTrim(rsdatos!ap_pat)
            Columna = 3
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = RTrim(rsdatos!ap_mat)
            Columna = 4
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = RTrim(rsdatos!nom_1)
            Columna = 5
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = RTrim(rsdatos!nom_2)
            Columna = 6
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = rsdatos!banco
            Columna = 7
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = rsdatos!BCO_CTA
            
           Cadena = "SP_TRAE_IMPORTE_BOLETA_PROV_VACA '" & wcia & "', " & Year(dtEnd) & ", " & Month(dtEnd) & ", '" & Trim(rsPersonal!PlaCod) & "'"
            Set rsInfo = OpenRecordset(Cadena, cn)
            ObjExcel.Range(m_LetraColumna(Columna + 1) & Fila & ":" & m_LetraColumna(Columna + 1) & Fila).Value = rsInfo!Total
            ObjExcel.Range(m_LetraColumna(Columna + 2) & Fila & ":" & m_LetraColumna(Columna + 2) & Fila).Value = Round(rsInfo!Total * 4, 2)
            
            ObjExcel.Range(m_LetraColumna(Columna + 3) & Fila & ":" & m_LetraColumna(Columna + 3) & Fila).Value = rsdatos!ctsmoneda
            
            Fila = Fila + 1
            rsPersonal.MoveNext
        Loop
    End If
    
    'FORMATO PARA LA CABECERA
    ObjExcel.Range("A1:A1").Font.Bold = True
    ObjExcel.Range("A2:A2").Font.Bold = True
    ObjExcel.Range("A3:A3").Font.Bold = True
    ObjExcel.Range("A1:A1").Font.Size = 9
    ObjExcel.Range("A2:A2").Font.Size = 9
    ObjExcel.Range("A3:A3").Font.Size = 9
    ObjExcel.Range("A1:" & m_LetraColumna(Diferencia + 4) & "1").Merge
    ObjExcel.Range("A2:" & m_LetraColumna(Diferencia + 4) & "2").Merge
    ObjExcel.Range("A3:" & m_LetraColumna(Diferencia + 4) & "3").Merge
    
    'PARA LAS CELDAS DEL DETALLE
    ObjExcel.Range("A5:D5").Font.Bold = True
    ObjExcel.Range("E5:" & m_LetraColumna(Diferencia + 4) & "5").Font.Bold = True
    ObjExcel.Range("A5:A6,B5:B6,C5:C6,D5:D6,E5:E6,F5:F6,G5:G6,H5:H6,I5:I6,J5:J6" & "," & m_LetraColumna(Diferencia + 5) & "5:" & m_LetraColumna(Diferencia + 5) & "6").Select
    ObjExcel.Selection.VerticalAlignment = xlCenter
    ObjExcel.Selection.HorizontalAlignment = xlCenter
    ObjExcel.Selection.MergeCells = True
    ObjExcel.Selection.Font.Bold = True
    ObjExcel.Selection.Interior.ColorIndex = 41
    ObjExcel.Selection.Font.ColorIndex = 2
    ObjExcel.Columns("A:D").EntireColumn.AutoFit
    ObjExcel.Range("E6:" & m_LetraColumna(Diferencia + 4) & "6").Select
    ObjExcel.Selection.Font.Bold = True
    ObjExcel.Selection.Interior.ColorIndex = 41
    ObjExcel.Selection.Font.ColorIndex = 2
    ObjExcel.Range("E5:" & m_LetraColumna(Diferencia + 4) & "5").Select
    ObjExcel.Selection.Interior.ColorIndex = 41
    ObjExcel.Selection.Font.ColorIndex = 2
    'ObjExcel.Selection.MergeCells = True
    ObjExcel.Range("A5:" & m_LetraColumna(Diferencia + 4) & "5").HorizontalAlignment = xlCenter

    ObjExcel.Columns(m_LetraColumna(8) & ":" & m_LetraColumna(9)).NumberFormat = Formato '"##0.00"
    ObjExcel.Columns(m_LetraColumna(1) & ":" & m_LetraColumna(9)).EntireColumn.AutoFit

    If MsgBox("Desea visualizar el archivo ahora?", vbQuestion + vbYesNo, "Sistema") = vbYes Then
        ObjExcel.Visible = True
    Else
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).SaveAs Path_Excel & mEXCEL & ".xls"
        MsgBox "Archivo grabado." & vbCrLf & "UBICACION : " & Path_Excel & mEXCEL, vbInformation + vbOKOnly, "Sistema"
        ObjExcel.Quit
    End If
    
    If Not ObjExcel Is Nothing Then
            Set ObjExcel = Nothing
    End If
    
End Sub

