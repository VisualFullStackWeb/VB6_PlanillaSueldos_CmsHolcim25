VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPromedios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "» Reporte Detalle Promedios «"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "FrmPromedios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4440
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   4260
      Begin VB.CheckBox ChkTodos 
         Caption         =   "Todos Los Trabajadores"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox TxtCodTrabajador 
         Height          =   315
         Left            =   2520
         TabIndex        =   10
         Top             =   1560
         Width           =   1515
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Semestral"
         Height          =   200
         Index           =   1
         Left            =   1080
         TabIndex        =   1
         Top             =   2040
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.ComboBox CmbTrabTipo 
         Height          =   315
         ItemData        =   "FrmPromedios.frx":030A
         Left            =   1350
         List            =   "FrmPromedios.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   975
         Width           =   2760
      End
      Begin VB.TextBox TxtAño 
         Height          =   285
         Left            =   1350
         MaxLength       =   4
         TabIndex        =   4
         Top             =   225
         Width           =   1140
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Anual"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   2040
         Width           =   1080
      End
      Begin MSComCtl2.DTPicker dtfecha 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "mmmm-aa"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   330
         Left            =   1350
         TabIndex        =   3
         Top             =   570
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM"
         Format          =   56492035
         UpDown          =   -1  'True
         CurrentDate     =   40878
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Trabajador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   11
         Top             =   1620
         Width           =   1560
      End
      Begin MSForms.CommandButton cmdproc 
         Height          =   375
         Left            =   2475
         TabIndex        =   9
         Top             =   1980
         Width           =   1605
         VariousPropertyBits=   268435483
         Caption         =   "     Procesar"
         PicturePosition =   327683
         Size            =   "2831;661"
         Picture         =   "FrmPromedios.frx":030E
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesar al Mes"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   645
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   540
      End
   End
End
Attribute VB_Name = "FrmPromedios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsParametros As ADODB.Recordset
Dim rscalculo As ADODB.Recordset
Dim RsTrabajadores As ADODB.Recordset
Dim RsValida As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cadi As String
Dim cadi2 As String
Dim fecha1 As String
Dim VTipo As String
Dim fecha2 As String
Dim cont As Integer
Dim año As String
Dim Mes As Integer
Dim Sql As String
Dim ObjExcel As Object
Dim Fila As Integer
Dim FILAMES As Integer
Dim FILAINT As Integer
Dim Columna As Integer
Dim FILAMON As Integer
Dim sw As Boolean
Dim suma As Double
Dim contador As Integer
Dim J As Integer
Dim I As Integer
Dim s As Integer



Private Sub ChkTodos_Click()
If ChkTodos.Value = 1 Then
    TxtCodTrabajador.Enabled = False
    TxtCodTrabajador.Text = ""
Else
    TxtCodTrabajador.Enabled = True
    TxtCodTrabajador.Text = ""
End If
End Sub

Private Sub CmbTrabTipo_Click()
    VTipo = fc_CodigoComboBox(CmbTrabTipo, 2)
End Sub

Private Sub cmdproc_Click()
    Call Exp_Excel2
End Sub

Private Sub Form_Load()
    Cadena = "SP_TRAE_TIPO_TRAB"
    Call rCarCbo(CmbTrabTipo, Cadena, "XX", "00")
    CmbTrabTipo.ListIndex = 0
    ChkTodos.Value = 1
    Me.Top = 0
    Me.Left = 0
    Me.dtfecha.Value = Date
End Sub

Public Sub Exp_Excel2()
    Dim iYear       As Integer
    Dim iMonth      As Integer
    Dim Inicio      As String
    Dim Final       As String
    Dim Cadena      As String
    Dim rs          As ADODB.Recordset
    Dim rsTmp       As ADODB.Recordset
    Dim rsTemp      As ADODB.Recordset
    Dim rsAux       As ADODB.Recordset
    Dim Ors         As ADODB.Recordset
    Dim bAnual      As Boolean
    Dim FlgTrabajador As Integer
    Screen.MousePointer = vbHourglass
    If CmbTrabTipo.ListIndex = -1 Then Screen.MousePointer = vbDefault: Exit Sub
    If IsNull(VTipo) Or VTipo = "" Then VTipo = Trim(Format(CmbTrabTipo.ItemData(CmbTrabTipo.ListIndex), "00"))
    If Opt(0).Value = True Then bAnual = True Else bAnual = False
    If TxtCodTrabajador.Text <> "" Then FlgTrabajador = 1 Else FlgTrabajador = 0
    
    If bAnual Then
        iYear = Val(TxtAño.Text)
    Else
        iYear = Val(TxtAño.Text)
        iYear = IIf(dtfecha.Month = 1, iYear - 1, iYear)
        iMonth = dtfecha.Month
        If iMonth = 1 Then iMonth = 12 Else iMonth = iMonth - 1
        Final = Ultimo_Dia(iMonth, iYear) & "/" & iMonth & "/" & iYear
        'Inicio = Ultimo_Dia(Format(DateAdd("m", -5, Final), "MM"), Format(DateAdd("m", -5, Final), "yyyy")) & "/" & Format(DateAdd("m", -5, Final), "MM/yyyy")
        Inicio = "01/" & Format(DateAdd("m", -5, Final), "MM/yyyy")
    End If

    Sql = "SELECT A.CODINTERNO,B.DESCRIPCION FROM PLATASAANEXO A,PLACONSTANTE B WHERE A.MODULO='01' AND A.STATUS<>'*' " & _
        "AND A.TIPOMOVIMIENTO='" & VTipo & "' AND A.BASECALCULO='16' AND A.TIPOTRAB='" & VTipo & "' AND " & _
        "B.TIPOMOVIMIENTO='02' AND A.CARGO='' AND B.STATUS<>'*' AND A.CIA=B.CIA AND A.CODINTERNO=B.CODINTERNO " & _
        "AND B.CIA='" & wcia & "' ORDER BY A.CODINTERNO"
    Set rsTmp = OpenRecordset(Sql, cn)
    If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If rsTmp.AbsolutePosition = rsTmp.RecordCount Then
                Cadena = Cadena & "i" & Trim(rsTmp!codinterno)
            Else
                Cadena = Cadena & "i" & Trim(rsTmp!codinterno) & "+"
            End If
            rsTmp.MoveNext
        Loop
    End If
    rsTmp.MoveFirst
    Dim cadt As String
    If FlgTrabajador = 1 Then
        cadt = " AND PLACOD='" & Trim(TxtCodTrabajador.Text) & "' "
    Else
        cadt = ""
    End If
    If bAnual Then
        Sql = "SET DATEFORMAT MDY " & _
            "SELECT *FROM (SELECT PLACOD,SUM(" & Cadena & ")/12 AS PROMEDIO " & _
            "FROM PLAHISTORICO WHERE YEAR(FECHAPROCESO)=" & iYear & " " & _
            "AND PROCESO='01' AND STATUS<>'*' AND CIA='" & wcia & "' " & _
             cadt & _
            " GROUP BY PLACOD) B WHERE B.PROMEDIO<>0 ORDER BY B.PLACOD"
    Else
        Sql = "SET DATEFORMAT MDY " & _
            "SELECT *FROM (SELECT PLACOD,SUM(" & Cadena & ")/6 AS PROMEDIO " & _
            "FROM PLAHISTORICO WHERE FECHAPROCESO BETWEEN '" & Format(Inicio, "MM/DD/YYYY") & "' " & _
            "AND '" & Format(Final, "MM/DD/YYYY") & "' AND PROCESO='01' AND STATUS<>'*' AND CIA='" & wcia & "' " & _
             cadt & _
            "GROUP BY PLACOD) B WHERE B.PROMEDIO<>0 ORDER BY B.PLACOD"
    End If
    Set rsTemp = OpenRecordset(Sql, cn)
    
    If Not ObjExcel Is Nothing Then
        ObjExcel.Close False
        Set ObjExcel = Nothing
    End If
    
    Set ObjExcel = CreateObject("Excel.Application")
    ObjExcel.Workbooks.Add
    ObjExcel.Application.StandardFont = "Arial"
    ObjExcel.Application.StandardFontSize = "8"
    ObjExcel.ActiveWorkbook.Sheets(1).Name = "Detalle.de.Promedios"
    ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
    ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
    With ObjExcel.ActiveSheet.PageSetup
        .LeftMargin = .Application.InchesToPoints(0)
        .RightMargin = .Application.InchesToPoints(0)
        .TopMargin = ObjExcel.Application.InchesToPoints(0.393700787401575)
        .BottomMargin = ObjExcel.Application.InchesToPoints(0.590551181102362)
        On Error Resume Next
        .PaperSize = xlPaperLetter
        .Zoom = 100
    End With
    Fila = 1
    ObjExcel.Range("A1:F1").Font.Bold = True
    ObjExcel.Range("A3:F3").Font.Bold = True
    ObjExcel.Range("A4:F4").Font.Bold = True
    ObjExcel.Range("C3:F6").Font.Bold = True
    ObjExcel.Range("G3:J6").Font.Bold = True
    ObjExcel.Range("A" & Fila & ":A" & Fila).Value = Trae_CIA(wcia)
    ObjExcel.Range("A" & Fila & ":B" & Fila).Merge
    Fila = Fila + 2 '=3

    ObjExcel.Range("D" & Fila & ":D" & Fila).Value = "DETALLE DE PROMEDIOS A " & Name_Month(Format(iMonth, "00")) & " DE " & iYear
   'ObjExcel.Range("D" & Fila & ":D" & Fila).Value = "DETALLE DE PROMEDIOS A " & UCase(Format(dtfecha, "MMMM")) & " DE " & iYear
    
    ObjExcel.Range("A" & Fila & ":A" & Fila).Font.Size = 12
    ObjExcel.Range("A" & Fila & ":J" & Fila).Merge
    ObjExcel.Range("A" & Fila & ":A" & Fila).HorizontalAlignment = xlCenter
    ObjExcel.Range("A" & Fila & ":A" & Fila).VerticalAlignment = xlCenter
    
    
    suma = 0
    sw = True
    Do While Not rsTemp.EOF 'PERSONAL
        Do While Not rsTmp.EOF 'CONCEPTO
            suma = 0
            Cadena = "SUM(I" & rsTmp!codinterno & ") AS MONTO,"
            If bAnual Then
                Sql = "SET DATEFORMAT MDY " & _
                    "SELECT PLACOD," & Cadena & " MONTH(FECHAPROCESO) AS MES " & _
                    "FROM PLAHISTORICO WHERE YEAR(FECHAPROCESO)=" & iYear & " " & _
                    "AND PROCESO='01' AND PLACOD='" & Trim(rsTemp!PlaCod) & "' AND " & _
                    "CIA='" & wcia & "' AND STATUS<>'*' GROUP BY MONTH(FECHAPROCESO),PLACOD ORDER BY MONTH(FECHAPROCESO)"
            Else
                Sql = "SET DATEFORMAT MDY " & _
                    "SELECT PLACOD," & Cadena & " MONTH(FECHAPROCESO) AS MES " & _
                    "FROM PLAHISTORICO WHERE FECHAPROCESO BETWEEN '" & Format(Inicio, "MM/DD/YYYY") & "' AND " & _
                    "'" & Format(Final, "MM/DD/YYYY") & "' AND PROCESO='01' AND PLACOD='" & Trim(rsTemp!PlaCod) & "' AND " & _
                    "CIA='" & wcia & "' AND STATUS<>'*' GROUP BY MONTH(FECHAPROCESO),PLACOD ORDER BY MONTH(FECHAPROCESO)"
            End If
            'Set Rs = OpenRecordset(Sql, cn)
            Set Ors = OpenRecordset(Sql, cn)
            Ors.MoveFirst
            'Do While Not Rs.EOF
                Sql = "SELECT DISTINCT PLACOD,RTRIM(AP_PAT)+' '+RTRIM(AP_MAT)+' '+RTRIM(NOM_1)+' '+RTRIM(NOM_2) AS NOMBRE " & _
                      "FROM PLANILLAS WHERE PLACOD='" & Trim(rsTemp!PlaCod) & "' AND CIA='" & wcia & "' AND STATUS<>'*'"
                Set rsAux = OpenRecordset(Sql, cn)
                
                Columna = 1
                If sw Then
                    'COLOCA EL NOMBRE EN LA FILA 5
                    Fila = Fila + 2
                    If Not rsAux.EOF Then
                        For I = 0 To 1
                            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = Trim(rsAux(I).Value)
                            Columna = Columna + 1
                        Next I
                    End If
                    
                    Columna = 2
                    Fila = Fila + 2 '=7
                    'PINTAMOS EL MES
                    Ors.MoveFirst
                    For I = 0 To Ors.RecordCount - 1
                        ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = mes_palabras2(Ors!Mes)
                        Columna = Columna + 1
                        Ors.MoveNext
                    Next I
                    'PINTAMOS LA COLUMNAS DE TOTAL Y PROMEDIO
                    ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = "TOTAL"
                    ObjExcel.Range(m_LetraColumna(Columna + 1) & Fila & ":" & m_LetraColumna(Columna + 1) & Fila).Value = "PROMEDIO"
                    sw = False
                End If
                'LA DESCRIPCION DEL CONCEPTO
                Fila = Fila + 1 '=8
                Columna = 1
                ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = Trim(rsTmp!Descripcion)
                Columna = 2
                Ors.MoveFirst
                For I = 0 To Ors.RecordCount - 1
                    ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = Val(Ors!Monto)
                    ObjExcel.Columns(m_LetraColumna(Columna) & ":" & m_LetraColumna(Columna)).NumberFormat = "##0.00"
                    suma = suma + Val(Ors!Monto)
                    Columna = Columna + 1
                    Ors.MoveNext
                Next I
                ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = suma
                ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Font.Bold = True
                ObjExcel.Range(m_LetraColumna(Columna + 1) & Fila & ":" & m_LetraColumna(Columna + 1) & Fila).Value = "=AVERAGE(RC[-" & Ors.RecordCount + 1 & "]:RC[-2])"
                ObjExcel.Columns(m_LetraColumna(Columna + 1) & ":" & m_LetraColumna(Columna + 1)).NumberFormat = "##0.00"
                ObjExcel.Range(m_LetraColumna(Columna + 1) & Fila & ":" & m_LetraColumna(Columna + 1) & Fila).Font.Bold = True
                'Rs.MoveNext
            'Loop
            rsTmp.MoveNext
        Loop
        Columna = 2
        Fila = Fila + 1
        For I = 0 To Ors.RecordCount
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = "=SUM(R[-" & rsTmp.RecordCount & "]C:R[-1]C)"
            ObjExcel.Columns(m_LetraColumna(Columna) & ":" & m_LetraColumna(Columna)).NumberFormat = "##0.00"
            ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Font.Bold = True
            Columna = Columna + 1
        Next I
        ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = "=AVERAGE(RC[-" & Ors.RecordCount + 1 & "]:RC[-2])"
        ObjExcel.Columns(m_LetraColumna(Columna) & ":" & m_LetraColumna(Columna)).NumberFormat = "##0.00"
        ObjExcel.Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Font.Bold = True

        
        Fila = Fila + 3
        sw = True
        rsTemp.MoveNext
        If Not rsTemp.EOF Then rsTmp.MoveFirst
    Loop
    ObjExcel.Columns("A:A").EntireColumn.AutoFit
    ObjExcel.Visible = True
    Set ObjExcel = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Opt_Click(Index As Integer)
    If Opt(0).Value = True Then
        dtfecha.Enabled = False
    Else
        dtfecha.Enabled = True
    End If
End Sub

Private Sub OptPersonal_Click(Index As Integer)
    If Opt(2).Value = True Then
        dtfecha.Enabled = False
    Else
        dtfecha.Enabled = True
    End If
End Sub

Private Sub TxtCodTrabajador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Frmgrdpla.Show
    Frmgrdpla.ZOrder 0
End If
End Sub

