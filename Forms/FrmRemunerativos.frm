VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmRemunerativos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Detalle de Conceptos Remunerativos «"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "FrmRemunerativos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRemunra 
      Caption         =   "Solo Conceptos Remunerativos"
      Height          =   435
      Left            =   3360
      TabIndex        =   17
      Top             =   1200
      Width           =   1545
   End
   Begin VB.ComboBox CmbMes2 
      Height          =   315
      ItemData        =   "FrmRemunerativos.frx":030A
      Left            =   5320
      List            =   "FrmRemunerativos.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ComboBox Cmbmes1 
      Height          =   315
      ItemData        =   "FrmRemunerativos.frx":039A
      Left            =   2280
      List            =   "FrmRemunerativos.frx":03C2
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CheckBox Chk 
      Caption         =   "Formato Alternativo"
      Height          =   200
      Left            =   5280
      TabIndex        =   13
      Top             =   1335
      Width           =   1680
   End
   Begin MSDataListLib.DataCombo CboTrabajador 
      Height          =   315
      Left            =   1425
      TabIndex        =   3
      Top             =   870
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.ComboBox CboTipo_Trab 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   495
      Width           =   5550
   End
   Begin VB.TextBox Txt_Year_Finish 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4515
      MaxLength       =   4
      TabIndex        =   1
      Top             =   120
      Width           =   525
   End
   Begin VB.TextBox Txt_Year_Star 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   0
      Top             =   120
      Width           =   525
   End
   Begin MSComctlLib.ProgressBar pbCadebecera 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   10
      Top             =   1770
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar pbDetalle 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   11
      Top             =   1665
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSForms.CommandButton btn 
      Height          =   375
      Left            =   1425
      TabIndex        =   14
      Top             =   1245
      Width           =   1515
      Caption         =   "     Procesar"
      PicturePosition =   327683
      Size            =   "2672;661"
      Picture         =   "FrmRemunerativos.frx":042A
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
      TabIndex        =   12
      Top             =   1245
      Width           =   1200
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      Height          =   195
      Index           =   3
      Left            =   3990
      TabIndex        =   9
      Top             =   160
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   870
      Width           =   765
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Trabajador"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   495
      Width           =   1125
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   465
   End
   Begin MSForms.SpinButton Sb_Year_Finish 
      Height          =   315
      Left            =   5040
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
      Size            =   "450;556"
   End
   Begin MSForms.SpinButton Sb_Year_Star 
      Height          =   315
      Left            =   1935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
      Size            =   "450;556"
   End
End
Attribute VB_Name = "FrmRemunerativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cadena      As String
Dim mTipo       As String
Dim Ors         As ADODB.Recordset
Public lRepor As String

Private Sub btn_Click()

Me.pbCadebecera.Value = 0
Me.pbDetalle.Value = 0
Me.lbl(4).Caption = Empty
Me.lbl(4).Caption = "Iniciando el Proceso..."
Me.lbl(4).Refresh
If Len(Me.Txt_Year_Star.Text) = 0 Or Len(Me.Txt_Year_Finish.Text) = 0 Then MsgBox "Error de Usuario: Verificar la Fecha.": Exit Sub
If Val(Me.Txt_Year_Finish.Text) < Val(Me.Txt_Year_Star.Text) Then MsgBox "Error de Usuario: Verificar la Fecha.": Exit Sub
If mTipo = "99" Or mTipo = "999" Or mTipo = Empty Then MsgBox "Error de Usuario: Debe de seleccionar un tipo de empleado.": Exit Sub
If Len(Me.CboTrabajador.BoundText) > 6 Then MsgBox "Error de Usuario: Debe de seleccion un empleado.": Exit Sub
If Not Trae_Informacion Then MsgBox "No se han encontrado resultados.": Exit Sub
    
If lRepor = "INGRESOS" Then
   If Trim(Cmbmes1.Text & "") = "" Or Trim(CmbMes2.Text & "") = "" Then MsgBox "Error de Usuario: Verificar la Fecha.": Exit Sub
   Call Exp_Excel_Ingresos
Else
    Call Exp_Excel
End If
End Sub

Private Sub CboTipo_Trab_Click()
    mTipo = Empty
    mTipo = Trim(fc_CodigoComboBox(Me.CboTipo_Trab, 2))
    Call Trae_Trabajador
End Sub

Private Sub Trae_Tipo_Trab()
    Cadena = "SP_TRAE_TIPO_TRAB"
    Call rCarCbo(Me.CboTipo_Trab, Cadena, "XX", "00")
    Me.CboTipo_Trab.ListIndex = 0
    
If wTipoPla <> "99" And UCase(wuser) <> "SA" Then
   If wTipoPla = "" Then
      Call rUbiIndCmbBox(CboTipo_Trab, "02", "00")
   Else
      Call rUbiIndCmbBox(CboTipo_Trab, Trim(wTipoPla), "00")
   End If
   CboTipo_Trab.Enabled = False
ElseIf UCase(wuser) = "RRHH01" Then
   Call rUbiIndCmbBox(CboTipo_Trab, Trim("02"), "00")
   CboTipo_Trab.Enabled = False
End If
End Sub

Private Sub Trae_Trabajador()
    mTipo = Trim(fc_CodigoComboBox(Me.CboTipo_Trab, 2))
    If mTipo = "99" Then mTipo = Empty
    Cadena = "SP_TRAE_TRABAJADOR '" & wcia & "'," & _
            "'" & Trim(mTipo) & "'"
    Set Ors = New ADODB.Recordset
    Ors.CursorLocation = adUseClient
    Ors.Open Cadena, cn, adOpenStatic, adLockReadOnly
    If Not Ors.EOF Then
        With Me.CboTrabajador
            Set .RowSource = Ors
            .ListField = "NOMBRE"
            .DataField = "NOMBRE"
            .BoundColumn = "CODIGO"
        End With
    End If
    Me.CboTrabajador.Refresh
    If mTipo = Empty Then Me.CboTrabajador.BoundText = "99999"
End Sub

'Private Sub Up(Ctrl As Object)
'    If Trim(Ctrl.Text) = "" Then
'       Ctrl.Text = Format(Year(Date), "0000")
'    Else
'       Ctrl.Text = Ctrl.Text + 1
'    End If
'End Sub
'
'Private Sub Down(Ctrl As Object)
'    If Trim(Ctrl.Text) = "" Then
'       Ctrl.Text = Format(Year(Date), "0000")
'    Else
'       Ctrl.Text = Ctrl.Text - 1
'    End If
'End Sub

Private Sub Form_Load()
    Call Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Ors Is Nothing Then
        If Ors.State = adStateOpen Then Ors.Close
        Set Ors = Nothing
    End If
End Sub

Private Sub Sb_Year_Finish_SpinDown()
    Call Down(Me.Txt_Year_Finish)
End Sub

Private Sub Sb_Year_Finish_SpinUp()
    Call Up(Me.Txt_Year_Finish)
End Sub

Private Sub Sb_Year_Star_SpinDown()
    Call Down(Me.Txt_Year_Star)
End Sub

Private Sub Sb_Year_Star_SpinUp()
    Call Up(Me.Txt_Year_Star)
End Sub

Private Sub Load()
    Call Trae_Tipo_Trab
    Call Trae_Trabajador
    Me.Txt_Year_Star.Text = Empty
    Me.Txt_Year_Finish.Text = Empty
    lbl(4).Caption = Empty
    Me.pbCadebecera.Value = 0
    Me.pbDetalle.Value = 0
    Me.Top = 0
    Me.Left = 0
    If lRepor = "INGRESOS" Then
       Cmbmes1.Visible = True
       CmbMes2.Visible = True
       Chk.Visible = False
       Txt_Year_Star.Text = Year(Date)
       Txt_Year_Finish.Text = Year(Date)
       Cmbmes1.ListIndex = Month(Date) - 1
       CmbMes2.ListIndex = Month(Date) - 1
    Else
       Cmbmes1.Visible = False
       CmbMes2.Visible = False
       Chk.Visible = True
    End If
End Sub

'Private Sub NumberOnly(ByRef Key As Integer)
'    If InStr("0123456789" & Chr(8), Chr(Key)) = 0 Then Key = 0
'End Sub

Private Sub Txt_Year_Finish_KeyPress(KeyAscii As Integer)
    Call NumberOnly(KeyAscii)
End Sub

Private Sub Txt_Year_Star_KeyPress(KeyAscii As Integer)
    Call NumberOnly(KeyAscii)
End Sub

Private Function Trae_Informacion() As Boolean
    Screen.MousePointer = vbHourglass
    Trae_Informacion = False
    Me.lbl(4).Refresh
    Me.lbl(4).Caption = "Espere por favor...."
    Me.lbl(4).Refresh
    Cadena = Empty
    Cadena = "EXEC SP_TRAE_CONCEPTO_REMUNERATIVO " & _
            "'" & wcia & "'," & _
            "'" & mTipo & "'," & _
            "" & CInt(Me.Txt_Year_Star.Text) & "," & _
            "" & CInt(Me.Txt_Year_Finish.Text) & "," & _
            "'" & Trim(Me.CboTrabajador.BoundText) & "'," & _
            "" & IIf(Chk.Value = Checked, 1, 0) & ""
    If (Not EXEC_SQL(Cadena, cn)) Then Screen.MousePointer = vbDefault: Exit Function
    Trae_Informacion = True
    Screen.MousePointer = vbDefault
End Function
Private Sub Exp_Excel()
    'NOTA : PARA ESTE REPORTE SE ASUMIO QUE SIEMPRE SERAN 14 COLUMNAS
    '       CONCEPTO + PERIODO(AÑO) + LOS 12 MESES.(ES DECIR NO SON DINAMICAS)
    Dim ObjExcel    As Object
    Dim Columna     As Integer
    Dim Fila        As Integer
    Dim rsTmp       As ADODB.Recordset
    Dim rsTemp      As ADODB.Recordset
    Dim Periodo     As String
    Screen.MousePointer = vbHourglass
    Me.lbl(4).Caption = "Procesando..."
    Cadena = ""
    Cadena = "SP_LISTADO_VARIOS " & 1 & ""
    Set rsTmp = OpenRecordset(Cadena, cn)
    If Not rsTmp.EOF Then
        Me.pbCadebecera.Min = 0: Me.pbCadebecera.Max = rsTmp.RecordCount
        'ARMAMOS LA CABECERA DE LA HORA EN EXCEL
        If Not ObjExcel Is Nothing Then
            ObjExcel.Close False
            Set ObjExcel = Nothing
        End If
        
        If Txt_Year_Star.Text = Txt_Year_Finish Then Periodo = Me.Txt_Year_Star.Text Else Periodo = Me.Txt_Year_Star.Text & " A " & Me.Txt_Year_Finish.Text
        Set ObjExcel = CreateObject("Excel.Application")
        ObjExcel.Workbooks.Add
        ObjExcel.Application.StandardFont = "Arial"
        ObjExcel.Application.StandardFontSize = "8"
        ObjExcel.ActiveWorkbook.Sheets(1).Name = Me.CboTrabajador.BoundText & "_" & Periodo
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
        ObjExcel.ActiveSheet.Protect DrawingObjects:=False, Contents:=False, Scenarios:=True, Password:="rodasa"
        With ObjExcel.ActiveSheet.PageSetup
            .LeftMargin = .Application.InchesToPoints(0)
            .RightMargin = .Application.InchesToPoints(0)
            .TopMargin = ObjExcel.Application.InchesToPoints(0.393700787401575)
            .BottomMargin = ObjExcel.Application.InchesToPoints(0.590551181102362)
            .Zoom = 100
        End With
        Fila = 1
        With ObjExcel
            .Range("A1:N1").Font.Bold = True
            .Range("A2:N2").Font.Bold = True
            .Range("A4:N4").Font.Bold = True
            .Range("A" & Fila & ":A" & Fila).Value = "PERIODO A INFORMAR : " & Periodo
            .Range("A" & Fila & ":N" & Fila).Merge
            .Range("A" & Fila & ":N" & Fila).HorizontalAlignment = xlCenter
            Fila = Fila + 1 '=2
            .Range("A" & Fila & ":A" & Fila).Value = "DETALLE DE LOS CONCEPTOS REMUNERATIVOS QUE SE CONSIGNAN EN PLANILLA"
            .Range("A" & Fila & ":N" & Fila).Merge
            .Range("A" & Fila & ":N" & Fila).HorizontalAlignment = xlCenter
            Fila = Fila + 2 '=4
            .Range("A" & Fila & ":A" & Fila).Value = Me.CboTrabajador.Text
            .Range("A" & Fila & ":N" & Fila).Merge
            If Chk.Value = Checked Then
                Fila = Fila + 1
                Cadena = "SP_TRAE_CENTRO_COSTO '" & wcia & "'," & _
                        "'" & Trim(Me.CboTrabajador.BoundText) & "'"
                Dim rsAux As ADODB.Recordset
                Set rsAux = OpenRecordset(Cadena, cn)
                If (Not rsAux.EOF) Then
                    .Range("A" & Fila & ":A" & Fila).Value = "CENTRO DE COSTO : " & IIf(CStr(rsAux!CENCOSTO) = "", "No Indicado", CStr(rsAux!CENCOSTO))
                    .Range("A" & Fila & ":N" & Fila).Merge
                End If
                rsAux.Close
                Set rsAux = Nothing
            End If
        End With
        Fila = Fila + 2 '=6
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Call Armado_dCabezeras(ObjExcel, Fila)
            Cadena = "SP_LISTADO_VARIOS " & 2 & "," & _
                    "" & CInt(rsTmp!Periodo) & "," & _
                    "'" & Trim(Me.CboTrabajador.BoundText) & "'"
            Set rsTemp = OpenRecordset(Cadena, cn)
            If (Not rsTemp.EOF) Then
                Call Armado_dDetalle(ObjExcel, Fila, rsTemp)
            End If
            Me.pbCadebecera.Value = Me.pbCadebecera.Value + 1
            Fila = Fila + 1
            rsTmp.MoveNext
        Loop
        Call Formato_dColumna(ObjExcel)
        Me.lbl(4).Caption = "Proceso Terminado..."
        MsgBox "Proceso Terminado con Exito.", vbInformation, "Detalle de Conceptos Remunerativos"
        Me.lbl(4).Caption = Empty
        ObjExcel.Visible = True
        Set ObjExcel = Nothing
    Else
        Me.lbl(4).Caption = "No hay Resultados..."
        MsgBox "No se ha encontrado información para el Trabajador : " & vbCrLf & Me.CboTrabajador.BoundText & " " & Me.CboTrabajador.Text, vbInformation + vbOKOnly, "Detalle de Conceptos Remunerativos"
        Me.lbl(4).Caption = Empty
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Armado_dCabezeras(ByRef oBj As Object, ByRef Fila As Integer)
    Dim Columna As Integer
    Dim I As Integer
    Columna = 1
    With oBj
        .Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = "CONCEPTO"
        Columna = Columna + 1
        .Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = "AÑO"
        For I = 1 To 12
            .Range(m_LetraColumna(Columna + I) & Fila & ":" & m_LetraColumna(Columna + I) & Fila).Value = mes_palabras2(I)
        Next I
        .Range("A" & Fila & ":N" & Fila).HorizontalAlignment = xlCenter
        .Range("A" & Fila & ":N" & Fila).Font.Bold = True
    End With
    Fila = Fila + 1
End Sub

Private Sub Armado_dDetalle(ByRef oBj As Object, ByRef Fila As Integer, ByRef mRecordSet As ADODB.Recordset)
    Dim Columna As Integer
    Const Formato   As String = "#,##0.00"
    mRecordSet.MoveFirst
    Me.pbDetalle.Min = 0: Me.pbDetalle.Max = mRecordSet.RecordCount
    Me.pbDetalle.Value = 0
    Do While Not mRecordSet.EOF
        Columna = 1
        For I = 0 To mRecordSet.Fields.count - 1
            With oBj
                .Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).Value = IIf(mRecordSet(I) = 0, "-", mRecordSet(I))
                If Columna > 2 Then .Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).HorizontalAlignment = xlRight
                If Columna = 2 Then .Range(m_LetraColumna(Columna) & Fila & ":" & m_LetraColumna(Columna) & Fila).HorizontalAlignment = xlCenter
            End With
            Columna = Columna + 1
        Next I
        Fila = Fila + 1
        Me.pbDetalle.Value = Me.pbDetalle.Value + 1
        mRecordSet.MoveNext
    Loop
    With oBj
        If Chk.Value = Checked Then
            .Range(m_LetraColumna(3) & Fila & ":" & m_LetraColumna(14) & Fila).Value = "=SUM(R[-" & mRecordSet.RecordCount & "]C:R[-1]C)"
            .Range(m_LetraColumna(3) & Fila & ":" & m_LetraColumna(14) & Fila).Select
            .Selection.Font.Bold = True
            Fila = Fila + 1
        End If
        .Range(m_LetraColumna(1) & 1 & ":" & m_LetraColumna(14) & Fila).Select
    End With
End Sub

Private Sub Formato_dColumna(ByRef oBj As Object)
    With oBj
        For I = 1 To 14
            .Columns(m_LetraColumna(I) & ":" & m_LetraColumna(I)).NumberFormat = "##0.00"
            .Columns(m_LetraColumna(I) & ":" & m_LetraColumna(I)).EntireColumn.AutoFit
        Next I
    End With
End Sub

Private Sub Exp_Excel_Ingresos()
    Dim ObjExcel    As Object
    Dim Columna     As Integer
    Dim Fila        As Integer
    Dim rsTmp       As ADODB.Recordset
    Dim rsTemp      As ADODB.Recordset
    Dim Periodo1     As String
    Dim Periodo2     As String
    Dim I As Integer
    Dim lCol As Integer
    Dim msum As Integer
    
    msum = 0
    Periodo1 = Format(Cmbmes1.ListIndex + 1, "00") + Txt_Year_Star.Text
    Periodo2 = Format(CmbMes2.ListIndex + 1, "00") + Txt_Year_Finish.Text
    
    Screen.MousePointer = vbHourglass
    Me.lbl(4).Caption = "Procesando..."
    
    Cadena = ""
    If chkRemunra.Value = 1 Then
        Cadena = "uSp_Pla_Detelle_Ingresos_Remunera '" & wcia & "','" & Periodo1 & "','" & Periodo2 & "','" & Trim(Me.CboTrabajador.BoundText) & "'"
    Else
        Cadena = "uSp_Pla_Detelle_Ingresos '" & wcia & "','" & Periodo1 & "','" & Periodo2 & "','" & Trim(Me.CboTrabajador.BoundText) & "'"
    End If

    Set rsTmp = OpenRecordset(Cadena, cn)
    If Not rsTmp.EOF Then
        Me.pbCadebecera.Min = 0: Me.pbCadebecera.Max = rsTmp.RecordCount
        If Not ObjExcel Is Nothing Then
            ObjExcel.Close False
            Set ObjExcel = Nothing
        End If
        
        Set ObjExcel = CreateObject("Excel.Application")
        ObjExcel.Workbooks.Add
        ObjExcel.Application.StandardFont = "Arial"
        ObjExcel.Application.StandardFontSize = "8"
        ObjExcel.ActiveWorkbook.Sheets(1).Name = Me.CboTrabajador.BoundText
        
        ObjExcel.ActiveWorkbook.Sheets(ObjExcel.ActiveWorkbook.Sheets(1).Name).Activate
        With ObjExcel.ActiveSheet.PageSetup
            .LeftMargin = .Application.InchesToPoints(0)
            .RightMargin = .Application.InchesToPoints(0)
            .TopMargin = ObjExcel.Application.InchesToPoints(0.393700787401575)
            .BottomMargin = ObjExcel.Application.InchesToPoints(0.590551181102362)
            .Zoom = 100
        End With
        Fila = 1
        With ObjExcel
            Fila = Fila + 1 '=2
            .Range("A" & Fila & ":A" & Fila).Value = "DETALLE DE LOS INGRESOS QUE SE CONSIGNAN EN PLANILLA"
            Fila = Fila + 2 '=4
            .Range("A" & Fila & ":A" & Fila).Value = Me.CboTrabajador.Text
            .Range("A" & Fila & ":A" & Fila).Font.Bold = True
            Fila = Fila + 2
            rsTmp.MoveFirst
            Columna = 2
            lCol = 0
            .Cells(Fila, 1).Value = "CONCEPTO"
            Do While Not rsTmp.EOF
               .Cells(Fila, Columna).Value = rsTmp!Periodo
               Columna = Columna + 1
               lCol = lCol + 1
               rsTmp.MoveNext
            Loop
            Fila = Fila + 1
            Dim lConcepto As Boolean
            lConcepto = True
            For I = 0 To rsTmp.Fields.count - 1
                Columna = 1
                If Left(rsTmp.Fields(I).Name, 1) = "I" Then
                   .Cells(Fila, Columna).Value = rsTmp.Fields(I).Name
                   rsTmp.MoveFirst

                   Do While Not rsTmp.EOF
                      If lConcepto Then
                      Cadena = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno='" & Mid(rsTmp.Fields(I).Name, 2, 2) & "' and status<>'*'"
                      Dim rsAux As ADODB.Recordset
                      Set rsAux = OpenRecordset(Cadena, cn)
                      If (Not rsAux.EOF) Then
                         If Mid(rsTmp.Fields(I).Name, 2, 2) = "10" Then
                            .Cells(Fila, 1).Value = "HORAS EXTRAS"
                         Else
                            .Cells(Fila, 1).Value = Trim(rsAux(0) & "")
                         End If
                      End If
                      rsAux.Close
                      msum = msum + 1
                      lConcepto = False
                      End If
                      Set rsAux = Nothing
                      .Cells(Fila, Columna + 1).Value = rsTmp.Fields(I)
                      Columna = Columna + 1
                      rsTmp.MoveNext
                   Loop
                   lConcepto = True
                   Fila = Fila + 1
                   Columna = Columna + 1
                End If
               Next
        End With
        
        msum = msum * -1
        ObjExcel.Cells(Fila, 1).Value = "TOTALES"
        ObjExcel.Cells(Fila, 1).HorizontalAlignment = xlCenter
        For I = 2 To lCol + 1
               ObjExcel.Cells(Fila, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
        Next
        ObjExcel.Range(ObjExcel.ActiveSheet.Cells(Fila, 1), ObjExcel.ActiveSheet.Cells(Fila, lCol + 1)).HorizontalAlignment = xlCenter
        ObjExcel.Range(ObjExcel.ActiveSheet.Cells(Fila, 1), ObjExcel.ActiveSheet.Cells(Fila, lCol + 1)).Borders.LineStyle = xlContinuous
        ObjExcel.Range(ObjExcel.ActiveSheet.Cells(Fila, 1), ObjExcel.ActiveSheet.Cells(Fila, lCol + 1)).Font.Bold = True


        Me.lbl(4).Caption = "Proceso Terminado..."
        MsgBox "Proceso Terminado con Exito.", vbInformation, "Detalle de Ingresos"
        Me.lbl(4).Caption = Empty
        ObjExcel.Application.ActiveWindow.DisplayGridLines = False
        ObjExcel.Range("A:A").ColumnWidth = 33
        
        ObjExcel.Range("B:Z").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
        
        ObjExcel.Range(ObjExcel.ActiveSheet.Cells(6, 1), ObjExcel.ActiveSheet.Cells(6, lCol + 1)).HorizontalAlignment = xlCenter
        ObjExcel.Range(ObjExcel.ActiveSheet.Cells(6, 1), ObjExcel.ActiveSheet.Cells(6, lCol + 1)).Borders.LineStyle = xlContinuous
        ObjExcel.Range(ObjExcel.ActiveSheet.Cells(6, 1), ObjExcel.ActiveSheet.Cells(6, lCol + 1)).Font.Bold = True

        ObjExcel.Visible = True
        Set ObjExcel = Nothing
    Else
        Me.lbl(4).Caption = "No hay Resultados..."
        MsgBox "No se ha encontrado información para el Trabajador : " & vbCrLf & Me.CboTrabajador.BoundText & " " & Me.CboTrabajador.Text, vbInformation + vbOKOnly, "Detalle de Conceptos Remunerativos"
        Me.lbl(4).Caption = Empty
    End If
    Screen.MousePointer = vbDefault
End Sub


