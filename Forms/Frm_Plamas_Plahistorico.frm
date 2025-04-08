VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frm_Plamas_Plahistorico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Plamas - Plahistorico «"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "Frm_Plamas_Plahistorico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_Year 
      Height          =   315
      Left            =   1155
      TabIndex        =   2
      Top             =   975
      Width           =   615
   End
   Begin VB.ComboBox Cbo_Month 
      Height          =   315
      ItemData        =   "Frm_Plamas_Plahistorico.frx":030A
      Left            =   1155
      List            =   "Frm_Plamas_Plahistorico.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1395
      Width           =   2385
   End
   Begin VB.ComboBox Cbo_Formato 
      Height          =   315
      ItemData        =   "Frm_Plamas_Plahistorico.frx":039A
      Left            =   1155
      List            =   "Frm_Plamas_Plahistorico.frx":03A4
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   165
      Width           =   2310
   End
   Begin VB.ComboBox Cbo_Cia 
      Height          =   315
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   570
      Width           =   4830
   End
   Begin MSForms.CommandButton btn_Salir 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1950
      Width           =   1320
      Caption         =   "     Salir"
      PicturePosition =   327683
      Size            =   "2328;661"
      Picture         =   "Frm_Plamas_Plahistorico.frx":03BE
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton btn_Exportar 
      Height          =   375
      Left            =   3375
      TabIndex        =   4
      Top             =   1950
      Width           =   1320
      Caption         =   "     Exportar"
      PicturePosition =   327683
      Size            =   "2328;661"
      Picture         =   "Frm_Plamas_Plahistorico.frx":0958
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.SpinButton sp_Year 
      Height          =   315
      Left            =   1785
      TabIndex        =   10
      Top             =   975
      Width           =   255
      Size            =   "450;556"
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
      Height          =   195
      Index           =   3
      Left            =   225
      TabIndex        =   9
      Top             =   1395
      Width           =   300
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
      Height          =   195
      Index           =   2
      Left            =   225
      TabIndex        =   8
      Top             =   975
      Width           =   285
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compañia"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   7
      Top             =   570
      Width           =   705
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Formato"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   165
      Width           =   570
   End
End
Attribute VB_Name = "Frm_Plamas_Plahistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pCia As String

Private Sub Load_Cia()
    Cadena = "SP_TRAE_CIA"
    Call rCarCbo(Cbo_Cia, Cadena, "XX", "00")
End Sub

'Private Function Exportar_Excel(ByRef mRecordSet As ADODB.Recordset) As Boolean
'
'    On Error GoTo errSub
'
'    Dim cn          As New ADODB.Connection
'    Dim rec         As New ADODB.Recordset
'    Dim Excel       As Object
'    Dim Libro       As Object
'    Dim Hoja        As Object
'    Dim arrData     As Variant
'    Dim iRec        As Long
'    Dim iCol        As Integer
'    Dim iRow        As Integer
'
'    Set Excel = CreateObject("Excel.Application")
'    Set Libro = Excel.Workbooks.Add
'    Set Hoja = Libro.Worksheets(1)
'    Excel.Visible = True: Excel.UserControl = True
'    iCol = mRecordSet.Fields.count
'    For iCol = 1 To mRecordSet.Fields.count
'        Hoja.Cells(1, iCol).Value = mRecordSet.Fields(iCol - 1).Name
'    Next
'
'    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
'        Hoja.Cells(2, 1).CopyFromRecordset mRecordSet
'    Else
'        arrData = mRecordSet.GetRows
'        iRec = UBound(arrData, 2) + 1
'        For iCol = 0 To mRecordSet.Fields.count - 1
'            For iRow = 0 To iRec - 1
'                If IsDate(arrData(iCol, iRow)) Then
'                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
'
'                ElseIf IsArray(arrData(iCol, iRow)) Then
'                    arrData(iCol, iRow) = "Array Field"
'                End If
'            Next iRow
'        Next iCol
'        Hoja.Cells(2, 1).Resize(iRec, mRecordSet.Fields.count).Value = GetData(arrData)
'    End If
'
'    Excel.Selection.CurrentRegion.Columns.AutoFit
'    Excel.Selection.CurrentRegion.Rows.AutoFit
'
'    If Not mRecordSet Is Nothing Then
'        If mRecordSet.State = adStateOpen Then mRecordSet.Close
'        Set mRecordSet = Nothing
'    End If
'
'    'Libro.SaveAs sOutputPathXLS
'    'Libro.Close
'
'    Set Hoja = Nothing
'    Set Libro = Nothing
'    Excel.Visible = True
'    'Set Excel = Nothing
'
'    Exportar_Excel = True
'    Exit Function
'errSub:
'    MsgBox Err.Description, vbCritical, "Error"
'    Exportar_Excel = False
'    Me.Enabled = True
'End Function
'
'Private Function GetData(vValue As Variant) As Variant
'    Dim x As Long, y As Long, xMax As Long, yMax As Long, T As Variant
'
'    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
'
'    ReDim T(xMax, yMax)
'    For x = 0 To xMax
'        For y = 0 To yMax
'            T(x, y) = vValue(y, x)
'        Next y
'    Next x
'    GetData = T
'End Function

Private Sub btn_Exportar_Click()
    If wGrupoPla = "01" Then
       Exporta_Detalle_Planilla
    End If
    If pCia = "" Then Exit Sub
    Select Case Cbo_Formato.ListIndex
        Case 0
            Cadena = "SP_TRAE_MOV_PLAHISTORICO '" & pCia & "', " & CInt(txt_Year.Text) & ", " & Cbo_Month.ListIndex + 1 & ""
        Case 1
            Cadena = "SP_MOV_PLAMAS '" & pCia & "'"
        Case Else
            Exit Sub
    End Select
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = OpenRecordset(Cadena, cn)
    If Not rsTmp.EOF Then
        Call Exportar_Excel(rsTmp)
    Else
        MsgBox "No se encontraron datos para mostrar.", vbExclamation + vbOKOnly, Me.Caption
    End If
End Sub

Private Sub Seleccion(ByVal Opcion As Integer)
    txt_Year.Text = Empty
    Cbo_Month.ListIndex = -1
    Select Case Opcion
        Case 0
            txt_Year.Enabled = True
            Cbo_Month.Enabled = True
            sp_Year.Enabled = True
        Case 1
            txt_Year.Enabled = False
            Cbo_Month.Enabled = False
            sp_Year.Enabled = False
    End Select
End Sub

Private Sub btn_salir_Click()
    Unload Me
End Sub

Private Sub Cbo_Cia_Click()
    pCia = Empty
    pCia = Trim(fc_CodigoComboBox(Cbo_Cia, 2))
End Sub

Private Sub Cbo_Formato_Click()
    If Cbo_Formato.ListIndex <> -1 Then Call Seleccion(Cbo_Formato.ListIndex)
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call Load_Cia
    If wGrupoPla = "01" Then
       Cbo_Formato.Visible = False
       lbl(0).Visible = False
       Cbo_Cia.Visible = False
       lbl(1).Visible = False
       Cbo_Month.Visible = False
       lbl(3).Visible = False
    End If
End Sub

Private Sub sp_Year_SpinDown()
    Call Down(txt_Year)
End Sub

Private Sub sp_Year_SpinUp()
    Call Up(txt_Year)
End Sub
Private Sub Exporta_Detalle_Planilla()
Sql$ = "uSp_Detalle_Planilla " & txt_Year.Text & ""
If Not (fAbrRst(rs, Sql$)) Then
   MsgBox "No Hay Registros", vbInformation
   rs.Close
   Exit Sub
End If
rs.MoveFirst
Dim I As Integer
I = 0
Do While Not rs.EOF
  For I = 0 To rs.Fields.count - 1
      MsgBox (rs.Fields(I).Name)
  Next
  rs.MoveNext
Loop
End Sub
