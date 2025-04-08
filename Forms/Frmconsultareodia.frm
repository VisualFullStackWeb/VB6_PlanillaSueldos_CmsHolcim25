VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmconsultareodia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Detallada de Tareo"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Frmconsultareodia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6375
      Begin VB.Frame FrameGenera 
         Height          =   735
         Left            =   4680
         TabIndex        =   12
         Top             =   560
         Visible         =   0   'False
         Width           =   1575
         Begin VB.CommandButton Command3 
            Caption         =   "Generar"
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
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "LISTADO"
         Height          =   255
         Left            =   4680
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TAREO SEMANAL"
         Height          =   255
         Left            =   4680
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker FecFin 
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   12615808
         CalendarTitleForeColor=   16777215
         Format          =   63111169
         CurrentDate     =   37616
      End
      Begin MSComCtl2.DTPicker FecIni 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   12615808
         CalendarTitleForeColor=   16777215
         Format          =   63111169
         CurrentDate     =   37616
      End
      Begin VB.ComboBox Cmbtipotra 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "al"
         Height          =   195
         Left            =   2760
         TabIndex        =   9
         Top             =   720
         Width           =   120
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.Trabajador"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   5295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   840
      End
   End
End
Attribute VB_Name = "Frmconsultareodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VCCosto As String
Dim VTipotrab As String
Dim VConcepto As String
Public lRepor As String
Private Sub Cmbccosto_Click()
If CmbCcosto.Text = "TOTAL" Then
   VCCosto = ""
Else
   VCCosto = fc_CodigoComboBox(CmbCcosto, 2)
End If
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2(wcia & "055", "", Cmbtipotra)
'Cmbtipotra.AddItem "TOTAL"
'Call fc_Descrip_Maestros2(wcia & "044", "", Cmbccosto)
'Cmbccosto.AddItem "TOTAL"
'Call fc_Descrip_Maestros2(wcia & "077", "", Cmbconcepto)
'Cmbconcepto.AddItem "TOTAL"
End Sub

Private Sub Cmbconcepto_Click()
If CmbConcepto.Text = "TOTAL" Then
   VConcepto = ""
Else
   VConcepto = fc_CodigoComboBox(CmbConcepto, 2)
End If
End Sub

Private Sub Cmbtipotra_Click()
If Cmbtipotra.Text = "TOTAL" Then
   VTipotrab = ""
Else
   VTipotrab = fc_CodigoComboBox(Cmbtipotra, 2)
End If
End Sub

Private Sub Command1_Click()
Procesa_TareoSemanal
End Sub

Private Sub Command2_Click()
Procesa_Listado
End Sub

Private Sub Command3_Click()
Dim f1 As String
Dim f2 As String
f1 = Format(Month(FecIni.Value), "00") & "/" & Format(Day(FecIni.Value), "00") & "/" & Format(Year(FecIni.Value), "0000")
f2 = Format(Month(FecFin.Value), "00") & "/" & Format(Day(FecFin.Value), "00") & "/" & Format(Year(FecFin.Value), "0000")
If lRepor = "MARCA" Then Call Marcaciones(wcia, f1, f2, fc_CodigoComboBox(Cmbtipotra, 2))
If lRepor = "TARDA" Then Call Tardanzas(wcia, f1, f2, fc_CodigoComboBox(Cmbtipotra, 2))
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 6750
Me.Height = 2595
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
FecIni.Value = Date
FecFin.Value = Date
End Sub

'Private Sub Txtobra_Change()
'If Len(Trim(Txtobra.Text)) < 8 Then Lblobra.Caption = ""
'End Sub
'
'Private Sub Txtobra_KeyPress(KeyAscii As Integer)
'Txtobra.Text = Txtobra.Text + fc_ValNumeros(KeyAscii)
'If KeyAscii = 13 And Trim(Txtobra.Text <> "") Then
'    Dim SQL As String
'    Dim Rq As ADODB.Recordset
'    Txtobra.Text = Format(Txtobra.Text, "00000000")
'    SQL = "select cod_obra,descrip from plaobras where cod_cia='" & wcia & "' and status<>'*' AND COD_OBRA='" & Txtobra.Text & "'"
'    Screen.MousePointer = 11
'    If Not fAbrRst(Rq, SQL) Then
'        ResaltarTexto Txtobra
'        MsgBox "Codigo de Obra No existe ", vbCritical, Me.Caption
'        GoTo TERMINA:
'    End If
'    Lblobra.Caption = Rq!descrip & ""
'
'TERMINA:
'    Screen.MousePointer = 0
'    Rq.Close
'    Set Rq = Nothing
'End If
'End Sub

Private Sub Txtreabajador_Change()

End Sub

'Private Sub Txtreabajador_KeyPress(KeyAscii As Integer)
'Txtrabajador.Text = Txtrabajador.Text + fc_ValNumeros(KeyAscii)
'
'
'End Sub
'
'Private Sub Txtrabajador_Change()
'If Len(Trim(Txtrabajador.Text)) < 8 Then Lbltrabajador.Caption = ""
'End Sub
'
'Private Sub Txtrabajador_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And Trim(Txtrabajador.Text <> "") Then
'    Dim SQL As String
'    Dim Rq As ADODB.Recordset
'    SQL = Nombre()
'    SQL = SQL + "placod from planillas where cia='" & wcia & "' and status<>'*' AND placod='" & Txtrabajador.Text & "'"
'    Screen.MousePointer = 11
'    If Not fAbrRst(Rq, SQL) Then
'        ResaltarTexto Txtrabajador
'        MsgBox "Codigo de Trabajador No existe ", vbCritical, Me.Caption
'        GoTo TERMINA:
'    End If
'    Lbltrabajador.Caption = Rq!Nombre & ""
'
'TERMINA:
'    Screen.MousePointer = 0
'    Rq.Close
'    Set Rq = Nothing
'End If
'
'End Sub

Public Sub Procesa_TareoSemanal()
Dim Sql As String
Dim Letras As String

If Cmbtipotra.ListIndex = -1 Then
    MsgBox "Elija tipo de trabajador", vbCritical, Me.Caption
    Exit Sub
End If

Dim xCodObra, xCC, xConcepto, xTipTra, xTipTra2, xCodTra As String
xTipTra = "": xTipTra2 = ""
If Cmbtipotra.Text <> "TOTAL" Then
    xTipTra = " AND P.TIPOTRABAJADOR='" & fc_CodigoComboBox(Cmbtipotra, 2) & "' "
    xTipTra2 = " and tipo_trab='" & fc_CodigoComboBox(Cmbtipotra, 2) & "' "
End If

Dim nFil As Integer
Dim nCol As Integer

Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
   
   Set xlApp1 = CreateObject("Excel.Application")
    xlApp1.Workbooks.Add
    Set xlApp2 = xlApp1.Application
    Set xlBook = xlApp2.Workbooks(1)
    Set xlSheet = xlApp2.Worksheets("HOJA1")
    'xlSheet.Name = "C" & NomCtaContable(wcia, vBco, Trim(vNroCta))
    xlSheet.Range("A:A").ColumnWidth = 1
    xlSheet.Range("B:B").ColumnWidth = 5
    xlSheet.Range("C:C").ColumnWidth = 10
    xlSheet.Range("D:D").ColumnWidth = 25
    xlSheet.Range("E:E").ColumnWidth = 15
    xlSheet.Range("F:F").ColumnWidth = 8
    
    
   Dim RZ As New ADODB.Recordset
   Sql$ = "select RAZSOC from cia where cod_cia='" & wcia & "' and status<>'*'"
   RZ.Open Sql$, cn, adOpenStatic, adLockReadOnly
   
'xlSheet.Cells(1, 2).Value = "CIA. MINERA AGREGADOS CALCAREOS S.A."
xlSheet.Cells(1, 2).Value = UCase(RZ("RAZSOC"))
   RZ.Close

xlSheet.Range("B1:N1").Merge
xlSheet.Range("B1:N1").HorizontalAlignment = xlLeft

xlSheet.Cells(2, 2).Value = "TAREO SEMANAL "
xlSheet.Cells(2, 2).Font.Underline = True
xlSheet.Cells(2, 2).Font.Size = 12
xlSheet.Cells(2, 2).Font.Bold = True
xlSheet.Range("B2:N2").Merge
xlSheet.Range("B2:N2").HorizontalAlignment = xlCenter

'xlSheet.Cells(3, 2).Value = "TAREO DE LA SEMANA N°"
xlSheet.Cells(4, 2).Value = "FECHA DEL " & Format(FecIni.Value, "DD/MM/YYYY") & " AL " & Format(FecFin.Value, "DD/MM/YYYY")
xlSheet.Cells(4, 12).Value = "TIEMPO EN HORAS"
xlSheet.Cells(5, 7).Value = "CONCEPTOS"

xlSheet.Cells(5, 2).Value = "N°"
xlSheet.Range("B5:B7").Merge
xlSheet.Range("B5:B7").HorizontalAlignment = xlCenter
xlSheet.Range("B5:B7").VerticalAlignment = xlCenter

xlSheet.Cells(5, 3).Value = "CODIGO"
xlSheet.Range("C5:C7").Merge
xlSheet.Range("C5:C7").HorizontalAlignment = xlCenter
xlSheet.Range("C5:C7").VerticalAlignment = xlCenter

xlSheet.Cells(5, 4).Value = "APELLIDOS Y NOMBRES"
xlSheet.Range("D5:D7").Merge
xlSheet.Range("D5:D7").HorizontalAlignment = xlCenter
xlSheet.Range("D5:D7").VerticalAlignment = xlCenter

xlSheet.Cells(5, 5).Value = "CARGO"
xlSheet.Range("E5:E7").Merge
xlSheet.Range("E5:E7").HorizontalAlignment = xlCenter
xlSheet.Range("E5:E7").VerticalAlignment = xlCenter

xlSheet.Cells(5, 6).Value = "H.TRAB."
xlSheet.Range("F5:F7").Merge
xlSheet.Range("F5:F7").HorizontalAlignment = xlCenter
xlSheet.Range("F5:F7").VerticalAlignment = xlCenter

Dim wciamae As String
Letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Dim Rc As ADODB.Recordset
wciamae = Determina_Maestro("01077")
Sql = "Select cod_maestro2,descrip from maestros_2 a,plaverhoras b where a.status<>'*' and b.status<>'*' and b.cia='" & wcia & "' " _
     & xTipTra2 & " and b.codigo=a.cod_maestro2 "
Sql = Sql & wciamae
nFil = 6
nCol = 7
Dim Li, Lf As String * 1
If fAbrRst(Rc, Sql) Then
    Do Until Rc.EOF
        xlSheet.Cells(nFil, nCol).Value = Rc!DESCRIP & ""
        Li = Mid(Letras, nCol, 1)
        xlSheet.Range(Li & nFil & ":" & Li & nFil + 1).ColumnWidth = 16
        xlSheet.Range(Li & nFil & ":" & Li & nFil + 1).Merge
        xlSheet.Range(Li & nFil & ":" & Li & nFil + 1).HorizontalAlignment = xlCenter
        xlSheet.Range(Li & nFil & ":" & Li & nFil + 1).VerticalAlignment = xlCenter
        xlSheet.Range(Li & nFil & ":" & Li & nFil + 1).WrapText = True
        xlSheet.Range(Li & nFil & ":" & Li & nFil + 1).ShrinkToFit = True
                
        nCol = nCol + 1
        Rc.MoveNext
    Loop
Else
    MsgBox "No existen conceptos para el Tareo", vbCritical, Me.Caption
    GoTo Termina
End If
 'Lf = Mid(Letras, nCol - 1, 1)
 Lf = Letra_ColumnaExcel(nCol - 1)
 
xlSheet.Range("G" & nFil - 1 & ":" & Lf & nFil - 1).Merge
xlSheet.Range("G" & nFil - 1 & ":" & Lf & nFil - 1).HorizontalAlignment = xlCenter

Dim Rq As ADODB.Recordset

Sql = nombre()
Sql = Sql & "P.PLACOD,P.CARGO,M.DESCRIP,P.TIPOTRABAJADOR FROM PLANILLAS P, MAESTROS_2 M WHERE P.CIA='" & wcia & "' AND P.STATUS<>'*' AND P.CARGO=M.COD_MAESTRO2 AND M.CIAMAESTRO='" & wcia & "008' AND (P.fcese is null OR P.fcese='') " _
& xTipTra & xCodTra
Screen.MousePointer = 11
If Not fAbrRst(Rq, Sql) Then
    Set xlSheet = Nothing
    Set xlApp1 = Nothing
    Set xlApp2 = Nothing
    MsgBox "No existen registros para el criterio especificado", vbCritical, Me.Caption
    GoTo Termina:
End If
Dim Rh As ADODB.Recordset
wciamae = Determina_Maestro("01055")
Sql = "Select cod_maestro2,flag2 from maestros_2 where status<>'*'"
Sql = Sql$ & wciamae
If fAbrRst(Rh, Sql) Then
End If

Dim xHoras As Integer
xHoras = 0
Dim Rm As ADODB.Recordset
wciamae = Determina_Maestro("01076")
Sql = "Select cod_maestro2,flag2 from maestros_2 where status<>'*'"
Sql = Sql$ & wciamae
If fAbrRst(Rm, Sql) Then xHoras = Val(Rm!flag2)
Rm.Close
Set Rm = Nothing

Dim I As Integer

nFil = 8
nCol = 2
I = 0
Do While Not Rq.EOF
    I = I + 1
    xlSheet.Cells(nFil, 2).Value = I
    xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
    xlSheet.Cells(nFil, 3).Value = Rq!PlaCod & ""
    xlSheet.Cells(nFil, 3).HorizontalAlignment = xlCenter
    xlSheet.Cells(nFil, 4).Value = Rq!nombre & ""
    xlSheet.Cells(nFil, 5).Value = Rq!DESCRIP & ""
    
    Rh.MoveFirst
    Do While Not Rh.EOF
        If Trim(Rq!TipoTrabajador) = Trim(Rh!COD_MAESTRO2) Then
            xlSheet.Cells(nFil, 6).Value = Rh!flag2 & ""
        End If
        Rh.MoveNext
    Loop
    Dim Rt As ADODB.Recordset
    Dim Campos As ADODB.Recordset
    
    Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    Sql = Sql & "select concepto,SUM(tiempo) AS TIEMPO,motivo," & _
    "CODIGOTRAB from platareo where cia='" & wcia & "' and fecha between '" & Format(FecIni.Value, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(FecFin.Value, FormatFecha) & Space(1) & FormatTimef & "' and status<>'*' " & _
    "group by codigotrab,concepto,motivo"
    
    If fAbrRst(Rt, Sql) Then
        Rt.MoveFirst
        Do While Not Rt.EOF
            Rc.MoveFirst
            nCol = 6
            Do While Not Rc.EOF
                nCol = nCol + 1
                If Rc.Fields("cod_maestro2") = Rt.Fields("concepto") And Rt.Fields("CODIGOTRAB") = xlSheet.Cells(nFil, 3).Value Then
                    Select Case Rt!motivo
                        Case "DI"
                            xlSheet.Cells(nFil, nCol).Value = Round(Rt!tiempo * xHoras, 2)
                        Case "HO"
                            xlSheet.Cells(nFil, nCol).Value = Rt!tiempo & ""
                        Case "MI"
                            xlSheet.Cells(nFil, nCol).Value = Round(Rt!tiempo / 60, 3)
                    End Select
                    With xlSheet.Cells(nFil, nCol).Interior
                        .ColorIndex = 36
                        .Pattern = xlSolid
                    End With
                End If
                Rc.MoveNext
            Loop
            Rt.MoveNext
        Loop
    End If
    Rt.Close
    Set Rt = Nothing
    nFil = nFil + 1
    Rq.MoveNext
Loop

'ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
xlSheet.Cells(nFil, 2).Value = "TOTALES"
xlSheet.Range("B" & nFil & ":F" & nFil).Merge
xlSheet.Range("B" & nFil & ":F" & nFil).HorizontalAlignment = xlCenter
xlSheet.Range("B" & nFil & ":F" & nFil).VerticalAlignment = xlCenter
Rc.MoveFirst
nCol = 6
Do While Not Rc.EOF
   nCol = nCol + 1
   xlSheet.Cells(nFil, nCol).FormulaR1C1 = "=SUM(R[-" & nFil - 7 & "]C:R[-1]C)"
   xlSheet.Cells(nFil, nCol).Font.Bold = True
   xlSheet.Cells(nFil, nCol).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
   Rc.MoveNext
Loop
With xlSheet.Range("B5:" & Lf & nFil)
    .Select
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone

    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
xlSheet.Cells(nFil + 1, 2).Value = "Fecha Emisión " & Format(Now, "DD/MM/YYYY hh:mm:ss ampm")

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "TAREO SEMANAL"
xlApp1.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True


Termina:
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = 0
Rc.Close
Set Rc = Nothing
Rq.Close
Set Rq = Nothing
End Sub

Public Function BuscarHoras()
wciamae = Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where status<>'*'"
Sql$ = Sql$ & wciamae

End Function

Public Sub Procesa_Listado()
Dim Sql As String
Dim Letras As String

If Cmbtipotra.ListIndex = -1 Then
    MsgBox "Elija tipo de trabajador", vbCritical, Me.Caption
    Exit Sub
End If

Dim xCodObra, xCC, xConcepto, xTipTra, xTipTra2, xCodTra As String
xTipTra = "": xTipTra2 = ""
If Cmbtipotra.Text <> "TOTAL" Then
    xTipTra = " AND P.TIPOTRABAJADOR='" & fc_CodigoComboBox(Cmbtipotra, 2) & "' "
    xTipTra2 = " and tipo_trab='" & fc_CodigoComboBox(Cmbtipotra, 2) & "' "
End If

Dim nFil As Integer
Dim nCol As Integer

Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
   
   Set xlApp1 = CreateObject("Excel.Application")
    xlApp1.Workbooks.Add
    Set xlApp2 = xlApp1.Application
    Set xlBook = xlApp2.Workbooks(1)
    Set xlSheet = xlApp2.Worksheets("HOJA1")
    'xlSheet.Name = "C" & NomCtaContable(wcia, vBco, Trim(vNroCta))
    xlSheet.Range("A:A").ColumnWidth = 1
    xlSheet.Range("B:B").ColumnWidth = 5
    xlSheet.Range("C:C").ColumnWidth = 10
    xlSheet.Range("D:D").ColumnWidth = 40
    xlSheet.Range("E:E").ColumnWidth = 25
    xlSheet.Range("F:F").ColumnWidth = 15
    xlSheet.Range("G:G").ColumnWidth = 15
    xlSheet.Range("H:H").ColumnWidth = 15
    xlSheet.Range("I:I").ColumnWidth = 10
    xlSheet.Range("J:J").ColumnWidth = 5
    xlSheet.Range("K:K").ColumnWidth = 12
    
'SEGA
xlSheet.Cells(1, 2).Value = "CIA. MINERA AGREGADOS CALCAREOS S.A."
xlSheet.Range("B1:K1").Merge
xlSheet.Range("B1:K1").HorizontalAlignment = xlLeft

xlSheet.Cells(2, 2).Value = "RELACION DE " & Cmbtipotra.Text
xlSheet.Cells(2, 2).Font.Underline = True
xlSheet.Cells(2, 2).Font.Size = 12
xlSheet.Cells(2, 2).Font.Bold = True
xlSheet.Range("B2:K2").Merge
xlSheet.Range("B2:K2").HorizontalAlignment = xlCenter

'xlSheet.Cells(3, 2).Value = "TAREO DE LA SEMANA N°"
xlSheet.Cells(4, 2).Value = "FECHA INGRESO DEL " & Format(FecIni.Value, "DD/MM/YYYY") & " AL " & Format(FecFin.Value, "DD/MM/YYYY")

xlSheet.Cells(5, 2).Value = "N°"
xlSheet.Range("B5:B6").Merge
xlSheet.Range("B5:B6").HorizontalAlignment = xlCenter
xlSheet.Range("B5:B6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 3).Value = "CODIGO"
xlSheet.Range("C5:C6").Merge
xlSheet.Range("C5:C6").HorizontalAlignment = xlCenter
xlSheet.Range("C5:C6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 4).Value = "APELLIDOS Y NOMBRES"
xlSheet.Range("D5:D6").Merge
xlSheet.Range("D5:D6").HorizontalAlignment = xlCenter
xlSheet.Range("D5:D6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 5).Value = "CARGO"
xlSheet.Range("E5:E6").Merge
xlSheet.Range("E5:E6").HorizontalAlignment = xlCenter
xlSheet.Range("E5:E6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 6).Value = "OCUPACION"
xlSheet.Range("F5:F6").Merge
xlSheet.Range("F5:F6").HorizontalAlignment = xlCenter
xlSheet.Range("F5:F6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 7).Value = "D.N.I."
xlSheet.Range("G5:G6").Merge
xlSheet.Range("G5:G6").HorizontalAlignment = xlCenter
xlSheet.Range("G5:G6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 8).Value = "CARNET IPSS"
xlSheet.Range("H5:H6").Merge
xlSheet.Range("H5:H6").HorizontalAlignment = xlCenter
xlSheet.Range("H5:H6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 9).Value = "F.INGRESO"
xlSheet.Range("I5:I6").Merge
xlSheet.Range("I5:I6").HorizontalAlignment = xlCenter
xlSheet.Range("I5:I6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 10).Value = "MON"
xlSheet.Range("J5:J6").Merge
xlSheet.Range("J5:J6").HorizontalAlignment = xlCenter
xlSheet.Range("J5:J6").VerticalAlignment = xlCenter

xlSheet.Cells(5, 11).Value = "BASICO"
xlSheet.Range("K5:K6").Merge
xlSheet.Range("K5:K6").HorizontalAlignment = xlCenter
xlSheet.Range("K5:K6").VerticalAlignment = xlCenter

Dim wciamae As String
Letras = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Dim Rq As ADODB.Recordset

Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql = Sql & nombre()
Sql = Sql & "P.PLACOD,P.CARGO,P.DNI,P.IPSS,P.FINGRESO,P.NIVELEDUCATIVO,M.DESCRIP,P.TIPOTRABAJADOR FROM PLANILLAS P, MAESTROS_2 M WHERE P.CIA='" & wcia & "' AND P.STATUS<>'*' AND P.CARGO=M.COD_MAESTRO2 AND M.CIAMAESTRO='" & wcia & "008' AND (P.fcese is null OR P.fcese='') and fingreso between '" & Format(FecIni.Value, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(FecFin.Value, FormatFecha) & Space(1) & FormatTimef & "' " _
& xTipTra & xCodTra & " ORDER BY NOMBRE"

Screen.MousePointer = 11
If Not fAbrRst(Rq, Sql) Then
    Set xlApp1 = Nothing
    Set xlApp2 = Nothing
    MsgBox "No existen registros para el criterio especificado", vbCritical, Me.Caption
    GoTo Termina:
End If

Dim I As Integer

nFil = 7
nCol = 2
I = 0
Do While Not Rq.EOF
    I = I + 1
    xlSheet.Cells(nFil, 2).Value = I
    xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter
    xlSheet.Cells(nFil, 3).Value = Rq!PlaCod & ""
    xlSheet.Cells(nFil, 3).HorizontalAlignment = xlCenter
    xlSheet.Cells(nFil, 4).Value = Rq!nombre & ""
    xlSheet.Cells(nFil, 5).Value = Rq!DESCRIP & ""
    'xlSheet.Cells(nFil, 6).Value = Rq!niveleducativo & ""
    xlSheet.Cells(nFil, 7).Value = Rq!DNI & ""
    xlSheet.Cells(nFil, 8).Value = Rq!ipss & ""
    xlSheet.Cells(nFil, 9).Value = Rq!fIngreso & ""
    
    Dim Rp As ADODB.Recordset
    Sql = "select moneda,importe from plaremunbase where cia='" & _
    wcia & "' and placod='" & Trim(Rq!PlaCod) & "' and concepto='01' " & _
    "and status<>'*'"
    If fAbrRst(Rp, Sql) Then
        xlSheet.Cells(nFil, 10).Value = Rp!moneda & ""
        xlSheet.Cells(nFil, 11).Value = Rp!importe & ""
        xlSheet.Cells(nFil, 11).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
    Else
        MsgBox "no existe sueldo del trabajador " & Rq!PlaCod & Space(1) & Rq!nombre, vbCritical, Me.Caption
    End If
    Rp.Close
    Set Rp = Nothing
    nFil = nFil + 1
    Rq.MoveNext
Loop

''ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
'xlSheet.Cells(nFil, 2).Value = "TOTALES"
'xlSheet.Range("B" & nFil & ":F" & nFil).Merge
'xlSheet.Range("B" & nFil & ":F" & nFil).HorizontalAlignment = xlCenter
'xlSheet.Range("B" & nFil & ":F" & nFil).VerticalAlignment = xlCenter
'Rc.MoveFirst
'nCol = 6
'Do While Not Rc.EOF
'   nCol = nCol + 1
'   xlSheet.Cells(nFil, nCol).FormulaR1C1 = "=SUM(R[-" & nFil - 7 & "]C:R[-1]C)"
'   xlSheet.Cells(nFil, nCol).Font.Bold = True
'   xlSheet.Cells(nFil, nCol).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
'   Rc.MoveNext
'Loop
With xlSheet.Range("B5:K" & nFil - 1)
    .Select
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone

    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End With
xlSheet.Cells(nFil, 2).Value = "Fecha Emisión " & Format(Now, "DD/MM/YYYY hh:mm:ss ampm")

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "LISTADO DE " & Trim(Cmbtipotra.Text)
xlApp1.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True


Termina:
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = 0
'Rc.Close
'Set Rc = Nothing
Rq.Close
Set Rq = Nothing
End Sub


