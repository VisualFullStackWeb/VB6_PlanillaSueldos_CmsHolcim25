VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmRptPlame 
   Caption         =   "Reporte de Vaciones -Subsidios - Faltas Comacsa"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4545
   LinkTopic       =   "FrmReporte"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdProceso 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker FecProceso2 
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   12615808
      CalendarTitleForeColor=   16777215
      Format          =   56426497
      CurrentDate     =   37616
   End
   Begin MSComCtl2.DTPicker FecProceso1 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   12615808
      CalendarTitleForeColor=   16777215
      Format          =   56426497
      CurrentDate     =   37616
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "FrmRptPlame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cadena As String
Dim Ors As ADODB.Recordset
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook

Private Sub CmdProceso_Click()
    Dim nFil As Integer
    Cadena = Empty
    Dim fec1 As String
    Dim fec2 As String
    Set Ors = New ADODB.Recordset
    fec1 = Format(FecProceso1, "dd/mm/yyyy")
    fec2 = Format(FecProceso2, "dd/mm/yyyy")
    Cadena = "PlameVaca '" & fec1 & "'," & "'" & fec2 & "'"

   If fAbrRst(Ors, Cadena) Then Ors.MoveFirst Else Ors.Close: Set Ors = Nothing: Exit Sub
   
    Set xlApp1 = CreateObject("Excel.Application")
    xlApp1.Workbooks.Add
    Set xlSheet = xlApp1.Worksheets("HOJA1")
    xlSheet.Name = "IMPORTA"
    xlApp1.Sheets("IMPORTA").Select
    
    xlSheet.Range("A:A").ColumnWidth = 6
    xlSheet.Range("B:B").ColumnWidth = 8
    xlSheet.Range("C:C").ColumnWidth = 60
    xlSheet.Range("D:D").ColumnWidth = 15
    xlSheet.Range("A:D").NumberFormat = "@"
    xlSheet.Range("E:E").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    xlSheet.Range("F:F").ColumnWidth = 30
    xlSheet.Range("G:G").ColumnWidth = 15
    xlSheet.Range("F:G").NumberFormat = "@"
     
    nFil = 1

    xlSheet.Cells(nFil, 1).Value = Trae_CIA(wcia)
    xlSheet.Cells(nFil, 1).Font.Bold = True
    nFil = nFil + 1
    xlSheet.Cells(nFil, 1).Value = "REPORTE DE MARCACIONES FALTAS VACACIONES SUBSIDIOS PARA PLAME " & fec1 & " Al " & fec2
    xlSheet.Cells(nFil, 1).HorizontalAlignment = xlCenter
    'xlSheet.Range(xlSheet.Cells(nFil, 2), xlSheet.Cells(nFil, Cmbmes.ListIndex + 6)).Merge
    
    xlSheet.Cells(nFil, 1).Font.Bold = True
    nFil = nFil + 1
    
    xlSheet.Cells(nFil, 1).Value = "EXPRESADO EN DIAS"
    xlSheet.Cells(nFil, 1).Font.Bold = True
    nFil = nFil + 2
    
    xlSheet.Cells(nFil, 1).Value = "ACTIVIDAD"
    xlSheet.Cells(nFil, 2).Value = "CODIGO"
    xlSheet.Cells(nFil, 3).Value = "NOMBRE"
    xlSheet.Cells(nFil, 4).Value = "NRO.DOC"
    xlSheet.Cells(nFil, 5).Value = "DIAS"
    xlSheet.Cells(nFil, 6).Value = "ACTIVIDAD2"
    xlSheet.Cells(nFil, 7).Value = "CODIGO SUNAT"
    xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).HorizontalAlignment = xlCenter
    xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).VerticalAlignment = xlCenter
    xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Borders.LineStyle = xlContinuous
    xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, 7)).Font.Bold = True
    
    nFil = 6
    Dim lCount As Integer
    lCount = 1
     
    If Ors.RecordCount > 0 Then Ors.MoveFirst
    Do While Not Ors.EOF
      xlSheet.Cells(nFil, 1).Value = Trim(Ors!Actividad & "")
      xlSheet.Cells(nFil, 2).Value = Trim(Ors!Codigo & "")
      xlSheet.Cells(nFil, 3).Value = Trim(Ors!nombre & "")
      xlSheet.Cells(nFil, 4).Value = Trim(Ors!Nrodoc & "")
      xlSheet.Cells(nFil, 5).Value = Trim(Ors!Dias & "")
      xlSheet.Cells(nFil, 6).Value = Trim(Ors!Actividad2 & "")
      xlSheet.Cells(nFil, 7).Value = Trim(Ors!CodigoSunat & "")
      nFil = nFil + 1
      Ors.MoveNext
    Loop
    nFil = nFil + 1
    Ors.Close: Set Ors = Nothing
    
    nFil = nFil + 5
    

    xlApp1.Application.ActiveWindow.DisplayGridLines = False
    
    xlApp1.Application.Visible = True
    
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
    If Not xlBook Is Nothing Then Set xlBook = Nothing
    If Not xlSheet Is Nothing Then Set xlSheet = Nothing
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
FecProceso1.Value = Now() - 30
FecProceso2.Value = Now()
End Sub
