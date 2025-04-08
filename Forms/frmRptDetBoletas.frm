VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptDetBoletas 
   Caption         =   "Form1"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   5745
   Begin VB.CommandButton btngenerar 
      Caption         =   "Generar"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1095
      Width           =   2415
   End
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Top             =   615
      Width           =   615
   End
   Begin VB.TextBox Txtcodigo 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1575
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "frmRptDetBoletas.frx":0000
      Left            =   1080
      List            =   "frmRptDetBoletas.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   615
      Width           =   2415
   End
   Begin VB.TextBox Txtsemana 
      Height          =   285
      Left            =   4650
      TabIndex        =   4
      Top             =   1095
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   80
         Width           =   4575
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
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   825
      End
   End
   Begin MSComCtl2.UpDown UpDown2 
      Height          =   300
      Left            =   5280
      TabIndex        =   3
      Top             =   1095
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   5280
      TabIndex        =   6
      Top             =   600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T. Trabajador"
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   1095
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Año"
      Height          =   195
      Left            =   4320
      TabIndex        =   14
      Top             =   615
      Width           =   285
   End
   Begin VB.Label Lblcodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1575
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Lblnombre 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   1575
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
      Height          =   195
      Left            =   720
      TabIndex        =   11
      Top             =   615
      Width           =   300
   End
   Begin VB.Label Lblsemana 
      AutoSize        =   -1  'True
      Caption         =   "Semana"
      Height          =   195
      Left            =   3720
      TabIndex        =   10
      Top             =   1095
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frmRptDetBoletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btngenerar_Click()
Call generaReporte
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 5865
Me.Height = 2490
Label3.Caption = "Año"
Lblsemana.Caption = "Semana"

Txtano.Text = Format(Year(Date), "0000")
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Me.Caption = "Tabla Dinámica"
Label3.Caption = "Año Inicial"
Label3.Left = 3840
Lblsemana.Caption = "Año Final"
Txtsemana.Text = Format(Year(Date), "0000")

If wTipoPla <> "99" And UCase(wuser) <> "SA" Then
   If wTipoPla = "" Then
      Call rUbiIndCmbBox(Cmbtipo, "02", "00")
   Else
      Call rUbiIndCmbBox(Cmbtipo, Trim(wTipoPla), "00")
   End If
   Cmbtipo.Enabled = False
End If
 Call rUbiIndCmbBox(Cmbtipo, "02", "01")
 Cmbtipo.Text = "EMPLEADO"
End Sub

Sub generaReporte()

Dim nFil As Integer
Dim nCol As Integer
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook

'Dim xlApp1  As Object
'Dim xlBook As Object
'Dim xlApp2 As Object

Sql = "exec [ACSVVM04].SISGNRSWEB_Comacsa.DBO.usp_rpt_detalle_boletas '" & Cmbmes.Text & " " & Txtano.Text & "',''"
If (fAbrRst(rs, Sql)) Then
rs.MoveFirst
Else
rs.Close
Set rs = Nothing
MsgBox "No existen datos para mostrar", vbInformation
Exit Sub
End If

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Datos"

With xlSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
     "OLEDB;User ID=" & Trim(wuser) & ";PWD=" & Trim(wclave) & ";Data Source=" & wserver & ";Initial Catalog=" & WDatabase & ";Provider=SQLOLEDB.1"), Destination:=xlApp1.ActiveSheet.Range("$A$1")).QueryTable
     .CommandType = xlCmdSql
     .CommandText = Array("" & Sql & "")
     .AdjustColumnWidth = True
     .Refresh BackgroundQuery:=False
End With
xlApp1.ActiveSheet.Range("F1:F10000").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
Dim I As Integer
Dim mTipoMov As String
For I = 0 To rs.Fields.count - 1
    If I > 18 Then
       If UCase(Left(rs.Fields(I).Name, 1)) = "I" Then
          mTipoMov = "02"
       ElseIf UCase(Left(rs.Fields(I).Name, 1)) = "D" Or UCase(Left(rs.Fields(I).Name, 1)) = "A" Then
          mTipoMov = "03"
       ElseIf UCase(Left(rs.Fields(I).Name, 1)) = "H" Then
          mTipoMov = "**"
       End If
       xlSheet.Cells(1, I + 1).Value = UCase(Left(rs.Fields(I).Name, 1))
    End If
Next

On Error GoTo Termina

Termina:

xlApp2.Application.ActiveWindow.DisplayGridLines = False
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing


Screen.MousePointer = vbDefault
End Sub
