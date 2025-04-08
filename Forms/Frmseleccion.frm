VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmseleccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Consultas de Planilla «"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   Icon            =   "Frmseleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkHasta 
      Caption         =   "Hasta:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox Cmbmeshasta 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Frmseleccion.frx":030A
      Left            =   1920
      List            =   "Frmseleccion.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   480
      Width           =   2400
   End
   Begin VB.TextBox txtAnohasta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1005
      TabIndex        =   29
      Top             =   495
      Width           =   615
   End
   Begin VB.TextBox Txtsemana 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      TabIndex        =   27
      Top             =   120
      Width           =   375
   End
   Begin Threed.SSPanel Panelprogress 
      Height          =   735
      Left            =   2985
      TabIndex        =   21
      Top             =   5010
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Procesando Informe"
      ForeColor       =   4210752
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodColor      =   4210752
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   180
         Left            =   105
         TabIndex        =   22
         Top             =   405
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.ComboBox Cmbtipobol 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   525
      Width           =   3015
   End
   Begin VB.ComboBox Cmbtipo 
      Height          =   315
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox Cmblistado 
      Height          =   315
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Txtano 
      Height          =   285
      Left            =   1005
      TabIndex        =   6
      Top             =   105
      Width           =   615
   End
   Begin VB.ComboBox Cmbmes 
      Height          =   315
      ItemData        =   "Frmseleccion.frx":039A
      Left            =   1920
      List            =   "Frmseleccion.frx":03C2
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   90
      Width           =   2400
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   5000
      Begin MSDataGridLib.DataGrid DgdApo 
         Height          =   1455
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2566
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Aportaciones"
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
            DataField       =   "tipo"
            Caption         =   "tipo"
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
            DataField       =   "numfield"
            Caption         =   "numfield"
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
               ColumnWidth     =   4169.764
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   3720
      Width           =   5000
      Begin MSDataGridLib.DataGrid DgdDed 
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   60
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Deducciones"
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
            DataField       =   "tipo"
            Caption         =   "tipo"
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
            DataField       =   "numfield"
            Caption         =   "numfield"
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
               ColumnWidth     =   4185.071
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5775
      Left            =   6240
      TabIndex        =   1
      Top             =   1440
      Width           =   5000
      Begin MSDataGridLib.DataGrid DgdReport 
         Height          =   5595
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   9869
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "descripcion"
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
            DataField       =   "tipo"
            Caption         =   "tipo"
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
            DataField       =   "numfield"
            Caption         =   "numfield"
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
            DataField       =   "item"
            Caption         =   "item"
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
               ColumnWidth     =   4050.142
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5000
      Begin MSDataGridLib.DataGrid DgdRem 
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Descripcion"
            Caption         =   "Remuneraciones"
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
            DataField       =   "tipo"
            Caption         =   "tipo"
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
            DataField       =   "numfield"
            Caption         =   "numfield"
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
               ColumnWidth     =   4215.118
            EndProperty
            BeginProperty Column01 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   615
      Left            =   5400
      TabIndex        =   14
      ToolTipText     =   "Generar Reporte"
      Top             =   6480
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "Frmseleccion.frx":042A
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   615
      Left            =   5370
      TabIndex        =   15
      ToolTipText     =   "Trasladar"
      Top             =   2745
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   1085
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "Frmseleccion.frx":09C4
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   615
      Left            =   5370
      TabIndex        =   16
      ToolTipText     =   "Trasladar"
      Top             =   3450
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "Frmseleccion.frx":0F5E
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   615
      Left            =   5370
      TabIndex        =   17
      ToolTipText     =   "Bajar una posicion"
      Top             =   4155
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "Frmseleccion.frx":14F8
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   615
      Left            =   5370
      TabIndex        =   18
      ToolTipText     =   "Subir una posicion"
      Top             =   2040
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1085
      _StockProps     =   78
      BevelWidth      =   1
      Picture         =   "Frmseleccion.frx":1A92
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   5670
      TabIndex        =   26
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSForms.SpinButton SpinButton2 
      Height          =   285
      Left            =   1635
      TabIndex        =   31
      Top             =   480
      Width           =   255
      VariousPropertyBits=   25
      Size            =   "450;503"
      PrevEnabled     =   0
      NextEnabled     =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Semana"
      Height          =   195
      Left            =   4560
      TabIndex        =   28
      Top             =   120
      Width           =   585
   End
   Begin VB.Label LblCia 
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Boleta"
      Height          =   195
      Left            =   6240
      TabIndex        =   23
      Top             =   525
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Trabajador"
      Height          =   195
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   1350
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   285
      Left            =   1635
      TabIndex        =   13
      Top             =   90
      Width           =   255
      Size            =   "450;503"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listado"
      Height          =   195
      Left            =   6240
      TabIndex        =   11
      Top             =   960
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "Frmseleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRem As New Recordset
Dim rsDed As New Recordset
Dim rsApo As New Recordset
Dim rsReport As New Recordset
Dim Sql As String
Dim VTipo As String
Dim VTipobol As String
Dim mFocus As String
Dim vID_Report As String
Dim Carga_Reports As String
Dim rs2 As ADODB.Recordset
Dim rsClon As New Recordset



Private Sub chkHasta_Click()
txtAnohasta.Text = Year(Now)
Cmbmeshasta.ListIndex = Month(Now) - 1
txtAnohasta.Enabled = CBool(chkHasta.Value)
SpinButton2.Enabled = CBool(chkHasta.Value)
Cmbmeshasta.Enabled = CBool(chkHasta.Value)

End Sub

Private Sub Cmblistado_Click()
vID_Report = fc_CodigoComboBox(Cmblistado, 4)
Carga_Report
If Cmblistado.Text = "NUEVO" Then
    Reporte_Seleccion
End If

End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
'Carga_Reports = "select distinct(id_report),name_report from plareports where cia='" & wcia & "' and tipo='02' and status<>'*' and user_crea='" & wuser & "' and referencia='" & VTipo & "' order by id_report"
Carga_Reports = "select distinct(id_report),name_report from plareports where cia='" & wcia & "' and tipo='02' and status<>'*' order by id_report"

Call rCarCbo(Cmblistado, Carga_Reports, "C", "0000")
Cmblistado.AddItem "NUEVO"
Txtsemana.Enabled = IIf(VTipo = "02" And VTipobol = "01", True, False)
If VTipo = "01" Then Txtsemana.Text = "  "
Reporte_Seleccion
End Sub

Private Sub Cmbtipobol_Click()
VTipobol = fc_CodigoComboBox(Cmbtipobol, 2)
Txtsemana.Enabled = IIf(VTipobol = "01" And VTipo = "02", True, False)
If VTipo <> "01" Then Txtsemana.Text = "  "
End Sub

Private Sub DgdApo_GotFocus()
mFocus = "A"
End Sub

Private Sub DgdDed_GotFocus()
mFocus = "D"
End Sub

Private Sub DgdRem_GotFocus()
mFocus = "I"
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 11265
Me.Height = 7680
Sql = "select razsoc from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(Rs, Sql)) Then Lblcia.Caption = Trim(Rs(0))
Rs.Close
Call Crea_Rs
Cmbmes.ListIndex = Month(Date) - 1
Txtano.Text = Format(Year(Date))
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Call fc_Descrip_Maestros2("01078", "", Cmbtipobol)
Cmbtipobol.AddItem "TOTAL"
Call Reporte_Seleccion
End Sub

Private Sub SpinButton1_SpinDown()
    Call Down(Txtano)
End Sub

Private Sub SpinButton1_SpinUp()
'If Not IsNumeric(Txtano) Then
'   Txtano = Format(Year(Date), "0000")
'Else
'   Txtano = Txtano + 1
'End If
    Call Up(Txtano)
End Sub

Private Sub SpinButton2_SpinDown()
Call Down(txtAnohasta)
End Sub

Private Sub SpinButton2_SpinUp()
Call Up(txtAnohasta)
End Sub

Private Sub SSCommand1_Click()
Dim MREC As Integer
Dim vIt As Integer
If rsReport.RecordCount <= 1 Then Exit Sub
If DgdReport.Columns(3) = "1" Then Exit Sub
MREC = rsReport.AbsolutePosition
DgdReport.Columns(3) = DgdReport.Columns(3) - 1
vIt = DgdReport.Columns(3)
If DgdReport.Row > 0 Then DgdReport.Row = DgdReport.Row - 1
DgdReport.Columns(3) = vIt + 1
DgdReport.Columns(3) = DgdReport.Columns(3) + 1
DgdReport.Row = DgdReport.Row + 1
rsReport.Sort = "item"
DgdReport.Refresh
If MREC > 1 Then rsReport.AbsolutePosition = MREC - 1
End Sub

Private Sub SSCommand2_Click()
Dim vIt As Integer
Dim MREC As Integer
If rsReport.RecordCount <= 1 Then Exit Sub
If rsReport.AbsolutePosition = rsReport.RecordCount Then Exit Sub
MREC = rsReport.AbsolutePosition
DgdReport.Columns(3) = DgdReport.Columns(3) + 1
vIt = DgdReport.Columns(3)
DgdReport.Row = DgdReport.Row + 1
DgdReport.Columns(3) = vIt - 1
If DgdReport.Row > 0 Then DgdReport.Row = DgdReport.Row - 1
rsReport.Sort = "item"
DgdReport.Refresh
rsReport.AbsolutePosition = MREC + 1
End Sub

Private Sub SSCommand3_Click()
If rsReport.RecordCount <= 0 Then Exit Sub

Select Case rsReport!tipo
       Case Is = "I"
            rsRem.AddNew
            rsRem!Descripcion = rsReport!Descripcion
            rsRem!numfield = rsReport!numfield
            rsRem!tipo = rsReport!tipo
            rsRem!Codigo = rsReport!Codigo
            rsReport.Delete
       Case Is = "D"
            rsDed.AddNew
            rsDed!Descripcion = rsReport!Descripcion
            rsDed!numfield = rsReport!numfield
            rsDed!tipo = rsReport!tipo
            rsDed!Codigo = rsReport!Codigo
            rsReport.Delete
       Case Is = "A"
            rsApo.AddNew
            rsApo!Descripcion = rsReport!Descripcion
            rsApo!numfield = rsReport!numfield
            rsApo!tipo = rsReport!tipo
            rsApo!Codigo = rsReport!Codigo
            rsReport.Delete
End Select
End Sub

Private Sub SSCommand4_Click()
If rsReport.RecordCount > 0 Then rsReport.MoveLast
Select Case mFocus
       Case Is = "I"
            If rsRem.RecordCount <= 0 Then Exit Sub
            rsReport.AddNew
            rsReport!Descripcion = rsRem!Descripcion
            rsReport!numfield = rsRem!numfield
            rsReport!tipo = rsRem!tipo
            rsReport!Codigo = rsRem!Codigo
            rsReport!Item = rsReport.RecordCount
            rsRem.Delete
       Case Is = "D"
            If rsDed.RecordCount <= 0 Then Exit Sub
            rsReport.AddNew
            rsReport!Descripcion = rsDed!Descripcion
            rsReport!numfield = rsDed!numfield
            rsReport!tipo = rsDed!tipo
            rsReport!Codigo = rsDed!Codigo
            rsReport!Item = rsReport.RecordCount
            rsDed.Delete
       Case Is = "A"
            If rsApo.RecordCount <= 0 Then Exit Sub
            rsReport.AddNew
            rsReport!Descripcion = rsApo!Descripcion
            rsReport!numfield = rsApo!numfield
            rsReport!tipo = rsApo!tipo
            rsReport!Codigo = rsApo!Codigo
            rsReport!Item = rsReport.RecordCount
            rsApo.Delete
End Select
End Sub

Private Sub SSCommand6_Click()
If chkHasta.Value = 0 Then
   Genera_Seleccion
Else
    Genera_Seleccion_Periodo
End If
End Sub

Private Sub Txtano_KeyPress(KeyAscii As Integer)
Txtano.Text = Txtano.Text + fc_ValNumeros(KeyAscii)
End Sub
Private Sub Crea_Rs()
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "descripcion", adVarChar, 60, adFldIsNullable
    rsRem.Fields.Append "tipo", adChar, 1, adFldIsNullable
    rsRem.Fields.Append "numfield", adInteger, 4, adFldIsNullable
    rsRem.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsRem.Open
    Set DgdRem.DataSource = rsRem

    If rsDed.State = 1 Then rsDed.Close
    rsDed.Fields.Append "descripcion", adVarChar, 60, adFldIsNullable
    rsDed.Fields.Append "tipo", adChar, 1, adFldIsNullable
    rsDed.Fields.Append "numfield", adInteger, 4, adFldIsNullable
    rsDed.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsDed.Open
    Set DgdDed.DataSource = rsDed

    If rsApo.State = 1 Then rsApo.Close
    rsApo.Fields.Append "descripcion", adVarChar, 60, adFldIsNullable
    rsApo.Fields.Append "tipo", adChar, 1, adFldIsNullable
    rsApo.Fields.Append "numfield", adInteger, 4, adFldIsNullable
    rsApo.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsApo.Open
    Set DgdApo.DataSource = rsApo
    
    If rsReport.State = 1 Then rsReport.Close
    rsReport.Fields.Append "descripcion", adVarChar, 60, adFldIsNullable
    rsReport.Fields.Append "tipo", adChar, 1, adFldIsNullable
    rsReport.Fields.Append "item", adInteger, 4, adFldIsNullable
    rsReport.Fields.Append "numfield", adInteger, 4, adFldIsNullable
    rsReport.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsReport.Open
    Set DgdReport.DataSource = rsReport
    

End Sub
Private Sub Reporte_Seleccion()
Cmblistado.ListIndex = -1
If Cmbtipo.ListIndex < 0 Then Exit Sub
If rsReport.State = 1 Then
    If rsReport.RecordCount > 0 Then
       rsReport.MoveFirst
       Do While Not rsReport.EOF
          rsReport.Delete
          rsReport.MoveNext
       Loop
    End If
End If
If rsRem.State = 1 Then
    If rsRem.RecordCount > 0 Then
       rsRem.MoveFirst
       Do While Not rsRem.EOF
          rsRem.Delete
          rsRem.MoveNext
       Loop
    End If
End If

Sql = "select distinct(c.codinterno) AS codigo,descripcion from plaseteoprint s,placonstante c " _
    & "where s.cia='" & wcia & "' and s.tipo='I' and s.tipo_trab='" & VTipo & "' and s.status<>'*' " _
    & "and c.cia=s.cia and c.tipomovimiento='02' and s.codigo=c.codinterno and c.status<>'*' "

Sql = Sql & " UNION SELECT DISTINCT(c.codinterno) AS codigo,descripcion FROM placonstante c " _
    & "WHERE c.cia='" & wcia & "' AND c.status<>'*' " _
    & "AND c.tipomovimiento='02' AND c.codinterno IN('39','40','45') " _
    & "ORDER BY c.codinterno"


'Sql = "select distinct(codigo),descripcion from plaseteoprint s,placonstante c " _
'    & "where s.cia='" & wcia & "' and s.tipo='I' and s.tipo_trab='" & VTipo & "' and s.status<>'*' " _
'    & "and c.cia=s.cia and c.tipomovimiento='02' and s.codigo=c.codinterno and c.status<>'*' " _
'    & "order by s.codigo"
    
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Do While Not Rs.EOF
   rsRem.AddNew
   rsRem!Descripcion = Rs!Descripcion
   rsRem!numfield = 43 + Val(Rs!Codigo)
   rsRem!tipo = "I"
   rsRem!Codigo = Rs!Codigo
   Rs.MoveNext
Loop
rsRem.AddNew
rsRem!Descripcion = "TOTAL INGRESO"
rsRem!numfield = 137
rsRem!tipo = "I"
rsRem!Codigo = "TI"
rsRem.AddNew
rsRem!Descripcion = "TOTAL NETO"
rsRem!numfield = 138
rsRem!tipo = "I"
rsRem!Codigo = "TN"

If rsRem.RecordCount > 0 Then rsRem.MoveFirst
Rs.Close

If rsDed.RecordCount > 0 Then
   rsDed.MoveFirst
   Do While Not rsDed.EOF
      rsDed.Delete
      rsDed.MoveNext
   Loop
End If

Sql = "select distinct(codigo),descripcion from plaseteoprint s,placonstante c " _
    & "where s.cia='" & wcia & "' and s.tipo='D' and s.tipo_trab='" & VTipo & "' and s.status<>'*' " _
    & "and c.cia=s.cia and c.tipomovimiento='03' and s.codigo=c.codinterno and c.status<>'*' " _
    & "order by s.codigo"
    
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Do While Not Rs.EOF
   rsDed.AddNew
   rsDed!Descripcion = Rs!Descripcion
   rsDed!numfield = 43 + 20 + Val(Rs!Codigo)
   rsDed!tipo = "D"
   rsDed!Codigo = Rs!Codigo
   Rs.MoveNext
Loop
rsDed.AddNew
rsDed!Descripcion = "TOTAL DEDUCCION"
rsDed!numfield = 136
rsDed!tipo = "D"
rsDed!Codigo = "TD"

If rsDed.RecordCount > 0 Then rsDed.MoveFirst
Rs.Close

If rsApo.RecordCount > 0 Then
   rsApo.MoveFirst
   Do While Not rsApo.EOF
      rsApo.Delete
      rsApo.MoveNext
   Loop
End If

Sql = "select distinct(codigo),descripcion from plaseteoprint s,placonstante c " _
    & "where s.cia='" & wcia & "' and s.tipo='A' and s.tipo_trab='" & VTipo & "' and s.status<>'*' " _
    & "and c.cia=s.cia and c.tipomovimiento='03' and s.codigo=c.codinterno and c.status<>'*' " _
    & "order by s.codigo"
    
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Do While Not Rs.EOF
   rsApo.AddNew
   rsApo!Descripcion = Rs!Descripcion
   rsApo!numfield = 43 + 20 + 20 + Val(Rs!Codigo)
   rsApo!tipo = "A"
   rsApo!Codigo = Rs!Codigo
   Rs.MoveNext
Loop
rsApo.AddNew
rsApo!Descripcion = "TOTAL APORTACION"
rsApo!numfield = 135
rsApo!tipo = "A"
rsApo!Codigo = "TA"

If rsApo.RecordCount > 0 Then rsApo.MoveFirst
Rs.Close

End Sub
Public Sub Grabar_Report_Seleccion()
Dim vName As String
Dim VNew As Boolean
Dim vOrden As Integer
Dim Mgrab As Integer
Dim NroTrans As Integer
On Error GoTo Salir
NroTrans = 0

If rsReport.RecordCount <= 0 Then Exit Sub
rsReport.MoveFirst
VNew = False
Mgrab = MsgBox("Seguro de Grabar Reporte", vbYesNo + vbQuestion, "Reporte de Importes")
If Mgrab <> 6 Then Exit Sub
If Cmblistado.Text = "NUEVO" Or Cmblistado.ListIndex < 0 Then
   VNew = True
   Do While Trim(vName) = ""
      vName = InputBox("Ingrese el Nombre del Reporte", "Reporte Nuevo")
   Loop
Else
   Mgrab = MsgBox("Desea Remplazar Reporte", vbYesNo + vbQuestion, "Reporte de Trabajadores")
   If Mgrab = 6 Then
      VNew = False
   Else
      VNew = True
      Do While Trim(vName) = ""
         vName = InputBox("Ingrese el Nombre del Reporte", "Reporte Nuevo")
      Loop
   End If
End If
Screen.MousePointer = vbArrowHourglass
If VNew = True Then
   Sql = "select max(id_report) from plareports where cia='" & wcia & "' and tipo='02' and status<>'*'"
   If (fAbrRst(Rs, Sql)) Then
      If IsNull(Rs(0)) Then vID_Report = "0001" Else vID_Report = Format(Val(Rs(0) + 1), "0000")
   Else
      vID_Report = "0001"
   End If
   Rs.Close
Else
   vName = Cmblistado.Text
End If
cn.BeginTrans
NroTrans = 1
Sql = "update plareports set status='*' where cia='" & wcia & "' and tipo='02' and id_report='" & vID_Report & "' and status<>'*'"
cn.Execute Sql
Do While Not rsReport.EOF
   Sql = "set dateformat " & Coneccion.FormatFechaSql & " "
   Sql = Sql & "insert into plareports values('" & wcia & "','02','" & vID_Report & _
   "','" & UCase(Trim(Left(vName, 60))) & "','" & rsReport!Codigo & "'," & rsReport!Item & _
   ",0,'" & rsReport!tipo & "','','" & wuser & "'," _
 & "0," & rsReport!numfield & ",'','','" & VTipo & "',''," & FechaSys & ")"
'       Debug.Print SQL
   cn.Execute Sql
   rsReport.MoveNext
Loop

cn.CommitTrans

Call rCarCbo(Cmblistado, Carga_Reports, "C", "0000")
Cmblistado.AddItem "NUEVO"
Call rUbiIndCmbBox(Cmblistado, vID_Report, "0000")

Screen.MousePointer = vbDefault
Exit Sub
Salir:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    MsgBox Err.Description, vbCritical, Me.Caption
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub Carga_Report()
If Cmblistado.ListIndex < 0 Then Exit Sub
If rsReport.RecordCount > 0 Then
   rsReport.MoveFirst
   Do While Not rsReport.EOF
      rsReport.Delete
      rsReport.MoveNext
   Loop
End If
Sql = "select * from plareports where cia='" & wcia & "' and tipo='02' and id_report='" & vID_Report & "' and status<>'*' order by item"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Do While Not Rs.EOF
   rsReport.AddNew
   Select Case Left(Rs!Quiebre, 1)
          Case Is = "I"
               Sql = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno='" & Left(Rs!Campo, 2) & "' and status<>'*'"
          Case Else
               Sql = "select descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & Left(Rs!Campo, 2) & "' and status<>'*'"
   End Select
   If (fAbrRst(rs2, Sql)) Then
      rsReport!Descripcion = rs2(0)
   Else
     If Left(Rs!Campo, 2) = "TI" Then rsReport!Descripcion = "TOTAL INGRESO"
     If Left(Rs!Campo, 2) = "TA" Then rsReport!Descripcion = "TOTAL APORTACION"
     If Left(Rs!Campo, 2) = "TD" Then rsReport!Descripcion = "TOTAL DEDUCCION"
     If Left(Rs!Campo, 2) = "TN" Then rsReport!Descripcion = "TOTAL NETO"
   End If
   rs2.Close
   rsReport!numfield = Rs!numfield
   rsReport!tipo = Left(Rs!Quiebre, 1)
   rsReport!Codigo = Left(Rs!Campo, 2)
   rsReport!Item = Rs!Item
   Rs.MoveNext
Loop
Rs.Close

Set DgdReport.DataSource = rsReport

End Sub
Private Sub Genera_Seleccion()
Dim mcamp As String
Dim mcad As String
Dim mcadbol As String
Dim mFields As Integer
Dim I As Integer
Dim nFil As Integer
Dim nCol As Integer
Dim msum As Integer
Dim NroMeses As Integer

On Error GoTo CORRIGE
Screen.MousePointer = vbArrowHourglass

Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 8
xlSheet.Range("B:B").ColumnWidth = 30
xlSheet.Range("C:AZ").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

mcadbol = " and h.proceso='" & VTipobol & "' "
If Cmbtipobol.ListIndex < 0 Then mcadbol = ""
If Cmbtipobol.Text = "TOTAL" Then mcadbol = ""
Set rsClon = rsReport.Clone
If rsClon.RecordCount > 0 Then rsClon.MoveFirst
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Max = rsClon.RecordCount
Sql = nombre()
Sql = Sql$ & "p.placod,"
mcad = ""

mFields = 0
Do While Not rsClon.EOF
   Barra.Value = rsClon.AbsolutePosition
   mcamp = rsClon!tipo & rsClon!Codigo
   If Left(rsClon!Codigo, 1) <> "T" Then
      mcad = mcad & "sum(" & mcamp & "),"
      xlSheet.Cells(6, mFields + 3).Value = rsClon!Descripcion
   Else
     If Left(rsClon!Codigo, 2) = "TI" Then mcad = mcad & "sum(totaling),": xlSheet.Cells(6, mFields + 3).Value = "TOTAL INGRESO"
     If Left(rsClon!Codigo, 2) = "TA" Then mcad = mcad & "sum(totalapo),": xlSheet.Cells(6, mFields + 3).Value = "TOTAL APORTACION"
     If Left(rsClon!Codigo, 2) = "TD" Then mcad = mcad & "sum(totalded),": xlSheet.Cells(6, mFields + 3).Value = "TOTAL DEDUCCION"
     If Left(rsClon!Codigo, 2) = "TN" Then mcad = mcad & "sum(totneto),": xlSheet.Cells(6, mFields + 3).Value = "TOTAL NETO"
   End If
   xlSheet.Range(xlSheet.Cells(6, mFields + 3), xlSheet.Cells(7, mFields + 3)).Merge
   xlSheet.Cells(6, mFields + 3).HorizontalAlignment = xlCenter
   xlSheet.Cells(6, mFields + 3).VerticalAlignment = xlDistributed
   mFields = mFields + 1
   rsClon.MoveNext
Loop
If Trim(mcad) <> "" Then mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
Sql = Sql & mcad
Sql = Sql & " from planillas p,plahistorico h " _
     & "where h.cia='" & wcia & "' and year(fechaproceso)=" & Val(Txtano.Text) & " and month(fechaproceso)=" & Cmbmes.ListIndex + 1 & " " _
     & "and h.status='T' and p.status<>'*' and tipotrab='" & VTipo & "' " & mcadbol
     
If Trim(Txtsemana.Text) <> "" Then
    Sql = Sql & " And semana='" & Format(Txtsemana.Text, "00") & "' "
End If

Sql = Sql & "and h.cia=p.cia and h.placod=p.placod group by " & _
     "p.placod,p.ap_pat,p.ap_mat,nom_1,nom_2 " & _
     "order by nombre"
     
'Debug.Print SQL
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst

xlSheet.Cells(1, 1).Value = Lblcia
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(1, 1).Font.Size = 12

xlSheet.Cells(3, 1).Value = Cmbtipo.Text & "  " & Cmbtipobol.Text
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).Font.Size = 12
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, mFields + 2)).Merge

xlSheet.Cells(4, 1).Value = Cmbmes.Text & " - " & Txtano.Text & IIf(RTrim(Txtsemana.Text) <> "", "SEM " + Txtsemana.Text, "")
xlSheet.Cells(4, 1).Font.Bold = True
xlSheet.Cells(4, 1).Font.Size = 12
xlSheet.Cells(4, 1).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, mFields + 2)).Merge

xlSheet.Cells(6, 1).Value = "CODIGO"
xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(7, 1)).Merge
xlSheet.Cells(6, 1).HorizontalAlignment = xlCenter
xlSheet.Cells(6, 1).VerticalAlignment = xlDistributed

xlSheet.Cells(6, 2).Value = "NOMBRE"
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 2)).Merge
xlSheet.Cells(6, 2).HorizontalAlignment = xlCenter
xlSheet.Cells(6, 2).VerticalAlignment = xlDistributed

xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(7, mFields + 2)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(7, mFields + 2)).Font.Bold = True

nFil = 8
Do While Not Rs.EOF
   xlSheet.Cells(nFil, 1) = Rs!PlaCod
   xlSheet.Cells(nFil, 2) = Rs!nombre
   nCol = 3
   For I = 1 To mFields
       xlSheet.Cells(nFil, nCol) = Rs(I + 1)
       nCol = nCol + 1
   Next
   nFil = nFil + 1
   Rs.MoveNext
Loop
Rs.Close
Set rsClon = Nothing

msum = ((nFil - 7) * -1)
nFil = nFil + 1
xlSheet.Cells(nFil, 2).Value = "TOTALES"
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter

If nCol >= 3 Then

   For I = 3 To nCol - 1
       xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
   Next I
   
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, nCol)).Font.Bold = True
End If

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "REPORTE DE PERSONAL"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = vbDefault
Panelprogress.Visible = False

Exit Sub

CORRIGE:

   MsgBox "Error : " & Err.Description, vbCritical, "Genera_Seleccion"

End Sub

Private Sub Genera_Seleccion_Periodo()
Dim mcamp As String
Dim mcad As String
Dim mcadbol As String
Dim mFields As Integer
Dim I As Integer
Dim nFil As Integer
Dim nCol As Integer
Dim msum As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim FechaTemp As Date
Dim Mescolumna As Integer
Dim rsPersonal As New ADODB.Recordset
Dim NroMeses As Integer
On Error GoTo CORRIGE

fecha1 = CDate(("01/" & Format(Cmbmes.ListIndex + 1, "00") & "/" & Format(Txtano.Text, "0000")))
fecha2 = CDate(("01/" & Format(Cmbmeshasta.ListIndex + 1, "00") & "/" & Format(txtAnohasta.Text, "0000")))
fecha2 = DateAdd("m", 1, fecha2)
fecha2 = DateAdd("s", -1, fecha2)

If fecha1 > fecha2 Then
    MsgBox "Ingrese un Periodo Correctamente", vbCritical, Me.Caption
    Exit Sub
End If

Screen.MousePointer = vbArrowHourglass

Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("A:A").ColumnWidth = 8
xlSheet.Range("B:B").ColumnWidth = 30
xlSheet.Range("C:AZ").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

mcadbol = " and h.proceso='" & VTipobol & "' "
If Cmbtipobol.ListIndex < 0 Then mcadbol = ""
If Cmbtipobol.Text = "TOTAL" Then mcadbol = ""
Set rsClon = rsReport.Clone
If rsClon.RecordCount > 0 Then rsClon.MoveFirst
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Max = rsClon.RecordCount
Sql = nombre()
Sql = Sql$ & " p.placod,year(h.fechaproceso) as anio,month(h.fechaproceso) as mes,"
mcad = ""

mFields = 0
Mescolumna = 0
NroMeses = 0

FechaTemp = fecha1
While FechaTemp <= fecha2
    NroMeses = NroMeses + 1
    rsClon.MoveFirst
    Do While Not rsClon.EOF
        
        If FechaTemp = fecha1 Then
            mcamp = rsClon!tipo & rsClon!Codigo
        End If
        
        If Left(rsClon!Codigo, 1) <> "T" Then
            If FechaTemp = fecha1 Then
                mcad = mcad & "sum(" & mcamp & ") as " & mcamp & ","
            End If
           xlSheet.Cells(6, Mescolumna + 3).Value = rsClon!Descripcion
        Else
            If Left(rsClon!Codigo, 2) = "TI" Then
                If FechaTemp = fecha1 Then
                    mcad = mcad & "sum(totaling) as 'totaling',"
                End If
                xlSheet.Cells(6, Mescolumna + 3).Value = "TOTAL INGRESO"
            End If
            If Left(rsClon!Codigo, 2) = "TA" Then
                If FechaTemp = fecha1 Then
                    mcad = mcad & "sum(totalapo) as 'totalapo',"
                End If
                xlSheet.Cells(6, Mescolumna + 3).Value = "TOTAL APORTACION"
            End If
            If Left(rsClon!Codigo, 2) = "TD" Then
                If FechaTemp = fecha1 Then
                    mcad = mcad & "sum(totalded) as 'totalded',"
                End If
                xlSheet.Cells(6, Mescolumna + 3).Value = "TOTAL DEDUCCION"
            End If
            If Left(rsClon!Codigo, 2) = "TN" Then
                If FechaTemp = fecha1 Then
                    mcad = mcad & "sum(totneto) as 'totneto',"
                End If
                xlSheet.Cells(6, Mescolumna + 3).Value = "TOTAL NETO"
            End If
        End If
        xlSheet.Range(xlSheet.Cells(6, Mescolumna + 3), xlSheet.Cells(7, Mescolumna + 3)).Merge
        xlSheet.Cells(6, Mescolumna + 3).HorizontalAlignment = xlCenter
        xlSheet.Cells(6, Mescolumna + 3).VerticalAlignment = xlDistributed
        If FechaTemp = fecha1 Then
            mFields = mFields + 1
        End If
        Mescolumna = Mescolumna + 1
        rsClon.MoveNext
    Loop
    FechaTemp = DateAdd("m", 1, FechaTemp)
Wend
Barra.Max = rsClon.RecordCount * NroMeses

If Trim(mcad) <> "" Then mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)
Sql = Sql & mcad
Sql = Sql & " into #Reporte" & wuser & " from planillas p,plahistorico h " _
     & " where h.cia='" & wcia & "' and fechaproceso between '" & Format(fecha1, "mm/dd/yyyy") & "' and '" & Format(fecha2, "mm/dd/yyyy") _
     & "' and h.status='T' and p.status<>'*' and tipotrab='" & VTipo & "' " & mcadbol
     
'If Trim(Txtsemana.Text) <> "" Then
'    Sql = Sql & " And semana='" & Format(Txtsemana.Text, "00") & "' "
'End If

Sql = Sql & " and h.cia=p.cia and h.placod=p.placod group by " & _
     " p.placod,p.ap_pat,p.ap_mat,nom_1,nom_2,year(h.fechaproceso),month(h.fechaproceso) " & _
     " order by nombre"
     
'Debug.Print SQL
cn.Execute Sql

Sql = "select distinct placod,nombre from #Reporte" & wuser & " order by 2"
If (fAbrRst(rsPersonal, Sql)) Then rsPersonal.MoveFirst


xlSheet.Cells(1, 1).Value = Lblcia
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(1, 1).Font.Size = 12

xlSheet.Cells(3, 1).Value = Cmbtipo.Text & "  " & Cmbtipobol.Text
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).Font.Size = 12
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, Mescolumna + 2)).Merge
xlSheet.Cells(4, 1).Font.Size = 12
xlSheet.Cells(5, 1).Font.Size = 12
Dim Desplazar As Integer
Desplazar = 0
FechaTemp = fecha1
While FechaTemp <= fecha2

    xlSheet.Cells(5, 3 + Desplazar).Value = Name_Month(Format(Month(FechaTemp), "00")) & " - " & Year(FechaTemp)
    xlSheet.Cells(5, 3 + Desplazar).Font.Bold = True
    xlSheet.Cells(5, 3 + Desplazar).Font.Size = 12
    xlSheet.Cells(5, 3 + Desplazar).HorizontalAlignment = xlCenter
    xlSheet.Range(xlSheet.Cells(5, 3 + Desplazar), xlSheet.Cells(5, mFields + Desplazar + 2)).Merge
    xlSheet.Range(xlSheet.Cells(5, 3 + Desplazar), xlSheet.Cells(5, mFields + Desplazar + 2)).Borders.LineStyle = xlContinuous
    
    Desplazar = Desplazar + mFields
    FechaTemp = DateAdd("m", 1, FechaTemp)
    
Wend

xlSheet.Cells(6, 1).Value = "CODIGO"
xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(7, 1)).Merge
xlSheet.Cells(6, 1).HorizontalAlignment = xlCenter
xlSheet.Cells(6, 1).VerticalAlignment = xlDistributed

xlSheet.Cells(6, 2).Value = "NOMBRE"
xlSheet.Range(xlSheet.Cells(6, 2), xlSheet.Cells(7, 2)).Merge
xlSheet.Cells(6, 2).HorizontalAlignment = xlCenter
xlSheet.Cells(6, 2).VerticalAlignment = xlDistributed

xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(7, Mescolumna + 2)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(7, Mescolumna + 2)).Font.Bold = True

nFil = 8
If rsPersonal.RecordCount > 0 Then
    Barra.Value = 0
    Barra.Max = rsPersonal.RecordCount
Else
    Barra.Value = 1
    Barra.Max = 1
End If

Do While Not rsPersonal.EOF
   Barra.Value = Barra.Value + 1
   DoEvents
   xlSheet.Cells(nFil, 1) = rsPersonal!PlaCod
   xlSheet.Cells(nFil, 2) = rsPersonal!nombre
   nCol = 3
   
   FechaTemp = fecha1
   While FechaTemp <= fecha2
       Sql = "select * from #Reporte" & wuser & " where placod='" & rsPersonal!PlaCod & "' and anio=" & Year(FechaTemp) & " and mes=" & Month(FechaTemp)
       If (fAbrRst(Rs, Sql)) Then
            Rs.MoveFirst
            Do While Not Rs.EOF
                 For I = 1 To mFields
                     xlSheet.Cells(nFil, nCol) = Rs(I + 1 + 2)
                     nCol = nCol + 1
                 Next
                 Rs.MoveNext
            Loop
        Else
            nCol = nCol + mFields
        End If
       Rs.Close
       FechaTemp = DateAdd("m", 1, FechaTemp)
   Wend
   nFil = nFil + 1
   rsPersonal.MoveNext
Loop
rsPersonal.Close
Set rsPersonal = Nothing
Set rsClon = Nothing

Sql = "drop table #Reporte" & wuser
cn.Execute Sql

msum = ((nFil - 7) * -1)
nFil = nFil + 1
xlSheet.Cells(nFil, 2).Value = "TOTALES"
xlSheet.Cells(nFil, 2).HorizontalAlignment = xlCenter

If nCol >= 3 Then

   For I = 3 To nCol - 1
       xlSheet.Cells(nFil, I).Value = "=SUM(R[" & msum & "]C:R[-1]C)"
   Next I
   
   xlSheet.Range(xlSheet.Cells(nFil, 1), xlSheet.Cells(nFil, nCol)).Font.Bold = True
End If

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "REPORTE DE PERSONAL"
xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = vbDefault
Panelprogress.Visible = False

Exit Sub

CORRIGE:
Sql = "drop table #Reporte" & wuser
cn.Execute Sql
   MsgBox "Error : " & Err.Description, vbCritical, "Genera_Seleccion"

End Sub

