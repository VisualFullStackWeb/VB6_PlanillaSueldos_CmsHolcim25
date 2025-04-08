VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{FE1D1F8B-EC4B-11D3-B06C-00500427A693}#1.1#0"; "vbalLBar6.ocx"
Begin VB.Form Frmgrdctacte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CUENTA CORRIENTE DEL PERSONAL"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "Frmgrdctacte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Framemovi 
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   45
      TabIndex        =   6
      Top             =   6795
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CommandButton cmdimporte 
         Appearance      =   0  'Flat
         Caption         =   "Importe"
         Height          =   209
         Left            =   9405
         TabIndex        =   22
         Top             =   540
         Width           =   980
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   16
         ToolTipText     =   "Cerrar Ventana"
         Top             =   0
         Width           =   375
      End
      Begin MSComCtl2.DTPicker Txtal 
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   37914
      End
      Begin MSComCtl2.DTPicker Txtdel 
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   37914
      End
      Begin MSDataGridLib.DataGrid DgrdMovi 
         Height          =   4875
         Left            =   165
         TabIndex        =   7
         Top             =   525
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   8599
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "fecha"
            Caption         =   "Fecha"
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
            Caption         =   "Tipo"
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
            DataField       =   "banco"
            Caption         =   "Banco"
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
            DataField       =   "documento"
            Caption         =   "Documento"
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
         BeginProperty Column04 
            DataField       =   "moneda"
            Caption         =   "Moneda"
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
         BeginProperty Column05 
            DataField       =   "importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "saldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "movi"
            Caption         =   "movi"
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
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4215.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   1800
         TabIndex        =   13
         Top             =   120
         Width           =   135
      End
      Begin VB.Label Label3 
         Caption         =   "Del"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Lblsaldoi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Lblal 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   8280
         TabIndex        =   9
         Top             =   180
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo al"
         Height          =   195
         Left            =   7560
         TabIndex        =   8
         Top             =   180
         Width           =   570
      End
      Begin VB.Label Lblnombre 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   7455
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8970
         TabIndex        =   5
         Top             =   120
         Width           =   1815
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
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6165
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   10815
      Begin VB.TextBox txtbuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   540
         Width           =   9375
      End
      Begin VB.CommandButton cmdsaldo 
         Caption         =   "Saldo"
         Height          =   210
         Left            =   9705
         TabIndex        =   23
         Top             =   315
         Visible         =   0   'False
         Width           =   1109
      End
      Begin MSComCtl2.DTPicker Cmbfecha 
         Height          =   315
         Left            =   6720
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   37915
      End
      Begin VB.ComboBox cmbtipo 
         Height          =   315
         Left            =   3030
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   120
         Width           =   2295
      End
      Begin MSDataGridLib.DataGrid Dgrdprestamo 
         Height          =   4935
         Left            =   1395
         TabIndex        =   4
         Top             =   6210
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   8705
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "placod"
            Caption         =   "Codigo"
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
            DataField       =   "nombre"
            Caption         =   "Nombre"
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
            DataField       =   "moneda"
            Caption         =   "Moneda"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "saldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "cuota"
            Caption         =   "Cuota"
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
               Alignment       =   2
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5009.953
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
         EndProperty
      End
      Begin vbalLbar6.vbalListBar vbalListBar1 
         Height          =   5490
         Left            =   -15
         TabIndex        =   24
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   9684
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbalIml6.vbalImageList ilsIcons32 
         Left            =   0
         Top             =   0
         _ExtentX        =   953
         _ExtentY        =   953
         IconSizeX       =   32
         IconSizeY       =   32
         ColourDepth     =   24
         Size            =   48532
         Images          =   "Frmgrdctacte.frx":030A
         Version         =   131072
         KeyCount        =   11
         Keys            =   "FindÿSystemÿExplorerÿFavouritesÿCalendarÿNetwork NeighbourhoodÿHistoryÿInternet ExplorerÿMailÿNewsÿChannels"
      End
      Begin vbalIml6.vbalImageList ilsIcons 
         Left            =   525
         Top             =   0
         _ExtentX        =   953
         _ExtentY        =   953
         ColourDepth     =   32
         Size            =   2296
         Images          =   "Frmgrdctacte.frx":C0BE
         Version         =   131072
         KeyCount        =   2
         Keys            =   "SORTASCÿSORTDESC"
      End
      Begin vbAcceleratorSGrid6.vbalGrid grdLib 
         Height          =   4635
         Left            =   1395
         TabIndex        =   25
         Top             =   855
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   8176
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         ScrollBarStyle  =   2
         DisableIcons    =   -1  'True
      End
      Begin VB.Label Lbltot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   8670
         TabIndex        =   21
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo al "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5940
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1485
         TabIndex        =   17
         Top             =   120
         Width           =   1350
      End
   End
End
Attribute VB_Name = "Frmgrdctacte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsmovi As New Recordset
Dim rssaldo As New Recordset
Dim VTipo As String
Dim nFil As Integer
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim FLAG As Boolean


'**************************
'*** COLUMNAS DE GRILLA ***
Const COL_PERSONA = 1
Const COL_PRESTAMO = 2
Const COL_FECHA = 3
Const COL_TIPO = 4
Const COL_BANCO = 5
Const COL_DOCUMENTO = 6
Const COL_MONEDA = 7
Const COL_IMPORTE = 8
Const COL_BUSQUEDA = 9
'**************************

Private Sub Cmbcia_Click()
Crea_Rs
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01055", "", Cmbtipo)
Cmbtipo.AddItem "TOTAL"
Cmbtipo.ItemData(Cmbtipo.NewIndex) = "99"
Procesa_Prestamos
End Sub

Private Sub Cmbfecha_Change()
Procesa_Prestamos
End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
'Procesa_Prestamos
ProcesaNew
End Sub

Private Sub cmdimporte_Click()
   rsmovi.Sort = "importe asc"
End Sub

Private Sub cmdsaldo_Click()
 rssaldo.Sort = "SALDO asc"
End Sub

Private Sub Command1_Click()
Framemovi.Visible = False
vbalListBar1.Visible = True
End Sub

Private Sub Command2_Click()
'frmdesccta.Show
End Sub

Private Sub Dgrdprestamo_DblClick()
On Error GoTo CORRIGE
 frmactpres.Show
 frmactpres.lblnombre = Trim(rssaldo(1))
 frmactpres.codaux = Trim(rssaldo(5))
 'frmactpres.FECHA = rsmovi("MES") & "/" & rsmovi("DIA") & "/" & rsmovi("ANO")
 frmactpres.PlaCod = Trim(rssaldo(0))
 Exit Sub
CORRIGE:
 MsgBox "Error :" & Err.Description, vbCritical, Me.Caption
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
    ProcesaNew
End If
End Sub

Private Sub Form_Load()
Txtal.Value = Now
Txtdel.Value = Now
FLAG = False
Me.Top = 0
Me.Left = 0
'10920
setUpGrid
Me.Width = 10980
Frame2.Width = 10980
Frame1.Width = 10980
Dgrdprestamo.Width = 10980
Me.Height = 6500
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
LblFecha.Caption = Format(Date, "dd/mm/yyyy")
Cmbfecha.Value = Format(Date, "dd/mm/yyyy")


Dim barX As cListBar
Dim itmX As cListBarItem
Dim i As Long

   With vbalListBar1
      
      .ImageList(evlbLargeIcon) = ilsIcons32
      
      Set barX = .Bars.Add("OPCIONES", , " ")
      Set itmX = barX.Items.Add("prestamos", , "Prestamos", 9)
      itmX.HelpText = "Prestamos Relizados al Personal"
      
      Set itmX = barX.Items.Add("movimientos", , "Movimientos", 2)
      itmX.HelpText = "Movimientos Realizados a Cada Personal"
      
      Set itmX = barX.Items.Add("descuentos", , "Descuentos", 4)
      itmX.HelpText = "Descuentos Relizados al Personal"
               
   End With

End Sub

Public Sub Procesa_Prestamos()
Dim mf1 As String
Dim mf2 As String
Dim f1 As String
Dim f2 As String
Dim Sql$
Dim rs2 As ADODB.Recordset
Dim mSaldo As Currency
Dim mcad As String

On Error GoTo CORRIGE

If rssaldo.RecordCount > 0 Then rssaldo.MoveFirst
Do While Not rssaldo.EOF
   rssaldo.Delete
   rssaldo.MoveNext
Loop

If VTipo = "99" Or Cmbtipo.ListIndex < 0 Then VTipo = ""
If VTipo = "" Then mcad = "" Else mcad = " and p.tipotrabajador='" & VTipo & "' "
lbltot.Caption = "0.00"

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " " & nombre()
Sql = Sql$ & "c.placod,c.moneda,sum(importe-pago_acuenta) as saldo,c.paRTES,C.CODAUXINTERNO from plactacte c,planillas p where c.cia='" & wcia & "' and c.status<>'*' " _
    & "and p.cia=c.cia and c.placod=p.placod and c.importe>0 and p.status<>'*'"
    Sql = Sql & mcad
    Sql = Sql & " and fecha<='" & Format(Cmbfecha.Value, FormatFecha) & FormatTimef & "' and (fecha_cancela >'" & Format(Cmbfecha.Value, FormatFecha) & FormatTimef & "' or fecha_cancela is null) " _
    & "group by c.placod,c.moneda,p.ap_pat,p.ap_mat,p.nom_1,p.nom_2,c.paRTES,C.CODAUXINTERNO"

If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Do While Not rs.EOF
   rssaldo.AddNew
   rssaldo!PlaCod = rs!PlaCod
   rssaldo!nombre = rs!nombre
   rssaldo!moneda = rs!moneda
   mSaldo = rs!saldo
   Sql$ = "Select sum(importe) as saldo from plabolcte where cia='" & wcia & "' and placod='" & Trim(rs!PlaCod) & "' and status<>'*' " _
        & "and fechaproceso>'" & Format(Cmbfecha.Value, FormatFecha) & FormatTimef & "' " _
        & "and fecha_cte<='" & Format(Cmbfecha.Value, FormatFecha) & FormatTimef & "'"
    If (fAbrRst(rs2, Sql)) Then If Not IsNull(rs2(0)) Then mSaldo = mSaldo + rs2(0)
    rs2.Close
    rssaldo!saldo = mSaldo
    lbltot = CCur(lbltot) + mSaldo
    rssaldo!cuota = rs!partes
    rssaldo(5) = rs("CODAUXINTERNO")
   rs.MoveNext
Loop
lbltot = Format(lbltot, "###,###,###.00")
rs.Close
If rssaldo.RecordCount > 0 Then rssaldo.MoveFirst
Dgrdprestamo.Refresh
Screen.MousePointer = vbDefault
FLAG = True
Exit Sub
CORRIGE:
MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub SSCommand1_Click()
Frmplacte.Nuevo_Prestamo
End Sub

Private Sub Txtcod_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SSCommand1.SetFocus
End Sub
Private Sub SSCommand2_Click()

If Dgrdprestamo.ApproxCount = 0 Then
   MsgBox "Ingrese Datos", vbInformation, "Prestamos al Trabajador"
   Exit Sub
End If

lblnombre.Caption = Trim(rssaldo(1))
Framemovi.Visible = True
Framemovi.ZOrder 0
Procesa_Movimiento_Ctacte
'SSPanel1.Visible = False
End Sub
Public Sub Procesa_Movimiento_Ctacte()
Dim saldo As Currency
Dim rs2 As ADODB.Recordset
Dim wciamae As String
Dim cod As String
Dim Sql$
Dim PlaCod As String

On Error GoTo CORRIGE
Screen.MousePointer = 11

cod = "01007"
wciamae = Determina_Maestro(cod)
If wMaeGen = True Then
   wciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
Else
   wciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
End If

saldo = 0
PlaCod = Trim(rssaldo(0))
Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "Select sum(importe-pago_Acuenta) as saldo from plactacte " & _
"where cia='" & wcia & "' and placod='" & Trim(PlaCod) & _
"' and status<>'*' and importe>0 " & _
"and fecha<'" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "' and (fecha_cancela >='" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "' or fecha_cancela is null)"

If (fAbrRst(rs, Sql)) Then If Not IsNull(rs(0)) Then saldo = rs(0)
rs.Close

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "Select sum(importe) as saldo from plabolcte where cia='" & wcia & "' and placod='" & PlaCod & "' and status<>'*' " _
     & "and fechaproceso>='" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "' " _
     & "and fecha_cte<'" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "'"
If (fAbrRst(rs, Sql)) Then If Not IsNull(rs(0)) Then saldo = saldo + rs(0)
rs.Close

Lblal.Caption = Txtdel.Value - 1
Lblsaldoi = Format(saldo, "###,###,###.00")

If rsmovi.RecordCount Then rsmovi.MoveFirst
Do While Not rsmovi.EOF
   rsmovi.Delete
   rsmovi.MoveNext
Loop

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "Select * from plactacte where cia='" & wcia & "' and placod='" & _
PlaCod & "' and status<>'*' and importe>0 " _
& "and fecha>='" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "' " _
& "and fecha<='" & Format(Txtal.Value, FormatFecha) & FormatTimef & "'"
     
If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Do While Not rs.EOF
   If rs!importe > 0 Then
      rsmovi.AddNew
      rsmovi!fecha = Format(rs!fecha, "dd/mm/yyyy")
      
      rsmovi!Tipo = rs!Tipo
      Sql = "select descrip from maestros_2 where cod_maestro2='" & rs!banco & "' " & wciamae
      If (fAbrRst(rs2, Sql)) Then rsmovi!banco = rs2!DESCRIP
      rs2.Close
      rsmovi!documento = rs!num_doc
      rsmovi!moneda = rs!moneda
      rsmovi!importe = rs!importe
      saldo = saldo + rs!importe
      rsmovi!saldo = saldo
      rsmovi!movi = "I"
      rsmovi!Dia = Mid(rsmovi!fecha, 1, 2)
      rsmovi!Mes = Mid(rsmovi!fecha, 4, 2)
      rsmovi!ano = Mid(rsmovi!fecha, 7, 4)
   End If
   rs.MoveNext
Loop
rs.Close

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "Select fechaproceso,d07 from plahistorico where cia='" & _
wcia & "' and placod='" & PlaCod & _
"' and status<>'*' and d07<>0 " _
& "and fechaproceso>='" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "' " _
& "and fechaproceso<='" & Format(Txtal.Value, FormatFecha) & FormatTimef & "'"
     
If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Do While Not rs.EOF
   If rs!d07 > 0 Then
      rsmovi.AddNew
      rsmovi!fecha = Format(rs!FechaProceso, "dd/mm/yyyy")
      rsmovi!Tipo = "PL"
      rsmovi!documento = ""
      rsmovi!moneda = wmoncont
      rsmovi!importe = rs!d07 * -1
      saldo = saldo - rs!d07
      rsmovi!saldo = saldo
      rsmovi!movi = "S"
      rsmovi!Dia = Mid(rsmovi!fecha, 1, 2)
      rsmovi!Mes = Mid(rsmovi!fecha, 4, 2)
      rsmovi!ano = Mid(rsmovi!fecha, 7, 4)
      
   End If
   rs.MoveNext
Loop
rs.Close

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql & "Select * from plactacte where cia='" & wcia & "' and placod='" & _
PlaCod & "' and status<>'*' and importe<0 " _
& "and fecha>='" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "' " _
& "and fecha<='" & Format(Txtal.Value, FormatFecha) & FormatTimef & "'"
     
If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Do While Not rs.EOF
   If Abs(rs!importe) > 0 Then
      rsmovi.AddNew
      rsmovi!fecha = Format(rs!fecha, "dd/mm/yyyy")
      rsmovi!Tipo = "DV"
      Sql = "select descrip from maestros_2 where cod_maestro2='" & _
      rs!banco & "' " & wciamae
      If (fAbrRst(rs2, Sql)) Then rsmovi!banco = rs2!DESCRIP
      rs2.Close
      rsmovi!documento = ""
      rsmovi!moneda = rs!moneda
      rsmovi!importe = rs!importe
      saldo = saldo + rs!importe
      rsmovi!saldo = saldo
      rsmovi!movi = "S"
      rsmovi!Dia = Mid(rsmovi!fecha, 1, 2)
      rsmovi!Mes = Mid(rsmovi!fecha, 4, 2)
      rsmovi!ano = Mid(rsmovi!fecha, 7, 4)
   End If
   rs.MoveNext
Loop
rs.Close
rsmovi.Sort = "ano,mes,dia,movi"
Screen.MousePointer = vbDefault
Exit Sub
CORRIGE:
       MsgBox "Error :" & Err.Description, vbCritical, Me.Caption
End Sub
Private Sub Crea_Rs()
    If rsmovi.State = 1 Then rsmovi.Close
    rsmovi.Fields.Append "fecha", adChar, 10, adFldIsNullable
    rsmovi.Fields.Append "tipo", adChar, 2, adFldIsNullable
    rsmovi.Fields.Append "banco", adChar, 40, adFldIsNullable
    rsmovi.Fields.Append "documento", adChar, 25, adFldIsNullable
    rsmovi.Fields.Append "moneda", adChar, 3, adFldIsNullable
    rsmovi.Fields.Append "importe", adCurrency, 18, adFldIsNullable
    rsmovi.Fields.Append "saldo", adCurrency, 18, adFldIsNullable
    rsmovi.Fields.Append "movi", adChar, 2, adFldIsNullable
    rsmovi.Fields.Append "dia", adChar, 2, adFldIsNullable
    rsmovi.Fields.Append "mes", adChar, 2, adFldIsNullable
    rsmovi.Fields.Append "ano", adChar, 4, adFldIsNullable
    rsmovi.Open
    Set DgrdMovi.DataSource = rsmovi
    
    If rssaldo.State = 1 Then rssaldo.Close
    rssaldo.Fields.Append "placod", adChar, 8, adFldIsNullable
    rssaldo.Fields.Append "nombre", adChar, 40, adFldIsNullable
    rssaldo.Fields.Append "moneda", adChar, 3, adFldIsNullable
    rssaldo.Fields.Append "saldo", adCurrency, 18, adFldIsNullable
    rssaldo.Fields.Append "Cuota", adCurrency, 18, adFldIsNullable
    rssaldo.Fields.Append "codauxinterno", adVarChar, 30, adFldIsNullable
    rssaldo.Open
    Set Dgrdprestamo.DataSource = rssaldo
   
End Sub

Private Sub grdLib_ColumnWidthChanged(ByVal lCol As Long, lWidth As Long, bCancel As Boolean)
bCancel = True
End Sub

Private Sub grdLib_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, bDoDefault As Boolean)
Dim Fila As Long
Dim nombre As String
Dim partes As String
Dim ipos As Long
Dim rsdata As ADODB.Recordset
Dim sSQL As String

Dim codigo As String

If Button = 2 Then
    Fila = grdLib.SelectedRow
    If Fila > 0 Then
        If grdLib.RowIsGroup(Fila) Then
            If grdLib.RowGroupingLevel(Fila) = 2 Then
                nombre = grdLib.CellText(Fila - 1, 1)
                partes = grdLib.CellText(Fila, 2)
                     
                ipos = InStr(1, partes, "Cuotas de : ")
                
                frmactpres.PlaCod = Trim(Mid(nombre, InStr(1, nombre, " - ") + 3, Len(nombre) - InStr(1, nombre, " - ")))
                frmactpres.lblnombre.Caption = nombre
                frmactpres.txtcuota.Text = Format(Trim(Mid(partes, ipos + 12, Len(partes) - ipos)), "#0.00")
                frmactpres.txtcuota.Tag = Trim(Mid(partes, 14, 6))
                
                sSQL = "SELECT sn_grati FROM plactacte WHERE placod='" & Trim(Mid(nombre, InStr(1, nombre, " - ") + 3, Len(nombre) - InStr(1, nombre, " - "))) & "' and cia='" & wcia & "' AND id_doc=" & Trim(Mid(partes, 14, 6))

                Set rsdata = cn.Execute(sSQL)
                
                If Not rsdata.EOF Then
                    If rsdata(0) = True Then
                        frmactpres.chkgrati.Value = vbChecked
                    Else
                        frmactpres.chkgrati.Value = vbUnchecked
                    End If
                    rsdata.Close
                End If
                
                Set rsdata = Nothing
                
                frmactpres.Show
            End If
        End If
    End If
End If

Set rsdata = Nothing

End Sub

Private Sub Txtal_Change()
Procesa_Movimiento_Ctacte
End Sub

Private Sub txtbuscar_Change()
Dim ipos As Long
ipos = grdLib.FindSearchMatchRow(txtbuscar.Text, True, False)
If ipos > 0 Then grdLib.SelectedRow = ipos
End Sub

Private Sub Txtdel_Change()
Procesa_Movimiento_Ctacte
End Sub
Public Sub Imprime_Saldos()
Dim mTipo As String
If rssaldo.RecordCount <= 0 Then Exit Sub
Screen.MousePointer = 11
If Cmbtipo.ListIndex < 0 Then mTipo = "TOTAL" Else mTipo = Cmbtipo.Text
rssaldo.MoveFirst
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("B:B").ColumnWidth = 40
xlSheet.Range("C:C").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 1).Value = Cmbcia.Text
xlSheet.Cells(1, 1).Font.Size = 12
xlSheet.Cells(1, 1).Font.Bold = True

xlSheet.Cells(3, 1).Value = "CUENTA CORRIENTE DEL PERSONAL AL " & Cmbfecha.Value
xlSheet.Range("A3:C3").Merge
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(4, 1).Value = mTipo
xlSheet.Cells(4, 1).Font.Bold = True
xlSheet.Range("A4:C4").Merge

xlSheet.Cells(6, 1).Value = "CODIGO"
xlSheet.Cells(6, 2).Value = "NOMBRE"
xlSheet.Cells(6, 3).Value = "SALDO"
xlSheet.Range("A6:C6").Font.Bold = True
xlSheet.Range("A3:C6").HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(6, 3)).Borders.LineStyle = xlContinuous

nFil = 7
Do While Not rssaldo.EOF
   xlSheet.Cells(nFil, 1).Value = rssaldo!PlaCod
   xlSheet.Cells(nFil, 2).Value = rssaldo!nombre
   xlSheet.Cells(nFil, 3).Value = rssaldo!saldo
   nFil = nFil + 1
   rssaldo.MoveNext
Loop
nFil = nFil + 1
xlSheet.Cells(nFil, 3).Value = lbltot.Caption

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "SALDOS DE CTA.CTE."
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = 0

End Sub
Public Sub Imprime_Movimientos()
Dim mTipo As String
If rsmovi.RecordCount <= 0 Then Exit Sub
Screen.MousePointer = 11
rsmovi.MoveFirst
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Range("B:B").ColumnWidth = 5
xlSheet.Range("C:C").ColumnWidth = 21
xlSheet.Range("D:D").ColumnWidth = 17
xlSheet.Range("E:E").ColumnWidth = 8

xlSheet.Range("F:G").NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"

xlSheet.Cells(1, 1).Value = Cmbcia.Text
xlSheet.Cells(1, 1).Font.Size = 12
xlSheet.Cells(1, 1).Font.Bold = True

xlSheet.Cells(3, 1).Value = "MOVIMIENTOS DE CTA. CTE. DEL PERSONAL "
xlSheet.Range("A3:G3").Merge
xlSheet.Range("A3:A3").HorizontalAlignment = xlCenter
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(4, 1).Value = "DEL " & Txtdel.Value & " AL " & Txtal.Value
xlSheet.Cells(4, 1).Font.Bold = True
xlSheet.Range("A4:G4").Merge
xlSheet.Range("A4:A4").HorizontalAlignment = xlCenter

xlSheet.Cells(6, 1).Value = Dgrdprestamo.Columns(0)
xlSheet.Cells(6, 2).Value = lblnombre.Caption
xlSheet.Cells(6, 2).Font.Bold = True
xlSheet.Cells(6, 5).Value = Label2.Caption & " " & Lblal.Caption
xlSheet.Cells(6, 7).Value = Lblsaldoi.Caption
xlSheet.Cells(6, 7).Font.Bold = True

xlSheet.Cells(8, 1).Value = "FECHA"
xlSheet.Cells(8, 2).Value = "TIPO"
xlSheet.Cells(8, 3).Value = "BANCO"
xlSheet.Cells(8, 4).Value = "DOCUMENTO"
xlSheet.Cells(8, 5).Value = "MONEDA"
xlSheet.Cells(8, 6).Value = "IMPORTE"
xlSheet.Cells(8, 7).Value = "SALDO"
xlSheet.Range("A8:G8").Font.Bold = True
xlSheet.Range("A8:G8").HorizontalAlignment = xlCenter

xlSheet.Range(xlSheet.Cells(8, 1), xlSheet.Cells(8, 7)).Borders.LineStyle = xlContinuous

nFil = 9
Do While Not rsmovi.EOF
   xlSheet.Cells(nFil, 1).Value = rsmovi!fecha
   xlSheet.Cells(nFil, 2).Value = rsmovi!Tipo
   xlSheet.Cells(nFil, 3).Value = rsmovi!banco
   xlSheet.Cells(nFil, 4).Value = rsmovi!documento
   xlSheet.Cells(nFil, 5).Value = rsmovi!moneda
   xlSheet.Cells(nFil, 6).Value = rsmovi!importe
   xlSheet.Cells(nFil, 7).Value = rsmovi!saldo
   nFil = nFil + 1
   rsmovi.MoveNext
Loop
nFil = nFil + 1

xlApp2.Application.ActiveWindow.DisplayGridLines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "MOVIMIENTOS DE CTA.CTE."
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Screen.MousePointer = 0

End Sub

Private Sub vbalListBar1_ItemClick(Item As vbalLbar6.cListBarItem, Bar As vbalLbar6.cListBar)
   
   Select Case Item.Key
   Case "prestamos"
        Frmplacte.Nuevo_Prestamo
   Case "movimientos"
        If Dgrdprestamo.ApproxCount = 0 Then
           MsgBox "Ingrese Datos", vbInformation, "Prestamos al Trabajador"
           Exit Sub
        End If
        
        lblnombre.Caption = Trim(rssaldo(1))
        Framemovi.Visible = True
        Framemovi.ZOrder 0
        Procesa_Movimiento_Ctacte
        vbalListBar1.Visible = False
   Case "descuentos"
   
   End Select
End Sub

Public Sub ProcesaNew()
Dim sSQL As String
Dim rs As ADODB.Recordset

sSQL = "EXEC sp_c_ctacte '" & wcia & "','" & VTipo & "','" & Format(Txtdel.Value, FormatFecha) & FormatTimei & "'"

Set rs = cn.Execute(sSQL)

grdLib.Redraw = False
grdLib.ColumnIsGrouped(1) = False
grdLib.ColumnIsGrouped(2) = False
grdLib.Clear
If Not rs.EOF Then
    Do While Not rs.EOF
        With grdLib
            .AddRow
            .CellDetails .Rows, COL_PERSONA, Trim(CStr(rs!tiptrab)) & "  " & Trim(CStr(rs!nombre)) & " -  " & Trim(UCase(CStr(rs!PlaCod))), DT_LEFT
            .CellDetails .Rows, COL_PRESTAMO, "N° Prestamo : " & rs!id_doc & Space(5) & "  Moneda : " & CStr(rs!monprest) & Space(5) & "  Importe : " & Format(CStr(rs!impprest), "###,###,###.00") & Space(5) & "  Saldo : " & Format((rs!impprest - rs!pago_acuenta), "#0.00") & Space(5) & " Cuotas de : " & rs!partes, DT_LEFT
            .CellDetails .Rows, COL_FECHA, Format(Trim(rs!FechaProceso & " "), "DD/MM/YYYY"), DT_LEFT
            .CellDetails .Rows, COL_TIPO, Trim(CStr(rs!Tipo)), DT_CENTER
            .CellDetails .Rows, COL_MONEDA, Trim(CStr(rs!monpago)), DT_CENTER
            .CellDetails .Rows, COL_IMPORTE, Format(Trim(CStr(rs!imppago)), "###,###,###.00"), DT_RIGHT
            .CellDetails .Rows, COL_BUSQUEDA, Trim(CStr(rs!nombre)), DT_LEFT
        End With
        rs.MoveNext
    Loop
    
    grdLib.ColumnVisible("persona") = True
    grdLib.ColumnVisible("prestamo") = True
    
    grdLib.ColumnIsGrouped(1) = True
    grdLib.ColumnIsGrouped(2) = True

End If



Dim i As Long

For i = 1 To grdLib.Rows
    If grdLib.RowIsGroup(i) Then
        If grdLib.RowGroupingState(i) = ecgExpanded Then grdLib.RowGroupingState(i) = ecgCollapsed
    End If
Next i

grdLib.Redraw = True

End Sub

Private Sub setUpGrid()
   
   ' Set general options:
   With grdLib
      .HideGroupingBox = True
      .AllowGrouping = True
       
      .BackColor = RGB(255, 255, 235)
      '.AlternateRowBackColor = RGB(255, 255, 185)
      .GroupRowBackColor = RGB(180, 173, 176)
      .GroupingAreaBackColor = .BackColor
      .ForeColor = RGB(0, 0, 0)
      .GroupRowForeColor = .ForeColor
      .HighlightForeColor = vbWindowText
      .HighlightBackColor = RGB(255, 255, 205)
      .NoFocusHighlightBackColor = RGB(200, 200, 200)
      .SelectionAlphaBlend = True
      .SelectionOutline = True
      .DrawFocusRectangle = False
      .HighlightSelectedIcons = False
      .HotTrack = True
      .RowMode = True
      .MultiSelect = True
   
   
      ' Add the columns:
      .AddColumn "persona", "", , , 192, False, , , , , , CCLSortStringNoCase
      .AddColumn "prestamo", "", , , 128, False, , , , , , CCLSortStringNoCase
      .AddColumn "fecha", "Fecha", , , 100, True, , , , , , CCLSortStringNoCase
      .AddColumn "tipo", "Tipo", ecgHdrTextALignLeft, , 70, True, , , , , , CCLSortStringNoCase
      .AddColumn "banco", "Banco", , , 125, True, , , , , , CCLSortStringNoCase
      .AddColumn "documento", "Documento", , , 105, True, , , , , , CCLSortBackColor
      .AddColumn "moneda", "Moneda", , , 70, True, , , , , , CCLSortStringNoCase
      .AddColumn "importe", "Importe", ecgHdrTextALignLeft, True, 110, , , , , , , CCLSortNumeric
      .AddColumn "BUSCAR", "", , , 0, True
      
      .KeySearchColumn = .ColumnIndex("BUSCAR")
        .SetHeaders
      
      .HeaderImageList = ilsIcons
      
      .StretchLastColumnToFit = True
      
   End With
   
End Sub

