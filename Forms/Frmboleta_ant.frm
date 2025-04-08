VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmboleta 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Boletas"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   Icon            =   "Frmboleta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "Descuentos Adicionales"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   4680
      TabIndex        =   4
      Top             =   5160
      Width           =   4335
      Begin MSDataGridLib.DataGrid DgrdDesAdic 
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2990
         _Version        =   393216
         Appearance      =   0
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "monto"
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
         BeginProperty Column02 
            DataField       =   "codigo"
            Caption         =   "codigo"
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
               Locked          =   -1  'True
               ColumnWidth     =   2640.189
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pagos Adicionales"
      Enabled         =   0   'False
      Height          =   3495
      Left            =   4680
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
      Begin MSDataGridLib.DataGrid DgrdPagAdic 
         Height          =   3135
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5530
         _Version        =   393216
         Appearance      =   0
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
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
            DataField       =   "monto"
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
         BeginProperty Column02 
            DataField       =   "codigo"
            Caption         =   "codigo"
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
               Locked          =   -1  'True
               ColumnWidth     =   2580.095
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Centro de Costo"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   4335
      Begin VB.ListBox LstCcosto 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   435
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid Dgrdccosto 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Centro de Costo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "monto"
            Caption         =   "Porcentaje"
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
         BeginProperty Column02 
            DataField       =   "codigo"
            Caption         =   "codigo"
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
               Button          =   -1  'True
               ColumnWidth     =   2654.929
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Lbltot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Constantes de Calculo"
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
      Begin MSDataGridLib.DataGrid Dgrdhoras 
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   6165
         _Version        =   393216
         Appearance      =   0
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "descripcion"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Monto"
            Caption         =   "Horas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "codigo"
            Caption         =   "codigo"
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
               Locked          =   -1  'True
               ColumnWidth     =   2475.213
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin MSMask.MaskEdBox Txtvacaf 
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   645
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txtvacai 
         Height          =   255
         Left            =   1440
         TabIndex        =   24
         Top             =   630
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Cmbturno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox Txtvaca 
         Height          =   255
         Left            =   7680
         TabIndex        =   15
         Top             =   645
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Txtcese 
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   645
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox Txtcodpla 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3465
         MaxLength       =   8
         TabIndex        =   10
         Top             =   225
         Width           =   855
      End
      Begin VB.Label Lblbasico 
         Height          =   15
         Left            =   7200
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Lblnumafp 
         Height          =   135
         Left            =   5640
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Lblcargo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6600
         TabIndex        =   27
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Lblvaca1 
         AutoSize        =   -1  'True
         Caption         =   "F. Ret. Vac."
         Height          =   195
         Left            =   3480
         TabIndex        =   23
         Top             =   660
         Width           =   855
      End
      Begin VB.Label LblFingreso 
         Height          =   135
         Left            =   3120
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Lblcodaux 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Periodo Vacacional"
         Height          =   195
         Left            =   5760
         TabIndex        =   14
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Lblvaca2 
         AutoSize        =   -1  'True
         Caption         =   "F. Inicio Vaca."
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   690
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Lblnombre 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Lblcodafp 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Lbltope 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5640
         TabIndex        =   26
         Top             =   240
         Width           =   45
      End
   End
   Begin Threed.SSPanel Panelprogress 
      Height          =   735
      Left            =   1680
      TabIndex        =   32
      Top             =   3360
      Visible         =   0   'False
      Width           =   5655
      _Version        =   65536
      _ExtentX        =   9975
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "Generando Vacaciones Devengadas"
      ForeColor       =   8388608
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Label Lblctacte 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Caption         =   "0.00"
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
      Height          =   255
      Left            =   7320
      TabIndex        =   31
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Saldo Cta.Cte.                   S/."
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
      Height          =   255
      Left            =   4680
      TabIndex        =   30
      Top             =   4800
      Width           =   2655
   End
End
Attribute VB_Name = "Frmboleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VokDevengue As Boolean
Dim vDevengue As Boolean
Dim Mstatus As String

Dim BolDevengada As Boolean
Dim NumDev As Integer

Public rshoras As New Recordset
Dim rsccosto As New Recordset
Dim rspagadic As New Recordset
Public rsdesadic As New Recordset
Dim VTipobol As String
Dim vTipoTra As String
Dim VTurno As String
Dim VItem As Integer
Dim VSemana As String
Dim VFProceso As String
Dim VPerPago As String
Dim VNewBoleta As Boolean
Dim Vano As Integer
Dim Vmes As Integer
Dim Vdia As Integer
Dim VHoras As Currency
Dim VHorasnormal As Currency
Dim VfDel As String
Dim VfAl As String
Dim mdiasfalta As String
Dim VAltitud As String
Dim VVacacion As String
Dim VArea As String
Dim macui As Currency
Dim macus As Currency
Dim mcancel As Boolean, MSINDICATO As Boolean
Dim VFechaNac As String
Dim VFechaJub As String
Dim VObra As String
Dim manos As Integer
Dim mHourDay As Currency

Dim ArrDsctoCTACTE() As Variant
Dim MAXROW As Long
Dim BoletaCargada As Boolean
Dim sn_essaludvida As Byte, sn_sindicato As Byte

Dim sn_quinta As Boolean

Private Sub Cmbturno_Click()
VTurno = Funciones.fc_CodigoComboBox(Cmbturno, 2)
End Sub

Private Sub Dgrdccosto_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
    Case Is = 1
         Dgrdccosto.Columns(1) = Format(Dgrdccosto.Columns(1), "###,###.00")
         Total_porcentaje ("N")
End Select
End Sub

Private Sub Dgrdccosto_AfterDelete()
Total_porcentaje ("S")
End Sub

Private Sub Dgrdccosto_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If Dgrdccosto.Col = 0 Then
        KeyAscii = 0
        Cancel = True
        Dgrdccosto_ButtonClick (ColIndex)
End If
End Sub

Private Sub Dgrdccosto_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer

Y = Dgrdccosto.Row
xtop = Dgrdccosto.Top + Dgrdccosto.RowTop(Y) + Dgrdccosto.RowHeight
Select Case ColIndex
Case 0:
       xleft = Dgrdccosto.Left + Dgrdccosto.Columns(0).Left
       With LstCcosto
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = Dgrdccosto.Top + Dgrdccosto.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub
Private Sub DgrdDesAdic_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 1
            If DgrdDesAdic.Columns(2) = "07" And CCur(Val(DgrdDesAdic.Columns(1))) > CCur(Format(Lblctacte.Caption, "########0.00")) Then
               MsgBox "El Importe no debe ser Mayor al Saldo", vbInformation, "Cuenta Corriente"
               DgrdDesAdic.Columns(1) = "0.00"
            Else
                ProrrateaCtaCte CCur(Val(DgrdDesAdic.Columns(1)))
            End If
End Select
End Sub

Private Sub DgrdDesAdic_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'**********DESHABILITAR PARA PPM****************************
If DgrdDesAdic.Columns(2) = "09" Then KeyAscii = 0
'***********************************************************
End Sub

Private Sub DgrdDesAdic_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'**********DESHABILITAR PARA PPM****************************
If DgrdDesAdic.Columns(2) = "09" Then Cancel = True
'***********************************************************
End Sub

Private Sub DgrdDesAdic_DblClick()
' Call Frmplacte.Show
' Frmplacte.Txtcodpla = Txtcodpla
' Frmplacte.proc = 1 'DE DONDE VIENE
'
' Call frmdesccta.Show
' frmdesccta.proc = 1
' frmdesccta.txtcod.Text = Me.Txtcodpla
' Call frmdesccta.Txtcod_KeyPress(13)
 
End Sub

Private Sub DgrdDesAdic_KeyPress(KeyAscii As Integer)
'**********DESHABILITAR PARA PPM GIOVANNI*******************
If Format(DgrdDesAdic.Columns(2) & "", "00") = "09" Then KeyAscii = 0
'***********************************************************
End Sub

Private Sub DgrdHoras_AfterColEdit(ByVal ColIndex As Integer)
If Dgrdhoras.Columns(1).Text = "" Then Dgrdhoras.Columns(1) = "0.00"
Dgrdhoras.Columns(1) = Format(Dgrdhoras.Columns(1), "###,###.00")
If Not rshoras.EOF Then rshoras.MoveNext
End Sub
'
'Private Sub Dgrdhoras_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
''If vTipoTra = "02" And rshoras!codigo = "01" Then Cancel = True
''If vTipoTra = "02" And rshoras!codigo = "02" Then Cancel = True
'End Sub

Private Sub Dgrdhoras_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim VALOR  As Currency

If vTipoTra <> "02" Or rshoras!codigo <> 14 Then Exit Sub
VALOR = Dgrdhoras.Columns(1)

Do While Not rshoras.EOF
    If rshoras!codigo = "01" Then
        rshoras!Monto = VALOR * 8
        Exit Do
    End If
    rshoras.MoveNext
Loop

rshoras.MoveFirst

Do While Not rshoras.EOF
    If rshoras!codigo = "02" Then
        rshoras!Monto = (VALOR / 6) * 8
        Exit Do
    End If
    rshoras.MoveNext
Loop

rshoras.MoveFirst
End Sub

Private Sub Form_Activate()
If wTipoDoc = True Then
   Me.Caption = "Ingreso de Boletas"
   Frmboleta.BackColor = &H80000001
Else
   Me.Caption = "Adelanto de Quincena"
   Frmboleta.BackColor = &H808000
End If
End Sub

Private Sub Form_Load()
Dim wciamae As String
Me.Top = 0
Me.Left = 0
Me.Width = 9135
Me.Height = 7680
Crea_Rs

wciamae = Determina_Maestro("01076")
sql$ = "Select flag2 from maestros_2 where cod_maestro2='01' and status<>'*'"
sql$ = sql$ & wciamae
mHourDay = 0
If (fAbrRst(rs, sql$)) Then mHourDay = Val(rs!flag2)
rs.Close

BoletaCargada = False
End Sub
Private Sub Crea_Rs()
    If rshoras.State = 1 Then rshoras.Close
    rshoras.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rshoras.Fields.Append "descripcion", adChar, 100, adFldIsNullable
    rshoras.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    'If Not VNewBoleta Then rshoras.LockType = adLockReadOnly
    rshoras.Open
    Set Dgrdhoras.DataSource = rshoras
    
    If rsccosto.State = 1 Then rsccosto.Close
    rsccosto.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsccosto.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsccosto.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rsccosto.Fields.Append "item", adChar, 2, adFldIsNullable
    'If Not VNewBoleta Then rsccosto.LockType = adLockReadOnly
    rsccosto.Open
    Set Dgrdccosto.DataSource = rsccosto
    
    If rspagadic.State = 1 Then rspagadic.Close
    rspagadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rspagadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rspagadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    'If Not VNewBoleta Then rspagadic.LockType = adLockReadOnly
    rspagadic.Open
    Set DgrdPagAdic.DataSource = rspagadic
    
    If rsdesadic.State = 1 Then rsdesadic.Close
    rsdesadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsdesadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsdesadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    'If Not VNewBoleta Then rsdesadic.LockType = adLockReadOnly
    rsdesadic.Open
    Set DgrdDesAdic.DataSource = rsdesadic
End Sub
Private Sub Procesa()

'Pagos Adicionales
sql$ = "Select * from placonstante where cia='" & Trim(wcia) & "' and tipomovimiento='02' and calculo='N' and status<>'*' order by codinterno"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(sql$, 64)
If rspagadic.RecordCount > 0 Then
   rspagadic.MoveFirst
   Do While Not rspagadic.EOF
      rspagadic.Delete
      rspagadic.MoveNext
   Loop
End If

If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rspagadic.AddNew
   rspagadic!codigo = rs!codinterno
   rspagadic!Descripcion = rs!Descripcion
   rspagadic!Monto = "0.00"
   rs.MoveNext
Loop

If rs.State = 1 Then rs.Close

'Descuentos Adicionales
If wTipoDoc = True Then
   sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='03' and adicional='S' and status<>'*' order by codinterno"
Else
   sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='03' and adicional='S' and status<>'*' and codinterno<>'09' order by codinterno"
End If
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(sql$, 64)
If rsdesadic.RecordCount > 0 Then
   rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      rsdesadic.Delete
      rsdesadic.MoveNext
   Loop
End If

If Not rs.RecordCount > 0 Then Exit Sub
rs.MoveFirst
Do While Not rs.EOF
   rsdesadic.AddNew
   rsdesadic!codigo = rs!codinterno
   rsdesadic!Descripcion = Trim(rs!Descripcion)
   rsdesadic!Monto = "0.00"
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

If Not sn_quinta Then
    rsdesadic.AddNew
    rsdesadic!codigo = "13"
    rsdesadic!Descripcion = "QUINTA CATEGORIA"
    rsdesadic!Monto = "0.00"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmCabezaBol.Procesa_Cabeza_Boleta
End Sub

Private Sub LstCcosto_Click()
Dim m As Integer

If LstCcosto.ListIndex > -1 Then
   m = Len(LstCcosto.Text) - 3
    
   Dgrdccosto.Columns(0) = Trim(Left(LstCcosto.Text, m))
   Dgrdccosto.Columns(2) = Format(Right(LstCcosto.Text, 2), "00")
   Dgrdccosto.Col = 1
   Dgrdccosto.SetFocus
   LstCcosto.Visible = False
End If
End Sub

Private Sub LstCcosto_LostFocus()
LstCcosto.Visible = False
End Sub


Private Sub Txtcodpla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Frmgrdpla.Show
    Frmgrdpla.ZOrder 0
End If
End Sub

Public Sub Txtcodpla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Cmbturno.SetFocus
End Sub
Private Sub Txtcodpla_LostFocus()
Dim xciamae As String
Dim cod As String
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

On Error GoTo CORRIGE

BolDevengada = False
cn.CursorLocation = adUseClient

Set rs = New ADODB.Recordset
cod = "01055"
Txtcodpla.Text = UCase(Txtcodpla.Text)
' TIPO DE TRABAJADOR
sql$ = "SELECT GENERAL FROM maestros where " & _
"right(ciamaestro,3)='" & Right(cod, 3) & "' and " & _
"status<>'*' "

If (fAbrRst(rs, sql$)) Then
   If rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If

If rs.State = 1 Then rs.Close

'OBTENER NOMBRE DE EMPLEADO
sql$ = Funciones.nombre()
sql$ = sql$ & "codauxinterno,a.status,a.tipotrabajador,a.fingreso," & _
     "a.fcese,a.codafp,a.numafp,a.area,a.placod," & _
     "a.codauxinterno,b.descrip,a.tipotasaextra," & _
     "a.cargo,a.altitud,a.vacacion,a.area,a.fnacimiento," & _
     "a.fec_jubila,a.sindicato,a.ESSALUDVIDA,a.quinta " & _
     "from planillas a,maestros_2 b where a.status<>'*' "
     sql$ = sql$ & xciamae
     sql$ = sql$ & " and a.tipotrabajador=b.cod_maestro2 " _
     & "and cia='" & wcia & "' AND placod='" & Trim(Txtcodpla.Text) & "' "
     sql$ = sql$ & " order by nombre"

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(sql$)

If rs.RecordCount > 0 Then
   If rs!TipoTrabajador <> vTipoTra Or (Not IsNull(rs!fcese) And VTipobol <> "04") Then
      If rs!TipoTrabajador <> vTipoTra Then MsgBox "Trabajador no es del tipo seleccionado", vbExclamation, "Codigo N° => " & Txtcodpla.Text
      If Not IsNull(rs!fcese) Then MsgBox "Trabajador ya fue Cesado", vbExclamation, "Con Fecha => " & Format(rs!fcese, "dd/mm/yyyy")
      Txtcodpla.Text = ""
      Limpia_Boleta
      lblnombre.Caption = ""
      Lblctacte.Caption = "0.00"
      Lblcodaux.Caption = ""
      Lblcodafp.Caption = ""
      Lblnumafp.Caption = ""
      Lblbasico.Caption = ""
      Lbltope.Caption = ""
      Lblcargo.Caption = ""
      LblFingreso.Caption = ""
      VAltitud = ""
      VVacacion = ""
      VArea = ""
      VFechaNac = ""
      VFechaJub = ""
      Txtcodpla.SetFocus
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      Frame5.Enabled = False
   Else
      LblFingreso.Caption = Format(rs!fingreso, "mm/dd/yyyy")
      If Val(Right(LblFingreso.Caption, 4)) > Vano Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
         Limpia_Boleta
         LblFingreso.Caption = ""
         Exit Sub
      ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) > Vmes Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
         Limpia_Boleta
         LblFingreso.Caption = ""
         Exit Sub
      ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) = Vmes And Val(Mid(LblFingreso.Caption, 4, 2)) > Val(Left(VFProceso, 2)) Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
         Limpia_Boleta
         LblFingreso.Caption = ""
         Exit Sub
      End If
      
      If Month(CDate(VFProceso)) = 7 Or Month(CDate(VFProceso)) = 12 Then
        If vTipoTra = "01" Then
          If Not wTipoDoc Then
            sql = "SELECT COUNT(*) FROM PLAHISTORICO WHERE CIA='" & wcia & "' AND PLACOD='" & Txtcodpla.Text & "' AND status!='*' and proceso='03' and month(fechaproceso)=" & Month(VFProceso) & " and year(fechaproceso)=" & Year(VFProceso)
            Set rs3 = cn.Execute(sql$)
            
            If Not rs3.EOF Then
              If rs3(0) = 0 Then
                  MsgBox "Primero debe Generar la Boleta de Gratificacion", vbExclamation, "Planilla"
                  Unload Me
                  Exit Sub
              End If
            End If
          End If
        End If
      End If
      
      
      'cargamos si la person esta afecto o no a la quinta categoria
      sn_quinta = True
      If Trim(rs!quinta & "") = "N" Or Trim(rs!quinta & "") = "" Then sn_quinta = False
      '***********************************************************
      
      lblnombre.Caption = rs!nombre
      Lblcodaux.Caption = rs!codauxinterno
      Lblcodafp.Caption = rs!CodAfp
      Lblnumafp.Caption = Trim(rs!NUMAFP)
      VFechaNac = Format(rs!fnacimiento, "dd/mm/yyyy")
      VFechaJub = Format(rs!fec_jubila, "dd/mm/yyyy")
      Lbltope.Caption = rs!tipotasaextra
      
      If vTipoTra = "05" Then Lblcargo.Caption = rs!cargo: VAltitud = rs!altitud: VVacacion = rs!vacacion
      
      VArea = Trim(rs!Area)
      Frame2.Enabled = True
      Frame3.Enabled = True
      Frame4.Enabled = True
      Frame5.Enabled = True
      
      'Basico  de la persona
      
      sql$ = "select importe from plaremunbase where cia='" & wcia & "' " & _
      "and placod='" & Trim(Txtcodpla.Text) & "' and concepto='01' and status<>'*'"
     
      Lblbasico.Caption = ""
      If (fAbrRst(rs2, sql$)) Then Lblbasico.Caption = rs2!importe
      If rs2.State = 1 Then rs2.Close
      'Centro de Costo
      If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
      Do While Not rsccosto.EOF
        rsccosto.Delete
        rsccosto.MoveNext
      Loop
      
      '=======================
       Dim RX As New ADODB.Recordset
       
       sn_essaludvida = 0: snmanual = 0
       
       If Trim(UCase(rs("ESSALUDVIDA"))) = "S" Then
          rsdesadic.MoveFirst
          sn_essaludvida = 1
         rsdesadic.FIND "CODIGO='06'", 1 'ESSALUDVIDA
         con = "select DEDUCCION,adicional from placonstante where cia='01' and tipomovimiento='03' and " & _
         "codinterno='06' and deduccion<>0 and status<>'*'"
         RX.Open con, cn, adOpenStatic, adLockReadOnly
         If RX.RecordCount > 0 Then
            If Not rsdesadic.EOF Then rsdesadic("MONTO") = RX("DEDUCCION"): snmanual = IIf(RX!adicional = "S", 0, 1)
         End If
         RX.Close
          Call Acumula_Mes("06", "D")
          If rsdesadic("MONTO") > 0 Then
            rsdesadic("MONTO") = rsdesadic("MONTO") - macui
          End If
      End If
        
      MSINDICATO = False
      sn_sindicato = 0
      If Trim(UCase(rs("sindicato"))) = "S" Then
         'rsdesadic.Filter = "CODIGO='15'"
         rsdesadic.FIND "CODIGO='15'", 1 'SOLIC SINDICATO
         con = "select deduccion from placonstante where cia='01' and tipomovimiento='03' " & _
         " and codinterno='15' and deduccion<>0 and status<>'*'"
         RX.Open con, cn, adOpenStatic, adLockReadOnly
         If RX.RecordCount > 0 Then
            If Not rsdesadic.EOF Then rsdesadic("MONTO") = RX("DEDUCCION") Else sn_sindicato = 1
         End If
         MSINDICATO = True
         RX.Close
      End If
      
      xciamae = Funciones.Determina_Maestro("01044")
    
      sql$ = "Select cod_maestro2,descrip from maestros_2 where " & _
      "status<>'*' and cod_maestro2='" & Trim(rs!Area) & "'"
      sql$ = sql$ & xciamae
          
      Set rs = cn.Execute(sql$)
      
      'LLENA GRID CENTRO DE COSTOS
      If rs.RecordCount > 0 Then
         rsccosto.AddNew
         rsccosto.MoveFirst
         rsccosto!codigo = Trim(rs!cod_maestro2)
         rsccosto!Descripcion = UCase(rs!descrip)
         rsccosto!Monto = "100.00"
         lbltot.Caption = "100.00"
         rsccosto!Item = VItem
      End If
      For I = I To 4
         If rsccosto.RecordCount < 5 Then rsccosto.AddNew
      Next I
      rsccosto.MoveFirst
      Dgrdccosto.Refresh
            
      If rs.State = 1 Then rs.Close
      Txtcodpla.Enabled = False
      
   End If
ElseIf Trim(Txtcodpla.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodpla.Text
   Txtcodpla.Text = ""
   Limpia_Boleta
   lblnombre.Caption = ""
   Lblctacte.Caption = "0.00"
   Lblcodaux.Caption = ""
   Lblcodafp.Caption = ""
   Lblnumafp.Caption = ""
   Lblbasico.Caption = ""
   LblFingreso.Caption = ""
   Lbltope.Caption = ""
   Lblcargo.Caption = ""
   VAltitud = ""
   VVacacion = ""
   VArea = ""
   VFechaNac = ""
   VFechaJub = ""
   Txtcodpla.SetFocus
   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   Frame5.Enabled = False
End If
VItem = 0
If VTipobol <> "01" Then
   Dgrdhoras.Columns(1).Locked = False
'Else
'   Dgrdhoras.Columns(1).Locked = True
End If

Call Carga_Horas

If VTipobol = "02" Or VTipobol = "03" Then Otros_Pagos_Vac
If wTipoDoc = True And VTipobol = "01" Then
   If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      If rsdesadic!codigo = "09" Then
         sql$ = "select totneto from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
         If (fAbrRst(rs2, sql$)) Then rsdesadic!Monto = rs2(0)
         rs2.Close
      End If
      rsdesadic.MoveNext
   Loop

End If
If wTipoDoc = True Then
   sql = "select sum(importe-pago_acuenta) from plactacte where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and importe-pago_acuenta<>0 and status<>'*' and importe>0"
   If (fAbrRst(rs, sql$)) Then
      If IsNull(rs(0)) Then Lblctacte.Caption = "0.00" Else Lblctacte.Caption = Format(rs(0), "###,###,###.00")
   End If
   rs.Close
End If

If VTipobol = "02" And vDevengue = False Then
   sql = "select i16 from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
   If (fAbrRst(rs, sql$)) Then
      rs.MoveFirst
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         rspagadic.Delete
         rspagadic.MoveNext
      Loop
      If rs!I16 <> 0 Then
         rspagadic.AddNew
         rspagadic!codigo = "16"
         rspagadic!Monto = rs!I16
         rspagadic!Descripcion = "OTROS PAGOS"
      End If
      MsgBox "Trabajador Tiene Vacaciones Devengadas" & Chr(13) & "No podra modificar Datos, Solo Grabar", vbInformation, "Vacaciones Devengadas"
      BolDevengada = True
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      Frame5.Enabled = False
   End If
End If
If VTipobol = "01" Or VTipobol = "03" Then MUESTRA_CUENTACORRIENTE
Exit Sub
CORRIGE:
  MsgBox "Error:" & Err.Description, vbCritical, Me.Caption
End Sub
Public Sub Carga_Boleta(codigo As String, tipo As String, Nuevo As Boolean, semana As String, fproce As String, Tipot As String, perpago As String, horas As Integer, mdel As String, mal As String, obra As String, devengue As Boolean)
Dim MField As String
Dim wciamae As String
Load Frmboleta
VTipobol = tipo
VObra = obra
vTipoTra = Tipot
VSemana = semana
VFProceso = fproce
VPerPago = perpago
VNewBoleta = Nuevo
VHoras = horas
Vano = Val(Mid(VFProceso, 7, 4))
Vmes = Val(Mid(VFProceso, 4, 2))
Vdia = Val(Mid(VFProceso, 1, 2))
VfDel = mdel
VfAl = mal
LstCcosto.Clear
wciamae = Funciones.Determina_Maestro("01044")
sql$ = "Select cod_maestro2,descrip from maestros_2 where  status<>'*'"
sql$ = sql$ & wciamae

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(sql$, 64)
If rs.RecordCount > 0 Then rs.MoveFirst
Do Until rs.EOF
   LstCcosto.AddItem rs!descrip & Space(100) & rs!cod_maestro2
   rs.MoveNext
Loop
If rs.State = 1 Then rs.Close

'Turno
sql$ = "Select codturno,descripcion from platurno where cia='" & wcia & "' and status<>'*' order by codturno"
Cmbturno.Clear
VTurno = ""

If (fAbrRst(rs, sql$)) Then
   If (Not rs.EOF) Then
      Do Until rs.EOF
         Cmbturno.AddItem rs(1)
         Cmbturno.ItemData(Cmbturno.NewIndex) = rs(0)
         rs.MoveNext
       Loop
       Cmbturno.ListIndex = 0
    End If
    If rs.State = 1 Then rs.Close
End If
Select Case VTipobol 'VACACIONES
       Case Is = "02"
            Lblvaca1.Visible = True
            Txtvacai.Visible = True
            Lblvaca2.Caption = "F. Inicio Vaca."
            Txtvacaf.Visible = True
            Label4.Visible = True
            Txtvaca.Visible = True
            Lblvaca2.Visible = True
            Txtcese.Visible = False
       Case Else
            Lblvaca1.Visible = False
            Txtvacai.Visible = False
            Txtvacaf.Visible = False
            Label4.Visible = False
            Txtvaca.Visible = False
            'Txtcese.Visible = True
End Select
Procesa
If VNewBoleta = False Then
   Txtcodpla.Text = codigo
   Txtcodpla_LostFocus
   If rshoras.RecordCount > 0 Then rshoras.MoveFirst
   Do While Not rshoras.EOF
      rshoras!Monto = "0"
      rshoras.MoveNext
   Loop
   If VNewBoleta = False Then MsgBox "Los datos solo se mostraran como consulta " & Chr(13) & "Si desea modificar algun dato debera anular la boleta", vbInformation, "Sistema de Planilla"
   sql$ = ""
   If wTipoDoc = True Then
      Select Case VPerPago
             Case Is = "02"
                  sql$ = "select * from plahistorico " _
                     & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
                     & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
             Case Is = "04"
                  sql$ = "select * from plahistorico " _
                     & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and year( fechaproceso) = " & Vano & " And Month(fechaproceso) = " & Vmes & " " _
                     & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
      End Select
   Else
      sql$ = "select * from plaquincena " _
         & "where cia='" & wcia & "' and year( fechaproceso) = " & Vano & " And Month(fechaproceso) = " & Vmes & " " _
         & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
   End If
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(sql$, 64)
   If rs.RecordCount > 0 Then
      Call rUbiIndCmbBox(Cmbturno, Format(rs!turno, "00"), "00")
     
      If rshoras.RecordCount > 0 Then rshoras.MoveFirst
      Do While Not rshoras.EOF
         MField = "h" & rshoras!codigo
         If rshoras!codigo <> "14" Then
            rshoras!Monto = rs.Fields(MField)
         Else
            rshoras!Monto = rs.Fields(MField) / 8
         End If
         rshoras.MoveNext
      Loop
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         MField = "i" & rspagadic!codigo
         rspagadic!Monto = rs.Fields(MField)
         rspagadic.MoveNext
      Loop
      
      If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
         Do While Not rsdesadic.EOF
            MField = "d" & rsdesadic!codigo
            rsdesadic!Monto = rs.Fields(MField)
            rsdesadic.MoveNext
         Loop
   End If
End If
Frame2.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False
'If VTipobol = "02" Then Dgrdhoras.Enabled = False
vDevengue = devengue
If devengue = True Then Calcula_Devengue_Vaca
'If VTipobol = "01" Then MUESTRA_CUENTACORRIENTE
End Sub
Private Sub Total_porcentaje(mdele As String)
Dim mcolf As Integer
Dim mtotalf As Currency
Dim rST As New Recordset


If rsccosto.RecordCount <= 0 Then lbltot.Caption = "0.00": Exit Sub
If Not rsccosto.EOF And mdele <> "S" Then mi = rsccosto!Item
Set rST = rsccosto.Clone
rST.AbsolutePosition = 1
With rST
    mtotalf = 0
    mtotali = 0
    If rST.RecordCount > 0 Then
       mcolf = Dgrdccosto.Row
       rST.MoveFirst
        
       Do While Not rST.EOF
          If Not IsNull(rST!Monto) Then
             mtotalf = mtotalf + rST!Monto
          End If

          rST.MoveNext
       Loop
    End If
    lbltot.Caption = Format(mtotalf, "#,###,###.00")
End With

If CCur(lbltot.Caption) > 100 Then
   MsgBox "Total Porcentaje no puede exceder a 100%", vbCritical, TitMsg
   rsccosto!Monto = Format(CCur(rsccosto!Monto) - (CCur(lbltot.Caption) - 100), "###,###.00")
   lbltot.Caption = "100.00"
End If
Dgrdccosto.Refresh
Dgrdccosto.MarqueeStyle = 6
If Dgrdccosto.Enabled = True Then Dgrdccosto.SetFocus
End Sub
Public Sub Grabar_Boleta()
Dim MqueryH As String
Dim MqueryP As String
Dim MqueryD As String
Dim MqueryI As String
Dim MqueryCalD As String
Dim MqueryCalA As String
Dim mtoting As Currency
Dim itemcosto As Integer
Dim mcad As String
Dim QUINCENA As Currency

  If vDevengue = True Then
   Mstatus = "D"
   VTurno = Format(NumDev, "00")
Else
   Mstatus = "T"
End If

If Trim(Lblbasico.Caption) = "" Then
   MsgBox "Trabajador No Registra Sueldo Basico", vbInformation, "Boletas de Pago"
   Lblbasico.Caption = ""
   VokDevengue = False
   Exit Sub
End If

mcancel = False
mtoting = 0
Total_porcentaje ("S")

If Trim(Cmbturno.Text) = "" Then MsgBox "Debe Indicar Turno", vbCritical, TitMsg: Cmbturno.SetFocus: VokDevengue = False: Exit Sub
If CCur(lbltot.Caption) <> 100 Then MsgBox "Total Porcentaje de Centro de Costos debe ser 100%", vbCritical, TitMsg: VokDevengue = False: Exit Sub

If VTipobol = "02" And BolDevengada = True And vDevengue = False Then
   Grabar_Devengada
   Exit Sub
End If

If Verifica_Boleta = False Then
   If wTipoDoc = True And VTipobol <> "02" Then
      MsgBox "Ya existe boleta generada con el mismo periodo, Debe eliminarla para poder generar otra", vbExclamation, TitMsg: VokDevengue = False: Exit Sub
   ElseIf VTipobol <> "02" Then
      MsgBox "Ya existe Adelanto de Quincena con el mismo periodo, Debe eliminar para poder generar otra", vbExclamation, TitMsg: VokDevengue = False: Exit Sub
   End If
End If

If vDevengue <> True Then
   Mgrab = MsgBox("Seguro de Grabar Boleta", vbYesNo + vbQuestion, TitMsg)
   If Mgrab <> 6 Then Exit Sub
End If
Screen.MousePointer = vbArrowHourglass

manos = perendat(VFProceso, VFechaNac, "a")

sql$ = ""
sql$ = wInicioTrans
cn.Execute sql$

sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
sql$ = sql$ & "insert into platemphist(cia,placod," & _
       "codauxinterno,proceso,fechaproceso,fechavacai,fechavacaf,fecperiodovaca,semana," & _
       "fechaingreso,turno,codafp,status,fec_crea," & _
       "tipotrab,obra,numafp,basico,fec_modi," & _
       "user_crea,user_modi) values('" & wcia & _
       "','" & Trim(Txtcodpla.Text) & "','" & _
       Lblcodaux.Caption & "','" & VTipobol & "','" & _
       Format(VFProceso, FormatFecha) & "'," & IIf(Txtvacai.Text <> "__/__/____", "'" & Format(Txtvacai.Text, FormatFecha) & "'", "NULL") & "," & IIf(Txtvacaf.Text <> "__/__/____", "'" & Format(Txtvacaf.Text, FormatFecha) & "'", "NULL") & "," & _
       IIf(Txtvaca.Text <> "__/__/____", "'" & Format(Txtvaca.Text, FormatFecha) & "'", "NULL") & ",'" & VSemana & "','" & Format(LblFingreso.Caption, FormatFecha) & _
       "','" & Format(VTurno, "0") & "','" & _
       Lblcodafp.Caption & "','" & Mstatus & "'," & _
       FechaSys & ",'" & vTipoTra & "','" & VObra & _
       "','" & Lblnumafp.Caption & "'," & _
       CCur(Lblbasico.Caption) & "," & FechaSys & _
       ",'" & wuser & "','" & wuser & "')"
 
cn.Execute sql

sql = "SELECT COALESCE(SUM(importe),0) as quincena FROM plabolcte WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' AND sn_quincena=1"
sql = sql & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ""
If (fAbrRst(rs, sql)) Then QUINCENA = rs(0)

rsdesadic.MoveFirst
Do While Not rsdesadic.EOF
    If rsdesadic("CODIGO") = "07" Then
        rsdesadic("MONTO") = rsdesadic("MONTO") + QUINCENA
        Exit Do
    End If
    rsdesadic.MoveNext
Loop

'Horas
If rshoras.RecordCount > 0 Then
   rshoras.MoveFirst
   MqueryH = ""

   Do While Not rshoras.EOF
   If rshoras!codigo <> "14" Then
      MqueryH = MqueryH & "h" & rshoras!codigo & "=" & IIf(IsNull(rshoras!Monto), 0, rshoras!Monto) & ""
    Else
        MqueryH = MqueryH & "h" & rshoras!codigo & "=" & IIf(IsNull(rshoras!Monto), 0, rshoras!Monto) * 8 & ""
    End If
      rshoras.MoveNext
      If Not rshoras.EOF Then MqueryH = MqueryH & ","
   Loop
End If

'Pagos Adicionales
If rspagadic.RecordCount > 0 Then
   rspagadic.MoveFirst
   MqueryP = ""
   Do While Not rspagadic.EOF
      MqueryP = MqueryP & "i" & rspagadic!codigo & "=" & Val(rspagadic!Monto) & ""
      rspagadic.MoveNext
      If Not rspagadic.EOF Then MqueryP = MqueryP & ","
   Loop
End If

'Descuentos Adicionales
If rsdesadic.RecordCount > 0 Then
   rsdesadic.MoveFirst
   MqueryD = ""
   Do While Not rsdesadic.EOF
    If rsdesadic!codigo = "13" And wTipoDoc = False Then
        MqueryD = MqueryD & "d" & rsdesadic!codigo & "=" & Round(Val(rsdesadic!Monto) / 2, 2) & ""
    Else
      MqueryD = MqueryD & "d" & rsdesadic!codigo & "=" & Val(rsdesadic!Monto) & ""
    End If
      rsdesadic.MoveNext
      If Not rsdesadic.EOF Then MqueryD = MqueryD & ","
   Loop
End If

itemcosto = 1
mcad = ""
 
If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
Do While Not rsccosto.EOF
   If Not IsNull(rsccosto!Monto) Then
      If rsccosto!Monto <> 0 Then
         mcad = mcad & "ccosto" & Format(itemcosto, "0") & " = '" & rsccosto!codigo & "'," & "porc" & Format(itemcosto, "0") & " = " & Str(rsccosto!Monto) & ","
      End If
   End If
   rsccosto.MoveNext
   itemcosto = itemcosto + 1
Loop
mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)

sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
sql$ = sql$ & "Update platemphist set " & mcad & "," & IIf(Trim(MqueryH) = "", "", MqueryH & ",") & MqueryP & "," & MqueryD
sql$ = sql$ & " where cia='" & wcia & "' and placod='" & _
     Trim(Txtcodpla.Text) & "' and codauxinterno='" & _
     Trim(Lblcodaux.Caption) & "' and proceso='" & _
     Trim(VTipobol) & "' and fechaproceso='" & _
     Format(VFProceso, FormatFecha) & _
     "' and semana='" & VSemana & "' and status='" & _
     Mstatus & "'"
     
cn.Execute sql
mcad = ""

'Calculo de ingresos
MqueryI = ""
For I = 1 To 50
    Select Case VTipobol
           Case Is = "01" 'Normal
                sql$ = Me.F01(Format(I, "00"), Val(VSemana))
           Case Is = "02" 'Vacaciones
                sql$ = Me.V01(Format(I, "00"))
           Case Is = "03" 'Gratificaciones
                sql$ = Me.G01(Format(I, "00"))
        Case Is = "04"
            sql$ = Me.F01(Format(I, "00"), Val(VSemana))
        Case Is = "05"
            sql$ = Me.F01(Format(I, "00"), Val(VSemana))
    End Select
    
    If Trim(sql$) <> "" Then
       cn.CursorLocation = adUseClient
       Set rs = New ADODB.Recordset
       Set rs = cn.Execute(sql$, 64)
      'RS.Save "C:\ALEXTE.RS"

       If rs.RecordCount > 0 Then
          rs.MoveFirst
          If IsNull(rs(0)) Or rs(0) = 0 Then
          Else
            If wTipoDoc = False And (I = 10 Or I = 11 Or I = 13 Or I = 21 Or I = 22 Or I = 23 Or I = 24) Then
                MqueryI = MqueryI & "i" & Format(I, "00") & " = " & Round(rs(0) / 2, 2) & ","
            Else
             MqueryI = MqueryI & "i" & Format(I, "00") & " = " & rs(0) & ","
            End If
           
          End If
       End If
       If rs.State = 1 Then rs.Close
    End If
Next

If MqueryI <> "" Then
   MqueryI = Mid(MqueryI, 1, Len(Trim(MqueryI)) - 1)
   sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   sql$ = sql$ & " Update platemphist set " & MqueryI
   sql$ = sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute sql
End If

'Calculo de Deducciones
MqueryCalD = ""
For I = 1 To 20
     sql$ = Me.F02(Format(I, "00"))
    
    If mcancel = True Then
       VokDevengue = False
       MsgBox "Se Cancelo la Grabacion", vbCritical, "Calculo de Boleta"
       sql$ = wCancelTrans
       cn.Execute sql$
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If sql$ <> "" Then
       If (fAbrRst(rs, sql$)) Then
          rs.MoveFirst
          If IsNull(rs(0)) Or rs(0) = 0 Then
            MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
          Else
             If I = 11 Then
                For j = 1 To 5
                   MqueryCalD = MqueryCalD & "d" & Format(I, "00") & Format(j, "0") & " = " & rs(j - 1) & ","
                Next j
             ElseIf I = 13 And wTipoDoc = False Then
                MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) / 2 & ","
             Else
                MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
             End If
          End If
       End If
       If rs.State = 1 Then rs.Close
    End If
Next I
If MqueryCalD <> "" Then
   MqueryCalD = Mid(MqueryCalD, 1, Len(Trim(MqueryCalD)) - 1)
   sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   sql$ = sql$ & "Update platemphist set " & MqueryCalD
   sql$ = sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute sql
   
   sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   sql$ = sql$ & "Update platemphist set d11=d111+d112+d113+d114+d115"
   sql$ = sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute sql
   
End If

'Calculo de Aportaciones
MqueryCalA = ""
For I = 1 To 20
    sql$ = F03(Format(I, "00"))
    If sql$ <> "" Then
       cn.CursorLocation = adUseClient
       Set rs = New ADODB.Recordset
       Set rs = cn.Execute(sql$, 64)
       If rs.RecordCount > 0 Then
          rs.MoveFirst
          If IsNull(rs(0)) Or rs(0) = 0 Then
          Else
             MqueryCalA = MqueryCalA & "a" & Format(I, "00") & " = " & rs(0) & ","
          End If
       End If
       If rs.State = 1 Then rs.Close
    End If
Next
If MqueryCalA <> "" Then
   MqueryCalA = Mid(MqueryCalA, 1, Len(Trim(MqueryCalA)) - 1)
   sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   sql$ = sql$ & "Update platemphist set " & MqueryCalA
   sql$ = sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute sql
End If

Dim mi As String, md As String, ma As String

mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20"
ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20"
sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
sql$ = sql$ & "update platemphist set h14=round(h01+h02+h03/8,0),totaling=" & mi & "," _
     & "totalded=" & md & "," _
     & "totalapo=" & ma & "," _
     & "totneto=(" & mi & ")-" & "(" & md & ")"
   sql$ = sql$ & " where cia='" & wcia & "' and " & _
   " placod='" & Trim(Txtcodpla.Text) & "' and " & _
   "codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
   sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute sql

sql$ = "UPDATE platemphist set h14=round((COALESCE(h01,0)+COALESCE(h02,0)+COALESCE(h03,0))/8,0) where cia='" & wcia & "' and "
sql$ = sql$ & " placod='" & Trim(Txtcodpla.Text) & "' and "
sql$ = sql$ & " codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
sql$ = sql$ & " and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute sql

If wTipoDoc = True Then
   sql$ = "insert into plahistorico select * from platemphist"
Else
   sql$ = "insert into plaquincena select * from platemphist"
End If
sql$ = sql$ & " where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute sql

sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
sql$ = sql$ & "select cia,placod,proceso,fechaproceso,semana,tipotrab,d07 from platemphist"
sql$ = sql$ & " where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, Coneccion.FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute sql

If (fAbrRst(rs, sql$)) Then
   Call Descarga_ctaCte(rs!PLACOD, rs!Proceso, Format(rs!FECHAPROCESO, Coneccion.FormatFecha), rs!semana, rs!tipotrab, rs!d07)
End If

sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
sql$ = sql$ & " and cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
sql$ = sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute sql

sql$ = wFinTrans
cn.Execute sql$
Erase ArrDsctoCTACTE
Limpia_Boleta
Screen.MousePointer = vbDefault
End Sub

Public Sub Limpia_Boleta()
Frame2.Enabled = True
Frame3.Enabled = True
Frame4.Enabled = True
Frame5.Enabled = True

VNewBoleta = True
Txtcodpla.Text = ""
lblnombre.Caption = ""
Lblctacte.Caption = "0.00"
Lblcodaux.Caption = ""
Lblcodafp.Caption = ""
Lblnumafp.Caption = ""
Lblbasico.Caption = ""
Lbltope.Caption = ""
Lblcargo.Caption = ""
VAltitud = ""
VVacacion = ""
VArea = ""
VFechaNac = ""
VFechaJub = ""
LblFingreso.Caption = ""
Txtcodpla.Enabled = True
Txtcodpla.SetFocus
Frame2.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False

If rshoras.RecordCount > 0 Then rshoras.MoveFirst
Do While Not rshoras.EOF
   rshoras.Delete
   rshoras.MoveNext
Loop

If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
Do While Not rspagadic.EOF
   rspagadic!Monto = 0
   rspagadic.MoveNext
Loop

If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
Do While Not rsdesadic.EOF
   rsdesadic!Monto = 0
   rsdesadic.MoveNext
Loop

If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
Do While Not rsccosto.EOF
   rsccosto.Delete
   rsccosto.MoveNext
Loop
End Sub
Private Function Verifica_Boleta() As Boolean
On Error GoTo CORRIGE
 If wTipoDoc = True Then
   Select Case VPerPago
          Case Is = "02"
               sql$ = "select * from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "' and semana='" & VSemana & "' and Year(fechaproceso)=" & Vano & " and status<>'*' "
          Case Is = "04"
               sql$ = "select * from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
   End Select
Else
   sql$ = "select * from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
End If
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(sql$, 64)
If rs.RecordCount > 0 Then Verifica_Boleta = False Else Verifica_Boleta = True
If rs.State = 1 Then rs.Close
 Exit Function
CORRIGE:
        MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Function
Public Sub Elimina_Boleta()
  If wTipoDoc = True Then
   Mgrab = MsgBox("Seguro de Eliminar Boleta", vbYesNo + vbQuestion, TitMsg)
Else
   sql$ = "select placod from plahistorico " _
   & "where cia='" & wcia & "' and proceso='01' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
   & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
   If (fAbrRst(rs, sql$)) Then
      MsgBox "Ya se genero la Boleta de Pago " & Chr(13) & "No se Puede Anular el Adelanto de Quincena", vbCritical, "Sistema de Planilla"
      Exit Sub
   End If
   Mgrab = MsgBox("Seguro de Eliminar Adelanto de Quincena", vbYesNo + vbQuestion, TitMsg)
End If
If Mgrab <> 6 Then Exit Sub

sql$ = wInicioTrans
cn.Execute sql$
If wTipoDoc = True Then
   Select Case VPerPago
          Case Is = "02"
               sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
               & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
          Case Is = "04"
               sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
               & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
   End Select
Else
   sql$ = "update plaquincena set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
   & "where cia='" & wcia & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
   & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
End If
cn.Execute sql$

'If wTipoDoc = True Then
   sql$ = "select * from plabolcte " _
   & "where cia='" & wcia & "' and placod='" & UCase(Trim(Txtcodpla.Text)) & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
   & "and month( fechaproceso) = " & Vmes & " and day( fechaproceso) = " & Vdia & " and status<>'*'"
   If (fAbrRst(rs, sql$)) Then rs.MoveFirst
   Do While Not rs.EOF
      sql = "update plactacte set fecha_cancela=null,pago_acuenta=pago_acuenta-" & rs!importe & " where cia='" & wcia & "' and placod='" & UCase(Trim(Txtcodpla.Text)) & "' and tipo='" & rs!tipo & "' and id_doc=" & rs!id_doc
      cn.Execute sql$
      rs.MoveNext
   Loop
   sql$ = "update plabolcte set status='*'" _
   & "where cia='" & wcia & "' and placod='" & UCase(Trim(Txtcodpla.Text)) & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
   & "and month( fechaproceso) = " & Vmes & " and day( fechaproceso) = " & Vdia & " and status<>'*'"
   cn.Execute sql$
'End If

sql$ = wFinTrans
cn.Execute sql$

End Sub

Public Function F01(concepto As String, Optional ByVal pSemana As Integer) As String 'INGRESOS
Dim rsF01 As ADODB.Recordset
Dim mFactor As Currency
Dim nHijos As Integer
Dim RX As New ADODB.Recordset
Dim sSQL As String
Dim pSemanaPago As Integer
mFactor = 0
nHijos = 0
F01 = ""
Select Case concepto
       Case Is = "01" 'BASICO
            F01 = "select round((b.importe/factor_horas)*a.h01,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & Trim(concepto) & "' and b.status<>'*'"
       Case Is = "02" 'ASIGNACION FAMILIAR
            sSQL = "SELECT semana_pago FROM PLACONSTANTE WHERE STATUS!='*' AND CIA='" & wcia & "' and codinterno='02' and tipomovimiento='02'"
            If (fAbrRst(RX, sSQL)) Then pSemanaPago = RX(0)
            If RX.State = 1 Then RX.Close
            
            If vTipoTra = "02" Then
                If Semana_Calcular(pSemana, pSemanaPago, Year(FrmCabezaBol.Cmbfecha), wcia) Then
                    F01 = "select round((b.importe/factor_horas)*(factor_horas),2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                End If
            Else
                F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+a.h08),2) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
                F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            End If
            
       Case Is = "03" 'ASIGNACION MOVILIDAD
            ' CAMBIO DE H14 POR H01  {>MA<} 10/06/2007
            F01 = "select round((b.importe/factor_horas)*a.H01+A.H03,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
        
       Case Is = "04" 'BONIFICACION T. SERVICIO
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h02+a.h03+a.h04+a.h05+a.h12),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           
       Case Is = "05" 'INCREMENTO AFP 10.23%

            '*******************codigo Original**********************************
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"

            '************codigo agregado giovanni 05082007****************
            Select Case FrmCabezaBol.Cmbtipotrabajador.Text
                Case "EMPLEADO": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                Case "OBRERO": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            End Select
            '*************************************************************
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            '*****************************************************************************************
            
       Case Is = "06" 'INCREMENTO AFP 3%
            '*****************codigo modificado giovanni 29082007**************************************
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"

            '************codigo agregado giovanni 05082007****************
            Select Case FrmCabezaBol.Cmbtipotrabajador.Text
                Case "EMPLEADO": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                Case "OBRERO": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            End Select
            '*************************************************************
            
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            '******************************************************************************************
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            
       Case Is = "08" 'SOBRETASA (CONSTRUCCION CIVIL)
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
          
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select ROUND(" & mFactor & " *(a.h13),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               'F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h13),2) as basico from platemphist a,plaremunbase b "
               'F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
               'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "09" 'DOMINICAL
            If vTipoTra <> "01" Then
               F01 = "select round(((b.importe/factor_horas)*h02),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "10" 'EXTRAS L-S
            
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
''               F01 = "select round((b.importe/factor_horas)* ((" & mFactor & " *(a.h10))+ a.h10),2) as basico from platemphist a,plaremunbase b "
''               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
''               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
''               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               '=====================================
               
                 'bts = convierte_cant
               
'                 F01 = "select round((((b.importe+" & bts & "+A.I07+A.I06)/factor_horas)* " & mFactor & ") *(a.h10),2) as basico,B.IMPORTE,FACTOR_HORAS from platemphist a,plaremunbase b "
'                 F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
'                 F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'                 F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.status<>'*'"
                 
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h10,2)) as basico"
                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA) INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                 F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('02')"
                 
            End If
       Case Is = "11" 'EXTRAS D-F
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                'bts = convierte_cant
               
'               F01 = "select round((((b.importe+" & bts & ")/factor_horas)* " & mFactor & ") *(a.h11),2) as basico from platemphist a,plaremunbase b "
'               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
'               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h11,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('02')"

            End If
       Case Is = "12" 'FERIADOS
            F01 = "select round((b.importe/factor_horas)*a.h03,2) as basico,A.H03 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "13" 'REINTEGROS
       Case Is = "14" 'VACACIONES (CONSTRUCCION CIVIL)
            If VVacacion = "S" Then
               sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsF01, sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "15" 'GRATIFICACION
       Case Is = "16" 'OTROS PAGOS
       Case Is = "17" 'ASIGNACION ESCOLAR
            If vTipoTra = "01" Then
            End If
            If vTipoTra = "02" Then
            End If
            If vTipoTra = "05" Then
               sql$ = "select ultima from plasemanas where cia='" & wcia & "' and semana='" & Format(VSemana, "00") & "' and ano='" & Vano & "' and status<>'*'"
               If (fAbrRst(rsF01, sql$)) Then
                  If rsF01!ultima = "S" Then
                     sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
                     If (fAbrRst(rsF01, sql$)) Then
                        nHijos = Numero_Hijos(Trim(Txtcodpla.Text), "S", "S", VFProceso, 18)
                        mFactor = rsF01!factor
                        If rsF01.State = 1 Then rsF01.Close
                        F01 = "select round((((b.importe/factor_horas)*8)* " & mFactor & ")/12 * " & nHijos & ",2) as basico from platemphist a,plaremunbase b "
                        F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                        F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                        F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
                     End If
                  End If
               End If
            End If
       Case Is = "18" 'UTILIDADES
       Case Is = "20" 'BUC
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
            
        Case Is = "21" '3RO HORAS EXTRAS
                sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
                mFactor = rsF01!factor
                If rsF01.State = 1 Then rsF01.Close
                'select round((((b.importe+" & bts & "+A.I07+A.I06)/factor_horas)* " & mFactor & ") *(a.h10),2) as basico
                  'bts = convierte_cant
                  
'                F01 = "select round(((b.importe+" & bts & ")/factor_horas)* (" & mFactor & " *(a.h17)),2) as basico from platemphist a,plaremunbase b "
'                F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
'                F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h17,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(Txtcodpla.Text) & "' AND A.STATUS!='*' and a.concepto not in ('02')"
                
            End If
        
'       Case Is = "21" 'BONIFICACION POR ALTURA
'            SQL$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
'            If (fAbrRst(rsF01, SQL$)) Then
'               mFactor = rsF01!factor
'               If rsF01.State = 1 Then rsF01.Close
'               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h18),2) as basico from platemphist a,plaremunbase b "
'               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
'               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
'            End If
       Case Is = "22" 'BONIF. CONTACTO AGUA
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h19),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "23" 'BONIF.POR ALTITUD
            If VAltitud = "S" Then
               sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsF01, sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round((" & mFactor & " / 8) * (a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "24" 'BONIF. TURNO NOCHE
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h20),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "25" 'H.E. HASTA DECIMA HORA
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h21),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "26" 'H.E. HASTA ONCEAVA HORA
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h22),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "27" 'EXTRAS 3PRA L-S
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h23),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "28" 'EXT. NOCHE 2PR L-S
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h24),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "29" 'EXT. NOCHE 3PRA L-S
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h25),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "30" 'SOBRETASA NOCHE(CONSTRUCCION CIVIL)
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h07),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
End Select

End Function
Public Function F02(concepto As String) As String 'DEDUCCIONES
Dim rsF02 As ADODB.Recordset
Dim rsF02afp As ADODB.Recordset
Dim F02str As String
Dim rsTope As ADODB.Recordset
Dim mFactor As Currency
Dim mperiodoafp As String
Dim vNombField As String
Dim mtope As Currency
Dim mincremento As Currency
Dim difmincremento As Currency
Dim mproy As Currency
Dim muit As Currency
Dim mgra As Integer
Dim msemano As Integer
Dim mpertope As Integer
Dim j As Integer
Dim conceptosremu As String
Dim snmanual As Byte, snfijo As Byte
Dim cptoincrementos As String
Dim IMPORTEESVMES As Currency

mFactor = 0
F02 = ""
mtope = 0

If (concepto <> "04" Or Trim(Lblcodafp.Caption) = "01" Or Trim(Lblcodafp.Caption) = "" Or Trim(Lblcodafp.Caption) = "02") And concepto <> "11" And concepto <> "13" Then  'SIN AFP
   If Not IsDate(VFechaJub) Then
    sql$ = "select deduccion,adicional,status from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and deduccion<>0 and status<>'*'"
  
    If (fAbrRst(rsF02, sql$)) Then
       If Not IsNull(rsF02!deduccion) Then
          If rsF02!deduccion <> 0 Then mFactor = rsF02!deduccion: snmanual = IIf(rsF02!adicional = "S", 0, 1): snfijo = IIf(rsF02!status = "F", 1, 0)
       End If
    End If
    
    If rsF02.State = 1 Then rsF02.Close
    
    If snfijo = 1 And snmanual = 0 And ((sn_essaludvida = 1 And concepto = "06") Or (sn_sindicato = 1 And concepto = "08")) Then
        If wTipoDoc = True Then
            Call Acumula_Mes(concepto, "D")
            If macui = mFactor Then
                F02 = " SELECT " & 0
            Else
                F02 = " SELECT " & mFactor
            End If
        Else
            F02 = " SELECT " & mFactor / 2
        End If
        Exit Function
    End If
    
    If mFactor <> 0 Then
       sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
       If (fAbrRst(rsF02, sql$)) Then
          rsF02.MoveFirst
          F02str = ""
          Do While Not rsF02.EOF
             F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
             rsF02.MoveNext
          Loop
          F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
          If rsF02.State = 1 Then rsF02.Close
          Call Acumula_Mes(concepto, "D")
          F02 = "select round(((" & F02str & " + " & macui & ") * " & mFactor & " /100)-" & macus & ",2) as deduccion from platemphist "
          F02 = F02 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
          F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
       End If
    End If
   End If
ElseIf Trim(concepto) = "11" And Trim(Lblcodafp.Caption) <> "" Then 'AFP
   If Not IsDate(VFechaJub) Then
   If Trim(Lblcodafp.Caption) = "01" Or Trim(Lblcodafp.Caption) = "02" Then GoTo AFP
    sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"

    If (fAbrRst(rsF02, sql$)) Then
       rsF02.MoveFirst
       F02str = ""
       Do While Not rsF02.EOF
          F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
          
          rsF02.MoveNext
       Loop
       
       F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
       
       If rsF02.State = 1 Then rsF02.Close
       mperiodoafp = Format(Vano, "0000") & Format(Vmes, "00")
       sql$ = "select afp01,afp02,afp03,afp04,afp05,tope from  plaafp where periodo='" & mperiodoafp & "' and codafp='" & Lblcodafp.Caption & "' and status<>'*' and cia='" & wcia & "'"
    
       If Not (fAbrRst(rsF02, sql$)) Then
          MsgBox "No se Encuentran Factores de Calculo para AFP", vbCritical, "Calculo de Boleta"
          mcancel = True
          Exit Function
       End If
       sql$ = Acumula_Mes_Afp(concepto, "D")
       If (fAbrRst(rsF02afp, sql$)) Then
          For j = 1 To 5
              vNombField = " as D11" & Format(j, "0")
              mFactor = rsF02(j - 1)
              If j = 2 Then
                 If manos > 64 Then mFactor = 0
                 Call Acumula_Mes_Afp112(concepto, "D")
                 mtope = macui
                 sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
                      & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
                      
                 If (fAbrRst(rsTope, sql$)) Then mtope = mtope + rsTope!TOPE
                 
                 If wTipoDoc = False Then mtope = mtope + rsTope!TOPE
                 
                 If mtope > rsF02!TOPE Then mtope = rsF02!TOPE
                                 
                 If rsTope.State = 1 Then rsTope.Close
                 
                 If wTipoDoc = False Then
                    F02 = F02 & "round((((" & mtope & ") * " & mFactor & " /100)-" & macus & ")/2,2) "
                 Else
                    F02 = F02 & "round(((" & mtope & ") * " & mFactor & " /100)-" & macus & ",2) "
                 End If
                 F02 = F02 & vNombField & ","
              Else
                 If Not IsNull(rsF02afp(0)) Then
                    F02 = F02 & "round(((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(j) & ",2) "
                    F02 = F02 & vNombField & ","
                 Else
                    F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                    F02 = F02 & vNombField & ","
                 End If
              End If
          Next j
          If rsF02afp.State = 1 Then rsF02afp.Close
          If rsF02.State = 1 Then rsF02.Close
       End If
       F02 = Mid(F02, 1, Len(Trim(F02)) - 1)
       F02 = "select " & F02
       F02 = F02 & " from platemphist "
       F02 = F02 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
       F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
AFP:
    End If
   End If
               'de 13 lo cambie a 131 para que no entre
ElseIf concepto = "13" And VTipobol <> "03" Then 'Quinta Categoria

    'PREGUNTAMOS SI ESTA AFECTO O NO
    If VTipobol = "04" Then GoTo quinta
    If Not sn_quinta Then GoTo quinta
    
   sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF02, sql$)) Then
      rsF02.MoveFirst
      F02str = ""
      conceptosremu = ""
      cptoincrementos = ""
      Do While Not rsF02.EOF
         F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
'         If Trim(rsF02!cod_remu) <> "01" And Trim(rsF02!cod_remu) <> "14" And Trim(rsF02!cod_remu) <> "12" Then
'            conceptosremu = conceptosremu & "'" & Trim(rsF02!cod_remu) & "',"
'            cptoincrementos = cptoincrementos & "I" & Trim(rsF02!cod_remu) & "+"
'         End If
         rsF02.MoveNext
      Loop

'      conceptosremu = Mid(conceptosremu, 1, Len(Trim(conceptosremu)) - 1)
'      cptoincrementos = Mid(cptoincrementos, 1, Len(Trim(cptoincrementos)) - 1)
      
      F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      muit = 0
      mtope = 0
      mpertope = 0
      mgra = 0
      msemano = 0
      mincremento = 0
      If vTipoTra <> "01" Then
         sql$ = "select max(semana) as ultima from plasemanas where cia='" & wcia & "' and ano='" & Format(Vano, "0000") & "' and status<>'*'"
         If (fAbrRst(rsF02, sql$)) Then mpertope = rsF02(0): msemano = rsF02(0)
      Else
         If Vmes > 6 Then mpertope = 12 Else mpertope = 13
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      'ACUMULADO DE TODOS INGRESOS
      Call Acumula_Ano(concepto, "D")
      mtope = macui
'      cptoincrementos = F02str
      
      'OBTENER EL INCREMENTO
      If Len(Trim(cptoincrementos)) > 0 Then
        sql$ = "select " & cptoincrementos & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
             & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
        If (fAbrRst(rsF02, sql$)) Then mincremento = rsF02!TOPE
        If rsF02.State = 1 Then rsF02.Close
      End If
      
      sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
           & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      If (fAbrRst(rsF02, sql$)) Then mtope = mtope + rsF02!TOPE
      If rsF02.State = 1 Then rsF02.Close
      
      sql$ = "select concepto,moneda,sum((importe/factor_horas)) as base from plaremunbase a,plaafectos b where a.cia='" & wcia & "' " _
           & "and placod='" & Txtcodpla.Text & "' and a.status<>'*' and b.tipo='D' AND B.CIA='" & wcia & "' and b.codigo='" & concepto & "' and b.tboleta='" & VTipobol & "' and b.status<>'*' and a.concepto=b.cod_remu " _
           & "Group By Placod,a.concepto,a.moneda"
      
      If (fAbrRst(rsF02, sql$)) Then
        Do While Not rsF02.EOF
            mproy = mproy + rsF02!base
            rsF02.MoveNext
        Loop
        'mproy = Round(mproy, 2)
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      sql$ = "select uit from plauit where ano='" & Format(Vano, "0000") & "' and moneda='S/.' and status<>'*'"
      If (fAbrRst(rsF02, sql$)) Then muit = rsF02!uit
      If rsF02.State = 1 Then rsF02.Close
      
      If vTipoTra = "05" Then
         sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='20' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
         If (fAbrRst(rsF02, sql$)) Then mFactor = rsF02!factor
         If rsF02.State = 1 Then rsF02.Close
      
         sql$ = "select concepto,moneda,importe/factor_horas as base from plaremunbase where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and concepto='01' and status<>'*'"
         If (fAbrRst(rsF02, sql$)) Then mproy = mproy + Round(rsF02!base * mFactor, 2)
         If rsF02.State = 1 Then rsF02.Close
      End If
      If vTipoTra = "01" Then
         If wTipoDoc = True Then
            If Vmes < 12 Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Vmes + 1) Else mproy = 0
         Else
            If Vmes = 12 Then
               mproy = (mproy / 2)
            Else
'                SQL = "select sum(importe) from plaremunbase b where b.cia='" & wcia & "' and b.placod='" & Trim(Txtcodpla.Text) & "' and b.concepto in (" & conceptosremu & ") and b.status<>'*'"
'                difmincremento = 0
'                If (fAbrRst(rsF02, SQL$)) Then difmincremento = IIf(IsNull(rsF02(0)), 0, rsF02(0))
'                If rsF02.State = 1 Then rsF02.Close
                
                '{>MA<} 090407
              mproy = (((mproy * VHoras) + difmincremento) * (mpertope - Vmes + 1)) + Round((((mproy * VHoras) + difmincremento) / 2), 2)
              'mproy = ((mproy * VHoras) + mincremento) * (mpertope - Vmes + 1)
            End If
         End If
      Else
         mgra = Busca_Grati()
         If vTipoTra = "05" Then
            sql$ = "select importe/factor_horas as base,b.factor  from plaremunbase a,platasaanexo b where a.cia='" & wcia & "' and a.placod='" & Txtcodpla.Text & "'  and a.concepto='01' " _
                 & "and a.status<>'*' and b.cia='" & wcia & "' and b.tipomovimiento='01' and b.codinterno='15' and b.status<>'*' and b.tipotrab='" & vTipoTra & "' and b.cargo='" & Trim(Lblcargo.Caption) & "'"
            If (fAbrRst(rsF02, sql$)) Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(VSemana)) + ((rsF02!base * 8) * (rsF02!factor * mgra)) + ((rsF02!base * 8) * (mpertope - Val(VSemana)))
            If rsF02.State = 1 Then rsF02.Close
         Else
            sql$ = "select importe/factor_horas as base from plaremunbase where Cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and concepto='01' and status<>'*'"
            If (fAbrRst(rsF02, sql$)) Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(VSemana)) + (((mproy * 240) + mincremento) * mgra) + (rsF02!base * 8) * (mpertope - Val(VSemana))
            If rsF02.State = 1 Then rsF02.Close
            'mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((mproy * 240) * mgra) + (mproy * 8) * (mpertope - Val(VSemana))
         End If
      End If
      mtope = mtope + mproy
      If mtope > Round(muit * 7, 2) Then
         mtope = mtope - Round(muit * 7, 2)
         Select Case mtope
                Case Is < (Round(muit * 27, 2) + 1)
                     mFactor = Round(mtope * 0.15, 2)
                Case Is < (Round(muit * 54, 2) + 1)
                     mFactor = Round(((mtope - (muit * 27)) * 0.21) + (muit * 27) * 0.15, 2)
                Case Else
                     mFactor = Round(((mtope - (muit * 54)) * 0.27) + ((muit * 54) - (muit * 27)) * 0.21, 2) + ((muit * 27) * 0.15)
         End Select
         If vTipoTra = "01" Then
            If wTipoDoc = True Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
            Else
               F02 = "Select Round(((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1)), 2)"
            End If
         Else
            If VTipobol = "02" Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
            Else
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (" & msemano & " - " & Val(VSemana) & " + 1), 2)"
            End If
         End If
      End If
   End If
quinta:
   
ElseIf concepto = "15" Then
       MsgBox "666"
End If
macui = 0: macus = 0
End Function
Private Sub Carga_Horas()
Dim rs2 As ADODB.Recordset
Dim mbol As String
Dim mconceptos As String
Dim mhor As Currency
Dim mdiasfalta As String
Dim mdiasferiado As String
Dim wBeginMonth As String
Dim mHourTra As Currency
Dim con As String
Dim I As Integer

On Error GoTo CORRIGE

If Trim(Txtcodpla.Text) = "" Then Exit Sub
If VTipobol = "04" Or VTipobol = "05" Then Exit Sub

sql$ = "select iniciomes from cia where cod_cia='" & _
wcia & "' and status<>'*'"

If (fAbrRst(rs, sql$)) Then
   If IsNull(rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = rs!iniciomes
End If
rs.Close

If Trim(wBeginMonth) <> "1" Then
   If Vmes = 1 Then
      wBeginMonth = Format(wBeginMonth, "00") & "/12/" & Format(Vano - 1, "0000")
   Else
     wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes - 1, "00") & "/" & Format(Vano, "0000")
   End If
Else
   wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes, "00") & "/" & Format(Vano, "0000")
End If

If VPerPago = "02" Then wBeginMonth = VfDel

If rshoras.RecordCount > 0 Then rshoras.MoveFirst
Do While Not rshoras.EOF
   rshoras.Delete
   rshoras.MoveNext
Loop
 
sql$ = wInicioTrans
cn.Execute sql$
sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
cn.Execute sql
sql$ = wFinTrans
cn.Execute sql$

mdiasfalta = 0
mdiasferiado = 0
VHorasnormal = VHoras
If VTipobol = "01" Then mbol = "N"
If VTipobol = "02" Then mbol = "V"
If VTipobol = "03" Then mbol = "G"

wciamae = Funciones.Determina_Maestro("01077")

If VTipobol = "01" Or wTipoDoc <> True Then

   sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    
   If wTipoDoc = True Then
     
      Select Case VPerPago
             Case Is = "04" 'Mensual
                  sql$ = sql$ & "select distinct(concepto) from platareo where fecha " _
                       & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + Coneccion.FormatTimei & _
                       "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
                       
             Case Is = "02" 'Semanal
                  sql$ = sql$ & "select distinct(concepto) from platareo where fecha " _
                       & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      End Select
   Else
      sql$ = sql$ & "select distinct(concepto) from platareo where fecha " _
           & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & _
           "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
           & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      mbol = "N"
   End If
   
   If (fAbrRst(rs, sql$)) Then
      rs.MoveFirst
      mconceptos = ""
      Do While Not rs.EOF
         mconceptos = mconceptos & "'" & Trim(rs!concepto) & "',"
         rs.MoveNext
      Loop
      mconceptos = "(cod_maestro2 in (" & Mid(mconceptos, 1, Len(mconceptos) - 1) & ")"
   End If
   
   If rs.State = 1 Then rs.Close
   
   If Trim(mconceptos) <> "" Then
      sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 and " & mconceptos
   Else
      sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 "
   End If
   
   If vTipoTra = "01" Then
      If Trim(mconceptos) <> "" Then
         sql$ = sql$ & " or cod_maestro2 in('01'))"
      Else
         'SQL$ = SQL$ & " and cod_maestro2 in('01') "
         sql$ = sql$ & " and cod_maestro2 in  ('01','04','05','09','10','17','11','18','19')"
      End If
   Else
     If Trim(mconceptos) <> "" Then
         sql$ = sql$ & " or cod_maestro2 in ('14','01','02','03','04','06','05','07','08','09','10','17','11','12','13','15','18','19'))"
     Else
        sql$ = sql$ & " and cod_maestro2 in  ('14','01','02','03','04','06','05','07','08','09','10','17','11','12','13','15','18','19')"
     End If
   End If
   
'3     '-feriado
'4     'perm. pag.
'6     '-enferm. no pagadas
'5     '-enferm. pagadas
'7     'accidente de trabajo
'8     'faltas injustificadas
'9     'suspencion
'10    'extras l-s
'11    'extras d-f
'12    'vacaciones
'13    'sobretasa
'15    'otros
   
   
   sql$ = sql$ & wciamae
   
   If (fAbrRst(rs, sql$)) Then
      rs.MoveFirst
      con = "140102030406050708091017111213151819"
      
       
      For I = 1 To (Len(con) / 2)
      'Do While Not RS.EOF
         rs.Filter = "COD_MAESTRO2='" & Mid(con, I + (I - 1), 2) & "'"
        If Not rs.EOF Then  ' {<MA>} 01/02/2007
            rshoras.AddNew
            rshoras!codigo = Trim(rs!cod_maestro2)
            rshoras!Descripcion = UCase(rs!descrip)
            If Trim(rs("cod_maestro2")) = "14" Then
               rshoras("MONTO") = 6
            End If

            If vTipoTra = "02" Then
                If rshoras!codigo = "01" Then
                    rshoras!Monto = 6 * 8
                End If
            
                If rshoras!codigo = "02" Then
                    rshoras!Monto = (6 / 6) * 8
                End If
            End If
            
         mhor = 0
         
         sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                  
         Select Case VPerPago
               Case Is = "04"
                    sql$ = sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where fecha " _
                         & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & _
                         Trim(rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
               Case Is = "02"
                    sql$ = sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where fecha " _
                         & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & _
                         Format(VfAl, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & Trim(rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
         End Select
 
         If (fAbrRst(rs2, sql$)) Then
            rs2.MoveFirst
            Do While Not rs2.EOF
               Select Case rs2!motivo
                      Case Is = "DI"
                           'If S Then
                              mhor = mhor + (rs2!tiempo * 8 * 60)
                              mdiasfalta = mdiasfalta + rs2!tiempo
                              If rs!cod_maestro2 = "03" Then mdiasferiado = mdiasferiado + rs2!tiempo
                           'Else
                            '  MsgBox "HOLA"
                           'End If
                      Case Is = "HO"
                           mhor = mhor + Int(rs2!tiempo) * 60 + ((rs2!tiempo - Int(rs2!tiempo)) * 100)
                      Case Is = "MI"
                           mhor = mhor + rs2!tiempo
               End Select
               rs2.MoveNext
            Loop
            If rs2.State = 1 Then rs2.Close
            mhor = Int(mhor / 60) + ((mhor Mod 60) / 100)
                        
            rshoras!Monto = IIf(mhor = 0, Null, mhor)
            
            If Trim(rs!flag2) = "-" Then
               VHorasnormal = VHorasnormal - mhor
            End If
         End If
         
         If rshoras("CODIGO") = "14" And IsNull(rshoras("MONTO")) Then
            rshoras("MONTO") = 6
         End If
         
         rs.MoveNext
      'Loop
      End If
      Next
      
      'TIPICO
       
      rshoras.MoveFirst
      mHourTra = 0
      If rshoras!codigo = "01" Then
         If wTipoDoc = True Then
            mHourTra = Calc_Horas_FecIng(wBeginMonth)
            rshoras!Monto = VHorasnormal
         Else
            VHorasnormal = 0
            If VNewBoleta = True Then VHorasnormal = Calc_Horas_Quincena(wBeginMonth)
            rshoras!Monto = VHorasnormal
         End If
         If rshoras.RecordCount > 1 Then rshoras.MoveNext
      End If
      If vTipoTra <> "01" Then
        If rshoras!codigo = "02" Then rshoras!Monto = ((6 - mdiasfalta + mdiasferiado) * 8) / 6
      End If
      
      If mHourTra <> 0 Then
         rshoras.AddNew
         rshoras!codigo = "03"
         rshoras!Descripcion = "FERIADOS"
         rshoras!Monto = mHourTra
      End If
      
   End If
Else
    sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0"
    sql$ = sql$ & wciamae
    cn.CursorLocation = adUseClient
    Set rs = New ADODB.Recordset
    Set rs = cn.Execute(sql$, 64)
    If rshoras.RecordCount > 0 Then
       rshoras.MoveFirst
       Do While Not rshoras.EOF
          rshoras.Delete
          rshoras.MoveNext
       Loop
    End If
    If Not rs.RecordCount > 0 Then
        If (VTipobol <> "04" And VTipobol <> "05") Then
        MsgBox "No Existen Horas Registradas", vbCritical, TitMsg: Exit Sub
        End If
    End If
    rs.MoveFirst
    Do While Not rs.EOF
       rshoras.AddNew
       rshoras!codigo = Trim(rs!cod_maestro2)
       rshoras!Descripcion = Trim(rs!descrip)
       If VTipobol = "04" Then rshoras!Monto = VHoras Else rshoras!Monto = "0.00"
       rs.MoveNext
    Loop
    If rs.State = 1 Then rs.Close
End If

'solo para datos guardados
Call llena_horas

Exit Sub
CORRIGE:
MsgBox "Error :" & Err.Description, vbCritical, "Sistema de Planillas"
End Sub
Public Function F03(concepto As String) As String 'APORTACIONES
Dim rsF03 As ADODB.Recordset
Dim rscalculo As ADODB.Recordset
Dim F03str As String
Dim mFactor As Currency
mFactor = 0
F03 = ""
If concepto = "03" Then
   sql$ = "select senati from cia where cod_cia='" & wcia & "' and status<>'*'"
   If Not (fAbrRst(rsF03, sql$)) Then Exit Function
   If Trim(rsF03!senati) <> "S" Then Exit Function
   If rsF03.State = 1 Then rsF03.Close
   wciamae = Determina_Maestro("01044")
   sql$ = "Select * from maestros_2 where cod_maestro2='" & Trim(VArea) & "' and status<>'*'"
   sql$ = sql$ & wciamae
   If (fAbrRst(rsF03, sql$)) Then
      If rsF03!flag7 <> "S" Then Exit Function
   Else
      Exit Function
   End If
   If rsF03.State = 1 Then rsF03.Close
End If

sql$ = "select aportacion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and aportacion<>0 and status<>'*'"
If (fAbrRst(rsF03, sql$)) Then
   If Not IsNull(rsF03!aportacion) Then
      If rsF03!aportacion <> 0 Then mFactor = rsF03!aportacion
   End If
End If
If rsF03.State = 1 Then rsF03.Close
If mFactor <> 0 Then
   sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF03, sql$)) Then
      rsF03.MoveFirst
      F03str = ""
      Do While Not rsF03.EOF
         F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
         rsF03.MoveNext
      Loop
      
      F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
      If rsF03.State = 1 Then rsF03.Close
      
      Call Acumula_Mes(concepto, "A")
      
      F03 = "select (" & F03str & " + " & macui & ") , (" & mFactor & " /100)," & macus & " from platemphist "
      F03 = F03 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
      F03 = F03 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      
      If (fAbrRst(rscalculo, F03)) Then
        If rscalculo(0) > sueldominimo Then
            F03 = " SELECT " & Round((rscalculo(0) * (mFactor / 100)) - macus, 2)
        Else
            F03 = "SELECT " & Round((sueldominimo * (mFactor / 100)) - macus, 2)
        End If
      End If
      
   End If
End If

macui = 0: macus = 0
End Function
Public Function V01(concepto As String) As String
'INGRESOS
Dim rsV01 As ADODB.Recordset
Dim mFactor As Currency
Dim fACTORfAMILIA As String

If wcia = "05" Then fACTORfAMILIA = 2 Else fACTORfAMILIA = 1

mFactor = 0
V01 = ""
Select Case concepto
       Case Is = "02" 'ASIGNACION FAMILIAR
            V01 = "select round((b.importe/factor_horas)*(a.h12*" & fACTORfAMILIA & "),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "04" 'BONIFICACION T. SERVICIO
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "05" 'INCREMENTO AFP 10.23%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "06" 'INCREMENTO AFP 3%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "14" 'VACACIONES
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "20" 'BUC
            sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsV01, sql$)) Then
               mFactor = rsV01!factor
               If rsV01.State = 1 Then rsV01.Close
               V01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h12),2) as basico from platemphist a,plaremunbase b "
               V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
End Select
'Debug.Print concepto
'Debug.Print SQL$
End Function
Public Function G01(concepto As String) As String 'INGRESOS
Dim rsG01 As ADODB.Recordset
Dim mFactor As Currency
Dim mh As Integer
mFactor = 0
If Val(Mid(VFProceso, 4, 2)) = 7 Then mh = 7 Else mh = 5
G01 = ""
If vTipoTra <> "05" Then
    Select Case concepto
           Case Is = "02" 'ASIGNACION FAMILIAR
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "04" 'BONIFICACION T. SERVICIO
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "05" 'INCREMENTO AFP 10.23%
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "06" 'INCREMENTO AFP 3%
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "07" 'BONIFICACION COSTO DE VIDA
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "15" 'GRATIFICACION
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
    End Select
Else
   Select Case concepto
          Case Is = "15" 'GRATIFICACION
               sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsG01, sql$)) Then
                  mFactor = rsG01!factor
                  If rsG01.State = 1 Then rsG01.Close
                  G01 = "select round(((((b.importe/factor_horas)*8)* " & mFactor & ") / " & mh & ") * a.h14 ,2) as basico from platemphist a,plaremunbase b "
                  G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                  G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
   End Select
End If
End Function
Private Sub Acumula_Mes(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim rsacumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 3
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If concepto = "06" And tipo = "D" Then
        sql$ = "select 'D06' AS COD_REMU"
    Else
        sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    End If
    If (fAbrRst(rsacumula, sql$)) Then
       rsacumula.MoveFirst
       mcad = ""
       Do While Not rsacumula.EOF
        If concepto = "06" And tipo = "D" Then
            mcad = mcad & Trim(rsacumula!cod_remu) & "+"
        Else
          mcad = mcad & "i" & Trim(rsacumula!cod_remu) & "+"
        End If
          rsacumula.MoveNext
       Loop
       
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If rsacumula.State = 1 Then rsacumula.Close
End Sub
Private Function Acumula_Mes_Afp(concepto As String, tipo As String) As String
Dim mcad As String
Dim SqlAcu As String
Dim rsacumula As New Recordset

macui = 0: macus = 0
    sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(rsacumula, sql$)) Then
       rsacumula.MoveFirst
       mcad = ""
       Do While Not rsacumula.EOF
          mcad = mcad & "i" & Trim(rsacumula!cod_remu) & "+"
          rsacumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d111) as ded1, " _
              & "sum(d112) as ded2, " _
              & "sum(d113) as ded3, " _
              & "sum(d114) as ded4, " _
              & "sum(d115) as ded5 "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and "
       SqlAcu = SqlAcu & "proceso in('01','02','03','04','05')"
    End If
If rsacumula.State = 1 Then rsacumula.Close
Acumula_Mes_Afp = SqlAcu
End Function
Private Sub Acumula_Mes_Afp112(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim rsacumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 3
    If VTipobol = "02" And I <> 2 Then I = I + 1
    If VTipobol <> "02" And I = 2 Then I = I + 1
    If I > 3 Then Exit For
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(rsacumula, sql$)) Then
       rsacumula.MoveFirst
       mcad = ""
       Do While Not rsacumula.EOF
          mcad = mcad & "i" & Trim(rsacumula!cod_remu) & "+"
          rsacumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d112) as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If rsacumula.State = 1 Then rsacumula.Close
End Sub
Private Sub Acumula_Ano(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim rsacumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 4
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If I = 4 Then mtb = "04"
    sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(rsacumula, sql$)) Then
       rsacumula.MoveFirst
       mcad = ""
       Do While Not rsacumula.EOF
          mcad = mcad & "i" & Trim(rsacumula!cod_remu) & "+"
          rsacumula.MoveNext
       Loop

       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & Txtcodpla.Text & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)<=" & Vmes & " and proceso='" & mtb & "'"

       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If rsacumula.State = 1 Then rsacumula.Close
End Sub
Private Function Busca_Grati() As Integer
Dim rsGrati As New Recordset
Select Case Vmes
       Case Is = 1, 2, 3, 4, 5, 6
            Busca_Grati = 2
       Case Is = 7
            sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, sql$)) Then Busca_Grati = 1 Else Busca_Grati = 2
                 
       Case Is = 8, 9, 10, 11
            Busca_Grati = 1
       Case Is = 12
            sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, sql$)) Then Busca_Grati = 0 Else Busca_Grati = 1
End Select
End Function
Private Function Calc_Horas_FecIng(Inicio As String) As Currency
Dim mFIng As String
Dim mWorkNew As Boolean
Dim VHNew As String
Dim mDateBegin As String

Calc_Horas_FecIng = 0
mFIng = Mid(LblFingreso, 4, 2) & "/" & Left(LblFingreso, 2) & "/" & Right(LblFingreso, 4)
mWorkNew = False

If VNewBoleta = True Then
   If Compara_Fechas(mFIng, Inicio) = True Then
      mWorkNew = True
      VHNew = InputBox("Primera Boleta del Trabajador" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Trabajador Nuevo")
      Do While Not IsNumeric(VHNew)
         VHNew = InputBox("Primera Boleta del Trabajador" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Trabajador Nuevo")
      Loop
      Do While VHNew > VHoras Or Val(VHNew) = 0
         If VHNew >= VHoras Then MsgBox "Las Horas no deben ser mayores a " & Trim(Str(VHoras)), vbInformation, "Horas Trabajadas"
         VHNew = "0"
         VHNew = InputBox("Primera Boleta del Trabajador" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Trabajador Nuevo", VHNew)
         If Not IsNumeric(VHNew) Then VHNew = "0"
      Loop
      VHorasnormal = VHNew
   End If
End If
If mWorkNew = True Then mDateBegin = mFIng Else mDateBegin = Inicio


'FERIADOS
sql$ = ""
If VTipobol = "01" Then
     
   Select Case VPerPago
          Case Is = "04" 'Mensual
               sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
               sql$ = sql$ & "select count(fecha) from plaferiados where cia='" & wcia & "' and fecha " _
                     & "BETWEEN '" & Format(mDateBegin, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                    & "and status<>'*'"
          Case Is = "02" 'Semanal
               sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
               sql$ = sql$ & "select count(fecha) from plaferiados where cia='" & wcia & "' and fecha " _
                     & "BETWEEN '" & Format(mDateBegin, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                    & "and status<>'*'"
   End Select
   If Trim(sql$) <> "" Then If (fAbrRst(rs, sql$)) Then rs.MoveFirst
   If Not IsNull(rs(0)) Then
      If rs(0) <> 0 Then
         VHorasnormal = VHorasnormal - (rs(0) * mHourDay)
         Calc_Horas_FecIng = (rs(0) * mHourDay)
      End If
   Else
   
   End If
End If

End Function
Private Sub Otros_Pagos_Vac()
Dim mCadOtPag As String
Dim mPer As Integer
Dim mDateBeginVac As String
Dim mDateEndVac As String
Dim mdia As Integer
Dim meses As Integer
Dim RX As ADODB.Recordset
Dim FecIni As String, FecFin As String

If Month(VFProceso) = 7 Then
    FecIni = "01/01/" & Year(VFProceso)
    FecFin = "07/01/" & Year(VFProceso)
Else
    FecIni = "07/01/" & Year(VFProceso)
    FecFin = "12/01/" & Year(VFProceso)
End If

mPer = 0
If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
Do While Not rspagadic.EOF
   If rspagadic!codigo = "16" Then
      sql = "select distinct(codinterno),factor,factor_divisionario from platasaanexo where cia='" & wcia & "' and tipomovimiento='" & vTipoTra & "' and modulo='01'  and basecalculo='16' and status<>'*'"
      If (fAbrRst(rs, sql)) Then rs.MoveFirst: mPer = Val(rs(2)): meses = Val(rs(1))
      mCadOtPag = ""
      rs.MoveFirst
      Do While Not rs.EOF
        If VTipobol = "03" Then
              sql = "SELECT COUNT(" & "I" & rs(0) & ") FROM plahistorico WHERE cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status!='*' and (fechaproceso>'" & FecIni & "' and fechaproceso<'" & FecFin & "') and " & "I" & rs(0) & ">0 AND proceso='01'"
              If (fAbrRst(RX, sql)) Then
                  If RX(0) >= 3 Then
                       mCadOtPag = mCadOtPag & "I" & rs(0) & "+"
                  End If
              End If
        Else
           mCadOtPag = mCadOtPag & "I" & rs(0) & "+"
        End If
         rs.MoveNext
      Loop
      rs.Close
      Exit Do
   End If
   rspagadic.MoveNext
Loop
'AGREGAMOS LAS HORAS EXTRAS QUE NO FUESEN CARGADAS
If InStr(1, mCadOtPag, "10") > 0 Or InStr(1, mCadOtPag, "11") > 0 Or InStr(1, mCadOtPag, "21") > 0 Then
    If InStr(1, mCadOtPag, "10") = 0 Then mCadOtPag = mCadOtPag & "I10+"
    If InStr(1, mCadOtPag, "11") = 0 Then mCadOtPag = mCadOtPag & "I11+"
    If InStr(1, mCadOtPag, "21") = 0 Then mCadOtPag = mCadOtPag & "I21+"
End If

sql = ""
If Trim(mCadOtPag) <> "" Then
   mCadOtPag = Mid(mCadOtPag, 1, Len(Trim(mCadOtPag)) - 1)
   mCadOtPag = "sum(" & mCadOtPag & ")"
   sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " select " & mCadOtPag & " from plahistorico"
End If

mDateBeginVac = Fecha_Promedios(mPer, VFProceso)

If Val(Mid(VFProceso, 4, 2)) = 1 Then
   mdia = Ultimo_Dia(12, Val(Mid(VFProceso, 7, 4)) - 1)
   mDateEndVac = Format(mdia, "00") & "/12/" & Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
Else
   mdia = Ultimo_Dia(Val(Mid(VFProceso, 4, 2) - 1), Val(Mid(VFProceso, 7, 4)))
   mDateEndVac = Format(mdia, "00") & "/" & Format(Val(Mid(VFProceso, 4, 2) - 1), "00") & "/" & Mid(VFProceso, 7, 4)
End If

mDateBeginVac = "01/" & Month(DateAdd("m", (meses - 1) * -1, mDateEndVac)) & "/" & Year(DateAdd("m", -5, mDateEndVac))


If Trim(sql) <> "" Then
   sql = sql & " where cia='" & wcia & "' and fechaproceso Between '" & Format(mDateBeginVac, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(mDateEndVac, FormatFecha) & Space(1) & FormatTimef & "'"
   sql = sql & " and proceso='01' and placod='" & Txtcodpla.Text & "' and status<>'*'"
   If (fAbrRst(rs, sql)) Then rs.MoveFirst
   If Not IsNull(rs(0)) Then rspagadic!Monto = rs(0) / mPer
End If
End Sub
Private Function Calc_Horas_Quincena(Inicio As String) As Currency
Dim mFIng As String
Dim mWorkNew As Boolean
Dim VHNew As String

Calc_Horas_Quincena = 0
mFIng = Mid(LblFingreso, 4, 2) & "/" & Left(LblFingreso, 2) & "/" & Right(LblFingreso, 4)
mWorkNew = False
   
VHNew = InputBox("Adelanto de Quincena" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Adelanto de Quincena")
Do While Not IsNumeric(VHNew)
   VHNew = InputBox("Adelanto de Quincena" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Adelanto de Quincena")
Loop
Do While VHNew > VHoras Or Val(VHNew) = 0
   If VHNew >= VHoras Then MsgBox "Las Horas no deben ser mayores a " & Trim(Str(VHoras)), vbInformation, "Horas Trabajadas"
   VHNew = "0"
   VHNew = InputBox("Adelanto de Quincena" & Chr(13) & "Se Requiere que ingrese Horas Trabajadas", "Adelanto de Quincena", VHNew)
   If Not IsNumeric(VHNew) Then VHNew = "0"
Loop
Calc_Horas_Quincena = VHNew
End Function
Private Sub Descarga_ctaCte(codigo As String, tipobol As String, fecha As String, sem As String, tiptrab As String, importe As Currency)
Dim sql As String
Dim RX As ADODB.Recordset
'SQL = "select * from plactacte where cia='" & wcia & "' and placod='" & Trim(codigo) & "' and importe-pago_acuenta>0 and status<>'*' order by fecha"
'If (fAbrRst(rs, SQL$)) Then rs.MoveFirst
'Do While Not rs.EOF
'   saldo = (rs!importe - rs!pago_acuenta)
'   If saldo >= importe Then
'      If saldo = importe Then
'         'SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and numinterno='" & RS!numinterno & "' and tipo='" & RS!tipo & "' and status<>'*'"
'         SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'         cn.Execute SQL
'      Else
'         'SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & " where cia='" & wcia & "' and numinterno='" & RS!numinterno & "' and tipo='" & RS!tipo & "' and status<>'*'"
'         SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & " where cia='" & wcia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'         cn.Execute SQL
'      End If
      QUINCENA = 0
      For I = 0 To MAXROW - 1
        If ArrDsctoCTACTE(0, I) = Trim(Txtcodpla.Text) Then
            If ArrDsctoCTACTE(2, I) > 0 Then
            
                sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                sql = sql & "insert into plabolcte values('" & wcia & "','" & UCase(Trim(codigo)) & "','" & tipobol & "','', " _
                & "'" & fecha & "','" & sem & "','" & tiptrab & "','" & Lblcodaux.Caption & "','" & ArrDsctoCTACTE(3, I) & "','" & fecha & "', " _
                & "'" & wmoncont & "'," & ArrDsctoCTACTE(2, I) & ",'','" & wuser & "'," & FechaSys & "," & ArrDsctoCTACTE(1, I) & "," & IIf(wTipoDoc = True, 0, 1) & ")"
                
                cn.Execute sql
            End If
        
            sql = "UPDATE plactacte set pago_acuenta=pago_acuenta + " & CStr(ArrDsctoCTACTE(2, I)) + " WHERE cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and id_doc='" & CStr(ArrDsctoCTACTE(1, I)) & "' and status<>'*'"
            cn.Execute sql
            
            sql = "UPDATE plactacte set fecha_cancela='" & fecha & "' WHERE cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and id_doc='" & ArrDsctoCTACTE(1, I) & "' and status<>'*' AND IMPORTE=PAGO_ACUENTA"
            cn.Execute sql
                
        End If
      Next
      'Exit Do
'   Else
'      SQL = "update plactacte set pago_acuenta=pago_acuenta+" & saldo & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'      cn.Execute SQL
'
'      SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
'      SQL = SQL & "insert into plabolcte values('" & wcia & "','" & codigo & "','" & tipobol & "','', " _
'          & "'" & Format(fecha, FormatFecha) & "','" & sem & "','" & tiptrab & "','" & rs!codauxinterno & "','" & rs!tipo & "','" & Format(rs!fecha, FormatFecha) & "', " _
'          & "'" & wmoncont & "'," & saldo & ",'','" & wuser & "'," & FechaSys & ")"
'      cn.Execute SQL
'      importe = importe - saldo
'   End If
'   rs.MoveNext
'Loop
'rs.Close
End Sub

Private Sub Calcula_Devengue_Vaca()
Dim mNumBol As Integer
Dim I As Integer
Dim rsdevengue As ADODB.Recordset

VokDevengue = True
sql = "select placod,recordacu from plaprovvaca where cia='" & wcia & "' and year(fechaproceso)=" & Vano - 1 & " and month(fechaproceso)=12 and status<>'*'"
If (fAbrRst(rsdevengue, sql)) Then rsdevengue.MoveFirst
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Value = 0
Barra.Max = rsdevengue.RecordCount
Do While Not rsdevengue.EOF
   Barra.Value = rsdevengue.AbsolutePosition
   mNumBol = Val(Mid(rsdevengue!recordacu, 1, 4))
   If mNumBol > 0 Then
      NumDev = 0
      For I = 1 To mNumBol
          NumDev = NumDev + 1
          Txtcodpla.Text = rsdevengue!PLACOD
          Txtcodpla_LostFocus
          Grabar_Boleta
      Next
   End If
   rsdevengue.MoveNext
Loop
rsdevengue.Close
Panelprogress.Visible = False
If VokDevengue = False Then
   sql = wInicioTrans
   cn.Execute sql
   
   sql = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
        & "where cia='" & wcia & "' and proceso='02' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
        & "and status='D'"
   cn.Execute sql
   
   sql = wFinTrans
   cn.Execute sql
   MsgBox "Se detectaron Irregularidades " & Chr(13) & "Se cancelara el calculo", vbCritical, "Devengue de Vacaciones"
End If
Unload Me
Frmprovision.Provisiones ("D")
Frmprovision.Txtano.Text = Vano
Frmprovision.Cmbmes.ListIndex = Vmes - 1
Frmprovision.Show
Frmprovision.ZOrder 0
End Sub
Private Sub Grabar_Devengada()
Mgrab = MsgBox("Seguro de Grabar Boleta", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub
Screen.MousePointer = vbArrowHourglass

sql$ = wInicioTrans
cn.Execute sql$

sql = "select * from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
If (fAbrRst(rs, sql)) Then rs.MoveFirst
sql = "update plahistorico set status='T' where cia='" & rs!cia & "' and placod='" & rs!PLACOD & "' and status='" & rs!status & "' and turno='" & rs!turno & "'"
cn.Execute sql
rs.Close
sql$ = wFinTrans
cn.Execute sql$
Limpia_Boleta
Screen.MousePointer = vbDefault

End Sub

Sub llena_horas()
    Dim RZ As New ADODB.Recordset
    Dim Rt As New ADODB.Recordset
    Dim con As String, cod As String
    Dim horas As Variant, VALOR As Variant
    Dim I As Integer, ccosto As String, PORC As String
    Dim INGRESOS As String
        
    For I = 1 To 30
         horas = horas & "H" & Format(I, "00") & ","
    Next

    horas = Left(horas, Len(horas) - 1)
    
    'PAGOS ADICIONALES
    If rspagadic.RecordCount > 0 Then
       rspagadic.MoveFirst
       Do While Not rspagadic.EOF
                INGRESOS = INGRESOS & "I" & rspagadic(0) & ","
                rspagadic.MoveNext
       Loop

       INGRESOS = Left(INGRESOS, Len(INGRESOS) - 1)
    End If
    
    'DESCUENTOS
    If rsdesadic.RecordCount > 0 Then
       rsdesadic.MoveFirst
       Do While Not rsdesadic.EOF
                descuentos = descuentos & "d" & rsdesadic(0) & ","
                rsdesadic.MoveNext
       Loop

       descuentos = Left(descuentos, Len(descuentos) - 1)
    End If
        
    ccosto = "ccosto1,ccosto2,ccosto3,ccosto4,ccosto5"
    PORC = "PORC1,PORC2,PORC3,PORC4,PORC5"
    
    con = "select " & horas & "," & ccosto & "," & PORC & "," & INGRESOS & _
    "," & descuentos & " from plahistorico where cia='" & wcia & "' and " & "proceso='" & _
    VTipobol & "' and placod='" & Trim(Txtcodpla.Text) & "' and semana='" & _
    VSemana & "' and Year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and day(fechaproceso)=" & Vdia & " and status<>'*'"

    RZ.Open con, cn, adOpenStatic, adLockReadOnly
    
    If RZ.RecordCount = 0 Then Exit Sub
    
    BoletaCargada = True
    
    rshoras.MoveFirst
    Do While Not rshoras.EOF
      With rshoras
           cod = "h" & Format(.Fields.Item(0), "00")
           If .Fields.Item(0) <> 14 Then
            VALOR = RZ.Fields.Item(cod).Value
           Else
            VALOR = RZ.Fields.Item(cod).Value / 8
           End If
           .Fields.Item(2) = IIf(VALOR = 0, Null, VALOR)
              
           .MoveNext
      End With
    Loop
  
  'LLENA CENTRO DE COSTO
  con = "Select cod_maestro2,descrip from maestros_2 where status<>'*' " & _
  "and ciamaestro= '01044' ORDER BY cod_maestro2"
  Rt.Open con, cn, adOpenStatic, adLockReadOnly
  
    For I = 0 To 4
        con = "ccosto" & (I + 1)
        'rsccosto.AddNew
        VALOR = RZ.Fields.Item("PORC" & (I + 1))
        If VALOR <> 0 Then
           rsccosto.Fields("MONTO") = VALOR
           
           VALOR = RZ.Fields.Item(con)
           Rt.Filter = "cod_maestro2='" & VALOR & "'"
           If Not Rt.EOF Then
            rsccosto.Fields("CODIGO") = Trim(Rt("COD_MAESTRO2"))
            rsccosto.Fields("DESCRIPCION") = Rt("DESCRIP")
           End If
           rsccosto.MoveNext
        End If
    Next
  
    Rt.Close
    'RZ.Close
    'Set RZ = Nothing
    
    
    'LLENA PAGOSADICIONALES
    rspagadic.MoveFirst
    Do While Not rspagadic.EOF
          VALOR = "I" & rspagadic.Fields("CODIGO")
          rspagadic.Fields("MONTO") = RZ.Fields(VALOR)
          rspagadic.MoveNext
    Loop
    
    'llena descuentos
    rsdesadic.MoveFirst
    Do While Not rsdesadic.EOF
       VALOR = "d" & rsdesadic.Fields("CODIGO")
       rsdesadic.Fields("MONTO") = RZ.Fields(VALOR)
       rsdesadic.MoveNext
    Loop
    
End Sub

Sub MUESTRA_CUENTACORRIENTE()
Dim RX As New ADODB.Recordset
Dim I As Integer
Dim sumporc As Currency
On Error GoTo CORRIGE
MAXROW = 0
'VSemana
'CON = "SELECT DES.MONTO FROM PLACTACTE CTA,PLADESCTA DES WHERE " & _
'"CTA.PLACOD='" & Trim(Txtcodpla.Text) & "' AND (CTA.IMPORTE-CTA.PAGO_ACUENTA)>0 AND " & _
'"DES.CODAUXINTERNO=CTA.CODAUXINTERNO AND RIGHT(DES.FECHA,2)='" & _
'VSemana & "' AND CTA.STATUS<>'*'"

If BoletaCargada Then Exit Sub

If VTipobol = "01" Then
    con = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(b.importe),0) as pago_acuenta,a.partes,"
    con = con & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as saldomes,"
    con = con & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as quincena"
    con = con & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
    con = con & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' AND a.STATUS<>'*'"
    con = con & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
Else
    con = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(b.importe),0) as pago_acuenta,a.partes,"
    con = con & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as saldomes,"
    con = con & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as quincena"
    con = con & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
    con = con & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & wcia & "' and a.placod='" & Trim(Txtcodpla.Text) & "' AND a.STATUS<>'*' and a.sn_grati=1"
    con = con & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
End If

RX.Open con, cn, adOpenStatic, adLockReadOnly

   If RX.RecordCount > 0 Then

      rsdesadic.MoveFirst
      Erase ArrDsctoCTACTE
      MAXROW = 0
      Do While Not rsdesadic.EOF
         If rsdesadic("CODIGO") = "07" Then
            Do While Not RX.EOF
                ReDim Preserve ArrDsctoCTACTE(0 To 4, 0 To MAXROW)
                
                ArrDsctoCTACTE(0, MAXROW) = Txtcodpla
                ArrDsctoCTACTE(1, MAXROW) = RX!id_doc
                ArrDsctoCTACTE(3, MAXROW) = RX!tipo
                ArrDsctoCTACTE(4, MAXROW) = 0
                
                If RX("partes") >= (RX("IMPORTE") - RX("PAGO_ACUENTA")) Then
                   rsdesadic("MONTO") = rsdesadic("MONTO") + (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                   If wTipoDoc = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("IMPORTE") - RX("PAGO_ACUENTA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                Else
                   rsdesadic("MONTO") = rsdesadic("MONTO") + IIf(RX("partes") = 0, 0, RX("partes") - RX("QUINCENA"))
                   If wTipoDoc = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("partes") - RX("QUINCENA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = RX("partes") - RX("QUINCENA")
                End If
                
'                If wTipoDoc = False Then
'                    If rsdesadic("MONTO") > 0 Then
'                        rsdesadic("MONTO") = Round(rsdesadic("MONTO") / 2, 2)
'                    End If
'                End If
                
                MAXROW = MAXROW + 1
                
                RX.MoveNext
            Loop
            
             If wTipoDoc = False Then
                If rsdesadic("MONTO") > 0 Then
                    rsdesadic("MONTO") = Round(rsdesadic("MONTO") / 2, 2)
                End If
            End If
                
            sumporc = 0
            If rsdesadic("MONTO") > 0 Then
                For I = 0 To MAXROW - 1
                
                    ArrDsctoCTACTE(4, I) = Round((ArrDsctoCTACTE(2, I) * 100) / rsdesadic("MONTO"), 2)
                    sumporc = sumporc + Round((ArrDsctoCTACTE(2, I) * 100) / rsdesadic("MONTO"), 2)
                    If I = MAXROW - 1 Then
                        If sumporc > 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) - (sumporc - 100)
                        ElseIf sumporc < 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) + (100 - sumporc)
                        End If
                    End If
                Next
            End If
         End If
         rsdesadic.MoveNext
      Loop
   End If

RX.Close


Exit Sub
CORRIGE:
MsgBox "Error :" & Err.Description, vbCritical, Me.Caption

End Sub

Public Function convierte_cant() As Double
Dim RX As New ADODB.Recordset
Dim con As String
Dim importe As Double

con = "select plar.* from plaremunbase plar,platemphist plat where " & _
"plar.placod=plat.placod and plar.concepto='04' and plar.status<>'*'"
RX.Open con, cn, adOpenStatic, adLockReadOnly

'rx.Filter = "concepto='04'"
If RX.RecordCount = 0 Then Exit Function

con = RX("factor_horas")

If con <> 8 Then
   importe = RX("importe") / 30
Else
   importe = RX("importe")
End If

RX.Close

convierte_cant = importe

End Function

Private Sub NuevaBoleta()
Dim rs2 As ADODB.Recordset

On Error GoTo CORRIGE

BolDevengada = False
sn_quinta = True
cn.CursorLocation = adUseClient

Set rs = New ADODB.Recordset

sql$ = " EXEC sp_c_datos_personal '" & wcia & "','" & Trim(Txtcodpla.Text) & "'"

Set rs = cn.Execute(sql$)

If Not rs.EOF Then
   If rs!TipoTrabajador <> vTipoTra Or Not IsNull(rs!fcese) Then
      If rs!TipoTrabajador <> vTipoTra Then MsgBox "Trabajador no es del tipo seleccionado", vbExclamation, "Codigo N° => " & Txtcodpla.Text
      If Not IsNull(rs!fcese) Then MsgBox "Trabajador ya fue Cesado", vbExclamation, "Con Fecha => " & Format(rs!fcese, "dd/mm/yyyy")
      Txtcodpla.Text = ""
      Limpia_Boleta
      lblnombre.Caption = ""
      Lblctacte.Caption = "0.00"
      Lblcodaux.Caption = ""
      Lblcodafp.Caption = ""
      Lblnumafp.Caption = ""
      Lblbasico.Caption = ""
      Lbltope.Caption = ""
      Lblcargo.Caption = ""
      LblFingreso.Caption = ""
      VAltitud = ""
      VVacacion = ""
      VArea = ""
      VFechaNac = ""
      VFechaJub = ""
      Txtcodpla.SetFocus
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      Frame5.Enabled = False
   Else
      LblFingreso.Caption = Format(rs!fingreso, "mm/dd/yyyy")
      If Val(Right(LblFingreso.Caption, 4)) > Vano Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
         Limpia_Boleta
         LblFingreso.Caption = ""
         Exit Sub
      ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) > Vmes Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
         Limpia_Boleta
         LblFingreso.Caption = ""
         Exit Sub
      ElseIf Val(Right(LblFingreso.Caption, 4)) = Vano And Val(Left(LblFingreso.Caption, 2)) = Vmes And Val(Mid(LblFingreso.Caption, 4, 2)) > Val(Left(VFProceso, 2)) Then
         MsgBox "Fecha de Ingreso del Trabajador es Superior", vbCritical, "Calculo de Boletas"
         Limpia_Boleta
         LblFingreso.Caption = ""
         Exit Sub
      End If
      
      'cargamos si la person esta afecto o no a la quinta categoria
      sn_quinta = True
      If Trim(rs!quinta & "") = "N" Then sn_quinta = False
      '***********************************************************
      
      lblnombre.Caption = rs!nombre
      Lblcodaux.Caption = rs!codauxinterno
      Lblcodafp.Caption = rs!CodAfp
      Lblnumafp.Caption = Trim(rs!NUMAFP)
      VFechaNac = Format(rs!fnacimiento, "dd/mm/yyyy")
      VFechaJub = Format(rs!fec_jubila, "dd/mm/yyyy")
      Lbltope.Caption = rs!tipotasaextra
      
      If vTipoTra = "05" Then Lblcargo.Caption = rs!cargo: VAltitud = rs!altitud: VVacacion = rs!vacacion
      
      VArea = Trim(rs!Area)
      Frame2.Enabled = True
      Frame3.Enabled = True
      Frame4.Enabled = True
      Frame5.Enabled = True
      
      Lblbasico.Caption = rs!basico
      
      If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
      Do While Not rsccosto.EOF
        rsccosto.Delete
        rsccosto.MoveNext
      Loop
      
       If Trim(UCase(rs("ESSALUDVIDA"))) = "S" Then
         rsdesadic.MoveFirst
         rsdesadic.FIND "CODIGO='06'", 1 'ESSALUDVIDA
         If Not rsdesadic.EOF Then rsdesadic("MONTO") = rs!impessalud
      End If
        
      MSINDICATO = False
            
      If Trim(UCase(rs("sindicato"))) = "S" Then
         rsdesadic.FIND "CODIGO='15'", 1 'SOLIC SINDICATO
         If Not rsdesadic.EOF Then rsdesadic("MONTO") = rs!impsindicato
         
         MSINDICATO = True
      End If
      
        rsccosto.AddNew
        rsccosto.MoveFirst
        rsccosto!codigo = Trim(rs!Area)
        rsccosto!Descripcion = UCase(rs!descrip)
        rsccosto!Monto = "100.00"
        lbltot.Caption = "100.00"
        rsccosto!Item = VItem
      
      
      For I = I To 4
         If rsccosto.RecordCount < 5 Then rsccosto.AddNew
      Next I
      
      rsccosto.MoveFirst
      Dgrdccosto.Refresh
            
      rs.Close
      Txtcodpla.Enabled = False
   End If
ElseIf Trim(Txtcodpla.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodpla.Text
   Txtcodpla.Text = ""
   Limpia_Boleta
   lblnombre.Caption = ""
   Lblctacte.Caption = "0.00"
   Lblcodaux.Caption = ""
   Lblcodafp.Caption = ""
   Lblnumafp.Caption = ""
   Lblbasico.Caption = ""
   LblFingreso.Caption = ""
   Lbltope.Caption = ""
   Lblcargo.Caption = ""
   VAltitud = ""
   VVacacion = ""
   VArea = ""
   VFechaNac = ""
   VFechaJub = ""
   Txtcodpla.SetFocus
   Frame2.Enabled = False
   Frame3.Enabled = False
   Frame4.Enabled = False
   Frame5.Enabled = False
End If
VItem = 0
If VTipobol <> "01" Then
'   Dgrdhoras.Columns(1).Locked = False
'Else
   'Dgrdhoras.Columns(1).Locked = True
End If

Call Carga_Horas_NEW

If VTipobol = "02" Then Otros_Pagos_Vac
If wTipoDoc = True And VTipobol = "01" Then
   If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      If rsdesadic!codigo = "09" Then
         sql$ = "select totneto from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
         If (fAbrRst(rs2, sql$)) Then rsdesadic!Monto = rs2(0)
         rs2.Close
      End If
      rsdesadic.MoveNext
   Loop
End If
If wTipoDoc = True Then
   sql = "select sum(importe-pago_acuenta) from plactacte where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and importe-pago_acuenta<>0 and status<>'*' and importe>0"
   If (fAbrRst(rs, sql$)) Then
      If IsNull(rs(0)) Then Lblctacte.Caption = "0.00" Else Lblctacte.Caption = Format(rs(0), "###,###,###.00")
   End If
   rs.Close
End If

If VTipobol = "02" And vDevengue = False Then
   sql = "select i16 from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
   If (fAbrRst(rs, sql$)) Then
      rs.MoveFirst
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         rspagadic.Delete
         rspagadic.MoveNext
      Loop
      If rs!I16 <> 0 Then
         rspagadic.AddNew
         rspagadic!codigo = "16"
         rspagadic!Monto = rs!I16
         rspagadic!Descripcion = "OTROS PAGOS"
      End If
      MsgBox "Trabajador Tiene Vacaciones Devengadas" & Chr(13) & "No podra modificar Datos, Solo Grabar", vbInformation, "Vacaciones Devengadas"
      BolDevengada = True
      Frame2.Enabled = False
      Frame3.Enabled = False
      Frame4.Enabled = False
      Frame5.Enabled = False
   End If
End If
If VTipobol = "01" Then MUESTRA_CUENTACORRIENTE
Exit Sub
CORRIGE:
  MsgBox "Error:" & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Carga_Horas_NEW()
Dim rs2 As ADODB.Recordset
Dim mbol As String
Dim mconceptos As String
Dim mhor As Currency
Dim mdiasfalta As String
Dim mdiasferiado As String
Dim wBeginMonth As String
Dim mHourTra As Currency
Dim con As String
Dim I As Integer

On Error GoTo CORRIGE

If Trim(Txtcodpla.Text) = "" Then Exit Sub

sql$ = "select iniciomes from cia where cod_cia='" & _
wcia & "' and status<>'*'"

If (fAbrRst(rs, sql$)) Then
   If IsNull(rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = rs!iniciomes
End If
rs.Close

If Trim(wBeginMonth) <> "1" Then
   If Vmes = 1 Then
      wBeginMonth = Format(wBeginMonth, "00") & "/12/" & Format(Vano - 1, "0000")
   Else
     wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes - 1, "00") & "/" & Format(Vano, "0000")
   End If
Else
   wBeginMonth = Format(wBeginMonth, "00") & "/" & Format(Vmes, "00") & "/" & Format(Vano, "0000")
End If

If VPerPago = "02" Then wBeginMonth = VfDel

If rshoras.RecordCount > 0 Then rshoras.MoveFirst
Do While Not rshoras.EOF
   rshoras.Delete
   rshoras.MoveNext
Loop
 
sql$ = wInicioTrans
cn.Execute sql$

sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
cn.Execute sql

sql$ = wFinTrans
cn.Execute sql$
mdiasfalta = 0
mdiasferiado = 0
VHorasnormal = VHoras
If VTipobol = "01" Then mbol = "N"
If VTipobol = "02" Then mbol = "V"
If VTipobol = "03" Then mbol = "G"

wciamae = Funciones.Determina_Maestro("01077")

If VTipobol = "01" Or wTipoDoc <> True Then

   sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    
   If wTipoDoc = True Then
     
      Select Case VPerPago
             Case Is = "04" 'Mensual
                  sql$ = sql$ & "select distinct(concepto) from platareo where cia='" & wcia & "' and fecha " _
                       & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + Coneccion.FormatTimei & _
                       "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
                       
             Case Is = "02" 'Semanal
                  sql$ = sql$ & "select distinct(concepto) from platareo where cia='" & wcia & "' and fecha " _
                       & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      End Select
   Else
      sql$ = sql$ & "select distinct(concepto) from platareo where cia='" & wcia & "' and fecha " _
           & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & _
           "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
           & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      mbol = "N"
   End If
   
   If (fAbrRst(rs, sql$)) Then
      rs.MoveFirst
      mconceptos = ""
      Do While Not rs.EOF
         mconceptos = mconceptos & "'" & Trim(rs!concepto) & "',"
         rs.MoveNext
      Loop
      mconceptos = "(cod_maestro2 in (" & Mid(mconceptos, 1, Len(mconceptos) - 1) & ")"
   End If
   
   If rs.State = 1 Then rs.Close
   
   If Trim(mconceptos) <> "" Then
      sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 and " & mconceptos
   Else
      sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 and "
   End If
   
   If vTipoTra = "01" Then
      If Trim(mconceptos) <> "" Then
         sql$ = sql$ & " or cod_maestro2 in('01'))"
      Else
         sql$ = sql$ & "cod_maestro2 in('01')"
      End If
   Else
     If Trim(mconceptos) <> "" Then
         sql$ = sql$ & " or cod_maestro2 in ('14','01','02','03','04','06','05','07','08','09','10','17','11','12','13','15','18','19'))"
     Else
        sql$ = sql$ & "cod_maestro2 in  ('14','01','02','03','04','06','05','07','08','09','10','17','11','12','13','15','18','19')"
     End If
   End If
   
'3     '-feriado
'4     'perm. pag.
'6     '-enferm. no pagadas
'5     '-enferm. pagadas
'7     'accidente de trabajo
'8     'faltas injustificadas
'9     'suspencion
'10    'extras l-s
'11    'extras d-f
'12    'vacaciones
'13    'sobretasa
'15    'otros
   
   
   sql$ = sql$ & wciamae
   
   If (fAbrRst(rs, sql$)) Then
      rs.MoveFirst
      con = "140102030406050708091017111213151819"
      
       
      For I = 1 To (Len(con) / 2)
      'Do While Not RS.EOF
         rs.Filter = "COD_MAESTRO2='" & Mid(con, I + (I - 1), 2) & "'"
        If Not rs.EOF Then  ' {<MA>} 01/02/2007
            rshoras.AddNew
            rshoras!codigo = Trim(rs!cod_maestro2)
            rshoras!Descripcion = UCase(rs!descrip)
            If Trim(rs("cod_maestro2")) = "14" Then
               rshoras("MONTO") = 6
            End If

            If vTipoTra = "02" Then
                If rshoras!codigo = "01" Then
                    'rshoras!Monto = 6 * 8
                    rshoras!Monto = 7 * 8
                End If
            
                If rshoras!codigo = "02" Then
                    rshoras!Monto = (6 / 6) * 8
                End If
            End If
            
         mhor = 0
         
         sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                  
         Select Case VPerPago
               Case Is = "04"
                    sql$ = sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where cia='" & wcia & "' and fecha " _
                         & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & _
                         Trim(rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
               Case Is = "02"
                    sql$ = sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where cia='" & wcia & "' fecha " _
                         & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & _
                         Format(VfAl, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & Trim(rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
         End Select
 
         If (fAbrRst(rs2, sql$)) Then
            rs2.MoveFirst
            Do While Not rs2.EOF
               Select Case rs2!motivo
                      Case Is = "DI"
                            mhor = mhor + (rs2!tiempo * 8 * 60)
                            mdiasfalta = mdiasfalta + rs2!tiempo
                            If rs!cod_maestro2 = "03" Then mdiasferiado = mdiasferiado + rs2!tiempo
                      Case Is = "HO"
                           mhor = mhor + Int(rs2!tiempo) * 60 + ((rs2!tiempo - Int(rs2!tiempo)) * 100)
                      Case Is = "MI"
                           mhor = mhor + rs2!tiempo
               End Select
               rs2.MoveNext
            Loop
            If rs2.State = 1 Then rs2.Close
            mhor = Int(mhor / 60) + ((mhor Mod 60) / 100)
                        
            rshoras!Monto = IIf(mhor = 0, Null, mhor)
            
            If Trim(rs!flag2) = "-" Then
               VHorasnormal = VHorasnormal - mhor
            End If
         End If
         
         If rshoras("CODIGO") = "14" And IsNull(rshoras("MONTO")) Then
            rshoras("MONTO") = 6
         End If
         
         rs.MoveNext
      End If
      Next
      
      'TIPICO
       
      rshoras.MoveFirst
      mHourTra = 0
      If rshoras!codigo = "01" Then
         If wTipoDoc = True Then
            mHourTra = Calc_Horas_FecIng(wBeginMonth)
            rshoras!Monto = VHorasnormal
         Else
            VHorasnormal = 0
            If VNewBoleta = True Then VHorasnormal = Calc_Horas_Quincena(wBeginMonth)
            rshoras!Monto = VHorasnormal
         End If
         If rshoras.RecordCount > 1 Then rshoras.MoveNext
      End If
      If rshoras!codigo = "02" Then rshoras!Monto = ((6 - mdiasfalta + mdiasferiado) * 8) / 6
      If mHourTra <> 0 Then
         rshoras.AddNew
         rshoras!codigo = "03"
         rshoras!Descripcion = "FERIADOS"
         rshoras!Monto = mHourTra
      End If
      
   End If
Else
    sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0"
    sql$ = sql$ & wciamae
    cn.CursorLocation = adUseClient
    Set rs = New ADODB.Recordset
    Set rs = cn.Execute(sql$, 64)
    If rshoras.RecordCount > 0 Then
       rshoras.MoveFirst
       Do While Not rshoras.EOF
          rshoras.Delete
          rshoras.MoveNext
       Loop
    End If
    If Not rs.RecordCount > 0 Then MsgBox "No Existen Horas Registradas", vbCritical, TitMsg: Exit Sub
    rs.MoveFirst
    Do While Not rs.EOF
       rshoras.AddNew
       rshoras!codigo = Trim(rs!cod_maestro2)
       rshoras!Descripcion = Trim(rs!descrip)
       If VTipobol = "02" Then rshoras!Monto = VHoras Else rshoras!Monto = "0.00"
       rs.MoveNext
    Loop
    If rs.State = 1 Then rs.Close
End If

'solo para datos guardados
Call Llena_Horas_New

Exit Sub
CORRIGE:
MsgBox "Error :" & Err.Description, vbCritical, "Sistema de Planillas"
End Sub


Sub Llena_Horas_New()
    Dim RZ As New ADODB.Recordset
    Dim Rt As New ADODB.Recordset
    Dim con As String, cod As String
    Dim horas As Variant, VALOR As Variant
    Dim I As Integer, ccosto As String, PORC As String
    Dim INGRESOS As String
        
    For I = 1 To 30
         horas = horas & "H" & Format(I, "00") & ","
    Next
    horas = Left(horas, Len(horas) - 1)
    
    'PAGOS ADICIONALES
    If rspagadic.RecordCount > 0 Then
       rspagadic.MoveFirst
       Do While Not rspagadic.EOF
                INGRESOS = INGRESOS & "I" & rspagadic(0) & ","
                rspagadic.MoveNext
       Loop
       INGRESOS = Left(INGRESOS, Len(INGRESOS) - 1)
    End If
    
    'DESCUENTOS
    If rsdesadic.RecordCount > 0 Then
       rsdesadic.MoveFirst
       Do While Not rsdesadic.EOF
                descuentos = descuentos & "d" & rsdesadic(0) & ","
                rsdesadic.MoveNext
       Loop
       descuentos = Left(descuentos, Len(descuentos) - 1)
    End If
        
    ccosto = "ccosto1,ccosto2,ccosto3,ccosto4,ccosto5"
    PORC = "PORC1,PORC2,PORC3,PORC4,PORC5"
    
    con = "select " & horas & "," & ccosto & "," & PORC & "," & INGRESOS & _
    "," & descuentos & " from plahistorico where cia='" & wcia & "' and " & "proceso='" & _
    VTipobol & "' and placod='" & Trim(Txtcodpla.Text) & "' and semana='" & _
    VSemana & "' and Year(fechaproceso)=" & Vano & " and status<>'*'"

    RZ.Open con, cn, adOpenStatic, adLockReadOnly
    
    If RZ.RecordCount = 0 Then Exit Sub
    
    rshoras.MoveFirst
    Do While Not rshoras.EOF
      With rshoras
           cod = "h" & Format(.Fields.Item(0), "00")
           If .Fields.Item(0) <> 14 Then
            VALOR = RZ.Fields.Item(cod).Value
           Else
            VALOR = RZ.Fields.Item(cod).Value / 8
           End If
           .Fields.Item(2) = IIf(VALOR = 0, Null, VALOR)
              
           .MoveNext
      End With
    Loop
  
  'LLENA CENTRO DE COSTO
  con = "Select cod_maestro2,descrip from maestros_2 where status<>'*' " & _
  "and ciamaestro= '01044' ORDER BY cod_maestro2"
  Rt.Open con, cn, adOpenStatic, adLockReadOnly
  
    For I = 0 To 4
        con = "ccosto" & (I + 1)
        'rsccosto.AddNew
        VALOR = RZ.Fields.Item("PORC" & (I + 1))
        If VALOR <> 0 Then
           rsccosto.Fields("MONTO") = VALOR
           
           VALOR = RZ.Fields.Item(con)
           Rt.Filter = "cod_maestro2='" & VALOR & "'"
           rsccosto.Fields("CODIGO") = Trim(Rt("COD_MAESTRO2"))
           rsccosto.Fields("DESCRIPCION") = Rt("DESCRIP")
           rsccosto.MoveNext
        End If
    Next
  
    Rt.Close
    'RZ.Close
    'Set RZ = Nothing
    
    
    'LLENA PAGOSADICIONALES
    rspagadic.MoveFirst
    Do While Not rspagadic.EOF
          VALOR = "I" & rspagadic.Fields("CODIGO")
          rspagadic.Fields("MONTO") = RZ.Fields(VALOR)
          rspagadic.MoveNext
    Loop
    
    'llena descuentos
    rsdesadic.MoveFirst
    Do While Not rsdesadic.EOF
       VALOR = "d" & rsdesadic.Fields("CODIGO")
       rsdesadic.Fields("MONTO") = RZ.Fields(VALOR)
       rsdesadic.MoveNext
    Loop
    
End Sub


Private Sub ProrrateaCtaCte(ByVal pNuevoValor As Currency)
Dim X As Integer

For X = 0 To MAXROW - 1
    ArrDsctoCTACTE(2, X) = Round(pNuevoValor * (ArrDsctoCTACTE(4, X) / 100), 2)
Next X

End Sub
