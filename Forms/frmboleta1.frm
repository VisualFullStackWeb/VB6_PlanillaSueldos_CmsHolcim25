VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmboleta1 
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   9420
   Begin VB.Frame Frame6 
      Caption         =   "Basicos"
      Height          =   4575
      Left            =   0
      TabIndex        =   33
      Top             =   2400
      Width           =   4455
      Begin MSDataGridLib.DataGrid DgrBasico 
         Height          =   4215
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7435
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6045
      Left            =   4560
      TabIndex        =   26
      Top             =   1080
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   10663
      _Version        =   393216
      TabOrientation  =   2
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmboleta1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmboleta1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmboleta1.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Descuentos Adicionales"
         Enabled         =   0   'False
         Height          =   2055
         Left            =   360
         TabIndex        =   35
         Top             =   3720
         Width           =   4335
         Begin MSDataGridLib.DataGrid DgrdDesAdic 
            Height          =   1695
            Left            =   120
            TabIndex        =   36
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
         Height          =   3615
         Left            =   -74640
         TabIndex        =   31
         Top             =   240
         Width           =   4335
         Begin MSDataGridLib.DataGrid DgrdPagAdic 
            Height          =   3255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   5741
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
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   4335
         Begin VB.ListBox LstCcosto 
            Height          =   840
            Left            =   480
            TabIndex        =   28
            Top             =   720
            Visible         =   0   'False
            Width           =   2655
         End
         Begin MSDataGridLib.DataGrid Dgrdccosto 
            Height          =   1455
            Left            =   120
            TabIndex        =   29
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
            TabIndex        =   30
            Top             =   1680
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8775
      Begin VB.TextBox Txtcodpla 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         MaxLength       =   8
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox Cmbturno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox Txtvacaf 
         Height          =   255
         Left            =   4440
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   645
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
      Begin MSMask.MaskEdBox Txtvaca 
         Height          =   255
         Left            =   7680
         TabIndex        =   6
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
         TabIndex        =   7
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
      Begin VB.Label Lbltope 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Lblcodafp 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Lblnombre 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Lblvaca2 
         AutoSize        =   -1  'True
         Caption         =   "F. Inicio Vaca."
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Periodo Vacacional"
         Height          =   195
         Left            =   5760
         TabIndex        =   16
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Lblcodaux 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label LblFingreso 
         Height          =   135
         Left            =   3120
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Lblvaca1 
         AutoSize        =   -1  'True
         Caption         =   "F. Ret. Vac."
         Height          =   195
         Left            =   3480
         TabIndex        =   12
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Lblcargo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6600
         TabIndex        =   11
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Lblnumafp 
         Height          =   135
         Left            =   5640
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Lblbasico 
         Height          =   15
         Left            =   7200
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "Constantes de Calculo"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
      Begin MSDataGridLib.DataGrid Dgrdhoras 
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
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
               ColumnWidth     =   2654.929
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
   Begin Threed.SSPanel Panelprogress 
      Height          =   735
      Left            =   1560
      TabIndex        =   22
      Top             =   3240
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
         TabIndex        =   23
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
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
      Left            =   4440
      TabIndex        =   25
      Top             =   4740
      Width           =   2655
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
      Left            =   7080
      TabIndex        =   24
      Top             =   4740
      Width           =   1695
   End
End
Attribute VB_Name = "frmboleta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VokDevengue As Boolean
Dim vDevengue As Boolean
Dim Mstatus As String
 
Dim BolDevengada As Boolean
Dim NumDev As Integer

Dim rshoras As New Recordset
Dim rsccosto As New Recordset
Dim rspagadic As New Recordset
Dim rsdesadic As New Recordset
Dim VTipobol As String
Dim vTipoTra As String
Dim VTurno As String
Dim vItem As Integer
Dim VSemana As String
Dim VFProceso As String
Dim VPerPago As String
Dim VNewBoleta As Boolean
Dim Vano As Integer
Dim Vmes As Integer
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
Dim mcancel As Boolean
Dim VFechaNac As String
Dim VFechaJub As String
Dim VObra As String
Dim manos As Integer
Dim mHourDay As Currency
Dim rsbasico As New ADODB.Recordset

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
            If DgrdDesAdic.Columns(2) = "07" And CCur(DgrdDesAdic.Columns(1)) > CCur(Lblctacte.Caption) Then
               MsgBox "El Importe no debe ser Mayor al Saldo", vbInformation, "Cuenta Corriente"
               DgrdDesAdic.Columns(1) = "0.00"
            End If
End Select
End Sub

Private Sub DgrdDesAdic_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If DgrdDesAdic.Columns(2) = "09" Then KeyAscii = 0
End Sub

Private Sub DgrdDesAdic_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If DgrdDesAdic.Columns(2) = "09" Then Cancel = True
End Sub

Private Sub DgrdDesAdic_KeyPress(KeyAscii As Integer)
If DgrdDesAdic.Columns(2) = "09" Then KeyAscii = 0
End Sub

Private Sub DgrdHoras_AfterColEdit(ByVal ColIndex As Integer)
If Dgrdhoras.Columns(1).Text = "" Then Dgrdhoras.Columns(1) = "0.00"
Dgrdhoras.Columns(1) = Format(Dgrdhoras.Columns(1), "###,###.00")
If Not rshoras.EOF Then rshoras.MoveNext
End Sub

Private Sub Form_Activate()
If wtipodoc = True Then
   Me.Caption = "Ingreso de Boletas"
   Me.BackColor = &H80000001
Else
   Me.Caption = "Adelanto de Quincena"
   Me.BackColor = &H808000
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
Sql$ = "Select flag2 from maestros_2 where cod_maestro2='01' and status<>'*'"
Sql$ = Sql$ & wciamae
mHourDay = 0
If (fAbrRst(Rs, Sql$)) Then mHourDay = Val(Rs!flag2)
Rs.Close
Call Llena_basico
End Sub
Private Sub Crea_Rs()
    If rshoras.State = 1 Then rshoras.Close
    rshoras.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rshoras.Fields.Append "descripcion", adChar, 100, adFldIsNullable
    rshoras.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rshoras.Open
    Set Dgrdhoras.DataSource = rshoras
    
    If rsccosto.State = 1 Then rsccosto.Close
    rsccosto.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsccosto.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsccosto.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rsccosto.Fields.Append "item", adChar, 2, adFldIsNullable
    rsccosto.Open
    Set Dgrdccosto.DataSource = rsccosto
    
    If rspagadic.State = 1 Then rspagadic.Close
    rspagadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rspagadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rspagadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rspagadic.Open
    Set DgrdPagAdic.DataSource = rspagadic
    
    If rsdesadic.State = 1 Then rsdesadic.Close
    rsdesadic.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsdesadic.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsdesadic.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rsdesadic.Open
    Set DgrdDesAdic.DataSource = rsdesadic
    
     'Basico
    If rsbasico.State = 1 Then rsbasico.Close
    rsbasico.Fields.Append "descripcion", adChar, 35, adFldIsNullable
    rsbasico.Fields.Append "tipo", adChar, 15, adFldIsNullable
    'rsbasico.Fields.Append "moneda", adChar, 3, adFldIsNullable
    rsbasico.Fields.Append "importe", adCurrency, 18, adFldIsNullable
    rsbasico.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsbasico.Fields.Append "codtipo", adChar, 2, adFldIsNullable
    rsbasico.Fields.Append "horas", adInteger, 4, adFldIsNullable
    rsbasico.Open
    Set DgrBasico.DataSource = rsbasico
    DgrBasico.Columns(0).Locked = True
    
    
End Sub
Private Sub Procesa()
Dim wciamae As String
Dim rs2 As ADODB.Recordset
Dim I, pos As Integer

'Pagos Adicionales
Sql$ = "Select * from placonstante where cia='" & Trim(wcia) & "' and tipomovimiento='02' and calculo='N' and status<>'*' order by codinterno"

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If rspagadic.RecordCount > 0 Then
   rspagadic.MoveFirst
   Do While Not rspagadic.EOF
      rspagadic.Delete
      rspagadic.MoveNext
   Loop
End If
If Not Rs.RecordCount > 0 Then Exit Sub
Rs.MoveFirst
Do While Not Rs.EOF
   rspagadic.AddNew
   rspagadic!Codigo = Rs!codinterno
   rspagadic!Descripcion = Rs!Descripcion
   rspagadic!Monto = "0.00"
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'=====================================

'Descuentos Adicionales
If wtipodoc = True Then
   Sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='03' and adicional='S' and status<>'*' order by codinterno"
Else
   Sql$ = "Select * from placonstante where cia='" & wcia & "' and tipomovimiento='03' and adicional='S' and status<>'*' and codinterno<>'09' order by codinterno"
End If
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If rsdesadic.RecordCount > 0 Then
   rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      rsdesadic.Delete
      rsdesadic.MoveNext
   Loop
End If
If Not Rs.RecordCount > 0 Then Exit Sub

Rs.MoveFirst
Do While Not Rs.EOF
   rsdesadic.AddNew
   rsdesadic!Codigo = Rs!codinterno
   rsdesadic!Descripcion = Trim(Rs!Descripcion)
   rsdesadic!Monto = "0.00"
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
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
Public Sub Txtcodpla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Cmbturno.SetFocus
End Sub
Private Sub Txtcodpla_LostFocus()
Dim xciamae As String
Dim cod As String
Dim rs2 As ADODB.Recordset

BolDevengada = False
cn.CursorLocation = adUseClient

Set Rs = New ADODB.Recordset
cod = "01055"

Sql$ = "SELECT GENERAL FROM maestros where " & _
"right(ciamaestro,3)='" & Right(cod, 3) & "' and " & _
"status<>'*' "

If (fAbrRst(Rs, Sql$)) Then
   If Rs!General = "S" Then
      xciamae = " and right(ciamaestro,3)= '" & Right(cod, 3) & "'"
   Else
      xciamae = " and ciamaestro= '" & wcia + Right(cod, 3) & "'"
   End If
End If

If Rs.State = 1 Then Rs.Close

Sql$ = Funciones.nombre()
Sql$ = Sql$ & "codauxinterno,a.status,a.tipotrabajador,a.fingreso," & _
     "a.fcese,a.codafp,a.numafp,a.area,a.placod," & _
     "a.codauxinterno,b.descrip,a.tipotasaextra," & _
     "a.cargo,a.altitud,a.vacacion,a.area,a.fnacimiento," & _
     "a.fec_jubila " & _
     "from planillas a,maestros_2 b where a.status<>'*'"
     Sql$ = Sql$ & xciamae
     Sql$ = Sql$ & " and a.tipotrabajador=b.cod_maestro2 " _
     & "and cia='" & wcia & "' AND placod='" & Trim(Txtcodpla.Text) & "'" & xciamae
     Sql$ = Sql$ & " order by nombre"


cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$)

If Rs.RecordCount > 0 Then
   If Rs!TipoTrabajador <> vTipoTra Or Not IsNull(Rs!fcese) Then
      If Rs!TipoTrabajador <> vTipoTra Then MsgBox "Trabajador no es del tipo seleccionado", vbExclamation, "Codigo N° => " & Txtcodpla.Text
      If Not IsNull(Rs!fcese) Then MsgBox "Trabajador ya fue Cesado", vbExclamation, "Con Fecha => " & Format(Rs!fcese, "dd/mm/yyyy")
      Txtcodpla.Text = ""
      Limpia_Boleta
      Lblnombre.Caption = ""
      Lblctacte.Caption = "0.00"
      Lblcodaux.Caption = ""
      Lblcodafp.Caption = ""
      Lblnumafp.Caption = ""
      LblBasico.Caption = ""
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
      LblFingreso.Caption = Format(Rs!fIngreso, "mm/dd/yyyy")
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
      Lblnombre.Caption = Rs!nombre
      Lblcodaux.Caption = Rs!codauxinterno
      Lblcodafp.Caption = Rs!CodAfp
      Lblnumafp.Caption = Trim(Rs!NUMAFP)
      VFechaNac = Format(Rs!fnacimiento, "dd/mm/yyyy")
      VFechaJub = Format(Rs!fec_jubila, "dd/mm/yyyy")
      Lbltope.Caption = Rs!tipotasaextra
      If vTipoTra = "05" Then Lblcargo.Caption = Rs!Cargo
      If vTipoTra = "05" Then VAltitud = Rs!altitud
      If vTipoTra = "05" Then VVacacion = Rs!vacacion
      VArea = Rs!Area
      Frame2.Enabled = True
      Frame3.Enabled = True
      Frame4.Enabled = True
      Frame5.Enabled = True
      'Basico
      Sql$ = "select importe from plaremunbase where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and concepto='01' and status<>'*'"
     
      LblBasico.Caption = ""
      If (fAbrRst(rs2, Sql$)) Then LblBasico.Caption = rs2!importe
      If rs2.State = 1 Then rs2.Close
      'Centro de Costo
      If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
      Do While Not rsccosto.EOF
        rsccosto.Delete
        rsccosto.MoveNext
      Loop
      
      xciamae = Funciones.Determina_Maestro("01044")
    
      Sql$ = "Select cod_maestro2,descrip from maestros_2 where " & _
      "status<>'*' and cod_maestro2='" & Rs!Area & "'"
      Sql$ = Sql$ & xciamae
          
      Set Rs = cn.Execute(Sql$)
      
      'LLENA GRID CENTRO DE COSTOS
      If Rs.RecordCount > 0 Then
         rsccosto.AddNew
         rsccosto.MoveFirst
         rsccosto!Codigo = Trim(Rs!cod_maestro2)
         rsccosto!Descripcion = Rs!DESCRIP
         rsccosto!Monto = "100.00"
         lbltot.Caption = "100.00"
         rsccosto!Item = vItem
      End If
      For I = I To 4
         If rsccosto.RecordCount < 5 Then rsccosto.AddNew
      Next I
      rsccosto.MoveFirst
      Dgrdccosto.Refresh
      If Rs.State = 1 Then Rs.Close
      Txtcodpla.Enabled = False
   End If
ElseIf Trim(Txtcodpla.Text) <> "" Then
   MsgBox "Codigo de Planilla No Existe ", vbExclamation, "Codigo N° => " & Txtcodpla.Text
   Txtcodpla.Text = ""
   Limpia_Boleta
   Lblnombre.Caption = ""
   Lblctacte.Caption = "0.00"
   Lblcodaux.Caption = ""
   Lblcodafp.Caption = ""
   Lblnumafp.Caption = ""
   LblBasico.Caption = ""
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
vItem = 0
If VTipobol <> "01" Then
   Dgrdhoras.Columns(1).Locked = False
Else
   Dgrdhoras.Columns(1).Locked = True
End If
 Carga_Horas
If VTipobol = "02" Then Otros_Pagos_Vac
If wtipodoc = True And VTipobol = "01" Then
   If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
   Do While Not rsdesadic.EOF
      If rsdesadic!Codigo = "09" Then
         Sql$ = "select totneto from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
         If (fAbrRst(rs2, Sql$)) Then rsdesadic!Monto = rs2(0)
         rs2.Close
      End If
      rsdesadic.MoveNext
   Loop
End If
If wtipodoc = True Then
   Sql = "select sum(importe-pago_acuenta) from plactacte where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and importe-pago_acuenta<>0 and status<>'*' and importe>0"
   If (fAbrRst(Rs, Sql$)) Then
      If IsNull(Rs(0)) Then Lblctacte.Caption = "0.00" Else Lblctacte.Caption = Format(Rs(0), "###,###,###.00")
   End If
   Rs.Close
End If

If VTipobol = "02" And vDevengue = False Then
   Sql = "select i16 from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         rspagadic.Delete
         rspagadic.MoveNext
      Loop
      If Rs!I16 <> 0 Then
         rspagadic.AddNew
         rspagadic!Codigo = "16"
         rspagadic!Monto = Rs!I16
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
End Sub
Public Sub Carga_Boleta(Codigo As String, tipo As String, Nuevo As Boolean, semana As String, fproce As String, Tipot As String, perpago As String, horas As Integer, mdel As String, mal As String, obra As String, devengue As Boolean)
Dim MField As String
Dim wciamae As String
Load frmboleta1
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
VfDel = mdel
VfAl = mal
LstCcosto.Clear
wciamae = Funciones.Determina_Maestro("01044")
Sql$ = "Select cod_maestro2,descrip from maestros_2 where  status<>'*'"
Sql$ = Sql$ & wciamae

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then Rs.MoveFirst
Do Until Rs.EOF
   LstCcosto.AddItem Rs!DESCRIP & Space(100) & Rs!cod_maestro2
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'Turno
Sql$ = "Select codturno,descripcion from platurno where cia='" & wcia & "' and status<>'*' order by codturno"
Cmbturno.Clear
VTurno = ""

If (fAbrRst(Rs, Sql$)) Then
   If (Not Rs.EOF) Then
      Do Until Rs.EOF
         Cmbturno.AddItem Rs(1)
         Cmbturno.ItemData(Cmbturno.NewIndex) = Rs(0)
         Rs.MoveNext
       Loop
       Cmbturno.ListIndex = 0
    End If
    If Rs.State = 1 Then Rs.Close
End If
Select Case VTipobol 'VACACIONES
       Case Is = "02"
            Lblvaca1.Visible = True
            Txtvacai.Visible = True
            Lblvaca2.Caption = "F. Inicio Vaca."
            Txtvacaf.Visible = True
            Label4.Visible = True
            TxtVaca.Visible = True
            Txtcese.Visible = False
       Case Else
            Lblvaca1.Visible = False
            Txtvacai.Visible = False
            Txtvacaf.Visible = False
            Label4.Visible = False
            TxtVaca.Visible = False
            'Txtcese.Visible = True
End Select
Procesa
If VNewBoleta = False Then
   Txtcodpla.Text = Codigo
   Txtcodpla_LostFocus
   If rshoras.RecordCount > 0 Then rshoras.MoveFirst
   Do While Not rshoras.EOF
      rshoras!Monto = "0"
      rshoras.MoveNext
   Loop
   If VNewBoleta = False Then MsgBox "Los datos solo se mostraran como consulta " & Chr(13) & "Si desea modificar algun dato debera anular la boleta", vbInformation, "Sistema de Planilla"
   Sql$ = ""
   If wtipodoc = True Then
      Select Case VPerPago
             Case Is = "02"
                  Sql$ = "select * from plahistorico " _
                     & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
                     & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
             Case Is = "04"
                  Sql$ = "select * from plahistorico " _
                     & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and year( fechaproceso) = " & Vano & " And Month(fechaproceso) = " & Vmes & " " _
                     & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
      End Select
   Else
      Sql$ = "select * from plaquincena " _
         & "where cia='" & wcia & "' and year( fechaproceso) = " & Vano & " And Month(fechaproceso) = " & Vmes & " " _
         & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
   End If
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$, 64)
   If Rs.RecordCount > 0 Then
      Call rUbiIndCmbBox(Cmbturno, Format(Rs!turno, "00"), "00")
     
      If rshoras.RecordCount > 0 Then rshoras.MoveFirst
      Do While Not rshoras.EOF
         MField = "h" & rshoras!Codigo
         rshoras!Monto = Rs.Fields(MField)
         rshoras.MoveNext
      Loop
      If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
      Do While Not rspagadic.EOF
         MField = "i" & rspagadic!Codigo
         rspagadic!Monto = Rs.Fields(MField)
         rspagadic.MoveNext
      Loop
      
      If rsdesadic.RecordCount > 0 Then rsdesadic.MoveFirst
         Do While Not rsdesadic.EOF
            MField = "d" & rsdesadic!Codigo
            rsdesadic!Monto = Rs.Fields(MField)
            rsdesadic.MoveNext
         Loop
   End If
End If
Frame2.Enabled = False
Frame3.Enabled = False
Frame4.Enabled = False
Frame5.Enabled = False
If VTipobol = "02" Then Dgrdhoras.Enabled = False
vDevengue = devengue
If devengue = True Then Calcula_Devengue_Vaca
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


If vDevengue = True Then
   Mstatus = "D"
   VTurno = Format(NumDev, "00")
Else
   Mstatus = "T"
End If

If Trim(LblBasico.Caption) = "" Then
   MsgBox "Trabajador No Registra Sueldo Basico", vbInformation, "Boletas de Pago"
   LblBasico.Caption = ""
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
   If wtipodoc = True And VTipobol <> "02" Then
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

Sql$ = wInicioTrans
cn.Execute Sql$

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "insert into platemphist(cia,placod," & _
       "codauxinterno,proceso,fechaproceso,semana," & _
       "fechaingreso,turno,codafp,status,fec_crea," & _
       "tipotrab,obra,numafp,basico,fec_modi," & _
       "user_crea,user_modi) values('" & wcia & _
       "','" & Trim(Txtcodpla.Text) & "','" & _
       Lblcodaux.Caption & "','" & VTipobol & "','" & _
       Format(VFProceso, FormatFecha) & "','" & _
       VSemana & "','" & Format(LblFingreso.Caption, FormatFecha) & _
       "','" & Format(VTurno, "0") & "','" & _
       Lblcodafp.Caption & "','" & Mstatus & "'," & _
       FechaSys & ",'" & vTipoTra & "','" & VObra & _
       "','" & Lblnumafp.Caption & "'," & _
       CCur(LblBasico.Caption) & "," & FechaSys & _
       ",'" & wuser & "','" & wuser & "')"
 
cn.Execute Sql

'Horas
If rshoras.RecordCount > 0 Then
   rshoras.MoveFirst
   MqueryH = ""

   Do While Not rshoras.EOF
      If rshoras!Codigo <> "14" Then
        MqueryH = MqueryH & "h" & rshoras!Codigo & "=" & rshoras!Monto & ""
      Else
        MqueryH = MqueryH & "h" & rshoras!Codigo & "=" & rshoras!Monto * 8 & ""
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
      MqueryP = MqueryP & "i" & rspagadic!Codigo & "=" & rspagadic!Monto & ""
      rspagadic.MoveNext
      If Not rspagadic.EOF Then MqueryP = MqueryP & ","
   Loop
End If
'Debug.Print MqueryP

'Descuentos Adicionales
If rsdesadic.RecordCount > 0 Then
   rsdesadic.MoveFirst
   MqueryD = ""
   Do While Not rsdesadic.EOF
    
      MqueryD = MqueryD & "d" & rsdesadic!Codigo & "=" & rsdesadic!Monto & ""
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
         mcad = mcad & "ccosto" & Format(itemcosto, "0") & " = '" & rsccosto!Codigo & "'," & "porc" & Format(itemcosto, "0") & " = " & Str(rsccosto!Monto) & ","
      End If
   End If
   rsccosto.MoveNext
   itemcosto = itemcosto + 1
Loop

mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "Update platemphist set " & mcad & "," & MqueryH & "," & MqueryP & "," & MqueryD
Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & _
     Trim(Txtcodpla.Text) & "' and codauxinterno='" & _
     Trim(Lblcodaux.Caption) & "' and proceso='" & _
     Trim(VTipobol) & "' and fechaproceso='" & _
     Format(VFProceso, FormatFecha) & _
     "' and semana='" & VSemana & "' and status='" & _
     Mstatus & "'"
    
cn.Execute Sql
mcad = ""

'Calculo de ingresos
MqueryI = ""
For I = 1 To 50
    Select Case VTipobol
           Case Is = "01" 'Normal
                Sql$ = Me.F01(Format(I, "00"))
           Case Is = "02" 'Vacaciones
                Sql$ = Me.V01(Format(I, "00"))
           Case Is = "03" 'Gratificaciones
                Sql$ = Me.G01(Format(I, "00"))
    End Select
    
    If Trim(Sql$) <> "" Then
       cn.CursorLocation = adUseClient
       Set Rs = New ADODB.Recordset
       Set Rs = cn.Execute(Sql$, 64)

       If Rs.RecordCount > 0 Then
          Rs.MoveFirst
          If IsNull(Rs(0)) Or Rs(0) = 0 Then
          Else
             MqueryI = MqueryI & "i" & Format(I, "00") & " = " & Rs(0) & ","
           
          End If
       End If
       If Rs.State = 1 Then Rs.Close
    End If
Next

If MqueryI <> "" Then
   MqueryI = Mid(MqueryI, 1, Len(Trim(MqueryI)) - 1)
   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & " Update platemphist set " & MqueryI
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

'Calculo de Deducciones
MqueryCalD = ""
For I = 1 To 20
     
    Sql$ = Me.F02(Format(I, "00"))
    
    If mcancel = True Then
       VokDevengue = False
       MsgBox "Se Cancelo la Grabacion", vbCritical, "Calculo de Boleta"
       Sql$ = wCancelTrans
       cn.Execute Sql$
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    If Sql$ <> "" Then
       If (fAbrRst(Rs, Sql$)) Then
          Rs.MoveFirst
          If IsNull(Rs(0)) Or Rs(0) = 0 Then
          Else
             If I = 11 Then
                For J = 1 To 5
                   MqueryCalD = MqueryCalD & "d" & Format(I, "00") & Format(J, "0") & " = " & Rs(J - 1) & ","
                Next J
             ElseIf I = 13 And wtipodoc = False Then
                MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & Rs(0) / 2 & ","
             Else
                MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & Rs(0) & ","
             End If
          End If
       End If
       If Rs.State = 1 Then Rs.Close
    End If
Next I
If MqueryCalD <> "" Then
   MqueryCalD = Mid(MqueryCalD, 1, Len(Trim(MqueryCalD)) - 1)
   Sql$ = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & "Update platemphist set " & MqueryCalD
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

'Calculo de Aportaciones
MqueryCalA = ""
For I = 1 To 20
    Sql$ = F03(Format(I, "00"))
    If Sql$ <> "" Then
       cn.CursorLocation = adUseClient
       Set Rs = New ADODB.Recordset
       Set Rs = cn.Execute(Sql$, 64)
       If Rs.RecordCount > 0 Then
          Rs.MoveFirst
          If IsNull(Rs(0)) Or Rs(0) = 0 Then
          Else
             MqueryCalA = MqueryCalA & "a" & Format(I, "00") & " = " & Rs(0) & ","
          End If
       End If
       If Rs.State = 1 Then Rs.Close
    End If
Next
If MqueryCalA <> "" Then
   MqueryCalA = Mid(MqueryCalA, 1, Len(Trim(MqueryCalA)) - 1)
   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
   Sql$ = Sql$ & "Update platemphist set " & MqueryCalA
   Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
   cn.Execute Sql
End If

Dim mi As String, md As String, ma As String
mi = "i01+i02+i03+i04+i05+i06+i07+i08+i09+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23+i24+i25+i26+i27+i28+i29+i30+i31+i32+i33+i34+i35+i36+i37+i38+i39+i40+i41+i42+i43+i44+i45+i46+i47+i48+i49+i50"
md = "d01+d02+d03+d04+d05+d06+d07+d08+d09+d10+d11+d12+d13+d14+d15+d16+d17+d18+d19+d20"
ma = "a01+a02+a03+a04+a05+a06+a07+a08+a09+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20"
Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "update platemphist set totaling=" & mi & "," _
     & "d11=d111+d112+d113+d114+d115," _
     & "totalded=" & md & "," _
     & "totalapo=" & ma & "," _
     & "totneto=(" & mi & ")-" & "(" & md & ")"
   Sql$ = Sql$ & " where cia='" & wcia & "' and " & _
   " placod='" & Txtcodpla.Text & "' and " & _
   "codauxinterno='" & Trim(Lblcodaux.Caption) & "' and proceso='" & VTipobol & "' "
   Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute Sql

If wtipodoc = True Then
   Sql$ = "insert into plahistorico select * from platemphist"
Else
   Sql$ = "insert into plaquincena select * from platemphist"
End If
Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute Sql

Sql$ = "select cia,placod,proceso,fechaproceso,semana,tipotrab,d07 from platemphist"
Sql$ = Sql$ & " where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"
cn.Execute Sql

If (fAbrRst(Rs, Sql$)) Then
   Call Descarga_ctaCte(Rs!PlaCod, Rs!Proceso, Format(Rs!FechaProceso, "dd/mm/yyyy"), Rs!semana, Rs!TipoTrab, Rs!d07)
End If


Sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
Sql$ = Sql$ & " and cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and codauxinterno='" & Lblcodaux.Caption & "' and proceso='" & VTipobol & "' "
Sql$ = Sql$ & "and fechaproceso='" & Format(VFProceso, FormatFecha) & "' and semana='" & VSemana & "' and status='" & Mstatus & "'"


cn.Execute Sql

Sql$ = wFinTrans
cn.Execute Sql$
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
Lblnombre.Caption = ""
Lblctacte.Caption = "0.00"
Lblcodaux.Caption = ""
Lblcodafp.Caption = ""
Lblnumafp.Caption = ""
LblBasico.Caption = ""
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
 If wtipodoc = True Then
   Select Case VPerPago
          Case Is = "02"
               Sql$ = "select * from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "' and semana='" & VSemana & "' and Year(fechaproceso)=" & Vano & " and status<>'*' "
          Case Is = "04"
               Sql$ = "select * from plahistorico where cia='" & wcia & "' and proceso='" & VTipobol & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
             
   End Select
Else
   Sql$ = "select * from plaquincena where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and status<>'*' "
End If
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then Verifica_Boleta = False Else Verifica_Boleta = True
If Rs.State = 1 Then Rs.Close
End Function
Public Sub Elimina_Boleta()
If wtipodoc = True Then
   Mgrab = MsgBox("Seguro de Eliminar Boleta", vbYesNo + vbQuestion, TitMsg)
Else
   Sql$ = "select placod from plahistorico " _
   & "where cia='" & wcia & "' and proceso='01' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
   & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
   If (fAbrRst(Rs, Sql$)) Then
      MsgBox "Ya se genero la Boleta de Pago " & Chr(13) & "No se Puede Anular el Adelanto de Quincena", vbCritical, "Sistema de Planilla"
      Exit Sub
   End If
   Mgrab = MsgBox("Seguro de Eliminar Adelanto de Quincena", vbYesNo + vbQuestion, TitMsg)
End If
If Mgrab <> 6 Then Exit Sub

Sql$ = wInicioTrans
cn.Execute Sql$
If wtipodoc = True Then
   Select Case VPerPago
          Case Is = "02"
               Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
               & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
          Case Is = "04"
               Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
               & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
   End Select
Else
   Sql$ = "update plaquincena set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
   & "where cia='" & wcia & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
   & "and status<>'*' and placod='" & Txtcodpla.Text & "'"
End If
cn.Execute Sql$

If wtipodoc = True Then
   Sql$ = "select * from plabolcte " _
   & "where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
   & "and month( fechaproceso) = " & Vmes & " and status<>'*'"
   If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
   Do While Not Rs.EOF
      Sql = "update plactacte set fecha_cancela=null,pago_acuenta=pago_acuenta-" & Rs!importe & " where cia='" & wcia & "' and numinterno='" & Rs!numinterno & "' and tipo='" & Rs!tipo & "'"
      cn.Execute Sql$
      Rs.MoveNext
   Loop
   Sql$ = "update plabolcte set status='*'" _
   & "where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
   & "and month( fechaproceso) = " & Vmes & " and status<>'*'"
   cn.Execute Sql$
End If

Sql$ = wFinTrans
cn.Execute Sql$

End Sub

Public Function F01(concepto As String) As String 'INGRESOS
Dim rsF01 As ADODB.Recordset
Dim mFactor As Currency
Dim nHijos As Integer
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
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "03" 'ASIGNACION MOVILIDAD
            F01 = "select round((b.importe/factor_horas)*a.h01,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
        
       Case Is = "04" 'BONIFICACION T. SERVICIO
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           
       Case Is = "05" 'INCREMENTO AFP 10.23%
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
          
       Case Is = "06" 'INCREMENTO AFP 3%
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            
       Case Is = "08" 'SOBRETASA (CONSTRUCCION CIVIL)
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
          
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h13),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "09" 'DOMINICAL
            If vTipoTra <> "01" Then
               F01 = "select round(((b.importe/factor_horas)*h02),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "10" 'EXTRAS L-S
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h10),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "11" 'EXTRAS D-F
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h11),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "12" 'FERIADOS
            F01 = "select round((b.importe/factor_horas)*a.h03,2) as basico,A.H03 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(Txtcodpla.Text) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "13" 'REINTEGROS
       Case Is = "14" 'VACACIONES (CONSTRUCCION CIVIL)
            If VVacacion = "S" Then
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
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
               Sql$ = "select ultima from plasemanas where cia='" & wcia & "' and semana='" & Format(VSemana, "00") & "' and ano='" & Vano & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  If rsF01!ultima = "S" Then
                     Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
                     If (fAbrRst(rsF01, Sql$)) Then
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
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "21" 'BONIFICACION POR ALTURA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h18),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "22" 'BONIF. CONTACTO AGUA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h19),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "23" 'BONIF.POR ALTITUD
            If VAltitud = "S" Then
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round((" & mFactor & " / 8) * (a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "24" 'BONIF. TURNO NOCHE
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h20),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "25" 'H.E. HASTA DECIMA HORA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h21),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "26" 'H.E. HASTA ONCEAVA HORA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h22),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "27" 'EXTRAS 3PRA L-S
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h23),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "28" 'EXT. NOCHE 2PR L-S
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h24),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "29" 'EXT. NOCHE 3PRA L-S
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h25),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "30" 'SOBRETASA NOCHE(CONSTRUCCION CIVIL)
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
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
Dim mproy As Currency
Dim MUIT As Currency
Dim mgra As Integer
Dim msemano As Integer
Dim mpertope As Integer
Dim J As Integer

mFactor = 0
F02 = ""
mtope = 0

If (concepto <> "04" Or Trim(Lblcodafp.Caption) = "") And concepto <> "11" And concepto <> "13" Then 'SIN AFP
   If Not IsDate(VFechaJub) Then
    Sql$ = "select deduccion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and deduccion<>0 and status<>'*'"
  
    If (fAbrRst(rsF02, Sql$)) Then
       If Not IsNull(rsF02!deduccion) Then
          If rsF02!deduccion <> 0 Then mFactor = rsF02!deduccion
       End If
    End If
    
    If rsF02.State = 1 Then rsF02.Close
    
    If mFactor <> 0 Then
       Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
       If (fAbrRst(rsF02, Sql$)) Then
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
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"

    If (fAbrRst(rsF02, Sql$)) Then
       rsF02.MoveFirst
       F02str = ""
       Do While Not rsF02.EOF
          F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
          rsF02.MoveNext
       Loop
       F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
       
       If rsF02.State = 1 Then rsF02.Close
       mperiodoafp = Format(Vano, "0000") & Format(Vmes, "00")
       Sql$ = "select afp01,afp02,afp03,afp04,afp05,tope from  plaafp where periodo='" & mperiodoafp & "' and codafp='" & Lblcodafp.Caption & "' and status<>'*'"
    
       If Not (fAbrRst(rsF02, Sql$)) Then
          MsgBox "No se Encuentran Factores de Calculo para AFP", vbCritical, "Calculo de Boleta"
          mcancel = True
          Exit Function
       End If
       Sql$ = Acumula_Mes_Afp(concepto, "D")
       If (fAbrRst(rsF02afp, Sql$)) Then
          For J = 1 To 5
              vNombField = " as D11" & Format(J, "0")
              mFactor = rsF02(J - 1)
              If J = 2 Then
                 If manos > 64 Then mFactor = 0
                 Call Acumula_Mes_Afp112(concepto, "D")
                 mtope = macui
                 Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
                      & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
                      
                 If (fAbrRst(rsTope, Sql$)) Then mtope = mtope + rsTope!tope
                 If mtope > rsF02!tope Then mtope = rsF02!tope
                 If rsTope.State = 1 Then rsTope.Close
                 F02 = F02 & "round(((" & mtope & ") * " & mFactor & " /100)-" & macus & ",2) "
                 F02 = F02 & vNombField & ","
              Else
                 If Not IsNull(rsF02afp(0)) Then
                    F02 = F02 & "round(((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & ",2) "
                    F02 = F02 & vNombField & ","
                 Else
                    F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                    F02 = F02 & vNombField & ","
                 End If
              End If
          Next J
          If rsF02afp.State = 1 Then rsF02afp.Close
          If rsF02.State = 1 Then rsF02.Close
       End If
       F02 = Mid(F02, 1, Len(Trim(F02)) - 1)
       F02 = "select " & F02
       F02 = F02 & " from platemphist "
       F02 = F02 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
       F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
    End If
   End If
ElseIf concepto = "13" And VTipobol <> "03" Then 'Quinta Categoria
   Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF02, Sql$)) Then
      rsF02.MoveFirst
      F02str = ""
      Do While Not rsF02.EOF
         F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
         rsF02.MoveNext
      Loop
      F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      MUIT = 0
      mtope = 0
      mpertope = 0
      mgra = 0
      msemano = 0
      If vTipoTra <> "01" Then
         Sql$ = "select max(semana) as ultima from plasemanas where cia='" & wcia & "' and ano='" & Format(Vano, "0000") & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mpertope = rsF02(0): msemano = rsF02(0)
      Else
         If Vmes > 6 Then mpertope = 12 Else mpertope = 13
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      Call Acumula_Ano(concepto, "D")
      mtope = macui
      Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' " _
           & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      If (fAbrRst(rsF02, Sql$)) Then mtope = mtope + rsF02!tope
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select concepto,moneda,sum((importe/factor_horas)) as base from plaremunbase a,plaafectos b where a.cia='" & wcia & "' " _
           & "and placod='" & Txtcodpla.Text & "' and a.status<>'*' and b.tipo='D' and b.codigo='" & concepto & "' and b.tboleta='" & VTipobol & "' and b.status<>'*' and a.concepto=b.cod_remu " _
           & "Group By Placod,a.concepto,a.moneda"
      
       
      If (fAbrRst(rsF02, Sql$)) Then mproy = rsF02!base
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select uit from plauit where ano='" & Format(Vano, "0000") & "' and moneda='S/.' and status<>'*'"
      If (fAbrRst(rsF02, Sql$)) Then MUIT = rsF02!uit
      If rsF02.State = 1 Then rsF02.Close
      
      If vTipoTra = "05" Then
         Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='20' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mFactor = rsF02!factor
         If rsF02.State = 1 Then rsF02.Close
      
         Sql$ = "select concepto,moneda,importe/factor_horas as base from plaremunbase where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and concepto='01' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mproy = mproy + Round(rsF02!base * mFactor, 2)
         If rsF02.State = 1 Then rsF02.Close
      End If
      If vTipoTra = "01" Then
         If wtipodoc = True Then
            If Vmes < 12 Then mproy = (mproy * VHoras) * (mpertope - Vmes + 1) Else mproy = 0
         Else
            If Vmes = 12 Then
               mproy = (mproy / 2)
            Else
              mproy = ((mproy * VHoras) * (mpertope - Vmes + 1)) + ((mproy * VHoras) / 2)
            End If
         End If
      Else
         mgra = Busca_Grati()
         If vTipoTra = "05" Then
            Sql$ = "select importe/factor_horas as base,b.factor  from plaremunbase a,platasaanexo b where a.cia='" & wcia & "' and a.placod='" & Txtcodpla.Text & "'  and a.concepto='01' " _
                 & "and a.status<>'*' and b.cia='" & wcia & "' and b.tipomovimiento='01' and b.codinterno='15' and b.status<>'*' and b.tipotrab='" & vTipoTra & "' and b.cargo='" & Trim(Lblcargo.Caption) & "'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((rsF02!base * 8) * (rsF02!factor * mgra)) + ((rsF02!base * 8) * (mpertope - Val(VSemana)))
            If rsF02.State = 1 Then rsF02.Close
         Else
            Sql$ = "select importe/factor_horas as base from plaremunbase where Cia='" & wcia & "' and placod='" & Txtcodpla.Text & "'  and concepto='01' and status<>'*'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((mproy * 240) * mgra) + (rsF02!base * 8) * (mpertope - Val(VSemana))
            If rsF02.State = 1 Then rsF02.Close
            'mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((mproy * 240) * mgra) + (mproy * 8) * (mpertope - Val(VSemana))
         End If
      End If
      mtope = mtope + mproy
      If mtope > Round(MUIT * 7, 2) Then
         mtope = mtope - Round(MUIT * 7, 2)
         Select Case mtope
                Case Is < (Round(MUIT * 27, 2) + 1)
                     mFactor = Round(mtope * 0.15, 2)
                Case Is < (Round(MUIT * 54, 2) + 1)
                     mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
                Case Else
                     mFactor = Round(((mtope - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
         End Select
         If vTipoTra = "01" Then
            If wtipodoc = True Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
            Else
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
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

If Trim(Txtcodpla.Text) = "" Then Exit Sub

Sql$ = "select iniciomes from cia where cod_cia='" & _
wcia & "' and status<>'*'"

If (fAbrRst(Rs, Sql$)) Then
   If IsNull(Rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = Rs!iniciomes
End If
Rs.Close

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
 
Sql$ = wInicioTrans
cn.Execute Sql$

Sql$ = "delete from platemphist where cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
cn.Execute Sql

Sql$ = wFinTrans
cn.Execute Sql$
mdiasfalta = 0
mdiasferiado = 0
VHorasnormal = VHoras
If VTipobol = "01" Then mbol = "N"
If VTipobol = "02" Then mbol = "V"
If VTipobol = "03" Then mbol = "G"

wciamae = Funciones.Determina_Maestro("01077")

If VTipobol = "01" Or wtipodoc <> True Then

   Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
    
   If wtipodoc = True Then
     
      Select Case VPerPago
             Case Is = "04" 'Mensual
                  Sql$ = Sql$ & "select distinct(concepto) from platareo where fecha " _
                       & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + Coneccion.FormatTimei & _
                       "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
                       
             Case Is = "02" 'Semanal
                  Sql$ = Sql$ & "select distinct(concepto) from platareo where fecha " _
                       & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                       & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      End Select
   Else
      Sql$ = Sql$ & "select distinct(concepto) from platareo where fecha " _
           & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
           & "and codigotrab='" & Trim(Txtcodpla) & "' and status<>'*'"
      mbol = "N"
   End If
   
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      mconceptos = ""
      Do While Not Rs.EOF
         mconceptos = mconceptos & "'" & Trim(Rs!concepto) & "',"
         Rs.MoveNext
      Loop
      mconceptos = "(cod_maestro2 in(" & Mid(mconceptos, 1, Len(mconceptos) - 1) & ")"
   End If
   
   If Rs.State = 1 Then Rs.Close
   
   If Trim(mconceptos) <> "" Then
      Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0  and " & mconceptos
   Else
      Sql$ = "Select * from maestros_2 where status<>'*' and CHARINDEX('" & mbol & "',flag1 )>0 and "
   End If
   
   If vTipoTra = "01" Then
      If Trim(mconceptos) <> "" Then
         Sql$ = Sql$ & " or cod_maestro2 in('01'))"
      Else
         Sql$ = Sql$ & "cod_maestro2 in('01')"
      End If
   Else
     If Trim(mconceptos) <> "" Then
        Sql$ = Sql$ & " or cod_maestro2 in('01','02'))"
     Else
        Sql$ = Sql$ & "cod_maestro2 in('01','02')"
     End If
   End If
   
   Sql$ = Sql$ & wciamae
   
   If (fAbrRst(Rs, Sql$)) Then
      Rs.MoveFirst
      Do While Not Rs.EOF
         rshoras.AddNew
         rshoras!Codigo = Trim(Rs!cod_maestro2)
         rshoras!Descripcion = Rs!DESCRIP
         mhor = 0
         
         Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                  
         Select Case VPerPago
               Case Is = "04"
                    Sql$ = Sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where fecha " _
                         & "BETWEEN '" & Format(wBeginMonth, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & Trim(Rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
               Case Is = "02"
                    Sql$ = Sql$ & "select concepto,motivo,sum(tiempo) as tiempo from platareo where fecha " _
                         & " BETWEEN '" & Format(VfDel, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                         & "and codigotrab='" & Trim(Txtcodpla) & "' and concepto='" & Trim(Rs!cod_maestro2) & "' and status<>'*'  group by concepto,motivo order by concepto,motivo"
         End Select
 
         If (fAbrRst(rs2, Sql$)) Then
            rs2.MoveFirst
            Do While Not rs2.EOF
               Select Case rs2!motivo
                      Case Is = "DI"
                           mhor = mhor + (rs2!tiempo * 8 * 60)
                           mdiasfalta = mdiasfalta + rs2!tiempo
                           If Rs!cod_maestro2 = "03" Then mdiasferiado = mdiasferiado + rs2!tiempo
                      Case Is = "HO"
                           mhor = mhor + Int(rs2!tiempo) * 60 + ((rs2!tiempo - Int(rs2!tiempo)) * 100)
                      Case Is = "MI"
                           mhor = mhor + rs2!tiempo
               End Select
               rs2.MoveNext
            Loop
            If rs2.State = 1 Then rs2.Close
            mhor = Int(mhor / 60) + ((mhor Mod 60) / 100)
            rshoras!Monto = mhor
            If Rs!flag2 = "-" Then
               VHorasnormal = VHorasnormal - mhor
            End If
         End If
         Rs.MoveNext
      Loop
      
      'TIPICO
       
      rshoras.MoveFirst
      mHourTra = 0
      If rshoras!Codigo = "01" Then
         If wtipodoc = True Then
            mHourTra = Calc_Horas_FecIng(wBeginMonth)
            rshoras!Monto = VHorasnormal
         Else
            VHorasnormal = 0
            If VNewBoleta = True Then VHorasnormal = Calc_Horas_Quincena(wBeginMonth)
            rshoras!Monto = VHorasnormal
         End If
         If rshoras.RecordCount > 1 Then rshoras.MoveNext
      End If
      If rshoras!Codigo = "02" Then rshoras!Monto = ((6 - mdiasfalta + mdiasferiado) * 8) / 6
      If mHourTra <> 0 Then
         rshoras.AddNew
         rshoras!Codigo = "03"
         rshoras!Descripcion = "FERIADOS"
         rshoras!Monto = mHourTra
      End If

      
   End If
Else
    Sql$ = "Select * from maestros_2 where status<>'*' and  CHARINDEX('" & mbol & "',flag1 )>0"
    Sql$ = Sql$ & wciamae
    cn.CursorLocation = adUseClient
    Set Rs = New ADODB.Recordset
    Set Rs = cn.Execute(Sql$, 64)
    If rshoras.RecordCount > 0 Then
       rshoras.MoveFirst
       Do While Not rshoras.EOF
          rshoras.Delete
          rshoras.MoveNext
       Loop
    End If
    If Not Rs.RecordCount > 0 Then MsgBox "No Existen Horas Registradas", vbCritical, TitMsg: Exit Sub
    Rs.MoveFirst
    Do While Not Rs.EOF
       rshoras.AddNew
       rshoras!Codigo = Trim(Rs!cod_maestro2)
       rshoras!Descripcion = Trim(Rs!DESCRIP)
       If VTipobol = "02" Then rshoras!Monto = VHoras Else rshoras!Monto = "0.00"
       Rs.MoveNext
    Loop
    If Rs.State = 1 Then Rs.Close
End If
End Sub
Public Function F03(concepto As String) As String 'APORTACIONES
Dim rsF03 As ADODB.Recordset
Dim F03str As String
Dim mFactor As Currency
mFactor = 0
F03 = ""
If concepto = "03" Then
   Sql$ = "select senati from cia where cod_cia='" & wcia & "' and status<>'*'"
   If Not (fAbrRst(rsF03, Sql$)) Then Exit Function
   If rsF03!senati <> "S" Then Exit Function
   If rsF03.State = 1 Then rsF03.Close
   wciamae = Determina_Maestro("01044")
   Sql$ = "Select * from maestros_2 where cod_maestro2='" & Trim(VArea) & "' and status<>'*'"
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rsF03, Sql$)) Then
      If rsF03!flag7 <> "S" Then Exit Function
   Else
      Exit Function
   End If
   If rsF03.State = 1 Then rsF03.Close
End If

Sql$ = "select aportacion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and aportacion<>0 and status<>'*'"
If (fAbrRst(rsF03, Sql$)) Then
   If Not IsNull(rsF03!aportacion) Then
      If rsF03!aportacion <> 0 Then mFactor = rsF03!aportacion
   End If
End If
If rsF03.State = 1 Then rsF03.Close
If mFactor <> 0 Then
   Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF03, Sql$)) Then
      rsF03.MoveFirst
      F03str = ""
      Do While Not rsF03.EOF
         F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
         rsF03.MoveNext
      Loop
      F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
      If rsF03.State = 1 Then rsF03.Close
      Call Acumula_Mes(concepto, "A")
      F03 = "select round(((" & F03str & " + " & macui & ") * " & mFactor & " /100)-" & macus & ",2) as aportacion from platemphist "
      F03 = F03 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & Txtcodpla.Text & "' "
      F03 = F03 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
   End If
End If

macui = 0: macus = 0
End Function
Public Function V01(concepto As String) As String
'INGRESOS
Dim rsV01 As ADODB.Recordset
Dim mFactor As Currency
mFactor = 0
V01 = ""
Select Case concepto
       Case Is = "02" 'ASIGNACION FAMILIAR
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "04" 'BONIFICACION T. SERVICIO
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "05" 'INCREMENTO AFP 10.23%
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "06" 'INCREMENTO AFP 3%
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
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
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
            If (fAbrRst(rsV01, Sql$)) Then
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
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "04" 'BONIFICACION T. SERVICIO
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "05" 'INCREMENTO AFP 10.23%
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "06" 'INCREMENTO AFP 3%
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "07" 'BONIFICACION COSTO DE VIDA
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "15" 'GRATIFICACION
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h14),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Txtcodpla.Text & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
    End Select
Else
   Select Case concepto
          Case Is = "15" 'GRATIFICACION
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(Lbltope.Caption) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(Lblcargo.Caption) & "' and status<>'*'"
               If (fAbrRst(rsG01, Sql$)) Then
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
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 3
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
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
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Function Acumula_Mes_Afp(concepto As String, tipo As String) As String
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset

macui = 0: macus = 0
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
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
       SqlAcu = SqlAcu & "proceso in('01','02','03')"
    End If
If RsAcumula.State = 1 Then RsAcumula.Close
Acumula_Mes_Afp = SqlAcu
End Function
Private Sub Acumula_Mes_Afp112(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
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
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
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
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Sub Acumula_Ano(concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 3
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
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
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Function Busca_Grati() As Integer
Dim rsGrati As New Recordset
Select Case Vmes
       Case Is = 1, 2, 3, 4, 5, 6
            Busca_Grati = 2
       Case Is = 7
            Sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 1 Else Busca_Grati = 2
                 
       Case Is = 8, 9, 10, 11
            Busca_Grati = 1
       Case Is = 12
            Sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 0 Else Busca_Grati = 1
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
Sql$ = ""
If VTipobol = "01" Then
     
   Select Case VPerPago
          Case Is = "04" 'Mensual
               Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
               Sql$ = Sql$ & "select count(fecha) from plaferiados where cia='" & wcia & "' and fecha " _
                     & "BETWEEN '" & Format(mDateBegin, FormatFecha) + FormatTimei & "' AND '" & Format(VFProceso, FormatFecha) + FormatTimef & "' " _
                    & "and status<>'*'"
          Case Is = "02" 'Semanal
               Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
               Sql$ = Sql$ & "select count(fecha) from plaferiados where cia='" & wcia & "' and fecha " _
                     & "BETWEEN '" & Format(mDateBegin, FormatFecha) + FormatTimei & "' AND '" & Format(VfAl, FormatFecha) + FormatTimef & "' " _
                    & "and status<>'*'"
   End Select
   If Trim(Sql$) <> "" Then If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
   If Not IsNull(Rs(0)) Then
      If Rs(0) <> 0 Then
         VHorasnormal = VHorasnormal - (Rs(0) * mHourDay)
         Calc_Horas_FecIng = (Rs(0) * mHourDay)
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
mPer = 0
If rspagadic.RecordCount > 0 Then rspagadic.MoveFirst
Do While Not rspagadic.EOF
   If rspagadic!Codigo = "16" Then
      Sql = "select distinct(codinterno),factor from platasaanexo where cia='" & wcia & "' and tipomovimiento='" & vTipoTra & "' and modulo='01'  and basecalculo='16' and status<>'*'"
      If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst: mPer = Val(Rs(1))
      mCadOtPag = ""
      Do While Not Rs.EOF
         mCadOtPag = mCadOtPag & "I" & Rs(0) & "+"
         Rs.MoveNext
      Loop
      Rs.Close
      Exit Do
   End If
   rspagadic.MoveNext
Loop
Sql = ""
If Trim(mCadOtPag) <> "" Then
   mCadOtPag = Mid(mCadOtPag, 1, Len(Trim(mCadOtPag)) - 1)
   mCadOtPag = "sum(" & mCadOtPag & ")"
   Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " select " & mCadOtPag & " from plahistorico"
End If

'If Val(Mid(VFProceso, 4, 2)) > mPer Then mYearAnt = Mid(VFProceso, 7, 4) Else mYearAnt = Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
'If Val(Mid(VFProceso, 4, 2)) > mPer Then
'   mMonthAnt = Format(Val(Mid(VFProceso, 4, 2)) - mPer, "00")
'Else
'   mMonthAnt = Format((12 - mPer) + Val(Mid(VFProceso, 4, 2)), "00")
'   If Val(mMonthAnt) > 12 Then mMonthAnt = Format(Val(mMonthAnt) - 12, "00")
'End If

'mDateBeginVac = "01/" & mMonthAnt & "/" & mYearAnt

mDateBeginVac = Fecha_Promedios(mPer, VFProceso)

If Val(Mid(VFProceso, 4, 2)) = 1 Then
   mdia = Ultimo_Dia(12, Val(Mid(VFProceso, 7, 4)) - 1)
   mDateEndVac = Format(mdia, "00") & "/12/" & Format(Val(Mid(VFProceso, 7, 4)) - 1, "0000")
Else
   mdia = Ultimo_Dia(Val(Mid(VFProceso, 4, 2) - 1), Val(Mid(VFProceso, 7, 4)))
   mDateEndVac = Format(mdia, "00") & "/" & Format(Val(Mid(VFProceso, 4, 2) - 1), "00") & "/" & Mid(VFProceso, 7, 4)
End If

If Trim(Sql) <> "" Then
   Sql = Sql & " where cia='" & wcia & "' and fechaproceso Between '" & Format(mDateBeginVac, FormatFecha) & Space(1) & FormatTimei & "' and '" & Format(mDateEndVac, FormatFecha) & Space(1) & FormatTimef & "'"
   Sql = Sql & " and proceso='01' and placod='" & Txtcodpla.Text & "' and status<>'*'"
   If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
   If Not IsNull(Rs(0)) Then rspagadic!Monto = Rs(0) / mPer
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
Private Sub Descarga_ctaCte(Codigo As String, tipobol As String, fecha As String, sem As String, tiptrab As String, importe As Currency)
Dim Sql As String
Dim saldo As Currency
Sql = "select * from plactacte where cia='" & wcia & "' and placod='" & Codigo & "' and importe-pago_acuenta>0 and status<>'*' order by fecha"
If (fAbrRst(Rs, Sql$)) Then Rs.MoveFirst
Do While Not Rs.EOF
   saldo = (Rs!importe - Rs!pago_acuenta)
   If saldo >= importe Then
      If saldo = importe Then
         Sql = "update plactacte set pago_acuenta=pago_acuenta+" & importe & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and numinterno='" & Rs!numinterno & "' and tipo='" & Rs!tipo & "' and status<>'*'"
         cn.Execute Sql
      Else
         Sql = "update plactacte set pago_acuenta=pago_acuenta+" & importe & " where cia='" & wcia & "' and numinterno='" & Rs!numinterno & "' and tipo='" & Rs!tipo & "' and status<>'*'"
         cn.Execute Sql
      End If
      
      Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      Sql = Sql & "insert into plabolcte values('" & wcia & "','" & Codigo & "','" & tipobol & "','', " _
          & "'" & Format(fecha, FormatFecha) & "','" & sem & "','" & tiptrab & "','" & Rs!numinterno & "','" & Rs!tipo & "','" & Format(Rs!fecha, FormatFecha) & "', " _
          & "'" & wmoncont & "'," & importe & ",'','" & wuser & "'," & FechaSys & ")"
      cn.Execute Sql
      Exit Do
   Else
      Sql = "update plactacte set pago_acuenta=pago_acuenta+" & saldo & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and numinterno='" & Rs!numinterno & "' and tipo='" & Rs!tipo & "' and status<>'*'"
      cn.Execute Sql
      
      Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
      Sql = Sql & "insert into plabolcte values('" & wcia & "','" & Codigo & "','" & tipobol & "','', " _
          & "'" & Format(fecha, FormatFecha) & "','" & sem & "','" & tiptrab & "','" & Rs!numinterno & "','" & Rs!tipo & "','" & Format(Rs!fecha, FormatFecha) & "', " _
          & "'" & wmoncont & "'," & saldo & ",'','" & wuser & "'," & FechaSys & ")"
      cn.Execute Sql
      importe = importe - saldo
   End If
   Rs.MoveNext
Loop
Rs.Close
End Sub

Private Sub Calcula_Devengue_Vaca()
Dim mNumBol As Integer
Dim I As Integer
Dim rsdevengue As ADODB.Recordset

VokDevengue = True
Sql = "select placod,recordacu from plaprovvaca where cia='" & wcia & "' and year(fechaproceso)=" & Vano - 1 & " and month(fechaproceso)=12 and status<>'*'"
If (fAbrRst(rsdevengue, Sql)) Then rsdevengue.MoveFirst
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
          Txtcodpla.Text = rsdevengue!PlaCod
          Txtcodpla_LostFocus
          Grabar_Boleta
      Next
   End If
   rsdevengue.MoveNext
Loop
rsdevengue.Close
Panelprogress.Visible = False
If VokDevengue = False Then
   Sql = wInicioTrans
   cn.Execute Sql
   
   Sql = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
        & "where cia='" & wcia & "' and proceso='02' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
        & "and status='D'"
   cn.Execute Sql
   
   Sql = wFinTrans
   cn.Execute Sql
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

Sql$ = wInicioTrans
cn.Execute Sql$

Sql = "select * from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & Txtcodpla.Text & "' order by turno"
If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Sql = "update plahistorico set status='T' where cia='" & Rs!cia & "' and placod='" & Rs!PlaCod & "' and status='" & Rs!status & "' and turno='" & Rs!turno & "'"
cn.Execute Sql
Rs.Close
Sql$ = wFinTrans
cn.Execute Sql$
Limpia_Boleta
Screen.MousePointer = vbDefault

End Sub

Sub Llena_basico()
    Dim rtempo As New ADODB.Recordset
    
    Sql$ = "Select codinterno,descripcion from placonstante where " & _
    "tipomovimiento='02' and basico='S' and status<>'*' order by codinterno"
    rtempo.Open Sql, cn, adOpenStatic, adLockReadOnly

    'If Not RS.RecordCount > 0 Then Exit Sub
    'RS.MoveFirst
    'basico
    Do While Not rtempo.EOF
      rsbasico.AddNew
            rsbasico!Descripcion = Trim(rtempo!Descripcion)
            rsbasico!Codigo = rtempo!codinterno
            'rsbasico!moneda = wmoncont
            rsbasico!importe = "0.00"
      rtempo.MoveNext
    Loop
    rtempo.Close

End Sub

