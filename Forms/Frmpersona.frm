VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{92F3EE28-0DF4-4CB0-AAA5-E6578B12929B}#1.0#0"; "cbofacil.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmpersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Mantenimiento de Personal «"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   ForeColor       =   &H8000000D&
   Icon            =   "Frmpersona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameJud 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame19"
      Height          =   3135
      Left            =   0
      TabIndex        =   249
      Top             =   5760
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton CmdSalirJud 
         Caption         =   "Salir"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FF0000&
         TabIndex        =   250
         Top             =   2880
         Width           =   7695
      End
      Begin MSDataGridLib.DataGrid DgrdJud 
         Height          =   2295
         Left            =   120
         TabIndex        =   251
         Top             =   480
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         HeadLines       =   4
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "NroDni"
            Caption         =   "Dni"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """S/."" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Nombre"
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
            DataField       =   "TipoCta"
            Caption         =   "TipoCta"
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
            DataField       =   "Bco"
            Caption         =   "Bco"
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
            DataField       =   "NroCta"
            Caption         =   "NroCta"
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
            DataField       =   "Porcentaje"
            Caption         =   "Porcentaje"
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
         BeginProperty Column06 
            DataField       =   "CalcBaseBruto"
            Caption         =   "Calculo En BaseIngreso Total Bruto  (S=Si)"
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Descuentos Judicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   252
         Top             =   120
         Width           =   6600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   11655
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   120
         Width           =   6855
      End
      Begin VB.Label LblFecha 
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
         Left            =   9360
         TabIndex        =   70
         Top             =   165
         Width           =   2175
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
         Index           =   0
         Left            =   180
         TabIndex        =   68
         Top             =   225
         Width           =   825
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   53
      Top             =   720
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Personales"
      TabPicture(0)   =   "Frmpersona.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Framebuspla"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ChkDomiciliado"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdJud"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Datos Laborales"
      TabPicture(1)   =   "Frmpersona.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lblobra"
      Tab(1).Control(1)=   "Lbldesobra"
      Tab(1).Control(2)=   "lblPlaCos"
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(5)=   "Frame9"
      Tab(1).Control(6)=   "Frame12"
      Tab(1).Control(7)=   "Txtcodobra"
      Tab(1).Control(8)=   "Frame11"
      Tab(1).Control(9)=   "Frame21"
      Tab(1).Control(10)=   "frmcontrato"
      Tab(1).Control(11)=   "Frame22"
      Tab(1).Control(12)=   "Frame10"
      Tab(1).Control(13)=   "PnlPlanta"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "DerechoHabientes"
      TabPicture(2)   =   "Frmpersona.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame16"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Remuneraciones"
      TabPicture(3)   =   "Frmpersona.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame17"
      Tab(3).Control(1)=   "CmbPerRem"
      Tab(3).Control(2)=   "Frame19"
      Tab(3).Control(3)=   "Frame15"
      Tab(3).Control(4)=   "Frame14"
      Tab(3).Control(5)=   "Frame13"
      Tab(3).Control(6)=   "CmbMonBoleta"
      Tab(3).Control(7)=   "Label40(1)"
      Tab(3).Control(8)=   "Label40(0)"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Datos Complementarios"
      TabPicture(4)   =   "Frmpersona.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame30(3)"
      Tab(4).Control(1)=   "FraDobleImp(6)"
      Tab(4).Control(2)=   "FraAseguraTuPension"
      Tab(4).Control(3)=   "FraOtrosEmpleadores(4)"
      Tab(4).Control(4)=   "Frame30(2)"
      Tab(4).Control(5)=   "Frame30(0)"
      Tab(4).Control(6)=   "Frame29(0)"
      Tab(4).Control(7)=   "Frame30(6)"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Contratos"
      TabPicture(5)   =   "Frmpersona.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame30(4)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Suspension de 4ta categoria"
      TabPicture(6)   =   "Frmpersona.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame30(5)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Estab. donde labora el trab."
      TabPicture(7)   =   "Frmpersona.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame30(7)"
      Tab(7).ControlCount=   1
      Begin VB.CommandButton CmdJud 
         DisabledPicture =   "Frmpersona.frx":03EA
         Height          =   375
         Left            =   10800
         Picture         =   "Frmpersona.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   253
         Top             =   5880
         Width           =   255
      End
      Begin Threed.SSPanel PnlPlanta 
         Height          =   1455
         Left            =   -74960
         TabIndex        =   234
         Top             =   6480
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   2566
         _StockProps     =   15
         BackColor       =   14215660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox CmbPlanta 
            Height          =   315
            Left            =   510
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   120
            Width           =   2760
         End
         Begin VB.Label LblCodCantera 
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
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   480
            TabIndex        =   239
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label LblCantera 
            BackColor       =   &H00C0FFFF&
            Height          =   495
            Left            =   60
            TabIndex        =   238
            Top             =   840
            Width           =   3255
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   2880
            TabIndex        =   237
            Top             =   480
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Planta :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   236
            Top             =   120
            Width           =   555
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Datos Adicionales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1005
         Left            =   -66120
         TabIndex        =   208
         Top             =   4515
         Width           =   2295
         Begin MSMask.MaskEdBox txtAporteEmpleado 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   165
            TabIndex        =   210
            Top             =   570
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   556
            _Version        =   393216
            ForeColor       =   0
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "% Aporte Empleador"
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   135
            TabIndex        =   209
            Top             =   285
            Width           =   1515
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Centro de Costo"
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
         Height          =   1695
         Left            =   -71640
         TabIndex        =   201
         Top             =   6360
         Width           =   7815
         Begin VB.ListBox LstCcosto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   810
            Left            =   435
            TabIndex        =   202
            Top             =   675
            Visible         =   0   'False
            Width           =   5895
         End
         Begin MSDataGridLib.DataGrid Dgrdccosto 
            Height          =   1335
            Left            =   120
            TabIndex        =   203
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   2355
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
            BackColor       =   12648447
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
                  ColumnWidth     =   5880.189
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
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   3150
            TabIndex        =   204
            Top             =   1740
            Width           =   855
         End
      End
      Begin VB.CheckBox ChkDomiciliado 
         Alignment       =   1  'Right Justify
         Caption         =   "Domiciliado"
         Height          =   195
         Left            =   1740
         TabIndex        =   14
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Direccion Legal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   105
         TabIndex        =   73
         Top             =   2295
         Width           =   11085
         Begin VB.TextBox TXtNroEtapa1 
            Height          =   285
            Left            =   5280
            TabIndex        =   31
            Top             =   1665
            Width           =   700
         End
         Begin VB.TextBox TxtNroBlock1 
            Height          =   285
            Left            =   7320
            MaxLength       =   4
            TabIndex        =   26
            Top             =   1080
            Width           =   700
         End
         Begin VB.TextBox TxtNroKM1 
            Height          =   285
            Left            =   5880
            MaxLength       =   4
            TabIndex        =   25
            Top             =   1080
            Width           =   700
         End
         Begin VB.TextBox TxtNroLote1 
            Height          =   285
            Left            =   4200
            MaxLength       =   4
            TabIndex        =   24
            Top             =   1080
            Width           =   700
         End
         Begin VB.TextBox TxtNroMz1 
            Height          =   285
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   23
            Top             =   1080
            Width           =   700
         End
         Begin VB.TextBox TxtNroDpto1 
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   22
            Top             =   1080
            Width           =   700
         End
         Begin VB.TextBox TxtRefAnt 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   6270
            MaxLength       =   30
            TabIndex        =   38
            Top             =   2445
            Width           =   4680
         End
         Begin CboFacil.cbo_facil cbo_viatrab 
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   525
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   0   'False
            Enabled         =   -1  'True
            Ninguno         =   0   'False
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CboFacil.cbo_facil cbo_zonatrab 
            Height          =   315
            Left            =   90
            TabIndex        =   28
            Top             =   1665
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   0   'False
            Enabled         =   -1  'True
            Ninguno         =   0   'False
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox TxtRef 
            Height          =   330
            Left            =   90
            MaxLength       =   30
            TabIndex        =   37
            Top             =   2445
            Width           =   6120
         End
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   330
            Left            =   6270
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1665
            Width           =   4365
         End
         Begin VB.TextBox txtint 
            Height          =   285
            Left            =   7320
            MaxLength       =   10
            TabIndex        =   18
            Top             =   525
            Width           =   735
         End
         Begin VB.TextBox txtnro 
            Height          =   285
            Left            =   6270
            MaxLength       =   10
            TabIndex        =   17
            Top             =   525
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Height          =   330
            Left            =   2070
            MaxLength       =   30
            TabIndex        =   16
            Top             =   525
            Width           =   4140
         End
         Begin VB.TextBox Text9 
            Height          =   330
            Left            =   2070
            MaxLength       =   30
            TabIndex        =   30
            Top             =   1665
            Width           =   2940
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   345
            Left            =   10635
            TabIndex        =   21
            ToolTipText     =   "Enviar Correo Electronico"
            Top             =   480
            Width           =   345
            _Version        =   65536
            _ExtentX        =   609
            _ExtentY        =   609
            _StockProps     =   78
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            Picture         =   "Frmpersona.frx":0EFE
         End
         Begin VB.TextBox TxtEmail 
            Height          =   285
            Left            =   8160
            TabIndex        =   20
            Top             =   525
            Width           =   2385
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manzana:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   2175
            TabIndex        =   199
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Block:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   6855
            TabIndex        =   198
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilometro:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   5040
            TabIndex        =   197
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   3765
            TabIndex        =   196
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   195
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Etapa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   5380
            TabIndex        =   194
            Top             =   1440
            Width           =   420
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia (Sistema Anterior)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   6270
            TabIndex        =   129
            Top             =   2160
            Width           =   2130
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Referencia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   124
            Top             =   2160
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"Frmpersona.frx":0F1A
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   123
            Top             =   1440
            Width           =   6690
         End
         Begin VB.Label lblubigeo 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   10680
            TabIndex        =   34
            Top             =   1665
            Width           =   330
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"Frmpersona.frx":0FA2
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   8550
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Situación Especial del Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   3
         Left            =   -67440
         TabIndex        =   154
         Top             =   720
         Width           =   3615
         Begin VB.OptionButton OptSitEsp_ninguna 
            Caption         =   "Ninguna"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   157
            Top             =   720
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptSitEsp_confianza 
            Caption         =   "Trabajador de confianza"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   156
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton OptSitEsp_direccion 
            Caption         =   "Trabajador de dirección"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   155
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame FraDobleImp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   6
         Left            =   -67440
         TabIndex        =   183
         Top             =   1800
         Width           =   3615
         Begin VB.ComboBox CmbEvitaDobleImp 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   184
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label Label15 
            Caption         =   "Aplica convenio para evitar doble imposición"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   185
            Top             =   225
            Width           =   3375
         End
      End
      Begin VB.Frame FraAseguraTuPension 
         Caption         =   "¿Afiliación a Asegura tu Pensión?"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -67440
         TabIndex        =   186
         Top             =   2760
         Width           =   3615
         Begin VB.OptionButton OptAsegPenNo 
            Caption         =   "No"
            Height          =   255
            Left            =   1320
            TabIndex        =   188
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptAsegPenSi 
            Caption         =   "Si"
            Height          =   255
            Left            =   240
            TabIndex        =   187
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1725
         Left            =   120
         TabIndex        =   72
         Top             =   5400
         Width           =   11070
         Begin VB.Frame Frame6 
            Caption         =   "Deducciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            Left            =   6825
            TabIndex        =   191
            Top             =   180
            Width           =   4230
            Begin VB.CheckBox ChkRetJudUti 
               Caption         =   "Retencion Judicial en Boleta de Utilidades"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   120
               TabIndex        =   220
               Top             =   1080
               Width           =   3975
            End
            Begin MSDataGridLib.DataGrid DgrDeduccion 
               Height          =   825
               Left            =   90
               TabIndex        =   192
               Top             =   195
               Width           =   4050
               _ExtentX        =   7144
               _ExtentY        =   1455
               _Version        =   393216
               AllowUpdate     =   -1  'True
               Appearance      =   0
               ForeColor       =   0
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
               BeginProperty Column02 
                  DataField       =   "porcentaje"
                  Caption         =   "Fijo %"
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
                  DataField       =   "codigo"
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
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   1995.024
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     ColumnWidth     =   524.976
                  EndProperty
                  BeginProperty Column03 
                     Object.Visible         =   0   'False
                  EndProperty
               EndProperty
            End
         End
         Begin VB.CheckBox ChkCalcula_AccidenteTrabajo 
            Caption         =   "Calcular Accidente de Trabajo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   3525
            TabIndex        =   182
            Top             =   615
            Width           =   3300
         End
         Begin VB.CheckBox ChkMadreResp 
            Caption         =   "Indicador de madre con responsabilidad Familiar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   240
            TabIndex        =   165
            Top             =   960
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.CheckBox ChkDiscapacidad 
            Caption         =   "Trabajador Discapacitado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   130
            Top             =   720
            Width           =   2580
         End
         Begin CboFacil.cbo_facil cbo_cattrab 
            Height          =   315
            Left            =   1080
            TabIndex        =   19
            Top             =   225
            Width           =   5640
            _ExtentX        =   9948
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cat. Trab :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   125
            Top             =   255
            Width           =   795
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Establecimientos donde labora el trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Index           =   7
         Left            =   -74880
         TabIndex        =   180
         Top             =   840
         Width           =   11055
         Begin TrueOleDBGrid70.TDBGrid DgrdrTipEst 
            Height          =   4695
            Left            =   120
            TabIndex        =   181
            Top             =   360
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   8281
            _LayoutType     =   4
            _RowHeight      =   17
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   4
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Add"
            Columns(0).DataField=   "add"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "codest"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Establecimiento"
            Columns(2).DataField=   "nomest"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo"
            Columns(3).DataField=   "tipest"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Ruc"
            Columns(4).DataField=   "ruc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Razón Social"
            Columns(5).DataField=   "razsoc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1111"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1032"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1244"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1164"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8196"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=4128"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4048"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8192"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=4233"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4154"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=8192"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(3).DropDownList=1"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2778"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2699"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=8708"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=4948"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=4868"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=8196"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   3
            FootLines       =   1
            Caption         =   "Tipos de establecimientos"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   13160660
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H800000&"
            _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HEFEFEF&,.appearance=0"
            _StyleDefs(14)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFECD9&"
            _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFEBD7&"
            _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HEFEFEF&"
            _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFFFFFF&"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0,.locked=-1"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15,.alignment=3"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=0,.locked=-1"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.locked=-1"
            _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(63)  =   "Named:id=33:Normal"
            _StyleDefs(64)  =   ":id=33,.parent=0"
            _StyleDefs(65)  =   "Named:id=34:Heading"
            _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   ":id=34,.wraptext=-1"
            _StyleDefs(68)  =   "Named:id=35:Footing"
            _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(70)  =   "Named:id=36:Selected"
            _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(72)  =   "Named:id=37:Caption"
            _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(74)  =   "Named:id=38:HighlightRow"
            _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(76)  =   "Named:id=39:EvenRow"
            _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(78)  =   "Named:id=40:OddRow"
            _StyleDefs(79)  =   ":id=40,.parent=33"
            _StyleDefs(80)  =   "Named:id=41:RecordSelector"
            _StyleDefs(81)  =   ":id=41,.parent=34"
            _StyleDefs(82)  =   "Named:id=42:FilterBar"
            _StyleDefs(83)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame Frame30 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5400
         Index           =   5
         Left            =   -74760
         TabIndex        =   177
         Top             =   720
         Width           =   10815
         Begin TrueOleDBGrid70.TDBGrid DgrdrsSuspension4ta 
            Height          =   5010
            Left            =   165
            TabIndex        =   178
            Top             =   225
            Width           =   10470
            _ExtentX        =   18468
            _ExtentY        =   8837
            _LayoutType     =   4
            _RowHeight      =   17
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nro de Operacion de la solicitud de suspensión"
            Columns(0).DataField=   "numop"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fecha de Presentacion"
            Columns(1).DataField=   "fecha"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Ejercicio(año)"
            Columns(2).DataField=   "ejercicio"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   49
            Columns(3)._MaxComboItems=   5
            Columns(3).ValueItems(0)._DefaultItem=   0
            Columns(3).ValueItems(0).Value=   "Internet"
            Columns(3).ValueItems(0).Value.vt=   8
            Columns(3).ValueItems(0).DisplayValue=   "Internet"
            Columns(3).ValueItems(0).DisplayValue.vt=   8
            Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems(1)._DefaultItem=   0
            Columns(3).ValueItems(1).Value=   "Dependencia SUNAT"
            Columns(3).ValueItems(1).Value.vt=   8
            Columns(3).ValueItems(1).DisplayValue=   "Dependencia SUNAT"
            Columns(3).ValueItems(1).DisplayValue.vt=   8
            Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(3).ValueItems.Count=   2
            Columns(3).Caption=   "Medio de Presentación"
            Columns(3).DataField=   "medio"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2778"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2699"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2514"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2434"
            Splits(0)._ColumnProps(9)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=1376"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1296"
            Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(15)=   "Column(3).Width=10874"
            Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=10795"
            Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=0"
            Splits(0)._ColumnProps(19)=   "Column(3).Button=1"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(3).DropDownList=1"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
            AllowAddNew     =   -1  'True
            Appearance      =   2
            DefColWidth     =   0
            HeadLines       =   3
            FootLines       =   1
            Caption         =   "Datos Suspension de 4ta categoria"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   13160660
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H800000&"
            _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HEFEFEF&,.appearance=0"
            _StyleDefs(14)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFECD9&"
            _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFEBD7&"
            _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HEFEFEF&"
            _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFFFFFF&"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2,.locked=0"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15,.alignment=3"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.wraptext=-1,.locked=0"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=0"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=0,.locked=0"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(55)  =   "Named:id=33:Normal"
            _StyleDefs(56)  =   ":id=33,.parent=0"
            _StyleDefs(57)  =   "Named:id=34:Heading"
            _StyleDefs(58)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   ":id=34,.wraptext=-1"
            _StyleDefs(60)  =   "Named:id=35:Footing"
            _StyleDefs(61)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(62)  =   "Named:id=36:Selected"
            _StyleDefs(63)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(64)  =   "Named:id=37:Caption"
            _StyleDefs(65)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(66)  =   "Named:id=38:HighlightRow"
            _StyleDefs(67)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(68)  =   "Named:id=39:EvenRow"
            _StyleDefs(69)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(70)  =   "Named:id=40:OddRow"
            _StyleDefs(71)  =   ":id=40,.parent=33"
            _StyleDefs(72)  =   "Named:id=41:RecordSelector"
            _StyleDefs(73)  =   ":id=41,.parent=34"
            _StyleDefs(74)  =   "Named:id=42:FilterBar"
            _StyleDefs(75)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame Frame30 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5295
         Index           =   4
         Left            =   -74880
         TabIndex        =   166
         Top             =   720
         Width           =   11055
         Begin VB.CommandButton CmdGenContrato 
            Caption         =   "Generar Contrato"
            Height          =   375
            Index           =   1
            Left            =   4680
            TabIndex        =   176
            Top             =   4800
            Width           =   2055
         End
         Begin VB.CommandButton CmdGrabarContrato 
            Caption         =   "Grabar Contrato"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   169
            Top             =   4800
            Width           =   2055
         End
         Begin VB.CommandButton CmdNewContrato 
            Caption         =   "Nuevo Contrato"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   168
            Top             =   4800
            Width           =   2055
         End
         Begin TrueOleDBGrid70.TDBGrid DgrdContrato 
            Height          =   4455
            Left            =   120
            TabIndex        =   167
            Top             =   240
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   7858
            _LayoutType     =   4
            _RowHeight      =   17
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nro Contrato"
            Columns(0).DataField=   "nro_contrato"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fecha Inicio o reinicio de Actividad"
            Columns(1).DataField=   "fecini"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fecha Fin"
            Columns(2).DataField=   "fecfin"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Importe Sueldo"
            Columns(3).DataField=   "imp_sueldo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tipo de Contrato"
            Columns(4).DataField=   "contrato"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Reporte Contrato"
            Columns(5).DataField=   "reporte"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Fecha firma contrato"
            Columns(6).DataField=   "FechaFirmaContrato"
            Columns(6).NumberFormat=   "dd/mm/yyyy"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2778"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2699"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8193"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(9)=   "Column(1).WrapText=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2196"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2117"
            Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(15)=   "Column(3).Width=1984"
            Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1905"
            Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8194"
            Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(20)=   "Column(4).Width=3413"
            Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=3334"
            Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
            Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(25)=   "Column(5).Width=10134"
            Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=10054"
            Splits(0)._ColumnProps(28)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(30)=   "Column(5).ButtonText=1"
            Splits(0)._ColumnProps(31)=   "Column(5).ButtonAlways=1"
            Splits(0)._ColumnProps(32)=   "Column(6).Width=3784"
            Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=3704"
            Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=8196"
            Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   2
            FootLines       =   1
            Caption         =   "Relación de contratos del trabajador"
            MultipleLines   =   2
            CellTipsWidth   =   0
            DeadAreaBackColor=   13160660
            RowDividerColor =   13160660
            RowSubDividerColor=   13160660
            DirectionAfterEnter=   1
            MaxRows         =   250000
            ViewColumnCaptionWidth=   0
            ViewColumnWidth =   0
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H800000&"
            _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
            _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HEFEFEF&,.appearance=0"
            _StyleDefs(14)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFECD9&"
            _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFEBD7&"
            _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HEFEFEF&"
            _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFFFFFF&"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2,.locked=-1"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15,.alignment=3"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.wraptext=-1,.locked=0"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=0"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1,.locked=-1"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(67)  =   "Named:id=33:Normal"
            _StyleDefs(68)  =   ":id=33,.parent=0"
            _StyleDefs(69)  =   "Named:id=34:Heading"
            _StyleDefs(70)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   ":id=34,.wraptext=-1"
            _StyleDefs(72)  =   "Named:id=35:Footing"
            _StyleDefs(73)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(74)  =   "Named:id=36:Selected"
            _StyleDefs(75)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(76)  =   "Named:id=37:Caption"
            _StyleDefs(77)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(78)  =   "Named:id=38:HighlightRow"
            _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(80)  =   "Named:id=39:EvenRow"
            _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(82)  =   "Named:id=40:OddRow"
            _StyleDefs(83)  =   ":id=40,.parent=33"
            _StyleDefs(84)  =   "Named:id=41:RecordSelector"
            _StyleDefs(85)  =   ":id=41,.parent=34"
            _StyleDefs(86)  =   "Named:id=42:FilterBar"
            _StyleDefs(87)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.ComboBox CmbPerRem 
         Height          =   315
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   164
         Top             =   960
         Width           =   2175
      End
      Begin VB.Frame FraOtrosEmpleadores 
         BackColor       =   &H8000000B&
         Caption         =   "Otros Empleadores"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Index           =   4
         Left            =   -69240
         TabIndex        =   158
         Top             =   4200
         Width           =   5415
         Begin MSDataGridLib.DataGrid DgrdOtrosEmpleadores 
            Height          =   2175
            Left            =   120
            TabIndex        =   159
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3836
            _Version        =   393216
            AllowUpdate     =   -1  'True
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
               DataField       =   "ruc"
               Caption         =   "RUC"
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
               DataField       =   "razsoc"
               Caption         =   "RAZON SOCIAL"
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
                  ColumnWidth     =   3284.788
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Trabajador sujeto a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Index           =   2
         Left            =   -74760
         TabIndex        =   148
         Top             =   4200
         Width           =   5415
         Begin VB.CheckBox Chk5taExonerada 
            Caption         =   "Indicador de rentas de quinta categoría exoneradas o inafectas "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   153
            Top             =   1920
            Width           =   3840
         End
         Begin VB.CheckBox ChkOtrosIngta 
            Caption         =   "Tiene otros ingresos de Quinta Categoría"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   152
            Top             =   1560
            Width           =   3840
         End
         Begin VB.CheckBox ChkHorario_nocturno 
            Caption         =   "Trabajador sujeto a horario nocturno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   151
            Top             =   1200
            Width           =   4560
         End
         Begin VB.CheckBox ChkJornada_max 
            Caption         =   "Trabajador sujeto a jornada de trabajo máxima"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   150
            Top             =   840
            Width           =   4680
         End
         Begin VB.CheckBox ChkRegimen_Alternativo 
            Caption         =   "Trabajador sujeto a régimen alternativo, acumulativo o atípico de jornada de trabajo y descanso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   255
            TabIndex        =   149
            Top             =   240
            Width           =   5040
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Régimen Pensionario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   0
         Left            =   -74760
         TabIndex        =   144
         Top             =   2280
         Width           =   7215
         Begin VB.Frame Frame18 
            Caption         =   "Tipo Comisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   735
            Left            =   120
            TabIndex        =   211
            Top             =   960
            Width           =   2655
            Begin VB.OptionButton optAFPFlujo 
               Caption         =   "Flujo"
               Height          =   255
               Left            =   1440
               TabIndex        =   213
               Top             =   360
               Width           =   975
            End
            Begin VB.OptionButton optAFPMixta 
               Caption         =   "Mixta"
               Height          =   255
               Left            =   240
               TabIndex        =   212
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.TextBox TxtNroAfp 
            Height          =   285
            Left            =   3480
            TabIndex        =   140
            Top             =   600
            Width           =   2295
         End
         Begin CboFacil.cbo_facil cbo_pensiones 
            Height          =   315
            Left            =   1440
            TabIndex        =   138
            Top             =   240
            Width           =   5610
            _ExtentX        =   9895
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox TxtFecAfilia 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   139
            Top             =   600
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reg. Pens :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   147
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fec. Afiliacion"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   146
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CUSPP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   145
            Top             =   600
            Width           =   480
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "Prestaciones de Salud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Index           =   0
         Left            =   -74760
         TabIndex        =   131
         Top             =   720
         Width           =   7215
         Begin VB.TextBox TxtIpss 
            Height          =   285
            Left            =   5640
            TabIndex        =   135
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton OptEpsSi 
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   133
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton OptEpsNo 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   132
            Top             =   240
            Width           =   615
         End
         Begin CboFacil.cbo_facil cbo_eps 
            Height          =   315
            Left            =   1440
            TabIndex        =   134
            Top             =   600
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CboFacil.cbo_facil cbo_situacioneps 
            Height          =   315
            Left            =   1440
            TabIndex        =   137
            Top             =   960
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EPS :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   143
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carnet IPSS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   4665
            TabIndex        =   142
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "¿Afiliado a EPS/Servicios propios?"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   141
            Top             =   240
            Width           =   2385
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Situacion EPS :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   136
            Top             =   960
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1620
         Left            =   105
         TabIndex        =   71
         Top             =   645
         Width           =   11070
         Begin VB.ComboBox CmbCivil 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3375
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   765
            Width           =   2655
         End
         Begin VB.TextBox Text8 
            Height          =   330
            Left            =   4410
            TabIndex        =   9
            Top             =   1170
            Width           =   1590
         End
         Begin VB.ComboBox Cbotipodoc 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1170
            Width           =   2520
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   330
            Index           =   0
            Left            =   9720
            TabIndex        =   13
            ToolTipText     =   "Telefonos"
            Top             =   1200
            Width           =   330
            _Version        =   65536
            _ExtentX        =   582
            _ExtentY        =   582
            _StockProps     =   78
            BevelWidth      =   1
            AutoSize        =   2
            Picture         =   "Frmpersona.frx":1056
         End
         Begin VB.Frame Frame5 
            Caption         =   " Sexo "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   6240
            TabIndex        =   10
            Top             =   1080
            Width           =   3240
            Begin VB.OptionButton OpcDama 
               Caption         =   "Femenino"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1800
               TabIndex        =   12
               Top             =   195
               Width           =   1125
            End
            Begin VB.OptionButton OpcVaron 
               Caption         =   "Masculino"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   435
               TabIndex        =   11
               Top             =   195
               Width           =   1185
            End
         End
         Begin VB.TextBox TxtSegNom 
            Height          =   285
            Left            =   8760
            MaxLength       =   20
            TabIndex        =   4
            Top             =   405
            Width           =   2175
         End
         Begin VB.TextBox TxtPriNom 
            Height          =   285
            Left            =   6600
            MaxLength       =   20
            TabIndex        =   3
            Top             =   405
            Width           =   2175
         End
         Begin VB.TextBox TxtApeCas 
            Height          =   285
            Left            =   4440
            MaxLength       =   20
            TabIndex        =   2
            Top             =   405
            Width           =   2175
         End
         Begin VB.TextBox TxtApeMat 
            Height          =   285
            Left            =   2280
            MaxLength       =   20
            TabIndex        =   1
            Top             =   405
            Width           =   2175
         End
         Begin VB.TextBox TxtApePat 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   0
            Top             =   405
            Width           =   2175
         End
         Begin MSMask.MaskEdBox TxtFecNac 
            Height          =   315
            Left            =   975
            TabIndex        =   5
            Top             =   765
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin CboFacil.cbo_facil cbo_naciontrab 
            Height          =   315
            Left            =   7470
            TabIndex        =   7
            Top             =   765
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCommand SSCommand6 
            Height          =   360
            Index           =   6
            Left            =   10200
            TabIndex        =   247
            ToolTipText     =   "Telefonos"
            Top             =   1200
            Width           =   360
            _Version        =   65536
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   78
            BevelWidth      =   1
            AutoSize        =   1
            Picture         =   "Frmpersona.frx":15F0
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   $"Frmpersona.frx":26802
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   193
            Top             =   120
            Width           =   9930
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fec. Nac.:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   162
            Top             =   825
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civil"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2370
            TabIndex        =   161
            Top             =   825
            Width           =   825
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nacionalidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6300
            TabIndex        =   160
            Top             =   825
            Width           =   900
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Doc :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3690
            TabIndex        =   121
            Top             =   1215
            Width           =   600
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Doc :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   120
            Top             =   1200
            Width           =   720
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   " SCTR - ESSALUD "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   -74865
         TabIndex        =   119
         Top             =   5040
         Width           =   3135
         Begin CboFacil.cbo_facil cboessalud 
            Height          =   315
            Left            =   180
            TabIndex        =   46
            Top             =   225
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   0   'False
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame frmcontrato 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   -69120
         TabIndex        =   117
         Top             =   4320
         Width           =   5205
         Begin VB.CheckBox chkJubilado 
            Caption         =   "Jubilado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   215
            Top             =   1560
            Width           =   1095
         End
         Begin CboFacil.cbo_facil cbo_tipocont 
            Height          =   315
            Left            =   780
            TabIndex        =   52
            Top             =   180
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox TxtFecCese 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   960
            TabIndex        =   170
            Top             =   600
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Txtfecjubila 
            Height          =   315
            Left            =   2400
            TabIndex        =   172
            Top             =   1560
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin CboFacil.cbo_facil cbo_TipMotFinPer 
            Height          =   315
            Left            =   120
            TabIndex        =   175
            Top             =   1200
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSCommand BtnMemo 
            Height          =   735
            Left            =   2160
            TabIndex        =   205
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   1296
            _StockProps     =   78
            Caption         =   "MEMORAND"
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox TxtFec_Afil_Sindicato 
            Height          =   315
            Left            =   4080
            TabIndex        =   254
            Top             =   720
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label26 
            Caption         =   "Fecha Afiliacion Sindicato"
            Height          =   615
            Left            =   3360
            TabIndex        =   255
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo fin de periodo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   174
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Jub. Anti."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   173
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Cese"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   171
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contrato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   45
            TabIndex        =   118
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   " Intermediacion Laboral "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -71640
         TabIndex        =   115
         Top             =   4365
         Width           =   2535
         Begin VB.TextBox Text7 
            Height          =   330
            Left            =   990
            MaxLength       =   15
            TabIndex        =   48
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° RUC :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   116
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   " ESSALUD VIDA "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -74865
         TabIndex        =   43
         Top             =   4365
         Width           =   3135
         Begin VB.OptionButton OpcVidaSi 
            Caption         =   "Afiliado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   44
            Top             =   285
            Width           =   1215
         End
         Begin VB.OptionButton OpcVidaNo 
            Caption         =   "No Afiliado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1635
            TabIndex        =   45
            Top             =   285
            Width           =   1230
         End
      End
      Begin VB.Frame Frame19 
         Height          =   1185
         Left            =   -68550
         TabIndex        =   113
         Top             =   5520
         Width           =   4800
         Begin VB.CheckBox chk_integro 
            Caption         =   "Asignación Fam. No Prorrateada"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   200
            Top             =   840
            Visible         =   0   'False
            Width           =   4500
         End
         Begin VB.CheckBox ChkVaca 
            Caption         =   "Prorrateo de Vacaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   90
            TabIndex        =   63
            Top             =   495
            Width           =   2220
         End
         Begin VB.CheckBox ChkAltitud 
            Caption         =   "Bonificacion por Altitud"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   105
            TabIndex        =   62
            Top             =   195
            Width           =   2100
         End
      End
      Begin VB.TextBox Txtcodobra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74280
         MaxLength       =   8
         TabIndex        =   111
         Top             =   6870
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame15 
         Caption         =   "Datos CTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3180
         Left            =   -66120
         TabIndex        =   88
         Top             =   1290
         Width           =   2310
         Begin VB.TextBox TxtCCICts 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   218
            Top             =   2760
            Width           =   2145
         End
         Begin VB.TextBox TxtCtaCts 
            Height          =   285
            Left            =   75
            MaxLength       =   30
            TabIndex        =   67
            Top             =   2205
            Width           =   2100
         End
         Begin VB.ComboBox CmbCtaCts 
            Height          =   315
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1635
            Width           =   2175
         End
         Begin VB.ComboBox CmbMonCts 
            Height          =   315
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   1035
            Width           =   2175
         End
         Begin VB.ComboBox CmbBcoCts 
            Height          =   315
            Left            =   75
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   420
            Width           =   2175
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CCI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   219
            Top             =   2520
            Width           =   270
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No de Cuenta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   90
            TabIndex        =   97
            Top             =   2010
            Width           =   990
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cuenta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   75
            TabIndex        =   96
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   105
            TabIndex        =   95
            Top             =   855
            Width           =   570
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   75
            TabIndex        =   94
            Top             =   225
            Width           =   435
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4230
         Left            =   -68520
         TabIndex        =   87
         Top             =   1290
         Width           =   2370
         Begin VB.TextBox TxtCCI 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   216
            Top             =   3860
            Width           =   2145
         End
         Begin VB.TextBox Txt_Sucursal 
            Height          =   285
            Left            =   135
            MaxLength       =   3
            TabIndex        =   190
            Top             =   2835
            Width           =   750
         End
         Begin VB.TextBox TxtCtaPag 
            Height          =   285
            Left            =   135
            MaxLength       =   30
            TabIndex        =   61
            Top             =   3330
            Width           =   2145
         End
         Begin VB.ComboBox CmbCtaPag 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   2235
            Width           =   2175
         End
         Begin VB.ComboBox CmbMonPago 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1650
            Width           =   2175
         End
         Begin VB.ComboBox CmbBcoPago 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   1050
            Width           =   2175
         End
         Begin VB.ComboBox CmbTipoPago 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   450
            Width           =   2175
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CCI"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   217
            Top             =   3620
            Width           =   270
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   189
            Top             =   2610
            Width           =   660
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No de Cuenta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   93
            Top             =   3105
            Width           =   990
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de cuenta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   2040
            Width           =   1065
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   1455
            Width           =   570
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   855
            Width           =   435
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Pago"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   255
            Width           =   930
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Remuneracion Basica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5445
         Left            =   -74925
         TabIndex        =   86
         Top             =   1290
         Width           =   6315
         Begin VB.ListBox LstMoneda 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   3840
            TabIndex        =   108
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ListBox LstTipo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   2760
            TabIndex        =   107
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSDataGridLib.DataGrid DgrBasico 
            Height          =   5025
            Left            =   75
            TabIndex        =   56
            Top             =   240
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   8864
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
            ColumnCount     =   7
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
            BeginProperty Column03 
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
            BeginProperty Column04 
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
            BeginProperty Column05 
               DataField       =   "codtipo"
               Caption         =   "codtipo"
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
            BeginProperty Column06 
               DataField       =   "horas"
               Caption         =   "horas"
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
                  ColumnWidth     =   2684.977
               EndProperty
               BeginProperty Column01 
                  Button          =   -1  'True
                  Locked          =   -1  'True
                  ColumnWidth     =   1035.213
               EndProperty
               BeginProperty Column02 
                  Button          =   -1  'True
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
      End
      Begin VB.ComboBox CmbMonBoleta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74820
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   975
         Width           =   1815
      End
      Begin VB.Frame Frame12 
         Caption         =   " SCTR - PENSION "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   -74865
         TabIndex        =   77
         Top             =   5760
         Width           =   3135
         Begin CboFacil.cbo_facil cbopension 
            Height          =   315
            Left            =   225
            TabIndex        =   47
            Top             =   225
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   0   'False
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   -71640
         TabIndex        =   76
         Top             =   5040
         Width           =   2535
         Begin VB.CheckBox Check1 
            Caption         =   "Sujeto a Control Inmediato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   51
            Top             =   840
            Width           =   2340
         End
         Begin VB.CheckBox ChkSindicato 
            Caption         =   "Afiliado al Sindicato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   50
            Top             =   555
            Width           =   2295
         End
         Begin VB.CheckBox ChkNoQuinta 
            Caption         =   "No se Calcula Qta Categoria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   49
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1905
         Left            =   -74880
         TabIndex        =   75
         Top             =   2430
         Width           =   11040
         Begin VB.ComboBox CmbCargo 
            Height          =   315
            ItemData        =   "Frmpersona.frx":268B1
            Left            =   6765
            List            =   "Frmpersona.frx":268B3
            Style           =   2  'Dropdown List
            TabIndex        =   243
            Top             =   585
            Width           =   4155
         End
         Begin VB.ComboBox CmbArea 
            Height          =   315
            ItemData        =   "Frmpersona.frx":268B5
            Left            =   3165
            List            =   "Frmpersona.frx":268B7
            Style           =   2  'Dropdown List
            TabIndex        =   242
            Top             =   195
            Width           =   7755
         End
         Begin VB.ComboBox CmbResponsable 
            Height          =   315
            ItemData        =   "Frmpersona.frx":268B9
            Left            =   7155
            List            =   "Frmpersona.frx":268BB
            Style           =   2  'Dropdown List
            TabIndex        =   233
            Top             =   1000
            Width           =   3770
         End
         Begin VB.TextBox txtPresentacion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   4440
            MaxLength       =   8
            TabIndex        =   36
            Top             =   577
            Width           =   1680
         End
         Begin VB.ComboBox CmbTrabSunat 
            Height          =   315
            Left            =   3030
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   990
            Width           =   3015
         End
         Begin MSMask.MaskEdBox TxtFecIngreso 
            Height          =   315
            Left            =   1155
            TabIndex        =   39
            Top             =   585
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txtcodpla 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   35
            Top             =   180
            Width           =   1440
         End
         Begin VB.ComboBox CmbSegCom 
            Height          =   315
            ItemData        =   "Frmpersona.frx":268BD
            Left            =   7560
            List            =   "Frmpersona.frx":268BF
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   1410
            Width           =   3375
         End
         Begin VB.ComboBox CmbHorasExtras 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   1395
            Width           =   2895
         End
         Begin VB.ComboBox CmbTipoTrabajador 
            Height          =   315
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Código de Presentación:"
            Height          =   255
            Left            =   2640
            TabIndex        =   241
            Top             =   555
            Width           =   1815
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   240
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Responsable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6120
            TabIndex        =   232
            Top             =   1060
            Width           =   915
         End
         Begin VB.Label LblTrabSunat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUNAT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2520
            TabIndex        =   207
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   2310
            TabIndex        =   127
            Top             =   585
            Width           =   330
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fec. Ing. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   81
            Top             =   585
            Width           =   765
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cargo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6240
            TabIndex        =   109
            Top             =   585
            Width           =   435
         End
         Begin VB.Label Lblplacod 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1440
            TabIndex        =   106
            Top             =   225
            Width           =   975
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod Trab. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   103
            Top             =   225
            Width           =   825
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seguro Compl. Trabajador Riesgo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5040
            TabIndex        =   84
            Top             =   1485
            Width           =   2415
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de horas Extras"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   83
            Top             =   1440
            Width           =   1485
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Trab. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   1035
            Width           =   840
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1680
         Left            =   -74880
         TabIndex        =   74
         Top             =   720
         Width           =   11070
         Begin VB.TextBox TxtCarrera 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5760
            TabIndex        =   229
            Top             =   1320
            Width           =   3540
         End
         Begin VB.TextBox TxtAnoEgreso 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   10440
            TabIndex        =   228
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TxtNombreUni 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6150
            TabIndex        =   224
            Top             =   840
            Width           =   4500
         End
         Begin VB.CheckBox ChkPais 
            Caption         =   "Estudio en el Exterior"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   8760
            TabIndex        =   223
            Top             =   160
            Width           =   2175
         End
         Begin VB.OptionButton OptPrivada 
            Caption         =   "Privada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   7200
            TabIndex        =   222
            Top             =   160
            Width           =   975
         End
         Begin VB.OptionButton OptPublica 
            Caption         =   "Pública"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   6120
            TabIndex        =   221
            Top             =   160
            Width           =   975
         End
         Begin CboFacil.cbo_facil cbo_TipModFor 
            Height          =   315
            Left            =   1275
            TabIndex        =   33
            Top             =   720
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            NameTab         =   ""
            NameCod         =   ""
            NameDesc        =   ""
            Filtro          =   ""
            OrderBy         =   ""
            SetIndex        =   ""
            NameSistema     =   ""
            Mensaje         =   0   'False
            ToolTip         =   -1  'True
            Enabled         =   -1  'True
            Ninguno         =   -1  'True
            BackColor       =   -2147483643
            BeginProperty ComboFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtprofesion 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   114
            Top             =   1280
            Width           =   3660
         End
         Begin VB.ComboBox CmbCentroForma 
            Height          =   315
            Left            =   6075
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   420
            Width           =   4815
         End
         Begin VB.ComboBox CmbNivelEdu 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   180
            Width           =   4005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Carrera"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5160
            TabIndex        =   231
            Top             =   1365
            Width           =   555
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   9360
            TabIndex        =   230
            Top             =   1320
            Width           =   330
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Egreso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9840
            TabIndex        =   227
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Inst."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5160
            TabIndex        =   226
            Top             =   855
            Width           =   945
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   10665
            TabIndex        =   225
            Top             =   840
            Width           =   330
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mod. Formativa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   128
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00EBFEFC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   4680
            TabIndex        =   126
            Top             =   1280
            Width           =   330
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ocupación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   80
            Top             =   1320
            Width           =   750
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Inst."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   5160
            TabIndex        =   79
            Top             =   465
            Width           =   690
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Situac. Edu."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   78
            Top             =   225
            Width           =   870
         End
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5820
         Left            =   -74880
         TabIndex        =   98
         Top             =   660
         Width           =   11220
      End
      Begin VB.Frame Frame30 
         Caption         =   "Frame30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1575
         Index           =   6
         Left            =   -67680
         TabIndex        =   179
         Top             =   4200
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Frame Framebuspla 
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Caption         =   "Frame19"
         Height          =   4575
         Left            =   120
         TabIndex        =   104
         Top             =   1260
         Visible         =   0   'False
         Width           =   11175
         Begin MSAdodcLib.Adodc AdoBusPla 
            Height          =   330
            Left            =   720
            Top             =   1680
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   8
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "AdoBusPla"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DgrdBusPla 
            Bindings        =   "Frmpersona.frx":268C1
            Height          =   4455
            Left            =   120
            TabIndex        =   105
            Top             =   120
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7858
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "codauxinterno"
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
               Caption         =   "Nombres"
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
               DataField       =   "ap_pat"
               Caption         =   "ap_pat"
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
               DataField       =   "ap_mat"
               Caption         =   "ap_mat"
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
               DataField       =   "ap_cas"
               Caption         =   "ap_cas"
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
               DataField       =   "nom_1"
               Caption         =   "nom_1"
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
            BeginProperty Column06 
               DataField       =   "nom_2"
               Caption         =   "nom_2"
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
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   9269.858
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column06 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label lblPlaCos 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   -72960
         TabIndex        =   214
         Top             =   6720
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label40 
         Caption         =   "Periodicidad de la Remuneración"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -72000
         TabIndex        =   163
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Lbldesobra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -74910
         TabIndex        =   112
         Top             =   7350
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label Lblobra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   110
         Top             =   6870
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label40 
         Caption         =   "Moneda Calculo de Boleta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74820
         TabIndex        =   85
         Top             =   750
         Width           =   2895
      End
   End
   Begin VB.Frame FrameUnif 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame19"
      Height          =   3135
      Left            =   600
      TabIndex        =   244
      Top             =   4680
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton CmdFRamUnif 
         Caption         =   "Salir"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FF0000&
         TabIndex        =   248
         Top             =   2880
         Width           =   5055
      End
      Begin MSDataGridLib.DataGrid DgrdUnif 
         Height          =   2295
         Left            =   120
         TabIndex        =   245
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
            DataField       =   "Referencia"
            Caption         =   "Referencia"
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
            DataField       =   "Talla"
            Caption         =   "Talla"
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Tallas de Uniformes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   246
         Top             =   120
         Width           =   5040
      End
   End
   Begin VB.Frame FrameTelf 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame19"
      Height          =   2655
      Left            =   5880
      TabIndex        =   99
      Top             =   4680
      Visible         =   0   'False
      Width           =   5295
      Begin Threed.SSCommand SSCommand5 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   102
         Top             =   2280
         Width           =   5055
         _Version        =   65536
         _ExtentX        =   8916
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Salir"
      End
      Begin MSDataGridLib.DataGrid DgrdTelf 
         Height          =   1815
         Left            =   120
         TabIndex        =   100
         Top             =   480
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
            DataField       =   "telefono"
            Caption         =   "Telefono"
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
            DataField       =   "fax"
            Caption         =   "Fax"
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
            DataField       =   "referencia"
            Caption         =   "Referencia"
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
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2594.835
            EndProperty
         EndProperty
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Lista de Telefonos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   101
         Top             =   120
         Width           =   5040
      End
   End
End
Attribute VB_Name = "Frmpersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VCivil As String
Dim VNacio As String
Dim VCodAfp As String
Dim VZona As String
Dim VVia As String
Dim VNivelEdu As String
Dim VCentroForma As String
Dim VProfesion As String
Dim VCargo As String
Dim Vpla_area As String
Dim VPlanta As String
Dim TipoDoc As String
Dim VTipotrab As String
Dim VTipoHorasExt As String
Dim VDistribCalcSeg As String
Dim VMonCalculo As String
Dim VTipoPago As String
Dim VBcoPago As String
Dim VMonPago As String
Dim VTipoCtaPago As String
Dim VBcoCts As String
Dim VMonCts As String
Dim VTipoCtaCts As String
Dim VCodAuxi As String
Dim Vcode As Integer
Dim VNewDerechoHab As Boolean
Dim VVinculoDh As String
Dim VsituacionDh As String
Dim VTipoDocDh As String
Dim VNew As Boolean
Dim Vbuscaaux As Boolean
Dim rstlf As New Recordset
Dim rsJud As New Recordset
Dim rsunif As New Recordset
Dim rscontrato As New Recordset
Dim rsdeduccion As New Recordset
Dim rsbasico As New Recordset
Dim rsderechohab As New Recordset
Dim rsOtrosEmpleadores As New Recordset
Dim rsSuspension4ta As New ADODB.Recordset
Dim rsTipEst As New ADODB.Recordset
Dim rsccosto As New Recordset
Dim TipoTrabSunat As String
Dim wciamae As String
Private Sub Carga_Combos()
Dim I, pos As Integer
Call fc_Descrip_Maestros2("01054", "", CmbCivil)
Call fc_Descrip_Maestros2("01032", "", Cbotipodoc, True)
Call rUbiIndCmbBox(Cbotipodoc, "01", "00")

Call fc_Descrip_Maestros2("01057", "", CmbNivelEdu, True)
Call fc_Descrip_Maestros2("01070", "", CmbPlanta)
Call fc_Descrip_Maestros2("01055", "", CmbTipoTrabajador, True)
Call fc_Descrip_Maestros2("01073", "", CmbHorasExtras)
Call rUbiIndCmbBox(CmbHorasExtras, "01", "00")
'se cambio el maestro 01161 por 01074
Call fc_Descrip_Maestros2("01074", "", CmbSegCom)

Call fc_Descrip_Maestros2_Mon("01006", "", CmbMonBoleta)
For I = 0 To CmbMonBoleta.ListCount - 1
    If Right(Left(CmbMonBoleta.List(I), 4), 3) = wmoncont Then CmbMonBoleta.ListIndex = I: Exit For
Next
Call fc_Descrip_Maestros2_Mon("01006", "", CmbMonPago)
For I = 0 To CmbMonPago.ListCount - 1
    If Right(Left(CmbMonPago.List(I), 4), 3) = wmoncont Then CmbMonPago.ListIndex = I: Exit For
Next
Call fc_Descrip_Maestros2("01072", "", CmbCtaPag)
Call fc_Descrip_Maestros2("01072", "", CmbCtaCts)
Call fc_Descrip_Maestros2("01060", "", CmbTipoPago, True)
Call fc_Descrip_Maestros2("01007", "", CmbBcoPago, False)
Call fc_Descrip_Maestros2_Mon("01006", "", CmbMonCts)
Call fc_Descrip_Maestros2("01007", "", CmbBcoCts, False)
Call fc_Descrip_Maestros2("01076", "", CmbPerRem, True)

Call fc_Descrip_Maestros2("01154", "", CmbEvitaDobleImp, True)

wciamae = Determina_Maestro("01006")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where status=''"
Sql$ = Sql$ & wciamae
Set Rs = cn.Execute(Sql$)
If Rs.RecordCount = 0 Then Exit Sub
Rs.MoveFirst
LstMoneda.Clear
Do Until Rs.EOF
   LstMoneda.AddItem Rs!flag1 & Space(1) & Rs!DESCRIP
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

LstTipo.Clear

wciamae = Determina_Maestro("01076")
Sql$ = "SELECT cod_maestro2,descrip,flag1 FROM maestros_2 where status='' and rtrim(isnull(codsunat,''))<>''"
Sql$ = Sql$ & wciamae

cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then Rs.MoveFirst
Do Until Rs.EOF
   LstTipo.AddItem Rs!DESCRIP & Space(100) & Rs!cod_maestro2
   Rs.MoveNext
Loop

' ZONAS DEL TRABAJADOR
cbo_zonatrab.NameTab = "maestros_2"
cbo_zonatrab.NameCod = "cod_maestro2"
cbo_zonatrab.NameDesc = "descrip"
cbo_zonatrab.Filtro = "right(ciamaestro,3)='035' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_zonatrab.conexion = cn
cbo_zonatrab.Execute
'**************************************
' VIAS DEL TRABAJADOR
cbo_viatrab.NameTab = "maestros_2"
cbo_viatrab.NameCod = "cod_maestro2"
cbo_viatrab.NameDesc = "descrip"
cbo_viatrab.Filtro = "right(ciamaestro,3)='036' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_viatrab.conexion = cn
cbo_viatrab.Execute
'**************************************
' CATEGORIA DEL TRABAJADOR
cbo_cattrab.NameTab = "maestros_2"
cbo_cattrab.NameCod = "cod_maestro2"
cbo_cattrab.NameDesc = "descrip"
cbo_cattrab.Filtro = "right(ciamaestro,3)='141' and status!='*'"
cbo_cattrab.conexion = cn
cbo_cattrab.Execute
cbo_cattrab.SetIndice "01"
'**************************************
' SITUACION DEL EPS
cbo_situacioneps.NameTab = "maestros_2"
cbo_situacioneps.NameCod = "cod_maestro2"
cbo_situacioneps.NameDesc = "descrip"
cbo_situacioneps.Filtro = "right(ciamaestro,3)='142' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_situacioneps.conexion = cn
cbo_situacioneps.Execute
'**************************************
' EPS
cbo_eps.NameTab = "maestros_2"
cbo_eps.NameCod = "cod_maestro2"
cbo_eps.NameDesc = "descrip"
cbo_eps.Filtro = "right(ciamaestro,3)='143' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_eps.conexion = cn
cbo_eps.Execute
'**************************************
' PENSIONES
cbo_pensiones.NameTab = "maestros_2"
cbo_pensiones.NameCod = "cod_maestro2"
cbo_pensiones.NameDesc = "descrip"
cbo_pensiones.Filtro = "right(ciamaestro,3)='069' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_pensiones.conexion = cn
cbo_pensiones.Execute
'**************************************
' TIPO DE PENSION
'cbo_pension.NameTab = "maestros_2"
'cbo_pension.NameCod = "cod_maestro2"
'cbo_pension.NameDesc = "descrip"
'cbo_pension.Filtro = "right(ciamaestro,3)='140' and status!='*'"
'cbo_pension.Conexion = cn
'cbo_pension.Execute
'**************************************
' TIPO DE CONTRATOS
cbo_tipocont.NameTab = "maestros_2"
cbo_tipocont.NameCod = "cod_maestro2"
cbo_tipocont.NameDesc = "descrip"
cbo_tipocont.Filtro = "RIGHT(ciamaestro,3)='144' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_tipocont.conexion = cn
cbo_tipocont.Execute


'**************************************
' SCTR ESSALUD
cboessalud.NameTab = "maestros_2"
cboessalud.NameCod = "cod_maestro2"
cboessalud.NameDesc = "descrip"
cboessalud.Filtro = "RIGHT(ciamaestro,3)='145' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cboessalud.conexion = cn
cboessalud.Execute
'**************************************
' SCTR PENSION

cbopension.NameTab = "maestros_2"
cbopension.NameCod = "cod_maestro2"
cbopension.NameDesc = "descrip"
cbopension.Filtro = "RIGHT(ciamaestro,3)='146' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbopension.conexion = cn
cbopension.Execute
'**************************************
'NACIONES
'With cbo_nacion
'    .NameCod = "codigo"
'    .NameDesc = "descripcion"
'    .NameTab = "tnacionalidad"
'    .Conexion = cn
'    .Execute
'End With
'****************************
'NACIONES
With cbo_naciontrab
    .NameCod = "codigo"
    .NameDesc = "descripcion"
    .NameTab = "tnacionalidad"
    .conexion = cn
    .Execute
End With

'****************************
'TIPO MOTIVO FIN DE PERIODO
With cbo_TipMotFinPer
    .NameTab = "maestros_2"
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .Filtro = "RIGHT(ciamaestro,3)='149' and status!='*' and rtrim(isnull(codsunat,''))<>''"
    .conexion = cn
    .Execute
End With


'****************************
'TIPO MODALIDAD FORMATIVA
With cbo_TipModFor
    .NameTab = "maestros_2"
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .Filtro = "RIGHT(ciamaestro,3)='150' and status!='*' and rtrim(isnull(codsunat,''))<>''"
    .conexion = cn
    .Execute
End With

'*****************

'Responsables
Sql$ = "SELECT COD_MAESTRO3,DESCRIP FROM MAESTROS_31 WHERE CIAMAESTRO = '" & wcia & "087" & "' AND STATUS != '*' ORDER BY DESCRIP"
Call rCarCbo(CmbResponsable, Sql$, "C", "000")


End Sub

Private Sub Cmbafp_Click()
VCodAfp = fc_CodigoComboBox(Cmbafp, 2)
If VCodAfp = "99" Then
   TxtNroAfp.Text = ""
   TxtFecAfilia.Text = "__/__/____"
   TxtNroAfp.Enabled = False
   TxtFecAfilia.Enabled = False
   CMBZONA.SetFocus
Else
   TxtNroAfp.Enabled = True
   TxtFecAfilia.Enabled = True
   TxtNroAfp.SetFocus
End If
End Sub

Private Sub BtnMemo_Click()
Memo_Cese (Trim(Txtcodpla.Text))
End Sub

Private Sub cbo_cattrab_Click()
Dim xId As String
xId = fc_CodigoComboBox(cbo_cattrab, 2)
Select Case xId
Case "05" 'PRESTADOR DE SERVICIOS - MODALIDAD FORMATIVA
    ChkMadreResp(1).Visible = True
'Case "03" 'PRESTADOR DE SERVICIOS - 4TA CATEGORIA
    
Case Else
    ChkMadreResp(1).Visible = False
    
End Select
End Sub



Private Sub cbo_pensiones_Click()
VCodAfp = Format(cbo_pensiones.ReturnCodigo, "00")
'add jcms 05/08/2008
If Trim(VCodAfp) = "01" Then 'ley 19990
    FraAseguraTuPension.Enabled = True
Else
    FraAseguraTuPension.Enabled = False
    Me.OptAsegPenNo.Value = True
End If

End Sub


Private Sub Cbotipodoc_Click()
TipoDoc = fc_CodigoComboBox(Cbotipodoc, 2)
End Sub

Private Sub Cbotipodoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub ChkAltitud_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub ChkDomiciliado_Click()
If Me.ChkDomiciliado.Value = 1 Then
    Frame4.Enabled = True
Else
    LimpiarNoDomiciliado
    Frame4.Enabled = False
End If
End Sub

Private Sub ChkNoQuinta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub ChkOtrosIngta_Click(Index As Integer)
If Me.ChkOtrosIngta(Index).Value = 1 Then
    FraOtrosEmpleadores(4).Enabled = True
 '   rsOtrosEmpleadores.AddNew
    DgrdOtrosEmpleadores.Col = 0
    DgrdOtrosEmpleadores.SetFocus
Else
    FraOtrosEmpleadores(4).Enabled = False
    If rsOtrosEmpleadores.RecordCount > 0 Then rsOtrosEmpleadores.MoveFirst
    Do While Not rsOtrosEmpleadores.EOF
        rsOtrosEmpleadores.Delete
        rsOtrosEmpleadores.MoveNext
    Loop
End If
End Sub

Private Sub ChkSindicato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub ChkVaca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbArea_Click()
Vpla_area = fc_CodigoComboBox(CmbArea, 4)
End Sub

Private Sub CmbArea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbBcoCts_Click()
VBcoCts = fc_CodigoComboBox(CmbBcoCts, 2)
CmbBcoCts.ToolTipText = CmbBcoCts.Text
End Sub

Private Sub CmbBcoCts_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbBcoPago_Click()
VBcoPago = fc_CodigoComboBox(CmbBcoPago, 2)
CmbBcoPago.ToolTipText = CmbBcoPago.Text
End Sub

Private Sub CmbBcoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbCargo_Click()
VCargo = fc_CodigoComboBox(CmbCargo, 3)
CmbCargo.ToolTipText = CmbCargo.Text
End Sub

Private Sub CmbCargo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbCentroForma_Click()
'VCentroForma = fc_CodigoComboBox(CmbCentroForma, 2)
VCentroForma = Left(CmbCentroForma, 2)
End Sub

Private Sub CmbCentroForma_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Carga_Combos
SSTab1.Tab = 0
End Sub

Private Sub CmbCivil_Click()
VCivil = fc_CodigoComboBox(CmbCivil, 2)
End Sub

Private Sub CmbCivil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbCtaCts_Click()
VTipoCtaCts = fc_CodigoComboBox(CmbCtaCts, 2)
End Sub

Private Sub CmbCtaCts_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbCtaPag_Click()
VTipoCtaPago = fc_CodigoComboBox(CmbCtaPag, 2)
 
End Sub

Private Sub CmbCtaPag_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbSegCom_Click()
' VDistribCalcSeg = fc_CodigoComboBox(CmbSegCom, 2)
End Sub

Private Sub CmbSegCom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbEvitaDobleImp_KeyDown(KeyCode As Integer, Shift As Integer)
If vbKeyDelete = KeyCode Then CmbEvitaDobleImp.ListIndex = -1
End Sub

Private Sub CmbHorasExtras_Click()
VTipoHorasExt = fc_CodigoComboBox(CmbHorasExtras, 2)
End Sub

Private Sub CmbHorasExtras_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbMonBoleta_Click()
VMonCalculo = Mid(CmbMonBoleta.Text, 2, 3)
End Sub

Private Sub CmbMonBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbMonCts_Click()
VMonCts = Mid(CmbMonCts.Text, 2, 3)
End Sub

Private Sub CmbMonCts_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbMonPago_Click()
VMonPago = Mid(CmbMonPago.Text, 2, 3)
End Sub

Private Sub CmbNacio_Click()
VNacio = fc_CodigoComboBox(CmbNacio, 3)
TxtDni.SetFocus
End Sub

Private Sub CmbMonPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbNivelEdu_Click()
VNivelEdu = fc_CodigoComboBox(CmbNivelEdu, 2)
End Sub

Private Sub CmbNivelEdu_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbPlanta_Click()
VPlanta = fc_CodigoComboBox(CmbPlanta, 2)
If VPlanta = "00" Then
   LblCodCantera.Visible = True
   Label1(5).Visible = True
   LblCantera.Visible = True
Else
   LblCodCantera.Visible = False
   Label1(5).Visible = False
   LblCantera.Visible = False
End If
Carga_Centro_Costos
End Sub

Private Sub CmbProfesion_Click()
VProfesion = fc_CodigoComboBox(CmbProfesion, 3)
End Sub

Private Sub CmbSituaDh_Click()
VsituacionDh = Left(CmbSituaDh.Text, 1)
End Sub

Private Sub CmbPlanta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub


Private Sub CmbTipoPago_Click()
VTipoPago = fc_CodigoComboBox(CmbTipoPago, 2)
End Sub

Private Sub CmbTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Cmbtipotrabajador_Click()
VTipotrab = fc_CodigoComboBox(CmbTipoTrabajador, 2)
Sql$ = "SELECT COD_MAESTRO3,DESCRIP FROM MAESTROS_31 WHERE CIAMAESTRO = '" & wcia & "055" & "' AND STATUS != '*' ORDER BY DESCRIP"
Call rCarCbo(CmbCargo, Sql$, "C", "000")

Sql$ = "SELECT cod_area,gerencia + ' / ' + dpto  FROM pla_areas WHERE STATUS != '*' ORDER BY cod_area"
Call rCarCbo(CmbArea, Sql$, "C", "0000")

Txtcodobra.Text = ""
Lblobra.Caption = ""
If VTipotrab = "05" Then
   Lblobra.Visible = True
   Txtcodobra.Visible = True
   Lbldesobra.Visible = True
   ChkAltitud.Visible = True
Else
   Lblobra.Visible = False
   Txtcodobra.Visible = False
   Lbldesobra.Visible = False
   ChkAltitud.Visible = False
   ChkAltitud.Value = 0
End If

'Centro de Costos

'Sql$ = "Select cod_maestro3,descrip from maestros_32 where ciamaestro='" & wcia & "055' and "
'Sql$ = Sql$ & "cod_maestro2='" & Trim(VTipotrab) & "' and status<>'*' ORDER BY 2"
Carga_Centro_Costos
End Sub
Private Sub Carga_Centro_Costos()
LstCcosto.Clear
If CmbPlanta.Text = "CANTERA" Then
   Sql$ = "select codigo,descripcion From Pla_ccostos where codigo in('24','37') and status<>'*' order by descripcion"
Else
   'Sql$ = "select codigo,descripcion From Pla_ccostos where status<>'*' and CODIGO NOT IN (SELECT COD_MAESTRO2 FROM MAESTROS_2 WHERE ciamaestro='01044' AND FLAG1='' AND status<>'*') order by descripcion"
   'add jcms se excluye 081122 CV+JJ
   Sql$ = "select codigo,descripcion From Pla_ccostos where status<>'*' order by descripcion"
End If

cn.CursorLocation = adUseClient
Set Rq = New ADODB.Recordset
Set Rq = cn.Execute(Sql$, 64)
If Rq.RecordCount > 0 Then Rq.MoveFirst
Do Until Rq.EOF
   'LstCcosto.AddItem Rq!DESCRIP & Space(100) & Rq!COD_MAESTRO3
   LstCcosto.AddItem Rq!Descripcion & Space(100) & Rq!Codigo
   Rq.MoveNext
Loop
If Rq.State = 1 Then Rq.Close
If rsccosto.State = 1 Then
    If rsccosto.RecordCount > 0 Then
       rsccosto.MoveFirst
       Do While Not rsccosto.EOF
          rsccosto!Codigo = ""
          rsccosto!Descripcion = ""
          rsccosto!Monto = 0
          rsccosto.MoveNext
       Loop
    End If
End If

End Sub


Private Sub CmbVia_Click()
VVia = fc_CodigoComboBox(CmbVia, 2)
TxtVia.SetFocus
End Sub

Private Sub CmbTipoTrabajador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub CmbZona_Click()
VZona = fc_CodigoComboBox(CMBZONA, 2)
txtzona.SetFocus
End Sub

Private Sub CmdFRamUnif_Click()
FrameUnif.Visible = False
SSCommand6(6).Enabled = True
End Sub

Private Sub CmdGenContrato_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    If rscontrato.RecordCount = 0 Then
        MsgBox "No existen contratos definidos", vbExclamation, Me.Caption
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf rscontrato.AbsolutePosition = -1 Then
        MsgBox "Elija un periodo de contrato", vbExclamation, Me.Caption
        Screen.MousePointer = vbDefault
        Exit Sub
    ElseIf Trim(rscontrato!nro_contrato) = "" Then
        MsgBox "El contrato no ha sido grabado aun", vbExclamation, Me.Caption
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    Dim RUTA As New FileSystemObject
    Dim Path As String
    Dim Carpeta As String
    Dim rsGenerar As ADODB.Recordset
    Dim Nuevo As String
    Dim pContrato As String
    Dim pArchivo As String
    
    If rscontrato.RecordCount > 0 Then
        Sql$ = "exec paGeneraContrato '" & wcia & "', '" & Trim(Txtcodpla) & "', '" & rscontrato!nro_contrato & "'"
        Set rsGenerar = OpenRecordset(Sql$, cn)
        If rsGenerar.EOF = True Then MsgBox "No se han encontrados los datos necesarios para generar el archivo.", vbInformation + vbOKOnly, "Sistema de Planillas": Exit Sub
        
        If rsGenerar.RecordCount > 0 Then
            Carpeta = App.Path & "\" & "Contratos"
            If fso.FolderExists(Carpeta) = False Then
                fso.CreateFolder (Carpeta)
            End If
            
            Carpeta = Carpeta & "\" & wruc
            If fso.FolderExists(Carpeta) = False Then
                fso.CreateFolder (Carpeta)
            End If
            
            Carpeta = Carpeta & "\" & Format(Date, "YYYY")
            If fso.FolderExists(Carpeta) = False Then
                fso.CreateFolder (Carpeta)
            End If
                                 
            pArchivo = Trim(Txtcodpla.Text)
            pArchivo = pArchivo & Space(1) & Mid(Trim(rsGenerar!nombre), 1, InStr(Trim(rsGenerar!nombre), ",") - 1)
            pArchivo = pArchivo & Space(1) & Trim(rscontrato!nro_contrato) & ".doc"
            pArchivo = Carpeta & "\" & pArchivo
            If fso.FileExists(pArchivo) = True Then
                fso.DeleteFile (pArchivo)
            End If
            
            Select Case wruc
                Case "20100084253"
                    pContrato = "CONTRATO.doc"
                Case "20100574005"
                    pContrato = "CONTRATO2.doc"
                Case Else
                    pContrato = "CONTRATX.doc"
            End Select
            pContrato = App.Path & "\Reports\" & pContrato
            fso.CopyFile pContrato, pArchivo, True
            Call EnviaWord(rsGenerar, pArchivo)
        Else
            MsgBox "No se han encontrados los datos necesarios para generar el archivo.", vbInformation + vbOKOnly, "Sistema de Planillas": Exit Sub
        End If
        MsgBox "Archivo Generado con Exito.", vbInformation + vbOKOnly, "Sistema de Planillas"
    Else
        MsgBox "Debe de crear al menos un registro para poder realizar este proceso.", vbInformation + vbOKOnly, "Sistema de Planillas": Exit Sub
    End If
    Set rsGenerar = Nothing
End Sub

Private Sub CmdGrabarContrato_Click(Index As Integer)
On Error GoTo ErrorTrans
Dim NroTrans As Integer
NroTrans = 0
If rscontrato.RecordCount = 0 Then
    MsgBox "No existen contratos definidos", vbExclamation, Me.Caption
    Exit Sub
End If
Dim xId As String
xId = fc_CodigoComboBox(cbo_cattrab, 2)
Dim xMeses As Integer
xMeses = 0
With rscontrato
    If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Not IsDate(!FecIni) Then
                        MsgBox "Ingrese fecha de inicio del contrato", vbCritical, Me.Caption
                        Me.DgrdContrato.Col = 1
                        Me.DgrdContrato.SetFocus
                        Exit Sub
                    ElseIf Not IsDate(!FecFin) Then
                        MsgBox "Ingrese fecha de Fin del contrato", vbCritical, Me.Caption
                        Me.DgrdContrato.Col = 2
                        Me.DgrdContrato.SetFocus
                        Exit Sub
                    ElseIf CCur(!imp_sueldo) = 0 Then
                        MsgBox "El importe del suelto es cero", vbCritical, Me.Caption
                        Me.DgrdContrato.Col = 1
                        Me.DgrdContrato.SetFocus
                        Exit Sub
                    End If
                    If Trim(!Tipo_contrato) = "" Then
                        SSTab1.Tab = 0
                        MsgBox "Elija Tipo de  Contrato del trabajador de la lista, correctamente antes de generar su contrato", vbExclamation, Me.Caption
                        cbo_tipocont.SetFocus
                        Exit Sub
                    End If
                
                    If Trim(!cod_Cargo) = "" Then
                        SSTab1.Tab = 1
                        MsgBox "Elija Cargo del trabajador de la lista, correctamente antes de generar su contrato", vbExclamation, Me.Caption
                        CmbCargo.SetFocus
                        Exit Sub
                    End If
                    If IsDate(!FecFin) Then
                            If CDate(!FecIni) < CDate(!FecIni) Then
                                MsgBox "la Fecha final no puede ser anterior a la fecha inicial, verifique", vbCritical, Me.Caption
                                Me.DgrdContrato.Col = 1
                                Me.DgrdContrato.SetFocus
                                Exit Sub
                            End If
                     End If
                    
                .MoveNext
            Loop
    End If
End With
Screen.MousePointer = 11

cn.BeginTrans
NroTrans = 1
Sql = "update placontrato set status='*',user_modI='" & Trim(wuser) & "',fec_modi=getdate() where codcia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and status<>'*'"
cn.Execute Sql, 64

With rscontrato
    If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!nro_contrato & "") = "" Then
                        Dim xNumContrato As String
                        Dim Rq As ADODB.Recordset
                        Sql = "select isnull(max(num_contrato),0) from placontrato where codcia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "'"
                        If fAbrRst(Rq, Sql) Then
                            xNumContrato = Format(Rq(0) + 1, "00000000")
                        Else
                            xNumContrato = "00000001"
                        End If
                    Else
                        xNumContrato = Trim(!nro_contrato)
                    End If
                    
                    Sql = "insert into placontrato (codcia,num_contrato,placod,fec_ini,fec_fin,cargo,importe,fec_crea,user_crea,fec_modi,user_modi,cod_fperiodo,status,cod_mod_formativa,cod_tip_contrato)"
                    Sql = Sql & " values('" & wcia & "','" & xNumContrato & "','" & Trim(Txtcodpla.Text) & "','" & Format(!FecIni, "mm/dd/yyyy") & "'," & IIf(IsDate(!FecFin), "'" & Format(!FecFin, "mm/dd/yyyy") & "'", "null")
                    Sql = Sql & ",'" & Trim(!cod_Cargo & "") & "'," & CCur(!imp_sueldo) & ",getdate(),'" & Trim(wuser) & "',null,'','" & Trim(!cod_mot_fin_periodo & "") & "','','" & Trim(!cod_tip_modalidad_formativa & "") & "','" & !Tipo_contrato & "')"
                    cn.Execute Sql, 64
                                  
                .MoveNext
            Loop
    End If
End With
        
cn.CommitTrans
Screen.MousePointer = 0
Carga_detalle_contratos Trim(Txtcodpla.Text)
MsgBox "Contrato guardado correctamente", vbInformation, Me.Caption
Exit Sub

ErrorTrans:
    If NroTrans = 1 Then
        cn.RollbackTrans
    End If
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub CmdJud_Click()
FrameJud.Visible = True
FrameJud.ZOrder 0
CmdJud.Enabled = False
End Sub

Private Sub CmdNewContrato_Click(Index As Integer)
 Dim xId As String
    xId = fc_CodigoComboBox(cbo_cattrab, 2)
    If Trim(Txtcodpla.Text) = "" Then
        SSTab1.Tab = 1
        MsgBox "Ingrese Código de Trabajador", vbCritical, Me.Caption
        Txtcodpla.SetFocus
        Exit Sub
    End If
    If Trim(cbo_tipocont.Text) = "NINGUNO" Or Trim(cbo_tipocont.Text) = "" Or cbo_tipocont.ListIndex = -1 Then
        SSTab1.Tab = 0
        MsgBox "Elija Tipo de  Contrato del trabajador de la lista, correctamente antes de generar su contrato", vbExclamation, Me.Caption
        cbo_tipocont.SetFocus
        Exit Sub
    End If

    If Trim(CmbCargo.Text) = "NINGUNO" Or Trim(CmbCargo.Text) = "" Or CmbCargo.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Cargo del trabajador de la lista, correctamente antes de generar su contrato", vbExclamation, Me.Caption
        CmbCargo.SetFocus
        Exit Sub
    End If
    If rscontrato.RecordCount = 0 Then
        Apertura_Registro_Contrato_Trabajador
    Else
        If rscontrato.RecordCount > 0 Then
            rscontrato.MoveFirst
            Do While Not rscontrato.EOF
                    If Not IsDate(rscontrato!FecIni) Then
                        MsgBox "Ingrese fecha de inicio del periodo de actividad antes de crear un nuevo contrato", vbCritical, Me.Caption
                        Me.DgrdContrato.Col = 1
                        Me.DgrdContrato.SetFocus
                        Exit Sub
                    End If
                    If Not IsDate(rscontrato!FecFin) Then
                        MsgBox "Cierre el periodo de actividad antes de crear un nuevo contrato", vbCritical, Me.Caption
                        Me.DgrdContrato.Col = 2
                        Me.DgrdContrato.SetFocus
                        Exit Sub
                    End If
                rscontrato.MoveNext
            Loop
        End If
         Apertura_Registro_Contrato_Trabajador
        
    End If
    
End Sub
Private Sub cmdubigeo_Click()
Framebuspla.Visible = False
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub CmdSalirUnif_Click()
FrameUnif.Visible = False
'SSCommand1.Enabled = True
End Sub

Private Sub CmdSalirJud_Click()

With rsJud
If .RecordCount > 0 Then
    .MoveFirst
    Do While Not .EOF
            If Trim(!CalcBaseBruto & "") <> "" Then
                If Not (Trim(!CalcBaseBruto & "") = "S" Or Trim(!CalcBaseBruto & "") = "") Then
                    MsgBox "El único valor permitido es la letra (S), para indicar que el calculo de la retención judicial será en base al Importe Total Bruto de los ingresos", vbExclamation, Me.Caption
                    Exit Sub
                End If
            End If
            
        .MoveNext
    Loop
End If
End With
FrameJud.Visible = False
CmdJud.Enabled = True
End Sub

Private Sub DgrBasico_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If DgrBasico.Col = 2 Then
        KeyAscii = 0
        Cancel = True
        DgrBasico_ButtonClick (ColIndex)
End If
End Sub
Private Sub DgrBasico_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer
Y = DgrBasico.Row
xtop = DgrBasico.Top + DgrBasico.RowTop(Y) + DgrBasico.RowHeight
Select Case ColIndex
Case 1:
       xleft = DgrBasico.Left + DgrBasico.Columns(1).Left
       With LstTipo
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = DgrBasico.Top + DgrBasico.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
Case 2:
       xleft = DgrBasico.Left + DgrBasico.Columns(2).Left
       With LstMoneda
       If Y < 8 Then
         .Top = xtop
       Else
         .Top = DgrBasico.Top + DgrBasico.RowTop(Y) - .Height
       End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub

Private Sub DgrdBusPla_DblClick()
Sql = "select placod from planillas where cia='" & wcia & "' and codauxinterno='" & Trim(DgrdBusPla.Columns(0)) & "' and status<>'*'"
If (fAbrRst(Rs, Sql)) Then
   VNew = False
   Carga_Trabajador (Rs!PlaCod)
Else
   TxtApePat.Text = Trim(DgrdBusPla.Columns(2))
   TxtApeMat.Text = Trim(DgrdBusPla.Columns(3))
   TxtApeCas.Text = Trim(DgrdBusPla.Columns(4))
   TxtPriNom.Text = Trim(DgrdBusPla.Columns(5))
   TxtSegNom.Text = Trim(DgrdBusPla.Columns(6))
End If
If Rs.State = 1 Then Rs.Close
Framebuspla.Visible = False
TxtFecNac.SetFocus
End Sub

Private Sub Dgrdccosto_ButtonClick(ByVal ColIndex As Integer)
Dim Y As Integer, xtop As Integer, xleft As Integer

Y = Dgrdccosto.Row

xtop = Dgrdccosto.Top + 500

Select Case ColIndex
Case 0:
       xleft = Dgrdccosto.Left + Dgrdccosto.Columns(0).Left
       With LstCcosto
       'If Y < 8 Then
         .Top = xtop
       'Else
       '  .Top = Dgrdccosto.Top + Dgrdccosto.RowTop(Y) - .Height
       'End If
        .Left = xleft
        .Visible = True
        .SetFocus
        .ZOrder 0
       End With
End Select
End Sub

Private Sub DgrdContrato_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Select Case ColIndex
Case 2
'    If IsDate(rscontrato!FecFin) Then
'
'        If Trim(cbo_TipMotFinPer.Text) = "NINGUNO" Or Trim(cbo_TipMotFinPer.Text) = "" Or cbo_TipMotFinPer.ListIndex = -1 Then
'            SSTab1.Tab = 1
'            rscontrato!FecFin = ""
'            MsgBox "Elija Motivo de fin de Periodo de la lista, correctamente" & Chr(13) & "antes de poner la fecha fin de periodo", vbExclamation, Me.Caption
'            cbo_TipMotFinPer.SetFocus
'            Exit Sub
'        End If
'    End If
End Select
End Sub

Private Sub DgrdContrato_ButtonClick(ByVal ColIndex As Integer)
Select Case ColIndex
Case 6
    
End Select
End Sub

Private Sub DgrdContrato_OnAddNew()
Apertura_Registro_Contrato_Trabajador
End Sub

Private Sub DgrDeduccion_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 1
            If DgrDeduccion.Columns(1) <> 0 Then DgrDeduccion.Columns(2) = "0.00"
       Case Is = 2
            If DgrDeduccion.Columns(2) <> 0 Then DgrDeduccion.Columns(1) = "0.00"
End Select
End Sub

Private Sub DgrdJud_OnAddNew()
rsJud.AddNew
End Sub

Private Sub DgrdOtrosEmpleadores_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
Case 0
    Sql = "select * from plaOtrosEmpleadores where cia='" & wcia & "' and ruc='" & rsOtrosEmpleadores!RUC & "' and status<>'*'"
    Dim Ro As ADODB.Recordset
    If fAbrRst(Ro, Sql) Then
        rsOtrosEmpleadores!razsoc = Trim(Ro!razsoc & "")
    End If
    Ro.Close
    Set Ro = Nothing
End Select
End Sub

Private Sub DgrdOtrosEmpleadores_OnAddNew()
rsOtrosEmpleadores.AddNew
End Sub

Private Sub DgrdrsSuspension4ta_OnAddNew()
rsSuspension4ta.AddNew
End Sub

Private Sub DgrdrTipEst_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 0 Then DgrdrTipEst.Update
End Sub

Private Sub DgrdrTipEst_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
On Error Resume Next:
If Trim(DgrdrTipEst.Splits(Split).Columns(0).CellValue(Bookmark)) = True Then
    RowStyle.BackColor = &HEDDCDC          '&HC0FFFF
Else
    RowStyle.BackColor = vbWhite
End If
End Sub

Private Sub DgrdTelf_OnAddNew()
rstlf.AddNew
End Sub


Private Sub DgrdUnif_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 8 'para borrar
Case 48 To 57 'numeros
Case 65 To 90 'mayusculas
Case 97 To 122 'minusculas conevertidas en mayusculas
KeyAscii = KeyAscii - 32
Case Else 'el resto no tipea
KeyAscii = 0
End Select
End Sub

Private Sub DgrdUnif_OnAddNew()
'rsunif.AddNew
End Sub

Private Sub Form_Activate()
TxtApePat.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 11610
Me.Height = 9305
Vbuscaaux = True
LblFecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Crea_Rs
Carga_Establecimientos True
TipoTrabSunat = ""
Dim RqTS As ADODB.Recordset
Sql = "select TipoTrabSunat from cia where cod_cia='" & wcia & "' and status<>'*'"
If fAbrRst(RqTS, Sql) Then TipoTrabSunat = Trim(RqTS(0) & "")
RqTS.Close: Set RqTS = Nothing
If TipoTrabSunat = "S" Then
   CmbTipoTrabajador.Width = 1095
   CmbTrabSunat.Visible = True
   LblTrabSunat.Visible = True
   Sql = "select * from TipoTrabSunat"
   If fAbrRst(RqTS, Sql) Then RqTS.MoveFirst
   Do While Not RqTS.EOF
      CmbTrabSunat.AddItem Trim(RqTS!DESCRIP)
      CmbTrabSunat.ItemData(CmbTrabSunat.NewIndex) = Trim(RqTS!Codigo)
      RqTS.MoveNext
   Loop
Else
   CmbTipoTrabajador.Width = 4695
   CmbTrabSunat.Visible = False
   LblTrabSunat.Visible = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmGrdPersonal.Procesa_Personal (False)
End Sub

Private Sub Label1_Click(Index As Integer)
If Index = 1 Then
   FrmAyuda.Busqueda = "OCUPACION"
   FrmAyuda.Regimen = 0
   FrmAyuda.Tipoinst = 0
ElseIf Index = 3 Then
   FrmAyuda.Busqueda = "UNIVERSIDAD"
   
   Dim Regimen As Integer
   Regimen = 0
   If OptPublica.Value = True Then Regimen = 1
   If OptPrivada.Value = True Then Regimen = 2
   FrmAyuda.Regimen = Regimen
   
   FrmAyuda.Tipoinst = Val(Left(CmbCentroForma.Text, 2))
ElseIf Index = 4 Then
   FrmAyuda.Busqueda = "CARRERA"
   FrmAyuda.Regimen = 0
   FrmAyuda.Tipoinst = 0
ElseIf Index = 5 Then
   FrmAyuda.Busqueda = "CANTERAS"
   FrmAyuda.Regimen = 0
   FrmAyuda.Tipoinst = 0
End If

Load FrmAyuda
FrmAyuda.Visible = True
FrmAyuda.ZOrder 0

End Sub

Private Sub Label212_Click()
pcthis.Visible = False
End Sub


Private Sub Lblubigeo_Click(Index As Integer)
Select Case Index
Case 3
    FrmUbiSunat.TipoCon = 0
Case 1
    FrmUbiSunat.TipoCon = 1
Case 2
    FrmUbiSunat.TipoCon = 2
End Select
Load FrmUbiSunat
FrmUbiSunat.Show
FrmUbiSunat.ZOrder 0

End Sub

Private Sub LstCcosto_Click()
Dim m As Integer

If LstCcosto.ListIndex > -1 Then
   m = Len(LstCcosto.Text) - 3
    
   Dgrdccosto.Columns(0) = Mid(Trim(Left(LstCcosto.Text, m)), 1, 40)
   Dgrdccosto.Columns(2) = Format(Right(Trim(LstCcosto.Text), 2), "00")
   Dgrdccosto.Col = 1
   Dgrdccosto.SetFocus
   LstCcosto.Visible = False
End If
End Sub

Private Sub LstCcosto_LostFocus()
LstCcosto.Visible = False
End Sub

Private Sub LstMoneda_Click()
If Vcode = 0 Then Vcode = 13
Call LstMoneda_KeyDown(Vcode, 0)
Vcode = 0
End Sub

Private Sub LstMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
Vcode = KeyCode
If KeyCode <> 13 Then Exit Sub
If Trim(LstMoneda) <> "" Then
   DgrBasico.Columns(2) = Mid$(Trim(LstMoneda), 2, 3)
   DgrBasico.Col = DgrBasico.Col + 1
   LstMoneda.Visible = False
End If
Vcode = 0
End Sub

Private Sub LstMoneda_LostFocus()
LstMoneda.Visible = False
Vcode = 0
End Sub

Private Sub Lsttipo_Click()
Dim m, P As Integer

If LstTipo.ListIndex > -1 Then
   m = Len(Trim(LstTipo.Text)) - 2
   
   DgrBasico.Columns(1) = Trim(Mid(LstTipo.Text, 1, m))
   DgrBasico.Columns(5) = Format(Right(RTrim(LstTipo.Text), 2), "00")
   wciamae = Determina_Maestro("01076")
   Sql$ = "Select cod_maestro2,flag2,descrip from maestros_2 " & _
   "where  status<>'*' and cod_maestro2='" & _
   Format(Right(LstTipo.Text, 2), "00") & "' "
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$, 64)
   If Rs.RecordCount > 0 Then DgrBasico.Columns(6) = Rs!flag2
   If Rs.State = 1 Then Rs.Close
   DgrBasico.Col = 2
   DgrBasico.SetFocus
   LstTipo.Visible = False
End If
End Sub

Private Sub Lsttipo_LostFocus()
LstTipo.Visible = False
End Sub

Private Sub OpcVidaNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub OpcVidaSi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub optestrabpen_Click()
If optestrabpen.Value = True Then
    cbo_pension.Enabled = True
End If
End Sub

Private Sub opttrabpen_Click()
If opttrabpen.Value = True Then
    cbo_pension.ListIndex = 5
    cbo_pension.Enabled = False
End If
End Sub

Private Sub OptEpsNo_Click(Index As Integer)
cbo_eps.Enabled = False
cbo_eps.SetIndex = -1
End Sub

Private Sub OptEpsSi_Click(Index As Integer)
cbo_eps.Enabled = True
cbo_eps.SetFocus

End Sub

Private Sub OptPrivada_Click()
Carga_Educacion
End Sub

Private Sub OptPublica_Click()
Carga_Educacion
End Sub

'Private Sub SSCommand1_Click(0)
'FrameTelf.Visible = True
'FrameTelf.ZOrder 0
'SSCommand1(0).Enabled = False
'End Sub

Private Sub SSCommand1_Click(Index As Integer)
FrameTelf.Visible = True
FrameTelf.ZOrder 0
SSCommand1(0).Enabled = False
End Sub

'Private Sub SSCommand1_Click()
'FrameTelf.Visible = True
'FrameTelf.ZOrder 0
'SSCommand1.Enabled = False
'End Sub

Private Sub SSCommand2_Click()
If TxtEmail.Text <> "" Then
     Call ShellExecute(hwnd, "Open", "mailto:" & TxtEmail.Text & "?Subject=" & Cmbcia.Text & "", "", "", vbNormalFocus)
End If
End Sub

Private Sub SSCommand3_Click()
Load FrmUbigeo
FrmUbigeo.Show
FrmUbigeo.ZOrder 0
End Sub

Private Sub SSCommand2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub SSCommand5_Click(Index As Integer)
FrameTelf.Visible = False
SSCommand1(0).Enabled = True
End Sub

Private Sub SSCommand6_Click(Index As Integer)
FrameUnif.Visible = True
FrameUnif.ZOrder 0
SSCommand6(6).Enabled = False
End Sub

'Private Sub SSCommand7_Click()
'FrameUnif.Visible = False
'SSCommand6.Enabled = True
'End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim rsValidar As ADODB.Recordset

Select Case SSTab1.Tab
    Case 2
        SSTab1.Tab = 0
        MsgBox "Módulo DESHABILITADO.", vbCritical, Me.Caption
        Exit Sub
Case 3 'REMUNERACIONES
    Sql$ = "select * from users where cod_cia='" & wcia & "' and login='" & wuser & "' and sistema='04' and status !='*'"
    If fAbrRst(rsValidar, Sql) Then
        If rsValidar!autorizar_remuneracion = False Then
            SSTab1.Tab = 0
            MsgBox "Personal No Autorizado, a consultar este tipo de información", vbCritical, Me.Caption
            Exit Sub
        End If
    Else
        SSTab1.Tab = 0
        MsgBox "Personal No Autorizado, a consultar este tipo de información", vbCritical, Me.Caption
        Exit Sub
    End If
    
Case 5 'CONTRATOS
    Sql$ = "select * from users where cod_cia='" & wcia & "' and login='" & wuser & "' and sistema='04' and status !='*'"
    If fAbrRst(rsValidar, Sql) Then
        If rsValidar!autorizar_contrato = False Then
            SSTab1.Tab = 0
            MsgBox "Personal No Autorizado, a consultar este tipo de información", vbCritical, Me.Caption
            Exit Sub
        End If
    Else
        SSTab1.Tab = 0
        MsgBox "Personal No Autorizado, a consultar este tipo de información", vbCritical, Me.Caption
        Exit Sub
    End If
    
Case 6
    Dim xId As String
    xId = fc_CodigoComboBox(cbo_cattrab, 2)
    If xId <> "03" Then
        SSTab1.Tab = 0
        MsgBox "Esta opcion es solo para la categoria trabajadores de 4ta categoria", vbExclamation, Me.Caption
        Exit Sub
    Else
        If rsSuspension4ta.RecordCount = 0 Then rsSuspension4ta.AddNew
    End If

End Select
End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtApeCas_Change()
'Busca_Auxiliar
End Sub

Private Sub TxtApeCas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtApeCas_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtPriNom.SetFocus
End Sub

Private Sub TxtApeMat_Change()
'Busca_Auxiliar
End Sub

Private Sub TxtApeMat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtApeMat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtApeCas.SetFocus
End Sub

Private Sub TxtApePat_Change()
'Busca_Auxiliar
End Sub

Private Sub TxtApePat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtApeMat.SetFocus
End Sub

Private Sub TxtCarta_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then CmbDocDh.SetFocus
End Sub

Private Sub txtAporteEmpleado_KeyPress(KeyAscii As Integer)

KeyAscii = fc_ValidarDecimal(KeyAscii, Me.txtAporteEmpleado.Text)
    
End Sub


Private Sub Txtcodobra_GotFocus()
wbus = "OB"
NameForm = "Frmpersona"
End Sub

Private Sub Txtcodobra_KeyPress(KeyAscii As Integer)
Txtcodobra.Text = Txtcodobra.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Txtcodobra.Text = Format(Txtcodobra.Text, "00000000"): TxtFecCese.SetFocus
End Sub
Private Sub Txtcodobra_LostFocus()
wbus = ""
If Txtcodobra.Text <> "" Then
   Sql$ = "select cod_obra,descrip,status from plaobras where cod_cia='" & wcia & "' and cod_obra='" & Txtcodobra.Text & "'"
   If (fAbrRst(Rs, Sql$)) Then
      If Rs!status = "*" Then
         MsgBox "Obra Eliminada", vbInformation, "Registro de Personal"
         Lbldesobra.Caption = ""
         Txtcodobra.SetFocus
      Else
         Lbldesobra.Caption = Trim(Rs!DESCRIP)
      End If
   Else
     MsgBox "Codigo de Obra no Registrada", vbInformation, "Registro de Personal"
     Lbldesobra.Caption = ""
     Txtcodobra.SetFocus
   End If
End If
End Sub

Private Sub Txtcodpla_Change()
If VNew = True Then
    txtPresentacion.Text = Txtcodpla.Text
End If
End Sub

Private Sub Txtcodpla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub Txtcodpla_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub TxtCtaCts_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtCtaCts_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub TxtCtaPag_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtCtaPag_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub TxtDni_KeyPress(KeyAscii As Integer)
TxtDni.Text = TxtDni.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then TxtLm.SetFocus
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtApePat.SetFocus
End Sub



Private Sub TxtFec_Afil_Sindicato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtFec_Afil_Sindicato_LostFocus()
If TxtFec_Afil_Sindicato.Text <> "__/__/____" Then
   If Not IsDate(TxtFec_Afil_Sindicato.Text) Then
      MsgBox "Ingrese Fecha Correctamente", vbCritical, TitMsg
      TxtFec_Afil_Sindicato.SetFocus
      Call ResaltarTexto(TxtFec_Afil_Sindicato)
   End If
End If
End Sub

Private Sub TxtFecAfilia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtFecAfilia_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And CMBZONA.Enabled Then CMBZONA.SetFocus

'If Cmbafp.ListIndex = -1 Or VCodAfp = "99" Then KeyAscii = 0
End Sub

Private Sub TxtFecAfilia_LostFocus()
If TxtFecAfilia.Text <> "__/__/____" Then
   If Not IsDate(TxtFecAfilia.Text) Then
      MsgBox "Ingrese Fecha Correctamente", vbCritical, TitMsg
      TxtFecNac.SetFocus
      Call ResaltarTexto(TxtFecAfilia)
   End If
End If
End Sub

Private Sub TxtFecCese_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtFecCese_LostFocus()
If TxtFecCese.Text <> "__/__/____" Then
   If Not IsDate(TxtFecCese.Text) Then
      MsgBox "Ingrese Fecha Correctamente", vbCritical, TitMsg
      TxtFecCese.SetFocus
      Call ResaltarTexto(TxtFecCese)
    Else
        If CDate(TxtFecCese.Text) < CDate(TxtFecIngreso.Text) Then
            MsgBox "La Fecha de Cese es Menor a la Fecha de Ingreso", vbCritical, TitMsg
            TxtFecCese.SetFocus
            Call ResaltarTexto(TxtFecCese)
        End If
   End If
End If
End Sub

Private Sub TxtFecIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtFecIngreso_LostFocus()
If TxtFecIngreso.Text <> "__/__/____" Then
   If Not IsDate(TxtFecIngreso.Text) Then
      MsgBox "Ingrese Fecha Correctamente", vbCritical, TitMsg
      TxtFecIngreso.SetFocus
      Call ResaltarTexto(TxtFecIngreso)
   End If
End If
End Sub

Private Sub Txtfecjubila_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtFecNac_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtFecNac_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then CmbCivil.SetFocus
End Sub

Private Sub TxtFecNac_LostFocus()
If TxtFecNac.Text <> "__/__/____" Then
   If Not IsDate(TxtFecNac.Text) Then
      MsgBox "Ingrese Fecha Correctamente", vbCritical, TitMsg
      TxtFecNac.SetFocus
      Call ResaltarTexto(TxtFecNac)
   End If
End If
End Sub

Private Sub txtint_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtInt_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtEmail.SetFocus
End Sub

Private Sub TxtIpss_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtIpss_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then Cmbafp.SetFocus
End Sub

Private Sub TxtLm_KeyPress(KeyAscii As Integer)
TxtLm.Text = TxtLm.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Txtpasaporte.SetFocus
End Sub

Private Sub txtnro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtNro_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then txtint.SetFocus
End Sub

Private Sub TxtNroAfp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtNroAfp_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtFecAfilia.SetFocus
'If Cmbafp.ListIndex = -1 Or VCodAfp = "99" Then KeyAscii = 0
End Sub

Private Sub Txtpasaporte_KeyPress(KeyAscii As Integer)
Txtpasaporte.Text = Txtpasaporte.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then Txtruc.SetFocus
End Sub

Private Sub TxtPriNom_Change()
'Busca_Auxiliar
End Sub

Private Sub TxtPriNom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtPriNom_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtSegNom.SetFocus
End Sub

Private Sub Txtruc_KeyPress(KeyAscii As Integer)
Txtruc.Text = Txtruc.Text + fc_ValNumeros(KeyAscii)
If KeyAscii = 13 Then TxtIpss.SetFocus
End Sub

Private Sub TxtRef_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtRef_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtSegNom_Change()
'Busca_Auxiliar
End Sub

Private Sub TxtSegNom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub TxtSegNom_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then TxtFecNac.SetFocus
End Sub

Private Sub TxtVia_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then txtnro.SetFocus
End Sub

Private Sub TxtZona_KeyPress(KeyAscii As Integer)
If KeyAscii = No_Apostrofe(KeyAscii) = 0 Then KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then CmbVia.SetFocus
End Sub
Public Sub Grabar_Persona()
'On Error GoTo ErrorTrans
Dim MSexo As String
Dim MSINDICATO As String
Dim MSctr As String
Dim MQuinta As String
Dim MVida As String
Dim Mcodtrab As String
Dim NumAuxi As Integer
Dim MimpDeduc As Currency
Dim MStatusDed As String
Dim mmonto As Currency
Dim mhoras As Integer
Dim mafp As String
Dim mescolar As String
Dim maltitud As String
Dim mvacacion As String
Dim xPerRem As String
Dim xsegcom As String
Dim VTipoTrabSunat As String
Dim NroTrans As Integer
Dim sAFPTipoComision As String

NroTrans = 0

If Me.optAFPMixta.Value = True Then
    sAFPTipoComision = "M"
Else
    sAFPTipoComision = "F"
End If
VTipoTrabSunat = ""
If TipoTrabSunat = "S" And CmbTrabSunat.ListIndex >= 0 Then VTipoTrabSunat = fc_CodigoComboBox(CmbTrabSunat, 2)

xsegcom = fc_CodigoComboBox(CmbSegCom, 2)
xPerRem = fc_CodigoComboBox(CmbPerRem, 2)

If Trim(xsegcom & "") = "" Then
   SSTab1.Tab = 1
   MsgBox "Debe Indicar Seguro Compl. Trabajador Riesgo", vbInformation
   Screen.MousePointer = vbDefault
   CmbSegCom.SetFocus
   Exit Sub
End If

Dim xId As String
xId = fc_CodigoComboBox(cbo_cattrab, 2)

If ChkAltitud.Value = 1 Then maltitud = "S"
If ChkVaca.Value = 1 Then mvacacion = "S"
If VCodAfp = "99" Then mafp = "" Else mafp = VCodAfp
Screen.MousePointer = vbHourglass
If Not Validar Then Screen.MousePointer = vbDefault: Exit Sub

If xId = "03" Then 'suspension de 4ta
    If Not Validar_Suspension4taCategoria Then Screen.MousePointer = vbDefault: Exit Sub
End If

If Trim(TxtFecCese.Text & "") <> "__/__/____" And Not IsDate(TxtFecCese.Text) Then
   MsgBox "Ingrese Fecha de Cese Correctamente", vbInformation
   Screen.MousePointer = vbDefault
   Exit Sub
End If
If Trim(TxtFecCese.Text & "") = "__/__/____" Then
   If cbo_situacioneps.ReturnCodigo = 6 Or cbo_situacioneps.ReturnCodigo = 2 Then
      SSTab1.Tab = 4
      MsgBox "Situación de EPS no puede ser BAJA", vbInformation
      Screen.MousePointer = vbDefault
      cbo_situacioneps.SetFocus
      Exit Sub
   End If
End If

If (CmbResponsable.Text & "") = "" Then
   SSTab1.Tab = 1
   MsgBox "Debe Indicar Responsable", vbInformation
   Screen.MousePointer = vbDefault
   CmbResponsable.SetFocus
   Exit Sub
End If

If CmbPlanta.Text = "CANTERA" And Trim(LblCodCantera.Caption & "") = "" Then
   SSTab1.Tab = 1
   MsgBox "Debe Indicar Cantera", vbInformation
   Screen.MousePointer = vbDefault
   CmbPlanta.SetFocus
   Exit Sub
End If

If MsgBox("Desea Grabar Personal ", vbYesNo + vbQuestion) = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
If OpcVaron.Value = True Then MSexo = "M"
If OpcDama.Value = True Then MSexo = "F"
'If OpcSctrSi.Value = True Then MSctr = "S" Else MSctr = ""
If OpcVidaSi.Value = True Then MVida = "S" Else MVida = ""
If ChkSindicato.Value = 1 Then MSINDICATO = "S" Else MSINDICATO = ""
If ChkNoQuinta.Value = 1 Then MQuinta = "N" Else MQuinta = "S"
Mcodtrab = UCase(Trim(Txtcodpla.Text))

Dim xEpsServ As String
Dim xreg_alternativo As String
If Me.OptEpsSi(0).Value = True Then xEpsServ = "1" Else xEpsServ = "0"
Dim xSituacionTrabajo As String
If OptSitEsp_direccion(0).Value Then
    xSituacionTrabajo = "1"
ElseIf OptSitEsp_confianza(1).Value Then
    xSituacionTrabajo = "2"
ElseIf OptSitEsp_ninguna(2).Value Then
    xSituacionTrabajo = "0"
End If

'/*ADD JCMS 04/08/08 */
Dim xAfiliacion_Asegura_TuPension As String
If Me.OptAsegPenNo.Value = True Then xAfiliacion_Asegura_TuPension = "0"
If Me.OptAsegPenSi.Value = True Then xAfiliacion_Asegura_TuPension = "1"
Dim xEvita_Doble_Tributacion As String
xEvita_Doble_Tributacion = fc_CodigoComboBox(Me.CmbEvitaDobleImp, 2)

'Cambiar Situcacion de Trabajador cuando Cesa (Situacuin EPS)
If IsDate(TxtFecCese.Text) Then
   Sql$ = "select afiliado_eps_serv from planillas where cia='" & wcia & "' and placod='" & Mcodtrab & "' and status<>'*'"
    cn.CursorLocation = adUseClient
    Set Rs = New ADODB.Recordset
    Set Rs = cn.Execute(Sql$)
    If Rs.RecordCount > 0 Then
       If Rs(0) Then
          cbo_situacioneps.SetIndice "02"
       Else
          cbo_situacioneps.SetIndice "06"
       End If
    End If
    Rs.Close
End If
txtPresentacion.Text = UCase(Trim(txtPresentacion.Text))
If Trim(txtPresentacion.Text) = "" Then
    MsgBox "Ingrese el código de presentación del Trabajador", vbCritical, Me.Caption
    Screen.MousePointer = vbDefault
    txtPresentacion.SetFocus
    Exit Sub
Else
    Sql$ = "select afiliado_eps_serv from planillas where cia='" & wcia & _
           "' and placod !='" & Mcodtrab & "' and placodpresentacion='" & _
           Trim(txtPresentacion.Text) & "' and status<>'*'"
    cn.CursorLocation = adUseClient
    Set Rs = New ADODB.Recordset
    Set Rs = cn.Execute(Sql$)
    If Rs.RecordCount > 0 Then
        MsgBox "El código de presentación está siendo utilizado por el Trabajador: " & Trim(Rs!PlaCod), vbCritical, Me.Caption
        Screen.MousePointer = vbDefault
        txtPresentacion.SetFocus
        Exit Sub
    End If

End If

cn.BeginTrans
NroTrans = 1
If VNew = True Then
   Sql$ = "Select numero from numeracion where documento='AUX'"
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$)
   
   'GENERA NUMERACION
   If Rs.RecordCount > 0 Then
      VCodAuxi = Format(Rs!numero + 1, "00000000")
      NumAuxi = Val(VCodAuxi)
      Sql$ = "Update numeracion set numero=" & NumAuxi & " where documento='AUX'"
      cn.Execute Sql$
   Else
      VCodAuxi = "00000001"
      NumAuxi = Val(VCodAuxi)
      Sql$ = "Insert into numeracion values('" & wcia & "','AUX','CORRELATIVO DE AUXILIARES'," & NumAuxi & ",0)"
      cn.Execute Sql$
   End If
  
  If Rs.State = 1 Then Rs.Close
End If

Sql$ = "Update planillas set status='*' where cia='" & wcia & "' and placod='" & Mcodtrab & "' and codauxinterno='" & VCodAuxi & "' and status<>'*'"
cn.Execute Sql$
'PLAME
Dim VArea As String
VArea = ""
Dim mCosto As Double
mCosto = 0
If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
Do While Not rsccosto.EOF
   If Trim(rsccosto!Codigo & "") <> "" And IsNumeric(rsccosto!Monto) Then
      If rsccosto!Monto > mCosto Then mCosto = rsccosto!Monto: VArea = Trim(rsccosto!Codigo & "")
   End If
   rsccosto.MoveNext
Loop

VCentroForma = Trim(Left(CmbCentroForma, 2) & "")
Dim lCantera As String
lCantera = ""
If CmbPlanta.Text = "CANTERA" Then VPlanta = "00": lCantera = LblCodCantera.Caption

Sql$ = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "INSERT INTO Planillas VALUES('" & wcia & "','" & Mcodtrab & "','" & VCodAuxi & "','" & Trim(TxtApePat.Text) & "','" & Trim(TxtApeMat.Text) & "','" & Trim(TxtApeCas.Text) & "', " _
   & "'" & Trim(TxtPriNom.Text) & "','" & Trim(TxtSegNom.Text) & "','" & IIf(cbo_naciontrab.ReturnCodigo = -1, "", cbo_naciontrab.ReturnCodigo) & "',"
   Sql$ = Sql$ & IIf(IsDate(TxtFecNac.Text), "'" & Format(TxtFecNac, FormatFecha) & "'", "Null")
   Sql$ = Sql$ & ",'" & VCivil & "', " _
   & "'" & MSexo & "','" & Trim("") & "','" & Trim("") & "','" & Trim("") & "','" & Trim("") & "', " _
   & "'" & Trim(TxtIpss.Text) & "','" & Trim(IIf(cbo_pensiones.ReturnCodigo = -1, "", Format(cbo_pensiones.ReturnCodigo, "00"))) & "','" & Trim(TxtNroAfp.Text) & "',"
   Sql$ = Sql$ & IIf(IsDate(TxtFecAfilia.Text), "'" & Format(TxtFecAfilia, FormatFecha) & "'", "Null")
   Sql$ = Sql$ & ",'" & IIf(cbo_zonatrab.ReturnCodigo = -1, "", Format(cbo_zonatrab.ReturnCodigo, "00")) & "', " _
   & "'" & Trim(Text9.Text) & "','" & IIf(cbo_viatrab.ReturnCodigo = -1, "", Format(cbo_viatrab.ReturnCodigo, "00")) & "','" & Trim(Text10.Text) & "','" & Trim(txtnro.Text) & "','" & Trim(txtint.Text) & "','" & Trim(Text13.Tag) & "','" & Trim(TxtEmail.Text) & "', " _
   & "'" & VNivelEdu & "','" & VCentroForma & "','" & Trim(txtprofesion.Tag) & "','" & VCargo & "','" & VPlanta & "','" & VArea & "',"
   Sql$ = Sql$ & IIf(IsDate(TxtFecIngreso.Text), "'" & Format(TxtFecIngreso.Text, FormatFecha) & "'", "Null")
   Sql$ = Sql$ & ",'" & VTipotrab & "','',"
   Sql$ = Sql$ & IIf(IsDate(TxtFecCese.Text), "'" & Format(TxtFecCese, FormatFecha) & "'", "Null")
   Sql$ = Sql$ & ",'" & MSINDICATO & "','" & VTipoHorasExt & "','" & MSctr & "','" & VDistribCalcSeg & "','" & _
   MQuinta & "','" & MVida & "','" & VTipoPago & "', " _
   & "'" & VBcoPago & "','" & VMonPago & "','" & VTipoCtaPago & "','" & Trim(TxtCtaPag.Text) & "','" & _
   VBcoCts & "','" & VMonCts & "','" & VTipoCtaCts & "','" & Trim(TxtCtaCts.Text) & "','" & wuser & "'," & FechaSys & ",'" & wuser & "','',"
   Sql$ = Sql$ & IIf(IsDate(Txtfecjubila.Text), "'" & Format(Txtfecjubila, FormatFecha) & "'", "Null")
   Sql$ = Sql$ & ",'" & Trim(Txtcodobra.Text) & "','" & VMonCalculo & "','" & maltitud & "','" & mvacacion & _
   "','','" & IIf(chkJubilado.Value = 1, "S", "") & "','" & TipoDoc & "','" & Trim(Text8.Text) & "','" & IIf(cbo_cattrab.ReturnCodigo = -1, "", Format(cbo_cattrab.ReturnCodigo, "00")) & "','','',"
   Sql$ = Sql$ & "'" & Trim(Text7.Text) & "','" & IIf(cbo_situacioneps.ReturnCodigo = -1, "", Format(cbo_situacioneps.ReturnCodigo, "00")) & "','" & IIf(cbo_eps.ReturnCodigo = -1, "", Format(cbo_eps.ReturnCodigo, "00")) & "','" & IIf(cbopension.ReturnCodigo = -1, "", Format(cbopension.ReturnCodigo, "00")) & "','" & ChkDiscapacidad.Value & "','" & IIf(cbo_tipocont.ReturnCodigo = -1, "", Format(cbo_tipocont.ReturnCodigo, "00")) & "'," & _
   "'" & IIf(Check1.Value = vbChecked, "1", "0") & "','" & IIf(cboessalud.ReturnCodigo = -1, "", Format(cboessalud.ReturnCodigo, "00")) & "','" & Trim(Txt_Sucursal.Text) & "'"
   Sql$ = Sql$ & ",'" & IIf(cbo_TipModFor.ReturnCodigo = -1, "", Format(cbo_TipModFor.ReturnCodigo, "00")) & "','" & IIf(cbo_TipMotFinPer.ReturnCodigo = -1, "", Format(cbo_TipMotFinPer.ReturnCodigo, "00")) & "'"
   Sql$ = Sql$ & ",'" & Apostrofe(Trim(TxtRef.Text)) & "','" & IIf(ChkDomiciliado.Value = 1, "1", "0") & "','" & Apostrofe(Trim(TxtRefAnt.Text)) & "'"
   Sql$ = Sql$ & ",'" & xEpsServ & "','" & ChkRegimen_Alternativo(0).Value & "','" & ChkJornada_max(1).Value & "','" & ChkHorario_nocturno(2).Value & "','" & ChkOtrosIngta(3).Value & "','" & Chk5taExonerada(4).Value & "','" & xSituacionTrabajo & "','" & xPerRem & "','" & ChkMadreResp(1).Value & "','" & ChkCalcula_AccidenteTrabajo(1).Value & "'"
   Sql$ = Sql$ & ",'" & xAfiliacion_Asegura_TuPension & "','" & xEvita_Doble_Tributacion & "','" & Trim(lblPlaCos.Caption) & "','" & xsegcom & "'," & chk_integro.Value & ""
   Sql$ = Sql$ & ",'" & Trim(TxtNroDpto1.Text) & "','" & Trim(TxtNroMz1.Text) & "','" & Trim(TxtNroLote1.Text) & "','" & Trim(TxtNroKM1.Text) & "','" & Trim(TxtNroBlock1.Text) & "','" & Trim(TXtNroEtapa1.Text) & "','" & VTipoTrabSunat & "'," & Val(txtAporteEmpleado.Text) & ",'" & Trim(txtPresentacion.Text) & "','" & sAFPTipoComision & "','" & Trim(TxtCCI.Text & "") & "','" & Trim(TxtCCICts.Text) & "'"
   
   Dim lPubPriv As Integer
   lPubPriv = 0
   If OptPublica.Value = True Then lPubPriv = 1
   If OptPrivada.Value = True Then lPubPriv = 2
   Sql$ = Sql$ & "," & lPubPriv & " "
   If ChkPais.Value = 1 Then Sql$ = Sql$ & ",1" Else Sql$ = Sql$ & ",0"
   Sql$ = Sql$ & ",'" & TxtNombreUni.Tag & "','" & TxtCarrera.Tag & "'"
   If IsNumeric(TxtAnoEgreso.Text) Then Sql$ = Sql$ & "," & TxtAnoEgreso.Text & "," Else Sql$ = Sql$ & ",0,"
   Sql$ = Sql$ & "'" & fc_CodigoComboBox(CmbResponsable, 3) & "','" & lCantera & "','" & Vpla_area & "'," & IIf(IsDate(TxtFec_Afil_Sindicato.Text), "'" & Format(TxtFec_Afil_Sindicato.Text, FormatFecha) & "'", "Null") & ")"
   cn.Execute Sql$
   
If VNew = True Then
   Sql$ = "insert into maestroaux values('" & VCodAuxi & "','','" & Trim(TxtApePat.Text) & "','" & Trim(TxtApeMat.Text) & "','" & Trim(TxtApeCas.Text) & "','" & Trim(TxtPriNom.Text) & "','" & Trim(TxtSegNom.Text) & "','','','X','')"
   cn.Execute Sql$
Else
   Sql$ = "update maestroaux set ap_pat='" & Trim(TxtApePat.Text) & "',ap_mat='" & Trim(TxtApeMat.Text) & "',ap_cas='" & Trim(TxtApeCas.Text) & "',nom_1='" & Trim(TxtPriNom.Text) & "',nom_2='" & Trim(TxtSegNom.Text) & "' where codauxinterno='" & VCodAuxi & "' and status<>'*'"
   cn.Execute Sql$
End If

Sql$ = "Update platelefono set status='*' where cia='" & wcia & "' and placod='" & Mcodtrab & "' and codauxinterno='" & VCodAuxi & "' and status<>'*'"
cn.Execute Sql$

If rstlf.RecordCount > 0 Then rstlf.MoveFirst
Do While Not rstlf.EOF
   Sql$ = "Insert into platelefono values('" & wcia & "','" & Mcodtrab & "','" & Trim(VCodAuxi) & "','" & Trim(rstlf!referencia) & "','" & Trim(rstlf!telefono) & "','" & Trim(rstlf!fax) & "','')"
   cn.Execute Sql$
   rstlf.MoveNext
Loop

Sql$ = "Update TBL_BCO_CUENTA_DJ set status='*' where placod='" & Mcodtrab & "' and status<'*'"
cn.Execute Sql$
'Insereta Judicial
If rsJud.RecordCount > 0 Then rsJud.MoveFirst
Do While Not rsJud.EOF
   'Sql$ = "Insert into TBL_BCO_CUENTA_DJ values('" & Mcodtrab & "','" & Trim(rsJud!Nrodni) & "','" & Trim(rsJud!nombre) & "','" & Trim(rsJud!Tipocta) & "','" & Trim(rsJud!Bco) & "','" & Trim(rsJud!NroCta) & "'," & Trim(rsJud!Porcentaje) & " ,'')"
   Sql$ = "Insert into TBL_BCO_CUENTA_DJ ([PLACOD],[NRODNI],[NOMBRE],[TIPOCTA],[BCO],[NROCTA],[PORCENTAJE],[STATUS],[IndicadorCalculoEnBaseIngresoTotalBruto])   values('" & Mcodtrab & "','" & Trim(rsJud!Nrodni) & "','" & Trim(rsJud!nombre) & "','" & Trim(rsJud!Tipocta) & "','" & Trim(rsJud!Bco) & "','" & Trim(rsJud!NroCta) & "'," & Trim(rsJud!Porcentaje) & " ,''," & IIf(Trim(rsJud!CalcBaseBruto) = "S", "1", "0") & ")"
   cn.Execute Sql$
   rsJud.MoveNext
Loop

'Sql$ = "Update plaUniforme set status='*' where cia='" & wcia & "' and placod='" & Mcodtrab & "' and codauxinterno='" & VCodAuxi & "' and status<>'*'"
Sql$ = "Update plaUniformes set status='*' where cia='" & wcia & "' and placod='" & Mcodtrab & "' and status<>'*'"
cn.Execute Sql$

If rsunif.RecordCount > 0 Then rsunif.MoveFirst
Do While Not rsunif.EOF
  ' Sql$ = "Insert into plaUniforme values('" & wcia & "','" & Mcodtrab & "','" & Trim(VCodAuxi) & "','" & Trim(rsunif!botas) & "','" & Trim(rsunif!camisa) & "','" & Trim(rsunif!pantalon) & "','" & Trim(rsunif!polo) & "','')"
    Sql$ = "Insert into plaUniformes values('" & wcia & "','" & Mcodtrab & "','" & Trim(VCodAuxi) & "','" & Trim(rsunif!referencia) & "','" & Trim(rsunif!Talla) & "','')"
   cn.Execute Sql$
   rsunif.MoveNext
Loop

Uniformes_1

Sql$ = "Update pladeducper set status='*' where cia='" & wcia & "' and placod='" & Mcodtrab & "' and status<>'*'"
cn.Execute Sql$

If rsdeduccion.RecordCount > 0 Then rsdeduccion.MoveFirst
Do While Not rsdeduccion.EOF
   If IsNull(rsdeduccion!importe) Then mmonto = 0 Else mmonto = rsdeduccion!importe
   If CCur(mmonto) <> 0 Then
      MimpDeduc = mmonto
      MStatusDed = "F"
   Else
      If IsNull(rsdeduccion!Porcentaje) Then mmonto = 0 Else mmonto = rsdeduccion!Porcentaje
      MimpDeduc = mmonto
      MStatusDed = "P"
   End If
   If CCur(MimpDeduc) <> 0 Then
      Sql$ = "Insert into pladeducper values('" & wcia & "','" & Mcodtrab & "','" & Trim(VCodAuxi) & "','" & Trim(rsdeduccion!Codigo) & "'," & CCur(MimpDeduc) & ",'" & MStatusDed & "')"
      cn.Execute Sql$
   End If
   rsdeduccion.MoveNext
Loop
'Retencion Judicioal en boletas de utilidades
Sql$ = "Update PlaRetJudUti set status='*',user_crea='" & wuser & "',fec_modi=getdate() where placod='" & Mcodtrab & "' and status<>'*'"
cn.Execute Sql$
If ChkRetJudUti.Value = 1 Then
   Sql$ = "Insert Into PlaRetJudUti Values('" & wcia & "','" & Mcodtrab & "','','" & wuser & "',getdate(),'" & wuser & "',getdate())"
   cn.Execute Sql$
End If

Sql$ = "Update plaremunbase set status='*' where cia='" & wcia & "' and placod='" & Mcodtrab & "' and codauxinterno='" & VCodAuxi & "' and status<>'*'"
cn.Execute Sql$

If rsbasico.RecordCount > 0 Then rsbasico.MoveFirst
Do While Not rsbasico.EOF
   If IsNull(rsbasico!importe) Then mmonto = 0 Else mmonto = rsbasico!importe
   If CCur(mmonto) <> 0 Then
      If IsNull(rsbasico!horas) Then mhoras = 0 Else mhoras = rsbasico!horas
      Sql$ = "Insert into plaremunbase values('" & wcia & "','" & Mcodtrab & "','" & Trim(VCodAuxi) & "','" & Trim(rsbasico!Codigo) & "','" & rsbasico!moneda & "'," & CCur(mmonto) & ",'" & rsbasico!codtipo & "'," & mhoras & ",'')"
      cn.Execute Sql$
   End If
   rsbasico.MoveNext
Loop


Sql$ = "Update planilla_ccosto set status='*' where cia='" & wcia & "' and placod='" & Mcodtrab & "' and status<>'*'"
cn.Execute Sql$

If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
Do While Not rsccosto.EOF
   If Trim(rsccosto!Codigo & "") <> "" And IsNumeric(rsccosto!Monto) Then
      If CCur(rsccosto!Monto) <> 0 Then
         Sql$ = "Insert into planilla_ccosto values('" & wcia & "','" & Mcodtrab & "','" & Trim(rsccosto!Codigo) & "'," & rsccosto!Monto & ",'')"
         cn.Execute Sql$
      End If
   End If
   rsccosto.MoveNext
Loop


'///***** Otros Empleadores *****/////
If ChkOtrosIngta(3).Value = 1 Then
    Sql$ = "Update plaOtrosEmpleadores set status='*',user_modi='" & wuser & "',fec_modi=getdate() where cia='" & wcia & "' and placod='" & Mcodtrab & "' and status<>'*'"
    cn.Execute Sql$
    
    If rsOtrosEmpleadores.RecordCount > 0 Then rsOtrosEmpleadores.MoveFirst
    Do While Not rsOtrosEmpleadores.EOF
          If Trim(rsOtrosEmpleadores!RUC & "") <> "" Then
            Sql$ = "Insert into plaOtrosEmpleadores values('" & wcia & "','" & Mcodtrab & "','" & Trim(rsOtrosEmpleadores!RUC) & "','" & UCase(Apostrofe(rsOtrosEmpleadores!razsoc)) & "','','" & Trim(wuser) & "',getdate(),null,null)"
            cn.Execute Sql$, 64
         End If
       rsOtrosEmpleadores.MoveNext
    Loop
End If



'///***** suspension 4ta categoria *****/////

Sql$ = "Update tsuspension set status='*' where cod_cia='" & wcia & "' and cod_prov='" & Mcodtrab & "' and status<>'*'"
cn.Execute Sql$

If rsSuspension4ta.RecordCount > 0 Then rsSuspension4ta.MoveFirst
Do While Not rsSuspension4ta.EOF
      If Trim(rsSuspension4ta!numop & "") <> "" Then
        Sql$ = "Insert into tsuspension values('" & wcia & "','" & Mcodtrab & "','" & Format(Trim(rsSuspension4ta!fecha), "mm/dd/yyyy")
        Sql = Sql & "','" & Trim(rsSuspension4ta!numop) & "','" & Trim(rsSuspension4ta!ejercicio) & "','" & Trim(rsSuspension4ta!medio) & "',getdate(),'" & Trim(wuser) & "','')"
        cn.Execute Sql$, 64
     End If
   rsSuspension4ta.MoveNext
Loop


'///***** Establecimientos donde labora el trabajador *****/////

Sql$ = "Update plaCiaEstablecimientos_Labora_Trabajador set status='*',user_modi='" & wuser & "',fec_modi=getdate() where cod_cia='" & wcia & "' and placod='" & Mcodtrab & "' and status<>'*'"
cn.Execute Sql$

With rsTipEst
    If .RecordCount > 0 Then .MoveFirst
    Do While Not .EOF
            If !Add = True Then
                Sql = "Insert into plaCiaEstablecimientos_Labora_Trabajador "
                Sql = Sql & " values('" & wcia & "','" & Mcodtrab & "','"
                Sql = Sql & Trim(!RUC) & "','" & Trim(!codest) & "','','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql$, 64
            End If
       .MoveNext
    Loop
End With

cn.CommitTrans

Screen.MousePointer = vbDefault
If VNew = True Then
    Nuevo_Personal (True)
Else
    Unload Me
End If
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox Err.Description, vbCritical, Me.Caption

End Sub

Private Function Validar()
'ON ERROR GOTO CORRIGE
Dim xIdCategoriaTrabahador As String
xIdCategoriaTrabahador = fc_CodigoComboBox(cbo_cattrab, 2)

If Not Verifica_permiso Then Exit Function

If Trim(TxtApePat.Text) = "" Then SSTab1.Tab = 0: MsgBox "Debe Ingresar Primer Apellido", vbCritical, TitMsg: Validar = False: TxtApePat.SetFocus: Exit Function
If Trim(TxtPriNom.Text) = "" Then SSTab1.Tab = 0: MsgBox "Debe Ingresar Primer Nombre", vbCritical, TitMsg: Validar = False: TxtPriNom.SetFocus: Exit Function
If Not IsDate(TxtFecNac.Text) Then SSTab1.Tab = 0: MsgBox "Ingrese Fecha De Nacimiento ", vbCritical, TitMsg: Validar = False: TxtFecNac.SetFocus: Exit Function
If Trim(cbo_naciontrab.Text) = "NINGUNO" Or Trim(cbo_naciontrab.Text) = "" Or cbo_naciontrab.ListIndex = -1 Then SSTab1.Tab = 0: MsgBox "Elija Nacionalidad del Trabajador", vbCritical, TitMsg: Validar = False: cbo_naciontrab.SetFocus: Exit Function

If Cbotipodoc.ListIndex = -1 Then SSTab1.Tab = 0: MsgBox "Elija Tipo de Documento del Trabajador", vbCritical, TitMsg: Validar = False: Cbotipodoc.SetFocus: Exit Function
If Trim(Text8.Text) = "" Then SSTab1.Tab = 0: MsgBox "Ingrese Nro de Documento del Trabajador", vbCritical, TitMsg: Validar = False: Text8.SetFocus: Exit Function

If Trim(Text9.Text) = "" Then SSTab1.Tab = 0: MsgBox "Debe Ingresar Zona", vbCritical, TitMsg: Validar = False: Text9.SetFocus: Exit Function
If Trim(Text13.Tag) = "" Then SSTab1.Tab = 0: MsgBox "Debe Ingresar Ubicacion", vbCritical, TitMsg: Validar = False:  Exit Function
If Trim(cbo_cattrab.Text) = "NINGUNO" Or Trim(cbo_cattrab.Text) = "" Or cbo_cattrab.ListIndex = -1 Then SSTab1.Tab = 0: MsgBox "Debe Seleccionar Categoria del Trabajador", vbCritical, TitMsg: Validar = False: cbo_cattrab.SetFocus: Exit Function


If OpcVaron.Value = False And OpcDama.Value = False Then SSTab1.Tab = 0: MsgBox "Debe Indicar Sexo", vbCritical, TitMsg: Validar = False: Exit Function
If Trim(CmbTipoTrabajador.Text) = "NINGUNO" Or Trim(CmbTipoTrabajador.Text) = "" Then SSTab1.Tab = 1: MsgBox "Debe Seleccionar Tipo de Trabajador", vbCritical, TitMsg: Validar = False: CmbTipoTrabajador.SetFocus: Exit Function

If Trim(Txtcodpla.Text) = "" Then SSTab1.Tab = 1: MsgBox "Ingrese Codigo de Trabajador", vbCritical, TitMsg: Validar = False: Txtcodpla.SetFocus: Exit Function
If xIdCategoriaTrabahador <> "04" And Not IsDate(TxtFecIngreso.Text) Then SSTab1.Tab = 1: MsgBox "Ingrese Fecha De Ingreso ", vbCritical, TitMsg: Validar = False: TxtFecIngreso.SetFocus: Exit Function

If Trim(CmbCargo.Text) = "" Then SSTab1.Tab = 1: MsgBox "Debe Seleccionar Cargo", vbCritical, TitMsg: Validar = False: CmbCargo.SetFocus: Exit Function
If Trim(CmbArea.Text) = "" Then SSTab1.Tab = 1: MsgBox "Debe Seleccionar Cargo", vbCritical, TitMsg: Validar = False: CmbArea.SetFocus: Exit Function

'If Trim(CmbArea.Text) = "" Then SSTab1.Tab = 1: MsgBox "Debe Seleccionar Area", vbCritical, TitMsg: Validar = False: CmbArea.SetFocus: Exit Function
Dim mCCosto As Double
mCCosto = 0
If rsccosto.RecordCount > 0 Then rsccosto.MoveFirst
Do While Not rsccosto.EOF
   If Trim(rsccosto!Codigo & "") <> "" Then
      If Trim(rsccosto!Monto & "") = "" Then SSTab1.Tab = 1: MsgBox "Debe Indicar Procentaje para el Centro de Costo", vbCritical, TitMsg: Validar = False: Dgrdccosto.SetFocus: Exit Function
      If Not IsNumeric(Trim(rsccosto!Monto & "")) Then SSTab1.Tab = 1: MsgBox "Debe Indicar Procentaje Correcto para el Centro de Costo", vbCritical, TitMsg: Validar = False: Dgrdccosto.SetFocus: Exit Function
      If Val(Trim(rsccosto!Monto & "")) < 0 Then SSTab1.Tab = 1: MsgBox "Debe Indicar Procentaje Correcto para el Centro de Costo", vbCritical, TitMsg: Validar = False: Dgrdccosto.SetFocus: Exit Function
      mCCosto = mCCosto + rsccosto!Monto
   End If
   rsccosto.MoveNext
Loop
If mCCosto <> 100 Then SSTab1.Tab = 1: MsgBox "Porcentaje de Centro de Costo debe sumar 100", vbCritical, TitMsg: Validar = False: Dgrdccosto.SetFocus: Exit Function

Dim lanos As Integer
lanos = DateDiff("yyyy", TxtFecNac.Text, TxtFecIngreso.Text)
If CDate(TxtFecIngreso.Text) < CDate(DateSerial(Year(TxtFecIngreso.Text), Month(TxtFecNac.Text), Day(TxtFecNac.Text))) Then lanos = lanos - 1
If lanos < 18 Then MsgBox "Trabajador Debe ser Mayor de 18 años", vbCritical, TitMsg: Validar = False: Text8.SetFocus: Exit Function

Select Case xIdCategoriaTrabahador
Case "05" 'PRESTADOR DE SERVICIOS - MODALIDAD FORMATIVA
    If OpcVidaSi.Value = False And OpcVidaNo.Value = False Then
        SSTab1.Tab = 1
        MsgBox "Elija Tipo de seguro, correctamente", vbExclamation, Me.Caption
        OpcVidaSi.SetFocus
        Validar = False: Exit Function
    ElseIf CmbNivelEdu.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Nivel Educativo de la lista, correctamente", vbExclamation, Me.Caption
        CmbNivelEdu.SetFocus
        Validar = False: Exit Function
    ElseIf Trim(txtprofesion.Text) = "" Then
        SSTab1.Tab = 1
        MsgBox "Elija Ocupacion y/o Profesion de la lista, correctamente", vbExclamation, Me.Caption
        'Label1(1).SetFocus
        Validar = False: Exit Function
    ElseIf CmbCentroForma.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Centro de Formación de la lista, correctamente", vbExclamation, Me.Caption
        CmbCentroForma.SetFocus
        Validar = False: Exit Function
    ElseIf Trim(cbo_TipModFor.Text) = "NINGUNO" Or Trim(cbo_TipModFor.Text) = "" Or cbo_TipModFor.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Modalidad Formativa de la lista, correctamente", vbExclamation, Me.Caption
        cbo_TipModFor.SetFocus
        Validar = False: Exit Function
    End If
    
Case "03" 'consistencia PRESTADOR DE SERVICIOS - 4TA CATEGORIA
    
    If (Trim(Text7.Text) = "" Or Len(Trim(Text7.Text)) < 11) Then
        SSTab1.Tab = 1
        MsgBox "Ingrese Nro de RUC, correctamente", vbExclamation, Me.Caption
        Text7.SetFocus
        Validar = False: Exit Function
    End If
Case "01", "02"

    '//** CONSISTENCIA SUNAT**///
    'If optestrabpen.Value = True And cbo_pension.Text = "NINGUNO" Then
    '    SSTab1.Tab = 0
    '    MsgBox "Elija tipo de Pensión de la lista, correctamente", vbExclamation, Me.Caption
    '    cbo_pension.SetFocus
    '    VALIDAR = False: Exit Function
    If OptEpsSi(0).Value = True And (Trim(cbo_eps.Text) = "NINGUNO" Or Trim(cbo_eps.Text) = "" Or cbo_eps.ListIndex = -1) Then
        SSTab1.Tab = 4
        MsgBox "Elija tipo de EPS de la lista, correctamente", vbExclamation, Me.Caption
        cbo_eps.SetFocus
        Validar = False: Exit Function
    ElseIf cbo_eps.Enabled And (Trim(cbo_situacioneps.Text) = "NINGUNO" Or Trim(cbo_situacioneps.Text) = "" Or cbo_situacioneps.ListIndex = -1) Then
        SSTab1.Tab = 4
        MsgBox "Elija tipo Situación EPS de la lista, correctamente", vbExclamation, Me.Caption
        cbo_situacioneps.SetFocus
        Validar = False: Exit Function
    ElseIf Trim(cbo_pensiones.Text) = "NINGUNO" Or Trim(cbo_pensiones.Text) = "" Or cbo_pensiones.ListIndex = -1 Then
        SSTab1.Tab = 4
        MsgBox "Elija Régimen de pensiones de la lista, correctamente", vbExclamation, Me.Caption
        cbo_pensiones.SetFocus
        Validar = False: Exit Function
    ElseIf CmbNivelEdu.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Nivel Educativo de la lista, correctamente", vbExclamation, Me.Caption
        CmbNivelEdu.SetFocus
        Validar = False: Exit Function
    'ElseIf Trim(cbo_TipModFor.Text) = "NINGUNO" Or Trim(cbo_TipModFor.Text) = "" Or cbo_TipModFor.ListIndex = -1 Then
    '    SSTab1.Tab = 1
    '    MsgBox "Elija Modalidad Formativa de la lista, correctamente", vbExclamation, Me.Caption
    '    cbo_TipModFor.SetFocus
    '    VALIDAR = False: Exit Function
    ElseIf CmbTipoTrabajador.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Tipo de Trabajador de la lista, correctamente", vbExclamation, Me.Caption
        CmbTipoTrabajador.SetFocus
        Validar = False: Exit Function
    ElseIf Trim(cbo_tipocont.Text) = "NINGUNO" Or Trim(cbo_tipocont.Text) = "" Or cbo_tipocont.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Tipo de  Contrato del trabajador de la lista, correctamente", vbExclamation, Me.Caption
        cbo_tipocont.SetFocus
        Validar = False: Exit Function
    ElseIf IsDate(TxtFecCese) And cbo_TipMotFinPer.ListIndex = -1 Then
        SSTab1.Tab = 1
        MsgBox "Elija Tipo Motivo Fin de Periodo de la lista, correctamente", vbExclamation, Me.Caption
        cbo_TipMotFinPer.SetFocus
        Validar = False: Exit Function
    End If
    
    
    If Trim(CmbMonBoleta.Text) = "" Then SSTab1.Tab = 3: MsgBox "Debe Seleccionar Moneda de Calculo de Boleta", vbCritical, TitMsg: Validar = False: CmbMonBoleta.SetFocus: Exit Function
    If Trim(CmbMonBoleta.Text) = "" Then SSTab1.Tab = 3: MsgBox "Debe Seleccionar Moneda de Calculo de Boleta", vbCritical, TitMsg: Validar = False: CmbMonBoleta.SetFocus: Exit Function
    If Trim(CmbTipoPago.Text) = "" Then SSTab1.Tab = 3: MsgBox "Debe Seleccionar Tipo de Pago", vbCritical, TitMsg: Validar = False: CmbTipoPago.SetFocus: Exit Function
    If CmbPerRem.ListIndex = -1 Then SSTab1.Tab = 3: MsgBox "Elija periodicidad de la remuneración de la lista, correctamene", vbCritical, TitMsg: Validar = False: CmbPerRem.SetFocus: Exit Function
    
    
    If Trim(Txtcodpla.Text) <> Trim(Lblplacod.Caption) Then
       Sql$ = "Select placod from planillas where cia='" & Trim(wcia) & "' and placod='" & Trim(Txtcodpla.Text) & "'"
       cn.CursorLocation = adUseClient
       Set Rs = New ADODB.Recordset
       Set Rs = cn.Execute(Sql$)
       If Rs.RecordCount > 0 Then SSTab1.Tab = 1: MsgBox "Codigo de Planilla ya fue Asignado a Otra Persona", vbCritical, TitMsg: Validar = False: Txtcodpla.SetFocus: Exit Function
       If Rs.State = 1 Then Rs.Close
    End If
    
    If rsbasico.RecordCount > 0 Then rsbasico.MoveFirst
    Do While Not rsbasico.EOF
       If CCur(rsbasico!importe) <> 0 Then
          If rsbasico!moneda = "" Or IsNull(rsbasico!moneda) Then SSTab1.Tab = 3: MsgBox "Debe Seleccionar Moneda en Remuneracion Basica", vbCritical, TitMsg: Validar = False: DgrBasico.SetFocus: Exit Function
          If rsbasico!tipo = "" Or IsNull(rsbasico!tipo) Then SSTab1.Tab = 3: MsgBox "Debe Seleccionar Tipo en Remuneracion Basica", vbCritical, TitMsg: Validar = False: DgrBasico.SetFocus: Exit Function
       End If
       rsbasico.MoveNext
    Loop
    If rsbasico.RecordCount > 0 Then rsbasico.MoveFirst
    
    If Trim(TxtFecCese.Text) <> "__/__/____" And cbo_TipMotFinPer.ReturnCodigo = -1 Then
            SSTab1.Tab = 1: MsgBox "Debe elegir el motivo fin de periodo", vbCritical, TitMsg: Validar = False: cbo_TipMotFinPer.SetFocus: Exit Function
    End If


End Select





If ChkOtrosIngta(3).Value = 1 Then

    With rsOtrosEmpleadores
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                    If Trim(!RUC & "") <> "" Then
                        If Not IsNumeric(!RUC) Then
                            SSTab1.Tab = 4
                            MsgBox "Ingrese el número de RUC del Empleador, correctamente", vbExclamation, Me.Caption
                            DgrdOtrosEmpleadores.Col = 0
                            DgrdOtrosEmpleadores.SetFocus
                            Validar = False: Exit Function
                        ElseIf Trim(!RUC & "") <> "" And Trim(!razsoc) = "" Then
                            SSTab1.Tab = 4
                            MsgBox "Ingrese Razón Social del Empleador, correctamente", vbExclamation, Me.Caption
                            DgrdOtrosEmpleadores.Col = 1
                            DgrdOtrosEmpleadores.SetFocus
                            Validar = False: Exit Function
                        End If
                    End If
                .MoveNext
            Loop
        Else
            SSTab1.Tab = 4
            MsgBox "Ingrese Otros Empleadores requerido por percibir otros ingresos de 5ta Categoria", vbCritical, Me.Caption
            Validar = False: Exit Function
        End If
    End With
End If
If rsTipEst.RecordCount > 0 Then
    Dim m As Integer
    m = 0
    With rsTipEst
        .MoveFirst
        Do While Not .EOF
                If !Add Then m = m + 1
            .MoveNext
        Loop
        
    End With
    If m = 0 Then
        SSTab1.Tab = 7
        MsgBox "Elija establecimiento(s) donde labora el trabajador", vbExclamation, Me.Caption
        Validar = False: Exit Function
    End If
End If
If Not optAFPMixta And Not optAFPFlujo Then
   SSTab1.Tab = 4
   MsgBox "Seleccione Tipo Comisión de Regimen Pensionario", vbExclamation, Me.Caption
   Validar = False: Exit Function
End If

If Not IsNumeric(TxtAnoEgreso.Text) And TxtAnoEgreso.Text <> "" Then
   SSTab1.Tab = 1
   MsgBox "Ingrese Correctamente Fecha de Egreso"
   Validar = False: Exit Function
Else
   If Val(TxtAnoEgreso.Text) < 1900 And TxtAnoEgreso.Text <> "" Then
      SSTab1.Tab = 1
      MsgBox "Ingrese Correctamente Fecha de Egreso"
      Validar = False: Exit Function
    End If
End If

Validar = True

End Function
Private Sub Crea_Rs()
    'suspension 4ta categoria
    If rsTipEst.State = 1 Then rsTipEst.Close
    rsTipEst.Fields.Append "add", adBoolean, , adFldIsNullable
    rsTipEst.Fields.Append "id", adVarChar, 15, adFldIsNullable
    rsTipEst.Fields.Append "codest", adVarChar, 4, adFldIsNullable
    rsTipEst.Fields.Append "nomest", adVarChar, 40, adFldIsNullable
    rsTipEst.Fields.Append "tipest", adVarChar, 80, adFldIsNullable
    rsTipEst.Fields.Append "razsoc", adVarChar, 100, adFldIsNullable
    rsTipEst.Fields.Append "ruc", adChar, 11, adFldIsNullable
    rsTipEst.Fields.Append "tipo", adChar, 25, adFldIsNullable
    rsTipEst.Open
    Set DgrdrTipEst.DataSource = rsTipEst
    
    'suspension 4ta categoria
    If rsSuspension4ta.State = 1 Then rsSuspension4ta.Close
    rsSuspension4ta.Fields.Append "numop", adChar, 15, adFldIsNullable
    rsSuspension4ta.Fields.Append "fecha", adChar, 10, adFldIsNullable
    rsSuspension4ta.Fields.Append "ejercicio", adChar, 4, adFldIsNullable
    rsSuspension4ta.Fields.Append "medio", adChar, 25, adFldIsNullable
    rsSuspension4ta.Open
    Set DgrdrsSuspension4ta.DataSource = rsSuspension4ta
    

    'Otros Empleadores
    If rsOtrosEmpleadores.State = 1 Then rsOtrosEmpleadores.Close
    rsOtrosEmpleadores.Fields.Append "ruc", adChar, 15, adFldIsNullable
    rsOtrosEmpleadores.Fields.Append "razsoc", adVarChar, 100, adFldIsNullable
    rsOtrosEmpleadores.Open
    Set DgrdOtrosEmpleadores.DataSource = rsOtrosEmpleadores

    'Contratos
    If rscontrato.State = 1 Then rscontrato.Close
    rscontrato.Fields.Append "nro_contrato", adChar, 10, adFldIsNullable
    rscontrato.Fields.Append "tipo_contrato", adChar, 2, adFldIsNullable
    rscontrato.Fields.Append "contrato", adVarChar, 150, adFldIsNullable
    rscontrato.Fields.Append "fecini", adChar, 10, adFldIsNullable
    rscontrato.Fields.Append "fecfin", adChar, 10, adFldIsNullable
    rscontrato.Fields.Append "meses", adInteger, 4, adFldIsNullable
    rscontrato.Fields.Append "cod_mot_fin_periodo", adVarChar, 2, adFldIsNullable
    rscontrato.Fields.Append "mot_fin_periodo", adVarChar, 50, adFldIsNullable
    rscontrato.Fields.Append "cod_tip_modalidad_formativa", adChar, 2, adFldIsNullable
    rscontrato.Fields.Append "tip_modalidad_formativa", adChar, 50, adFldIsNullable
    rscontrato.Fields.Append "cod_cargo", adChar, 3, adFldIsNullable
    rscontrato.Fields.Append "cargo", adVarChar, 50, adFldIsNullable
    rscontrato.Fields.Append "imp_sueldo", adChar, 10, adFldIsNullable
    rscontrato.Fields.Append "Reporte", adChar, 20, adFldIsNullable
    rscontrato.Fields.Append "Eliminar", adChar, 20, adFldIsNullable
    
    rscontrato.Fields.Append "FechaFirmaContrato", adChar, 10, adFldIsNullable
    
    
    rscontrato.Open
    Set DgrdContrato.DataSource = rscontrato

    'Telefonos
    If rstlf.State = 1 Then rstlf.Close
    rstlf.Fields.Append "telefono", adChar, 15, adFldIsNullable
    rstlf.Fields.Append "fax", adChar, 15, adFldIsNullable
    rstlf.Fields.Append "referencia", adChar, 25, adFldIsNullable
    rstlf.Open
    Set DgrdTelf.DataSource = rstlf
    
    'Judicial
    If rsJud.State = 1 Then rsJud.Close
    rsJud.Fields.Append "Nrodni", adChar, 15, adFldIsNullable
    rsJud.Fields.Append "Nombre", adChar, 50, adFldIsNullable
    rsJud.Fields.Append "TipoCta", adChar, 5, adFldIsNullable
    rsJud.Fields.Append "Bco", adChar, 5, adFldIsNullable
    rsJud.Fields.Append "NroCta", adChar, 25, adFldIsNullable
    rsJud.Fields.Append "Porcentaje", adChar, 15, adFldIsNullable
    rsJud.Fields.Append "CalcBaseBruto", adChar, 1, adFldIsNullable
    rsJud.Open
    Set DgrdJud.DataSource = rsJud
    
    
    
    'Uniformes
    If rsunif.State = 1 Then rsunif.Close
    'rsunif.Fields.Append "Botas", adChar, 10, adFldIsNullable
    'rsunif.Fields.Append "Camisa", adChar, 10, adFldIsNullable
    'rsunif.Fields.Append "Pantalon", adChar, 20, adFldIsNullable
    'rsunif.Fields.Append "Polo", adChar, 10, adFldIsNullable
    rsunif.Fields.Append "referencia", adChar, 25, adFldIsNullable
    rsunif.Fields.Append "Talla", adChar, 10, adFldIsNullable
    rsunif.Open
    Set DgrdUnif.DataSource = rsunif
       
    'Deducciones
    If rsdeduccion.State = 1 Then rsdeduccion.Close
    rsdeduccion.Fields.Append "descripcion", adChar, 35, adFldIsNullable
    rsdeduccion.Fields.Append "importe", adCurrency, 18, adFldIsNullable
    rsdeduccion.Fields.Append "porcentaje", adCurrency, 18, adFldIsNullable
    rsdeduccion.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsdeduccion.Open
    Set DgrDeduccion.DataSource = rsdeduccion
    
    'Basico
    If rsbasico.State = 1 Then rsbasico.Close
    rsbasico.Fields.Append "descripcion", adChar, 35, adFldIsNullable
    rsbasico.Fields.Append "tipo", adChar, 15, adFldIsNullable
    rsbasico.Fields.Append "moneda", adChar, 3, adFldIsNullable
    rsbasico.Fields.Append "importe", adCurrency, 18, adFldIsNullable
    rsbasico.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsbasico.Fields.Append "codtipo", adChar, 2, adFldIsNullable
    rsbasico.Fields.Append "horas", adInteger, 4, adFldIsNullable
    rsbasico.Open
    Set DgrBasico.DataSource = rsbasico
    
    'Centro de Costo
    If rsccosto.State = 1 Then rsccosto.Close
    rsccosto.Fields.Append "codigo", adChar, 2, adFldIsNullable
    rsccosto.Fields.Append "descripcion", adChar, 45, adFldIsNullable
    rsccosto.Fields.Append "monto", adCurrency, 18, adFldIsNullable
    rsccosto.Open
    rsccosto.AddNew
    rsccosto.AddNew
    rsccosto.AddNew
    rsccosto.AddNew
    rsccosto.AddNew
    Set Dgrdccosto.DataSource = rsccosto
    
End Sub
Private Sub Busca_Auxiliar()
If Vbuscaaux = False Then Exit Sub
LblTrabajador.Caption = Trim(TxtApePat.Text) & " " & Trim(TxtApeMat.Text) & " " & Trim(TxtPriNom.Text) & " " & Trim(TxtSegNom.Text)
Sql$ = nombre("a")

Sql$ = Sql$ & "a.ap_pat , a.ap_mat, a.ap_cas, a.nom_1, a.nom_2, a.codauxinterno " & _
        " from maestroaux a INNER JOIN planillas b ON (b.cia='" & wcia & "' and a.codauxinterno=b.codauxinterno and b.status!='*')" & _
        " where a.ap_pat like '" & Trim(TxtApePat.Text) + "%" & "' and a.ap_mat like '" & Trim(TxtApeMat.Text) + "%" & "' and a.ap_cas like '" & Trim(TxtApeCas.Text) + "%" & "' and a.nom_1 like '" & Trim(TxtPriNom.Text) + "%" & "' and a.nom_2 like '" & Trim(TxtSegNom.Text) + "%" & "' " _
    & "order by nombre"
    
cn.CursorLocation = adUseClient
Set AdoBusPla.Recordset = cn.Execute(Sql, 64)
If AdoBusPla.Recordset.EOF Then Framebuspla.Visible = False: Exit Sub
Framebuspla.Visible = True
Framebuspla.ZOrder 0
End Sub
Public Sub Nuevo_Personal(NewPla As Boolean)
VNew = NewPla
lblPlaCos.Caption = ""
VCodAfp = ""
VCodAuxi = ""
TxtApePat.Text = ""
TxtApeMat.Text = ""
TxtApeCas.Text = ""
TxtPriNom.Text = ""
TxtSegNom.Text = ""
'TxtDni.Text = ""
'TxtLm.Text = ""
'Txtpasaporte.Text = ""
'txtruc.Text = ""
TxtFecNac.Text = "__/__/____"
OpcVaron.Value = False
OpcDama.Value = False
TxtIpss.Text = ""
TxtNroAfp.Text = ""
TxtFecAfilia.Text = "__/__/____"
TxtFec_Afil_Sindicato.Text = "__/__/____"
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text7.Text = ""
TxtCCI.Text = ""
TxtCCICts.Text = ""
'TxtUbica.Text = ""
'Lblubigeo.Caption = ""
txtAporteEmpleado.Text = 0#

ChkDomiciliado.Value = 1
txtnro.Text = ""
txtint.Text = ""
TxtEmail.Text = ""

'optAFPMixta.Value = True

cbo_eps.SetIndex = -1
cbo_situacioneps.SetIndex = -1
cbo_pensiones.SetIndex = -1
'cbo_pension.SetIndex = -1
cbo_naciontrab.SetIndex = 9589
CmbPerRem.ListIndex = -1

'lblfecmodi.Caption = ""
'Lbluser.Caption = ""
Txtcodpla.Text = ""
Txtcodpla.Enabled = True
txtPresentacion.Text = ""
Lblplacod.Caption = ""
TxtFecIngreso.Text = "__/__/____"
ChkNoQuinta.Value = 0
ChkSindicato.Value = 0
ChkAltitud.Value = 0
ChkVaca.Value = 0
OpcVidaSi.Value = False
OpcVidaNo.Value = False
'OpcSctrSi.Value = False
'OpcSctrNo.Value = False
TxtFecCese.Text = "__/__/____"
TxtCtaPag.Text = ""
TxtCtaCts.Text = ""
Txtcodobra.Text = ""
Lbldesobra.Caption = ""
TxtRef.Text = ""
TxtRefAnt.Text = ""
ChkDiscapacidad.Value = 0
Cmbcia_Click

CmbCentroForma.ListIndex = -1
ChkPais.Value = 0
TxtNombreUni.Text = ""
TxtNombreUni.Tag = ""
TxtCarrera.Text = ""
TxtCarrera.Tag = ""
txtprofesion.Text = ""
txtprofesion.Tag = ""
OptPublica.Value = False
OptPrivada.Value = False

'Carga Deducciones Personales
Sql$ = "Select * from placonstante where cia='" & wcia & "' and status<>'*' and tipomovimiento='03' and personal='S' order by codinterno"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If rsdeduccion.RecordCount > 0 Then
   rsdeduccion.MoveFirst
   Do While Not rsdeduccion.EOF
      rsdeduccion.Delete
      rsdeduccion.MoveNext
   Loop
End If
If Rs.RecordCount > 0 Then Rs.MoveFirst
Do While Not Rs.EOF
   rsdeduccion.AddNew
   rsdeduccion!Descripcion = Rs!Descripcion
   rsdeduccion!Codigo = Rs!codinterno
   Rs.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'Carga Basicos
Sql$ = "Select codinterno,descripcion from placonstante where " & _
"tipomovimiento='02' and basico='S' and status<>'*' AND cia='" & wcia & "' order by " & _
"codinterno "   ' {<MA>} 01/02/2007
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If rsbasico.RecordCount > 0 Then
   rsbasico.MoveFirst
   Do While Not rsbasico.EOF
      rsbasico.Delete
      rsbasico.MoveNext
   Loop
End If
If Rs.RecordCount > 0 Then Rs.MoveFirst
Do While Not Rs.EOF
   rsbasico.AddNew
   rsbasico!Descripcion = Trim(Rs!Descripcion)
   rsbasico!Codigo = Rs!codinterno
   rsbasico!moneda = wmoncont
   rsbasico!importe = "0.00"
   Rs.MoveNext
Loop

If Rs.State = 1 Then Rs.Close

If rscontrato.RecordCount > 0 Then rscontrato.MoveFirst
Do While Not rscontrato.EOF
   rscontrato.Delete
   rscontrato.MoveNext
Loop

'If rsderechohab.RecordCount > 0 Then rsderechohab.MoveFirst
'Do While Not rsderechohab.EOF
'   rsderechohab.Delete
'   rsderechohab.MoveNext
'Loop

If rstlf.RecordCount > 0 Then rstlf.MoveFirst
Do While Not rstlf.EOF
   rstlf.Delete
   rstlf.MoveNext
Loop

If rsJud.RecordCount > 0 Then rsJud.MoveFirst
Do While Not rsJud.EOF
   rsJud.Delete
   rsJud.MoveNext
Loop

If rsunif.RecordCount > 0 Then rsunif.MoveFirst
Do While Not rsunif.EOF
   rsunif.Delete
   rsunif.MoveNext
Loop

If rsOtrosEmpleadores.RecordCount > 0 Then rsOtrosEmpleadores.MoveFirst
Do While Not rsOtrosEmpleadores.EOF
   rsOtrosEmpleadores.Delete
   rsOtrosEmpleadores.MoveNext
Loop



If rsSuspension4ta.RecordCount > 0 Then rsSuspension4ta.MoveFirst
Do While Not rsSuspension4ta.EOF
   rsSuspension4ta.Delete
   rsSuspension4ta.MoveNext
Loop


If rsTipEst.RecordCount > 0 Then rsTipEst.MoveFirst
Do While Not rsTipEst.EOF
   rsTipEst.Delete
   rsTipEst.MoveNext
Loop

FraOtrosEmpleadores(4).Enabled = False


OptEpsSi(0).Value = False
OptEpsNo(1).Value = False
ChkRegimen_Alternativo(0).Value = False
ChkJornada_max(1).Value = False
ChkHorario_nocturno(2).Value = False
ChkOtrosIngta(3).Value = False
Chk5taExonerada(4).Value = False
ChkMadreResp(1).Value = False

OptSitEsp_direccion(0).Value = False
OptSitEsp_confianza(1).Value = False
OptSitEsp_ninguna(2).Value = True

'/*ADD JC 07/08/08*/
ChkCalcula_AccidenteTrabajo(1).Value = 0
CmbEvitaDobleImp.ListIndex = -1
Me.OptAsegPenSi.Value = False
Me.OptAsegPenNo.Value = True
OpcVidaNo.Value = True

TxtApePat.SetFocus
Vbuscaaux = True
Carga_Establecimientos True
BtnMemo.Visible = False
End Sub
Private Function VALIDAR_DERECHOHAB()

'//** CONSISTENCIA SUNAT**///
'If CmbDocDh.ListIndex = -1 Then
'    SSTab1.Tab = 2
'    MsgBox "Elija tipo de documento Derechohabiente de la lista, correctamente", vbExclamation, Me.Caption
'    CmbDocDh.SetFocus
'    GoTo Salir:
'ElseIf Trim(TxtNroDocDh.Text) = "" Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese nro Documento Derechohabiente, correctamente", vbExclamation, Me.Caption
'    TxtNroDocDh.SetFocus
'    GoTo Salir:
'ElseIf Trim(TxtApePatDh.Text) = "" Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Apellido Paterno Derechohabiente, correctamente", vbExclamation, Me.Caption
'    TxtApePatDh.SetFocus
'    GoTo Salir:
'ElseIf Trim(TxtApeMatDh.Text) = "" Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Apellido Materno Derechohabiente, correctamente", vbExclamation, Me.Caption
'    TxtApeMatDh.SetFocus
'    GoTo Salir:
'ElseIf Trim(TxtPriNomDh.Text) = "" Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Primer Nombre Derechohabiente, correctamente", vbExclamation, Me.Caption
'    TxtPriNomDh.SetFocus
'    GoTo Salir:
'ElseIf Not IsDate(TxtFecNacDh.Text) Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Fecha de Nacimiento Derechohabiente, correctamente", vbExclamation, Me.Caption
'    TxtFecNacDh.SetFocus
'    GoTo Salir:
'ElseIf OpcVaronDh.Value = False And OpcDamaDh.Value = False Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Sexo Derechohabiente, correctamente", vbExclamation, Me.Caption
'    SSTab1.SetFocus
'    GoTo Salir:
'
'ElseIf OpcActivoDh.Value = False And OpcBajaDh.Value = False Then
'    SSTab1.Tab = 2
'    MsgBox "Elija Situación Derechohabiente, correctamente", vbExclamation, Me.Caption
'    SSTab1.SetFocus
'    GoTo Salir:
'ElseIf OpcActivoDh.Value = True And Not IsDate(FecAltaDH(0).Text) Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Fecha Situación activo Derechohabiente, correctamente", vbExclamation, Me.Caption
'    FecAltaDH(0).SetFocus
'    GoTo Salir:
'ElseIf OpcActivoDh.Value = True And CDate(FecAltaDH(0).Text) < CDate(TxtFecNacDh.Text) Then
'    SSTab1.Tab = 2
'    MsgBox "La Fecha de alta no puede ser anterior a la fecha de nacimiento", vbExclamation, Me.Caption
'    FecAltaDH(0).SetFocus
'    GoTo Salir:
'ElseIf OpcBajaDh.Value = True And Not IsDate(FecBajaDH(1).Text) Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Fecha de baja Derechohabiente, correctamente", vbExclamation, Me.Caption
'    FecBajaDH(1).SetFocus
'    GoTo Salir:
''ElseIf OpcBajaDh.Value = True And CDate(FecBajaDH(1).Text) < CDate(TxtFecNacDh.Text) Then
''    SSTab1.Tab = 2
''    MsgBox "La Fecha de baja no puede ser anterior a la fecha de nacimiento", vbExclamation, Me.Caption
''    FecBajaDH(1).SetFocus
''    GoTo salir:
''ElseIf OpcBajaDh.Value = True And CDate(FecBajaDH(1).Text) < CDate(FecBajaDH(1).Text) Then
''    SSTab1.Tab = 2
''    MsgBox "La Fecha de baja no puede ser anterior a la fecha de alta", vbExclamation, Me.Caption
''    FecBajaDH(1).SetFocus
''    GoTo salir:
'ElseIf OpcBajaDh.Value = True And cbobajadh.ListIndex = -1 Then
'    SSTab1.Tab = 2
'    MsgBox "Elija motivo de baja Derechohabiente, correctamente", vbExclamation, Me.Caption
'    cbobajadh.SetFocus
'    GoTo Salir:
'ElseIf CmbVinculo.ListIndex = -1 Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese vínculo familiar Derechohabiente de la lista, correctamente", vbExclamation, Me.Caption
'    CmbVinculo.SetFocus
'    GoTo Salir:
'ElseIf CmbVinculo.ItemData(CmbVinculo.ListIndex) = 1 And Trim(TxtIncapaz.Text) = "" And DateDiff("yyyy", Year(CDate(TxtFecNacDh)), Year(Date)) >= 18 Then 'hijo
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Nro de Certificado de Incapacidad", vbExclamation, Me.Caption
'    TxtIncapaz.SetFocus
'    GoTo Salir:
'
'ElseIf CmbVinculo.ItemData(CmbVinculo.ListIndex) = 4 And CmbTipDocAcreditaPaternidad.ListIndex = -1 Then 'gestante
'    SSTab1.Tab = 2
'    MsgBox "Elija tipo de documento que acredita la paternidad", vbExclamation, Me.Caption
'    CmbTipDocAcreditaPaternidad.SetFocus
'    GoTo Salir:
'ElseIf CmbVinculo.ItemData(CmbVinculo.ListIndex) = 4 And CmbTipDocAcreditaPaternidad.ListIndex = -1 Then 'gestante
'    SSTab1.Tab = 2
'    MsgBox "Elija tipo de documento que acredita la paternidad", vbExclamation, Me.Caption
'    CmbTipDocAcreditaPaternidad.SetFocus
'    GoTo Salir:
'ElseIf CmbVinculo.ItemData(CmbVinculo.ListIndex) = 4 And Trim(TxtNroDocAcreditapaternidad.Text) = "" Then
'    SSTab1.Tab = 2
'    MsgBox "Ingrese Número de documento que acredita la paternidad", vbExclamation, Me.Caption
'    TxtNroDocAcreditapaternidad.SetFocus
'    GoTo Salir:
'End If
'
'
'
'
'If TxtApePatDh.Text = "" Then SSTab1.Tab = 2: MsgBox "Debe Ingresar Primer Apellido", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: TxtApePatDh.SetFocus: Exit Function
'If TxtPriNomDh.Text = "" Then SSTab1.Tab = 2: MsgBox "Debe Ingresar Primer Nombre", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: TxtPriNomDh.SetFocus: Exit Function
'If Not IsDate(TxtFecNacDh.Text) Then SSTab1.Tab = 2: MsgBox "Ingrese Fecha De Nacimiento ", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: TxtFecNacDh.SetFocus: Exit Function
'If CmbVinculo.Text = "" Then SSTab1.Tab = 2: MsgBox "Debe Seleccionar Vinculo", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: CmbVinculo.SetFocus: Exit Function
'If OpcVaronDh.Value = False And OpcDamaDh.Value = False Then SSTab1.Tab = 2: MsgBox "Debe Indicar Sexo", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: Exit Function
'If CmbDocDh.Text = "" Then SSTab1.Tab = 2: MsgBox "Debe Seleccionar Tipo de Documento", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: CmbDocDh.SetFocus: Exit Function
'If TxtNroDocDh.Text = "" Then SSTab1.Tab = 2: MsgBox "Debe Ingresar Numero de Documento", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: TxtNroDocDh.SetFocus: Exit Function
'If OpcIncapazSi.Value = True And TxtIncapaz.Text = "" Then SSTab1.Tab = 2: MsgBox "Debe Ingresar Certificado de Incapacidad", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: TxtIncapaz.SetFocus: Exit Function
'If OpcDomiclioDhNo.Value = False And OpcDomiclioDhSi.Value = False Then SSTab1.Tab = 2: MsgBox "Debe Indicar Domicilio", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: Exit Function
''If CmbSituaDh.Text = "" Then SSTab1.Tab = 2: MsgBox "Debe Seleccionar Situacion", vbCritical, TitMsg: VALIDAR_DERECHOHAB = False: CmbSituaDh.SetFocus: Exit Function
'
'VALIDAR_DERECHOHAB = True
'Exit Function
'Salir:
'VALIDAR_DERECHOHAB = False
End Function
Public Sub Carga_Trabajador(CODPLA As String)
'Dim EPersonal As New ClsEPlanilla
'Dim DPersonal As New ClsDAPlanilla
Call Nuevo_Personal(False)
Vbuscaaux = False
Dim rsprof As ADODB.Recordset

'With EPersonal
'    EPersonal.Companhia = wcia
'    EPersonal.PlaCod = CODPLA
'End With

If Rs.State = 1 Then Rs.Close
Sql$ = "select a.*,dbo.fc_nombre_ubigeo_sunat(A.ubigeo) as nom_ubigeo from planillas a where cia='" & wcia & "' and placod='" & CODPLA & "' and a.status<>'*'"

'Set rs = DPersonal.Consultar_Edit_Personal(EPersonal)
'If (rs.EOF) Then MsgBox "No Se Encuentra el Trabajador", vbCritical, TitMsg: Exit Sub
If Not fAbrRst(Rs, Sql) Then MsgBox "No Se Encuentra el Trabajador", vbCritical, TitMsg: Exit Sub
VCodAuxi = Rs!codauxinterno
Call rUbiIndCmbBox(CmbCivil, Rs!EstadoCivil, "00")
If Trim(Rs!nacionalidad) = "001" Or Trim(Rs!nacionalidad) = "" Then
    cbo_naciontrab.SetIndex = 9589
Else
    cbo_naciontrab.SetIndice Rs!nacionalidad
End If
chk_integro.Value = IIf(IsNull(Rs!asignacion_fam_prorrateada) Or Rs!asignacion_fam_prorrateada = 0, 0, 1)
Call rUbiIndCmbBox(CmbNivelEdu, Rs!niveleducativo, "00")
Call rUbiIndCmbBox(CmbTipoTrabajador, Trim(Rs!TipoTrabajador), "00")
Call rUbiIndCmbBox(CmbCargo, Rs!Cargo, "000")
Call rUbiIndCmbBox(CmbArea, Rs!cod_area, "0000")
Call rUbiIndCmbBox(CmbResponsable, Trim(Rs!responsable & ""), "000")
Call rUbiIndCmbBox(CmbPlanta, Rs!Planta, "00")
Call rUbiIndCmbBox(CmbSegCom, Rs!codsctr, "00")
Call rUbiIndCmbBox(CmbTrabSunat, Trim(Rs!CodTipoTrabSunat & ""), "00")

If rsccosto.RecordCount > 0 Then
   rsccosto.MoveFirst
   Do While Not rsccosto.EOF
      rsccosto!Codigo = ""
      rsccosto!Descripcion = ""
      rsccosto!Monto = 0
      rsccosto.MoveNext
   Loop
   
'   Sql = "select c.ccosto,m.descrip,c.porc from planilla_ccosto c,maestros_32 m where c.cia='" & wcia & _
'         "' and c.placod='" & CODPLA & "' and c.status<>'*' and m.ciamaestro= '" & wcia + "055" & _
'         "' and cod_maestro2='" & Trim(rs!TipoTrabajador) & "' and cod_maestro3=c.ccosto order by c.porc desc"

 Sql = "select c.ccosto,m.descripcion,c.porc from planilla_ccosto c,Pla_ccostos m where c.cia='" & wcia & _
         "' and c.placod='" & CODPLA & "' and c.status<>'*' and m.cia=c.cia and m.codigo=c.ccosto and m.status<>'*' order by c.porc desc"
         
   Dim Rq As ADODB.Recordset
   If fAbrRst(Rq, Sql) Then
      Do While Not Rq.EOF
         rsccosto.MoveFirst
         Do While Not rsccosto.EOF
            If Trim(rsccosto!Codigo & "") = "" Then
               rsccosto!Codigo = Trim(Rq!ccosto & "")
               rsccosto!Descripcion = Trim(Rq!Descripcion & "")
               rsccosto!Monto = Rq!PORC
               Exit Do
            End If
            rsccosto.MoveNext
         Loop
         Rq.MoveNext
      Loop
      Rq.Close: Set Rq = Nothing: Sql = ""
   Else
      Sql = "select descrip from maestros_2 where ciamaestro= '" & wcia + "044" & "' and cod_maestro2='" & Rs!Area & "'"
      If fAbrRst(Rq, Sql) Then
         rsccosto.MoveFirst
         rsccosto!Codigo = Trim(Rs!Area & "")
         rsccosto!Descripcion = Trim(Rq!DESCRIP & "")
         rsccosto!Monto = 100
      End If
      Rq.Close: Set Rq = Nothing: Sql = ""
   End If
   
End If
'Fin Centros de Costo

Call rUbiIndCmbBox(CmbHorasExtras, Rs!tipotasaextra, "00")
'Call rUbiIndCmbBox(CmbSegCom, rs!codsctr, "0")
Call rUbiIndCmbBox(CmbPerRem, Rs!trab_periodicidad_remuneracion & "", "00")

For I = 0 To CmbMonPago.ListCount - 1
    If Right(Left(CmbMonPago.List(I), 4), 3) = Rs!pagomoneda Then CmbMonPago.ListIndex = I: Exit For
Next
For I = 0 To CmbMonBoleta.ListCount - 1
    If Right(Left(CmbMonBoleta.List(I), 4), 3) = Rs!moneda Then CmbMonBoleta.ListIndex = I: Exit For
Next
Call rUbiIndCmbBox(CmbCtaPag, Rs!pagotipcta, "00")
Call rUbiIndCmbBox(CmbCtaCts, Rs!ctstipcta, "00")
Call rUbiIndCmbBox(CmbTipoPago, Rs!tipopago, "00")
Call rUbiIndCmbBox(CmbBcoPago, Rs!pagobanco, "00")
For I = 0 To CmbMonCts.ListCount - 1
    If Right(Left(CmbMonCts.List(I), 4), 3) = Rs!ctsmoneda Then CmbMonCts.ListIndex = I: Exit For
Next
Call rUbiIndCmbBox(CmbBcoCts, Rs!ctsbanco, "00")
If Trim(Rs!discapacidad & "") = True Then
    ChkDiscapacidad.Value = 1
Else
    ChkDiscapacidad.Value = 0
End If
TxtApePat.Text = UCase(Trim(Rs!ap_pat))
TxtApeMat.Text = UCase(Trim(Rs!ap_mat))
TxtApeCas.Text = UCase(Trim(Rs!ap_cas))
TxtPriNom.Text = UCase(Trim(Rs!nom_1))
TxtSegNom.Text = UCase(Trim(Rs!nom_2))
If IsNull(Rs!placos) = False Then
    lblPlaCos.Caption = Trim(Rs!placos)
Else
    lblPlaCos.Caption = ""
End If



Call rUbiIndCmbBox(Cbotipodoc, Trim(Rs!tipo_doc & ""), "00")
Text8.Text = Trim(Rs!nro_doc & "")

'TxtDni.Text = Trim(rs!dni)
'TxtLm.Text = Trim(rs!lmilitar)
'Txtpasaporte.Text = Trim(rs!pasaporte)
'txtruc.Text = Trim(rs!RUC)
If IsNull(Rs!fnacimiento) Then TxtFecNac.Text = "__/__/____" Else TxtFecNac.Text = Format(Rs!fnacimiento, "dd/mm/yyyy")
If Rs!sexo = "F" Then OpcDama.Value = True Else OpcVaron.Value = True

If Trim(Rs!afiliado_eps_serv & "") = True Then
    OptEpsSi(0).Value = True
Else
    OptEpsNo(1).Value = True
End If

ChkRegimen_Alternativo(0).Value = IIf(Trim(Rs!trab_reg_alternativo & "") = True, 1, 0)
ChkJornada_max(1).Value = IIf(Trim(Rs!trab_jornada_trab_max & "") = True, 1, 0)
ChkHorario_nocturno(2).Value = IIf(Trim(Rs!trab_hor_nocturno & "") = True, 1, 0)
ChkOtrosIngta(3).Value = IIf(Trim(Rs!trab_otros_ing_5ta & "") = True, 1, 0)
Chk5taExonerada(4).Value = IIf(Trim(Rs!trab_5ta_exonerada_inafecta & "") = True, 1, 0)
ChkMadreResp(1).Value = IIf(Trim(Rs!trab_madre_resp_familiar & "") = True, 1, 0)

Select Case Trim(Rs!trab_situacion_especial & "")
    Case "1": OptSitEsp_direccion(0).Value = True
    Case "2": OptSitEsp_confianza(1).Value = True
    Case "0": OptSitEsp_ninguna(2).Value = True
End Select


'/*ADD JC 04/08/08*/
ChkCalcula_AccidenteTrabajo(1).Value = IIf(Trim(Rs!trab_Calc_AccidenteTrabajo & "") = True, 1, 0)
Call rUbiIndCmbBox(CmbEvitaDobleImp, Trim(Rs!trab_evita_doble_tribu & ""), "00")
If Trim(Rs!trab_afiliacion_asegura_tu_pension) = True Then
    Me.OptAsegPenSi.Value = True
Else
    Me.OptAsegPenNo.Value = True
End If

If Len(Trim(Rs!profesion & "")) > 0 Then
    Sql = "SELECT DESCRIPCION FROM tocupaciones where codigo='" & Trim(Rs!profesion & "") & "' "
    If (fAbrRst(rsprof, Sql)) Then
        txtprofesion.Text = CStr(Trim(rsprof(0) & ""))
        txtprofesion.Tag = Trim(Rs!profesion & "")
    Else
        txtprofesion.Text = ""
    End If
    rsprof.Close
End If


If Rs!Est_Pub_Priv = 1 Then OptPublica.Value = True
If Rs!Est_Pub_Priv = 2 Then OptPrivada.Value = True
If Rs!Est_Ext = 1 Then ChkPais = 1 Else ChkPais = 0

If Trim(Rs!centroformacion & "") <> "" Then
   For I = 0 To CmbCentroForma.ListCount - 1
       If Left(CmbCentroForma.List(I), 2) = Trim(Rs!centroformacion & "") Then
          CmbCentroForma.ListIndex = I
       End If
   Next
End If

If Trim(Rs!Est_Cod_Inst & "") <> "" Then
    Sql = "select DISTINCT(DESC_INSTITUCION) from SUNAT_INSTITUCIONES_EDUCATIVAS where COD_INSTITUCION='" & Trim(Rs!Est_Cod_Inst & "") & "'"
    If (fAbrRst(rsprof, Sql)) Then
        TxtNombreUni.Text = CStr(Trim(rsprof(0) & ""))
        TxtNombreUni.Tag = Trim(Rs!Est_Cod_Inst & "")
    Else
        TxtNombreUni.Text = ""
    End If
    rsprof.Close
End If

If Trim(Rs!Est_Cod_Carrera & "") <> "" Then
    Sql = "select DISTINCT(DESC_CARRERA) from SUNAT_INSTITUCIONES_EDUCATIVAS where COD_carrera='" & Trim(Rs!Est_Cod_Carrera & "") & "'"
    If (fAbrRst(rsprof, Sql)) Then
        TxtCarrera.Text = CStr(Trim(rsprof(0) & ""))
        TxtCarrera.Tag = Trim(Rs!Est_Cod_Carrera & "")
    Else
        TxtCarrera.Text = ""
    End If
    rsprof.Close
End If

If Trim(Rs!Est_ano_egreso & "") <> "" Then
   
   If Rs!Est_ano_egreso > 0 Then TxtAnoEgreso.Text = Rs!Est_ano_egreso Else TxtAnoEgreso.Text = ""
Else
   TxtAnoEgreso.Text = ""
End If

cbo_zonatrab.SetIndice Rs!tzona
Me.cbo_viatrab.SetIndice Rs!tvia
cbo_cattrab.SetIndice Trim(Rs!cat_trab & "")

'If Not IsNull(Rs!sn_trab_pens) Then
'    If Rs!sn_trab_pens Then
'        Me.opttrabpen.Value = True
'    Else
'        optestrabpen.Value = True
'    End If
'Else
'    optestrabpen.Value = False
'    opttrabpen.Value = False
'End If
'cbo_pension.SetIndice Trim(Rs!tipo_pension & "")
cbo_situacioneps.SetIndice Trim(Rs!estado_eps & "")
cbo_eps.SetIndice Trim(Rs!codigo_eps & "")
cbo_pensiones.SetIndice Trim(Rs!CodAfp & "")
Text7.Text = Trim(Rs!ruc_intermedio & "")
cboessalud.SetIndice Trim(Rs!scrt_salud & "")
cbopension.SetIndice Trim(Rs!sctr_pension & "")

cbo_TipModFor.SetIndice Trim(Rs!tipo_modalida_formativa & "")
cbo_TipMotFinPer.SetIndice Trim(Rs!mot_fin_periodo & "")

If Not IsNull(Rs!sujeto_fiscalizacion) Then If CInt(Rs!sujeto_fiscalizacion) = 1 Then Check1.Value = vbChecked Else Check1.Value = vbUnchecked
cbo_tipocont.SetIndice Trim(Rs!Tipo_contrato & "")

TxtIpss.Text = Trim(Rs!ipss)
TxtNroAfp.Text = Trim(Rs!NUMAFP)
If IsNull(Rs!afpfechaafil) Then TxtFecAfilia.Text = "__/__/____" Else TxtFecAfilia.Text = Format(Rs!afpfechaafil, "dd/mm/yyyy")
If IsNull(Rs!Fecha_Sindicato) Then TxtFec_Afil_Sindicato.Text = "__/__/____" Else TxtFec_Afil_Sindicato.Text = Format(Rs!Fecha_Sindicato, "dd/mm/yyyy")
If Trim(Rs!domiciliado & "") = True Or IsNull(Rs!domiciliado) Then
    ChkDomiciliado.Value = 1
Else
    ChkDomiciliado.Value = 0
End If
Text9.Text = Trim(Rs!nomzona)
Text10.Text = Trim(Rs!nomvia)

'Text13.Text = Rs!pais & " - " & Rs!DPTO & " - " & Rs!PROV & " - " & Rs!DIST
Text13.Text = Trim(Rs!nom_ubigeo & "")

Text13.Tag = Trim(Rs!ubigeo)
txtnro.Text = Trim(Rs!nrokmmza)
txtint.Text = Trim(Rs!intdptolote)
TxtEmail.Text = Trim(Rs!Email)
TxtRef.Text = Trim(Rs!referencia & "")
TxtRefAnt.Text = Trim(Rs!referencia_antigua & "")

'PLAME
TxtNroDpto1.Text = Trim(Rs!nro_departamento & "")
TxtNroMz1.Text = Trim(Rs!nro_manzana & "")
TxtNroLote1.Text = Trim(Rs!nro_lote & "")
TxtNroKM1.Text = Trim(Rs!nro_kilometro & "")
TxtNroBlock1.Text = Trim(Rs!nro_block & "")
TXtNroEtapa1.Text = Trim(Rs!nro_etapa & "")

If Trim(Rs!afptipocomision & "") = "M" Then
    optAFPMixta = True
Else
    optAFPFlujo = True
End If

' ADD LFSA 26/10/2012 % DE APORTE DEL EMPLEADOR
If IsNull(Rs!PorcAport) = True Then
    txtAporteEmpleado.Text = 0#
Else
    txtAporteEmpleado.Text = Val(Rs!PorcAport)
End If
'/* ADD POR RICARDO HINOSTROZA 07/11/2008 */
'/* MOTIVO: TRANFERENCIA BCP */

Txt_Sucursal.Text = IIf(IsNull(Rs!Sucursal), "", Rs!Sucursal)

'/*FIN  DE MODIFICACION */

'If IsNull(rs!fechamodificacion) Then lblfecmodi.Caption = "" Else lblfecmodi.Caption = Format(rs!fechamodificacion, "dd/mm/yyyy")
'Lbluser.Caption = Trim(rs!usuariomodi)
Txtcodpla.Text = Trim(Rs!PlaCod)
Lblplacod.Caption = Trim(Rs!PlaCod)
txtPresentacion.Text = UCase(Trim(Rs!PlaCodpresentacion))

If Not IsNull(Rs!fIngreso) Then
    TxtFecIngreso.Text = Format(Rs!fIngreso, "dd/mm/yyyy")
End If
Txtcodpla.Enabled = False
If Rs!quinta = "N" Then ChkNoQuinta.Value = 1 Else ChkNoQuinta.Value = 0
If Rs!sindicato = "S" Then ChkSindicato.Value = 1 Else ChkSindicato.Value = 0
If Rs!essaludvida = "S" Then OpcVidaSi.Value = True Else OpcVidaNo.Value = True
'If rs!sctr = "S" Then OpcSctrSi.Value = True Else OpcSctrNo.Value = True
If Rs!altitud = "S" Then ChkAltitud.Value = 1
If Rs!vacacion = "S" Then ChkVaca.Value = 1
If IsNull(Rs!fcese) Then TxtFecCese.Text = "__/__/____" Else TxtFecCese.Text = Format(Rs!fcese, "dd/mm/yyyy"): BtnMemo.Visible = True
TxtCtaPag.Text = Trim(Rs!pagonumcta)
TxtCtaCts.Text = Trim(Rs!ctsnumcta)
TxtCCI.Text = Trim(Rs!pagocci & "")
TxtCCICts.Text = Trim(Rs!ctscci & "")
If IsNull(Rs!txtrajub) Then
    chkJubilado.Value = 0
Else
    If Trim(Rs!txtrajub) = "S" Then
        chkJubilado.Value = 1
    Else
        chkJubilado.Value = 0
    End If
End If
LblCodCantera.Caption = Trim(Rs!cantera & "")
If IsNull(Rs!fec_jubila) Then Txtfecjubila.Text = "__/__/____" Else Txtfecjubila.Text = Format(Rs!fec_jubila, "dd/mm/yyyy")
If IsNull(Rs!obra) Then Txtcodobra.Text = "" Else Txtcodobra.Text = Rs!obra
If Rs.State = 1 Then Rs.Close

'Telefonos
Sql$ = "select * from platelefono where cia='" & wcia & "' and placod='" & CODPLA & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   Rs.MoveFirst
   Do While Not Rs.EOF
      rstlf.AddNew
      rstlf!telefono = Rs!telefono
      rstlf!fax = Rs!fax
      rstlf!referencia = Rs!Descripcion
      Rs.MoveNext
   Loop
End If
If Rs.State = 1 Then Rs.Close

'Judiciales
Sql$ = "select * from TBL_BCO_CUENTA_DJ where placod='" & CODPLA & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   Rs.MoveFirst
   Do While Not Rs.EOF
      rsJud.AddNew
      rsJud!Nrodni = Rs!Nrodni
      rsJud!nombre = Rs!nombre
      rsJud!Tipocta = Rs!Tipocta
      rsJud!Bco = Rs!Bco
      rsJud!NroCta = Rs!NroCta
      rsJud!Porcentaje = Rs!Porcentaje
      rsJud!CalcBaseBruto = IIf(Rs!IndicadorCalculoEnBaseIngresoTotalBruto = True, "S", "")
      Rs.MoveNext
   Loop
End If
If Rs.State = 1 Then Rs.Close




Uniformes_1

'Deduccion Personal
If rsdeduccion.RecordCount > 0 Then rsdeduccion.MoveFirst
Do While Not rsdeduccion.EOF
   Set Rs = cn.Execute("select * from pladeducper where cia='" & wcia & "' and placod='" & CODPLA & "' and status<>'*' and concepto='" & rsdeduccion!Codigo & "'")
   If Rs.RecordCount > 0 Then
      If Rs!status = "P" Then
         rsdeduccion!importe = "0.00"
         rsdeduccion!Porcentaje = Rs!importe
      Else
         rsdeduccion!Porcentaje = "0.00"
         rsdeduccion!importe = Rs!importe
      End If
   End If
   rsdeduccion.MoveNext
Loop
If Rs.State = 1 Then Rs.Close
'Retencion Judicial en Boleta de Utilidades
Set Rs = cn.Execute("select placod from PlaRetJudUti where cia='" & wcia & "' and placod='" & CODPLA & "' and status<>'*'")
If Rs.RecordCount > 0 Then ChkRetJudUti.Value = 1
If Rs.State = 1 Then Rs.Close


'Remuneraciones Basicas
If rsbasico.RecordCount > 0 Then rsbasico.MoveFirst
Do While Not rsbasico.EOF
   wciamae = Determina_Maestro("01076")
   Sql$ = "select a.*,b.flag2,b.descrip from plaremunbase a," & _
   "maestros_2 b where a.cia='" & wcia & "' and a.tipo=" & _
   "b.cod_maestro2 and placod='" & CODPLA & "' and a.status<>'*' and concepto='" & rsbasico!Codigo & "'"
   Sql$ = Sql$ & wciamae
   Set Rs = cn.Execute(Sql$)
   If Rs.RecordCount > 0 Then
      rsbasico!tipo = Rs!DESCRIP
      rsbasico!codtipo = Rs!tipo
      rsbasico!moneda = Rs!moneda
      rsbasico!importe = Rs!importe
      rsbasico!horas = Rs!flag2
   End If
   rsbasico.MoveNext
Loop
If Rs.State = 1 Then Rs.Close

'Derecho Habientes
Sql$ = nombre()
wciamae = Determina_Maestro("01071")

'If rsderechohab.RecordCount > 0 Then rsderechohab.MoveFirst
'Do While Not rsderechohab.EOF
'   rsderechohab.Delete
'   rsderechohab.MoveNext
'Loop

'Sql$ = Sql$ & "a.*,b.descrip,dbo.fc_nombre_ubigeo_sunat(A.ubigeo) as nom_ubigeo "
'Sql$ = Sql$ & ",dbo.fc_nombre_ubigeo_sunat(A.ubigeo_lugar_nac) as nom_ubigeo_lugar_nac from pladerechohab a,maestros_2 b where a.cia='" & wcia & "' and a.status<>'*' and b.status<>'*' and a.codvinculo=b.cod_maestro2 and placod='" & CODPLA & "'"
'Sql$ = Sql$ & wciamae
'cn.CursorLocation = adUseClient
'Set rs = New ADODB.Recordset
'Set rs = cn.Execute(Sql$, 64)
'If rs.RecordCount > 0 Then
'   rs.MoveFirst
'   Do While Not rs.EOF
'      rsderechohab.AddNew
'      rsderechohab!nombre = Trim(rs!ap_pat) & " " & Trim(rs!ap_mat) & " " & Trim(rs!nom_1) & " " & Trim(rs!nom_2)
'      rsderechohab!vinculo = Trim(rs!DESCRIP)
'      rsderechohab!fecha_naci = Format(rs!fec_nac, "dd/mm/yyyy")
'      rsderechohab!ap_pat = Trim(rs!ap_pat)
'      rsderechohab!ap_mat = Trim(rs!ap_mat)
'      rsderechohab!nom_1 = Trim(rs!nom_1)
'      rsderechohab!nom_2 = Trim(rs!nom_2)
'      rsderechohab!cod_doc = Trim(rs!cod_doc)
'      rsderechohab!numero = Trim(rs!numero)
'      rsderechohab!codvinculo = Trim(rs!codvinculo)
'      'rsderechohab!atencionmed = "" 'Trim(Rs!atencionmed)
'      rsderechohab!sexo = Trim(rs!sexo)
'      rsderechohab!situacion = Trim(rs!situacion)
'      'rsderechohab!fallecimiento = Trim(rs!fallecimiento)
'      rsderechohab!incapacidad = Trim(rs!incapacidad)
'      rsderechohab!nrocertificado = Trim(rs!nrocertificado)
'      rsderechohab!domicilio = Trim(rs!domicilio)
'      rsderechohab!escolar = Trim(rs!escolar)
'
'      rsderechohab!nacionalidad = Trim(rs!nacionalidad & "")
'      rsderechohab!cod_mot_baja = Trim(rs!cod_mot_baja & "")
'      rsderechohab!ubigeo_lugar_nac = Trim(rs!ubigeo_lugar_nac & "")
'      rsderechohab!cod_zona = Trim(rs!cod_zona & "")
'      rsderechohab!NOM_ZONA = Trim(rs!NOM_ZONA & "")
'      rsderechohab!cod_via = Trim(rs!cod_via & "")
'      rsderechohab!NOM_VIA = Trim(rs!NOM_VIA & "")
'      rsderechohab!NRO = Trim(rs!NRO & "")
'      rsderechohab!Interior = Trim(rs!Interior & "")
'      rsderechohab!ubigeo = Trim(rs!ubigeo & "")
'
'
'      rsderechohab!ubigeo_desc = Trim(rs!nom_ubigeo & "")
'      rsderechohab!ubigeonac_desc = Trim(rs!nom_ubigeo_lugar_nac & "")
'
'
'    rsderechohab!tipdoc_acredita_paternidad = Trim(rs!tipdoc_acredita_paternidad & "")
'    rsderechohab!nrodoc_acredita_paternidad = Trim(rs!nrodoc_acredita_paternidad & "")
'    rsderechohab!fecha_alta = Trim(rs!fecha_alta & "")
'    rsderechohab!fecha_baja = Trim(rs!fecha_baja & "")
'    rsderechohab!ref = Trim(rs!referencia & "")
'
'
'
'      rs.MoveNext
'
'   Loop
'End If

Dim Rt As ADODB.Recordset
Sql = "select ruc,razsoc from plaOtrosEmpleadores where cia='" & wcia & "' and placod='" & CODPLA & "' and status<>'*'"
If fAbrRst(Rt, Sql) Then
    Do While Not Rt.EOF
            rsOtrosEmpleadores.AddNew
            rsOtrosEmpleadores!RUC = Trim(Rt!RUC & "")
            rsOtrosEmpleadores!razsoc = Trim(Rt!razsoc & "")
        Rt.MoveNext
    Loop
End If

Sql = "select * from tsuspension where cod_cia='" & wcia & "' and cod_prov='" & CODPLA & "' and status<>'*'"
If fAbrRst(Rt, Sql) Then
    Do While Not Rt.EOF
            rsSuspension4ta.AddNew
            rsSuspension4ta!numop = Trim(Rt!num_oper & "")
            rsSuspension4ta!fecha = Format(Trim(Rt!fec_susp & ""), "dd/mm/yyyy")
            rsSuspension4ta!ejercicio = Trim(Rt!ejercicio & "")
            rsSuspension4ta!medio = Trim(Rt!med_pres & "")
        Rt.MoveNext
    Loop
End If
LblCantera.Caption = ""
If Trim(LblCodCantera.Caption & "") <> "" Then
   Sql = "select descripcion from canteras where cia='" & wcia & "' and codcantera='" & Trim(LblCodCantera.Caption & "") & "'"
   If fAbrRst(Rt, Sql) Then LblCantera.Caption = Trim(Rt!Descripcion & "")
End If
Rt.Close
Set Rt = Nothing

Carga_detalle_contratos CODPLA
Carga_Establecimientos False

Vbuscaaux = True
End Sub

Public Sub HabilitaDh()
cbo_zona.Enabled = True
cbo_via.Enabled = True
Text6.Enabled = True
Text5.Enabled = True
Text4.Enabled = True
Text3.Enabled = True
Text2.Enabled = True
TxtRefdh.Enabled = True
lblubigeo(2).Enabled = True
LimpiarDh
End Sub

Public Sub DesHabilitaDh()
cbo_zona.Enabled = False
cbo_via.Enabled = False
Text6.Enabled = False
Text5.Enabled = False
Text4.Enabled = False
Text3.Enabled = False
Text2.Enabled = False
TxtRefdh.Enabled = False
lblubigeo(2).Enabled = False
LimpiarDh
End Sub

Public Sub LimpiarDh()
cbo_zona.ListIndex = -1
cbo_via.ListIndex = -1
Text6.Text = ""
Text5.Text = ""
Text4.Text = ""
Text3.Text = ""
Text2.Text = ""
lblubigeo(2).Caption = ""
Text2.Tag = ""

End Sub

Public Sub LimpiarNoDomiciliado()
cbo_viatrab.ListIndex = -1
Text10.Text = ""
txtnro.Text = ""
txtint.Text = ""
cbo_zonatrab.ListIndex = -1
Text9.Text = ""
Text13.Text = ""
Text13.Tag = ""
TxtRef.Text = ""
TxtEmail.Text = ""
End Sub

Public Function fc_idSunat(ByVal pCia As String, ByVal pCiaMaestro As String, ByVal pCodMaestro2 As String) As String
Dim Rq As ADODB.Recordset
Dim Sql As String
Sql = "select codsunat from maestros_2 where ciamaestro='" & pCiaMaestro & "' and cod_maestro2='" & pCodMaestro2 & "' and status<>'*'"
If fAbrRst(Rq, Sql) Then
    fc_idSunat = Trim(Rq!CODSUNAT & "")
Else
    fc_idSunat = ""
End If
Rq.Close
Set Rq = Nothing
End Function

Public Sub Apertura_Registro_Contrato_Trabajador()
Dim xImporte As Currency
xImporte = 0
xImporte = Calcula_ImporteContrato(wcia, Trim(Txtcodpla.Text))
If xImporte = 0 Then
    SSTab1.Tab = 3
    MsgBox "Ingrese Remuneracion del trabajador antes de generar el contrato", vbCritical, Me.Caption
    Exit Sub
End If
Dim xId As String
xId = fc_CodigoComboBox(cbo_cattrab, 2)
rscontrato.AddNew
'If xId = "01" Or xId = "02" Then
'    rscontrato!cod_mot_fin_periodo = Format(cbo_TipMotFinPer.ReturnCodigo, "00")
'    rscontrato!mot_fin_periodo = cbo_TipMotFinPer.Text
'Else
'    rscontrato!cod_mot_fin_periodo = ""
'    rscontrato!mot_fin_periodo = ""
'End If
If xId = 5 Then
    rscontrato!cod_tip_modalidad_formativa = Format(cbo_TipModFor.ReturnCodigo, "00")
    rscontrato!tip_modalidad_formativa = cbo_TipModFor.Text
Else
    rscontrato!cod_tip_modalidad_formativa = ""
    rscontrato!tip_modalidad_formativa = ""
End If
rscontrato!cod_Cargo = fc_CodigoComboBox(CmbCargo, 3)
rscontrato!Cargo = CmbCargo.Text
rscontrato!imp_sueldo = Format(xImporte, "###,##0.00")
rscontrato!Tipo_contrato = Format(cbo_tipocont.ReturnCodigo, "00")
rscontrato!contrato = cbo_tipocont.Text
rscontrato!Reporte = "Generar Contrato"
Me.DgrdContrato.Col = 1
Me.DgrdContrato.SetFocus

End Sub

Public Function Calcula_ImporteContrato(ByVal pCia As String, ByVal pCodTrab As String) As Currency
Dim Sql As String
Select Case wruc
Case "20100574005"
    Sql = "select importe from plaremunbase where cia = '" & pCia & "' and placod = '" & pCodTrab & "' and status != '*' and tipo = '01'"
Case Else
    Sql = "select sum("
    Sql = Sql & " case tipo  when '01' then importe*30 " '--diario
    Sql = Sql & " when '02' then importe*4 " 'semanal
    Sql = Sql & " when '03' then importe*2 " 'Quincenal
    Sql = Sql & " when '04' then importe*1 " '--mensual
    Sql = Sql & " else importe end"
    Sql = Sql & " ) AS  TOTAL"
    Sql = Sql & " from plaremunbase where  cia='" & pCia & "' and placod='" & Trim(pCodTrab) & "' AND STATUS<>'*' "
End Select

Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
    If Rq(0) = 0 Then
        MsgBox "El importe del Contrato es Cero, ingrese conceptos remunerativos del trabajador", vbCritical, Me.Caption
            Calcula_ImporteContrato = 0
    Else
        Calcula_ImporteContrato = IIf(IsNull(Rq(0)), 0, Rq(0))
    End If
Else
    Calcula_ImporteContrato = 0
    MsgBox "No existen conceptos remunerativos para calcular el importe del contrato", vbCritical, Me.Caption
End If
Rq.Close
Set Rq = Nothing
End Function

Public Sub Carga_detalle_contratos(ByVal CODPLA As String)
If rscontrato.RecordCount > 0 Then rscontrato.MoveFirst
Do While Not rscontrato.EOF
   rscontrato.Delete
   rscontrato.MoveNext
Loop

Dim Rt As ADODB.Recordset
Sql = "select * "
Sql = Sql & " ,(select descrip from maestros_2 where ciamaestro=placontrato.codcia+'144' and cod_maestro2=placontrato.cod_tip_contrato and status<>'*') as tipo_contrato"
Sql = Sql & " ,(select descrip from maestros_2 where ciamaestro=placontrato.codcia+'150' and cod_maestro2=placontrato.cod_mod_formativa and status<>'*') as modalidad_formativa"
Sql = Sql & " ,(select descrip from maestros_2 where ciamaestro=placontrato.cod_fperiodo+'149' and cod_maestro2=placontrato.cod_fperiodo and status<>'*') as mot_fin_periodo"
Sql = Sql & " from placontrato where codcia='" & wcia & "' and placod='" & CODPLA & "' and status<>'*'"
If fAbrRst(Rt, Sql) Then
    Do While Not Rt.EOF
        rscontrato.AddNew
        rscontrato!nro_contrato = Trim(Rt!num_contrato & "")
        rscontrato!Tipo_contrato = Trim(Rt!cod_tip_contrato & "")
        rscontrato!contrato = Trim(Rt!Tipo_contrato & "")
        rscontrato!FecIni = Format(Trim(Rt!fec_ini & ""), "dd/mm/yyyy")
        If IsDate(Rt!fec_fin) Then
            rscontrato!FecFin = Format(Trim(Rt!fec_fin & ""), "dd/mm/yyyy")
        Else
            rscontrato!FecFin = ""
        End If
        'rscontrato!meses = ""
        rscontrato!cod_mot_fin_periodo = Trim(Rt!cod_fperiodo & "")
        rscontrato!mot_fin_periodo = Trim(Rt!mot_fin_periodo & "")
        rscontrato!cod_tip_modalidad_formativa = Trim(Rt!cod_mod_formativa & "")
        rscontrato!tip_modalidad_formativa = Trim(Rt!modalidad_formativa & "")
        rscontrato!cod_Cargo = Trim(Rt!Cargo & "")
        rscontrato!Cargo = Trim(Rt!Cargo & "")
        rscontrato!imp_sueldo = Format(Rt!importe, "###,##0.00")
        rscontrato!Reporte = "Generar Contrato"
        
        If IsDate(Rt!fec_fin) Then
            rscontrato!FechaFirmaContrato = Format(Trim(Rt!fecha_firma_contrato & ""), "dd/mm/yyyy")
        Else
            rscontrato!FechaFirmaContrato = ""
        End If
        
        Rt.MoveNext
    Loop
End If
Rt.Close
Set Rt = Nothing
End Sub
Sub EnviaWord(ByVal pRs As ADODB.Recordset, pNomFIle As String)
'procedimiento evnvia word
'Carlos Pua ......07/01/2008

'On Error GoTo MsgErr:
Dim Cadena As String
Dim moneda As String

    Set W = CreateObject("word.application")
    'Temporal Cambiar po seteo diferentes contratos para la misma Cia
    If wruc = "20100574005" Then
       W.Documents.Open FileName:=App.Path & "\Reports\contrato2.doc"
    Else
       W.Documents.Open FileName:=App.Path & "\Reports\contrato.doc"
    End If
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO01"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!razsoc)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With

    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO02"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!RUC)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO03"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!direcc)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO04"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!rep_nom)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
     With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO05"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!rep_LE & "")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO06"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!nombre)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO07"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!DNI)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO08"
        .Replacement.ClearFormatting
        .Replacement.Text = IIf(IsNull(Trim(pRs!Direccion)), "", Trim(pRs!Direccion))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO04"
        .Replacement.ClearFormatting
        .Replacement.Text = Trim(pRs!rep_nom)
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO09"
        .Replacement.ClearFormatting
        .Replacement.Text = IIf(IsNull(Trim(pRs!DISTRITO)), "", Trim(pRs!DISTRITO))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO10"
        .Replacement.ClearFormatting
        .Replacement.Text = Day(pRs!fec_ini) & " " & "de " & Format(pRs!fec_ini, "MMMM") & " " & "del " & Format(pRs!fec_ini, "yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO11"
        .Replacement.ClearFormatting
        If IsNull(pRs!fec_fin) Then
            .Replacement.Text = "INDEFINIDO"
        Else
        .Replacement.Text = "el " & Day(pRs!fec_fin) & " " & "de " & Format(pRs!fec_fin, "MMMM") & " " & "del " & Format(pRs!fec_fin, "yyyy")
        End If
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO12"
        .Replacement.ClearFormatting
        .Replacement.Text = IIf(IsNull(Trim(pRs!Cargo)), "", Trim(pRs!Cargo))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO13"
        .Replacement.ClearFormatting
        .Replacement.Text = monto_palabras(pRs!importe) & " " & IIf(wmoncont = "S/.", "Soles", "Dolares Americanos") & "(" & wmoncont & " " & Format(pRs!importe, "##0.00") & ")"
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO14"
        .Replacement.ClearFormatting
        If IsNull(pRs!meses) Then
            .Replacement.Text = ""
        Else
            .Replacement.Text = fc_Numero(pRs!meses) & "(" & pRs!meses & ") " & "meses"
        End If
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    

    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO15"
        .Replacement.ClearFormatting
        .Replacement.Text = "Lima, " & Day(pRs!fec_crea) & " " & "de " & Format(pRs!fec_crea, "MMMM") & " " & "del " & Format(pRs!fec_crea, "yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO17"
        .Replacement.ClearFormatting
        .Replacement.Text = IIf(IsNull(Trim(pRs!EstadoCivil)), "", Trim(pRs!EstadoCivil))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO18"
        .Replacement.ClearFormatting
        .Replacement.Text = IIf(IsNull(Trim(pRs!nacionalidad)), "", Trim(pRs!nacionalidad))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO19"
        .Replacement.ClearFormatting
        .Replacement.Text = IIf(IsNull(Trim(pRs!edad)), "", Trim(pRs!edad))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO20"
        .Replacement.ClearFormatting
        .Replacement.Text = Day(pRs!fec_ini) & " " & "de " & Format(pRs!fec_ini, "MMMM") & " " & "del " & Format(pRs!fec_ini, "yyyy")
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO22"
        .Replacement.ClearFormatting
        If IsNull(pRs!meses) Then
            .Replacement.Text = ""
        Else
           Dim xFecFin As Date
           xFecFin = DateAdd("m", pRs!meses, pRs!fec_ini)
           .Replacement.ClearFormatting
           .Replacement.Text = Day(xFecFin) & " " & "de " & Format(xFecFin, "MMMM") & " " & "del " & Format(xFecFin, "yyyy")
        End If
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With

    
    With W.Selection.FIND
        .ClearFormatting
        .Text = "DATO21"
        .Replacement.ClearFormatting
        .Replacement.Text = IIf(IsNull(Trim(pRs!profesion)), "", Trim(pRs!profesion))
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
    
    W.ActiveDocument.SaveAs FileName:=pNomFIle
    W.Documents.Close SaveChanges:=wdDoNotSaveChanges
    W.Documents.Open FileName:=pNomFIle
    
    Set W = Nothing
 Exit Sub
MsgErr:
Set W = Nothing
MsgBox Err.Number & " - " & Err.Description, vbCritical, Me.Caption

End Sub





Public Function Validar_Suspension4taCategoria() As Boolean
With rsSuspension4ta
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                If Trim(!numop & "") = "" Then
                    SSTab1.Tab = 6
                    MsgBox "Ingrese Nro de documento del prestador de Servicios 4ta Categoria", vbExclamation, Me.Caption
                    DgrdrsSuspension4ta.Col = 0
                    DgrdrsSuspension4ta.SetFocus
                    GoTo Salir:
                ElseIf Not IsDate(!fecha & "") Then
                    SSTab1.Tab = 6
                    MsgBox "Ingrese fecha de prestación valida", vbExclamation, Me.Caption
                    DgrdrsSuspension4ta.Col = 1
                    DgrdrsSuspension4ta.SetFocus
                    GoTo Salir:
                ElseIf Trim(!ejercicio & "") = "" Then
                    SSTab1.Tab = 6
                    MsgBox "Ingrese Ejercicio de prestación", vbExclamation, Me.Caption
                    DgrdrsSuspension4ta.Col = 2
                    DgrdrsSuspension4ta.SetFocus
                    GoTo Salir:
                ElseIf Trim(!medio & "") = "" Then
                    SSTab1.Tab = 6
                    MsgBox "Elija medio de presentación", vbExclamation, Me.Caption
                    DgrdrsSuspension4ta.Col = 3
                    DgrdrsSuspension4ta.SetFocus
                    GoTo Salir:
                End If
            .MoveNext
        Loop
    End If
    Validar_Suspension4taCategoria = True
    Exit Function
Salir:
Validar_Suspension4taCategoria = False
End With
End Function



Public Sub Carga_Establecimientos(ByVal pSwNuevo As Boolean)

 If rsTipEst.RecordCount > 0 Then rsTipEst.MoveFirst
    Do While Not rsTipEst.EOF
       rsTipEst.Delete
       rsTipEst.MoveNext
    Loop
    
Dim Rq As ADODB.Recordset
Dim Sql As String
Sql = "usp_pla_consulta_establecimientos '" & wcia & "'"
If fAbrRst(Rq, Sql) Then
    Do While Not Rq.EOF
            rsTipEst.AddNew
            rsTipEst!Add = False
            rsTipEst!Id = Trim(Rq!RUC) & Trim(Rq!cod_establecimiento)
            rsTipEst!RUC = Trim(Rq!RUC)
            rsTipEst!codest = Trim(Rq!cod_establecimiento)
            rsTipEst!nomest = Trim(Rq!nom_establecimiento)
            rsTipEst!tipest = Trim(Rq!tipo_establecimiento)
            rsTipEst!razsoc = Trim(Rq!razsoc)
            
        Rq.MoveNext
    Loop
End If
If pSwNuevo = False Then
   
    Sql = "SELECT ruc,cod_establecimiento FROM plaCiaEstablecimientos_Labora_Trabajador "
    Sql = Sql & " where cod_cia='" & wcia & "' and placod='" & Trim(Txtcodpla.Text) & "' and status<>'*'"
    If fAbrRst(Rq, Sql) Then
        Dim xCriterio As String
        
        Do While Not Rq.EOF
            xCriterio = Trim(Rq!RUC) & Trim(Rq!cod_establecimiento)
            If rsTipEst.RecordCount > 0 Then rsTipEst.MoveFirst
            rsTipEst.FIND "id='" & xCriterio & "'"
            If Not rsTipEst.EOF Then
                rsTipEst!Add = True
                'rsTipEst.Update
            Else
                If Not rsTipEst.EOF Then rsTipEst!Add = False
            End If
            Rq.MoveNext
        Loop
    End If
End If
'Me.DgrdrTipEst.Rebind
Rq.Close
Set Rq = Nothing
End Sub

Private Function Verifica_permiso() As Boolean
Verifica_permiso = True

If UCase(wuser) <> "SA" Then
   Sql$ = "Select grabar from users_menu where name_menu='Mnupersona' and name_user='" & wuser & "' and sistema='04'"
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   Set Rs = cn.Execute(Sql$)
   If Rs.RecordCount > 0 Then
      If Rs!Grabar <> 1 Then
         MsgBox "Usuario no tiene Permisio para Grabar", vbInformation
         Verifica_permiso = False: If Rs.State = 1 Then Rs.Close: Exit Function
      End If
   Else
      MsgBox "Usuario no tiene Permisio para grabar", vbInformation
      Verifica_permiso = False: If Rs.State = 1 Then Rs.Close: Exit Function
   End If
   If Rs.State = 1 Then Rs.Close
End If


End Function
Private Sub Carga_Educacion()
Dim CodRegimen As String
CodRegimen = "0"
If OptPublica.Value = True Then CodRegimen = "1"
If OptPrivada.Value = True Then CodRegimen = "2"


Sql$ = "select Distinct('00') as Cod_maestro2, (replicate ('0',(2-len(COD_TIPO_INST)))+convert(varchar,COD_TIPO_INST)) + '-' + DESC_TIPO_INST as descrip from SUNAT_INSTITUCIONES_EDUCATIVAS Where Cod_Regimen=" & CodRegimen & " order BY descrip"

Dim rsRegimen  As ADODB.Recordset
Set rsRegimen = cn.Execute(Sql$)

CmbCentroForma.Clear
Do While Not rsRegimen.EOF
   CmbCentroForma.AddItem Trim(rsRegimen!DESCRIP)
   CmbCentroForma.ItemData(CmbCentroForma.NewIndex) = Trim(rsRegimen!cod_maestro2)
   rsRegimen.MoveNext
Loop
rsRegimen.Close: Set rsRegimen = Nothing
End Sub
Sub Uniformes_1()
'Uniformes
'Sql$ = "select * from plaUniforme where cia='" & wcia & "' and placod='" & CODPLA & "' and status<>'*'"
If rsunif.RecordCount > 0 Then rsunif.MoveFirst
Do While Not rsunif.EOF
   rsunif.Delete
   rsunif.MoveNext
Loop

Sql$ = "select * from plaUniformes where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   Rs.MoveFirst
   Do While Not Rs.EOF
       rsunif.AddNew
      'rsunif!botas = rs!botas
      'rsunif!camisa = rs!camisa
      'rsunif!pantalon = rs!pantalon
      'rsunif!polo = rs!polo
       rsunif!referencia = Rs!Descripcion
       rsunif!Talla = Trim(Rs!Talla)
       Rs.MoveNext
   Loop
Else
    rsunif.AddNew
    rsunif(0) = "BOTAS"
    rsunif.AddNew
    rsunif(0) = "CAMISA"
    rsunif.AddNew
    rsunif(0) = "PANTALON"
    rsunif.AddNew
    rsunif(0) = "POLO"
    rsunif.AddNew
    rsunif(0) = "CHOMPA"
    rsunif.AddNew
    rsunif(0) = "MANDIL"
End If
If Rs.State = 1 Then Rs.Close

End Sub
