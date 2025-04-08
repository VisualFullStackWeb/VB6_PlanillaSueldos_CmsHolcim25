VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AE8EA9F8-D7E0-452E-8699-0C3BA0F0EBBE}#1.0#0"; "Pryobjetos.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmGrdPersonal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Lista de Trabajadores «"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   Icon            =   "FrmGrdPersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000013&
      Height          =   6030
      Left            =   0
      TabIndex        =   1
      Top             =   585
      Width           =   2415
      Begin VB.TextBox TxtNroDoc 
         Height          =   285
         Left            =   240
         TabIndex        =   51
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox CmbTipDoc 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox TxtCodTrab 
         Height          =   285
         Left            =   1560
         TabIndex        =   47
         Top             =   360
         Width           =   750
      End
      Begin VB.ComboBox CmbPago 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   4920
         Width           =   2055
      End
      Begin VB.ComboBox CmbCargo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   4320
         Width           =   2055
      End
      Begin VB.ComboBox CmbArea 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3720
         Width           =   2055
      End
      Begin VB.ComboBox CmbTipo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3000
         Width           =   2055
      End
      Begin VB.ComboBox CmbPlanta 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   2055
      End
      Begin Threed.SSCommand SSCommand9 
         Height          =   375
         Left            =   240
         TabIndex        =   65
         Top             =   5520
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Exportar Trabajador"
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":030A
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. de trabajadores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   1605
      End
      Begin VB.Label LblNumTra 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   1800
         TabIndex        =   69
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Documento"
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
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento"
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
         Index           =   2
         Left            =   240
         TabIndex        =   48
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Trabajador"
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
         Index           =   1
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
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
         Left            =   240
         TabIndex        =   15
         Top             =   4680
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
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
         TabIndex        =   13
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
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
         TabIndex        =   11
         Top             =   3480
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Trabajador"
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
         TabIndex        =   9
         Top             =   2760
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Planta"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   6975
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Lblfecha"
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
         Left            =   10065
         TabIndex        =   6
         Top             =   120
         Width           =   1575
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
         Left            =   1515
         TabIndex        =   4
         Top             =   135
         Width           =   825
      End
   End
   Begin Threed.SSPanel Panelprogress 
      Height          =   735
      Left            =   3120
      TabIndex        =   42
      Top             =   3000
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FloodColor      =   -2147483637
      Alignment       =   6
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   6600
      Width           =   11775
      Begin VB.CheckBox Chkpc 
         BackColor       =   &H00808080&
         Caption         =   "Personal de confianza"
         Height          =   255
         Left            =   6480
         TabIndex        =   71
         Top             =   360
         Width           =   3135
      End
      Begin VB.ComboBox CmbCatTrab 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   120
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         Caption         =   "Cesados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6480
         TabIndex        =   52
         Top             =   120
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Txtfechaal 
         Height          =   255
         Left            =   10200
         TabIndex        =   19
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria del Trabajador"
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
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   120
         Width           =   2025
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Trabajadores al"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   7920
         TabIndex        =   18
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6015
      Left            =   2400
      TabIndex        =   2
      Top             =   600
      Width           =   9375
      Begin Pryobjetos.Usertxt TxtApMat 
         Height          =   285
         Left            =   1950
         TabIndex        =   45
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
      End
      Begin Pryobjetos.Usertxt TxtApPat 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
      End
      Begin VB.TextBox TxtseNom 
         Height          =   285
         Left            =   7440
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtPrnom 
         Height          =   285
         Left            =   5610
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtApCas 
         Height          =   285
         Left            =   3780
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc AdoPersonal 
         Height          =   375
         Left            =   960
         Top             =   1800
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   661
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
         Caption         =   "AdoPersonal"
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
      Begin VB.TextBox TxtApMat1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DgrdPersonal 
         Bindings        =   "FrmGrdPersonal.frx":0326
         Height          =   5295
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9340
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
         ColumnCount     =   6
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
            DataField       =   "fingreso"
            Caption         =   "Fec. Ingreso"
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
            DataField       =   "fcese"
            Caption         =   "Fec. Cese"
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
            DataField       =   "nro_doc"
            Caption         =   "Nro Documento"
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
            DataField       =   "cia"
            Caption         =   "Cia"
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
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4215.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seg. Nombre"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7440
         TabIndex        =   28
         Top             =   0
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pri. nombre"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5640
         TabIndex        =   27
         Top             =   0
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Casada"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3840
         TabIndex        =   26
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Materno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   25
         Top             =   0
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ap. Paterno"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.Frame FrmExporta 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      Height          =   1815
      Left            =   2640
      TabIndex        =   55
      Top             =   2400
      Visible         =   0   'False
      Width           =   8895
      Begin VB.ComboBox CmbCiaExporta 
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
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox TxtCodExporta 
         Height          =   285
         Left            =   5400
         TabIndex        =   59
         Top             =   1320
         Width           =   990
      End
      Begin MSMask.MaskEdBox TxtFecExporta 
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
         Left            =   6480
         TabIndex        =   56
         Top             =   1320
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   615
         Left            =   8040
         TabIndex        =   63
         Top             =   840
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   1085
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":0340
      End
      Begin Threed.SSCommand SSCommand8 
         Height          =   375
         Left            =   8520
         TabIndex        =   64
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "X"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":0792
      End
      Begin VB.Label LblNombre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   960
         TabIndex        =   67
         Top             =   600
         Width           =   6645
      End
      Begin VB.Label LblCodExportar 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia a Exportar"
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
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Cod."
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
         Index           =   4
         Left            =   5400
         TabIndex        =   60
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Ingreso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   6600
         TabIndex        =   58
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "EXPORTA TRABAJADOR"
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
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   8295
      End
   End
   Begin VB.Frame FrameReport 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   6495
      Left            =   0
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CheckBox ChkAllCia 
         Caption         =   "Todas la Empresas"
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
         Left            =   9480
         TabIndex        =   68
         Top             =   6200
         Width           =   2055
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   615
         Left            =   5160
         TabIndex        =   41
         ToolTipText     =   "Generar Reporte"
         Top             =   4440
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   1085
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":07AE
      End
      Begin VB.ComboBox Cmblistado 
         Height          =   315
         Left            =   6075
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   120
         Width           =   5490
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   615
         Left            =   5160
         TabIndex        =   38
         ToolTipText     =   "Salir"
         Top             =   5280
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   1085
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":0D48
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6255
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   4815
         Begin MSDataGridLib.DataGrid DgdField 
            Height          =   6075
            Left            =   60
            TabIndex        =   37
            Top             =   60
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   10716
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
               DataField       =   "campo"
               Caption         =   "campo"
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
               DataField       =   "nitem"
               Caption         =   "nitem"
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
                  ColumnWidth     =   3929.953
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
      Begin Threed.SSCommand SSCommand4 
         Height          =   615
         Left            =   5160
         TabIndex        =   30
         ToolTipText     =   "Trasladar"
         Top             =   1440
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1058
         _ExtentY        =   1085
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":12E2
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   615
         Left            =   5145
         TabIndex        =   31
         ToolTipText     =   "Trasladar"
         Top             =   2400
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   1085
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":187C
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   615
         Left            =   5145
         TabIndex        =   32
         ToolTipText     =   "Bajar una posicion"
         Top             =   3480
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   1085
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":1E16
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Left            =   5145
         TabIndex        =   33
         ToolTipText     =   "Subir una posicion"
         Top             =   480
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   1085
         _StockProps     =   78
         BevelWidth      =   1
         Picture         =   "FrmGrdPersonal.frx":23B0
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5775
         Left            =   6000
         TabIndex        =   34
         Top             =   600
         Width           =   5655
         Begin MSDataGridLib.DataGrid DgdReport 
            Height          =   5475
            Left            =   60
            TabIndex        =   35
            Top             =   60
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   9657
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
            ColumnCount     =   6
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
               DataField       =   "campo"
               Caption         =   "Campo"
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
            BeginProperty Column03 
               DataField       =   "orden"
               Caption         =   "Orden"
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
               DataField       =   "quiebre"
               Caption         =   "Quiebre"
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
               DataField       =   "nitem"
               Caption         =   "nitem"
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
                  ColumnWidth     =   4350.047
               EndProperty
               BeginProperty Column01 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column02 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   524.976
               EndProperty
               BeginProperty Column04 
                  Object.Visible         =   0   'False
                  ColumnWidth     =   705.26
               EndProperty
               BeginProperty Column05 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Listado"
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
         Left            =   5160
         TabIndex        =   39
         Top             =   120
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmGrdPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VPlanta As String
Dim VTipo As String
Dim VArea As String
Dim VCargo As String
Dim VProfesion As String
Dim VFPago As String
Dim rsField As New Recordset
Dim rsReport As New Recordset
Dim rsLista As New Recordset
Dim vID_Report As String
Dim Carga_Reports As String
Dim mCadReport As String
Dim rs2 As ADODB.Recordset
Dim rsClon As New Recordset
Dim wciaExporta As String

Private Sub Check1_Click()
Procesa_Personal (False)
End Sub

Private Sub Chkpc_Click()
Procesa_Personal (False)
End Sub

Private Sub CmbArea_Click()
VArea = fc_CodigoComboBox(CmbArea, 2)
Procesa_Personal (False)
End Sub
Private Sub CmbCargo_Click()
VCargo = fc_CodigoComboBox(CmbCargo, 3)
Procesa_Personal (False)
End Sub

Private Sub CmbCatTrab_Click()
If CmbCatTrab.ListIndex > -1 Then Procesa_Personal (False)
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Carga_Combos
Procesa_Personal (False)
End Sub

Private Sub Cmblistado_Click()
vID_Report = fc_CodigoComboBox(Cmblistado, 4)
Carga_Report
End Sub

Private Sub CmbPago_Click()
VFPago = fc_CodigoComboBox(CmbPago, 2)
Procesa_Personal (False)
End Sub

Private Sub CmbPlanta_Click()
VPlanta = fc_CodigoComboBox(CmbPlanta, 2)
Procesa_Personal (False)
End Sub

Private Sub CmbProfesion_Click()
VProfesion = fc_CodigoComboBox(CmbProfesion, 2)
Procesa_Personal (False)
End Sub

Private Sub CmbTipDoc_Click()
If CmbTipDoc.ListIndex > -1 Then
    Procesa_Personal (False)
End If
End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(CmbTipo, 2)
Procesa_Personal (False)
End Sub

Private Sub DgdReport_AfterColEdit(ByVal ColIndex As Integer)
If DgdReport.Columns(3).Text <> "" Then
   DgdReport.Columns(3) = Format(DgdReport.Columns(3), "###,###")
End If
End Sub

Private Sub DgrdPersonal_DblClick()
If DgrdPersonal.Row < 0 Then Exit Sub
Call Frmpersona.Carga_Trabajador(Trim(DgrdPersonal.Columns(0)))
End Sub

Private Sub DgrdPersonal_HeadClick(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 0
            AdoPersonal.Recordset.Sort = "placod"
       Case Is = 1
            AdoPersonal.Recordset.Sort = "nombre"
       Case Is = 2
            AdoPersonal.Recordset.Sort = "fingreso"
End Select
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Height = 7755
Me.Width = 11865
'If wGrupoPla = "01" Then ChkAllCia.Visible = True Else ChkAllCia.Visible = False
ChkAllCia.Visible = True

Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rCarCbo(CmbCiaExporta, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Crea_Rs
End Sub
Private Sub Carga_Combos()
Call fc_Descrip_Maestros2("01032", "", CmbTipDoc, True)
Call rUbiIndCmbBox(CmbTipDoc, "01", "00")

'Call fc_Descrip_Maestros2("01008", "", CmbCargo)
Sql$ = "SELECT COD_MAESTRO3,DESCRIP FROM MAESTROS_31 WHERE CIAMAESTRO = '" & wcia & "055" & "' AND STATUS != '*' ORDER BY DESCRIP"
Call rCarCbo(CmbCargo, Sql$, "C", "000")

CmbCargo.AddItem "TOTAL"
CmbCargo.ItemData(CmbCargo.NewIndex) = "999"
Call fc_Descrip_Maestros2("01070", "", CmbPlanta)
CmbPlanta.AddItem "TOTAL"
CmbPlanta.ItemData(CmbPlanta.NewIndex) = "99"

'Call fc_Descrip_Maestros2("01044", "", CmbArea)

Sql$ = "SELECT Codigo,Descripcion FROM Pla_Ccostos where cia='" & wcia & "' and status<>'*' ORDER BY DESCRIPCION"
Call rCarCbo(CmbArea, Sql$, "C", "000")

CmbArea.AddItem "TOTAL"
CmbArea.ItemData(CmbArea.NewIndex) = "99"

Call fc_Descrip_Maestros2("01055", "", CmbTipo, True)


CmbTipo.AddItem "TOTAL"
CmbTipo.ItemData(CmbTipo.NewIndex) = "99"
Call fc_Descrip_Maestros2("01060", "", CmbPago, True)
CmbPago.AddItem "TOTAL"
Call fc_Descrip_Maestros2("01141", "", CmbCatTrab)
CmbCatTrab.AddItem "TOTAL"

CmbPago.ItemData(CmbPago.NewIndex) = "99"
'If wGrupoPla = "01" Then
Carga_Reports = "select distinct(id_report),name_report from plareports where tipo='01' and status<>'*' and user_crea='" & wuser & "' order by id_report"
'Else
'   Carga_Reports = "select distinct(id_report),name_report from plareports where cia='" & wcia & "' and tipo='01' and status<>'*' and user_crea='" & wuser & "' order by id_report"
'End If
Call rCarCbo(Cmblistado, Carga_Reports, "C", "0000")
Cmblistado.AddItem "NUEVO"

End Sub
Public Sub Procesa_Personal(mAll As Boolean)
Dim mf As String
Dim mcese As String
Dim f1 As String
Dim ls_CadPersonalConfianza As String
ls_CadPersonalConfianza = ""
If Me.Chkpc.Value Then
    ls_CadPersonalConfianza = " and placod in (select placod FROM [dbo].[Pla_Personal_Confianza]) "
End If

If Check1.Value = 1 Then mcese = "S" Else mcese = ""
If Txtfechaal.Text = "__/__/____" Then f1 = "": mf = Format("01/01/2000", FormatFecha) Else f1 = "S": mf = Format(Txtfechaal.Text, FormatFecha)
If VPlanta = "99" Or CmbPlanta.ListIndex < 0 Then VPlanta = ""
If VTipo = "99" Or CmbTipo.ListIndex < 0 Then VTipo = ""
If VArea = "99" Or CmbArea.ListIndex < 0 Then VArea = ""
If VCargo = "999" Or CmbCargo.ListIndex < 0 Then VCargo = ""
'If VProfesion = "99" Or CmbProfesion.ListIndex < 0 Then VProfesion = ""
If VFPago = "99" Or CmbPago.ListIndex < 0 Then VFPago = ""
Dim xCatTrab As String
If Trim(CmbCatTrab.Text) <> "TOTAL" And Trim(CmbCatTrab.Text) <> "" Then
    xCatTrab = " and cat_trab='" & fc_CodigoComboBox(CmbCatTrab, 2) & "' "
End If

Sql$ = nombre()
Dim xDocumento As String
Dim xTipDoc As String
If Trim(TxtNroDoc.Text) <> "TOTAL" And Trim(TxtNroDoc.Text) <> "" Then
    xTipDoc = fc_CodigoComboBox(CmbTipDoc, 2)
    xDocumento = " and tipo_doc='" & xTipDoc & "' and nro_doc like '" & Me.TxtNroDoc.Text & "%'"
End If

If mAll Then
   Sql$ = Sql$ & "placod,fingreso,fcese,nro_doc,cia From planillas where status<>'*' "
   If f1 = "S" Then
       If mcese <> "S" Then
          Sql$ = Sql$ + " and fingreso<='" & mf & FormatTimef & "'  and fcese is null Order by placod,ap_pat"
       Else
          Sql$ = Sql$ + " and fingreso<='" & mf & FormatTimef & "' Order by placod,ap_pat"
      End If
   Else
       If mcese <> "S" Then
          Sql$ = Sql$ + " and fcese is null Order by ap_pat"
       Else
          Sql$ = Sql$ + " Order by ap_pat"
       End If
   End If
Else
   Sql$ = Sql$ & "placod,fingreso,fcese,nro_doc,cia " _
     & "From planillas " _
     & "where cia='" & wcia & "' and placod like '" & Trim(TxtCodTrab.Text) & "%' and status<>'*' and planta like '" & RTrim(VPlanta) + "%" & "' and tipotrabajador like '" & RTrim(VTipo) + "%" & "' and " _
     & "cargo like '" & RTrim(VCargo) + "%" & "' and profesion like '" & RTrim(VProfesion) + "%" & "' and tipopago like '" & RTrim(VFPago) + "%" & "' and " _
     & "ap_pat like '" & Trim(TxtApPat.Text) + "%" & "'  and ap_mat like '" & Trim(TxtApMat.Text) + "%" & "'  and ap_cas like '" & Trim(TxtApCas.Text) + "%" & "' and nom_1 like '" & Trim(TxtPrnom.Text) + "%" & "'  and nom_2 like '" & Trim(TxtseNom.Text) + "%" & "'"

   Sql$ = Sql$ & xCatTrab
     
   Sql$ = Sql$ & xDocumento
   
   'add jcms 021021
   Sql$ = Sql$ & ls_CadPersonalConfianza
   
   If Trim(VArea & "") <> "" Then
      Sql$ = Sql$ + " and (placod in (select placod from planilla_ccosto where cia='" & wcia & "' and ccosto='" & Trim(VArea & "") & "' and status<>'*')) "
   End If
   
   
   If f1 = "S" Then
       If mcese <> "S" Then
          Sql$ = Sql$ + " and fingreso<='" & mf & FormatTimef & "'  and fcese is null Order by ap_pat"
       Else
          Sql$ = Sql$ + " and fingreso<='" & mf & FormatTimef & "' Order by ap_pat"
      End If
   Else
       If mcese <> "S" Then
          Sql$ = Sql$ + " and fcese is null Order by ap_pat"
       Else
          Sql$ = Sql$ + " Order by ap_pat"
       End If
   End If
End If
cn.CursorLocation = adUseClient
 
Set AdoPersonal.Recordset = cn.Execute(Sql$, 64)
LblNumTra.Caption = "0"
If AdoPersonal.Recordset.RecordCount > 0 Then AdoPersonal.Recordset.MoveFirst: LblNumTra.Caption = AdoPersonal.Recordset.RecordCount

DgrdPersonal.Refresh
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not (Rs Is Nothing) Then
    If Rs.State = 1 Then Rs.Close
End If
Set DgdField.DataSource = Nothing
If rsField.State = 1 Then rsField.Close
Set rsField = Nothing
Set DgdReport.DataSource = Nothing
If rsReport.State = 1 Then rsReport.Close
Set rsReport = Nothing
End Sub

Private Sub SSCommand1_Click()
Dim MREC As Integer
Dim vIt As Integer
If rsReport.RecordCount <= 1 Then Exit Sub
If DgdReport.Columns(2) = "1" Then Exit Sub
MREC = rsReport.AbsolutePosition
DgdReport.Columns(2) = DgdReport.Columns(2) - 1
vIt = DgdReport.Columns(2)
If DgdReport.Row > 0 Then DgdReport.Row = DgdReport.Row - 1
DgdReport.Columns(2) = vIt + 1
DgdReport.Columns(2) = DgdReport.Columns(2) + 1
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
DgdReport.Columns(2) = DgdReport.Columns(2) + 1
vIt = DgdReport.Columns(2)
DgdReport.Row = DgdReport.Row + 1
DgdReport.Columns(2) = vIt - 1
If DgdReport.Row > 0 Then DgdReport.Row = DgdReport.Row - 1
rsReport.Sort = "item"
DgdReport.Refresh
rsReport.AbsolutePosition = MREC + 1
End Sub

Private Sub SSCommand3_Click()
If rsReport.RecordCount <= 0 Then Exit Sub
rsField.AddNew
rsField!Descripcion = rsReport!Descripcion
rsField!Campo = rsReport!Campo
rsField!nitem = rsReport!nitem
rsField!numfield = rsReport!numfield
rsField!directo = rsReport!directo
rsField!Tabla = rsReport!Tabla
rsField!referencia = rsReport!referencia
rsField!CORTO = rsReport!CORTO
rsReport.Delete
End Sub

Private Sub SSCommand4_Click()
If rsField.RecordCount <= 0 Then Exit Sub
Dim rsValidar As ADODB.Recordset
If Trim(rsField!Descripcion & "") = "REMUNERACIONES BASICAS" Then
    Sql$ = "select * from users where cod_cia='" & wcia & "' and login='" & wuser & "' and sistema='04' and status !='*'"
    If fAbrRst(rsValidar, Sql) Then
        If rsValidar!autorizar_remuneracion = False Then
            MsgBox "Personal No Autorizado, a consultar este tipo de información", vbCritical, Me.Caption
            Exit Sub
        End If
    Else
        MsgBox "Personal No Autorizado, a consultar este tipo de información", vbCritical, Me.Caption
        Exit Sub
    End If
    If rsValidar.State = 1 Then rsValidar.Close: Set rsValidar = Nothing
End If


If rsReport.RecordCount > 0 Then rsReport.MoveLast
rsReport.AddNew
rsReport!Descripcion = rsField!Descripcion
rsReport!Campo = rsField!Campo
rsReport!nitem = rsField!nitem
rsReport!numfield = rsField!numfield
rsReport!directo = rsField!directo
rsReport!Tabla = rsField!Tabla
rsReport!referencia = rsField!referencia
rsReport!CORTO = IIf(Trim(rsField!CORTO) = "UNIFORMES", "BOTAS;CAMISA;PANTALON;POLO;CHOMPA;MANDIL", rsField!CORTO)
rsReport!Item = rsReport.RecordCount
rsField.Delete
End Sub

Private Sub SSCommand5_Click()
FrameReport.Visible = False
End Sub

Private Sub SSCommand6_Click()
If ChkAllCia.Value Then Procesa_Personal (True)
Genera_Reporte
If ChkAllCia.Value Then Procesa_Personal (False)
End Sub

Private Sub SSCommand7_Click()
On Error GoTo ErrorTrans
Dim NroTrans As Integer
NroTrans = 0
TxtCodExporta.Text = UCase(Trim(TxtCodExporta.Text & ""))
If Trim(LblCodExportar.Caption) = "" Then
   MsgBox "Debse Seleccionar un trabajador", vbInformation
   Exit Sub
End If
wciaExporta = Trim(Right("00" & CmbCiaExporta.ItemData(CmbCiaExporta.ListIndex), 2))
If wcia = wciaExporta Then
   MsgBox "Se debe exportar a una Cia Distinta", vbInformation
   Exit Sub
End If
If Trim(TxtCodExporta.Text) = "" Then
   MsgBox "Debse Ingresar un Codigo", vbInformation
   Exit Sub
End If


Sql$ = "select rtrim(ap_pat)+' '+rtrim(ap_mat)+' '+rtrim(nom_1)+' '+rtrim(nom_2) from planillas where cia='" & wciaExporta & "' and placod='" & Trim(TxtCodExporta.Text) & "' "
If (fAbrRst(Rs, Sql)) Then
   MsgBox "Codigo ya registrado con el trabajdor " + RTrim(Rs(0) & ""), vbInformation
   Rs.Close: Set Rs = Nothing
   Exit Sub
End If

Me.MousePointer = 11

cn.BeginTrans
NroTrans = 1

Sql$ = "Usp_Exporta_Trabajador '" & wcia & "','" & LblCodExportar.Caption & "','" & wciaExporta & "','" & Trim(TxtCodExporta.Text) & "','" & TxtFecExporta.Text & "'"
cn.Execute Sql$, 64

cn.CommitTrans
Me.MousePointer = 0

FrmExporta.Visible = False
DgrdPersonal.Enabled = True

Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If
MsgBox Err.Description, vbCritical, Me.Caption
Me.MousePointer = 0
FrmExporta.Visible = False
DgrdPersonal.Enabled = True

End Sub

Private Sub SSCommand8_Click()
FrmExporta.Visible = False
DgrdPersonal.Enabled = True
End Sub

Private Sub SSCommand9_Click()
LblCodExportar.Caption = ""
LblNombre.Caption = ""
LblCodExportar.Caption = Trim(DgrdPersonal.Columns(0) & "")
LblNombre.Caption = Trim(DgrdPersonal.Columns(1) & "")
If LblCodExportar.Caption = "" Then Exit Sub
TxtCodExporta.Text = ""
TxtFecExporta = Date
CmbCiaExporta.ListIndex = -1
wciaExporta = ""
FrmExporta.Visible = True
End Sub

Private Sub TxtApCas_Change()
Procesa_Personal (False)
End Sub

Private Sub TxtApMat_Change()
Procesa_Personal (False)
End Sub

Private Sub TxtApPat_Change()
Procesa_Personal (False)
End Sub



Private Sub TxtCodTrab_Change()
Procesa_Personal (False)
End Sub

Private Sub Txtfechaal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmbPlanta.SetFocus
End Sub

Private Sub Txtfechaal_LostFocus()
Procesa_Personal (False)
End Sub

Private Sub TxtNroDoc_Change()
Procesa_Personal (False)
End Sub

Private Sub TxtPrnom_Change()
Procesa_Personal (False)
End Sub

Private Sub TxtseNom_Change()
Procesa_Personal (False)
End Sub

Private Sub Crea_Rs()
    If rsField.State = 1 Then rsField.Close
    rsField.Fields.Append "descripcion", adChar, 60, adFldIsNullable
    rsField.Fields.Append "campo", adChar, 60, adFldIsNullable
    rsField.Fields.Append "nitem", adInteger, 4, adFldIsNullable
    rsField.Fields.Append "numfield", adInteger, 4, adFldIsNullable
    rsField.Fields.Append "directo", adChar, 1, adFldIsNullable
    rsField.Fields.Append "tabla", adChar, 25, adFldIsNullable
    rsField.Fields.Append "referencia", adChar, 25, adFldIsNullable
    rsField.Fields.Append "corto", adVarChar, 20, adFldIsNullable
    rsField.Open
    Set DgdField.DataSource = rsField
    
    If rsReport.State = 1 Then rsReport.Close
    rsReport.Fields.Append "descripcion", adChar, 60, adFldIsNullable
    rsReport.Fields.Append "campo", adChar, 60, adFldIsNullable
    rsReport.Fields.Append "item", adInteger, 4, adFldIsNullable
    rsReport.Fields.Append "orden", adInteger, 4, adFldIsNullable
    rsReport.Fields.Append "quiebre", adChar, 1, adFldIsNullable
    rsReport.Fields.Append "nitem", adInteger, 4, adFldIsNullable
    rsReport.Fields.Append "numfield", adInteger, 4, adFldIsNullable
    rsReport.Fields.Append "directo", adChar, 1, adFldIsNullable
    rsReport.Fields.Append "tabla", adChar, 25, adFldIsNullable
    rsReport.Fields.Append "referencia", adChar, 25, adFldIsNullable
    rsReport.Fields.Append "general", adChar, 1, adFldIsNullable
    rsReport.Fields.Append "corto", adVarChar, 20, adFldIsNullable
    rsReport.Open
    Set DgdReport.DataSource = rsReport

End Sub
Public Sub Reporte_Personal()
Cmblistado.ListIndex = -1
FrameReport.Visible = True
FrameReport.ZOrder 0
If rsField.RecordCount > 0 Then
   rsField.MoveFirst
   Do While Not rsField.EOF
      rsField.Delete
      rsField.MoveNext
   Loop
End If

If rsReport.RecordCount > 0 Then
   rsReport.MoveFirst
   Do While Not rsReport.EOF
      rsReport.Delete
      rsReport.MoveNext
   Loop
End If
If wuser = "SUSANA" Then
Sql = "select * from plasetreport where reporte='PL' and tipo='01' AND ITEM!='33' order by item"
Else
Sql = "select * from plasetreport where reporte='PL' and tipo='01' order by item"
End If



If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Do While Not Rs.EOF
   rsField.AddNew
   rsField!Descripcion = Rs!Descripcion
   rsField!Campo = Rs!Campo
   rsField!nitem = Rs!Item
   rsField!numfield = Rs!numfield
   rsField!directo = Rs!directo
   rsField!Tabla = Rs!Tabla
   rsField!referencia = Rs!referencia
   rsField!CORTO = IIf(Trim(Rs!CORTO) = "UNIFORMES", "BT;CM;PT;PL;CH;MD", Rs!CORTO)
   Rs.MoveNext
Loop
If rsField.RecordCount > 0 Then rsField.MoveFirst
Rs.Close

End Sub
Public Sub Grabar_Report()
On Error GoTo ErrorTrans
Dim NroTrans As Integer

Dim vName As String
Dim VNew As Boolean
Dim vOrden As Integer
NroTrans = 0
If rsReport.RecordCount <= 0 Then Exit Sub

rsReport.MoveFirst
VNew = False
Mgrab = MsgBox("Seguro de Grabar Reporte", vbYesNo + vbQuestion, "Reporte de Trabajadores")
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
   Sql = "select max(id_report) from plareports where tipo='01' and status<>'*'"
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
Sql = "update plareports set status='*' where tipo='01' and id_report='" & vID_Report & "' and status<>'*'"
cn.Execute Sql
Do While Not rsReport.EOF
   If Not IsNumeric(rsReport!Orden) Then vOrden = 0 Else vOrden = rsReport!Orden
   Sql = "insert into plareports values('" & wcia & "','01','" & vID_Report & "','" & Trim(Left(vName, 60)) & "','" & Trim(rsReport!Campo) & "'," & rsReport!Item & "," & vOrden & ",'" & rsReport!Quiebre & "','','" & wuser & "'," _
       & "" & rsReport!nitem & "," & rsReport!numfield & ",'" & rsReport!directo & "','" & rsReport!Tabla & "','" & rsReport!referencia & "','" & rsReport!CORTO & "'," & FechaSys & ")"
   cn.Execute Sql
   rsReport.MoveNext
Loop

cn.CommitTrans

Call rCarCbo(Cmblistado, Carga_Reports, "C", "0000")
Cmblistado.AddItem "NUEVO"
Call rUbiIndCmbBox(Cmblistado, vID_Report, "0000")

Screen.MousePointer = vbDefault
Exit Sub
ErrorTrans:
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

Sql = "select s.descripcion,r.* from plareports r,plasetreport s " _
& "where r.tipo='01' and r.id_report='" & vID_Report & "' and r.status<>'*' " _
& "and s.item=r.nitem and r.user_crea='" & wuser & "' order by R.item"

If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
Do While Not Rs.EOF
   rsReport.AddNew
   rsReport!Descripcion = Rs!Descripcion & ""
   rsReport!Campo = Rs!Campo & ""
   rsReport!Item = Rs!Item
   rsReport!nitem = Rs!nitem
   rsReport!numfield = Rs!numfield
   rsReport!directo = Rs!directo
   rsReport!Tabla = Rs!Tabla
   rsReport!referencia = Rs!referencia
   rsReport!CORTO = IIf(Trim(Rs!CORTO) = "UNIFORMES", "BOTAS;CAMISA;PANTALON;CAMISA;POLO;CHOMPA;MANDIL", Rs!CORTO)
   
   If Rs!Orden <> 0 Then rsReport!Orden = Rs!Orden
   rsReport!Quiebre = Rs!Quiebre & ""
   Rs.MoveNext
Loop
Rs.Close
End Sub
Private Sub Genera_Reporte()
Dim mcamp As String
Dim mcadmae As String
Dim W As Integer
Screen.MousePointer = vbArrowHourglass
Set rsClon = rsReport.Clone
If rsLista.State = 1 Then rsLista.Close
If rsClon.RecordCount > 0 Then rsClon.MoveFirst
W = 0
Panelprogress.Visible = True
Panelprogress.ZOrder 0
Me.Refresh
Barra.Max = rsClon.RecordCount
Do While Not rsClon.EOF
   Barra.Value = rsClon.AbsolutePosition
   mcamp = "f" & Format(rsClon!Item, "0000")
   If rsClon!directo = "S" Then
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
   ElseIf rsClon!directo <> "S" And Trim(rsClon!Tabla) = "maestros_2" Then
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      Call Determina_Maestro(Left(rsClon!referencia, 5))
      If wMaeGen = True Then rsClon!General = "S" Else rsClon!General = "N"
   ElseIf rsClon!directo <> "S" And UCase(Trim(rsClon!Tabla)) = "MAESTROS_31" Then
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      Call Determina_Maestro(Left(rsClon!referencia, 5))
      If wMaeGen = True Then rsClon!General = "S" Else rsClon!General = "N"
      
   ElseIf rsClon!directo <> "S" And Trim(rsClon!Tabla) = "pla_areas" Then 'Gerencia
      rsClon!General = "X"
      Sql = "SELECT top 1 cod_ger, Gerencia FROM pla_areas WHERE STATUS<>'*'"
      If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
      Do While Not rs2.EOF
         mcamp = "x" & Format("01", "00")
         rsLista.Fields.Append mcamp, adVarChar, 200, adFldIsNullable
         W = W + 1
         rs2.MoveNext
      Loop
      rs2.Close
    ElseIf rsClon!directo <> "S" And Trim(rsClon!Tabla) = "pla_areas1" Then 'Departamento
      rsClon!General = "W"
      Sql = "SELECT top 1 cod_dpto, dpto FROM pla_areas WHERE STATUS<>'*'"
      If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
      Do While Not rs2.EOF
         mcamp = "w" & Format("01", "00")
         rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
         W = W + 1
         rs2.MoveNext
      Loop
      rs2.Close
   ElseIf rsClon!directo <> "S" And rsClon!nitem = 33 Then 'Remuneraciones Basicas
      rsClon!General = "B"
      Sql = "select distinct(concepto) from plaremunbase where cia='" & wcia & "' and status<>'*' order by concepto"
      If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
      Do While Not rs2.EOF
         mcamp = "b" & Format(rs2(0), "0000")
         rsLista.Fields.Append mcamp, adDecimal, 20, adFldIsNullable
         rsLista.Fields(mcamp).NumericScale = 2
         W = W + 1
         mcamp = "t" & Format(rs2(0), "0000")
         rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
         W = W + 1
         rs2.MoveNext
      Loop
      rs2.Close
   ElseIf rsClon!directo <> "S" And rsClon!nitem = 32 Then 'Telefono
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      rsClon!General = "T"
        'mgirao uNIFORMES
   ElseIf rsClon!directo <> "S" And rsClon!numfield = 82 Then 'UNiformes
      rsLista.Fields.Append mcamp, adVarChar, 80, adFldIsNullable
      W = W + 1
      rsClon!General = "K"
  'mgirao senati
   ElseIf rsClon!directo <> "S" And rsClon!numfield = 81 Then
      rsLista.Fields.Append mcamp, adVarChar, 1, adFldIsNullable
      W = W + 1
     ' rsClon!General = "SE"
      '  Sql = "select c.senati from planilla_ccosto P,Pla_CCostos C where p.cia='" & wcia & "' and p.placod='" & Trim(rs!PlaCod) & "' and p.status<>'*' and p.cia=c.cia and p.ccosto=c.codigo and c.status<>'*'"
   ElseIf rsClon!directo <> "S" And rsClon!nitem = 14 Then 'Distrito
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      rsClon!General = "U"
   ElseIf rsClon!directo <> "S" And rsClon!nitem = 51 Then 'Ubigeo
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      rsClon!General = "Y"
   ElseIf rsClon!directo <> "S" And rsClon!nitem = 15 Then 'Direccion
      rsLista.Fields.Append mcamp, adVarChar, 250, adFldIsNullable
      W = W + 1
      rsClon!General = "D"
   ElseIf rsClon!directo <> "S" And rsClon!nitem = 21 Then 'Areas
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      rsClon!General = "N"
   ElseIf rsClon!directo <> "S" And rsClon!nitem = 18 Then 'Profesiones
      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      rsClon!General = "N"
      
   ElseIf Trim(rsClon!directo) <> "S" And Trim(rsClon!Tabla) = "planilla_ccosto" Then 'CENTRO COSTO
      
       rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
      W = W + 1
      rsClon!General = "C"
      
'    ElseIf rsClon!directo <> "S" And Trim(rsClon!Tabla) = "tocupaciones" Then
'      rsLista.Fields.Append mcamp, adVarChar, 60, adFldIsNullable
'      W = W + 1
'      Call Determina_Maestro(Left(rsClon!referencia, 5))
'      If wMaeGen = True Then rsClon!General = "S" Else rsClon!General = "N"
      
   End If
   rsClon.MoveNext
Loop
rsLista.Open


If rsClon.RecordCount <= 0 Then Exit Sub
If AdoPersonal.Recordset.RecordCount > 0 Then AdoPersonal.Recordset.MoveFirst
mcadmae = ""
Barra.Max = AdoPersonal.Recordset.RecordCount
Do While Not AdoPersonal.Recordset.EOF
   Barra.Value = AdoPersonal.Recordset.AbsolutePosition
   'Sql = "select top 1  * from planillas where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*'"
   Sql = "select top 1  * "
   Sql = Sql & " "
   Sql = Sql & " from planillas where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*'"
   If (fAbrRst(Rs, Sql)) Then
      rsLista.AddNew
      rsClon.MoveFirst
      x = 0
      Do While Not rsClon.EOF
         mcamp = "f" & Format(rsClon!Item, "0000")
         If rsClon!directo = "S" Then
            'If rsClon!nitem = 6 Or rsClon!nitem = 13 Or rsClon!nitem = 22 Or rsClon!nitem = 24 Then
            If rsClon!nitem = 6 Or rsClon!nitem = 13 Or rsClon!nitem = 22 Or rsClon!nitem = 24 Or rsClon!nitem = 47 Then
               rsLista.Fields(mcamp) = "'" & Format(Trim(Rs(Val(rsClon!numfield))), "dd/mm/yyyy")
            ElseIf rsClon!nitem = 9 Then
               If Trim(Rs(64) & "") = "01" Then rsLista.Fields(mcamp) = Trim(Rs(65) & "")
            ElseIf rsClon!nitem = 28 Then
               rsLista.Fields(mcamp) = Trim(Rs(77) & "") + "-" + Trim(Rs(48) & "")
            
            ElseIf Val(rsClon!numfield) = 83 Then
               rsLista.Fields(mcamp) = IIf(Rs.Fields(Val(rsClon!numfield)) = True, "SI", "NO")
               
            Else
               rsLista.Fields(mcamp) = Trim(Rs(Val(rsClon!numfield)))
            End If
         ElseIf rsClon!General = "S" Or rsClon!General = "N" Then
            If rsClon!General = "S" Then
               mcadmae = " right(ciamaestro,3)='" & Mid(rsClon!referencia, 3, 3) & "'"
            Else
               mcadmae = " ciamaestro='" & AdoPersonal.Recordset!cia + Mid(rsClon!referencia, 3, 3) & "'"
            End If
            
            If rsClon!numfield = 31 Then
                Sql = "select descrip from maestros_31 where ciamaestro='" & AdoPersonal.Recordset!cia & "055" & "' and cod_maestro3 = '" & Rs!Cargo & "'"
            ElseIf rsClon!numfield = 33 Then
                Sql = "select c.descripcion from planilla_ccosto P,Pla_CCostos C where p.cia='" & wcia & "' and p.placod='" & Trim(Rs!PlaCod) & "' and p.status<>'*' and p.cia=c.cia and p.ccosto=c.codigo and c.status<>'*'"
   
            ElseIf rsClon!numfield = 30 Then
                'Sql = "SELECT descripcion FROM tocupaciones WHERE codigo='" & rs!profesion & "'"
                Sql = "select DISTINCT(DESC_CARRERA) from SUNAT_INSTITUCIONES_EDUCATIVAS where COD_carrera='" & Trim(Rs!Est_Cod_Carrera & "") & "'"
            ElseIf rsClon!numfield = 28 And rsClon!nitem = 35 Then
                Sql = "select codsunat from maestros_2 where" & mcadmae & " and cod_maestro2='" & Trim(Rs(Val(rsClon!numfield))) & "'"
            Else
                Dim TipAfp  As String
                TipAfp = ""
                If Trim(Rs!afptipocomision) = "F" And Trim(rsClon!referencia) = "01069" Then TipAfp = " (FLUJO)"
                If Trim(Rs!afptipocomision) = "M" And Trim(rsClon!referencia) = "01069" Then TipAfp = " (MIXTA)"
                If Trim(Rs(Val(rsClon!numfield))) = "01" Then TipAfp = ""
                Sql = "select descrip + '" & TipAfp & "' from maestros_2 where" & mcadmae & " and cod_maestro2='" & Trim(Rs(Val(rsClon!numfield))) & "'"
            End If
            
            If (fAbrRst(rs2, Sql)) Then rsLista.Fields(mcamp) = Mid(rs2(0), 1, 50)
            rs2.Close
            'MGIRAO SENATI
        ElseIf rsClon!numfield = 81 Then
              Sql = "select c.senati from planilla_ccosto P,Pla_CCostos C where p.cia='" & wcia & "' and p.placod='" & Trim(Rs!PlaCod) & "' and p.status<>'*' and p.cia=c.cia and p.ccosto=c.codigo and c.status<>'*'"
                If (fAbrRst(rs2, Sql)) Then rsLista.Fields(mcamp) = rs2(0)
                rs2.Close
        
'        ElseIf rsClon!numfield = 83 Then 'EPS
'               rsLista.Fields(mcamp) = IIf(Rs.Fields(Val(rsClon!numfield)) = True, "SI", "NO")

         ElseIf rsClon!General = "B" Then
            Sql = "select  p.concepto,p.importe,m.descrip from plaremunbase p,maestros_2 m " _
                & "where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and p.status<>'*' and m.ciamaestro='01076' and m.cod_maestro2=p.tipo"
            If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
            Do While Not rs2.EOF
               mcamp = "b" & Format(rs2(0), "0000")
               rsLista.Fields(mcamp) = rs2(1)
               mcamp = "t" & Format(rs2(0), "0000")
               rsLista.Fields(mcamp) = Left(rs2(2), 1)
               rs2.MoveNext
            Loop
            rs2.Close
            
            ElseIf rsClon!General = "X" Then
            'GERENCIA
            Sql = "SELECT distinct cod_ger, Gerencia FROM pla_areas where status<>'*' and cod_ger= substring(rtrim('" & Rs!cod_area & "'),1,2) "
            If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
            Do While Not rs2.EOF
               mcamp = "x" & Format("01", "00")
               rsLista.Fields(mcamp) = rs2(1)
               rs2.MoveNext
            Loop
            rs2.Close
            ElseIf rsClon!General = "W" Then
            'Departamento
            Sql = "SELECT distinct cod_dpto, dpto FROM pla_areas where status<>'*' and cod_ger= substring(rtrim('" & Rs!cod_area & "'),1,2) and cod_dpto= substring(rtrim('" & Rs!cod_area & "'),3,2) "
            If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
            Do While Not rs2.EOF
               mcamp = "w" & Format("01", "00")
               rsLista.Fields(mcamp) = rs2(1)
               rs2.MoveNext
            Loop
            rs2.Close
            
         ElseIf rsClon!General = "T" Then
            rsLista.Fields(mcamp) = ""
            'Sql = "select top 3 rtrim(left(descripcion,12)) + ' ' + telefono from platelefono where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*'"
            Sql = "select  top 4  telefono from platelefono where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*'"
            
            If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
            Do While Not rs2.EOF
               If Trim(rsLista.Fields(mcamp) & "") = "" Then
                  rsLista.Fields(mcamp) = Trim(rs2(0) & "")
               Else
                  rsLista.Fields(mcamp) = Trim(rsLista.Fields(mcamp)) & "," & Trim(rs2(0)) & ""
               End If
               rs2.MoveNext
            Loop
            rs2.Close
            
        ElseIf rsClon!General = "K" Then 'Uniforme
'           Sql = "select 'Botas :' + Botas +' - ' + 'Camisa :' + Camisa +' - ' + 'Pantalon :' + Pantalon +' - ' +'Polo :' + Polo  from plauniforme where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*'"
           'Sql = "select Talla  from plauniformes where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!placod & "' and status<>'*'"
'            If (fAbrRst(rs2, Sql)) Then rsLista.Fields(mcamp) = Trim(rs2(0))
'            rs2.Close
            Sql = "select top 1 isnull(Talla,'')  from plauniformes where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*' and descripcion='BOTAS' union all "
            Sql = Sql + "select top 1 isnull(Talla,'')  from plauniformes where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*' and descripcion='CAMISA' union all "
            Sql = Sql + "select top 1 isnull(Talla,'')  from plauniformes where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*' and descripcion='PANTALON' union all "
            Sql = Sql + "select top 1 isnull(Talla,'')  from plauniformes where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*' and descripcion='POLO' union all "
            Sql = Sql + "select top 1 isnull(Talla,'')  from plauniformes where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*' and descripcion='CHOMPA' union all "
            Sql = Sql + "select top 1 isnull(Talla,'')  from plauniformes where cia='" & AdoPersonal.Recordset!cia & "' and placod='" & AdoPersonal.Recordset!PlaCod & "' and status<>'*' and descripcion='MANDIL' "
            
            If (fAbrRst(rs2, Sql)) Then rs2.MoveFirst
            Do While Not rs2.EOF
               'If Trim(rsLista.Fields(mcamp) & "") = "" Then
               '   rsLista.Fields(mcamp) = Trim(rs2(0) & "")
               'Else
                  rsLista.Fields(mcamp) = Trim(rsLista.Fields(mcamp) & "") & Trim(rs2(0) & " ") + " ; "
               'End If
               rs2.MoveNext
            Loop
            rs2.Close
            
        ElseIf rsClon!General = "D" Then 'Direccion
            mcadmae = ""
            Sql = "select dbo.fc_Trae_Direccion('" & AdoPersonal.Recordset!cia & "','" & AdoPersonal.Recordset!PlaCod & "')"
            If (fAbrRst(rs2, Sql)) Then mcadmae = Trim(rs2(0) & "")
            rs2.Close
            
            'Sql = "select descrip from maestros_2 where right(ciamaestro,3)='036' and cod_maestro2='" & rs(22) & "' and status<>'*'"
            'If (fAbrRst(rs2, Sql)) Then mcadmae = Trim(rs2(0))
            'rs2.Close
            'mcadmae = mcadmae & " " & Trim(rs(23)) & " " & Trim(rs(24)) & "  " & Trim(rs(25))
            
            rsLista.Fields(mcamp) = Mid(mcadmae, 1, 250)
         ElseIf rsClon!General = "U" Then
            Sql = "select nombre from sunat_ubigeo where id_ubigeo ='" & Rs!ubigeo & "'"
            If (fAbrRst(rs2, Sql)) Then rsLista.Fields(mcamp) = Trim(rs2(0) & "")
            rs2.Close
         ElseIf rsClon!General = "Y" Then
            Sql = "select dbo.fc_nombre_ubigeo_sunat_SEE_v2('" & Rs!ubigeo & "',1)"
            If (fAbrRst(rs2, Sql)) Then rsLista.Fields(mcamp) = Trim(rs2(0) & "")
            rs2.Close
               
        ElseIf rsClon!General = "C" Then 'CENTRO COSTO
            
             Sql = "select m.descripcion from planilla_ccosto c,Pla_ccostos m where c.cia='" & wcia & _
             "' and c.placod='" & AdoPersonal.Recordset!PlaCod & "' and c.status<>'*' and m.cia=c.cia and m.codigo=c.ccosto and m.status<>'*' "
            If (fAbrRst(rs2, Sql)) Then rsLista.Fields(mcamp) = Trim(rs2(0) & "")
            rs2.Close
         End If
         rsClon.MoveNext
         
      Loop
   End If
   Rs.Close
   AdoPersonal.Recordset.MoveNext
Loop
rsClon.Close
Set rsClon = Nothing
If rsLista.RecordCount > 0 Then Carga_Excel (W)
End Sub
Private Sub Carga_Excel(FF As Integer)
Dim mcad As String
Dim nFil As Integer
Dim nCol As Integer
'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet

Dim xlApp2 As Object
Dim xlApp1  As Object
Dim xlBook As Object
Dim xlSheet As Object


Set rsClon = rsReport.Clone
rsClon.MoveFirst
Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

xlSheet.Cells(1, 1).Value = Cmbcia
xlSheet.Cells(1, 1).Font.Bold = True
xlSheet.Cells(1, 1).Font.Size = 12
xlSheet.Cells(1, 1).HorizontalAlignment = xlCenter

xlSheet.Cells(3, 1).Value = "REPORTE DE TRABAJADORES"
xlSheet.Cells(3, 1).Font.Bold = True
xlSheet.Cells(3, 1).Font.Size = 12
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter

nCol = 1
Barra.Max = rsClon.RecordCount
Do While Not rsClon.EOF
   Barra.Value = rsClon.AbsolutePosition
   If rsClon!General = "B" Then
      Sql = "select distinct c.descripcion,b.concepto from plaremunbase b,placonstante c where b.cia='" & wcia & "' and b.status<>'*' " _
          & "and c.cia=b.cia and c.tipomovimiento='02' and c.codinterno=b.concepto and c.status<>'*' " _
          & "order by b.concepto"
      If (fAbrRst(Rs, Sql)) Then Rs.MoveFirst
      nFil = 0
      Do While Not Rs.EOF
         If nFil = 0 Then
            xlSheet.Cells(5, nCol).Value = "REM. " & Trim(Rs(0))
         Else
            xlSheet.Cells(5, nCol).Value = Trim(Rs(0))
         End If
         nFil = nFil + 1
         nCol = nCol + 1
         xlSheet.Cells(5, nCol).Value = "PER"
         nCol = nCol + 1
         Rs.MoveNext
      Loop
      Rs.Close
   Else
      xlSheet.Cells(5, nCol).Value = Trim(rsClon!CORTO)
      nCol = nCol + 1
   End If
   rsClon.MoveNext
Loop
xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, nCol - 1)).Merge
xlSheet.Range(xlSheet.Cells(3, 1), xlSheet.Cells(3, nCol - 1)).Merge

xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, nCol - 1)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, nCol - 1)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(5, 1), xlSheet.Cells(5, nCol - 1)).Borders.LineStyle = xlContinuous
nFil = 6

rsClon.Sort = "orden"
rsClon.MoveFirst
mcad = ""
Do While Not rsClon.EOF
   If IsNumeric(rsClon!Orden) Then
      If rsClon!General = "B" Then
         mcad = mcad & "b0001,"
      Else
         mcad = mcad & "f" & Format(rsClon!Item, "0000") & ","
      End If
   End If
   rsClon.MoveNext
Loop
If Trim(mcad) <> "" Then mcad = Mid(mcad, 1, Len(Trim(mcad)) - 1)

rsLista.Sort = mcad

rsLista.MoveFirst
Barra.Max = rsLista.RecordCount
Do While Not rsLista.EOF
   Barra.Value = rsLista.AbsolutePosition
   nCol = 1
   For I = 0 To FF - 1
       xlSheet.Cells(nFil, nCol).Value = rsLista(I)
       If Left(rsLista(I).Name, 1) = "b" Then xlSheet.Cells(nFil, nCol).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
       nCol = nCol + 1
   Next
   nFil = nFil + 1
   rsLista.MoveNext
Loop

xlSheet.Range("A:AZ").EntireColumn.AutoFit

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
End Sub

