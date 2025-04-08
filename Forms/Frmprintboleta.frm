VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frmprintboleta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Impresión de Boletas «"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "Frmprintboleta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkdsctojudicial 
      Caption         =   "Descuento Judicial"
      Height          =   255
      Left            =   2040
      TabIndex        =   51
      Top             =   1940
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cruce Con Maestro"
      Height          =   375
      Left            =   240
      TabIndex        =   44
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CheckBox ChkPreImpreso 
      Caption         =   "Formato Pre Impreso"
      Height          =   255
      Left            =   2040
      TabIndex        =   43
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8535
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   60
         Width           =   5175
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
         TabIndex        =   20
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Lblfecha 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Left            =   6495
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5775
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Width           =   8055
      Begin MSAdodcLib.Adodc Adocabeza 
         Height          =   375
         Left            =   240
         Top             =   4005
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Adodc1"
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
      Begin VB.Frame FramePrint 
         Height          =   1695
         Left            =   360
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox TxtRefeBol 
            Height          =   375
            Left            =   2520
            MaxLength       =   25
            TabIndex        =   49
            Top             =   840
            Width           =   4575
         End
         Begin VB.OptionButton Opcrango 
            Caption         =   "A partir de Selección"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   25
            Top             =   900
            Width           =   2040
         End
         Begin VB.OptionButton Opctotal 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   24
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Opcindividual 
            Caption         =   "Individual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   23
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Caption         =   "Referencia para la impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2520
            TabIndex        =   50
            Top             =   480
            Width           =   4575
         End
         Begin MSForms.CommandButton Command1 
            Height          =   375
            Left            =   60
            TabIndex        =   29
            Top             =   1260
            Width           =   2055
            Caption         =   "  -   Imprimir Boletas"
            PicturePosition =   327683
            Size            =   "3625;661"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin Threed.SSPanel Panelprogress 
         Height          =   735
         Left            =   840
         TabIndex        =   27
         Top             =   2640
         Visible         =   0   'False
         Width           =   5655
         _Version        =   65536
         _ExtentX        =   9975
         _ExtentY        =   1296
         _StockProps     =   15
         Caption         =   "Generando Boletas"
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
         Alignment       =   6
         Begin MSComctlLib.ProgressBar Barra 
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "Frmprintboleta.frx":030A
         Height          =   5535
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   9763
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
            DataField       =   "totneto"
            Caption         =   "Neto"
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
            DataField       =   "fechaproceso"
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
         BeginProperty Column05 
            DataField       =   "Id_Boleta"
            Caption         =   "Id_Boleta"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4094.929
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Object.Visible         =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   7935
      Begin VB.Frame FrmLiquid 
         Height          =   495
         Left            =   2040
         TabIndex        =   31
         Top             =   520
         Visible         =   0   'False
         Width           =   1575
         Begin VB.CommandButton Command2 
            Caption         =   "Liquidación"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   160
            Width           =   1335
         End
      End
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Txtsemana 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cmbtipotrabajador 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker Cmbfecha 
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   37265
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   3255
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker Cmbal 
         Height          =   285
         Left            =   6360
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60817409
         CurrentDate     =   37265
      End
      Begin MSComCtl2.DTPicker Cmbdel 
         Height          =   285
         Left            =   4320
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   60817409
         CurrentDate     =   37267
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Boleta"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana"
         Height          =   195
         Left            =   2160
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   5880
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.TextBox Txtcodobra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   8
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame FrameBco 
      Height          =   5895
      Left            =   0
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   8175
      Begin MSDataGridLib.DataGrid DgrDepo 
         Height          =   5460
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   9631
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   17
         BeginProperty Column00 
            DataField       =   "excluir"
            Caption         =   "Excluir"
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
            DataField       =   "Codigo"
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
         BeginProperty Column02 
            DataField       =   "NOM_CLIE"
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
         BeginProperty Column03 
            DataField       =   "NUMEROCTA"
            Caption         =   "Cuenta"
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
            DataField       =   "IMPORTE"
            Caption         =   "Importe"
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
            DataField       =   "TIPO_REG"
            Caption         =   "TIPO_REG"
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
            DataField       =   "CUENTA"
            Caption         =   "tCUENTA"
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
         BeginProperty Column07 
            DataField       =   "TIPO_DOC"
            Caption         =   "TIPO_DOC"
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
         BeginProperty Column08 
            DataField       =   "NUMERO_DOC"
            Caption         =   "NUMERO_DOC"
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
         BeginProperty Column09 
            DataField       =   "REF_TRAB"
            Caption         =   "REF_TRAB"
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
         BeginProperty Column10 
            DataField       =   "REF_EMP"
            Caption         =   "REF_EMP"
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
         BeginProperty Column11 
            DataField       =   "MONEDA"
            Caption         =   "MONEDA"
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
         BeginProperty Column12 
            DataField       =   "FLAG"
            Caption         =   "FLAG"
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
         BeginProperty Column13 
            DataField       =   "CTA"
            Caption         =   "CTA"
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
         BeginProperty Column14 
            DataField       =   "CTAINTER"
            Caption         =   "CTAINTER"
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
         BeginProperty Column15 
            DataField       =   "SUCURSAL"
            Caption         =   "SUCURSAL"
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
         BeginProperty Column16 
            DataField       =   "PAGOCUENTA"
            Caption         =   "PAGOCUENTA"
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
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   2910.047
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column05 
               Object.Visible         =   0   'False
               ColumnWidth     =   315.213
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
               ColumnWidth     =   390.047
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
               ColumnWidth     =   374.74
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   315.213
            EndProperty
            BeginProperty Column09 
               Object.Visible         =   0   'False
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column10 
               Object.Visible         =   0   'False
               ColumnWidth     =   90.142
            EndProperty
            BeginProperty Column11 
               Object.Visible         =   0   'False
               ColumnWidth     =   420.095
            EndProperty
            BeginProperty Column12 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column13 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column14 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column15 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column16 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin Threed.SSPanel PnlCruce 
      Height          =   5895
      Left            =   0
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   10398
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
      BevelInner      =   1
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   5460
         Width           =   7815
         _Version        =   65536
         _ExtentX        =   13785
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "CERRAR"
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
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Frmprintboleta.frx":0322
         Height          =   4935
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8705
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483624
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
            DataField       =   "fingreso"
            Caption         =   "F. Ingreso"
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
            Caption         =   "F. Cese"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
               Object.Visible         =   -1  'True
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4199.811
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoCruce 
         Height          =   375
         Left            =   0
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Adodc1"
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Trabajadores Sin Boleta"
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
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   7815
      End
   End
   Begin VB.Frame FrameTxtBco 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   7935
      Begin VB.ComboBox CmbBco 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   120
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker Cbofecha 
         Height          =   285
         Left            =   1560
         TabIndex        =   36
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   37265
      End
      Begin MSForms.CommandButton CommandButton4 
         Height          =   375
         Left            =   7560
         TabIndex        =   52
         Tag             =   "I"
         Top             =   120
         Width           =   360
         Caption         =   "I"
         PicturePosition =   327683
         Size            =   "635;661"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton3 
         Height          =   375
         Left            =   7080
         TabIndex        =   42
         Top             =   120
         Width           =   345
         Caption         =   "B"
         PicturePosition =   327683
         Size            =   "609;661"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton2 
         Height          =   375
         Left            =   6600
         TabIndex        =   39
         Top             =   120
         Width           =   360
         PicturePosition =   327683
         Size            =   "635;661"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   38
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Deposito"
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
         Left            =   120
         TabIndex        =   37
         Top             =   180
         Width           =   1350
      End
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   375
      Left            =   6240
      TabIndex        =   33
      Top             =   1680
      Width           =   1800
      Caption         =   "     TXT Banco"
      PicturePosition =   327683
      Size            =   "3175;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton btn_Exportar 
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      Top             =   1680
      Width           =   1440
      Caption         =   "     Exportar"
      PicturePosition =   327683
      Size            =   "2540;661"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label LblTipDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Lblobra 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "Frmprintboleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VTipotrab As String
Dim VTipoPago As String
Dim VHorasBol As Integer
Dim rsboleta As New Recordset
Dim mlinea As Integer
Public wmBolQuin As String
Dim FLAG As Boolean
Dim rsdepo As New Recordset

Private Sub btn_Exportar_Click()
    If AdoCabeza.Recordset.RecordCount > 0 Then
        Call Exportar_Excel(AdoCabeza.Recordset)
    Else
        MsgBox ("No se tienen registros")
    End If

End Sub

Private Sub chkdsctojudicial_Click()
Call Procesa_Seteo_Boleta
End Sub

Private Sub CmbBco_Click()


If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If

Dim VBcoPago As String
VBcoPago = fc_CodigoComboBox(CmbBco, 2)
Dim rsbco As ADODB.Recordset
Dim xSem As String
If Txtsemana.Visible = True Then xSem = Trim(Txtsemana.Text) Else xSem = "**"


Dim f1 As String
Dim f2 As String

f1 = "01/" & Format(Cmbfecha.Month, "00") & "/" & Format(Cmbfecha.Year, "0000") & " 00:00:00"
f2 = Format(Ultimo_Dia(Cmbfecha.Month, Cmbfecha.Year), "00") & "/" & Format(Cmbfecha.Month, "00") & "/" & Format(Cmbfecha.Year, "0000") & " 11:59:59"

If VTipo = "02" And VTipotrab = "02" Then
   Dim SqlFec As String
   SqlFec = " set dateformat dmy select semana,fechai,fechaf from plasemanas where cia='" & wcia & "' and ano=" & Cmbfecha.Year & " and status<>'*' and '" & Cmbfecha.Value & "' between fechai and fechaf"
   If (fAbrRst(rs, SqlFec)) Then
      f1 = rs!fechai: f2 = rs!fechaf
   End If
   rs.Close
End If

Dim lQ As String
If wmBolQuin = "B" Then lQ = "N" Else lQ = "S"
If chkdsctojudicial.Value = 0 Then
Sql = "Usp_Pla_Archivo_Importa '" & wcia & "','" & VBcoPago & "','" & VTipotrab & "'," & Format(Cmbfecha.Year, "0000") & "," & Format(Cmbfecha.Month, "00") & ",'" & xSem & "','" & VTipo & "','" & f1 & "','" & f2 & "','" & lQ & "'"
Else
Sql = "Usp_Pla_Archivo_Importa_V2 '" & wcia & "','" & VBcoPago & "','" & VTipotrab & "'," & Format(Cmbfecha.Year, "0000") & "," & Format(Cmbfecha.Month, "00") & ",'" & xSem & "','" & VTipo & "','" & f1 & "','" & f2 & "','" & lQ & "'"
End If
Debug.Print Sql
If (fAbrRst(rsbco, Sql)) Then rsbco.MoveFirst
Do While Not rsbco.EOF
    rsdepo.AddNew
    rsdepo!Codigo = Trim(rsbco!Codigo & "")
    rsdepo!NOM_CLIE = Trim(rsbco!NOM_CLIE & "")
    rsdepo!NUMEROCTA = Trim(rsbco!NUMEROCTA & "")
    rsdepo!importe = rsbco!importe
    rsdepo!TIPO_REG = Trim(rsbco!TIPO_REG & "")
    rsdepo!Cuenta = Trim(rsbco!Cuenta & "")
    rsdepo!tipo_doc = Trim(rsbco!tipo_doc & "")
    rsdepo!NUMERO_DOC = Trim(rsbco!NUMERO_DOC & "")
    rsdepo!REF_TRAB = Trim(rsbco!REF_TRAB & "")
    rsdepo!REF_EMP = Trim(rsbco!REF_EMP & "")
    rsdepo!moneda = Trim(rsbco!moneda & "")
    rsdepo!FLAG = Trim(rsbco!FLAG & "")
    rsdepo!cta = Trim(rsbco!cta & "")
    rsdepo!CTAINTER = Trim(rsbco!CTAINTER & "")
    rsdepo!Sucursal = Trim(rsbco!Sucursal & "")
    rsdepo!PAGOCUENTA = Trim(rsbco!PAGOCUENTA & "")
    rsdepo!Excluir = ""
    rsbco.MoveNext
Loop
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & CmbCia.ItemData(CmbCia.ListIndex), 2))
'Call fc_Descrip_Maestros2("01078", "", Cmbtipo)

If wGrupoPla = "01" And UCase(wuser) <> "SA" Then
   wciamae = Determina_Maestro("01078")
   Sql$ = "Select COD_MAESTRO2,DESCRIP from maestros_2 where status<>'*' and (cod_maestro2 in(select tipo from pla_permisos where usuario='" & wuser & "' and calculo='B'))"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   If rs.RecordCount > 0 Then rs.MoveFirst
   Do While Not rs.EOF
      Cmbtipo.AddItem rs!DESCRIP
      Cmbtipo.ItemData(Cmbtipo.NewIndex) = Trim(rs!cod_maestro2)
      rs.MoveNext
   Loop
   rs.Close
Else
   Cadena = "SELECT RTRIM(COD_MAESTRO2) AS COD_MAESTRO2, RTRIM(DESCRIP) AS DESCRIP FROM MAESTROS_2 WHERE RIGHT(CIAMAESTRO, 3) = '078' AND STATUS = '' AND COD_MAESTRO2 NOT IN ('10','07') ORDER BY COD_MAESTRO2"
   Call rCarCbo(Cmbtipo, Cadena, "XX", "00")
End If
If Cmbtipo.ListCount = 1 Then Cmbtipo.ListIndex = 0

If wmBolQuin = "B" = True Then
   Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
   Procesa_Seteo_Boleta
Else
   wciamae = Determina_Maestro("01055")
   Sql$ = "Select * from maestros_2 where flag1='04' and status<>'*'"
   Sql$ = Sql$ & wciamae
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   If rs.RecordCount > 0 Then rs.MoveFirst
   Do While Not rs.EOF
      Cmbtipotrabajador.AddItem rs!DESCRIP
      Cmbtipotrabajador.ItemData(Cmbtipotrabajador.NewIndex) = Trim(rs!cod_maestro2)
      rs.MoveNext
   Loop
   rs.Close
   If Cmbtipotrabajador.ListCount >= 0 Then Cmbtipotrabajador.ListIndex = 0
End If
Crea_Rs
End Sub

Private Sub Cmbfecha_Change()
Procesa_Seteo_Boleta
End Sub

Private Sub CmbTipo_Click()
FrmLiquid.Visible = False
VTipo = Funciones.fc_CodigoComboBox(Cmbtipo, 2)
If VTipo = "02" Then
   Cmbdel.Visible = False
   Cmbal.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Label4.Visible = False
   Txtsemana.Visible = False
   UpDown1.Visible = False
   Label5.Visible = True
   Label6.Visible = True
   Cmbdel.Visible = True
   Cmbal.Visible = True
ElseIf VTipo = "03" Then
   Cmbdel.Visible = False
   Cmbal.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Label4.Visible = False
   Txtsemana.Visible = False
   UpDown1.Visible = False
   Label5.Visible = False
   Label6.Visible = False
   Cmbdel.Visible = False
   Cmbal.Visible = False
   
Else
   Cmbdel.Visible = True
   Cmbal.Visible = True
   Label5.Visible = True
   Label6.Visible = True
End If

If VTipo = "04" Then FrmLiquid.Visible = True

Cmbtipotrabajador_Click
Procesa_Seteo_Boleta
End Sub

Private Sub Cmbtipotrabajador_Click()

VTipotrab = Funciones.fc_CodigoComboBox(Cmbtipotrabajador, 2)
Dim wciamae As String

wciamae = Funciones.Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where cod_maestro2='" & _
VTipotrab & "' and status<>'*'"

Sql$ = Sql$ & wciamae

cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount > 0 Then VHorasBol = Val(rs!flag2): VTipoPago = Left(rs!flag1, 2)
If VTipo = "01" Or VTipo = "05" Or VTipo = "11" Then
   If rs.RecordCount > 0 Then
      Select Case Left(rs!flag1, 2)
             Case Is <> "02"
                  Txtsemana.Text = ""
                  Txtsemana.Visible = False
                  UpDown1.Visible = False
                  Label4.Visible = False
                  Label5.Visible = False
                  Label6.Visible = False
                  Cmbdel.Visible = False
                  Cmbal.Visible = False
             Case Else
                  Txtsemana.Visible = True
                  UpDown1.Visible = True
                  Label4.Visible = True
                  Label5.Visible = True
                  Label6.Visible = True
                  Cmbdel.Visible = True
                  Cmbal.Visible = True
      End Select
    End If
End If
If VTipotrab = "05" And VTipo = "01" Then
   Txtcodobra.Text = ""
   Lblobra.Caption = ""
   Txtcodobra.Visible = False
   Lblobra.Visible = False
Else
   Txtcodobra.Text = ""
   Lblobra.Caption = ""
   Txtcodobra.Visible = False
   Lblobra.Visible = False
End If
If rs.State = 1 Then rs.Close
Procesa_Seteo_Boleta
End Sub
Private Sub Command1_Click()
If wmBolQuin = "B" Then
   Panelprogress.Caption = "Generando Boletas de Pago"
   If ChkPreImpreso.Value = 1 Then
      Print_Bol_COMACSA
   Else
      Print_Bol
   End If
Else
   Panelprogress.Caption = "Generando Recibos de Quincena"
   Print_Quin
End If
End Sub

Private Sub Command2_Click()
On Error GoTo Salir
Liquidacion (AdoCabeza.Recordset!id_boleta)
Exit Sub

Salir:
MsgBox "Seleccione Boleta", vbCritical, Me.Caption
End Sub

Private Sub Command3_Click()
Sql$ = "Usp_Pla_Cruce_Boletas_Maestro '" & wcia & "','" & VTipotrab & "','" & VTipo & "'," & Cmbfecha.Year & "," & Cmbfecha.Month & "," & Ultimo_Dia(Cmbfecha.Month, Cmbfecha.Year) & ",'" & Format(Trim(Txtsemana.Text), "00") & "'"
cn.CursorLocation = adUseClient
Set AdoCruce.Recordset = cn.Execute(Sql$, 64)
Command3.Visible = False
PnlCruce.Visible = True
End Sub

Private Sub CommandButton1_Click()

If rsdepo.RecordCount > 0 Then
   rsdepo.MoveFirst
   Do While Not rsdepo.EOF
      rsdepo.Delete
      rsdepo.MoveNext
   Loop
End If


FrameTxtBco.Visible = True
Command3.Visible = False
ChkPreImpreso.Visible = False
chkdsctojudicial.Visible = False
FrameBco.ZOrder 0
FrameBco.Visible = True
Frame4.Enabled = False
End Sub

Private Sub DgrdRemun_Click()

End Sub

Private Sub CommandButton2_Click()
Dim VBcoPago As String
VBcoPago = fc_CodigoComboBox(CmbBco, 2)
If chkdsctojudicial.Value = 0 Then
    Select Case VBcoPago
       Case "01": Print_TxtBcoCredito
       Case "29": Print_TxtBcoScotia
       Case "02": Print_TxtBcoConti
       Case Else
          MsgBox "Implementación no desarrollada para el banco seleccionado", vbInformation
    End Select
Else
    Dim tipo As String
    tipo = ""
    If VTipo = "01" Then
    tipo = "N"
    End If
    If VTipo = "02" Then
    tipo = "V"
    End If
    If VTipo = "03" Then
    tipo = "G"
    End If
    If VTipo = "04" Then
    tipo = "L"
    End If
    If VTipo = "11" Then
    tipo = "U"
    End If
    
    Select Case VBcoPago
       Case "01": Print_TxtBcoCredito_DJ (tipo)
       Case "29": Print_TxtBcoScotia_DJ (tipo)
       Case "02": Print_TxtBcoConti_DJ (tipo)
       Case Else
          MsgBox "Implementación no desarrollada para el banco seleccionado", vbInformation
    End Select
End If
End Sub

Private Sub CommandButton3_Click()
FrameTxtBco.Visible = False
ChkPreImpreso.Visible = True
chkdsctojudicial.Visible = True
Command3.Visible = True
FrameBco.Visible = False
Frame4.Enabled = True
End Sub



Private Sub DgrDepo_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case Is = 0
            If Trim(UCase(DgrDepo.Columns(0))) <> "S" Then
               MsgBox "Excluir solo Puede ser [S]i", vbCritical, "Archivo para el banco"
               DgrDepo.Columns(0) = ""
            Else
               DgrDepo.Columns(0) = Trim(UCase(DgrDepo.Columns(0)))
            End If
End Select
End Sub

Private Sub Form_Activate()
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Me.Command1.Caption = "  -   Imprimir Boletas"
If wmBolQuin = "B" Then
   Me.Caption = "IMPRESION DE BOLETAS"
   'Frame3.BackColor = &H80000001
   Label2.Visible = True
   Cmbtipo.Visible = True
Else
   Me.Caption = "IMPRESION DE QUINCENAS"
   'Frame3.BackColor = &H80000008
   Label2.Visible = False
   Cmbtipo.Visible = False
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{tab}"
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 And FramePrint.Visible = True Then
           FramePrint.Visible = False
        End If
End Sub

Private Sub Form_Load()
FLAG = False
Me.Top = 0
Me.Left = 0
Me.Width = 8265
Me.Height = 8565
Call Funciones.rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call Funciones.rUbiIndCmbBox(CmbCia, wcia, "00")

Cmbfecha.Year = Year(Date)
Cmbfecha.Month = Month(Date)
Cmbfecha.Day = Day(Date)

Cmbdel.Year = Year(Date)
Cmbdel.Month = Month(Date)
Cmbdel.Day = Day(Date)

Cmbal.Year = Year(Date)
Cmbal.Month = Month(Date)
Cmbal.Day = Day(Date)
Call fc_Descrip_Maestros2("01007", "", CmbBco, False)
Me.Command1.Caption = "  -   Imprimir Boletas"
Me.KeyPreview = True
Cbofecha.Value = Date
End Sub

Private Sub SSCommand1_Click()
Command3.Visible = True
PnlCruce.Visible = False
End Sub

Private Sub Txtsemana_Change()
Procesa_Seteo_Boleta
End Sub
'
Public Sub Procesa_Seteo_Boleta()
Dim mano As Integer
Dim mmes As Integer
On Error GoTo CORRIGE

Command3.Visible = True
PnlCruce.Visible = False

If Trim(Txtsemana.Text) <> "" Then
   Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(Cmbfecha.Year, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"

   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   If rs.RecordCount > 0 Then
      Cmbdel.Value = Format(rs!fechai, "dd/mm/yyyy")
      Cmbal.Value = Format(rs!fechaf, "dd/mm/yyyy")
      Cmbfecha.Value = Format(rs!fechaf, "dd/mm/yyyy")
   End If
   If rs.State = 1 Then rs.Close
End If
If wmBolQuin = "B" Then
   If VTipoPago = "" Or IsNull(VTipoPago) Then Exit Sub
End If

mano = Val(Mid(Cmbfecha.Value, 7, 4))
mmes = Val(Mid(Cmbfecha.Value, 4, 2))
If wmBolQuin = "B" Then
   Select Case VTipoPago
          Case Is = "02"
               Sql$ = nombre()
               If VTipo = "04" Then
               If chkdsctojudicial.Value = 0 Then
                  Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=b.pagobanco),'') as Banco " _
                     & " from plahistorico a,planillas b " _
                     & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "'" _
                     & " and a.semana='" & Txtsemana.Text & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
                     & " and month(fechaproceso)=" & Format(Cmbfecha.Month, "00") & " " _
                     & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
                Else
                     Sql$ = "select c.nombre,a.placod,b.moneda,D05*(C.PORCENTAJE/(SELECT SUM(R.PORCENTAJE) FROM TBL_BCO_CUENTA_DJ R WHERE a.placod=R.placod AND r.status<>'*' )) as totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=c.bco),'') as Banco ,c.nrocta,c.nrodni" _
                     & " from plahistorico a,planillas b ,TBL_BCO_CUENTA_DJ c" _
                     & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "'" _
                     & " and a.semana='" & Trim(Txtsemana.Text) & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
                     & " and month(fechaproceso)=" & Format(Cmbfecha.Month, "00") & " "
                     Sql$ = Sql$ & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' and a.placod=c.placod  and D05<>0 ORDER BY A.PLACOD"
                End If
               ElseIf VTipo = "02" Then
                  Dim f1 As String
                  Dim f2 As String
                  Dim SqlFec As String
                  SqlFec = " set dateformat dmy select semana,fechai,fechaf from plasemanas where cia='" & wcia & "' and ano=" & Cmbfecha.Year & " and status<>'*' and '" & Cmbfecha.Value & "' between fechai and fechaf"
                  If (fAbrRst(rs, SqlFec)) Then
                     f1 = rs!fechai: f2 = rs!fechaf
                  End If
                  rs.Close
                  If chkdsctojudicial.Value = 0 Then
                  Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=b.pagobanco),'') as Banco " _
                     & " from plahistorico a,planillas b " _
                     & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "'" _
                     & " and fechaproceso between '" & f1 & "' and '" & f2 & "' " _
                     & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
                   Else
                   Sql$ = "select c.nombre,a.placod,b.moneda,D05*(C.PORCENTAJE/(SELECT SUM(R.PORCENTAJE) FROM TBL_BCO_CUENTA_DJ R WHERE a.placod=R.placod AND r.status<>'*' )) as totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=c.bco),'') as Banco,c.nrocta,c.nrodni " _
                     & " from plahistorico a,planillas b ,TBL_BCO_CUENTA_DJ c" _
                     & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "'" _
                     & " and a.semana='" & Trim(Txtsemana.Text) & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
                     & " and fechaproceso between '" & f1 & "' and '" & f2 & "' " _
                     & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' and a.placod=c.placod  and D05<>0 AND C.STATUS<>'*' ORDER BY A.PLACOD"
                   End If
               Else
                  Dim strMes As String
                  strMes = ""
                  If chkdsctojudicial.Value = 0 Then
                     Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=b.pagobanco),'') as Banco " _
                     & " from plahistorico a,planillas b " _
                     & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "'" _
                     & " and a.semana='" & Trim(Txtsemana.Text) & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
                     & " and month(fechaproceso)=" & Format(Cmbfecha.Month, "00") & " "
                     Sql$ = Sql$ & strMes
                     Sql$ = Sql$ & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
                  Else
                     Sql$ = "select c.nombre,a.placod,b.moneda,D05*(C.PORCENTAJE/(SELECT SUM(R.PORCENTAJE) FROM TBL_BCO_CUENTA_DJ R WHERE a.placod=R.placod AND r.status<>'*' )) as totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=c.bco),'') as Banco ,c.nrocta,c.nrodni" _
                     & " from plahistorico a,planillas b ,TBL_BCO_CUENTA_DJ c" _
                     & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "'" _
                     & " and a.semana='" & Trim(Txtsemana.Text) & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
                     & " and month(fechaproceso)=" & Format(Cmbfecha.Month, "00") & " "
                     Sql$ = Sql$ & strMes
                     Sql$ = Sql$ & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' and a.placod=c.placod  and D05<>0 AND C.STATUS<>'*' ORDER BY A.PLACOD"
                  End If
                  
               End If
          Case Is = "04"
                  Sql$ = nombre()
                  If chkdsctojudicial.Value = 0 Then
                    Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=b.pagobanco),'') as Banco " _
                    & "from plahistorico a,planillas b " _
                    & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' " _
                    & "and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
                    & "and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
                  Else
                    Sql$ = "select c.nombre,a.placod,b.moneda,D05*(C.PORCENTAJE/(SELECT SUM(R.PORCENTAJE) FROM TBL_BCO_CUENTA_DJ R WHERE a.placod=R.placod AND r.status<>'*' )) as totneto,a.fechaproceso,Id_Boleta,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=c.bco),'') as Banco " _
                     & " from plahistorico a,planillas b ,TBL_BCO_CUENTA_DJ c" _
                     & " where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "'" _
                     & " and a.semana='" & Trim(Txtsemana.Text) & "' AND YEAR(fechaproceso)=" & Format(Cmbfecha.Year, "0000") & " " _
                     & " and month(fechaproceso)=" & Format(Cmbfecha.Month, "00") & " "
                     Sql$ = Sql$ & strMes
                     Sql$ = Sql$ & " and a.status='T' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' and a.placod=c.placod  and D05<>0 AND C.STATUS<>'*' ORDER BY A.PLACOD"
                  End If
   End Select
Else
   Sql$ = nombre()
   Sql$ = Sql$ & "a.placod,b.moneda,a.totneto,a.fechaproceso,(SELECT descripcion FROM pla_ccostos where codigo=b.area) as Area,isnull((select descrip from maestros_2 where ciamaestro='01007' and cod_maestro2=b.pagobanco),'') as Banco " & _
   "from plaquincena a,planillas b " _
   & "where a.cia='" & wcia & "' and b.tipotrabajador='" & VTipotrab & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " " _
   & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' ORDER BY A.PLACOD"
End If
cn.CursorLocation = adUseClient
Set AdoCabeza.Recordset = cn.Execute(Sql$, 64)
FLAG = True
If AdoCabeza.Recordset.RecordCount > 0 Then AdoCabeza.Recordset.MoveFirst: If VTipo <> "02" Then Cmbfecha.Value = Format(AdoCabeza.Recordset!FechaProceso, "dd/mm/yyyy")
Dgrdcabeza.Refresh
Screen.MousePointer = vbDefault
Exit Sub
CORRIGE:
MsgBox "Error : " & Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Txtsemana_KeyPress(KeyAscii As Integer)
    Txtsemana.Text = Txtsemana.Text + fc_ValNumeros(KeyAscii)
End Sub

Private Sub UpDown1_DownClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "0"
If Txtsemana.Text > 0 Then Txtsemana = Format(Txtsemana - 1, "00")
End Sub

Private Sub UpDown1_UpClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "0"
Txtsemana = Format(Txtsemana + 1, "00")
End Sub
Public Sub Imprime_Boletas()

If Not FLAG Then Exit Sub
If AdoCabeza.Recordset.RecordCount <= 0 Then Exit Sub
If rsboleta.RecordCount > 0 Then rsboleta.MoveFirst
If wmBolQuin = "B" Then Command1.Caption = "  -   Imprimir Boleta" Else Command1.Caption = "  -   Imprimir Quincena"
TxtRefeBol.Text = ""
FramePrint.Visible = True
End Sub
Private Sub Crea_Rs()
    If rsboleta.State = 1 Then rsboleta.Close
    rsboleta.Fields.Append "texto", adChar, 65, adFldIsNullable
    rsboleta.Open
    
    
    If rsdepo.State = 1 Then rsdepo.Close
    rsdepo.Fields.Append "TIPO_REG", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "CUENTA", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "NUMEROCTA", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "TIPO_DOC", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "NUMERO_DOC", adChar, 12, adFldIsNullable
    rsdepo.Fields.Append "NOM_CLIE", adChar, 75, adFldIsNullable
    rsdepo.Fields.Append "REF_TRAB", adChar, 40, adFldIsNullable
    rsdepo.Fields.Append "REF_EMP", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "MONEDA", adChar, 4, adFldIsNullable
    rsdepo.Fields.Append "IMPORTE", adDouble, 2, adFldIsNullable
    rsdepo.Fields.Append "FLAG", adChar, 1, adFldIsNullable
    rsdepo.Fields.Append "cta", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "CODIGO", adChar, 8, adFldIsNullable
    rsdepo.Fields.Append "CTAINTER", adChar, 20, adFldIsNullable
    rsdepo.Fields.Append "SUCURSAL", adChar, 3, adFldIsNullable
    rsdepo.Fields.Append "PAGOCUENTA", adChar, 20, adFldIsNullable
    
  
    rsdepo.Fields.Append "EXCLUIR", adChar, 1, adFldIsNullable

    rsdepo.Open
    Set DgrDepo.DataSource = rsdepo

End Sub
Private Function No_Seteados(tipo As String, seteo As String, status As String) As Currency
Dim mHasta As Integer
Dim mFields As Integer
Dim mCadSet As String
Dim I As Integer
Dim J As Integer
Dim mFound As Boolean
Dim mLen As Integer
Dim mcadadd As String
Dim rsnoset As ADODB.Recordset
Dim rsnorem As ADODB.Recordset
Dim mbasico As Currency

No_Seteados = 0
mLen = 0
If Trim(seteo) <> "" Then mLen = Len(Trim(seteo))

Select Case tipo
       Case Is = "IN": mHasta = 50: mFields = 44
       Case Is = "AP": mHasta = 21: mFields = 115
       Case Is = "DE": mHasta = 21: mFields = 94
End Select
mCadSet = ""
For I = 1 To mHasta
'    MsgBox rs(i + mFields).Name
    If rs(I + mFields) <> 0 Then
       mCount = 0
       mFound = False
       If mLen > 0 Then
          For J = 1 To mLen - 1 Step 2
              If Mid(seteo, J, 2) = Format(I, "00") Then mFound = True: Exit For
          Next
          If mFound = False Then mCadSet = mCadSet & "'" & Format(I, "00") & "',"
       Else
          mCadSet = mCadSet & "'" & Format(I, "00") & "',"
       End If
    End If
Next
If Trim(mCadSet) = "" Then Exit Function

mCadSet = Mid(mCadSet, 1, Len(Trim(mCadSet)) - 1)
mCadSet = "in(" & mCadSet & ")"
    
If status = "R" Then
   Sql = "SELECT distinct(codinterno),descripcion FROM PLACONSTANTE  c, plaafectos a " _
       & "WHERE c.cia='" & wcia & "' and c.CODINTERNO " & mCadSet & " AND c.TIPOMOVIMIENTO='02'  and c.status<>'*' " _
       & "and a.cia=c.cia and c.codinterno=a.cod_remu and a.status<>'*'"
   If (fAbrRst(rsnoset, Sql)) Then rsnoset.MoveFirst
   Do While Not rsnoset.EOF
      mbasico = 0
      mnumh = Remun_Horas(rsnoset!codinterno)
      If mnumh <> 0 Then mbasico = rs(14 + mnumh)
      If mbasico <> 0 Then
         mcadadd = lentexto(10, Left(rsnoset!Descripcion, 10)) & fCadNum(mbasico, "##0.00")
      Else
         mcadadd = lentexto(16, Left(rsnoset!Descripcion, 16))
      End If
      mbasico = rs(mFields + Val(rsnoset!codinterno))
      mcadadd = mcadadd & Space(1) & fCadNum(mbasico, "###,##0.00")
      rsboleta.AddNew
      rsboleta!texto = mcadadd
      No_Seteados = No_Seteados + mbasico
      rsnoset.MoveNext
   Loop
   rsnoset.Close
ElseIf status = "N" Then
   Sql = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='02' and codinterno " & mCadSet & " and status<>'*'"
   If (fAbrRst(rsnoset, Sql)) Then rsnoset.MoveFirst
   Do While Not rsnoset.EOF
      mbasico = 0
      Sql = "select * from plaafectos where cia='" & wcia & "' and cod_remu ='" & rsnoset!codinterno & "' and status<>'*'"
      If Not (fAbrRst(rsnorem, Sql)) Then
         mbasico = rs(mFields + Val(rsnoset!codinterno))
         mcadadd = lentexto(16, Left(rsnoset!Descripcion, 16))
         mcadadd = mcadadd & Space(1) & fCadNum(mbasico, "###,##0.00")
         rsboleta.AddNew
         rsboleta!texto = mcadadd
         No_Seteados = No_Seteados + mbasico
      End If
      rsnorem.Close
      rsnoset.MoveNext
   Loop
   rsnoset.Close
Else
   Sql = "select codinterno,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno " & mCadSet & " and status<>'*'"
   If (fAbrRst(rsnoset, Sql)) Then rsnoset.MoveFirst
   Do While Not rsnoset.EOF
      mbasico = rs(mFields + Val(rsnoset!codinterno))
      If tipo = "AP" Then
         mcad = lentexto(12, Left(rsnoset!Descripcion, 12)) & Space(1) & fCadNum(mbasico, "##,##0.00")
      Else
         mcad = lentexto(12, Left(rsnoset!Descripcion, 12)) & Space(13) & fCadNum(mbasico, "##,##0.00")
      End If
      rsboleta.AbsolutePosition = mlinea
      rsboleta!texto = Left(rsboleta!texto, 27) & Space(4) & mcad
      mlinea = mlinea + 1
      rsnoset.MoveNext
   Loop
   rsnoset.Close
End If
End Function
Private Sub Print_Bol()
Dim mano As Integer
Dim mmes As Integer
Dim mperiodo As String
Dim mnombre As String
Dim mbasico As Currency
Dim mcargo As String
Dim mcad As String
Dim mafp As String
Dim mconcep As String
Dim rs2 As ADODB.Recordset
Dim rsremu As ADODB.Recordset
Dim wciamae As String
Dim mesnombre As String
Dim mcianom As String
Dim mciaregpat As String
Dim mciaruc As String
Dim mciadir As String
Dim mciadist As String
Dim mnumh As Integer
Dim totremu As Currency
Dim mtexto As String
Dim mCadBlanc As String
Dim mCadSeteo As String
Dim RX As New ADODB.Recordset
Dim FACTOR_HORAS As Variant
Dim NIC As Variant
Dim bCia As Boolean

On Error GoTo FUNKA
mtexto = ""
If VTipoPago = "" Or IsNull(VTipoPago) Then Exit Sub
'OBTENEMOS DATOS DE LA CIA
'Sql$ = "select a.*,dist from cia a,ubigeos b where cod_cia='" & wcia & "' and a.cod_ubi=b.cod_ubi"

Sql$ = " select c.*, " & _
      " (select nombre from sunat_departamento where id_dpto=left(c.cod_ubi,2)) as dp, " & _
      " (select nombre from sunat_provincia where id_prov=left(c.cod_ubi,4)) as prov, " & _
      "  u.nombre  as dist, 'PERU' as pais " & _
      " from cia c, sunat_ubigeo u,pla_cia_ubigeo pu " & _
      " WHERE cod_cia='" & wcia & "' AND u.id_ubigeo=pu.cod_ubi and c.cod_cia=pu.cia"
 
If (fAbrRst(rs, Sql$)) Then
   'mcianom = Trim(rs!razsoc)
   mcianom = Trae_CIA(wcia) 'Trim(rs!razsoc)
   mciaregpat = IIf(IsNull(rs!reg_pat), "", Trim(rs!reg_pat))
   mciaruc = rs!RUC
   mciadir = Trim(rs!direcc) & " " & rs!NRO & " " & IIf(IsNull(rs!DIST), "", Trim(rs!DIST))
   mciadist = IIf(IsNull(rs!DIST), "", Trim(rs!DIST))
   'bCia = CBool(IIf(IsNull(rs!Aportacion_EPS), 0, rs!Aportacion_EPS))
End If
If rs.State = 1 Then rs.Close

mano = Val(Mid(Cmbfecha.Value, 7, 4))
mmes = Val(Mid(Cmbfecha.Value, 4, 2))
mesnombre = Name_Month(Format(mmes, "00"))

mlinea = 1
Dim mArchBol As String
mArchBol = "BO" & wcia & VTipotrab & VTipo & ".txt"
RUTA$ = App.Path & "\REPORTS\" & mArchBol
Open RUTA$ For Output As #1

Barra.Max = AdoCabeza.Recordset.RecordCount

If Opctotal.Value = True Then AdoCabeza.Recordset.MoveFirst: Barra.Value = 0

If Opcindividual.Value <> True Then
   Panelprogress.Visible = True
   Panelprogress.ZOrder 0
   Me.Refresh
   If Opcrango.Value = True Then Barra.Value = AdoCabeza.Recordset.AbsolutePosition
End If

Do While Not AdoCabeza.Recordset.EOF
   Barra.Value = AdoCabeza.Recordset.AbsolutePosition
   If rsboleta.RecordCount > 0 Then rsboleta.MoveFirst
   Do While Not rsboleta.EOF
      rsboleta.Delete
      rsboleta.MoveNext
   Loop
   
   mcianom = Trae_CIA(wcia)
  
   Print #1, Chr(15) & Chr(27) & Chr(69) 'Negrita
   mcad = Trim(mcianom)
   mcad = lentexto(65, Left(mcad, 65))
   Print #1, mcad & Space(11) & mcad
   mcad = Space(33) & "BOLETA DE PAGO MES DE " & mesnombre
   Print #1, mcad & Space(14) & mcad
   mciadir = lentexto(49, Left(mciadir, 49))
   mcad = mciadir & " RUC " & mciaruc
   mcad = lentexto(65, Left(mcad, 65))
   Print #1, mcad & Space(11) & mcad & Chr(27) & Chr(70)
   
   If VTipoPago = "04" Then
      mcad = "               PLANILLA EMPLEADOS - SUELDOS"
   Else
      mcad = "               PLANILLA OBREROS - SALARIOS"
   End If
   mcad = lentexto(65, Left(mcad, 65))
   Print #1, mcad & Space(11) & mcad
   Sql$ = nombre()
   
     
'   If VTipoPago = "04" Then
'      Sql$ = Sql$ & "a.*,b.fingreso,b.cargo,b.ipss,b.codafp,b.numafp,b.dni,b.nro_doc,b.afiliado_eps_serv " _
'           & "from plahistorico a,planillas b " _
'           & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & _
'           VTipotrab & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " AND day(A.FECHAPROCESO)=" & Day(Adocabeza.Recordset!FechaProceso) & " and a.placod='" & Adocabeza.Recordset!PlaCod & "' and a.totneto='" & Adocabeza.Recordset!totneto & "' " _
'           & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.placod"
'           mperiodo = Format(mmes, "00")
'   Else
'        If VTipo <> "03" Then
'      Sql$ = Sql$ & "a.*,b.fingreso,b.cargo,b.ipss,b.codafp,b.numafp,b.dni,b.nro_doc,b.afiliado_eps_serv " _
'           & "from plahistorico a,planillas b " _
'           & "where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' and a.semana='" & Trim(Txtsemana.Text) & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " and a.placod='" & Adocabeza.Recordset!PlaCod & "' " _
'           & "and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.placod"
'           mperiodo = Txtsemana.Text
'        Else
'            Sql$ = Sql$ & "a.*,b.fingreso,b.cargo,b.ipss,b.codafp,b.numafp,b.dni,b.nro_doc,b.afiliado_eps_serv " _
'           & "from plahistorico a,planillas b where a.cia='" & wcia & "' and a.proceso='" & VTipo & "' and b.tipotrabajador='" & VTipotrab & "' and a.semana='" & Txtsemana.Text & "' and year(a.fechaproceso)=" & mano & " and month(a.fechaproceso)=" & mmes & " and a.placod='" & Trim(Adocabeza.Recordset!PlaCod) & "' " _
'           & " and a.status<>'*' and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.placod"
'           mperiodo = Txtsemana.Text
'        End If
'   End If

      If VTipoPago = "04" Then mperiodo = Format(mmes, "00") Else mperiodo = Txtsemana.Text
      
      Sql$ = Sql$ & "a.*,b.fingreso,b.cargo,b.ipss,b.codafp,b.numafp,b.dni,b.nro_doc,b.afiliado_eps_serv,b.placodpresentacion " _
           & "from plahistorico a,planillas b " _
           & "where a.Id_Boleta=" & AdoCabeza.Recordset!id_boleta & " and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.placod"
           
   If (fAbrRst(rs, Sql$)) Then
      mbasico = 0
      rs.MoveFirst
      mnombre = lentexto(40, Left(rs!nombre, 40))
      bCia = CBool(IIf(IsNull(rs!afiliado_eps_serv), 0, rs!afiliado_eps_serv))
      
      'OBTENEMOS VALOR BASE Y FACTOR DE LA REM DEL TRABAJADOR
      Sql$ = "select importe,FACTOR_HORAS from plaremunbase where cia='" & wcia & "' and concepto='01' and placod='" & rs!PlaCod & "' and status<>'*'"
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then rs2.MoveFirst: mbasico = rs!Basico: FACTOR_HORAS = rs2("FACTOR_HORAS") Else mbasico = 0
      If rs2.State = 1 Then rs2.Close
      '24/01/2008
      mcad = "CODIGO    : " & rs!PlaCodpresentacion
      Print #1, mcad & Space(56) & mcad
      mcad = ""
      mcad = "NOMBRE    : " & mnombre & " DNI:" & Left(rs!nro_doc, 8) & ")"
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(11) & mcad
      wciamae = Determina_Maestro_2("01055")
      
      'OBTENEMOS CARGO DEL TRABAJADOR
      Sql$ = "select cod_maestro3,descrip from maestros_31 where ciamaestro='" & wcia & "055" & "' AND cod_maestro3='" & rs!Cargo & "'"
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then rs2.MoveFirst: mcargo = rs2!DESCRIP Else mcargo = ""
      If rs2.State = 1 Then rs2.Close
      
      'CARGO DEL TRABAJADOR
      mcargo = lentexto(31, Left(mcargo, 31))
      If Not IsNull(rs!fechacese) Then mcad = Format(rs!fechacese, "dd/mm/yyyy") Else mcad = Space(10)
      mcad = "OCUPACION : " & mcargo & "BASICO : " & fCadNum(mbasico, "##,###,##0.00")
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(11) & mcad
      
      mcad = "FECHA ING : " & Format(rs!fIngreso, "dd/mm/yyyy") & "                CARNET IPSS " & rs!ipss
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(11) & mcad
      
      'OBTENEMOS NOMBRE DEL AFP
      wciamae = Determina_Maestro("01069")
      Sql$ = "select * from maestros_2 where cod_maestro2='" & rs!CodAfp & "' and status<>'*'"
      Sql$ = Sql$ & wciamae
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then rs2.MoveFirst: mafp = rs2!DESCRIP Else mafp = ""
      If rs2.State = 1 Then rs2.Close
      
      mafp = lentexto(26, Left(mafp, 26))
      mcad = "AFP       : " & mafp & "CARNET SPP  " & rs!NUMAFP
      
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(11) & mcad
      If VTipoPago = "04" Then
        If VTipo = "02" Then
            mcad = "VACACIONES DEL " & rs!fechavacai & " AL " & rs!fechavacaF & Space(5) & "(" & Trim(rs!PlaCod) & ")"
        ElseIf VTipo = "03" Then
                mcad = "GRATIFICACION    " & Name_Month(Format(mmes, "00")) & "   " & Year(Cmbfecha.Value) & Space(10) & "(" & Trim(rs!PlaCod) & ")"
            ElseIf VTipo = "02" Then
                mcad = "VACACIONES DEL " & rs!fechavacai & " AL " & rs!fechavacaF & Space(5) & "(" & Trim(rs!PlaCod) & ")"
            ElseIf VTipo = "11" Then
                mcad = "UTILIDADES"
            Else
                mcad = "Mes       : " & mesnombre & "   " & Format(Right(Cmbfecha.Value, 4), "0000") '& Space(5) & "(" & Trim(rs!PLACOD) & ")"
         End If
      ElseIf VTipoPago = "01" Or VTipoPago = "02" Then
            If VTipo = "01" Then
                 mcad = "SEMANA No : " & Txtsemana.Text & "  DEL " & Cmbdel.Value & " AL " & Cmbal.Value '& Space(5) & "(" & Trim(rs!PLACOD) & ")"
             ElseIf VTipo = "02" Then
                mcad = "VACACIONES DEL " & rs!fechavacai & " AL " & rs!fechavacaF & Space(5) & "(" & Trim(rs!PlaCod) & ")"
                ElseIf VTipo = "04" Then
                    mcad = "LIQUIDACION"
                ElseIf VTipo = "11" Then
                    mcad = "UTILIDADES"
                ElseIf VTipo = "03" Then
                    mcad = "GRATIFICACION    " & Name_Month(Format(mmes, "00")) & "   " & Year(Cmbfecha.Value) & Space(10) & "(" & Trim(rs!PlaCod) & ")"
             End If
         ElseIf VTipoPago = "03" Then
            mcad = "GRATIFICACION    " & Name_Month(Format(mmes, "00")) & "   " & Year(Cmbfecha.Value) & Space(10) & "(" & Trim(rs!PlaCod) & ")"
      End If
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(11) & mcad
      '24/01/2008 agregar dias trabajados
      
      mcad = ""
      mcad = "DIAS TRABAJADOS : " & Format(rs!h14, "00")
      If wGrupoPla = "01" Then
         mcad = mcad & "        Dias. Sub :" & Format(rs!h08 + rs!H25, "00")
         mcad = mcad & "        Dias. No Lab :" & Format(rs!h22 + rs!H07, "00")
         Print #1, mcad & Space(76 - Len(mcad)) & mcad
      Else
         'EPS
         If bCia = True Then
           mcad = mcad & Space(20) & "TRABAJADOR AFILIADO A EPS"
         End If
         Print #1, mcad & Space(76 - Len(mcad)) & mcad
      End If
      mcad = ""
      
      '24/01/2008 SE AGREGA PERIODO
    '  mcad = Space(25) & mesnombre & " - " & mano
     ' Print #1, mcad & Space(40) & mcad
     
      'TERMINO DE MODIFICAR BOLETA
      'Print #1, mcad & Space(11) & mcad
      Print #1, String(65, "-") & Space(11) & String(65, "-")
      mcad = "REMUNERACIONES:                DESCUENTOS:   Empleador Trabajador"
      mcad = lentexto(65, Left(mcad, 65))
      Print #1, mcad & Space(11) & mcad
      
      'Print #1,
      'INGRESOS
      
      mCadSeteo = ""
      Sql$ = "select a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
           & "where a.cia='" & wcia & "' and a.status<>'*' and a.tipo='I' " _
           & "and a.cia=b.cia and b.status<>'*' and b.tipomovimiento='02' and a.tipo_trab='" & VTipotrab & "' and a.codigo=b.codinterno " _
           & "order by a.orden,a.codigo"
           
        mlinea = 1
      'Remunerativos
      If (fAbrRst(rs2, Sql$)) Then
         rs2.MoveFirst
         totremu = 0
         Do While Not rs2.EOF
            mbasico = 0
'            If rs2!codigo = "20" Then
            Sql$ = "select * from plaafectos where cia='" & wcia & "' and status<>'*' and cod_remu='" & rs2!Codigo & "' and tipo in ('A','D') AND CODIGO!='13' "
            If (fAbrRst(rsremu, Sql$)) Then
               mnumh = Remun_Horas(rs2!Codigo)
               If mnumh <> 0 Then mbasico = rs(14 + mnumh)
               If mbasico <> 0 Then
                  If rs2!Codigo = "12" And mmes = 5 Then
                     mcad = lentexto(10, Left("1ero Mayo", 10)) & fCadNum(mbasico, "##0.00")
                  Else
                     mcad = lentexto(10, Left(rs2!Descripcion, 10)) & fCadNum(mbasico, "##0.00")
                  End If
               Else
                  If rs2!Codigo = "12" And mmes = 5 Then
                     mcad = lentexto(16, Left("1ero Mayo", 16))
                  Else
                     mcad = lentexto(16, Left(rs2!Descripcion, 16))
                  End If
               End If
               mbasico = rs(44 + Val(rs2!Codigo))
               If mbasico <> 0 Then
                  mcad = mcad & Space(1) & fCadNum(mbasico, "###,##0.00")
               Else
                  mcad = mcad & Space(1) & Space(10)
               End If
               rsboleta.AddNew
               rsboleta!texto = mcad
               totremu = totremu + mbasico
            Else
'               Debug.Print "ME VOY"
            End If
            If rsremu.State = 1 Then rsremu.Close
            mCadSeteo = mCadSeteo & rs2!Codigo
            
            If rs2("CODIGO") = "27" Then
               'CALCULA PROM/DIARIO
               con = "SELECT CAST( (" & totremu & "/(h01+h02+h03+h15+h04+h05+h12)) AS DECIMAL(5,2))*" & FACTOR_HORAS & " AS PROM" & " FROM " & _
               "PLAHISTORICO WHERE PLACOD='" & Trim(AdoCabeza.Recordset!PlaCod) & "' AND " & _
               "YEAR(FECHAPROCESO)=" & mano & " AND month(fechaproceso)=" & mmes & " AND " & _
               "SEMANA='" & Trim(Txtsemana) & "'"
               RX.Open con, cn, adOpenStatic, adLockReadOnly
                  rsboleta.AddNew
                  rsboleta("texto") = "PROM/DIARIO" & Space(1) & Space(10) & RX(0)
                   'totremu = totremu + RX(0)
               RX.Close
            End If
            
            rs2.MoveNext
         Loop
      End If
      
      totremu = totremu + No_Seteados("IN", mCadSeteo, "R")
      rsboleta.AddNew
      rsboleta!texto = Space(17) & "----------"
      rsboleta.AddNew
      rsboleta!texto = "Sub Total I        " & Format(totremu, "###,###,##0.00")
'      rsboleta.AddNew
'      rsboleta!texto = ""
      rsboleta.AddNew
      rsboleta!texto = "NO REMUNERATIVOS :"
'      rsboleta.AddNew
'      rsboleta!texto = ""
      
      'No Remunerativos
      mCadSeteo = ""
      If rs2.RecordCount > 0 Then rs2.MoveFirst
      totremu = 0
      Do While Not rs2.EOF
         mbasico = 0
         Sql$ = "select * from plaafectos where cia='" & wcia & "' and status<>'*' and cod_remu='" & rs2!Codigo & "' and tipo in ('A','D') AND CODIGO!='13' "
         If Not (fAbrRst(rsremu, Sql$)) Then
            mnumh = Remun_Horas(rs2!Codigo)
            If mnumh <> 0 Then mbasico = rs(14 + mnumh)
            If mbasico <> 0 Then
               mcad = lentexto(11, Left(rs2!Descripcion, 11)) & fCadNum(mbasico, "##0.00")
            Else
               mcad = lentexto(16, Left(rs2!Descripcion, 16))
            End If
            mbasico = rs(44 + Val(rs2!Codigo))
            If mbasico <> 0 Then
               mcad = mcad & Space(1) & fCadNum(mbasico, "###,##0.00")
            Else
               mcad = mcad & Space(1) & Space(10)
            End If
            rsboleta.AddNew
            rsboleta!texto = mcad
            totremu = totremu + mbasico
         End If
         If rsremu.State = 1 Then rsremu.Close
         mCadSeteo = mCadSeteo & rs2!Codigo
         rs2.MoveNext
      Loop
      totremu = totremu + No_Seteados("IN", mCadSeteo, "N")
      rsboleta.AddNew
      rsboleta!texto = Space(17) & "----------"
      rsboleta.AddNew
      If totremu = 0 Then
         rsboleta!texto = "Sub Total II           " & Format(totremu, "###,###,##0.00")
      Else
         rsboleta!texto = "Sub Total II         " & Format(totremu, "###,###,##0.00")
      End If
      'APORTACIONES
      mCadSeteo = ""
      mlinea = 1
      Sql$ = "select a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
           & "where a.cia='" & wcia & "' and a.status<>'*' and a.tipo='A' " _
           & "and a.cia=b.cia and b.status<>'*' and a.tipo_trab='" & VTipotrab & "' and b.tipomovimiento='03' and a.codigo=b.codinterno " _
           & "order by a.codigo"
      If (fAbrRst(rs2, Sql$)) Then
         Do While Not rs2.EOF
            mbasico = rs(44 + 50 + 21 + Val(rs2!Codigo))
                          
            If mbasico <> 0 Then
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(1) & fCadNum(mbasico, "##,##0.00")
            Else
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(1) & Space(9)
            End If
            rsboleta.AbsolutePosition = mlinea
            rsboleta!texto = Left(rsboleta!texto, 27) & Space(4) & mcad
            mlinea = mlinea + 1
            mCadSeteo = mCadSeteo & rs2!Codigo
            rs2.MoveNext
         Loop
         
         
      End If
      If rs2.State = 1 Then rs2.Close
      Call No_Seteados("AP", mCadSeteo, "")
      'DESCUENTOS"
      mCadSeteo = ""
      Sql$ = "select a.codigo,b.descripcion from  plaseteoprint a,placonstante b " _
           & "where a.cia='" & wcia & "' and a.status<>'*' and a.tipo='D' " _
           & "and a.cia=b.cia and b.status<>'*' and a.tipo_trab='" & VTipotrab & "' and b.tipomovimiento='03' and a.codigo=b.codinterno " _
           & "order by a.codigo"
      If (fAbrRst(rs2, Sql$)) Then
         Do While Not rs2.EOF
            mbasico = rs(44 + 50 + Val(rs2!Codigo))
                  
            If mbasico <> 0 Then
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(13) & fCadNum(mbasico, "##,##0.00")
            Else
               mcad = lentexto(12, Left(rs2!Descripcion, 12)) & Space(13) & Space(9)
            End If
            rsboleta.MoveFirst
            rsboleta.AbsolutePosition = mlinea
            mCadBlanc = ""
            
            If Len(RTrim(rsboleta!texto)) < 27 Then
               For I = Len(RTrim((rsboleta!texto))) To 26
                   mCadBlanc = mCadBlanc & Space(1)
               Next
               rsboleta!texto = RTrim(rsboleta!texto) & mCadBlanc & Space(4) & mcad
            Else
               rsboleta!texto = Left(rsboleta!texto, 27) & Space(4) & mcad
            End If
            
            mlinea = mlinea + 1
            mCadSeteo = mCadSeteo & rs2!Codigo
            rs2.MoveNext
         Loop
      End If
      Call No_Seteados("DE", mCadSeteo, "")
      If rs2.State = 1 Then rs2.Close
   End If
   If rsboleta.RecordCount > 0 Then rsboleta.MoveFirst
   Do While Not rsboleta.EOF
      mtexto = lentextosp(65, Left(rsboleta!texto, 65))
      Print #1, mtexto & Space(11) & mtexto
'      Debug.Print mtexto & Space(4) & mtexto
      rsboleta.MoveNext
   Loop
   Print #1, Space(18) & String(47, "-") & Space(4) & Space(18) & String(47, "-")
   mtexto = "*** TOTAL *** " & fCadNum(rs!totaling, "##,###,##0.00") & "   * TOTAL *    " & fCadNum(rs!totalapo, "###,##0.00") & "  " & fCadNum(rs!totalded, "###,##0.00")
   mtexto = lentexto(65, Left(mtexto, 65))
   Print #1, mtexto & Space(11) & mtexto
'   Debug.Print mtexto & Space(4) & mtexto
   Print #1, String(65, "-") & Space(11) & String(65, "-")
   Print #1, Space(34) & "NETO A PAGAR      " & fCadNum(rs!totaling - rs!totalded, "##,###,##0.00") & Space(4) & Space(34) & "NETO A PAGAR      " & fCadNum(rs!totaling - rs!totalded, "##,###,##0.00")
'   Debug.Print Space(34) & "NETO A PAGAR      " & fCadNum(rs!totaling - rs!totalded, "##,###,##0.00") & Space(4) & Space(34) & "NETO A PAGAR      " & fCadNum(rs!totaling - rs!totalded, "##,###,##0.00")
   'Print #1,
   'Print #1,
   'Print #1,
   'mtexto = "FECHA DE PAGO : " & Format(Left(Cmbfecha.Value, 2), "00") & " DE " & mesnombre & " DE " & Format(mano, "0000")
   mtexto = "FECHA DE PAGO : " & Format(Day(rs!FechaProceso), "00") & " DE " & mesnombre & " DE " & Format(mano, "0000")
   mtexto = lentexto(65, Left(mtexto, 65))
   Print #1, mtexto & Space(11) & mtexto
'   Debug.Print mtexto & Space(4) & mtexto
    Print #1, ""
    
    mcianom = Trae_CIA(wcia, cNomCiaCorto)
    
   mtexto = Left(mcianom & Space(47), 45) & "---------------------"
   mtexto = lentexto(65, Left(mtexto, 65))
   Print #1, mtexto & Space(11) & mtexto
   'If rs.State = 1 Then rs.Close
   If wGrupoPla = "01" Then
      Sql$ = "select rep_nom,cargo_rep_legal from cia where cod_cia='" & wcia & "'"
      If (fAbrRst(rs, Sql$)) Then
         Print #1, lentexto(45, Left(Trim(rs(0) & ""), 45)) & "  Recibi Conforme" & Space(14) & lentexto(45, Left(Trim(rs(0) & ""), 45)) & "   Recibi Conforme"
         Print #1, lentexto(45, Left(Trim(rs(1) & ""), 45)) & Space(31) & lentexto(45, Left(Trim(rs(1) & ""), 45))
         'If rs.State = 1 Then rs.Close
      Else
         Print #1, Space(47) & "Recibi Conforme" & Space(4) & Space(50) & "Recibi Conforme"
      End If
   Else
      Print #1, Space(47) & "Recibi Conforme" & Space(4) & Space(50) & "Recibi Conforme"
   End If
   
  'Pinta Cuenta Corriente
  Sql$ = "Usp_Pla_Pinta_ctacte_Boleta '" & wcia & "','" & Trim(rs!PlaCod & "") & "','" & rs!FechaProceso & "'," & rs!id_boleta & ""
  If (fAbrRst(rs2, Sql$)) Then
      Print #1,
      Print #1, "Cuenta Corriente:" & Space(59) & "Cuenta Corriente:"
      mcad = "Ant." & fCadNum(rs2(0), "###,##0.00") & " Cargo" & fCadNum(rs2(1), "###,##0.00") & " Abono" & fCadNum(rs2(2), "###,##0.00") & " Saldo" & fCadNum(rs2(3), "###,##0.00")
      Print #1, mcad & Space(14) & mcad
  End If
  If rs2.State = 1 Then rs2.Close
   
  If rs.State = 1 Then rs.Close
   If Opcindividual.Value = True Then Exit Do
   AdoCabeza.Recordset.MoveNext
   Print #1, SaltaPag
Loop
Close #1
Panelprogress.Visible = False
FramePrint.Visible = False

Call Imprime_Txt(mArchBol, RUTA$)
Exit Sub

FUNKA:
Close #1
MsgBox "Error : " & Err.Description, vbCritical, "Planillas"
End Sub

Private Sub Print_Quin()
Dim mano As Integer
Dim mmes As Integer
Dim mnombre As String
Dim wciamae As String
Dim mesnombre As String
Dim mcianom As String
Dim mciadep As String
Dim montolet As String
Dim NRODET As Integer
Dim sPlaCodPresentacion As String
NRODET = 0
'Sql$ = "select a.*,dist,b.dpto as dptou from cia a,ubigeos b where cod_cia='" & wcia & "' and a.cod_ubi=b.cod_ubi"

Sql$ = " select c.*, " & _
      " (select nombre from sunat_departamento where id_dpto=left(c.cod_ubi,2)) as dptou, " & _
      " (select nombre from sunat_provincia where id_prov=left(c.cod_ubi,4)) as prov, " & _
      "  u.nombre  as dist, 'PERU' as pais from cia c, sunat_ubigeo u " & _
      "  WHERE cod_cia='" & wcia & "' AND u.id_ubigeo=*c.cod_ubi"


If (fAbrRst(rs, Sql$)) Then
   mcianom = Trim(rs!razsoc)
   mciadep = Trim(rs!dptou & "")
End If
rs.Close
mesnombre = Name_Month(Format(Cmbfecha.Month, "00"))
mano = Val(Mid(Cmbfecha.Value, 7, 4))

RUTA$ = App.Path & "\REPORTS\Quincena.txt"
Open RUTA$ For Output As #1
If Opctotal.Value = True Then AdoCabeza.Recordset.MoveFirst
Do While Not AdoCabeza.Recordset.EOF
   If AdoCabeza.Recordset!totneto > 0 Then montolet = monto_palabras(AdoCabeza.Recordset!totneto) Else montolet = ""
   NRODET = NRODET + 1
   Sql$ = nombre()
   Sql$ = Sql$ & "cargo , placodpresentacion " _
           & "from planillas " _
           & "where cia='" & wcia & "' and placod='" & AdoCabeza.Recordset!PlaCod & "' " _
           & "and status<>'*' "
           mperiodo = Txtsemana.Text
   If (fAbrRst(rs, Sql$)) Then mnombre = lentexto(65, Left(rs!nombre, 65)): sPlaCodPresentacion = Trim(rs!PlaCodpresentacion)
   rs.Close

   Print #1,
   Print #1,
   Print #1, Chr(18) & Space(40) & "NETO A PAGAR S/.       " & fCadNum(AdoCabeza.Recordset!totneto, "#,###,###.00")
   Print #1,
   Print #1,
   Print #1, Chr(15) & Space(5) & "RECIBI DE LA CIA.  " & Chr(18) & mcianom & Chr(15)
   Print #1,
   Print #1, Space(5) & Chr(18) & "LA CANTIDAD DE           " & Chr(15) & AsteriscoR(80, montolet)
   Print #1,
   Print #1,
   Print #1, Space(5) & Chr(18) & "POR CONCEPTO DE           " & Chr(15) & "Pago de la 1ra Quincena del mes de " & mesnombre & " de " & Format(mano, "0000") & Chr(18)
   Print #1,
   Print #1,
   Print #1,
   Print #1,
   Print #1, Space(50) & mciadep & " " & Format(Cmbfecha.Day, "00") & " DE " & mesnombre & " DE " & Format(mano, "0000")
   Print #1,
   Print #1,
   Print #1,
   Print #1,
   Print #1, Chr(15) & Space(87) & "-------------------------------------"
   Print #1, Space(87) & "                FIRMA"
   Print #1,
   Print #1, Space(5) & Chr(18) & "NOMBRE        " & Chr(15) & mnombre
   Print #1, Space(5) & Chr(18) & "CODIGO        " & Chr(15) & sPlaCodPresentacion 'AdoCabeza.Recordset!PlaCod
   Print #1,
   Print #1,
   Print #1,
   Print #1, Space(50) & "-------------------------------------"
   Print #1, Space(50) & mcianom
   If Opcindividual.Value = True Then Exit Do
   AdoCabeza.Recordset.MoveNext
   NRODET = NRODET + 1
   If NRODET > 2 Then
    Print #1, Chr(12) + Chr(13)
    NRODET = 0
   End If
Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("Quincena.txt", RUTA$)
End Sub

Private Sub Print_TxtBcoCredito()
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double

nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If

   If UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      nTrab = nTrab + 1
      mscta = mscta + rsdepo!cta
      mneto = mneto + rsdepo!importe
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\BcoCred.txt"
Open RUTA$ For Output As #1

Dim mcad As String
Dim mt As String
mcad = "1" & Llenar_Ceros(Trim(Str(nTrab)), 6) & Trim(mfproceso)

If VTipo = "01" Then mt = "O"
If VTipo = "11" Then mt = "O"
If VTipo = "02" Then mt = "V"
If VTipo = "03" Then mt = "G"
If VTipo = "04" Then mt = "O"
If wmBolQuin = "Q" Then mt = "O"

If mt = "" Then
   MsgBox "Falda Indicar el tipo de boleta", vbInformation
   Exit Sub
End If

mcad = mcad & mt & "C00011910215470064       " & Llenar_Ceros(Trim(fCadNum(mneto, "#######0.00")), 17)
mcad = mcad & "Planilla de Haberes " & Space(20)

Dim MCOD_CTA As Double
MCOD_CTA = 215470064
mscta = mscta + MCOD_CTA

mcad = mcad & Llenar_Ceros(Trim(Str(mscta)), 15)
Print #1, mcad

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

Do While Not rsdepo.EOF
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!TIPO_REG + rsdepo!Cuenta + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + rsdepo!FLAG
   End If
   rsdepo.MoveNext
 Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("BcoCred.txt", RUTA$)

FrameTxtBco.Visible = False
ChkPreImpreso.Visible = True
chkdsctojudicial.Visible = True
Command3.Visible = False
FrameBco.Visible = False
Frame4.Enabled = True
End Sub

Private Sub Print_TxtBcoScotia()
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\pagscot.txt"
Open RUTA$ For Output As #1

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Dim mcpago As String
mcpago = "Planilla Normal             "

Dim importe As Currency, sImporte As String

Do While Not rsdepo.EOF
    importe = CCur(rsdepo!importe)
    sImporte = Llenar_Ceros(Int(importe * 100), 11)
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!Codigo + Mid(rsdepo!NOM_CLIE, 1, 30) + Mid(mcpago, 1, 20) + Trim(mfproceso) + sImporte + "3" + rsdepo!Sucursal + Mid(rsdepo!PAGOCUENTA, 1, 7) + Mid(rsdepo!NUMERO_DOC, 1, 8) + " " + rsdepo!CTAINTER
   End If
   rsdepo.MoveNext
 Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("pagscot.txt", RUTA$)

FrameTxtBco.Visible = False
ChkPreImpreso.Visible = True
chkdsctojudicial.Visible = True
Command3.Visible = False
FrameBco.Visible = False
Frame4.Enabled = True
End Sub
Private Sub Print_TxtBcoConti()
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double

nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If

   If UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      nTrab = nTrab + 1
      mscta = mscta + rsdepo!cta
      mneto = mneto + rsdepo!importe
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\BcoConti.txt"
Open RUTA$ For Output As #1

Dim mcad As String
Dim mt As String

If VTipo = "01" Then mt = "O"
If VTipo = "11" Then mt = "O"
If VTipo = "02" Then mt = "V"
If VTipo = "03" Then mt = "G"
If VTipo = "04" Then mt = "O"
If wmBolQuin = "Q" Then mt = "O"

If mt = "" Then
   MsgBox "Falda Indicar el tipo de boleta", vbInformation
   Exit Sub
End If

mcad = "700" & "00110686000100006678" & "PEN" & Llenar_Ceros(Trim(fCadNum(mneto * 100, "#######0")), 15) & "A" & Space(9) & "Planilla de Haberes      " & Llenar_Ceros(Trim(fCadNum(nTrab, "#######0")), 6) & "S"
Print #1, mcad & Space(15) & Space(3) & Space(30) & Space(20)

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Dim mDoc As String
Dim Lc_ImportePagoxTrab As Currency
Dim Lc_Acum_ImportePagoxTrab As Currency
Lc_Acum_ImportePagoxTrab = 0
Do While Not rsdepo.EOF
   If Trim(rsdepo!tipo_doc & "") = "1" Then mDoc = "L" Else mDoc = ""
   
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      'Print #1, rsdepo!TIPO_REG + rsdepo!Cuenta + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + rsdepo!FLAG
      'Print #1, "002" & mDoc & Space(4) & Trim(rsdepo!NUMERO_DOC & "") & "P" & Trim(rsdepo!cta & "") & Mid(rsdepo!NOM_CLIE, 1, 40) & Llenar_Ceros(Int(rsdepo!importe * 100), 15) & "Planilla de Haberes" + Space(21) & Space(1) & Space(50) & Space(2) & Space(30) & Space(18)
      
      'add jcms 060516 corrige ajuste de centimos, para que cuadre con la cabecera.
      Lc_ImportePagoxTrab = rsdepo!importe
      Print #1, "002" & mDoc & Trim(rsdepo!NUMERO_DOC & "") & Space(4) & "P" & Trim(rsdepo!cta & "") & Mid(rsdepo!NOM_CLIE, 1, 40) & Llenar_Ceros(Int(Lc_ImportePagoxTrab * 100), 15) & "Planilla de Haberes" + Space(21) & Space(1) & Space(50) & Space(2) & Space(30) & Space(18)
      Lc_Acum_ImportePagoxTrab = Lc_Acum_ImportePagoxTrab + Lc_ImportePagoxTrab
      
   End If
   rsdepo.MoveNext
 Loop
If mneto = Lc_Acum_ImportePagoxTrab Then
    Close #1
    FramePrint.Visible = False
    Call Imprime_Txt("BcoConti.txt", RUTA$)
    FrameTxtBco.Visible = False
    ChkPreImpreso.Visible = True
    chkdsctojudicial.Visible = True
    Command3.Visible = False
    FrameBco.Visible = False
    Frame4.Enabled = True
Else
    MsgBox "Los importes de cabecera y detalle no coinciden", vbCritical, Me.Caption
End If
End Sub

Private Sub Print_Bol_COMACSA()
Dim mano As Integer
Dim mmes As Integer
Dim mperiodo As String
Dim mnombre As String
Dim mbasico As Currency
Dim mcargo As String
Dim mcad As String
Dim mafp As String
Dim mconcep As String
Dim rs2 As ADODB.Recordset
Dim rsremu As ADODB.Recordset
Dim wciamae As String
Dim mesnombre As String
Dim mcianom As String
Dim mciaregpat As String
Dim mciaruc As String
Dim mciadir As String
Dim mciadist As String
Dim mnumh As Integer
Dim totremu As Currency
Dim mtexto As String
Dim mCadBlanc As String
Dim mCadSeteo As String
Dim RX As New ADODB.Recordset
Dim FACTOR_HORAS As Variant
Dim NIC As Variant
Dim bCia As Boolean

'On Error GoTo FUNKA
mtexto = ""
If VTipoPago = "" Or IsNull(VTipoPago) Then Exit Sub
'OBTENEMOS DATOS DE LA CIA
'Sql$ = "select a.*,dist from cia a,ubigeos b where cod_cia='" & wcia & "' and a.cod_ubi=b.cod_ubi"

Sql$ = " select c.*, " & _
      " (select nombre from sunat_departamento where id_dpto=left(c.cod_ubi,2)) as dp, " & _
      " (select nombre from sunat_provincia where id_prov=left(c.cod_ubi,4)) as prov, " & _
      "  u.nombre  as dist, 'PERU' as pais " & _
      " from cia c, sunat_ubigeo u,pla_cia_ubigeo pu " & _
      " WHERE cod_cia='" & wcia & "' AND u.id_ubigeo=pu.cod_ubi and c.cod_cia=pu.cia"
 
If (fAbrRst(rs, Sql$)) Then
   'mcianom = Trim(rs!razsoc)
   mcianom = Trae_CIA(wcia) 'Trim(rs!razsoc)
   mciaregpat = IIf(IsNull(rs!reg_pat), "", Trim(rs!reg_pat))
   mciaruc = rs!RUC
   mciadir = Trim(rs!direcc) & " " & rs!NRO & " " & IIf(IsNull(rs!DIST), "", Trim(rs!DIST))
   mciadist = IIf(IsNull(rs!DIST), "", Trim(rs!DIST))
   'bCia = CBool(IIf(IsNull(rs!Aportacion_EPS), 0, rs!Aportacion_EPS))
End If
If rs.State = 1 Then rs.Close

mano = Val(Mid(Cmbfecha.Value, 7, 4))
mmes = Val(Mid(Cmbfecha.Value, 4, 2))
mesnombre = Name_Month(Format(mmes, "00"))

mlinea = 1
Dim mArchBol As String
mArchBol = "BO" & wcia & VTipotrab & VTipo & ".txt"
RUTA$ = App.Path & "\REPORTS\" & mArchBol
Open RUTA$ For Output As #1

Barra.Max = AdoCabeza.Recordset.RecordCount

If Opctotal.Value = True Then AdoCabeza.Recordset.MoveFirst: Barra.Value = 0

If Opcindividual.Value <> True Then
   Panelprogress.Visible = True
   Panelprogress.ZOrder 0
   Me.Refresh
   If Opcrango.Value = True Then Barra.Value = AdoCabeza.Recordset.AbsolutePosition
End If

Print #1, Chr(27) + "C" + Chr(36) + Chr(15)

Do While Not AdoCabeza.Recordset.EOF
   Barra.Value = AdoCabeza.Recordset.AbsolutePosition
   If rsboleta.RecordCount > 0 Then rsboleta.MoveFirst
   Do While Not rsboleta.EOF
      rsboleta.Delete
      rsboleta.MoveNext
   Loop
   mcianom = Trae_CIA(wcia)
   Sql$ = nombre()
   If VTipoPago = "04" Then mperiodo = Format(mmes, "00") Else mperiodo = Txtsemana.Text
   If VTipo = "02" Then mperiodo = Format(mmes, "00")
   If VTipo = "03" Then mperiodo = Format(mmes, "00")
   
   Sql$ = Sql$ & "a.*,b.fcese,b.fingreso,b.cargo,b.ipss,b.codafp,b.numafp,b.dni,b.nro_doc,b.afiliado_eps_serv,b.placodpresentacion,b.tipo_doc " _
        & "from plahistorico a,planillas b " _
        & "where a.Id_Boleta=" & AdoCabeza.Recordset!id_boleta & " and a.placod=b.placod and a.cia=b.cia and b.status<>'*' order by a.placod"
           
   If (fAbrRst(rs, Sql$)) Then
      mbasico = 0
      rs.MoveFirst
      
      
      'Print #1, lentexto(25, Trim(TxtRefeBol.Text)) & Space(67) & lentexto(25, Trim(TxtRefeBol.Text)) & Chr(27) + Chr(71)
      If Cmbtipo.Text = "UTILIDADES" Then
          Print #1, " UTILIDADES " & lentexto(25, Trim(TxtRefeBol.Text)) & Space(55) & " UTILIDADES " & lentexto(25, Trim(TxtRefeBol.Text)) & Chr(27) + Chr(71)
      Else
          Print #1, lentexto(25, Trim(TxtRefeBol.Text)) & Space(67) & lentexto(25, Trim(TxtRefeBol.Text)) & Chr(27) + Chr(71)
      End If
       Print #1, Chr(27) + Chr(72)
      '1Ra Linea
      mnombre = lentexto(36, Left(rs!nombre, 36))
      bCia = CBool(IIf(IsNull(rs!afiliado_eps_serv), 0, rs!afiliado_eps_serv))
      
      'OBTENEMOS VALOR BASE Y FACTOR DE LA REM DEL TRABAJADOR
      Sql$ = "select importe,FACTOR_HORAS from plaremunbase where cia='" & wcia & "' and concepto='01' and placod='" & Trim(rs!PlaCod & "") & "' and status<>'*'"
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then rs2.MoveFirst: mbasico = rs!Basico: FACTOR_HORAS = rs2("FACTOR_HORAS") Else mbasico = 0
      If rs2.State = 1 Then rs2.Close


      
      mcad = lentexto(7, Left(Trim(rs!PlaCodpresentacion & ""), 7)) & " " & mnombre & " " & fCadNum(mbasico, "#,##,##0.00") & "     " & Format(rs!fIngreso, "dd/mm/yyyy") & "     " & mperiodo
      Print #1, mcad & Space(14) & mcad
      Print #1,
      
      '2Da Linea
      If Not IsNull(rs!fcese) Then mcad = Format(rs!fcese, "dd/mm/yyyy") Else mcad = Space(10)
      'Tipo de documento
      If rs!tipo_doc = "01" Then
         mcad = mcad & "   D.N.I.  " & Space(1)
      Else
         Sql$ = "select descrip from maestros_2 where ciamaestro='01032' and cod_maestro2='" & rs!tipo_doc & "'"
         cn.CursorLocation = adUseClient
         Set rs2 = New ADODB.Recordset
         Set rs2 = cn.Execute(Sql$, 64)
         If rs2.RecordCount > 0 Then rs2.MoveFirst: mcad = mcad & "   " & lentexto(8, Left(rs2!DESCRIP, 8)) & Space(1)
         If rs2.State = 1 Then rs2.Close
      End If
      mcad = mcad & lentexto(10, Left(Trim(rs!nro_doc & ""), 10))
      
      'OBTENEMOS CARGO DEL TRABAJADOR
      Sql$ = "select cod_maestro3,cuenta as descrip from maestros_31 where ciamaestro='" & wcia & "055" & "' AND cod_maestro3='" & rs!Cargo & "'"
      cn.CursorLocation = adUseClient
      Set rs2 = New ADODB.Recordset
      Set rs2 = cn.Execute(Sql$, 64)
      If rs2.RecordCount > 0 Then rs2.MoveFirst: mcargo = rs2!DESCRIP Else mcargo = ""
      If rs2.State = 1 Then rs2.Close
      
      If VTipo = "01" Then
         Dim mdiasemb As Currency
         mdiasemb = 0#
        'OBTENEMOS bonificaciones nov 19 embolsado TRABAJADOR
         Sql$ = "select dias from BON_ENBOLSADO_NOV19 where placod='" & Trim(rs!PlaCod & "") & "' and semana='" & Txtsemana.Text & "' and año=" & Cmbfecha.Year
         cn.CursorLocation = adUseClient
         Set rs2 = New ADODB.Recordset
         Set rs2 = cn.Execute(Sql$, 64)
         If rs2.RecordCount > 0 Then rs2.MoveFirst: mdiasemb = rs2!Dias Else mdiasemb = 0#
         If rs2.State = 1 Then rs2.Close
      End If
      
      mcad = mcad & lentexto(10, Left(Trim(mcargo & ""), 10)) & Space(1)
      
      If VTipo = "02" Then
         mcad = mcad & " " & Format(rs!fechavacai, "dd/mm/yyyy") & "  " & Format(rs!fechavacaF, "dd/mm/yyyy") & "   " & Format(rs!FechaProceso, "dd/mm/yyyy")
      Else
         mcad = mcad & Space(22) & "   " & Format(rs!FechaProceso, "dd/mm/yyyy")
      End If
      
      Print #1, mcad & Space(12) & mcad
      Print #1,
      
      '3Ra Linea
      mcad = lentexto(20, Left(Trim(rs!NUMAFP & ""), 20)) & "   "
      
      'OBTENEMOS NOMBRE DE LA AFP
      If Trim(rs!CodAfp & "") = "01" Or Trim(rs!CodAfp & "") = "02" Or Trim(rs!CodAfp & "") = "99" Then
         mafp = ""
      Else
         wciamae = Determina_Maestro("01069")
         Sql$ = "select descrip from maestros_2 where cod_maestro2='" & rs!CodAfp & "' and status<>'*'"
         Sql$ = Sql$ & wciamae
         cn.CursorLocation = adUseClient
         Set rs2 = New ADODB.Recordset
         Set rs2 = cn.Execute(Sql$, 64)
         If rs2.RecordCount > 0 Then rs2.MoveFirst: mafp = rs2!DESCRIP Else mafp = ""
         If rs2.State = 1 Then rs2.Close
      End If
      mafp = lentexto(20, Left(mafp, 20))
      mcad = mcad & mafp
      
      'lalo
      If VTipo = "03" Then
         mcad = mcad & Space(2) & Space(2) & "          " & Space(2) & "          " & Space(2)
      Else
         mcad = mcad & Space(2) & Formato_Numero_Str(rs!h14, 2) & "          " & Formato_Numero_Str(rs!h22, 2) & "          " & Formato_Numero_Str(rs!h24, 2)
      End If
      
      Print #1, mcad & Space(18) & mcad
      Print #1,
      Print #1,
      
      '4Ta Linea
      mcad = Formato_Numero_Str(rs!h01, 6) & Space(3)           'NORMAL
      mcad = mcad & Formato_Numero_Str(rs!h02, 6) & Space(2)    'DSO
      mcad = mcad & Formato_Numero_Str(rs!h03, 6) & Space(3)   'FERIADOS
      mcad = mcad & Formato_Numero_Str(rs!H04, 6) & Space(1)    'PERM. PAG
      mcad = mcad & Formato_Numero_Str(rs!h05, 6) & Space(2)    'ENFER. PAG
      mcad = mcad & Formato_Numero_Str(rs!H25, 6) & Space(2)    'LIC CON GOCE DE HABER
      
      If rs!h06 > 0 Then
         mcad = mcad & Formato_Numero_Str(rs!h06, 6) & Space(2) 'SUBSIDIO POR ENFERMEDAD O ENFERMEDADES NO PAGADAS
      Else
         mcad = mcad & Formato_Numero_Str(rs!h30, 6) & Space(2) 'SUBSIDIO POR ACC. DE TRABAJO O ENFERMEDADES NO PAGADAS
      End If
      mcad = mcad & Formato_Numero_Str(rs!h08, 6) & Space(2)    'INSISTENCIA INJUSTIFICADA
      mcad = mcad & Formato_Numero_Str(rs!h26, 6) & Space(2)    'LIC. SIN GOCE HABER
      mcad = mcad & Formato_Numero_Str(rs!h09, 6)               'SUSPENCION SIN GOCE DE HABER
          
      Print #1, mcad & Space(10) & mcad
      Print #1,
      
      '5Ta Linea
      mcad = Formato_Numero_Str(rs!h10, 6) & Space(3)           'EXTRAS L-S 25%
      mcad = mcad & Formato_Numero_Str(rs!h17, 6) & Space(2)    'EXTRAS L-S 35%
      mcad = mcad & Formato_Numero_Str(rs!h11, 6) & Space(3)    'EXTRAS DSO Y FERIADOS
      mcad = mcad & Formato_Numero_Str(rs!H12, 6) & Space(1)    'VACACIONES
      mcad = mcad & Formato_Numero_Str(rs!h13, 6) & Space(2)    'SOBRETASA 2DO TURNO
      mcad = mcad & Formato_Numero_Str(rs!h20, 6) & Space(2)    'SOBRETASA 3DO TURNO
      mcad = mcad & Formato_Numero_Str(rs!h29, 6) & Space(2)    'LIC. SINDICAL
      mcad = mcad & Formato_Numero_Str(rs!h23, 6) & Space(2)    'LICENCIA POR PATERNIDAD
      mcad = mcad & Formato_Numero_Str(rs!h15, 6) & Space(2)         'DIA 1RO DE MAYO EN DSO
      mcad = mcad & Formato_Numero_Str(rs!h01 + rs!h02 + rs!h03 + rs!h15 + rs!H04 + rs!h05 + rs!H25 + rs!h23 + rs!h29 + rs!h10 + rs!h17 + rs!h11, 6) ' TOTAL PAGADAS
      
      Print #1, mcad & Space(11) & mcad
      Print #1,
      Print #1,
      
      '6Ta Linea
      mcad = Formato_Numero_Str(rs!i01, 9) & Space(2)           'NORMAL
      mcad = mcad & Formato_Numero_Str(rs!i09, 9) & Space(2)    'DSO
      mcad = mcad & Formato_Numero_Str(rs!i12, 9) & Space(2)    'FERIADO
      mcad = mcad & Formato_Numero_Str(rs!i19, 9) & Space(2)    'PERMISO PAGADO
      mcad = mcad & Formato_Numero_Str(rs!i32, 9) & Space(4)    'ENFERMEDADES PAGADAS
      mcad = mcad & Formato_Numero_Str(rs!i49, 9) & Space(3)    'LIC. CON GOCE DE HABER
      mcad = mcad & Formato_Numero_Str(rs!i02, 9)               'ASIGNACION FAMILIAR
      
      Print #1, mcad & Space(11) & mcad
      Print #1,
      
      '7Ma Linea
      mcad = Formato_Numero_Str(rs!i10, 9) & Space(2)           'EXTRAS L-S 25%
      mcad = mcad & Formato_Numero_Str(rs!i21, 9) & Space(2)    'EXTRAS L-S 35%
      mcad = mcad & Formato_Numero_Str(rs!i11, 9) & Space(2)    'EXTRAS DSO Y FERIADOS
      mcad = mcad & Formato_Numero_Str(rs!i14, 9) & Space(2)    'REMUNERACION VACACIONAL
      mcad = mcad & Formato_Numero_Str(rs!i08, 9) & Space(4)    'SOBRETASA 2DO TURNO
      mcad = mcad & Formato_Numero_Str(rs!i22, 9) & Space(3)    'SOBRETASA 3DO TURNO
      mcad = mcad & Formato_Numero_Str(rs!i04, 9)               'BONIF. TIEM. SERV.
      
      Print #1, mcad & Space(12) & mcad
      Print #1,
      
      Dim iedias As Double
      '8Va Linea
      mcad = Formato_Numero_Str(rs!i27, 9) & Space(2)           'BONIFICACION POR DOBLAR TURNO
      mcad = mcad & Formato_Numero_Str(rs!i26, 9) & Space(2)    'BONIFICACION POR LUBRICAR
      
      If mdiasemb > 0 Then
 
         iedias = Round((((rs!h28 - mdiasemb) * rs!i37) / rs!h28), 2)
         mcad = mcad & Formato_Numero_Str(iedias, 9) & Space(2)    'BONIFICACION POR EMBOLSADO
      Else
         mcad = mcad & Formato_Numero_Str(rs!i37, 9) & Space(2)    'BONIFICACION POR EMBOLSADO
      End If
      mcad = mcad & Formato_Numero_Str(rs!i23, 9) & Space(2)    'BONIFICACION POR HIDRATACION
      mcad = mcad & Formato_Numero_Str(rs!i44, 9) & Space(4)    'BONIFICACION POR REEMPLAZO
      mcad = mcad & Formato_Numero_Str(rs!i36, 9) & Space(3)    'BONIFICACION POR PUNTUALIDAD
      mcad = mcad & Formato_Numero_Str(rs!i28, 9)               'BONIFICACION POR REFRIGERIO
      
           
      Print #1, mcad & Space(11) & mcad
      Print #1,
      
      
      '9Na Linea
      mcad = Formato_Numero_Str(rs!i05, 9) & Space(2)           'AFP 10%
      mcad = mcad & Formato_Numero_Str(rs!i06, 9) & Space(2)    'AFP 3%
      mcad = mcad & Formato_Numero_Str(rs!i48, 9) & Space(2)    'LICENCIA SINDICAL
      mcad = mcad & Formato_Numero_Str(rs!i07, 9) & Space(2)    'LICENCIA POR PATERNIDAD
      mcad = mcad & Formato_Numero_Str(rs!i34, 9) & Space(4)    'DIA 1RO DE MAYO EN DSO
      mcad = mcad & Formato_Numero_Str(rs!i13, 9) & Space(3)    'REINTEGROS
      mcad = mcad & Formato_Numero_Str(rs!i15, 9)               'GRATRIFICACION

 
      Print #1, mcad & Space(11) & mcad
      
      
      '10ma Linea TOTAL REMUNERATIVOS
'      If rs!I16 <> 0 And VTipotrab = "01" Then
'         mcad = Chr(27) + Chr(83) + Chr(1) & "Otros Pagos" & Chr(27) + Chr(84) & Space(79) & Chr(27) + Chr(83) + Chr(1) & "Otros Pagos" & Chr(27) + Chr(84)
'         Print #1, mcad
'      Else
'         Print #1,
'      End If

      'If rs!i38 <> 0 And VTipotrab = "01" Then
      If mdiasemb > 0 Then
         mcad = Chr(27) + Chr(83) + Chr(1) & "           REINT.EMBOLS.            " & Chr(27) + Chr(84) & Space(55) & Chr(27) + Chr(83) + Chr(1) & "           REINT.EMBOLS.            " & Chr(27) + Chr(84)
         Print #1, mcad
      Else
      If rs!I24 <> 0 Then
         mcad = Chr(27) + Chr(83) + Chr(1) & "           BONIF.PRODUC.            " & Chr(27) + Chr(84) & Space(55) & Chr(27) + Chr(83) + Chr(1) & "           BONIF.PRODUC.            " & Chr(27) + Chr(84)
         Print #1, mcad
      Else
        If rs!i38 <> 0 And rs!i29 <> 0 Then
           mcad = Chr(27) + Chr(83) + Chr(1) & "           Bon. x Cump. Comision Vta" & Chr(27) + Chr(84) & Space(55) & Chr(27) + Chr(83) + Chr(1) & "           Bon. x Cump. Comision Vta" & Chr(27) + Chr(84)
           Print #1, mcad
         ElseIf rs!i29 <> 0 Then
           mcad = Chr(27) + Chr(83) + Chr(1) & "                        Comision Vta" & Chr(27) + Chr(84) & Space(55) & Chr(27) + Chr(83) + Chr(1) & "                        Comision Vta" & Chr(27) + Chr(84)
           Print #1, mcad
         ElseIf rs!i38 <> 0 Then
           mcad = Chr(27) + Chr(83) + Chr(1) & "           Bon. x Cump.             " & Chr(27) + Chr(84) & Space(55) & Chr(27) + Chr(83) + Chr(1) & "           Bon. x Cump.             " & Chr(27) + Chr(84)
           Print #1, mcad
         Else
           Print #1,
         End If
     End If
     End If
      'mcad = Formato_Numero_Str(rs!I16, 10) & Space(58) & Formato_Numero_Str(rs!i01 + rs!i09 + rs!i12 + rs!I16 + rs!i19 + rs!i32 + rs!i49 + rs!i02 + rs!i10 + rs!i21 + rs!i11 + rs!i14 + rs!i08 + rs!i22 + rs!i04 + rs!i13 + rs!i15 + rs!i27 + rs!i26 + rs!i37 + rs!i23 + rs!i44 + rs!i36 + rs!i28 + rs!i05 + rs!i06 + rs!i48 + rs!i07, 10)
      mcad = Formato_Numero_Str(rs!I16, 9) & Space(2)
      
      If mdiasemb > 0 Then
         mcad = mcad & Formato_Numero_Str((rs!i37 - iedias), 9) & Space(2)
         mcad = mcad & Formato_Numero_Str(rs!i29, 9) & Space(2)
      Else
      If rs!I24 <> 0 Then
        mcad = mcad & Formato_Numero_Str(rs!I24, 9) & Space(2)
        mcad = mcad & Formato_Numero_Str(rs!i29, 9) & Space(2)
      Else
        mcad = mcad & Formato_Numero_Str(rs!i38, 9) & Space(2)
        mcad = mcad & Formato_Numero_Str(rs!i29, 9) & Space(2)
      End If
      End If
      mcad = mcad & Space(37) & Formato_Numero_Str(rs!i01 + rs!i09 + rs!i12 + rs!i34 + rs!I16 + rs!i19 + rs!i32 + rs!i49 + rs!i02 + rs!i10 + rs!i21 + rs!i11 + rs!i14 + rs!i08 + rs!i22 + rs!i04 + rs!i13 + rs!i15 + rs!i27 + rs!i26 + rs!i37 + rs!i23 + rs!i44 + rs!i36 + rs!i28 + rs!i05 + rs!i06 + rs!i48 + rs!i07 + rs!i38 + rs!i29 + rs!I24, 10)
      
      
      Print #1, mcad & Space(11) & mcad
      Print #1,
      
      '11Va Linea
      mcad = Formato_Numero_Str(rs!i03, 9) & Space(2)           'ASIGNACION POR MOVILIDAD
      mcad = mcad & Formato_Numero_Str(rs!i17, 9) & Space(2)    'ASIGNACION POR ESCOLARIDAD
      mcad = mcad & Formato_Numero_Str(rs!i18, 9) & Space(2)    'UTILIDADES
      mcad = mcad & Formato_Numero_Str(rs!i20, 9) & Space(2)    'VIATICOS
      mcad = mcad & Formato_Numero_Str(rs!i46, 9) & Space(4)    'CANASTA NAVIDEÑA
      mcad = mcad & Formato_Numero_Str(rs!i30, 9) & Space(3)    'BONIFICACION EXTRAORDINARIA
      mcad = mcad & Formato_Numero_Str(rs!i35, 9)               'ASIGNACION POR FALEECIMIENTO
      
      Print #1, mcad & Space(11) & mcad
      Print #1,
      
      '12Va Linea
      mcad = Space(69) & Formato_Numero_Str(rs!i03 + rs!i17 + rs!i18 + rs!i20 + rs!i46 + rs!i30 + rs!i35, 9) 'TOTAL NO REMUNERATIVOS
      
      Print #1, mcad & Space(12) & mcad
      Print #1,
      Print #1,
      
      '13Va Linea
      If rs!h30 <> 0 Then
         mcad = Formato_Numero_Str(0, 9) & Space(2)             'SUBSIDIO POR ENFERMEDAD
         mcad = mcad & Formato_Numero_Str(rs!i43, 9) & Space(2) 'SUBSIDIO POR ACC. DE TRABAJO
      Else
         mcad = Formato_Numero_Str(rs!i43, 9) & Space(2)        'SUBSIDIO POR ENFERMEDAD
         mcad = mcad & Formato_Numero_Str(0, 9) & Space(2)      'SUBSIDIO POR ACC. DE TRABAJO
      End If
      
      mcad = mcad & Formato_Numero_Str(rs!i42, 9) & Space(2)    'SUBSIDIO POR MATERNIDAD
      mcad = mcad & Formato_Numero_Str(0, 9) & Space(2)         'LIBRE
      mcad = mcad & Formato_Numero_Str(0, 9) & Space(4)         'LIBRE
      mcad = mcad & Formato_Numero_Str(0, 9) & Space(3)         'LIBRE
      
      mcad = mcad & Formato_Numero_Str(rs!i43 + rs!i42, 9)      ' TOTAL SUBSIDIOS
      
      Print #1, mcad & Space(11) & mcad
      If (rs!h19 > 0) Then
          mcad = "Tardanza Minutos     -    Tardanza Soles"             'horas tardanzas
      Else
          mcad = Space(40)
      End If
      
      Print #1, mcad + Space(50) + mcad
      
      '14Ava Linea
      'If (rs!h01 + rs!h02 + rs!h03 + rs!h15 + rs!H04 + rs!h05 + rs!H12 + rs!H25 + rs!h29) > 0 Then
      '   mcad = Space(56) & Formato_Numero_Str(Round(rs!totaling / (rs!h01 + rs!h02 + rs!h03 + rs!h15 + rs!H04 + rs!h05 + rs!H12 + rs!H25 + rs!h29) * 8, 2), 9) & Space(4)  'PROM / DIARIO
      'Else
      '   mcad = Space(56) & Formato_Numero_Str(Round(0, 2), 9) & Space(4) 'PROM / DIARIO
      'End If
      '
      'mcad = mcad & Formato_Numero_Str(rs!totaling, 9)      'TOTAL INGRESOS
      '
      'Print #1, mcad & Space(11) & mcad
      'Print #1,
      'Print #1,
 '     If (rs!h19 > 0) Then
 '         mcad = "Tardanza Minutos     -    Tardanza Soles"             'horas tardanzas
 '     Else
 '         mcad = ".                                      ."
 '     End If
     mcad = Space(40)
    ' Tardanza esta ingresada en horas
     If (rs!h19 > 0) Then
          'mcad = Space(5) & Formato_Numero_Str(Round(rs!h19 * 60, 2), 9) + Space(10) & Formato_Numero_Str(Round(rs!tottardanza, 2), 9) + Space(9)
          mcad = Space(5) & Formato_Numero_Str(Int(rs!h19 * 60), 9) + Space(10) & Formato_Numero_Str(Round(rs!tottardanza, 2), 9) + Space(9)

      End If
 
   
      If (rs!h01 + rs!h02 + rs!h03 + rs!h15 + rs!H04 + rs!h05 + rs!H12 + rs!H25 + rs!h29) > 0 Then
         'mcad = Space(56) & Formato_Numero_Str(Round(rs!totaling / (rs!h01 + rs!h02 + rs!h03 + rs!h15 + rs!H04 + rs!h05 + rs!H12 + rs!H25 + rs!H29) * 8, 2), 9) & Space(4)  'PROM / DIARIO
         mcad = mcad + Space(16) & Formato_Numero_Str(Round(rs!totaling / (rs!h01 + rs!h02 + rs!h03 + rs!h15 + rs!H04 + rs!h05 + rs!H12 + rs!H25 + rs!h29) * 8, 2), 9) & Space(4)  'PROM / DIARIO
      Else
         mcad = mcad + Space(16) & Formato_Numero_Str(Round(0, 2), 9) & Space(4) 'PROM / DIARIO
      End If
      
      mcad = mcad & Formato_Numero_Str(rs!totaling, 9)      'TOTAL INGRESOS
      
      Print #1, mcad & Space(11) & mcad
      
     ' If (rs!h19 > 0) Then
     '     mcad = Space(5) & Formato_Numero_Str(Round(rs!h19, 2), 9) + Space(10) & Formato_Numero_Str(Round(rs!tottardanza, 2), 9)
     '    Print #1, mcad & Space(56) & mcad
     ' Else
     '     Print #1,
     ' End If
      Print #1,
      Print #1,
      
      '15Va Linea
      mcad = Formato_Numero_Str(rs!d04, 9) & Space(2)           'SNP
      mcad = mcad & Formato_Numero_Str(rs!d111, 9) & Space(2)   'SPP - APORTACION OBLIGATORIA
      mcad = mcad & Formato_Numero_Str(rs!d112, 9) & Space(2)   'PRIMA DE SEGURO AFP
      If rs!d111 <> 0 And rs!d115 = 0 Then
         mcad = mcad & Formato_Numero_Str(rs!d114, 9) & Space(2)   'COMISION AFP PORCENTUAL POR FLUJO
      Else
         mcad = mcad & Formato_Numero_Str(rs!d115, 9) & Space(2)   'COMISION AFP PORCENTUAL MIXTA
      End If
      mcad = mcad & Formato_Numero_Str(rs!d07, 9) & Space(4)    'CTA. CTE.
      mcad = mcad & Formato_Numero_Str(rs!d20, 9) & Space(3)    'DSCTO. A SOLICITUD DEL SINDICATO
      mcad = mcad & Formato_Numero_Str(rs!d08, 9)               'CUOTA SINDICAL
      
      
      Print #1, mcad & Space(11) & mcad
      Print #1,
      
      '16Va Linea
      mcad = Formato_Numero_Str(rs!d05, 9) & Space(2)               'DESCUENTO JUDICIAL
      If rs!d13 < 0 Then
          mcad = mcad & "(" & Formato_Numero_Str(Abs(rs!d13), 9) & ")"
      Else
          mcad = mcad & Formato_Numero_Str(rs!d13, 9) & Space(2)        'QUINTA CATEGORIA
      End If
      'mcad = mcad & Formato_Numero_Str(rs!d13, 9) & Space(2)        'QUINTA CATEGORIA
      mcad = mcad & Formato_Numero_Str(rs!d06, 9) & Space(2)        'ESSALUD VIDA
      mcad = mcad & Formato_Numero_Str(rs!d09, 9) & Space(2)        'ADELANTO
      mcad = mcad & Formato_Numero_Str(0, 9) & Space(4)             'LIBRE
      mcad = mcad & Formato_Numero_Str(rs!d12, 9) & Space(3)        'OTROS DESCUENTOS
      mcad = mcad & Formato_Numero_Str(rs!totalded, 9)              'TOTAL DSCTO.
      
     
      Print #1, mcad & Space(11) & mcad
      Print #1,
      Print #1,
      
      '17Va Linea
      mcad = Formato_Numero_Str(rs!a01, 9) & Space(2)               'ESSALUD
      mcad = mcad & Formato_Numero_Str(rs!a18, 9) & Space(2)        'SCTR-SALUD
      mcad = mcad & Formato_Numero_Str(rs!a19, 9) & Space(2)        'SCTR-PENSION
      mcad = mcad & Formato_Numero_Str(rs!a03, 9) & Space(2)        'SENATI
      mcad = mcad & Formato_Numero_Str(0, 9) & Space(4)             'LIBRE
      mcad = mcad & Formato_Numero_Str(0, 9) & Space(3)             'LIBRE
      mcad = mcad & Formato_Numero_Str(rs!totalapo, 9)              'TOTAL APORT. PATR.
      
      
      Print #1, mcad & Space(11) & mcad
      Print #1,
      
      '18Va Linea
       'Pinta Cuenta Corriente
      Sql$ = "Usp_Pla_Pinta_ctacte_Boleta '" & wcia & "','" & Trim(rs!PlaCod & "") & "','" & rs!FechaProceso & "'," & rs!id_boleta & ""
      If (fAbrRst(rs2, Sql$)) Then
         mcad = Formato_Numero_Str(rs2(0), 9) & Space(4) & Formato_Numero_Str(rs2(1), 9) & Space(2) & Formato_Numero_Str(rs2(2), 9) & Space(2) & Formato_Numero_Str(rs2(3), 9)
      Else
         mcad = Space(44)
      End If
      mcad = mcad & Space(25) & Formato_Numero_Str(rs!totneto, 9)
      If rs2.State = 1 Then rs2.Close
      Print #1, mcad & Space(11) & mcad
   End If
   If rs.State = 1 Then rs.Close
   If Opcindividual.Value = True Then Exit Do
   AdoCabeza.Recordset.MoveNext
   'Print #1, SaltaPag
   For I = 1 To 6
      Print #1,
   Next
Loop

Close #1
Panelprogress.Visible = False
FramePrint.Visible = False

Call Imprime_Txt(mArchBol, RUTA$)
Exit Sub

FUNKA:
Close #1
MsgBox "Error : " & Err.Description, vbCritical, "Planillas"
End Sub

Private Function Formato_Numero_Str(importe As Double, mLen As Integer) As String
If importe = 0 Then
   Formato_Numero_Str = lentexto(mLen, "")
Else
   If mLen = 2 Then
      Formato_Numero_Str = fCadNum(importe, "##")
   ElseIf mLen = 6 Then
      Formato_Numero_Str = fCadNum(importe, "##0.00")
   ElseIf mLen = 7 Then
      Formato_Numero_Str = fCadNum(importe, "###0.00")
   ElseIf mLen = 8 Then
      Formato_Numero_Str = fCadNum(importe, "#,##0.00")
   ElseIf mLen = 9 Then
      Formato_Numero_Str = fCadNum(importe, "##,##0.00")
   ElseIf mLen = 10 Then
      Formato_Numero_Str = fCadNum(importe, "###,##0.00")
   ElseIf mLen = 11 Then
      Formato_Numero_Str = fCadNum(importe, "####,##0.00")
   ElseIf mLen = 12 Then
      Formato_Numero_Str = fCadNum(importe, "#,###,##0.00")
   ElseIf mLen = 13 Then
      Formato_Numero_Str = fCadNum(importe, "##,###,##0.00")
   End If
End If
End Function


Private Sub Print_TxtBcoCredito_DJ(tipo As String)
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double

nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If

   If UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      nTrab = nTrab + 1
      mscta = mscta + rsdepo!cta
      mneto = mneto + rsdepo!importe
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\BcoCred_DJ_" & tipo & ".txt"
Open RUTA$ For Output As #1

Dim mcad As String
Dim mt As String
mcad = "1" & Llenar_Ceros(Trim(Str(nTrab)), 6) & Trim(mfproceso)

If VTipo = "01" Then mt = "O"
If VTipo = "11" Then mt = "O"
If VTipo = "02" Then mt = "V"
If VTipo = "03" Then mt = "G"
If VTipo = "04" Then mt = "O"
If wmBolQuin = "Q" Then mt = "O"

If mt = "" Then
   MsgBox "Falda Indicar el tipo de boleta", vbInformation
   Exit Sub
End If

mcad = mcad & mt & "C00011910215470064       " & Llenar_Ceros(Trim(fCadNum(mneto, "#######0.00")), 17)
mcad = mcad & "Descuento Judicial  " & Space(20)

Dim MCOD_CTA As Double
MCOD_CTA = 215470064
mscta = mscta + MCOD_CTA

mcad = mcad & Llenar_Ceros(Trim(Str(mscta)), 15)
Print #1, mcad

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst

Do While Not rsdepo.EOF
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!TIPO_REG + rsdepo!Cuenta + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + rsdepo!FLAG
   End If
   rsdepo.MoveNext
 Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("BcoCred_DJ_" & tipo & ".txt", RUTA$)

FrameTxtBco.Visible = False
ChkPreImpreso.Visible = True
chkdsctojudicial.Visible = True
Command3.Visible = False
FrameBco.Visible = False
Frame4.Enabled = True
End Sub

Private Sub Print_TxtBcoScotia_DJ(tipo As String)
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\pagscot_DJ_" & tipo & ".txt"
Open RUTA$ For Output As #1

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Dim mcpago As String
mcpago = "Descuentos Judiciales       "

Do While Not rsdepo.EOF
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!Codigo + Mid(rsdepo!NOM_CLIE, 1, 30) + Mid(mcpago, 1, 20) + Trim(mfproceso) + Llenar_Ceros(Int(rsdepo!importe * 100), 11) + "3" + rsdepo!Sucursal + Mid(rsdepo!PAGOCUENTA, 1, 7) + Mid(rsdepo!NUMERO_DOC, 1, 8) + " " + rsdepo!CTAINTER
   End If
   rsdepo.MoveNext
 Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("pagscot_DJ_" & tipo & ".txt", RUTA$)

FrameTxtBco.Visible = False
ChkPreImpreso.Visible = True
chkdsctojudicial.Visible = True
Command3.Visible = False
FrameBco.Visible = False
Frame4.Enabled = True
End Sub
Private Sub Print_TxtBcoConti_DJ(tipo As String)
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

Dim nTrab As Integer
Dim mneto As Double
Dim mscta As Double

nTrab = 0: mneto = 0
If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If

   If UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      nTrab = nTrab + 1
      mscta = mscta + rsdepo!cta
      mneto = mneto + rsdepo!importe
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\BcoConti_DJ_" & tipo & ".txt"
Open RUTA$ For Output As #1

Dim mcad As String
Dim mt As String

If VTipo = "01" Then mt = "O"
If VTipo = "11" Then mt = "O"
If VTipo = "02" Then mt = "V"
If VTipo = "03" Then mt = "G"
If VTipo = "04" Then mt = "O"
If wmBolQuin = "Q" Then mt = "O"

If mt = "" Then
   MsgBox "Falda Indicar el tipo de boleta", vbInformation
   Exit Sub
End If

mcad = "700" & "00110686000100006678" & "PEN" & Llenar_Ceros(Trim(fCadNum(mneto * 100, "#######0")), 15) & "A" & Space(9) & "Descuentos Judiciales    " & Llenar_Ceros(Trim(fCadNum(nTrab, "#######0")), 6) & "S"
Print #1, mcad & Space(15) & Space(3) & Space(30) & Space(20)

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Dim mDoc As String
Dim Lc_ImportePagoxTrab As Currency
Dim Lc_Acum_ImportePagoxTrab As Currency
Lc_Acum_ImportePagoxTrab = 0
Do While Not rsdepo.EOF
   If Trim(rsdepo!tipo_doc & "") = "1" Then mDoc = "L" Else mDoc = ""
   
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      'Print #1, rsdepo!TIPO_REG + rsdepo!Cuenta + rsdepo!NUMEROCTA + rsdepo!tipo_doc + rsdepo!NUMERO_DOC + Space(3) + rsdepo!NOM_CLIE + rsdepo!REF_TRAB + rsdepo!REF_EMP + rsdepo!moneda + Llenar_Ceros(Trim(fCadNum(rsdepo!importe, "#######0.00")), 17) + rsdepo!FLAG
      'Print #1, "002" & mDoc & Space(4) & Trim(rsdepo!NUMERO_DOC & "") & "P" & Trim(rsdepo!cta & "") & Mid(rsdepo!NOM_CLIE, 1, 40) & Llenar_Ceros(Int(rsdepo!importe * 100), 15) & "Planilla de Haberes" + Space(21) & Space(1) & Space(50) & Space(2) & Space(30) & Space(18)
      
      'add jcms 060516 corrige ajuste de centimos, para que cuadre con la cabecera.
      Lc_ImportePagoxTrab = rsdepo!importe
      Print #1, "002" & mDoc & Trim(rsdepo!NUMERO_DOC & "") & Space(4) & "P" & Trim(rsdepo!cta & "") & Mid(rsdepo!NOM_CLIE, 1, 40) & Llenar_Ceros(Int(Lc_ImportePagoxTrab * 100), 15) & "Planilla de Haberes" + Space(21) & Space(1) & Space(50) & Space(2) & Space(30) & Space(18)
      Lc_Acum_ImportePagoxTrab = Lc_Acum_ImportePagoxTrab + Lc_ImportePagoxTrab
      
   End If
   rsdepo.MoveNext
 Loop
If mneto = Lc_Acum_ImportePagoxTrab Then
    Close #1
    FramePrint.Visible = False
    Call Imprime_Txt("BcoConti_DJ_" & tipo & ".txt", RUTA$)
    FrameTxtBco.Visible = False
    ChkPreImpreso.Visible = True
    chkdsctojudicial.Visible = True
    Command3.Visible = False
    FrameBco.Visible = False
    Frame4.Enabled = True
Else
    MsgBox "Los importes de cabecera y detalle no coinciden", vbCritical, Me.Caption
End If
End Sub
Private Sub CommandButton4_Click()
Dim VBcoPago As String
VBcoPago = fc_CodigoComboBox(CmbBco, 2)
    Select Case VBcoPago
       Case "01": Print_TxtBcoScotia_CCI_Credito
       Case "29": Print_TxtBcoScotia
       Case "02": Print_TxtBcoScotia_CCI_Conti
       Case Else
          MsgBox "Implementación no desarrollada para el banco seleccionado", vbInformation
    End Select

End Sub
Private Sub Print_TxtBcoScotia_CCI_Credito()
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\pagscot_cci_Cred.txt"
Open RUTA$ For Output As #1

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Dim mcpago As String
mcpago = "Planilla Normal             "

Dim importe As Currency, sImporte As String

Do While Not rsdepo.EOF
    importe = CCur(rsdepo!importe)
    sImporte = Llenar_Ceros(Int(importe * 100), 11)
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!Codigo + Mid(rsdepo!NOM_CLIE, 1, 30) + Mid(mcpago, 1, 20) + Trim(mfproceso) + sImporte + "3" + rsdepo!Sucursal + Mid(rsdepo!PAGOCUENTA, 1, 7) + Mid(rsdepo!NUMERO_DOC, 1, 8) + " " + rsdepo!CTAINTER
   End If
   rsdepo.MoveNext
 Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("pagscot_cci_Cred.txt", RUTA$)

FrameTxtBco.Visible = False
ChkPreImpreso.Visible = True
chkdsctojudicial.Visible = True
Command3.Visible = False
FrameBco.Visible = False
Frame4.Enabled = True
End Sub
Private Sub Print_TxtBcoScotia_CCI_Conti()
Dim mfproceso As String
mfproceso = Str(Cbofecha.Year) + Format(Cbofecha.Month, "00") + Format(Cbofecha.Day, "00")

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Do While Not rsdepo.EOF
   If rsdepo!importe < 0 Then
      MsgBox "Existen Netos en negativos, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   If Trim((rsdepo!cta & "") & "") = "" Then
      MsgBox "Existen trabajadores sin cuenta, no se generara el archivo", vbInformation
      rsdepo.Close: Set rsdepo = Nothing
      Exit Sub
   End If
   rsdepo.MoveNext
Loop

RUTA$ = App.Path & "\REPORTS\pagscot_CCI_Cont.txt"
Open RUTA$ For Output As #1

If rsdepo.RecordCount > 0 Then rsdepo.MoveFirst
Dim mcpago As String
mcpago = "Planilla Normal             "

Dim importe As Currency, sImporte As String

Do While Not rsdepo.EOF
    importe = CCur(rsdepo!importe)
    sImporte = Llenar_Ceros(Int(importe * 100), 11)
   If rsdepo!importe > 0 And UCase(Trim(rsdepo!Excluir & "")) <> "S" Then
      Print #1, rsdepo!Codigo + Mid(rsdepo!NOM_CLIE, 1, 30) + Mid(mcpago, 1, 20) + Trim(mfproceso) + sImporte + "3" + rsdepo!Sucursal + Mid(rsdepo!PAGOCUENTA, 1, 7) + Mid(rsdepo!NUMERO_DOC, 1, 8) + " " + rsdepo!CTAINTER
   End If
   rsdepo.MoveNext
 Loop
Close #1
FramePrint.Visible = False
Call Imprime_Txt("pagscot_CCI_Cont.txt", RUTA$)

FrameTxtBco.Visible = False
ChkPreImpreso.Visible = True
chkdsctojudicial.Visible = True
Command3.Visible = False
FrameBco.Visible = False
Frame4.Enabled = True
End Sub

