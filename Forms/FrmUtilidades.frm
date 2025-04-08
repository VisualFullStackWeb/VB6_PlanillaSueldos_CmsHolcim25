VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmUtilidades 
   Caption         =   "Calculo de Utilidades"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16365
   Icon            =   "FrmUtilidades.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   16365
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   16335
      Begin VB.Frame FrameFecRecibo 
         BackColor       =   &H00808080&
         Height          =   735
         Left            =   11280
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   5055
         Begin Threed.SSCommand SSCommand15 
            Height          =   375
            Left            =   3720
            TabIndex        =   38
            Top             =   240
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "Generar"
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
         End
         Begin MSMask.MaskEdBox Txtfecha 
            Height          =   315
            Left            =   2280
            TabIndex        =   40
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648384
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Indique Fecha de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   285
            Width           =   1995
         End
      End
      Begin Threed.SSCommand SSCommand8 
         Height          =   495
         Left            =   10080
         TabIndex        =   24
         Top             =   240
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Dcto. Cta.Cte."
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
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   380
         Width           =   855
      End
      Begin VB.TextBox Txtano 
         Height          =   315
         Left            =   720
         MaxLength       =   4
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   495
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "CALCULAR"
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
      End
      Begin MSAdodcLib.Adodc AdoUtilidad 
         Height          =   330
         Left            =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin Threed.SSCommand SSCommand2 
         Height          =   495
         Left            =   8160
         TabIndex        =   10
         Top             =   240
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Cuentas de Asiento"
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
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   495
         Left            =   5400
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Reporte"
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
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   495
         Left            =   14760
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Ver Asiento"
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
      End
      Begin Threed.SSCommand SSCommand7 
         Height          =   495
         Left            =   6840
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Recibos"
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
      End
      Begin Threed.SSCommand SSCommand9 
         Height          =   495
         Left            =   13200
         TabIndex        =   35
         Top             =   240
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Cruce con Boletas"
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
      End
      Begin Threed.SSCommand SSCommand14 
         Height          =   495
         Left            =   15000
         TabIndex        =   36
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Sustento"
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
      End
      Begin Threed.SSCommand SSCommand16 
         Height          =   495
         Left            =   11640
         TabIndex        =   41
         Top             =   240
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "Agregar Dìas"
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
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(S/.)"
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
         Left            =   3240
         TabIndex        =   8
         Top             =   405
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
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
         Left            =   1800
         TabIndex        =   6
         Top             =   405
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Peiodo"
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
         TabIndex        =   5
         Top             =   405
         Width           =   600
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   315
         Left            =   1335
         TabIndex        =   4
         Top             =   360
         Width           =   255
         Size            =   "450;556"
      End
   End
   Begin VB.Frame FrameCalc 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   16335
      Begin VB.Frame FrameCtaCte 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   3360
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   10455
         Begin Threed.SSCommand SSCommand13 
            Height          =   495
            Left            =   6720
            TabIndex        =   34
            Top             =   360
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   873
            _StockProps     =   78
            Caption         =   "Aceptar"
            ForeColor       =   16711680
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
         Begin Threed.SSCommand SSCommand10 
            Height          =   495
            Left            =   8280
            TabIndex        =   33
            Top             =   360
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   873
            _StockProps     =   78
            Caption         =   "Eliminar"
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
         Begin VB.TextBox TxtImporteCtaCte 
            Height          =   285
            Left            =   5160
            TabIndex        =   29
            Top             =   405
            Width           =   1215
         End
         Begin VB.TextBox TxtCodCtaCte 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   420
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid DgrdCtaCte 
            Bindings        =   "FrmUtilidades.frx":030A
            Height          =   5655
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   9975
            _Version        =   393216
            AllowArrows     =   -1  'True
            Appearance      =   0
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
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
               DataField       =   "Importe"
               Caption         =   "Monto a Descontar"
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
               DataField       =   "Id_Dcto"
               Caption         =   "ID"
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
                  ColumnWidth     =   1319.811
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   6524.788
               EndProperty
               BeginProperty Column02 
                  Alignment       =   1
                  Object.Visible         =   -1  'True
                  ColumnWidth     =   1649.764
               EndProperty
               BeginProperty Column03 
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin Threed.SSCommand SSCommand11 
            Height          =   360
            Left            =   9960
            TabIndex        =   31
            ToolTipText     =   "Salir"
            Top             =   45
            Width           =   465
            _Version        =   65536
            _ExtentX        =   820
            _ExtentY        =   635
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "FrmUtilidades.frx":0322
         End
         Begin Threed.SSCommand SSCommand12 
            Height          =   480
            Left            =   9840
            TabIndex        =   32
            ToolTipText     =   "Salir"
            Top             =   600
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   847
            _StockProps     =   78
            BevelWidth      =   1
            Picture         =   "FrmUtilidades.frx":08BC
         End
         Begin VB.Label LblId 
            BackColor       =   &H8000000C&
            Caption         =   "ID"
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
            Left            =   5280
            TabIndex        =   42
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label LblTrabajadorCtaCte 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1080
            TabIndex        =   28
            Top             =   450
            Width           =   3975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trabajador"
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
            Left            =   45
            TabIndex        =   27
            Top             =   180
            Width           =   930
         End
      End
      Begin VB.Frame FrameAsiento 
         BackColor       =   &H8000000A&
         Height          =   3135
         Left            =   4680
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   6495
         Begin VB.TextBox TxtCta4O 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2400
            TabIndex        =   17
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox TxtCta4E 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2400
            TabIndex        =   15
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox TxtCta9 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2400
            TabIndex        =   13
            Top             =   960
            Width           =   1575
         End
         Begin Threed.SSCommand SSCommand3 
            Height          =   495
            Left            =   4560
            TabIndex        =   19
            Top             =   960
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   873
            _StockProps     =   78
            Caption         =   "Aceptar"
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
         End
         Begin Threed.SSCommand SSCommand4 
            Height          =   495
            Left            =   4560
            TabIndex        =   20
            Top             =   2280
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   873
            _StockProps     =   78
            Caption         =   "Cancelar"
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
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Seteo de Asiento"
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
            Height          =   375
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   6375
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cuenta Obligación por pagar Obrero (4)"
            Height          =   495
            Left            =   360
            TabIndex        =   16
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cuenta Obligación por pagar Empleado (4)"
            Height          =   495
            Left            =   360
            TabIndex        =   14
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "Cuenta de Gasto (9)"
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   960
            Width           =   1935
         End
      End
      Begin MSDataGridLib.DataGrid Dgrdcabeza 
         Bindings        =   "FrmUtilidades.frx":0BD6
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   12303
         _Version        =   393216
         AllowArrows     =   -1  'True
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
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
         ColumnCount     =   9
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
            DataField       =   "totalremu"
            Caption         =   "Total Remuneración"
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
            DataField       =   "calcporremu"
            Caption         =   "Calc. x Remun."
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
            DataField       =   "totalhoras"
            Caption         =   "Total Horas"
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
         BeginProperty Column05 
            DataField       =   "totaldias"
            Caption         =   "Total Días"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "calcpordias"
            Caption         =   "Calc. x Días"
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
            DataField       =   "Total"
            Caption         =   "Total por Pagar"
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
         BeginProperty Column08 
            DataField       =   "monto"
            Caption         =   "MONTO"
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
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4589.858
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1830.047
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1470.047
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoCtaCte 
         Height          =   330
         Left            =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
   End
End
Attribute VB_Name = "FrmUtilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 16485: Me.Height = 8700
Txtano.Text = Year(Date)
End Sub

Private Sub SpinButton1_SpinDown()
   If Txtano.Text = "" Then Txtano.Text = "0"
   Txtano = Txtano + 1
End Sub


Private Sub SpinButton1_SpinUp()
   If Txtano.Text = "" Then Txtano.Text = "0"
   If Txtano.Text > 0 Then Txtano = Txtano - 1
End Sub

Private Sub SSCommand1_Click()

Dim Rq As ADODB.Recordset
Sql$ = ""
'Sql$ = "Select cta9,cta4e,cta4o From PlautilidadAsiento where cia='" & wcia & "' and status<>'*'"
'If fAbrRst(Rq, Sql$) Then
'   If Trim(Rq!cta9 & "") = "" Or Trim(Rq!cta4e & "") = "" Or Trim(Rq!cta4o & "") = "" Then
'      MsgBox "Deben Registrarse las cuentas para el asiento contable primero", vbInformation
'      Rq.Close: Set Rq = Nothing
'      Exit Sub
'   End If
'Else
'   MsgBox "Deben Registrarse las cuentas para el asiento contable primero", vbInformation
'   Rq.Close: Set Rq = Nothing
'   Exit Sub
'End If

Dim xTipo As Integer
If SSCommand1.Caption = "CALCULAR" Then xTipo = 1 Else xTipo = 3
Sql$ = "usp_PlaCalculaUtilidades '" & wcia & "'," & Txtano.Text & "," & CCur(TxtMonto.Text) & ",'" & wuser & "'," & xTipo & ""
cn.Execute Sql$, 64

'Sql$ = "usp_PlaAsientoUtilidades '" & wcia & "'," & Txtano.Text & ",'" & wuser & "','" & wNamePC & "'"
'cn.Execute Sql$, 64

Procesa_Utilidades
End Sub


Private Sub SSCommand10_Click()
If Trim(AdoCtaCte.Recordset!Id_dcto & "") = "" Then Exit Sub
If LblId.Caption = "Dias a Agregar" Then
   Elimina_Dias
Else
   Elimina_CtaCte
End If

End Sub
Private Sub Elimina_CtaCte()
If MsgBox("Desea Eliminar Descuento del trabajador" & Chr(13) & AdoCtaCte.Recordset!nombre, vbYesNo + vbQuestion) = vbYes Then
   Sql$ = "Usp_Pla_Cta_Cte_Utilidades '" & wcia & "'," & Txtano.Text & ",'',0,'" & wuser & "',3," & AdoCtaCte.Recordset!Id_dcto & ""
   cn.Execute Sql$, 64
   Procesa_Cta_Cte
   TxtCodCtaCte.Text = ""
   LblTrabajadorCtaCte.Caption = ""
End If

End Sub
Private Sub Elimina_Dias()
If MsgBox("Desea Eliminar Días del trabajador" & Chr(13) & AdoCtaCte.Recordset!nombre, vbYesNo + vbQuestion) = vbYes Then
   Sql$ = "Usp_Pla_Cta_Cte_Dias '" & wcia & "'," & Txtano.Text & ",'',0,'" & wuser & "',3," & AdoCtaCte.Recordset!Id_dcto & ""
   cn.Execute Sql$, 64
   Procesa_Dias
   TxtCodCtaCte.Text = ""
   LblTrabajadorCtaCte.Caption = ""
End If

End Sub

Private Sub SSCommand11_Click()
FrameCtaCte.Visible = False
Frame2.Enabled = True
End Sub

Private Sub SSCommand12_Click()

Dim Rs As Object
Dim xlApp2 As Excel.Application
Dim xlApp1  As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim I As Integer
Dim Fila As Integer

Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")

mcianom = Trae_CIA(wcia)

xlSheet.Range("B:B").ColumnWidth = 50
xlSheet.Cells(1, 1).Value = Trim(mcianom)
xlSheet.Cells(1, 1).Font.Bold = True

Dim lDias As Boolean
lDias = False
If LblId.Caption = "Dias a Agregar" Then lDias = True

If lDias Then
   xlSheet.Cells(3, 1).Value = "DIAS A AGREGAR PARA EL CALCULO DE UTILIDADES"
Else
   xlSheet.Cells(3, 1).Value = "DESCUENTOS POR CUENTA CORRIENTE A EFECTUAR EN LAS BOLETAS DE UTILIDADES"
End If
xlSheet.Cells(3, 1).HorizontalAlignment = xlCenter
xlSheet.Range("A3:C3").Merge

xlSheet.Cells(4, 1).Value = "CORRESPONDIENTE AL PERIODO " & Txtano.Text
xlSheet.Cells(4, 1).HorizontalAlignment = xlCenter
xlSheet.Range("A4:C4").Merge

xlSheet.Cells(6, 1).Value = "CODIGO"
xlSheet.Cells(6, 2).Value = "NOMBRE"
If lDias Then
   xlSheet.Cells(6, 3).Value = "DIAS"
Else
   xlSheet.Cells(6, 3).Value = "IMPORTE"
End If

xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(6, 3)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(6, 3)).Font.Bold = True
xlSheet.Range(xlSheet.Cells(6, 1), xlSheet.Cells(6, 3)).Borders.LineStyle = xlContinuous
Fila = 7

xlSheet.Range("C:C").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "

If lDias Then
   Procesa_Dias
Else
   Procesa_Cta_Cte
End If

Fila = 7
If AdoCtaCte.Recordset.RecordCount > 0 Then
    Do While Not AdoCtaCte.Recordset.EOF
        xlSheet.Cells(Fila, 1).Value = AdoCtaCte.Recordset!PlaCod
        xlSheet.Cells(Fila, 2).Value = AdoCtaCte.Recordset!nombre
        xlSheet.Cells(Fila, 3).Value = AdoCtaCte.Recordset!importe
        Fila = Fila + 1
        AdoCtaCte.Recordset.MoveNext
    Loop
End If

xlApp2.Application.ActiveWindow.DisplayGridlines = False
xlSheet.Range("A1:A1").Select
xlApp2.Application.Caption = "PROVISION DE CTS"
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True
End Sub

Private Sub SSCommand13_Click()
If LblId.Caption = "Dias a Agregar" Then
   Graba_Dias
Else
   Graba_CtaCte
End If
End Sub
Private Sub Graba_CtaCte()
If Trim(TxtCodCtaCte.Text & "") = "" Then Exit Sub
Sql$ = "Usp_Pla_Cta_Cte_Utilidades '" & wcia & "'," & Txtano.Text & ",'" & TxtCodCtaCte.Text & "'," & CCur(TxtImporteCtaCte.Text) & ",'" & wuser & "',1,0"
cn.Execute Sql$, 64
Procesa_Cta_Cte
TxtCodCtaCte.Text = ""
LblTrabajadorCtaCte.Caption = ""
End Sub

Private Sub Graba_Dias()
If Trim(TxtCodCtaCte.Text & "") = "" Then Exit Sub
Sql$ = "Usp_Pla_Cta_Cte_Dias '" & wcia & "'," & Txtano.Text & ",'" & TxtCodCtaCte.Text & "'," & CCur(TxtImporteCtaCte.Text) & ",'" & wuser & "',1,0"
cn.Execute Sql$, 64
Procesa_Dias
TxtCodCtaCte.Text = ""
LblTrabajadorCtaCte.Caption = ""
End Sub


Private Sub SSCommand14_Click()
Sustento_Utilidades (Txtano.Text)
End Sub

Private Sub SSCommand15_Click()
If Txtfecha.Text <> "__/__/____" And Not IsDate(Txtfecha.Text) Then
   MsgBox "Ingrese Correctamente la Fecha", vbInformation, "Recibos"
   Txtfecha.SetFocus
   Exit Sub
End If

If Val(Right(Txtfecha.Text, 4)) <> Val(Txtano.Text) + 1 Then
   MsgBox "Ingrese Correctamente el Año", vbInformation, "Recibos"
   Txtfecha.SetFocus
   Exit Sub
End If

Call Carga_Recibos_Utilidades(Txtano.Text, Val(Mid(Txtfecha.Text, 1, 2)), Val(Mid(Txtfecha.Text, 4, 2)))
FrameFecRecibo.Visible = False
End Sub

Private Sub SSCommand16_Click()
TxtImporteCtaCte.Text = ""
Procesa_Dias
FrameCtaCte.Visible = True
LblId.Caption = "Dias a Agregar"
DgrdCtaCte.Columns.Item(2).Caption = "DIAS"
Frame2.Enabled = False
End Sub

Private Sub SSCommand2_Click()
Dim Rq As ADODB.Recordset
Sql$ = ""
TxtCta9.Text = "": TxtCta4E.Text = "": TxtCta4O.Text = ""
Sql$ = "Select cta9,cta4e,cta4o from PlautilidadAsiento where cia='" & wcia & "' and status<>'*'"
If fAbrRst(Rq, Sql$) Then
   TxtCta9.Text = Trim(Rq("cta9") & "")
   TxtCta4E.Text = Trim(Rq("cta4e") & "")
   TxtCta4O.Text = Trim(Rq("cta4o") & "")
End If
Rq.Close: Set Rq = Nothing
FrameAsiento.Visible = True
End Sub

Private Sub SSCommand3_Click()
If Trim(TxtCta9.Text) = "" Then MsgBox "Ingrese Cuenta de Gasto", vbInformation: Exit Sub
If Trim(TxtCta4E.Text) = "" Then MsgBox "Ingrese Cuenta Para Empleados", vbInformation: Exit Sub
If Trim(TxtCta4O.Text) = "" Then MsgBox "Ingrese Cuenta Para Obreros", vbInformation: Exit Sub
If Not Valida_Cuenta(Trim(TxtCta9.Text)) Then MsgBox "Cuenta de Gasto No Registrada", vbInformation: Exit Sub
If Not Valida_Cuenta(Trim(TxtCta4E.Text)) Then MsgBox "Cuenta de Empleado No Registrada", vbInformation: Exit Sub
If Not Valida_Cuenta(Trim(TxtCta4O.Text)) Then MsgBox "Cuenta de Obrero No Registrada", vbInformation: Exit Sub

Sql$ = "usp_PlaSeteoUtilidadesAsiento '" & wcia & "','" & Trim(TxtCta9.Text) & "','" & Trim(TxtCta4E.Text) & "','" & Trim(TxtCta4O.Text) & "','" & wuser & "'"
cn.Execute Sql$, 64
FrameAsiento.Visible = False
End Sub

Private Sub SSCommand4_Click()
FrameAsiento.Visible = False
End Sub

Private Sub SSCommand5_Click()
Carga_Utilidades (Txtano.Text)
End Sub

Private Sub SSCommand6_Click()
Call Carga_Asiento_Excel(Txtano.Text, 12, "PPU", "1103000001", "PROVISION DE PARTICIPACION DE TRABAJADORES EJERCICIO " & Txtano.Text, "DICIEMBRE", "")
End Sub

Private Sub SSCommand7_Click()
FrameFecRecibo.Visible = True
FrameFecRecibo.ZOrder 0

End Sub

Private Sub SSCommand8_Click()
TxtImporteCtaCte.Text = ""
Procesa_Cta_Cte
FrameCtaCte.Visible = True
LblId.Caption = "Importe a Descontar"
DgrdCtaCte.Columns.Item(2).Caption = "IMPORTE"
Frame2.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)


End Sub

Private Sub SSCommand9_Click()
UtilidadesVsBoletas (Txtano.Text)
End Sub

Private Sub Txtano_Change()
Procesa_Utilidades
End Sub
Private Sub Procesa_Utilidades()
Sql$ = "usp_PlaCargaUtilidades '" & wcia & "'," & Txtano.Text & ""
cn.CursorLocation = adUseClient
Set AdoUtilidad.Recordset = cn.Execute(Sql$, 64)
If AdoUtilidad.Recordset.RecordCount > 0 Then
   SSCommand1.Caption = "ELIMINAR"
   TxtMonto.Text = Format(AdoUtilidad.Recordset(9), "###,###.00")
   TxtMonto.Enabled = False
   SSCommand6.Enabled = True
   SSCommand5.Enabled = True
   SSCommand7.Enabled = True
Else
   SSCommand1.Caption = "CALCULAR"
   TxtMonto.Text = ""
   TxtMonto.Enabled = True
   SSCommand6.Enabled = False
   SSCommand5.Enabled = False
   SSCommand7.Enabled = False
End If
End Sub

Private Sub Verifica_Trabajador()

If TxtCodCtaCte.Text = "" Then LblTrabajadorCtaCte.Caption = "": Exit Sub

Sql$ = Funciones.nombre()
Sql$ = Sql$ & "status from planillas where status<>'*' "
Sql$ = Sql$ & " and cia='" & wcia & "' AND placod='" & Trim(TxtCodCtaCte) & "'"


cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$)
If Rs.RecordCount <= 0 Then
   MsgBox "Codigo de Trabajador no Registrado", vbInformation
   Rs.Close: Set Rs = Nothing
   TxtCodCtaCte.Text = ""
   LblTrabajadorCtaCte.Caption = ""
   Exit Sub
End If
LblTrabajadorCtaCte.Caption = Trim(Rs!nombre)
TxtCodCtaCte.Text = UCase(TxtCodCtaCte.Text)
Rs.Close: Set Rs = Nothing

End Sub

Private Sub TxtCodCtaCte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtImporteCtaCte.SetFocus
End Sub

Private Sub TxtCodCtaCte_LostFocus()
Verifica_Trabajador
If TxtCodCtaCte.Text = "" Then Exit Sub

End Sub
Private Sub Procesa_Cta_Cte()
Sql$ = "select u.placod,rtrim(p.ap_pat)+' '+rtrim(p.ap_mat)+' '+rtrim(p.nom_1)+' '+rtrim(p.nom_2) as nombre ,Importe,Id_Dcto "
Sql$ = Sql$ & "from pla_cta_cte_utilidades u,planillas p "
Sql$ = Sql$ & "Where u.cia='" & wcia & "' and u.Ano=" & Txtano.Text & " "
Sql$ = Sql$ & "And p.cia=u.cia and u.placod=p.placod and u.status<>'*' and p.status<>'*' order by u.placod"
cn.CursorLocation = adUseClient
Set AdoCtaCte.Recordset = cn.Execute(Sql$, 64)
End Sub

Private Sub Procesa_Dias()
Sql$ = "select u.placod,rtrim(p.ap_pat)+' '+rtrim(p.ap_mat)+' '+rtrim(p.nom_1)+' '+rtrim(p.nom_2) as nombre ,Dias as Importe,Id_Dias as Id_dcto "
Sql$ = Sql$ & "from pla_dias_utilidades u,planillas p "
Sql$ = Sql$ & "Where u.cia='" & wcia & "' and u.Ano=" & Txtano.Text & " "
Sql$ = Sql$ & "And p.cia=u.cia and u.placod=p.placod and u.status<>'*' and p.status<>'*' order by u.placod"
cn.CursorLocation = adUseClient
Set AdoCtaCte.Recordset = cn.Execute(Sql$, 64)
End Sub

