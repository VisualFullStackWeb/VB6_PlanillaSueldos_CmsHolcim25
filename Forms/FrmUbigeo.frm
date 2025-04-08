VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmUbigeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubigeos"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "FrmUbigeo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Filtrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9120
      TabIndex        =   23
      Top             =   120
      Width           =   645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   9975
      Begin MSAdodcLib.Adodc AdoUbigeos 
         Height          =   330
         Left            =   1200
         Top             =   1800
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
         Caption         =   "Ubigeos"
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
      Begin MSDataGridLib.DataGrid DgrdUbigeos 
         Bindings        =   "FrmUbigeo.frx":030A
         Height          =   3255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5741
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "codpais"
            Caption         =   "Pais"
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
            DataField       =   "coddpto"
            Caption         =   "Dpto"
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
            DataField       =   "codprov"
            Caption         =   "Prov"
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
            DataField       =   "coddist"
            Caption         =   "Dist."
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
            DataField       =   "cod_postal"
            Caption         =   "Postal"
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
            DataField       =   "pais"
            Caption         =   "              Pais"
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
            DataField       =   "dpto"
            Caption         =   "              Dpto."
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
            DataField       =   "prov"
            Caption         =   "              Prov."
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
            DataField       =   "dist"
            Caption         =   "               Dist."
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
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   420.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1785.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1679.811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1769.953
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtcriterio 
      Height          =   315
      Left            =   3150
      TabIndex        =   4
      Top             =   120
      Width           =   5385
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2205
      Left            =   0
      TabIndex        =   8
      Top             =   3960
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3889
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos del Lugar"
      TabPicture(0)   =   "FrmUbigeo.frx":0323
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "AdoProv"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "AdoDpto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "AdoPais"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "sscmd(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "sscmd(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "sscmd(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCerrar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Cmbpais"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbdpto"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmbprov"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Txtdist"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Txtcodpostal"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Buscar por"
      TabPicture(1)   =   "FrmUbigeo.frx":033F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ssopt1(4)"
      Tab(1).Control(1)=   "ssopt1(3)"
      Tab(1).Control(2)=   "ssopt1(2)"
      Tab(1).Control(3)=   "ssopt1(1)"
      Tab(1).Control(4)=   "ssopt1(0)"
      Tab(1).ControlCount=   5
      Begin VB.TextBox Txtcodpostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8520
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtdist 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         TabIndex        =   27
         Top             =   1560
         Width           =   3735
      End
      Begin MSDataListLib.DataCombo cmbprov 
         Bindings        =   "FrmUbigeo.frx":035B
         Height          =   315
         Left            =   6120
         TabIndex        =   26
         Top             =   1080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "prov"
         BoundColumn     =   "codprov"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbdpto 
         Bindings        =   "FrmUbigeo.frx":0371
         Height          =   315
         Left            =   1200
         TabIndex        =   25
         Top             =   1560
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "dpto"
         BoundColumn     =   "coddpto"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo Cmbpais 
         Bindings        =   "FrmUbigeo.frx":0387
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   1080
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ListField       =   "pais"
         BoundColumn     =   "codpais"
         Text            =   ""
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Default         =   -1  'True
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
         Left            =   8760
         TabIndex        =   22
         Top             =   0
         Width           =   855
      End
      Begin Threed.SSCommand sscmd 
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   9
         Top             =   480
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Nuevo"
      End
      Begin Threed.SSOption ssopt1 
         Height          =   195
         Index           =   0
         Left            =   -74430
         TabIndex        =   10
         Top             =   600
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "País"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption ssopt1 
         Height          =   195
         Index           =   1
         Left            =   -74430
         TabIndex        =   11
         Top             =   930
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Departamento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption ssopt1 
         Height          =   195
         Index           =   2
         Left            =   -74430
         TabIndex        =   12
         Top             =   1260
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Provincia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption ssopt1 
         Height          =   195
         Index           =   3
         Left            =   -74430
         TabIndex        =   13
         Top             =   1590
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Distrito"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSCommand sscmd 
         Height          =   285
         Index           =   1
         Left            =   4260
         TabIndex        =   14
         Top             =   480
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Actualizar"
      End
      Begin Threed.SSOption ssopt1 
         Height          =   195
         Index           =   4
         Left            =   -74430
         TabIndex        =   15
         Top             =   1890
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Código Postal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand sscmd 
         Height          =   285
         Index           =   2
         Left            =   1980
         TabIndex        =   16
         Top             =   480
         Width           =   1275
         _Version        =   65536
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "Modificar"
      End
      Begin MSAdodcLib.Adodc AdoPais 
         Height          =   330
         Left            =   2160
         Top             =   1080
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
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
         Caption         =   "AdoPais"
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
      Begin MSAdodcLib.Adodc AdoDpto 
         Height          =   330
         Left            =   1920
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "AdoDpto"
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
      Begin MSAdodcLib.Adodc AdoProv 
         Height          =   330
         Left            =   6240
         Top             =   1080
         Visible         =   0   'False
         Width           =   1920
         _ExtentX        =   3387
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
         Caption         =   "AdoProv"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "País"
         Height          =   195
         Index           =   0
         Left            =   690
         TabIndex        =   21
         Top             =   1140
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   20
         Top             =   1590
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
         Height          =   195
         Index           =   2
         Left            =   5220
         TabIndex        =   19
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Distrito"
         Height          =   195
         Index           =   3
         Left            =   5370
         TabIndex        =   18
         Top             =   1620
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Postal"
         Height          =   195
         Index           =   4
         Left            =   7560
         TabIndex        =   17
         Top             =   510
         Width           =   810
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Criterio"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label lblcod1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblcod4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblcod3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblcod2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FrmUbigeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim COL1 As MSDataGridLib.Column, COL2 As MSDataGridLib.Column, COL3 As MSDataGridLib.Column, col4 As MSDataGridLib.Column, ColPostal As MSDataGridLib.Column
Dim COL5 As MSDataGridLib.Column, COL6 As MSDataGridLib.Column, COL7 As MSDataGridLib.Column, COL8 As MSDataGridLib.Column

Dim Nuevo As Boolean, FLAG As Boolean
Dim First As Boolean
Dim Ubi As String
Public TXTOPCION As Integer

Private Sub cmbdpto_Click(Area As Integer)
If cmbdpto <> "" And Area = 2 Then
   cn.CursorLocation = adUseClient
   Set AdoProv.Recordset = cn.Execute("SELECT distinct prov,codprov from ubigeos WHERE codpais='" & Cmbpais.BoundText & "' and coddpto='" & cmbdpto.BoundText & "' ", 64)
   cmbprov = ""
   cmbprov.ListField = "prov"
   cmbprov.BoundColumn = "codprov"
End If
End Sub

Private Sub Cmbpais_Click(Area As Integer)
If Cmbpais <> "" And Area = 2 Then
   cn.CursorLocation = adUseClient
   Set AdoDpto.Recordset = cn.Execute("SELECT distinct dpto,coddpto from ubigeos WHERE codpais='" & Cmbpais.BoundText & "'", 64)
   cmbdpto = ""
   cmbprov = ""
   cmbdpto.ListField = "dpto"
   cmbdpto.BoundColumn = "CODDPTO"
   Set AdoProv.Recordset = Nothing
End If

End Sub

Private Sub cmdbuscar_Click()
If Trim(txtcriterio) <> "" Then
If txtcriterio = "*" Then txtcriterio = "%"

Dim Qry$
  If ssopt1(0).Value = True Then
    Qry$ = "SELECT Cod_ubi,codpais,cod_postal,coddpto,codprov,coddist,pais,dpto,prov,dist FROM ubigeos WHERE" _
             & " pais LIKE '" & Trim(txtcriterio) + "%" & "' ORDER BY pais,dpto,prov,dist"
  ElseIf ssopt1(1).Value = True Then
    Qry$ = "SELECT Cod_ubi,codpais,cod_postal,coddpto,codprov,coddist,pais,dpto,prov,dist FROM ubigeos WHERE" _
             & " dpto LIKE '" & Trim(txtcriterio) + "%" & "' ORDER BY dpto,prov,dist"
  ElseIf ssopt1(2).Value = True Then
    Qry$ = "SELECT Cod_ubi,codpais,cod_postal,coddpto,codprov,coddist,pais,dpto,prov,dist FROM ubigeos WHERE" _
             & " prov LIKE '" & Trim(txtcriterio) + "%" & "' ORDER BY prov,dist"
  ElseIf ssopt1(3).Value = True Then
  
    Qry$ = "SELECT Cod_ubi,codpais,cod_postal,coddpto,codprov,coddist,pais,dpto,prov,dist FROM ubigeos WHERE" _
             & " dist LIKE '" & Trim(txtcriterio) + "%" & "' ORDER BY dist"
  Else
    Qry$ = "SELECT Cod_ubi,codpais,cod_postal,coddpto,codprov,coddist,pais,dpto,prov,dist FROM ubigeos WHERE" _
             & " cod_postal LIKE '" & Trim(txtcriterio) + "%" & "' ORDER BY codpais,cod_postal"
  End If
cn.CursorLocation = adUseClient
Set AdoUbigeos.Recordset = cn.Execute(Qry$, 64)

If txtcriterio = "%" Then txtcriterio = "*"
DgrdUbigeos.SetFocus
txtcriterio.SetFocus
End If
FLAG = True
End Sub
Private Sub DgrdUbigeos_DblClick()
Dim Lugar As String

If Not FLAG Then Exit Sub
lblcod1 = COL1.Value
lblcod2 = COL2.Value
lblcod3 = COL3.Value
lblcod4 = col4.Value
Lugar = Trim(COL5.Value & " - " & COL6.Value & " - " & COL7.Value & " - " & COL8.Value)
Ubi = COL1.Value & COL2.Value & COL3.Value & col4.Value
  Select Case MDIplared.ActiveForm.Name
        Case "Frmpersona"
            If TXTOPCION = 0 Then
              Frmpersona.Text13.Text = Lugar
              Frmpersona.Text13.Tag = Ubi
            ElseIf TXTOPCION = 1 Then
               'Frmpersona.Text2.Text = Lugar
               'Frmpersona.Text2.Tag = Ubi
            End If
        Case "Frmcia"
              Frmcia.lbllugar = Lugar
              Frmcia.Ciacodubi = Ubi
              Frmcia.lbllugar.Tag = Ubi
        Case "Frmobras"
              Frmobras.LblUbica = Lugar
              Frmobras.Lblubigeo = Ubi
  End Select

Unload Me
End Sub
Private Sub DgrdUbigeos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Salir
 Nuevo = False
 If LastRow <> "" Then
 lblcod1 = COL1.Value
 lblcod2 = COL2.Value
 lblcod3 = COL3.Value
 lblcod4 = col4.Value
 Nuevo = False
 Ubi = COL1.Value & COL2.Value & COL3.Value & col4.Value

 Cmbpais = Trim(COL5.Text)
 cmbdpto = Trim(COL6.Text)
 cmbprov = Trim(COL7.Text)
 Txtdist = Trim(COL8.Text)
 Txtcodpostal = Trim(ColPostal)
 End If
 
 Cmbpais.Enabled = False
 cmbdpto.Enabled = False
 cmbprov.Enabled = False
 Txtdist.Enabled = False
 Txtcodpostal.Enabled = False
Salir: Exit Sub
End Sub

Private Sub Form_Load()
Set COL1 = DgrdUbigeos.Columns(0)
Set COL2 = DgrdUbigeos.Columns(1)
Set COL3 = DgrdUbigeos.Columns(2)
Set col4 = DgrdUbigeos.Columns(3)
Set ColPostal = DgrdUbigeos.Columns(4)
Set COL5 = DgrdUbigeos.Columns(5)
Set COL6 = DgrdUbigeos.Columns(6)
Set COL7 = DgrdUbigeos.Columns(7)
Set COL8 = DgrdUbigeos.Columns(8)

FLAG = False
End Sub

Private Sub sscmd_Click(Index As Integer)
Select Case Index
 Case 0
        Me.MousePointer = 11
        Cmbpais.Enabled = True
        cmbdpto.Enabled = True
        cmbprov.Enabled = True
        Txtdist.Enabled = True
        Txtcodpostal.Enabled = True
        
        Cmbpais.BoundText = ""
        Cmbpais.Text = ""
        cmbdpto.BoundText = ""
        cmbdpto.Text = ""
        cmbprov.BoundText = ""
        cmbprov.Text = ""
        Txtdist.Text = ""
        Txtcodpostal.Text = ""
        
        If First = False Then
          cn.CursorLocation = adUseClient
          Set AdoPais.Recordset = cn.Execute("SELECT distinct pais,codpais from ubigeos", 64)
          First = True
        End If

        Nuevo = True
        GoTo Termina
 Case 1
       If MsgBox("Desea Actualizar ", vbQuestion + vbYesNo) = vbNo Then GoTo Termina
       If Trim(Cmbpais) = "" Then
         MsgBox "País no puede ser en blanco", vbCritical
         Cmbpais.SetFocus
         GoTo Termina
       ElseIf Trim(cmbdpto) = "" Then
         MsgBox "Departamento no puede ser en blanco", vbCritical
         cmbdpto.SetFocus
         GoTo Termina
       ElseIf Trim(cmbprov) = "" Then
         MsgBox "Departamento no puede ser en blanco", vbCritical
         cmbdpto.SetFocus
         GoTo Termina
       ElseIf Trim(Txtdist) = "" Then
         MsgBox "Distrito no puede ser en blanco", vbCritical
         Txtdist.SetFocus
         GoTo Termina
       ElseIf Txtcodpostal = "" Then
         MsgBox "Cod. Postal no puede ser en blanco", vbCritical
         Txtcodpostal.SetFocus
         GoTo Termina
       End If
       If Nuevo = True Then
       Dim DPTO$, PROV$, DIST$
       
       If Trim(Cmbpais.BoundText) = "" Then
         Sql = "select max(codpais) from ubigeos "
         cn.CursorLocation = adUseClient
         Set rs = New ADODB.Recordset
         Set rs = cn.Execute(Sql)

         AdoUbigeos.Recordset.AddNew
         'AdoUbigeos.Recordset.EditMode
         AdoUbigeos.Recordset!codpais = Format(rs(0) + 1, "000")
         AdoUbigeos.Recordset!coddpto = "01"
         AdoUbigeos.Recordset!codprov = "01"
         AdoUbigeos.Recordset!CODDIST = "01"
         AdoUbigeos.Recordset!cod_ubi = Format(rs(0) + 1, "000") & "010101"
       Else
         If Trim(cmbdpto.BoundText) = "" Then
            Sql = "select max(coddpto) from ubigeos where codpais='" & Cmbpais.BoundText & "'"
            cn.CursorLocation = adUseClient
            Set rs = New ADODB.Recordset
            Set rs = cn.Execute(Sql)
            DPTO$ = Format(rs(0) + 1, "00"): PROV$ = "01": DIST$ = "01"
         Else
            If Trim(cmbprov.BoundText) = "" Then
               Sql = "select max(codprov) from ubigeos where codpais='" & Cmbpais.BoundText & "' AND CODDPTO='" & cmbdpto.BoundText & "'"
               cn.CursorLocation = adUseClient
               Set rs = New ADODB.Recordset
               Set rs = cn.Execute(Sql)
               PROV$ = Format(rs(0) + 1, "00"): DPTO$ = cmbdpto.BoundText: DIST$ = "01"
            Else
               Sql = "select max(codDIST) from ubigeos " & _
               "where codpais='" & Cmbpais.BoundText & _
               "' AND CODDPTO='" & cmbdpto.BoundText & _
               "' AND CODPROV='" & cmbprov.BoundText & "'"
               cn.CursorLocation = adUseClient
               Set rs = New ADODB.Recordset
               Set rs = cn.Execute(Sql)
               DIST$ = Format(rs(0) + 1, "00")
               PROV$ = cmbprov.BoundText
               DPTO$ = cmbdpto.BoundText
            End If
         End If
         
         Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
         Sql = Sql & " INSERT INTO ubigeos VALUES ('" & Cmbpais.BoundText & DPTO$ & PROV$ & DIST$ & "','" & _
                Cmbpais.BoundText & "','" & DPTO$ & "','" & PROV$ & "','" & DIST & "','" & _
                Cmbpais.Text & "','" & cmbdpto.Text & "','" & cmbprov.Text & "','" & _
                Trim(Txtdist) & "','" & Trim(Txtcodpostal) & "','C','" & wuser & "'," & FechaSys & _
                ",'','')"
         
         cn.Execute Sql, rdExecDirect
       End If
  Else
  
       'AdoUbigeos.Recordset.EditMode
'       AdoUbigeos.Recordset!pais = Trim(Cmbpais)
'       AdoUbigeos.Recordset!DPTO = Trim(cmbdpto.Text)
'       AdoUbigeos.Recordset!PROV = Trim(cmbprov.Text)
'       AdoUbigeos.Recordset!DIST = Trim(Txtdist)
'       AdoUbigeos.Recordset!cod_postal = Trim(Txtcodpostal)
'       AdoUbigeos.Recordset.Update
  End If
       txtcriterio.Text = Trim(Txtdist.Text)
       cmdbuscar_Click
       DgrdUbigeos.SetFocus
       
        Cmbpais.Enabled = False
        cmbdpto.Enabled = False
        cmbprov.Enabled = False
        Txtdist.Enabled = False
        Txtcodpostal.Enabled = False
        
Case 2
        Cmbpais.Enabled = True
        cmbdpto.Enabled = True
        cmbprov.Enabled = True
        Txtdist.Enabled = True
        Txtcodpostal.Enabled = True
End Select
Termina:
   Me.MousePointer = vbDefault
   Exit Sub
End Sub

Private Sub txtcriterio_GotFocus()
sscmd(1).default = False
cmdbuscar.default = True
End Sub


