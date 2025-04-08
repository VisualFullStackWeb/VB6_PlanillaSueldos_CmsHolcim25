VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSeteoAsiento 
   Caption         =   "Seteo de Cuentas Contables"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   Icon            =   "FrmSeteoAsiento.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   9360
   Begin Threed.SSPanel PnlModi 
      Height          =   1935
      Left            =   240
      TabIndex        =   40
      Top             =   2640
      Visible         =   0   'False
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   3413
      _StockProps     =   15
      BackColor       =   4210752
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
      Begin VB.TextBox TxtCtag 
         Height          =   285
         Left            =   7800
         TabIndex        =   81
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtCtav 
         Height          =   285
         Left            =   6680
         TabIndex        =   80
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtCta4 
         Height          =   285
         Left            =   7845
         TabIndex        =   78
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TxtCta 
         Height          =   285
         Left            =   5565
         TabIndex        =   44
         Top             =   825
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta G"
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
         Height          =   195
         Left            =   7800
         TabIndex        =   84
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta V"
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
         Height          =   195
         Left            =   6720
         TabIndex        =   83
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta N"
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
         Height          =   195
         Left            =   5640
         TabIndex        =   82
         Top             =   600
         Width           =   810
      End
      Begin VB.Label LblCta4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta 4"
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
         Height          =   195
         Left            =   6960
         TabIndex        =   79
         Top             =   1485
         Width           =   780
      End
      Begin VB.Label LblIdSeteo 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSForms.CommandButton SSCommand4 
         Height          =   375
         Left            =   8385
         TabIndex        =   47
         Top             =   120
         Width           =   435
         ForeColor       =   4210752
         PicturePosition =   327683
         Size            =   "767;661"
         Picture         =   "FrmSeteoAsiento.frx":030A
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label LblTipoConcepto 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   46
         Top             =   1440
         Width           =   375
      End
      Begin MSForms.CommandButton CommandButton3 
         Height          =   375
         Left            =   5040
         TabIndex        =   45
         Top             =   1365
         Width           =   1755
         ForeColor       =   4210752
         Caption         =   "   Aceptar"
         PicturePosition =   327683
         Size            =   "3096;661"
         Picture         =   "FrmSeteoAsiento.frx":075C
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label LblConcepto 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label LblCod 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   375
      End
      Begin VB.Label LblTit 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
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
         Left            =   120
         TabIndex        =   41
         Top             =   180
         Width           =   8220
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   40
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.ComboBox CmbConcepto 
         Height          =   315
         ItemData        =   "FrmSeteoAsiento.frx":0CF6
         Left            =   1200
         List            =   "FrmSeteoAsiento.frx":0D03
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   200
         Width           =   7935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Conceptos"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
   End
   Begin MSAdodcLib.Adodc AdoAsiento 
      Height          =   375
      Left            =   0
      Top             =   0
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
      Caption         =   "AdoAsiento"
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
   Begin VB.Frame FramIng 
      Height          =   8895
      Left            =   40
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   9255
      Begin VB.TextBox TxtCts 
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox TxtGrati 
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   8400
         Width           =   1335
      End
      Begin VB.TextBox TxtVaca 
         Height          =   285
         Left            =   360
         TabIndex        =   9
         Top             =   8400
         Width           =   1335
      End
      Begin VB.ComboBox CmbCcosto 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   7455
      End
      Begin VB.ComboBox CmbTrabajador 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   7455
      End
      Begin MSDataGridLib.DataGrid DGrdRemunera 
         Bindings        =   "FrmSeteoAsiento.frx":0D2C
         Height          =   6135
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   10821
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "codinterno"
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
         BeginProperty Column02 
            DataField       =   "cuenta"
            Caption         =   "CuentaN"
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
            DataField       =   "Id_Seteo"
            Caption         =   "Id_Seteo"
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
            DataField       =   "cuentav"
            Caption         =   "CuentaV"
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
            DataField       =   "cuentag"
            Caption         =   "CuentaG"
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
            DataField       =   "cuenta2"
            Caption         =   "Cuenta 4"
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
               Object.Visible         =   0   'False
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4110.236
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin VB.Label LblIdSeteoIC 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   7080
         TabIndex        =   51
         Top             =   8400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblIdSeteoIG 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6600
         TabIndex        =   50
         Top             =   8400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblIdSeteoIV 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   6120
         TabIndex        =   49
         Top             =   8400
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSForms.CommandButton SSCommand3 
         Height          =   435
         Left            =   4440
         TabIndex        =   37
         Top             =   8340
         Width           =   1395
         ForeColor       =   4210752
         Caption         =   "   Aceptar"
         PicturePosition =   327683
         Size            =   "2461;767"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Provisiones"
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
         Left            =   15
         TabIndex        =   15
         Top             =   7680
         Width           =   9075
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CTS"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   8160
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Gratificación"
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
         Left            =   1680
         TabIndex        =   13
         Top             =   8160
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vacaciones"
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
         Left            =   360
         TabIndex        =   12
         Top             =   8160
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Centro de Costo"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Frame FrmDeducc 
      Height          =   8895
      Left            =   40
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   7575
      Begin Threed.SSPanel SSPanel2 
         Height          =   1935
         Left            =   3780
         TabIndex        =   28
         Top             =   6885
         Width           =   3690
         _Version        =   65536
         _ExtentX        =   6509
         _ExtentY        =   3413
         _StockProps     =   15
         BackColor       =   8421504
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
         Begin VB.TextBox TxtCteD 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2520
            TabIndex        =   66
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox TxtRemuD 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2520
            TabIndex        =   65
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TxtRemuO 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1680
            TabIndex        =   32
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TxtCteO 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1680
            TabIndex        =   31
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox TxtRemuE 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   840
            TabIndex        =   30
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TxtCteE 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   840
            TabIndex        =   29
            Top             =   960
            Width           =   855
         End
         Begin VB.Label LblIdCteD 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   480
            TabIndex        =   77
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdRemuD 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "Dirección"
            Height          =   255
            Left            =   2520
            TabIndex        =   72
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "Obrero"
            Height          =   255
            Left            =   1680
            TabIndex        =   71
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "Empleado"
            Height          =   255
            Left            =   840
            TabIndex        =   70
            Top             =   480
            Width           =   855
         End
         Begin VB.Label LblIdCteO 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   61
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdCteE 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   60
            Top             =   960
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdRemuO 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   59
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdRemuE 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   58
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSForms.CommandButton CommandButton1 
            Height          =   375
            Left            =   855
            TabIndex        =   38
            Top             =   1320
            Width           =   2595
            ForeColor       =   4210752
            BackColor       =   16744576
            Caption         =   "   Aceptar"
            PicturePosition =   327683
            Size            =   "4577;661"
            Picture         =   "FrmSeteoAsiento.frx":0D45
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "PLANILLAS"
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
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   3495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cta Cte"
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
            TabIndex        =   34
            Top             =   1020
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rem"
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
            TabIndex        =   33
            Top             =   720
            Width           =   390
         End
      End
      Begin MSDataGridLib.DataGrid DGrdDedudcc 
         Bindings        =   "FrmSeteoAsiento.frx":12DF
         Height          =   6615
         Left            =   75
         TabIndex        =   17
         Top             =   240
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   11668
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
            DataField       =   "codinterno"
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
         BeginProperty Column02 
            DataField       =   "cuenta"
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
         BeginProperty Column03 
            DataField       =   "id_seteo"
            Caption         =   "id_seteo"
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
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4559.811
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1935
         Left            =   75
         TabIndex        =   18
         Top             =   6885
         Width           =   3690
         _Version        =   65536
         _ExtentX        =   6509
         _ExtentY        =   3413
         _StockProps     =   15
         BackColor       =   8421504
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
         Begin VB.TextBox TxtCtsD 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2400
            TabIndex        =   64
            Top             =   1140
            Width           =   855
         End
         Begin VB.TextBox TxtGratiD 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2400
            TabIndex        =   63
            Top             =   900
            Width           =   855
         End
         Begin VB.TextBox TxtVacaD 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2400
            TabIndex        =   62
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox TxtVacaO 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1560
            TabIndex        =   27
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox TxtGratiO 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1560
            TabIndex        =   26
            Top             =   900
            Width           =   855
         End
         Begin VB.TextBox TxtCtsO 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1560
            TabIndex        =   25
            Top             =   1140
            Width           =   855
         End
         Begin VB.TextBox TxtVacaE 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   720
            TabIndex        =   21
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox TxtGratiE 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   720
            TabIndex        =   20
            Top             =   900
            Width           =   855
         End
         Begin VB.TextBox TxtCtsE 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   720
            TabIndex        =   19
            Top             =   1140
            Width           =   855
         End
         Begin VB.Label LblIdVacaD 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdGratiD 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdCtsD 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "Dirección"
            Height          =   255
            Left            =   2400
            TabIndex        =   69
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "Obrero"
            Height          =   255
            Left            =   1560
            TabIndex        =   68
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "Empleado"
            Height          =   255
            Left            =   720
            TabIndex        =   67
            Top             =   480
            Width           =   855
         End
         Begin VB.Label LblIdCtsO 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   57
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdCtsE 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   56
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdGratiO 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   55
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdGratiE 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdVacaO 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   53
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblIdVacaE 
            BackColor       =   &H8000000E&
            Height          =   255
            Left            =   3360
            TabIndex        =   52
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSForms.CommandButton CommandButton2 
            Height          =   375
            Left            =   720
            TabIndex        =   39
            Top             =   1440
            Width           =   2595
            ForeColor       =   4210752
            BackColor       =   16744576
            Caption         =   "   Aceptar"
            PicturePosition =   327683
            Size            =   "4577;661"
            Picture         =   "FrmSeteoAsiento.frx":12F9
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "PROVISIONES"
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
            Left            =   140
            TabIndex        =   35
            Top             =   90
            Width           =   3500
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vac"
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
            Left            =   240
            TabIndex        =   24
            Top             =   690
            Width           =   345
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grat"
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
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTS"
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
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "FrmSeteoAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VCCosto As String
Dim VConcepto As String
Dim Sql As String
Dim rsremunera As New Recordset
Dim rsDeducc As New Recordset
Dim gIdSeteo As Integer
Dim lBaseDados As String

Private Sub Cmbccosto_Click()
If VConcepto = "02" Then Carga_Remunera
If VConcepto = "04" Then Carga_Aportacion
End Sub

Private Sub Cmbconcepto_Click()
LblCta4.Visible = False: TxtCta4.Visible = False
If CmbConcepto.ListIndex = 0 Then
   FramIng.Visible = True
   FrmDeducc.Visible = False
   VConcepto = "02"
   Carga_Remunera
ElseIf CmbConcepto.ListIndex = 1 Then
   FramIng.Visible = False
   FrmDeducc.Visible = True
   VConcepto = "03"
   Carga_Deducciones
Else
   FramIng.Visible = True
   FrmDeducc.Visible = False
   LblCta4.Visible = True: TxtCta4.Visible = True
   VConcepto = "04"
   Carga_Aportacion
End If
End Sub

Private Sub CmbTrabajador_Click()
If VConcepto = "02" Then Carga_Remunera
If VConcepto = "04" Then Carga_Aportacion
End Sub

Private Sub CommandButton1_Click()
If Graba("05", "01", "", VConcepto, TxtRemuE.Text, "09", LblIdRemuE, 0, "''", "''", "''") Then LblIdRemuE.Caption = gIdSeteo
If Graba("05", "02", "", VConcepto, TxtRemuO.Text, "09", LblIdRemuO, 0, "''", "''", "''") Then LblIdRemuO.Caption = gIdSeteo
If Graba("05", "01", "", VConcepto, TxtRemuD.Text, "09", LblIdRemuD, 1, "''", "''", "''") Then LblIdRemuD.Caption = gIdSeteo
If Graba("06", "01", "", VConcepto, TxtCteE.Text, "07", LblIdCteE, 0, "''", "''", "''") Then LblIdCteE.Caption = gIdSeteo
If Graba("06", "02", "", VConcepto, TxtCteO.Text, "07", LblIdCteO, 0, "''", "''", "''") Then LblIdCteO.Caption = gIdSeteo
If Graba("06", "01", "", VConcepto, TxtCteD.Text, "07", LblIdCteD, 1, "''", "''", "''") Then LblIdCteD.Caption = gIdSeteo
End Sub

Private Sub CommandButton2_Click()
If Graba("02", "01", "", VConcepto, TxtVacaE.Text, "", LblIdVacaE, 0, "''", "''", "''") Then LblIdVacaE.Caption = gIdSeteo
If Graba("02", "02", "", VConcepto, TxtVacaO.Text, "", LblIdVacaO, 0, "''", "''", "''") Then LblIdVacaO.Caption = gIdSeteo
If Graba("02", "01", "", VConcepto, TxtVacaD.Text, "", LblIdVacaD, 1, "''", "''", "''") Then LblIdVacaD.Caption = gIdSeteo
If Graba("03", "01", "", VConcepto, TxtGratiE.Text, "", LblIdGratiE, 0, "''", "''", "''") Then LblIdGratiE.Caption = gIdSeteo
If Graba("03", "02", "", VConcepto, TxtGratiO.Text, "", LblIdGratiO, 0, "''", "''", "''") Then LblIdGratiO.Caption = gIdSeteo
If Graba("03", "01", "", VConcepto, TxtGratiD.Text, "", LblIdGratiD, 1, "''", "''", "''") Then LblIdGratiD.Caption = gIdSeteo
If Graba("04", "01", "", VConcepto, TxtCtsE.Text, "", LblIdCtsE, 0, "''", "''", "''") Then LblIdCtsE.Caption = gIdSeteo
If Graba("04", "02", "", VConcepto, TxtCtsO.Text, "", LblIdCtsO, 0, "''", "''", "''") Then LblIdCtsO.Caption = gIdSeteo
If Graba("04", "01", "", VConcepto, TxtCtsD.Text, "", LblIdCtsD, 1, "''", "''", "''") Then LblIdCtsD.Caption = gIdSeteo
End Sub

Private Sub CommandButton3_Click()
Dim mCta4 As String
mCta4 = ""
If TxtCta4.Visible = True Then mCta4 = TxtCta4.Text
If Graba("01", VTipo, VCCosto, VConcepto, TxtCta.Text, LblCod.Caption, LblIdSeteo, 0, mCta4, TxtCtav.Text, TxtCtag) Then
   LblIdSeteo.Caption = gIdSeteo
   PnlModi.Visible = False
   If VConcepto = "02" Or VConcepto = "04" Then
      DgrdRemunera.Enabled = True
      DgrdRemunera.Columns(2) = TxtCta.Text
      DgrdRemunera.Columns(3) = LblIdSeteo.Caption
      DgrdRemunera.Columns(4) = TxtCtav.Text
      DgrdRemunera.Columns(5) = TxtCtag.Text
   Else
      DGrdDedudcc.Enabled = True
      DGrdDedudcc.Columns(2) = TxtCta.Text
      DGrdDedudcc.Columns(3) = LblIdSeteo.Caption
   End If
End If
End Sub

Private Sub DGrdDedudcc_DblClick()
Acepta_Deduccion
End Sub
Private Sub DGrdRemunera_DblClick()
TxtCta.Text = "": TxtCtav.Text = "": TxtCtag.Text = "": TxtCta4.Text = ""
Acepta_Remunera
End Sub

Private Sub DgrdRemunera_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Acepta_Remunera
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 9480: Me.Height = 9975
Call fc_Descrip_Maestros2("01055", "", CmbTrabajador)
'Call fc_Descrip_Maestros2("01044", "", Cmbccosto)

Sql$ = "select codigo,descripcion From Pla_ccostos where status<>'*' order by descripcion"
Call rCarCbo(CmbCcosto, Sql$, "C", "00")

Crea_Rs

lBaseDados = ""
Sql = "select nombrebd from cia where cod_cia='" & wcia & "'"
If fAbrRst(Rs, Sql) Then lBaseDados = Trim(Rs(0) & "")
Rs.Close: Set Rs = Nothing
'Campo TipoSeteo
'01=BOLETA
'02=PROV VAC
'03=PROV GRATI
'04=PROV CTS
'05=CTAS REMUNERACION
'06=CTA CTE
End Sub
Private Sub Carga_Remunera()
If rsremunera.RecordCount > 0 Then rsremunera.MoveFirst
Do While Not rsremunera.EOF
   rsremunera.Delete
   rsremunera.MoveNext
Loop
DgrdRemunera.Refresh

VTipo = fc_CodigoComboBox(CmbTrabajador, 2)
VCCosto = fc_CodigoComboBox(CmbCcosto, 2)

If Trim(VTipo & "") = "" Then Exit Sub
If Trim(VCCosto & "") = "" Then Exit Sub
If Trim(VConcepto & "") = "" Then Exit Sub
TxtVaca.Text = ""
TxtGrati.Text = ""
TxtCts.Text = ""

LblIdSeteoIV.Caption = "0"
LblIdSeteoIG.Caption = "0"
LblIdSeteoIC.Caption = "0"

Sql = "usp_Seteo_Asiento '" & wcia & "','01','" & VTipo & "','" & VCCosto & "','" & VConcepto & "'"
If fAbrRst(Rs, Sql) Then Rs.MoveFirst
Do While Not Rs.EOF
   If Trim(Rs!TipoSeteo & "") = "01" Or Trim(Rs!TipoSeteo & "") = "" Then
      rsremunera.AddNew
      rsremunera!id_seteo = 0
      rsremunera!codinterno = Trim(Rs!codinterno & "")
      rsremunera!Descripcion = Trim(Rs!Descripcion & "")
      rsremunera!Cuenta = Trim(Rs!Cuenta & "")
      rsremunera!Cuentav = Trim(Rs!Cuentav & "")
      rsremunera!Cuentag = Trim(Rs!Cuentag & "")
      rsremunera!Cuenta2 = Trim(Rs!Cuenta2 & "")
      
      If Trim(Rs!id_seteo & "") <> "" Then rsremunera!id_seteo = Rs!id_seteo
   ElseIf Trim(Rs!TipoSeteo & "") = "02" Then
      If Trim(Rs!id_seteo & "") <> "" Then LblIdSeteoIV = Rs!id_seteo
      TxtVaca.Text = Trim(Rs!Cuenta & "")
   ElseIf Trim(Rs!TipoSeteo & "") = "03" Then
      If Trim(Rs!id_seteo & "") <> "" Then LblIdSeteoIG = Rs!id_seteo
      TxtGrati.Text = Trim(Rs!Cuenta & "")
   ElseIf Trim(Rs!TipoSeteo & "") = "04" Then
      If Trim(Rs!id_seteo & "") <> "" Then LblIdSeteoIC = Rs!id_seteo
      TxtCts.Text = Trim(Rs!Cuenta & "")
   End If
   Rs.MoveNext
Loop
DgrdRemunera.Refresh
Rs.Close: Set Rs = Nothing

Screen.MousePointer = vbDefault
End Sub

Private Sub SSCommand3_Click()
If Graba("02", VTipo, VCCosto, VConcepto, TxtVaca.Text, "", LblIdSeteoIV, 0, "''", "''", "''") Then LblIdSeteoIV = gIdSeteo
If Graba("03", VTipo, VCCosto, VConcepto, TxtGrati.Text, "", LblIdSeteoIG, 0, "''", "''", "''") Then LblIdSeteoIG = gIdSeteo
If Graba("04", VTipo, VCCosto, VConcepto, TxtCts.Text, "", LblIdSeteoIC, 0, "''", "''", "''") Then LblIdSeteoIC = gIdSeteo
End Sub

Private Sub SSCommand4_Click()
PnlModi.Visible = False
DGrdDedudcc.Enabled = True
DgrdRemunera.Enabled = True
End Sub
Private Function Valida_Cuenta(cta As String) As Boolean
Dim Rq As ADODB.Recordset
Valida_Cuenta = False
If Len(Trim(cta)) <> 7 And Trim(cta & "") <> "" Then
   MsgBox "Ingrese Cuenta Correctamente", vbInformation
   Exit Function
End If
If Trim(cta & "") <> "" Then
   If wGrupoPla = "01" And wcia = "21" Then
      Sql = "select cgcod from conmaspcge21 where cod_cia='" & wcia & "' and cgcod='" & cta & "'"
   Else
      'Sql = "select cgcod from ." & lBaseDados & "..conmaspcge where cod_cia='" & wcia & "' and cgcod='" & cta & "'"
      Sql = "select cgcod from conmaspcge where cod_cia='" & wcia & "' and cgcod='" & cta & "'"
   End If
   If Not fAbrRst(Rq, Sql) Then
      MsgBox "Cuenta no Registrada => " & cta, vbInformation
      Rq.Close: Set Rq = Nothing
      Exit Function
   End If
Rq.Close: Set Rq = Nothing
End If
Valida_Cuenta = True
End Function
Private Sub Crea_Rs()
    If rsremunera.State = 1 Then rsremunera.Close
    rsremunera.Fields.Append "codinterno", adChar, 2, adFldIsNullable
    rsremunera.Fields.Append "descripcion", adVarChar, 100, adFldIsNullable
    rsremunera.Fields.Append "cuenta", adChar, 7, adFldIsNullable
    rsremunera.Fields.Append "cuentav", adChar, 7, adFldIsNullable
    rsremunera.Fields.Append "cuentag", adChar, 7, adFldIsNullable
    rsremunera.Fields.Append "id_seteo", adInteger, 4, adFldIsNullable
    rsremunera.Fields.Append "cuenta2", adChar, 7, adFldIsNullable
    rsremunera.Open
    Set DgrdRemunera.DataSource = rsremunera
    
    If rsDeducc.State = 1 Then rsDeducc.Close
    rsDeducc.Fields.Append "codinterno", adChar, 2, adFldIsNullable
    rsDeducc.Fields.Append "descripcion", adVarChar, 100, adFldIsNullable
    rsDeducc.Fields.Append "cuenta", adChar, 7, adFldIsNullable
    rsDeducc.Fields.Append "id_seteo", adInteger, 4, adFldIsNullable
    rsDeducc.Open
    Set DGrdDedudcc.DataSource = rsDeducc
End Sub
Private Sub Acepta_Remunera()
LblIdSeteo.Caption = "0"
LblCod.Caption = ""
LblConcepto.Caption = ""
TxtCta.Text = ""
TxtCta4.Text = ""
LblTipoConcepto.Caption = ""

If DgrdRemunera.Row < 0 Then Exit Sub
LblTit.Caption = CmbConcepto.Text
LblCod.Caption = (Trim(DgrdRemunera.Columns(0)))
LblConcepto.Caption = (Trim(DgrdRemunera.Columns(1)))
TxtCta.Text = (Trim(DgrdRemunera.Columns(2)) & "")
TxtCta4.Text = (Trim(DgrdRemunera.Columns(6)) & "")

TxtCtav.Text = (Trim(DgrdRemunera.Columns(4)) & "")
TxtCtag.Text = (Trim(DgrdRemunera.Columns(5)) & "")
LblTipoConcepto.Caption = VConcepto
If Trim(DgrdRemunera.Columns(3) & "") <> "" Then LblIdSeteo.Caption = DgrdRemunera.Columns(3)
PnlModi.Visible = True
DgrdRemunera.Enabled = False
TxtCta.SetFocus
End Sub
Private Function Graba(TipoSet As String, VTipotrab As String, VCCosto As String, VConcepto As String, cta As String, lCodConcepto As String, IdSeteo As Integer, Direccion As Integer, mCta4 As String, Ctav As String, Ctag As String) As Boolean
Graba = False
Dim lC As String
lC = VConcepto
If VConcepto = "04" Then lC = "03"
If Trim(cta & "") <> "" Then
   If Valida_Cuenta(cta) Then
      Sql = "usp_inserta_SeteoAsiento '" & wcia & "','" & TipoSet & "','" & VTipotrab & "','" & VCCosto & "','" & lC & "','" & lCodConcepto & "','" & cta & "','" & wuser & "'," & IdSeteo & "," & Direccion & ",'" & mCta4 & "','" & Ctav & "','" & Ctag & "'"
      If fAbrRst(Rs, Sql) Then
         gIdSeteo = Rs(0): Graba = True
      End If
      Rs.Close: Set Rs = Nothing
   End If
End If
End Function
Private Sub Carga_Deducciones()
If rsDeducc.RecordCount > 0 Then rsDeducc.MoveFirst
Do While Not rsDeducc.EOF
   rsDeducc.Delete
   rsDeducc.MoveNext
Loop
DGrdDedudcc.Refresh

If Trim(VConcepto & "") = "" Then Exit Sub

VTipo = ""
VCCosto = ""
TxtVacaE.Text = ""
TxtVacaO.Text = ""
TxtVacaD.Text = ""
TxtGratiE.Text = ""
TxtGratiO.Text = ""
TxtGratiD.Text = ""
TxtCtsE.Text = ""
TxtCtsO.Text = ""
TxtCtsD.Text = ""
TxtRemuE.Text = ""
TxtRemuO.Text = ""
TxtRemuD.Text = ""
TxtCteE.Text = ""
TxtCteO.Text = ""
TxtCteD.Text = ""

LblIdVacaE.Caption = "0"
LblIdVacaO.Caption = "0"
LblIdVacaD.Caption = "0"
LblIdGratiE.Caption = "0"
LblIdGratiO.Caption = "0"
LblIdGratiD.Caption = "0"
LblIdCtsE.Caption = "0"
LblIdCtsO.Caption = "0"
LblIdCtsD.Caption = "0"
LblIdRemuE.Caption = "0"
LblIdRemuO.Caption = "0"
LblIdRemuD.Caption = "0"
LblIdCteE.Caption = "0"
LblIdCteO.Caption = "0"
LblIdCteD.Caption = "0"

Sql = "usp_Seteo_Asiento '" & wcia & "','01','','','" & VConcepto & "'"
If fAbrRst(Rs, Sql) Then Rs.MoveFirst
Do While Not Rs.EOF
   If Rs!aportacion = 0 Then
    If Trim(Rs!TipoSeteo & "") = "01" Or Trim(Rs!TipoSeteo & "") = "" Then
       rsDeducc.AddNew
       rsDeducc(3) = 0
       rsDeducc(0) = Trim(Rs!codinterno & "")
       rsDeducc(1) = Trim(Rs!Descripcion & "")
       rsDeducc(2) = Trim(Rs!Cuenta & "")
       If Trim(Rs!id_seteo & "") <> "" Then rsDeducc(3) = Rs!id_seteo
    ElseIf Trim(Rs!TipoSeteo & "") = "02" Then
       If Rs!Direccion = "1" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdVacaD = Rs!id_seteo
          TxtVacaD.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "01" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdVacaE = Rs!id_seteo
          TxtVacaE.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "02" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdVacaO = Rs!id_seteo
          TxtVacaO.Text = Trim(Rs!Cuenta & "")
       End If
    ElseIf Trim(Rs!TipoSeteo & "") = "03" Then
       If Rs!Direccion = "1" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdGratiD = Rs!id_seteo
          TxtGratiD.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "01" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdGratiE = Rs!id_seteo
          TxtGratiE.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "02" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdGratiO = Rs!id_seteo
          TxtGratiO.Text = Trim(Rs!Cuenta & "")
       End If
    ElseIf Trim(Rs!TipoSeteo & "") = "04" Then
       If Rs!Direccion = "1" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdCtsD = Rs!id_seteo
          TxtCtsD.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "01" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdCtsE = Rs!id_seteo
          TxtCtsE.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "02" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdCtsO = Rs!id_seteo
          TxtCtsO.Text = Trim(Rs!Cuenta & "")
       End If
    ElseIf Trim(Rs!TipoSeteo & "") = "05" Then
       If Rs!Direccion = "1" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdRemuD = Rs!id_seteo
          TxtRemuD.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "01" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdRemuE = Rs!id_seteo
          TxtRemuE.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "02" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdRemuO = Rs!id_seteo
          TxtRemuO.Text = Trim(Rs!Cuenta & "")
       End If
    ElseIf Trim(Rs!TipoSeteo & "") = "06" Then
       If Rs!Direccion = "1" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdCteD = Rs!id_seteo
          TxtCteD.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "01" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdCteE = Rs!id_seteo
          TxtCteE.Text = Trim(Rs!Cuenta & "")
       ElseIf Trim(Rs!TipoTrab & "") = "02" Then
          If Trim(Rs!id_seteo & "") <> "" Then LblIdCteO = Rs!id_seteo
          TxtCteO.Text = Trim(Rs!Cuenta & "")
       End If
    End If
   End If
   Rs.MoveNext
Loop
DGrdDedudcc.Refresh
Rs.Close: Set Rs = Nothing

Screen.MousePointer = vbDefault
End Sub
Private Sub Acepta_Deduccion()
LblIdSeteo.Caption = "0"
LblCod.Caption = ""
LblConcepto.Caption = ""
TxtCta.Text = ""
LblTipoConcepto.Caption = ""

If DGrdDedudcc.Row < 0 Then Exit Sub
LblTit.Caption = CmbConcepto.Text
LblCod.Caption = (Trim(DGrdDedudcc.Columns(0)))
LblConcepto.Caption = (Trim(DGrdDedudcc.Columns(1)))
TxtCta.Text = (Trim(DGrdDedudcc.Columns(2)) & "")
LblTipoConcepto.Caption = VConcepto
LblIdSeteo.Caption = DGrdDedudcc.Columns(3)
PnlModi.Visible = True
DGrdDedudcc.Enabled = False
TxtCta.SetFocus
End Sub
Private Sub Carga_Aportacion()
If rsremunera.RecordCount > 0 Then rsremunera.MoveFirst
Do While Not rsremunera.EOF
   rsremunera.Delete
   rsremunera.MoveNext
Loop
DgrdRemunera.Refresh

VTipo = fc_CodigoComboBox(CmbTrabajador, 2)
VCCosto = fc_CodigoComboBox(CmbCcosto, 2)

If Trim(VTipo & "") = "" Then Exit Sub
If Trim(VCCosto & "") = "" Then Exit Sub
If Trim(VConcepto & "") = "" Then Exit Sub
TxtVaca.Text = ""
TxtGrati.Text = ""
TxtCts.Text = ""

LblIdSeteoIV.Caption = "0"
LblIdSeteoIG.Caption = "0"
LblIdSeteoIC.Caption = "0"

Sql = "usp_Seteo_Asiento '" & wcia & "','01','" & VTipo & "','" & VCCosto & "','03'"
If fAbrRst(Rs, Sql) Then Rs.MoveFirst
Do While Not Rs.EOF
   If Trim(Rs!Cuenta & "") <> "" Then
   If Trim(Rs!TipoSeteo & "") = "01" Or Trim(Rs!TipoSeteo & "") = "" Then
      rsremunera.AddNew
      rsremunera!id_seteo = 0
      rsremunera!codinterno = Trim(Rs!codinterno & "")
      rsremunera!Descripcion = Trim(Rs!Descripcion & "")
      rsremunera!Cuenta = Trim(Rs!Cuenta & "")
      rsremunera!Cuentav = Trim(Rs!Cuentav & "")
      rsremunera!Cuentag = Trim(Rs!Cuentag & "")
      rsremunera!Cuenta2 = Trim(Rs!Cuenta2 & "")
      
      
      If Trim(Rs!id_seteo & "") <> "" Then rsremunera!id_seteo = Rs!id_seteo
   ElseIf Trim(Rs!TipoSeteo & "") = "02" Then
      If Trim(Rs!id_seteo & "") <> "" Then LblIdSeteoIV = Rs!id_seteo
      TxtVaca.Text = Trim(Rs!Cuenta & "")
   ElseIf Trim(Rs!TipoSeteo & "") = "03" Then
      If Trim(Rs!id_seteo & "") <> "" Then LblIdSeteoIG = Rs!id_seteo
      TxtGrati.Text = Trim(Rs!Cuenta & "")
   ElseIf Trim(Rs!TipoSeteo & "") = "04" Then
      If Trim(Rs!id_seteo & "") <> "" Then LblIdSeteoIC = Rs!id_seteo
      TxtCts.Text = Trim(Rs!Cuenta & "")
   End If
   End If
   Rs.MoveNext
Loop
DgrdRemunera.Refresh
Rs.Close: Set Rs = Nothing

Screen.MousePointer = vbDefault
End Sub

