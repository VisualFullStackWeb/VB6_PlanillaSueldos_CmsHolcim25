VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{92F3EE28-0DF4-4CB0-AAA5-E6578B12929B}#1.0#0"; "CboFacil.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form aFrmDH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Maestro Derechohabientes «"
   ClientHeight    =   6870
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   11520
   Icon            =   "aFrmDH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tabdh 
      Height          =   5235
      Left            =   60
      TabIndex        =   49
      Top             =   1560
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9234
      _Version        =   393216
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Datos Derechohabientes"
      TabPicture(0)   =   "aFrmDH.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraDir1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dirección 1"
      TabPicture(1)   =   "aFrmDH.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame25"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Dirección 2"
      TabPicture(2)   =   "aFrmDH.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Domicilio del Derecho Habiente - Dirección 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4590
         Left            =   -74835
         TabIndex        =   92
         Top             =   480
         Width           =   11085
         Begin VB.CheckBox ChkInd2 
            Caption         =   "Indicador centro asistencial de salud"
            Height          =   495
            Left            =   375
            TabIndex        =   106
            Top             =   3015
            Width           =   3375
         End
         Begin VB.TextBox TxtRef2 
            Height          =   315
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   105
            Top             =   2640
            Width           =   9465
         End
         Begin VB.TextBox TxtNomUbicacion2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   2160
            Width           =   9075
         End
         Begin VB.TextBox TxtNroDpto2 
            Height          =   285
            Left            =   5070
            TabIndex        =   103
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNroInt2 
            Height          =   285
            Left            =   6750
            TabIndex        =   102
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNroMz2 
            Height          =   285
            Left            =   8430
            TabIndex        =   101
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNroLote2 
            Height          =   285
            Left            =   10110
            TabIndex        =   100
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNroKM2 
            Height          =   285
            Left            =   5070
            TabIndex        =   99
            Top             =   1230
            Width           =   700
         End
         Begin VB.TextBox TxtNroBlock2 
            Height          =   285
            Left            =   6750
            TabIndex        =   98
            Top             =   1230
            Width           =   700
         End
         Begin VB.TextBox TXtNroEtapa2 
            Height          =   285
            Left            =   8430
            TabIndex        =   97
            Top             =   1230
            Width           =   700
         End
         Begin VB.TextBox TxtNomZona2 
            Height          =   285
            Left            =   3825
            MaxLength       =   30
            TabIndex        =   96
            Top             =   1695
            Width           =   6960
         End
         Begin VB.TextBox TxtNomVia2 
            Height          =   330
            Left            =   3825
            MaxLength       =   30
            TabIndex        =   95
            Top             =   420
            Width           =   5325
         End
         Begin VB.CommandButton CmdUbigeo2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1305
            Picture         =   "aFrmDH.frx":035E
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   2160
            Width           =   360
         End
         Begin VB.TextBox TxtNroVia2 
            Height          =   285
            Left            =   10110
            TabIndex        =   93
            Top             =   420
            Width           =   700
         End
         Begin CboFacil.cbo_facil cbo_via2 
            Height          =   315
            Left            =   1305
            TabIndex        =   107
            Top             =   420
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
         Begin CboFacil.cbo_facil cbo_zona2 
            Height          =   315
            Left            =   1305
            TabIndex        =   108
            Top             =   1695
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
         Begin VB.Label Label64 
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
            Index           =   5
            Left            =   390
            TabIndex        =   122
            Top             =   2655
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación"
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
            Left            =   390
            TabIndex        =   121
            Top             =   2115
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interior"
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
            Left            =   6090
            TabIndex        =   120
            Top             =   870
            Width           =   540
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
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
            Index           =   9
            Left            =   3870
            TabIndex        =   119
            Top             =   870
            Width           =   1035
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
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
            Index           =   8
            Left            =   9495
            TabIndex        =   118
            Top             =   870
            Width           =   315
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilometro"
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
            Index           =   7
            Left            =   3885
            TabIndex        =   117
            Top             =   1185
            Width           =   660
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Block"
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
            Index           =   6
            Left            =   6165
            TabIndex        =   116
            Top             =   1230
            Width           =   360
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manzana"
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
            Left            =   7755
            TabIndex        =   115
            Top             =   870
            Width           =   645
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
            Index           =   3
            Left            =   7755
            TabIndex        =   114
            Top             =   1230
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Zona"
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
            Left            =   390
            TabIndex        =   113
            Top             =   1695
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zona"
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
            Left            =   3420
            TabIndex        =   112
            Top             =   1695
            Width           =   360
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Via"
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
            Left            =   390
            TabIndex        =   111
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Via"
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
            Left            =   3420
            TabIndex        =   110
            Top             =   420
            Width           =   210
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Via"
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
            Left            =   9480
            TabIndex        =   109
            Top             =   420
            Width           =   510
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Domicilio del Derecho Habiente - Dirección 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4590
         Left            =   -74835
         TabIndex        =   74
         Top             =   480
         Width           =   11085
         Begin VB.TextBox TxtNroVia1 
            Height          =   285
            Left            =   10110
            TabIndex        =   31
            Top             =   420
            Width           =   700
         End
         Begin VB.CommandButton CmdUbigeo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1305
            Picture         =   "aFrmDH.frx":08E8
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   2160
            Width           =   360
         End
         Begin VB.TextBox TxtNomVia1 
            Height          =   330
            Left            =   3825
            MaxLength       =   30
            TabIndex        =   30
            Top             =   420
            Width           =   5325
         End
         Begin VB.TextBox TxtNomZona1 
            Height          =   285
            Left            =   3825
            MaxLength       =   30
            TabIndex        =   28
            Top             =   1695
            Width           =   6960
         End
         Begin VB.TextBox TXtNroEtapa1 
            Height          =   285
            Left            =   8430
            TabIndex        =   38
            Top             =   1230
            Width           =   700
         End
         Begin VB.TextBox TxtNroBlock1 
            Height          =   285
            Left            =   6750
            TabIndex        =   37
            Top             =   1230
            Width           =   700
         End
         Begin VB.TextBox TxtNroKM1 
            Height          =   285
            Left            =   5070
            TabIndex        =   36
            Top             =   1230
            Width           =   700
         End
         Begin VB.TextBox TxtNroLote1 
            Height          =   285
            Left            =   10110
            TabIndex        =   35
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNroMz1 
            Height          =   285
            Left            =   8430
            TabIndex        =   34
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNroInt1 
            Height          =   285
            Left            =   6750
            TabIndex        =   33
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNroDpto1 
            Height          =   285
            Left            =   5070
            TabIndex        =   32
            Top             =   870
            Width           =   700
         End
         Begin VB.TextBox TxtNomUbicacion1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   2160
            Width           =   9075
         End
         Begin VB.TextBox TxtRef1 
            Height          =   315
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   41
            Top             =   2640
            Width           =   9465
         End
         Begin VB.CheckBox ChkInd1 
            Caption         =   "Indicador centro asistencial de salud"
            Height          =   495
            Left            =   375
            TabIndex        =   42
            Top             =   3015
            Width           =   3375
         End
         Begin CboFacil.cbo_facil cbo_via1 
            Height          =   315
            Left            =   1305
            TabIndex        =   29
            Top             =   420
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
         Begin CboFacil.cbo_facil cbo_zona1 
            Height          =   315
            Left            =   1305
            TabIndex        =   27
            Top             =   1695
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
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Via"
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
            Left            =   9480
            TabIndex        =   90
            Top             =   420
            Width           =   510
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Via"
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
            Left            =   3420
            TabIndex        =   87
            Top             =   420
            Width           =   210
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Via"
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
            Left            =   390
            TabIndex        =   86
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zona"
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
            Left            =   3420
            TabIndex        =   85
            Top             =   1695
            Width           =   360
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Zona"
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
            Left            =   390
            TabIndex        =   84
            Top             =   1695
            Width           =   705
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
            Left            =   7755
            TabIndex        =   83
            Top             =   1230
            Width           =   420
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Manzana"
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
            Left            =   7755
            TabIndex        =   82
            Top             =   870
            Width           =   645
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Block"
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
            Left            =   6165
            TabIndex        =   81
            Top             =   1230
            Width           =   360
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kilometro"
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
            Left            =   3885
            TabIndex        =   80
            Top             =   1185
            Width           =   660
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lote"
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
            Left            =   9495
            TabIndex        =   79
            Top             =   870
            Width           =   315
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento"
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
            Left            =   3870
            TabIndex        =   78
            Top             =   870
            Width           =   1035
         End
         Begin VB.Label TxtNroInt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interior"
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
            Left            =   6090
            TabIndex        =   77
            Top             =   870
            Width           =   540
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubicación"
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
            Left            =   390
            TabIndex        =   76
            Top             =   2115
            Width           =   675
         End
         Begin VB.Label Label64 
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
            Index           =   1
            Left            =   390
            TabIndex        =   75
            Top             =   2655
            Width           =   780
         End
      End
      Begin VB.Frame FraDir1 
         Height          =   4725
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   11175
         Begin VB.Frame FramederchoH 
            Caption         =   " Información del Derecho Habiente "
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
            Height          =   3450
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   10935
            Begin VB.CheckBox Chkestudio 
               Caption         =   "Estudios Escolares"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   6840
               TabIndex        =   15
               Top             =   1680
               Width           =   1875
            End
            Begin VB.TextBox TxtEmail 
               Height          =   315
               Left            =   6840
               TabIndex        =   14
               Top             =   1320
               Width           =   3975
            End
            Begin VB.Frame Fratele 
               Caption         =   "Teléfono"
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
               Height          =   1275
               Left            =   120
               TabIndex        =   70
               Top             =   2040
               Width           =   5655
               Begin VB.TextBox TxtTelef 
                  Height          =   285
                  Left            =   1920
                  TabIndex        =   17
                  Top             =   600
                  Width           =   2895
               End
               Begin VB.ComboBox CboCodCiudad 
                  Height          =   315
                  Left            =   1920
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   240
                  Width           =   2895
               End
               Begin VB.Label Label59 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Número:"
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
                  TabIndex        =   72
                  Top             =   600
                  Width           =   615
               End
               Begin VB.Label Label59 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código de Ciudad:"
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
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1320
               End
            End
            Begin VB.ComboBox CboPais 
               Height          =   315
               Left            =   8160
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   840
               Width           =   2655
            End
            Begin VB.Frame Frame26 
               Caption         =   "Situacion"
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
               Height          =   1275
               Left            =   5940
               TabIndex        =   59
               Top             =   2040
               Width           =   4830
               Begin VB.OptionButton OpcActivoDh 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Activo"
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
                  TabIndex        =   18
                  Top             =   270
                  Width           =   810
               End
               Begin VB.OptionButton OpcBajaDh 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Baja"
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
                  Left            =   2730
                  TabIndex        =   20
                  Top             =   240
                  Width           =   675
               End
               Begin CboFacil.cbo_facil cbobajadh 
                  Height          =   315
                  Left            =   195
                  TabIndex        =   22
                  Top             =   840
                  Width           =   4500
                  _ExtentX        =   7938
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
               Begin MSMask.MaskEdBox FecAltaDH 
                  Height          =   315
                  Index           =   0
                  Left            =   1155
                  TabIndex        =   19
                  Top             =   240
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox FecBajaDH 
                  Height          =   315
                  Index           =   1
                  Left            =   3495
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   556
                  _Version        =   393216
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label29 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Motivo de Baja"
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
                  TabIndex        =   60
                  Top             =   615
                  Width           =   1065
               End
            End
            Begin VB.Frame Frame24 
               Caption         =   "Incapacidad"
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
               Height          =   735
               Left            =   1575
               TabIndex        =   57
               Top             =   1215
               Width           =   4200
               Begin VB.OptionButton OpcIncapazSi 
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
                  Height          =   195
                  Left            =   195
                  TabIndex        =   11
                  Top             =   480
                  Width           =   705
               End
               Begin VB.OptionButton OpcIncapazNo 
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
                  Height          =   195
                  Left            =   195
                  TabIndex        =   13
                  Top             =   240
                  Width           =   510
               End
               Begin VB.TextBox TxtIncapaz 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   990
                  TabIndex        =   12
                  Top             =   380
                  Width           =   2670
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nro Certificado"
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
                  Left            =   990
                  TabIndex        =   58
                  Top             =   180
                  Width           =   1080
               End
            End
            Begin VB.Frame Frame23 
               Caption         =   "Sexo"
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
               Height          =   720
               Left            =   105
               TabIndex        =   56
               Top             =   1215
               Width           =   1395
               Begin VB.OptionButton OpcVaronDh 
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
                  Height          =   255
                  Left            =   90
                  TabIndex        =   9
                  Top             =   180
                  Width           =   1185
               End
               Begin VB.OptionButton OpcDamaDh 
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
                  Left            =   90
                  TabIndex        =   10
                  Top             =   450
                  Width           =   1185
               End
            End
            Begin VB.ComboBox CmbDocDh 
               Height          =   315
               ItemData        =   "aFrmDH.frx":0E72
               Left            =   1080
               List            =   "aFrmDH.frx":0E74
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   840
               Width           =   2685
            End
            Begin VB.TextBox TxtNroDocDh 
               Height          =   285
               Left            =   4800
               MaxLength       =   15
               TabIndex        =   7
               Top             =   855
               Width           =   1395
            End
            Begin VB.TextBox TxtSegNomDh 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6870
               TabIndex        =   4
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox TxtPriNomDh 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4620
               TabIndex        =   3
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox TxtApeMatDh 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2355
               TabIndex        =   2
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox TxtApePatDh 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   1
               Top             =   480
               Width           =   2175
            End
            Begin MSMask.MaskEdBox TxtFecNacDh 
               Height          =   315
               Left            =   9300
               TabIndex        =   5
               Top             =   450
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email :"
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
               Left            =   6120
               TabIndex        =   73
               Top             =   1320
               Width           =   465
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pais emisor del doc."
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
               Left            =   6360
               TabIndex        =   68
               Top             =   900
               Width           =   1410
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fec. Nac (dd/mm/yyyy) :"
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
               Left            =   8685
               TabIndex        =   67
               Top             =   225
               Width           =   1800
            End
            Begin VB.Label Label60 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "N° Doc. :"
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
               Left            =   3960
               TabIndex        =   66
               Top             =   900
               Width           =   660
            End
            Begin VB.Label Label59 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Doc. :"
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
               Left            =   180
               TabIndex        =   65
               Top             =   900
               Width           =   780
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Segundo Nombre"
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
               Left            =   6915
               TabIndex        =   64
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Primer Nombre"
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
               Left            =   4620
               TabIndex        =   63
               Top             =   240
               Width           =   1050
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ap_Materno"
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
               Left            =   2400
               TabIndex        =   62
               Top             =   240
               Width           =   885
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ap_Paterno"
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
               TabIndex        =   61
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Vínculo Familiar"
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
            Height          =   840
            Left            =   120
            TabIndex        =   51
            Top             =   3750
            Width           =   10935
            Begin MSMask.MaskEdBox TxtConcepcion 
               Height          =   315
               Left            =   9720
               TabIndex        =   26
               Top             =   435
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   7
               Mask            =   "##/####"
               PromptChar      =   "_"
            End
            Begin VB.ComboBox CmbTipDocAcreditaPaternidad 
               Height          =   315
               Left            =   2760
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   435
               Width           =   4455
            End
            Begin VB.TextBox TxtNroDocAcreditapaternidad 
               Height          =   285
               Left            =   7320
               TabIndex        =   25
               Top             =   450
               Width           =   1455
            End
            Begin VB.ComboBox CmbVinculo 
               Height          =   315
               ItemData        =   "aFrmDH.frx":0E76
               Left            =   75
               List            =   "aFrmDH.frx":0E78
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   435
               Width           =   2655
            End
            Begin VB.Label Label55 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Mes/Año de Concepción"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   675
               Index           =   1
               Left            =   9120
               TabIndex        =   69
               Top             =   180
               Width           =   1845
            End
            Begin VB.Label Label55 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Nro. Doc. Acredita vínculo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   0
               Left            =   6960
               TabIndex        =   54
               Top             =   180
               Width           =   1875
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Doc. Acredita  vínculo"
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
               Left            =   2745
               TabIndex        =   53
               Top             =   180
               Width           =   1905
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Vinculo"
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
               TabIndex        =   52
               Top             =   180
               Width           =   1065
            End
         End
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "Trabajador"
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
      Height          =   825
      Left            =   120
      TabIndex        =   47
      Top             =   630
      Width           =   11340
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00EBFEFC&
         Caption         =   "Agregar Nuevo Derechohabiente"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   8745
         Picture         =   "aFrmDH.frx":0E7A
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   135
         Width           =   2565
      End
      Begin VB.CommandButton CmdDH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Picture         =   "aFrmDH.frx":1404
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox TxtCodTrab 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   330
         Width           =   855
      End
      Begin VB.Label LblIdInternodh 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   8040
         TabIndex        =   91
         Top             =   330
         Width           =   615
      End
      Begin VB.Label LblNomTrab 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1680
         TabIndex        =   48
         Top             =   330
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   11655
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   120
         Width           =   7830
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
         Left            =   240
         TabIndex        =   46
         Top             =   120
         Width           =   825
      End
      Begin VB.Label LblFecha 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   11295
         TabIndex        =   45
         Top             =   120
         Width           =   45
      End
   End
End
Attribute VB_Name = "aFrmDH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SwNuevo As Boolean
Dim sCondInternoTrab As String

Dim Sql As String

Private Sub CboCodCiudad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(CboCodCiudad.ListIndex > -1) <> "" Then SendKeys "{tab}"
If KeyCode = vbKeyDelete Then CboCodCiudad.ListIndex = -1
End Sub

Private Sub CboPais_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(CboPais.ListIndex > -1) <> "" Then SendKeys "{tab}"
End Sub

Private Sub ChkInd1_Click()
'If ((Trim(TxtNomZona1.Text) = "" And Trim(TxtNomVia1.Text) = "")) Then
'    MsgBox "Ingrese Nombre de zona o via antes de activar el indicador del centro asistencial de salud (dirección 1)", vbExclamation, Me.Caption
'    Tabdh.Tab = 1
'    ChkInd1.Value = False
'    TxtNomZona1.SetFocus
'End If
End Sub

Private Sub ChkInd2_Click()
'If ((Trim(TxtNomZona2.Text) = "" And Trim(TxtNomVia2.Text) = "")) Then
'    MsgBox "Ingrese Nombre de zona o via antes de activar el indicador del centro asistencial de salud (dirección 2)", vbExclamation, Me.Caption
'    Tabdh.Tab = 1
'    ChkInd2.Value = False
'    TxtNomZona1.SetFocus
'End If
End Sub

Private Sub CmbDocDh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(CmbDocDh.ListIndex > -1) <> "" Then SendKeys "{tab}"
If KeyCode = vbKeyDelete Then CmbDocDh.ListIndex = -1
End Sub

Private Sub CmbTipDocAcreditaPaternidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(CmbTipDocAcreditaPaternidad.ListIndex > -1) <> "" Then SendKeys "{tab}"
If KeyCode = vbKeyDelete Then CmbTipDocAcreditaPaternidad.ListIndex = -1
End Sub

Private Sub CmbVinculo_Click()
Dim sCodVinculo As String
sCodVinculo = fc_CodigoComboBox(CmbVinculo, 2)
TxtConcepcion.Text = "__/____"
If sCodVinculo = "04" Then 'getsantes
    Me.Label55(1).Visible = True
    TxtConcepcion.Visible = True
Else
    Me.Label55(1).Visible = False
    TxtConcepcion.Visible = False
End If
End Sub

Private Sub CmbVinculo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(CmbVinculo.ListIndex > -1) <> "" Then SendKeys "{tab}"
End Sub

Private Sub CmdAdd_Click()
If Trim(TxtCodTrab.Text) = "" Or Trim(LblNomTrab.Caption) = "" Then
    MsgBox "Ingrese Codigo de Trabajador, antes de ingresar derechohabientes", vbExclamation, Me.Caption
    Me.TxtCodTrab.SetFocus
    Exit Sub
End If
Me.TxtCodTrab.Enabled = False

Nuevo
End Sub

Private Sub CmdDH_Click()
Unload Frmgrdpla
Load Frmgrdpla
Frmgrdpla.Show vbModal
End Sub

Private Sub CmdUbigeo1_Click()

FrmUbiSunat.TipoCon = 1
Load FrmUbiSunat
FrmUbiSunat.Show
FrmUbiSunat.ZOrder 0
End Sub

Private Sub CmdUbigeo2_Click()
FrmUbiSunat.TipoCon = 2
Load FrmUbiSunat
FrmUbiSunat.Show
FrmUbiSunat.ZOrder 0

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 11610
Me.Height = 7245
Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(CmbCia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(CmbCia, wcia, "00")
CargaCombos
Tabdh.Enabled = False

End Sub

Public Sub Nuevo()
Limpiar
LblIdInternodh.Caption = "0"
SwNuevo = True
Tabdh.Tab = 0
Tabdh.Enabled = True
If Trim(TxtCodTrab.Text) = "" Then
    TxtCodTrab.SetFocus
Else
    Me.TxtApePatDh.SetFocus
End If

End Sub

Private Sub OpcActivoDh_Click()
If OpcActivoDh.Value = True Then
    'FecAltaDH(0).Text = "__/__/____"
    FecAltaDH(0).Enabled = True
    'FecBajaDH(1).Text = "__/__/____"
    FecBajaDH(1).Enabled = False
    
    cbobajadh.Enabled = False
    cbobajadh.ListIndex = -1
    FecBajaDH(1).BackColor = &H8000000F
    cbobajadh.BackColor = &H8000000F
    If FecAltaDH(0).Enabled Then FecAltaDH(0).SetFocus
Else
    FecBajaDH(1).BackColor = vbWhite
    cbobajadh.BackColor = vbWhite
    'FecAltaDH(0).Text = "__/__/____"
    FecAltaDH(0).Enabled = False
    'FecBajaDH(1).Text = "__/__/____"
    FecBajaDH(1).Enabled = True
    cbobajadh.Enabled = True
    cbobajadh.ListIndex = -1
    FecBajaDH(1).SetFocus
End If
End Sub

Private Sub OpcBajaDh_Click()
If OpcActivoDh.Value = True Then
    'FecAltaDH(0).Text = "__/__/____"
    FecAltaDH(0).Enabled = True
    'FecBajaDH(1).Text = "__/__/____"
    FecBajaDH(1).Enabled = False
    cbobajadh.Enabled = False
    cbobajadh.ListIndex = -1
    FecBajaDH(1).BackColor = &H8000000F
    cbobajadh.BackColor = &H8000000F
    FecAltaDH(0).SetFocus
Else
    FecBajaDH(1).BackColor = vbWhite
    cbobajadh.BackColor = vbWhite
    
    
    'FecAltaDH(0).Text = "__/__/____"
    FecAltaDH(0).Enabled = False
    'FecBajaDH(1).Text = "__/__/____"
    FecBajaDH(1).Enabled = True
    cbobajadh.Enabled = True
    cbobajadh.ListIndex = -1
    If FecBajaDH(1).Enabled Then FecBajaDH(1).SetFocus
End If
End Sub

Private Sub OpcIncapazNo_Click()
If OpcIncapazNo.Value = True Then
    TxtIncapaz.Enabled = False
    TxtIncapaz.BackColor = &H8000000F
Else
    TxtIncapaz.BackColor = vbWhite
    TxtIncapaz.Enabled = True
End If
End Sub

Private Sub OpcIncapazSi_Click()
If OpcIncapazNo.Value = True Then
    TxtIncapaz.Enabled = False
    TxtIncapaz.BackColor = &H8000000F
Else
    TxtIncapaz.BackColor = vbWhite
    TxtIncapaz.Enabled = True
End If

End Sub

Public Sub CargaCombos()
Call fc_Descrip_Maestros2("01032", "", CmbDocDh, True)
CargaPais
CargaTelfCodCiudad
'TIPO DE BAJA DH
With cbobajadh
    .NameTab = "maestros_2"
    .NameCod = "cod_maestro2"
    .NameDesc = "descrip"
    .Filtro = "RIGHT(ciamaestro,3)='147' and status!='*' and rtrim(isnull(codsunat,''))<>''"
    .conexion = cn
    .Execute
End With
Call fc_Descrip_Maestros2("01071", "", CmbVinculo, True)
Call fc_Descrip_Maestros2("01032", "", CmbDocDh, True)
Call fc_Descrip_Maestros2("01152", "", CmbTipDocAcreditaPaternidad)
'**************************************
' ZONAS 1
cbo_zona1.NameTab = "maestros_2"
cbo_zona1.NameCod = "cod_maestro2"
cbo_zona1.NameDesc = "descrip"
cbo_zona1.Filtro = "right(ciamaestro,3)='035' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_zona1.conexion = cn
cbo_zona1.Execute
'**************************************
' VIAS 1
cbo_via1.NameTab = "maestros_2"
cbo_via1.NameCod = "cod_maestro2"
cbo_via1.NameDesc = "descrip"
cbo_via1.Filtro = "right(ciamaestro,3)='036' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_via1.conexion = cn
cbo_via1.Execute
'**************************************


' ZONAS 2
cbo_zona2.NameTab = "maestros_2"
cbo_zona2.NameCod = "cod_maestro2"
cbo_zona2.NameDesc = "descrip"
cbo_zona2.Filtro = "right(ciamaestro,3)='035' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_zona2.conexion = cn
cbo_zona2.Execute
'**************************************
' VIAS 2
cbo_via2.NameTab = "maestros_2"
cbo_via2.NameCod = "cod_maestro2"
cbo_via2.NameDesc = "descrip"
cbo_via2.Filtro = "right(ciamaestro,3)='036' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_via2.conexion = cn
cbo_via2.Execute
'**************************************


End Sub

Public Sub CargaPais()
Dim Rq As ADODB.Recordset
Dim Sql As String
Sql = "select idpais,nombre from sunat_pais_emisor order by nombre"
Me.CboPais.Clear
If fAbrRst(Rq, Sql) Then
    Dim I As Integer
    I = 0
    Do While Not Rq.EOF
            Me.CboPais.AddItem Trim(Rq!nombre & "")
            Me.CboPais.ItemData(Me.CboPais.NewIndex) = Trim(Rq!idpais & "")
            I = I + 1
            Rq.MoveNext
    Loop
End If
Rq.Close
Set Rq = Nothing
End Sub

Public Sub CargaTelfCodCiudad()
Dim Rq As ADODB.Recordset
Dim Sql As String
Sql = "select idciudad,ciudad from sunat_telef_cod_ciudad order by ciudad"
Me.CboCodCiudad.Clear
If fAbrRst(Rq, Sql) Then
    Dim I As Integer
    I = 0
    Do While Not Rq.EOF
            Me.CboCodCiudad.AddItem Trim(Rq!ciudad & "")
            Me.CboCodCiudad.ItemData(Me.CboCodCiudad.NewIndex) = Trim(Rq!idciudad & "")
            I = I + 1
            Rq.MoveNext
    Loop
End If
Rq.Close
Set Rq = Nothing
End Sub

Private Sub TxtApeMatDh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtApeMatDh.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtApeMatDh_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtApePatDh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtApePatDh.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtApePatDh_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtCodTrab_Change()
LblNomTrab.Caption = ""
If Len(Trim(TxtCodTrab.Text)) = 5 Then Txtcodtrab_LostFocus
End Sub

Private Sub TxtCodTrab_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtCodTrab.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub Txtcodtrab_LostFocus()
If Trim(TxtCodTrab.Text) = "" Then Exit Sub
Dim Sql As String
Sql = "select dbo.fc_Razsoc_Trabajador(placod,cia) as nom_trab,fcese,codauxinterno "
Sql = Sql & " from planillas where cia='" & wcia & "' and placod='" & Trim(TxtCodTrab.Text) & "' AND status<>'*'"
Dim Rq As ADODB.Recordset
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    If Not IsNull(Rq!fcese) Then
        MsgBox "El Trabajador ya fue Cesado", vbExclamation, Me.Caption
        GoTo Salir:
    End If
    LblNomTrab.Caption = Rq(0)
    sCondInternoTrab = Trim(Rq!codauxinterno & "")
    'Me.Tabdh.Enabled = True
    CmdAdd.Enabled = True
    'CmdAdd.SetFocus
Else
    CmdAdd.Enabled = False
    LblNomTrab.Caption = ""
    'Me.Tabdh.Enabled = False
    MsgBox "No existe codigo,verifique", vbExclamation, Me.Caption
End If
Salir:
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
End Sub


Public Sub NuevoTrab()
If MDIplared.ActiveForm.Name = "FrmConDerechoHab" Then
    Me.TxtCodTrab.Text = Trim(FrmConDerechoHab.Lblplacod.Caption)
    Me.LblNomTrab.Caption = ""
    Txtcodtrab_LostFocus
    Me.Tabdh.Enabled = False
Else
    Me.TxtCodTrab.Text = ""
    Me.LblNomTrab.Caption = ""
    Me.Tabdh.Enabled = False
End If
Limpiar
Tabdh.Tab = 0
Me.TxtCodTrab.Enabled = True
End Sub

Private Sub TxtConcepcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtConcepcion.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtEmail.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtFecNacDh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtFecNacDh.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNomVia1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNomVia1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNomVia1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNomVia2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNomVia2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNomVia2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNomZona1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNomZona1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNomZona1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNomZona2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNomZona2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNomZona2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtNroBlock1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroBlock1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroBlock2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroBlock2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroDocAcreditapaternidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroDocAcreditapaternidad.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroDocDh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroDocDh.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroDpto1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroDpto1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroDpto2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroDpto2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TXtNroEtapa1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TXtNroEtapa1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TXtNroEtapa2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TXtNroEtapa2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroInt1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroInt1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroInt2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroInt2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroKM1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroKM1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroKM2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroKM2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroLote1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroLote1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroLote2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroLote2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroMz1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroMz1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroMz2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroMz2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroVia1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroVia1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtNroVia2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtNroVia2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtPriNomDh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtPriNomDh.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtPriNomDh_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtRef1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtRef1.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtRef1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtRef2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtRef2.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtRef2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtSegNomDh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtSegNomDh.Text) <> "" Then SendKeys "{tab}"
End Sub

Private Sub TxtSegNomDh_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtTelef_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(TxtTelef.Text) <> "" Then SendKeys "{tab}"
End Sub

Public Sub Limpiar()
LblIdInternodh.Caption = "0"
TxtApePatDh.Text = ""
TxtApeMatDh.Text = ""
TxtPriNomDh.Text = ""
TxtSegNomDh.Text = ""
TxtFecNacDh.Text = "__/__/____"
CmbDocDh.ListIndex = -1
TxtNroDocDh.Text = ""
Call rUbiIndCmbBox(CboPais, "604", "000") 'PERU
TxtIncapaz.Text = ""
TxtEmail.Text = ""
Chkestudio.Value = 0
'CboCodCiudad.ListIndex = -1
Call rUbiIndCmbBox(CboCodCiudad, "01", "00") 'LIMA Y CALLAO

TxtTelef.Text = ""
FecAltaDH(0).Text = "__/__/____"
FecBajaDH(1).Text = "__/__/____"
cbobajadh.ListIndex = -1
CmbVinculo.ListIndex = -1
CmbTipDocAcreditaPaternidad.ListIndex = -1
TxtNroDocAcreditapaternidad.Text = ""
TxtConcepcion.Text = "__/____"

cbo_zona1.ListIndex = -1
TxtNomZona1.Text = ""
cbo_via1.ListIndex = -1
TxtNomVia1.Text = ""
TxtNroVia1.Text = ""
TxtNroDpto1.Text = ""
TxtNroKM1.Text = ""
TxtNroInt1.Text = ""
TxtNroBlock1.Text = ""
TxtNroMz1.Text = ""
TXtNroEtapa1.Text = ""
TxtNroLote1.Text = ""
TxtNomUbicacion1.Text = ""
TxtRef1.Text = ""
ChkInd1.Value = 0

cbo_zona2.ListIndex = -1
TxtNomZona2.Text = ""
cbo_via2.ListIndex = -1
TxtNomVia2.Text = ""
TxtNroVia2.Text = ""
TxtNroDpto2.Text = ""
TxtNroKM2.Text = ""
TxtNroInt2.Text = ""
TxtNroBlock2.Text = ""
TxtNroMz2.Text = ""
TXtNroEtapa2.Text = ""
TxtNroLote2.Text = ""
TxtNomUbicacion2.Text = ""
TxtRef2.Text = ""
ChkInd2.Value = 0

OpcVaronDh.Value = True
OpcIncapazNo.Value = True
OpcActivoDh.Value = True

End Sub

Public Function Validar()

Dim sCodVinculo As String
sCodVinculo = fc_CodigoComboBox(CmbVinculo, 2)
Dim sCodDoc As String
sCodDoc = fc_CodigoComboBox(CmbDocDh, 2)

If Trim(TxtCodTrab.Text) = "" Or Trim(LblNomTrab.Caption) = "" Then
    MsgBox "Ingrese código de trabajador", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtCodTrab.SetFocus
    GoTo Salir:
ElseIf Trim(sCondInternoTrab) = "" Then
    MsgBox "Código interno del trabajador no identificado", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtCodTrab.SetFocus
    GoTo Salir:
    
ElseIf SwNuevo = False And Trim(Me.LblIdInternodh.Caption) = "" Then
    MsgBox "Código interno del DERECHOHABIENTE no identificado", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtCodTrab.SetFocus
    GoTo Salir:
ElseIf Trim(TxtApePatDh.Text) = "" Then
    MsgBox "Ingrese Apellido paterno", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtApePatDh.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtApePatDh.Text)) > 20 Then
    MsgBox "La longitud máx. del Apellido paterno es 20 caracteres.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtApePatDh.SetFocus
    GoTo Salir:
ElseIf Trim(TxtApeMatDh.Text) = "" Then
    MsgBox "Ingrese Apellido materno", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtApeMatDh.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtApeMatDh.Text)) > 40 Then
    MsgBox "La longitud máx. del Apellido materno es 40 caracteres.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtApeMatDh.SetFocus
    GoTo Salir:
ElseIf Trim(TxtPriNomDh.Text) = "" Then
    MsgBox "Ingrese primer nombre", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtPriNomDh.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtPriNomDh.Text)) > 20 Then
    MsgBox "La longitud máx. del primer nombre es 20 caracteres.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtPriNomDh.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtSegNomDh.Text)) > 20 Then
    MsgBox "La longitud máx. del segundo nombre es 20 caracteres.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtSegNomDh.SetFocus
    GoTo Salir:
ElseIf Not IsDate(TxtFecNacDh.Text) Then
    MsgBox "La fecha de nacimiento no es válida.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtFecNacDh.SetFocus
    GoTo Salir:
ElseIf CDate(TxtFecNacDh.Text) > Date Then
    MsgBox "La fecha de nacimiento no puede ser mayor al dia de hoy.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtFecNacDh.SetFocus
    GoTo Salir:
ElseIf CmbDocDh.ListIndex = -1 Then
    MsgBox "Elija Tipo de documento del derechohabiente.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    CmbDocDh.SetFocus
    GoTo Salir:
ElseIf Trim(TxtNroDocDh.Text) = "" Or Len(Trim(TxtNroDocDh.Text)) = 0 Then
    MsgBox "Ingrese Número de documento del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtNroDocDh.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroDocDh.Text)) > 15 Then
    MsgBox "La longitud máx. del documento es 15 caracteres.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtNroDocDh.SetFocus
    GoTo Salir:
ElseIf CboPais.ListIndex = -1 Then
    MsgBox "Elija pais emisor documento del derechohabiente.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    CboPais.SetFocus
    GoTo Salir:
ElseIf OpcVaronDh.Value = False And OpcDamaDh.Value = False Then
    MsgBox "Elija sexo del derechohabiente.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    OpcVaronDh.SetFocus
    GoTo Salir:
ElseIf OpcIncapazSi.Value = False And OpcIncapazNo.Value = False Then
    MsgBox "Elija incapacidad del derechohabiente.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    OpcIncapazNo.SetFocus
    GoTo Salir:
ElseIf OpcIncapazSi.Value = True And Trim(TxtIncapaz.Text) = "" Then
    MsgBox "Ingrese Nro Certificado de Incapacidad del derechohabiente.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtIncapaz.SetFocus
    GoTo Salir:
ElseIf Trim(TxtEmail.Text) <> "" And InStr(1, Trim(TxtEmail.Text), "@") = 0 Then
    MsgBox "El email ingresado no es válido", vbExclamation, "Dato Obligatorio - Verifique"
    Tabdh.Tab = 0
    TxtEmail.SetFocus
    GoTo Salir:
ElseIf Trim(TxtTelef.Text) <> "" And CboCodCiudad.ListIndex = -1 Then
    MsgBox "Elija Código telefónico de la ciudad", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    CboCodCiudad.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtTelef.Text)) > 10 Then
    MsgBox "La longitud máx. del telefono es 10 caracteres.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtTelef.SetFocus
    GoTo Salir:
ElseIf OpcActivoDh.Value And Not IsDate(Me.FecAltaDH(0).Text) Then
    MsgBox "La fecha de alta no es válida.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    FecAltaDH(0).SetFocus
    GoTo Salir:
ElseIf CDate(Me.FecAltaDH(0).Text) > Date Then
    MsgBox "La fecha de alta no puede ser mayor al dia de hoy.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    Me.FecAltaDH(0).SetFocus
    GoTo Salir:
ElseIf OpcBajaDh.Value And Not IsDate(Me.FecBajaDH(1).Text) Then
    MsgBox "La fecha de baja no es válida.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    FecBajaDH(1).SetFocus
    GoTo Salir:
'ElseIf OpcBajaDh.Value And CDate(Me.FecBajaDH(1).Text) > Date Then
'    MsgBox "La fecha de baja no puede ser mayor al dia de hoy.", vbExclamation, Me.Caption
'    Me.FecAltaDH(0).SetFocus
'    GoTo Salir:
ElseIf OpcBajaDh.Value And cbobajadh.ListIndex = -1 Then
    MsgBox "Elija motivo de baja", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    cbobajadh.SetFocus
    GoTo Salir:
ElseIf CmbVinculo.ListIndex = -1 Then
    MsgBox "Elija tipo de vinculo", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    CmbVinculo.SetFocus
    GoTo Salir:
   'OBLIGATORIO PARA CONCUBINO E  HIJO MAYOR DE EDAD INCAPACITADO PERMANENTE
ElseIf (sCodVinculo = "03" Or sCodVinculo = "05") And CmbTipDocAcreditaPaternidad.ListIndex = -1 Then
    MsgBox "Elija tipo de documento que acredita el vinculo", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    CmbTipDocAcreditaPaternidad.SetFocus
    GoTo Salir:
ElseIf (sCodVinculo = "03" Or sCodVinculo = "05") And CmbTipDocAcreditaPaternidad.ListIndex > -1 And Trim(TxtNroDocAcreditapaternidad.Text) = "" Then
    MsgBox "Ingrese número de  documento que acredita el vinculo", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtNroDocAcreditapaternidad.SetFocus
    GoTo Salir:
    'GESTANTE
ElseIf sCodVinculo = "04" And Me.TxtConcepcion.Text = "__/____" Then    'solo gestantes
    MsgBox "Ingrese mes y año de concepción,para vinculo gestantes. ", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtConcepcion.SetFocus
    GoTo Salir:
'DIRECCION 1 OBLIGATORIA PARA DOC. PASAPORTE,CARNE DE EXTRANJERIA Y DNI
ElseIf (sCodDoc = "05" Or sCodDoc = "04" Or sCodDoc = "01") And (Trim(TxtNomZona1.Text) = "" And Trim(TxtNomVia1.Text) = "") Then
    MsgBox "Ingrese nombre de zona o via para la dirección 1 del derechohabiente, dato obligatorio", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNomZona1.SetFocus
    GoTo Salir:
ElseIf (sCodDoc = "05" Or sCodDoc = "04" Or sCodDoc = "01") And (Trim(TxtNomUbicacion1.Tag) = "" Or Trim(TxtNomUbicacion1.Text) = "") Then
    MsgBox "Ingrese ubicación para la dirección 1 del derechohabiente, dato obligatorio", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    CmdUbigeo1.SetFocus
    GoTo Salir:


'DIRECCION 1
ElseIf cbo_zona1.ListIndex > -1 And Trim(TxtNomZona1.Text) = "" Then
    MsgBox "Ingrese nombre de zona para la dirección 1 del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNomZona1.SetFocus
    GoTo Salir:
ElseIf cbo_via1.ListIndex > -1 And Trim(TxtNomVia1.Text) = "" Then
    MsgBox "Ingrese nombre de via para la dirección 1 del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNomVia1.SetFocus
    GoTo Salir:
ElseIf (Trim(TxtNomZona1.Text) <> "" Or Trim(TxtNomVia1.Text) <> "") And (Trim(TxtNomUbicacion1.Tag) = "" Or Trim(TxtNomUbicacion1.Text) = "") Then
    MsgBox "Ingrese ubicación para la dirección 1 del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    CmdUbigeo1.SetFocus
    GoTo Salir:
'DIRECCION2
ElseIf cbo_zona2.ListIndex > -1 And Trim(TxtNomZona2.Text) = "" Then
    MsgBox "Ingrese nombre de zona para la dirección 2 del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNomZona2.SetFocus
    GoTo Salir:
ElseIf cbo_via2.ListIndex > -1 And Trim(TxtNomVia2.Text) = "" Then
    MsgBox "Ingrese nombre de via para la dirección 2 del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNomVia2.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroVia1.Text)) > 4 Then
    MsgBox "Para el Nro de via 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNroVia1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroDpto1.Text)) > 4 Then
    MsgBox "Para el Nro de departamento 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNroDpto1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroInt1.Text)) > 4 Then
    MsgBox "Para el Nro de Interior 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNroInt1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroMz1.Text)) > 4 Then
    MsgBox "Para el Nro de manzana 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNroMz1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroLote1.Text)) > 4 Then
    MsgBox "Para el Nro de lote 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNroLote1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroKM1.Text)) > 4 Then
    MsgBox "Para el Nro de Kilometro 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNroKM1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroBlock1.Text)) > 4 Then
    MsgBox "Para el Nro de block 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TxtNroBlock1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TXtNroEtapa1.Text)) > 4 Then
    MsgBox "Para el Nro de etapa 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    TXtNroEtapa1.SetFocus
    GoTo Salir:


'direccion 2
ElseIf Len(Trim(TxtNroVia2.Text)) > 4 Then
    MsgBox "Para el Nro de via 2, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNroVia1.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroDpto1.Text)) > 4 Then
    MsgBox "Para el Nro de departamento 1, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNroDpto2.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroInt2.Text)) > 4 Then
    MsgBox "Para el Nro de Interior 2, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNroInt2.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroMz2.Text)) > 4 Then
    MsgBox "Para el Nro de manzana 2, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNroMz2.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroLote2.Text)) > 4 Then
    MsgBox "Para el Nro de lote 2, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNroLote2.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroKM2.Text)) > 4 Then
    MsgBox "Para el Nro de Kilometro 2, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNroKM2.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TxtNroBlock2.Text)) > 4 Then
    MsgBox "Para el Nro de block 2, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TxtNroBlock2.SetFocus
    GoTo Salir:
ElseIf Len(Trim(TXtNroEtapa2.Text)) > 4 Then
    MsgBox "Para el Nro de etapa 2, solo se permiten 4 caracteres alfanuméricos", vbExclamation, Me.Caption
    Tabdh.Tab = 2
    TXtNroEtapa1.SetFocus
    GoTo Salir:

'ElseIf (Trim(TxtNomZona2.Text) <> "" Or Trim(TxtNomVia2.Text) <> "") And (Trim(TxtNomUbicacion2.Tag) = "" Or Trim(TxtNomUbicacion2.Text) = "") Then
'    MsgBox "Ingrese ubicación para la dirección 2 del derechohabiente", vbExclamation, Me.Caption
'    Tabdh.Tab = 2
'    CmdUbigeo2.SetFocus
'    GoTo Salir:
'
'ElseIf ((Trim(TxtNomZona1.Text) <> "" Or Trim(TxtNomVia1.Text) <> "") And (Trim(TxtNomZona2.Text) <> "" Or Trim(TxtNomVia2.Text) <> "")) And (Me.ChkInd1.Value = 0 Or Me.ChkInd2.Value = 0) Then
'    MsgBox "Indique cual de las direcciones se considerará como Centro asistencial de ESSALUD para el derechohabiente", vbExclamation, Me.Caption
'    Tabdh.Tab = 1
'    CmdUbigeo1.SetFocus
'    GoTo Salir:
ElseIf (sCodDoc = "05" Or sCodDoc = "04" Or sCodDoc = "01") And (Me.ChkInd1.Value = 0 And Me.ChkInd2.Value = 0) Then
    MsgBox "Indique centro asistencial de salud del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    CmdUbigeo1.SetFocus
    GoTo Salir:

ElseIf (Me.ChkInd1.Value = 1 And Me.ChkInd2.Value = 1) Then
    MsgBox "Indique un solo centro asistencial de salud del derechohabiente", vbExclamation, Me.Caption
    Tabdh.Tab = 1
    CmdUbigeo1.SetFocus
    GoTo Salir:
End If

If sCodVinculo = "04" And Me.TxtConcepcion.Text <> "__/____" Then    'solo gestantes
    If Val(Left(Me.TxtConcepcion.Text, 2)) > 12 Then
        MsgBox "El mes ingresado no es válido", vbExclamation, Me.Caption
        Tabdh.Tab = 0
        TxtConcepcion.SetFocus
        GoTo Salir:
    End If
    If Val(Right(Me.TxtConcepcion.Text, 4)) < Year(CDate(FecAltaDH(0).Text)) Then
        MsgBox "El año ingresado debe ser mayor o igual al año de inicio de actividad", vbExclamation, Me.Caption
        Tabdh.Tab = 0
        TxtConcepcion.SetFocus
        GoTo Salir:
    End If
End If

If SwNuevo Then
Dim rsTemporal As ADODB.Recordset
Cadena = "SELECT NUMERO FROM PLADERECHOHAB WHERE CIA = '" & wcia & "' AND (PLACOD <> '" & Trim(TxtCodTrab.Text) & "' OR PLACOD = '" & Trim(TxtCodTrab.Text) & "') AND STATUS != '*' AND NUMERO = '" & Trim(TxtNroDocDh.Text) & "'"
Set rsTemporal = OpenRecordset(Cadena, cn)
If Not rsTemporal.EOF Then
    MsgBox "El Documento de identidad del DERECHOHABIENTE ya se encuentra registrado." & _
    vbCrLf & "Recuerde que el No. de documento de identidad es unico.", vbExclamation, Me.Caption
    Tabdh.Tab = 0
    TxtNroDocDh.SetFocus
    If rsTemporal.State = adStateOpen Then rsTemporal.Close
    GoTo Salir:
End If
If rsTemporal.State = adStateOpen Then rsTemporal.Close
End If

If sCodVinculo = "01" Then
    Dim mEdad As Integer
    Cadena = "SELECT GETDATE() AS FECHA"
    Set rsTemporal = OpenRecordset(Cadena, cn)
    If Not rsTemporal.EOF Then
        mEdad = DateDiff("yyyy", TxtFecNacDh.Text, Format(rsTemporal!fecha, "dd/MM/yyyy"))
    Else
        mEdad = DateDiff("yyyy", TxtFecNacDh.Text, Format(Date, "dd/MM/yyyy"))
    End If
    'If mEdad >= 18 Then MsgBox "DERECHOHABIENTE no puede tener como tipo de vínculo " & CmbVinculo.Text & " teniedo como edad " & CStr(mEdad) & " años.", vbExclamation, Me.Caption: GoTo Salir
    
    If mEdad > 18 Then MsgBox "DERECHOHABIENTE no puede tener como tipo de vínculo " & CmbVinculo.Text & " teniedo como edad " & CStr(mEdad) & " años.", vbExclamation, Me.Caption: GoTo Salir
    
End If

Validar = True
Exit Function
Salir:
    Validar = False
End Function

Public Sub Grabar()
On Error GoTo ErrorTrans
Dim NroTrans As Integer
NroTrans = 0
If Not Validar Then Exit Sub
If MsgBox("Guardar Cambios ??", vbDefaultButton2 + vbYesNo + vbQuestion) = vbNo Then Exit Sub

Dim sCodDoc As String
Dim sCodVinculo As String
Dim sCodZona1 As String
Dim sCodVia1 As String
Dim sCodDocAcreditaVicnculo As String
Dim sCodZona2 As String
Dim sCodVia2 As String
Dim sCodPais As String
Dim sCodTelfCiudad As String

sCodDoc = fc_CodigoComboBox(CmbDocDh, 2)
sCodVinculo = fc_CodigoComboBox(CmbVinculo, 2)
sCodZona1 = IIf(cbo_zona1.ReturnCodigo = -1, "", Format(cbo_zona1.ReturnCodigo, "00"))
sCodVia1 = IIf(cbo_via1.ReturnCodigo = -1, "", Format(cbo_via1.ReturnCodigo, "00"))
sCodDocAcreditaVicnculo = fc_CodigoComboBox(CmbTipDocAcreditaPaternidad, 2)
 
sCodZona2 = IIf(cbo_zona2.ReturnCodigo = -1, "", Format(cbo_zona2.ReturnCodigo, "00"))
sCodVia2 = IIf(cbo_via2.ReturnCodigo = -1, "", Format(cbo_via2.ReturnCodigo, "00"))
sCodPais = fc_CodigoComboBox(Me.CboPais, 3)
sCodTelfCiudad = fc_CodigoComboBox(Me.CboCodCiudad, 2)

Me.MousePointer = 11

cn.BeginTrans
NroTrans = 1
If SwNuevo = False Then
    Sql = "usp_d_pladerechohab '" & wcia & "'," & Me.LblIdInternodh.Caption
    cn.Execute Sql, 64
End If
Sql = "usp_i_pladerechohab "
Sql = Sql & "'" & wcia & "'"
Sql = Sql & ",'" & Me.TxtCodTrab.Text & "'"
Sql = Sql & ",'" & sCondInternoTrab & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtApePatDh.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtApeMatDh.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtPriNomDh.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtSegNomDh.Text)) & "'"
Sql = Sql & ",'" & sCodDoc & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroDocDh.Text)) & "'"
Sql = Sql & ",'" & Format(TxtFecNacDh, FormatFecha) & "'"
Sql = Sql & ",'" & sCodVinculo & "'"
Sql = Sql & ",''"
Sql = Sql & ",'" & IIf(OpcVaronDh.Value, "M", "F") & "'"
Sql = Sql & ",'" & IIf(OpcActivoDh.Value = True, 1, 0) & "'"
Sql = Sql & ",''"
Sql = Sql & ",'" & IIf(OpcIncapazSi.Value, "S", "N") & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtIncapaz.Text)) & "'"
Sql = Sql & ",'N'"
Sql = Sql & ",''"
Sql = Sql & ",NULL"
Sql = Sql & ",'" & IIf(Chkestudio.Value, "S", "N") & "'"
Sql = Sql & ",''"
Sql = Sql & ",'" & Format(IIf(cbobajadh.ReturnCodigo = -1, "", cbobajadh.ReturnCodigo), "00") & "'"
Sql = Sql & ",''"
Sql = Sql & ",'" & sCodZona1 & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNomZona1.Text)) & "'"
Sql = Sql & ",'" & sCodVia1 & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNomVia1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroDpto1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroInt1.Text)) & "'"
Sql = Sql & ",'" & TxtNomUbicacion1.Tag & "'"
Sql = Sql & ",'" & sCodDocAcreditaVicnculo & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroDocAcreditapaternidad.Text)) & "'"
Sql = Sql & ",'" & Format(FecAltaDH(0).Text, "mm/dd/yyyy") & "'"
If FecBajaDH(1).Text = "__/__/____" Then
    Sql = Sql & ",null"
Else
    Sql = Sql & ",'" & Format(FecBajaDH(1).Text, "mm/dd/yyyy") & "'"
End If
Sql = Sql & ",'" & Apostrofe(Trim(TxtRef1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroVia1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroMz1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroLote1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroKM1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroBlock1.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TXtNroEtapa1.Text)) & "'"
Sql = Sql & "," & ChkInd1.Value & ""
Sql = Sql & ",'" & sCodVia2 & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNomVia2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroVia2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroDpto2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroInt2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroMz2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroLote2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroKM2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNroBlock2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TXtNroEtapa2.Text)) & "'"
Sql = Sql & ",'" & sCodZona2 & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtNomZona2.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtRef2.Text)) & "'"
Sql = Sql & ",'" & TxtNomUbicacion2.Tag & "'"
Sql = Sql & "," & ChkInd2.Value & ""
Sql = Sql & ",'" & sCodPais & "'"
If TxtConcepcion.Text = "__/____" Then
    Sql = Sql & ",''"
Else
    Sql = Sql & ",'" & Left(TxtConcepcion, 2) & Right(TxtConcepcion.Text, 4) & "'"
End If
Sql = Sql & ",'" & sCodTelfCiudad & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtTelef.Text)) & "'"
Sql = Sql & ",'" & Apostrofe(Trim(TxtEmail.Text)) & "'"
Sql = Sql & "," & Me.LblIdInternodh.Caption
Sql = Sql & "," & IIf(SwNuevo = True, "0", "1")
cn.Execute Sql, 64
Debug.Print Sql

cn.CommitTrans
Limpiar
Me.CmdAdd.Enabled = True
Me.CmdAdd.SetFocus
Me.MousePointer = 0
MsgBox "Se guardaron los datos correctamente", vbInformation, Me.Caption
If Not SwNuevo Then Unload Me

Exit Sub
ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

MsgBox ERR.Description, vbCritical, Me.Caption

Me.MousePointer = 0

End Sub

Public Sub CargaDerechohabiente()
Dim Rq As ADODB.Recordset
Screen.MousePointer = 11
Sql = "usp_c_pladerechohab '" & wcia & "'," & Me.LblIdInternodh.Caption
If Not fAbrRst(Rq, Sql) Then
    MsgBox "No existe derechohabiente", vbExclamation, Me.Caption
    GoTo Salir:
End If

TxtApePatDh.Text = Trim(Rq!ap_pat & "")
TxtApeMatDh.Text = Trim(Rq!ap_mat & "")
TxtPriNomDh.Text = Trim(Rq!nom_1 & "")
TxtSegNomDh.Text = Trim(Rq!nom_2 & "")
TxtFecNacDh.Text = Format(Rq!fec_nac, "dd/mm/yyyy")

Call rUbiIndCmbBox(Me.CmbDocDh, Trim(Rq!cod_doc), "00")

TxtNroDocDh.Text = Trim(Rq!numero & "")

Call rUbiIndCmbBox(CboPais, Trim(Rq!cod_pais_emision & ""), "000") 'PERU
If Trim(Rq!nrocertificado & "") = "S" Then
    OpcIncapazSi.Value = True
Else
    OpcIncapazNo.Value = True
End If
TxtIncapaz.Text = Trim(Rq!nrocertificado & "")
TxtEmail.Text = Trim(Rq!Email & "")
If Trim(Rq!escolar & "") = "S" Then Chkestudio.Value = 1 Else Chkestudio.Value = 0


Call rUbiIndCmbBox(CboCodCiudad, Trim(Rq!telef_cod_ciudad & ""), "00")

TxtTelef.Text = Trim(Rq!telefono & "")
If Trim(Rq!situacion & "") = 1 Then Me.OpcActivoDh.Value = True Else Me.OpcBajaDh.Value = True
FecAltaDH(0).Text = Format(Rq!fecha_alta, "dd/mm/yyyy")
If IsNull(Rq!fecha_baja) Then
    FecBajaDH(1).Text = "__/__/____"
Else
    FecBajaDH(1).Text = Format(Rq!fecha_baja, "dd/mm/yyyy")
End If
If Trim(Rq!cod_mot_baja & "") = "00" Then
    cbobajadh.ListIndex = -1
Else
    cbobajadh.SetIndice Trim(Rq!cod_mot_baja & "")
End If

Call rUbiIndCmbBox(CmbVinculo, Trim(Rq!codvinculo & ""), "00")


Call rUbiIndCmbBox(CmbTipDocAcreditaPaternidad, Trim(Rq!tipdoc_acredita_paternidad & ""), "00")

TxtNroDocAcreditapaternidad.Text = Trim(Rq!nrodoc_acredita_paternidad & "")
If Trim(Rq!mes_concepcion & "") = "" Then
    TxtConcepcion.Text = "__/____"
Else
    TxtConcepcion.Text = Left(Trim(Rq!mes_concepcion & ""), 2) & "/" & Right(Trim(Rq!mes_concepcion & ""), 4)
End If

If Trim(Rq!sexo & "") = "M" Then Me.OpcVaronDh.Value = True Else Me.OpcDamaDh.Value = True
Me.TxtIncapaz = Trim(Rq!nrocertificado & "")

If Trim(Rq!cod_zona & "") = "00" Or Trim(Rq!cod_zona & "") = "" Then
    cbo_zona1.ListIndex = -1
Else
    cbo_zona1.SetIndice Trim(Rq!cod_zona & "")
End If
TxtNomZona1.Text = Trim(Rq!NOM_ZONA & "")
If Trim(Rq!cod_via & "") = "00" Or Trim(Rq!cod_via & "") = "" Then
    cbo_via1.ListIndex = -1
Else
    cbo_via1.SetIndice Trim(Rq!cod_via & "")
End If
TxtNomVia1.Text = Trim(Rq!NOM_VIA & "")
TxtNroVia1.Text = Trim(Rq!nro_via1 & "")
TxtNroDpto1.Text = Trim(Rq!NRO & "")
TxtNroKM1.Text = Trim(Rq!nro_kilometro1 & "")
TxtNroInt1.Text = Trim(Rq!Interior & "")
TxtNroBlock1.Text = Trim(Rq!nro_block1 & "")
TxtNroMz1.Text = Trim(Rq!nro_manzana1 & "")
TXtNroEtapa1.Text = Trim(Rq!nro_etapa1 & "")
TxtNroLote1.Text = Trim(Rq!nro_lote1 & "")
TxtNomUbicacion1.Text = Trim(Rq!nom_ubigeo & "")
TxtNomUbicacion1.Tag = Trim(Rq!ubigeo & "")
TxtRef1.Text = Trim(Rq!referencia & "")

If Rq!indicador_centro_essalud1 = True Then
    ChkInd1.Value = 1
Else
    ChkInd1.Value = 0
End If

If Trim(Rq!cod_zona2 & "") = "00" Or Trim(Rq!cod_zona2 & "") = "" Then
    cbo_zona2.ListIndex = -1
Else
    cbo_zona2.SetIndice Trim(Rq!cod_zona2 & "")
End If
TxtNomZona2.Text = Trim(Rq!NOM_ZONA2 & "")

If Trim(Rq!cod_via2 & "") = "00" Or Trim(Rq!cod_via2 & "") = "" Then
  cbo_via2.ListIndex = -1
Else
  cbo_via2.SetIndice Trim(Rq!cod_via2 & "")
End If

TxtNomVia2.Text = Trim(Rq!NOM_VIA2 & "")
TxtNroVia2.Text = Trim(Rq!nro_via2 & "")
TxtNroDpto2.Text = Trim(Rq!nro_departamento2 & "")
TxtNroKM2.Text = Trim(Rq!nro_kilometro2 & "")
TxtNroInt2.Text = Trim(Rq!nro_interior2 & "")
TxtNroBlock2.Text = Trim(Rq!nro_block2 & "")
TxtNroMz2.Text = Trim(Rq!nro_manzana2 & "")
TXtNroEtapa2.Text = Trim(Rq!nro_etapa2 & "")
TxtNroLote2.Text = Trim(Rq!nro_lote2 & "")
TxtNomUbicacion2.Text = Trim(Rq!nom_ubigeo2 & "")
TxtNomUbicacion2.Tag = Trim(Rq!ubigeo2 & "")
TxtRef2.Text = Trim(Rq!referencia2 & "")
'ChkInd2.Value = IIf(IsNull(Trim(Rq!indicador_centro_essalud2 & "")) Or Trim(Rq!indicador_centro_essalud2 & "") = False, 0, 1)

If Rq!indicador_centro_essalud2 = True Then
    ChkInd2.Value = 1
Else
    ChkInd2.Value = 0
End If


TxtCodTrab.Enabled = False
Me.LblNomTrab.Enabled = False
Me.Tabdh.Enabled = True
Me.CmdAdd.Enabled = False
Me.Tabdh.Tab = 0
SwNuevo = False
Salir:
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
End Sub
