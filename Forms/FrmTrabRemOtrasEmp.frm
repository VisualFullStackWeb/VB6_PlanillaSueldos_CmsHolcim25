VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmTrabRemOtrasEmp 
   Caption         =   "Remuneraciones de otras empresas"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15555
   Icon            =   "FrmTrabRemOtrasEmp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   15555
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   17895
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   90
         Width           =   14250
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
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   825
      End
   End
   Begin TabDlg.SSTab SStabOtras 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   13996
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Trabajadores con otrasremunerciones"
      TabPicture(0)   =   "FrmTrabRemOtrasEmp.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DgrdTrab"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSpanelNewTrab"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Otras empresas"
      TabPicture(1)   =   "FrmTrabRemOtrasEmp.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSPnelEmpresa"
      Tab(1).Control(1)=   "FrameOtrasEmpr"
      Tab(1).Control(2)=   "FrameTrab"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Remuneraciones otras empresas"
      TabPicture(2)   =   "FrmTrabRemOtrasEmp.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSPanelremunera"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).ControlCount=   3
      Begin Threed.SSPanel SSPanelremunera 
         Height          =   6375
         Left            =   -71160
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   11245
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
         Begin VB.Frame Frame4 
            Height          =   1455
            Left            =   120
            TabIndex        =   46
            Top             =   2040
            Width           =   8055
            Begin VB.TextBox TxtQuinta 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6720
               MaxLength       =   8
               TabIndex        =   49
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox TxtRemuneracion 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6720
               MaxLength       =   8
               TabIndex        =   47
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H80000009&
               Caption         =   "Remuneración del periodo sin extraordinarios"
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
               Left            =   60
               TabIndex        =   51
               Top             =   120
               Width           =   7935
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Ingrese Quinta Categoria descontada correspondiente al periido"
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
               Index           =   10
               Left            =   1080
               TabIndex        =   50
               Top             =   1080
               Width           =   5445
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Ingrese Remuneración afecta a quinta correspondiente al periodo"
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
               Index           =   9
               Left            =   960
               TabIndex        =   48
               Top             =   600
               Width           =   5595
            End
         End
         Begin VB.TextBox TxtBasico 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7200
            MaxLength       =   8
            TabIndex        =   44
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox TxtIdEmpRemuneracion 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   41
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox TxtCodTrabremuneracion 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   35
            Top             =   600
            Width           =   855
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   825
            Left            =   1800
            TabIndex        =   36
            ToolTipText     =   "Salir"
            Top             =   5400
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
            _StockProps     =   78
            Caption         =   "Anular"
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":035E
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   825
            Left            =   240
            TabIndex        =   37
            ToolTipText     =   "Salir"
            Top             =   5400
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":07B0
         End
         Begin Threed.SSCommand SSCommand5 
            Height          =   825
            Left            =   6840
            TabIndex        =   38
            ToolTipText     =   "Salir"
            Top             =   5400
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
            _StockProps     =   78
            Caption         =   "Salir"
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":0D4A
         End
         Begin VB.Label LblMesNumero 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3720
            TabIndex        =   60
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
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
            Index           =   12
            Left            =   360
            TabIndex        =   59
            Top             =   120
            Width           =   660
         End
         Begin VB.Label LblAyoRemuneracion 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   58
            Top             =   120
            Width           =   735
         End
         Begin VB.Label LblMesRemuneracion 
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1920
            TabIndex        =   57
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ingrese Remuneracion Basica para proyectar"
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
            Index           =   8
            Left            =   3120
            TabIndex        =   45
            Top             =   1680
            Width           =   3840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
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
            Index           =   7
            Left            =   240
            TabIndex        =   43
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label LblNomEmpRemuneracion 
            BackColor       =   &H80000014&
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
            Left            =   2040
            TabIndex        =   42
            Top             =   1110
            Width           =   6015
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Index           =   6
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   930
         End
         Begin VB.Label LblNomRemuneracion 
            BackColor       =   &H80000014&
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
            Left            =   2040
            TabIndex        =   39
            Top             =   600
            Width           =   6015
         End
      End
      Begin Threed.SSPanel SSPnelEmpresa 
         Height          =   1815
         Left            =   -71400
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   3201
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
         Begin VB.TextBox Txtempresa 
            Height          =   285
            Left            =   1080
            TabIndex        =   31
            Top             =   240
            Width           =   7095
         End
         Begin Threed.SSCommand BtnEmpDeshabilita 
            Height          =   825
            Left            =   1800
            TabIndex        =   27
            ToolTipText     =   "Salir"
            Top             =   840
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
            _StockProps     =   78
            Caption         =   "Deshabilitar"
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":1064
         End
         Begin Threed.SSCommand BtnEmpAcepta 
            Height          =   825
            Left            =   240
            TabIndex        =   28
            ToolTipText     =   "Salir"
            Top             =   840
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":16DE
         End
         Begin Threed.SSCommand SSCommand4 
            Height          =   825
            Left            =   6840
            TabIndex        =   29
            ToolTipText     =   "Salir"
            Top             =   840
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
            _StockProps     =   78
            Caption         =   "Salir"
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":1C78
         End
         Begin Threed.SSCommand BtnEmpAnula 
            Height          =   825
            Left            =   3240
            TabIndex        =   32
            ToolTipText     =   "Salir"
            Top             =   840
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
            _StockProps     =   78
            Caption         =   "Anular"
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":1F92
         End
         Begin VB.Label LblIdEmpresa 
            AutoSize        =   -1  'True
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
            Index           =   6
            Left            =   7320
            TabIndex        =   33
            Top             =   0
            Width           =   75
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
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
            Index           =   5
            Left            =   240
            TabIndex        =   30
            Top             =   285
            Width           =   735
         End
      End
      Begin Threed.SSPanel SSpanelNewTrab 
         Height          =   1815
         Left            =   3480
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   8295
         _Version        =   65536
         _ExtentX        =   14631
         _ExtentY        =   3201
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
         Begin VB.TextBox TxtCodNewTrab 
            Height          =   315
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin Threed.SSCommand BtnAnula 
            Height          =   825
            Left            =   1800
            TabIndex        =   23
            ToolTipText     =   "Salir"
            Top             =   840
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
            _StockProps     =   78
            Caption         =   "Anular"
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":23E4
         End
         Begin Threed.SSCommand BtnAceptar 
            Height          =   825
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "Salir"
            Top             =   840
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":2836
         End
         Begin Threed.SSCommand SSCommand3 
            Height          =   825
            Left            =   6840
            TabIndex        =   25
            ToolTipText     =   "Salir"
            Top             =   840
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   1455
            _StockProps     =   78
            Caption         =   "Salir"
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
            Picture         =   "FrmTrabRemOtrasEmp.frx":2DD0
         End
         Begin VB.Label LblNombreNewTrab 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2160
            TabIndex        =   22
            Top             =   270
            Width           =   5895
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6015
         Left            =   -74880
         TabIndex        =   15
         Top             =   1800
         Width           =   15135
         Begin TrueOleDBGrid70.TDBGrid DgrdRemunera 
            Height          =   5865
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   10345
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "Codigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nombre"
            Columns(1).DataField=   "Nombre"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "IdPlaOtracia"
            Columns(2).DataField=   "IdPlaOtracia"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Empresas"
            Columns(3).DataField=   "NombreCia"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Basico"
            Columns(4).DataField=   "basico"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Ingreso"
            Columns(5).DataField=   "ingreso"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Quinta"
            Columns(6).DataField=   "quinta"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Año"
            Columns(7).DataField=   "ayo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Mes"
            Columns(8).DataField=   "Mes"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1508"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1429"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=9155"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=9075"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=1349"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1270"
            Splits(0)._ColumnProps(12)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=7620"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=7541"
            Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(18)=   "Column(4).Width=2117"
            Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2037"
            Splits(0)._ColumnProps(21)=   "Column(4)._ColStyle=2"
            Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(23)=   "Column(5).Width=2540"
            Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=2461"
            Splits(0)._ColumnProps(26)=   "Column(5)._ColStyle=2"
            Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(28)=   "Column(6).Width=2381"
            Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=2302"
            Splits(0)._ColumnProps(31)=   "Column(6)._ColStyle=2"
            Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(33)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(38)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(41)=   "Column(8).Visible=0"
            Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HD7D7D7&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HFF8000&"
            _StyleDefs(23)  =   ":id=11,.appearance=0"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFF8000&"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
            _StyleDefs(28)  =   ":id=14,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFF8000&"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
            _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
            _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(75)  =   "Named:id=33:Normal"
            _StyleDefs(76)  =   ":id=33,.parent=0"
            _StyleDefs(77)  =   "Named:id=34:Heading"
            _StyleDefs(78)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   ":id=34,.wraptext=-1"
            _StyleDefs(80)  =   "Named:id=35:Footing"
            _StyleDefs(81)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(82)  =   "Named:id=36:Selected"
            _StyleDefs(83)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(84)  =   "Named:id=37:Caption"
            _StyleDefs(85)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(86)  =   "Named:id=38:HighlightRow"
            _StyleDefs(87)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(88)  =   "Named:id=39:EvenRow"
            _StyleDefs(89)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(90)  =   "Named:id=40:OddRow"
            _StyleDefs(91)  =   ":id=40,.parent=33"
            _StyleDefs(92)  =   "Named:id=41:RecordSelector"
            _StyleDefs(93)  =   ":id=41,.parent=34"
            _StyleDefs(94)  =   "Named:id=42:FilterBar"
            _StyleDefs(95)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   15135
         Begin VB.TextBox TxtIdEmpRem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   53
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox Cmbmes 
            Height          =   315
            ItemData        =   "FrmTrabRemOtrasEmp.frx":30EA
            Left            =   12300
            List            =   "FrmTrabRemOtrasEmp.frx":3112
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   2475
         End
         Begin VB.TextBox Txtano 
            Height          =   285
            Left            =   11415
            TabIndex        =   16
            Top             =   360
            Width           =   825
         End
         Begin VB.TextBox TxtCodRemun 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   12
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            Caption         =   "Empresa"
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
            Index           =   11
            Left            =   120
            TabIndex        =   55
            Top             =   645
            Width           =   735
         End
         Begin VB.Label LblNomEmpRem 
            BackColor       =   &H80000014&
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
            Left            =   2400
            TabIndex        =   54
            Top             =   630
            Width           =   7935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            Caption         =   "Periodo"
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
            Index           =   3
            Left            =   10680
            TabIndex        =   18
            Top             =   405
            Width           =   660
         End
         Begin VB.Label LblNomRemun 
            BackColor       =   &H80000014&
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
            Left            =   2400
            TabIndex        =   14
            Top             =   150
            Width           =   7935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
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
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   165
            Width           =   930
         End
      End
      Begin VB.Frame FrameOtrasEmpr 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6615
         Left            =   -74880
         TabIndex        =   9
         Top             =   1200
         Width           =   15135
         Begin TrueOleDBGrid70.TDBGrid DgrdOtras 
            Height          =   6345
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   11192
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "Codigo"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nombre"
            Columns(1).DataField=   "Nombre"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "IdPlaOtracia"
            Columns(2).DataField=   "IdPlaOtracia"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Empresas"
            Columns(3).DataField=   "NombreCia"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Estado"
            Columns(4).DataField=   "estado"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=10874"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=10795"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=1349"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1270"
            Splits(0)._ColumnProps(12)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=9525"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=9446"
            Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HD7D7D7&,.bold=0"
            _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HFF8000&"
            _StyleDefs(23)  =   ":id=11,.appearance=0"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFF8000&"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
            _StyleDefs(28)  =   ":id=14,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFF8000&"
            _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
            _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(59)  =   "Named:id=33:Normal"
            _StyleDefs(60)  =   ":id=33,.parent=0"
            _StyleDefs(61)  =   "Named:id=34:Heading"
            _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   ":id=34,.wraptext=-1"
            _StyleDefs(64)  =   "Named:id=35:Footing"
            _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(66)  =   "Named:id=36:Selected"
            _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(68)  =   "Named:id=37:Caption"
            _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(70)  =   "Named:id=38:HighlightRow"
            _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(72)  =   "Named:id=39:EvenRow"
            _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(74)  =   "Named:id=40:OddRow"
            _StyleDefs(75)  =   ":id=40,.parent=33"
            _StyleDefs(76)  =   "Named:id=41:RecordSelector"
            _StyleDefs(77)  =   ":id=41,.parent=34"
            _StyleDefs(78)  =   "Named:id=42:FilterBar"
            _StyleDefs(79)  =   ":id=42,.parent=33"
         End
      End
      Begin VB.Frame FrameTrab 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   15135
         Begin VB.TextBox Txtcodpla 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   6
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label LblIdEmpRemun 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
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
            Index           =   11
            Left            =   14040
            TabIndex        =   52
            Top             =   360
            Width           =   75
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
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
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   165
            Width           =   930
         End
         Begin VB.Label Lblnombre 
            BackColor       =   &H80000014&
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
            Left            =   2400
            TabIndex        =   7
            Top             =   150
            Width           =   12615
         End
      End
      Begin TrueOleDBGrid70.TDBGrid DgrdTrab 
         Height          =   7305
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   12885
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Codigo"
         Columns(0).DataField=   "Codigo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nombre"
         Columns(1).DataField=   "Nombre"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=23178"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=23098"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
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
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&HD7D7D7&,.bold=0"
         _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41,.bgcolor=&HFF8000&"
         _StyleDefs(23)  =   ":id=11,.appearance=0"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFF8000&"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bgcolor=&HFF8000&,.fgcolor=&HFFFFFF&"
         _StyleDefs(28)  =   ":id=14,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(29)  =   ":id=14,.fontname=MS Sans Serif"
         _StyleDefs(30)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bgcolor=&HFF8000&"
         _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(47)  =   "Named:id=33:Normal"
         _StyleDefs(48)  =   ":id=33,.parent=0"
         _StyleDefs(49)  =   "Named:id=34:Heading"
         _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   ":id=34,.wraptext=-1"
         _StyleDefs(52)  =   "Named:id=35:Footing"
         _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   "Named:id=36:Selected"
         _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=37:Caption"
         _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "FrmTrabRemOtrasEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnAceptar_Click()
If Trim(TxtCodNewTrab.Text & "") <> "" Then
   Graba_Nuevo_Trabajador (1)
   SSpanelNewTrab.Visible = False
Else
   MsgBox "Ingrese código de Trabajador", vbInformation, Me.Caption
   Exit Sub
End If
End Sub
Private Sub Graba_Nuevo_Trabajador(Accion As Integer)

Dim Sql As String
Dim Rq As ADODB.Recordset

If Accion = 1 Then
   Sql = "SELECT placod from Pla_Trab_Otras_Empresas where cia='" & wcia & "' and placod='" & TxtCodNewTrab.Text & "' and status<>'*'"
Else
   Sql = "select nombrecia from Pla_Trab_Otras_Empresas_Cias where cia='" & wcia & "' and placod='" & TxtCodNewTrab.Text & "' and status<>'*'"
End If

If fAbrRst(Rq, Sql) Then
   If Accion = 1 Then
      MsgBox "Trabajador ya esta registrado", vbInformation
   Else
      MsgBox "Trabajador tiene asignadas empresas, no se puede anular", vbInformation
   End If
   Rq.Close: Set Rq = Nothing
   Exit Sub
End If
Rq.Close: Set Rq = Nothing

Sql$ = "Usp_Pla_Mant_Trab_Otras_Empresas '" & wcia & "','T','" & TxtCodNewTrab.Text & "'," & Accion & ",'" & wuser & "','',0,0,0,0,0,0"
cn.Execute Sql$
Carga_Trabajadores
End Sub

Private Sub BtnAnula_Click()
   Graba_Nuevo_Trabajador (3)
   SSpanelNewTrab.Visible = False
End Sub

Private Sub BtnEmpAcepta_Click()
If Trim(Txtempresa.Text & "") <> "" Then
   Graba_Nueva_Empresa (1)
   SSPnelEmpresa.Visible = False
Else
   MsgBox "Ingrese Nombre de la Empresa", vbInformation
   Exit Sub
End If
End Sub

Private Sub BtnEmpAnula_Click()
If Trim(Txtempresa.Text & "") <> "" Then
   Graba_Nueva_Empresa (3)
   SSPnelEmpresa.Visible = False
Else
   MsgBox "Seleccione Empresa", vbInformation
   Exit Sub
End If
End Sub

Private Sub BtnEmpDeshabilita_Click()
If Trim(Txtempresa.Text & "") <> "" Then
   If BtnEmpDeshabilita.Caption = "Activar" Then
      Graba_Nueva_Empresa (4)
   Else
      Graba_Nueva_Empresa (2)
   End If
   SSPnelEmpresa.Visible = False
Else
   MsgBox "Seleccione Empresa", vbInformation
   Exit Sub
End If

End Sub

Private Sub CmbMes_Click()
If Trim(TxtCodRemun.Text & "") = "" Then Exit Sub
Call Carga_Remuneraciones(Trim(TxtCodRemun.Text & ""), Trim(LblNomRemun.Caption & ""), Trim(TxtIdEmpRem.Text & ""), Trim(LblNomEmpRem.Caption & ""), Txtano.Text, Cmbmes.ListIndex + 1)
End Sub

Private Sub DgrdOtras_DblClick()
SSPnelEmpresa.Visible = False
If Trim(DgrdOtras.Columns(0) & "") <> "" Then Call Carga_Remuneraciones(Trim(DgrdOtras.Columns(0) & ""), Trim(DgrdOtras.Columns(1) & ""), Trim(DgrdOtras.Columns(2) & ""), Trim(DgrdOtras.Columns(3) & ""), Txtano.Text, Cmbmes.ListIndex + 1)
End Sub
Private Sub Carga_Remuneraciones(lCod As String, lNombre As String, lIDEmpresa As Integer, lNomEmpresa As String, layo As Integer, lMEs As Integer)

If Not IsNumeric(Txtano.Text) Then
   Set DgrdRemunera.DataSource = Nothing
   Exit Sub
End If
If Val(Txtano.Text) < 2015 Then
   Set DgrdRemunera.DataSource = Nothing
   Exit Sub
End If

TxtCodRemun.Text = lCod
LblNomRemun.Caption = lNombre
LblIdEmpRemun(11).Caption = lIDEmpresa
TxtIdEmpRem.Text = lIDEmpresa
LblNomEmpRem.Caption = lNomEmpresa

Dim rstrab As New ADODB.Recordset

Sql$ = "usp_pla_trab_otras_empresas_quinta '" & wcia & "','R','" & lCod & "'," & layo & "," & lMEs & "," & lIDEmpresa & ""

cn.CursorLocation = adUseClient
Set rstrab = New ADODB.Recordset
Set rstrab = cn.Execute(Sql$, 64)
Set DgrdRemunera.DataSource = rstrab


SStabOtras.Tab = 2
End Sub

Private Sub DGrdRemunera_DblClick()
If Trim(DgrdRemunera.Columns(0) & "") = "" Then Exit Sub
Call Carga_Remuneraciones(Trim(DgrdRemunera.Columns(0) & ""), Trim(DgrdRemunera.Columns(1) & ""), Trim(DgrdRemunera.Columns(2) & ""), Trim(DgrdRemunera.Columns(3) & ""), Trim(DgrdRemunera.Columns(7) & ""), Trim(DgrdRemunera.Columns(8) & ""))

TxtCodTrabremuneracion.Text = Trim(DgrdRemunera.Columns(0) & "")
LblNomRemuneracion.Caption = Trim(DgrdRemunera.Columns(1) & "")
TxtIdEmpRemuneracion.Text = Trim(DgrdRemunera.Columns(2) & "")
LblNomEmpRemuneracion.Caption = Trim(DgrdRemunera.Columns(3) & "")
LblAyoRemuneracion.Caption = Trim(DgrdRemunera.Columns(7) & "")
LblMesRemuneracion.Caption = Cmbmes.Text
LblMesNumero.Caption = Cmbmes.ListIndex + 1
TxtBasico.Text = Trim(DgrdRemunera.Columns(4) & "")
TxtRemuneracion.Text = Trim(DgrdRemunera.Columns(5) & "")
TxtQuinta.Text = Trim(DgrdRemunera.Columns(6) & "")
SSPanelremunera.Visible = True
End Sub

Private Sub DgrdTrab_DblClick()
SSpanelNewTrab.Visible = False
If Trim(DgrdTrab.Columns(0) & "") <> "" Then Call Carga_Empresas(Trim(DgrdTrab.Columns(0) & ""), Trim(DgrdTrab.Columns(1) & ""))
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 15675: Me.Height = 9255
SStabOtras.Tab = 0
Txtano.Text = Format(Year(Date), "0000")
If Month(Date) = 1 Then
   Cmbmes.ListIndex = 0
Else
   Cmbmes.ListIndex = Month(Date) - 1
End If
Carga_Trabajadores
End Sub
Public Sub Carga_Trabajadores()
Dim rstrab As New ADODB.Recordset

Sql$ = "usp_pla_trab_otras_empresas_quinta '" & wcia & "','T','',0,0,0"

cn.CursorLocation = adUseClient
Set rstrab = New ADODB.Recordset
Set rstrab = cn.Execute(Sql$, 64)
Set DgrdTrab.DataSource = rstrab
End Sub

Public Sub Carga_Empresas(lCod As String, lNombre As String)
Dim rstrab As New ADODB.Recordset

Sql$ = "usp_pla_trab_otras_empresas_quinta '" & wcia & "','E','" & lCod & "',0,0,0"

cn.CursorLocation = adUseClient
Set rstrab = New ADODB.Recordset
Set rstrab = cn.Execute(Sql$, 64)
Set DgrdOtras.DataSource = rstrab

Txtcodpla.Text = lCod
Lblnombre.Caption = lNombre

SStabOtras.Tab = 1

End Sub

Private Sub SSCommand1_Click()
Graba_Remuneracion (3)
End Sub

Private Sub SSCommand2_Click()
Graba_Remuneracion (1)
End Sub

Private Sub SSCommand3_Click()
SSpanelNewTrab.Visible = False
End Sub
Public Sub Nuevo_Trabajador_Otras()
BtnAnula.Visible = False
BtnAceptar.Visible = True
TxtCodNewTrab.Text = ""
LblNombreNewTrab.Caption = ""
SSpanelNewTrab.Visible = True
End Sub

Public Sub Anula_Trabajador_Otras()

If Trim(DgrdTrab.Columns(0) & "") = "" Then Exit Sub

BtnAceptar.Visible = False
BtnAnula.Visible = True

TxtCodNewTrab.Text = Trim(DgrdTrab.Columns(0) & "")
LblNombreNewTrab.Caption = Trim(DgrdTrab.Columns(1) & "")
SSpanelNewTrab.Visible = True
End Sub
Public Sub Anula_Empresa_Otras()

If Trim(DgrdOtras.Columns(0) & "") = "" Then Exit Sub

BtnEmpAcepta.Visible = False

If Trim(DgrdOtras.Columns(4) & "") = "DESHABILITADO" Then
   BtnEmpDeshabilita.Caption = "Activar"
Else
   BtnEmpDeshabilita.Caption = "Deshabilitar"
End If

BtnEmpDeshabilita.Visible = True
BtnEmpAnula.Visible = True

Txtempresa.Text = Trim(DgrdOtras.Columns(3) & "")
LblIdEmpresa(6#).Caption = Trim(DgrdOtras.Columns(2) & "")

SSPnelEmpresa.Visible = True
   
End Sub


Public Sub Nueva_Empresa()
   Txtempresa.Text = ""
   LblIdEmpresa(6).Caption = ""
   
   BtnEmpAcepta.Visible = True
   BtnEmpDeshabilita.Visible = False
   BtnEmpAnula.Visible = False
   SSPnelEmpresa.Visible = False
   
   SSPnelEmpresa.Visible = True
End Sub

Private Sub SSCommand4_Click()
SSPnelEmpresa.Visible = False
End Sub
Private Sub Graba_Nueva_Empresa(Accion As Integer)

Dim Sql As String
Dim Rq As ADODB.Recordset

If Accion = 1 Then
   Sql = "SELECT Nombrecia from Pla_Trab_Otras_Empresas_Cias where cia='" & wcia & "' and placod='" & Txtcodpla.Text & "' and nombrecia='" & Trim(Txtempresa.Text & "") & "' and status<>'*'"
Else
   Sql = "select top 1 mes from Pla_Trab_Otras_Empresas_Meses where idplaotracia=" & LblIdEmpresa(6).Caption & " and status<>'*'"
End If

If fAbrRst(Rq, Sql) Then
   If Accion = 1 Then
      MsgBox "Empresa ya esta registrada", vbInformation
      Rq.Close: Set Rq = Nothing
      Exit Sub
   ElseIf Accion = 3 Then
      MsgBox "Trabajador tiene remuneraciones en la empresa, no se puede anular", vbInformation
      Rq.Close: Set Rq = Nothing
      Exit Sub
   End If
End If
Rq.Close: Set Rq = Nothing

Dim lIDEmpresa As Integer
lIDEmpresa = 0
If LblIdEmpresa(6).Caption <> "" Then lIDEmpresa = LblIdEmpresa(6).Caption

Sql$ = "Usp_Pla_Mant_Trab_Otras_Empresas '" & wcia & "','E','" & Txtcodpla.Text & "'," & Accion & ",'" & wuser & "','" & Trim(Txtempresa.Text) & "'," & lIDEmpresa & ",0,0,0,0,0"
cn.Execute Sql$

Call Carga_Empresas(Trim(Txtcodpla.Text & ""), Trim(Lblnombre.Caption & ""))
End Sub

Private Sub SSCommand5_Click()
SSPanelremunera.Visible = False
End Sub
Private Sub Graba_Remuneracion(Accion As Integer)
If SSPanelremunera.Visible = flase Then Exit Sub
If Not IsNumeric(TxtBasico.Text) Then
   MsgBox "Ingrese Correctamente Basico", vbInformation
   Exit Sub
End If
If Not IsNumeric(TxtRemuneracion.Text) Then
   MsgBox "Ingrese Correctamente Remuneracion", vbInformation
   Exit Sub
End If
If Not IsNumeric(TxtQuinta.Text) Then
   MsgBox "Ingrese Correctamente Quinta", vbInformation
   Exit Sub
End If


Sql$ = "Usp_Pla_Mant_Trab_Otras_Empresas '" & wcia & "','R','" & TxtCodTrabremuneracion.Text & "'," & Accion & ",'" & wuser & "',''," & TxtIdEmpRemuneracion.Text & "," & LblAyoRemuneracion.Caption & "," & LblMesNumero.Caption & "," & CCur(TxtBasico.Text) & "," & CCur(TxtRemuneracion.Text) & "," & CCur(TxtQuinta.Text) & ""
cn.Execute Sql$

SSPanelremunera.Visible = False
Call Carga_Remuneraciones(Trim(TxtCodTrabremuneracion.Text & ""), Trim(LblNomRemuneracion.Caption & ""), Trim(TxtIdEmpRemuneracion.Text & ""), Trim(LblNomEmpRemuneracion.Caption & ""), LblAyoRemuneracion.Caption, LblMesNumero.Caption)
End Sub
Public Sub Nueva_Remuneracion()
If TxtCodRemun.Text = "" Then Exit Sub

Sql$ = "usp_pla_trab_otras_empresas_quinta '" & wcia & "','R','" & Trim(TxtCodRemun.Text & "") & "'," & Txtano.Text & "," & Cmbmes.ListIndex + 1 & "," & Trim(TxtIdEmpRem.Text & "") & ""
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
   MsgBox "Ya hay ingresos registrados para el periodo", vbInformation
   Rq.Close: Set Rq = Nothing
   Exit Sub
End If
Rq.Close: Set Rq = Nothing

TxtCodTrabremuneracion.Text = Trim(TxtCodRemun.Text & "")
LblNomRemuneracion.Caption = Trim(LblNomRemun.Caption & "")
TxtIdEmpRemuneracion.Text = Trim(TxtIdEmpRem.Text & "")
LblNomEmpRemuneracion.Caption = Trim(LblNomEmpRem.Caption & "")
LblAyoRemuneracion.Caption = Txtano.Text
LblMesRemuneracion.Caption = Cmbmes.Text
LblMesNumero.Caption = Cmbmes.ListIndex + 1
TxtBasico.Text = ""
TxtRemuneracion.Text = ""
TxtQuinta.Text = ""
SSPanelremunera.Visible = True
End Sub

Private Sub Txtano_Change()
If Cmbmes.ListIndex = -1 Then Exit Sub
Call Carga_Remuneraciones(Trim(TxtCodRemun.Text & ""), Trim(LblNomRemun.Caption & ""), Trim(TxtIdEmpRem.Text & ""), Trim(LblNomEmpRem.Caption & ""), Txtano.Text, Cmbmes.ListIndex + 1)
End Sub

