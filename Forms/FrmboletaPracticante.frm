VERSION 5.00
Begin VB.Form FrmboletaPracticante 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Ingreso de Boletas «"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   17070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtasigfam 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         MaxLength       =   8
         TabIndex        =   78
         Top             =   1320
         Width           =   855
      End
      Begin VB.PictureBox PnlActivos 
         BackColor       =   &H00D8E9EC&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   4995
         TabIndex        =   65
         Top             =   1680
         Width           =   5055
         Begin VB.TextBox TxtAct4 
            Height          =   285
            Left            =   4080
            TabIndex        =   72
            Top             =   30
            Width           =   855
         End
         Begin VB.TextBox TxtAct3 
            Height          =   285
            Left            =   3160
            TabIndex        =   70
            Top             =   30
            Width           =   855
         End
         Begin VB.TextBox TxtHor4 
            Height          =   285
            Left            =   4080
            TabIndex        =   73
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtHor1 
            Height          =   285
            Left            =   1440
            TabIndex        =   67
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtHor2 
            Height          =   285
            Left            =   2325
            TabIndex        =   69
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtHor3 
            Height          =   285
            Left            =   3165
            TabIndex        =   71
            Top             =   330
            Width           =   855
         End
         Begin VB.TextBox TxtAct2 
            Height          =   285
            Left            =   2320
            TabIndex        =   68
            Top             =   30
            Width           =   855
         End
         Begin VB.TextBox TxtAct1 
            Height          =   285
            Left            =   1440
            TabIndex        =   66
            Top             =   30
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Horas por Equipo"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Activos Asignados"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   45
            Width           =   1455
         End
      End
      Begin VB.PictureBox BtnVerCalculo 
         Height          =   420
         Left            =   5880
         ScaleHeight     =   360
         ScaleWidth      =   1635
         TabIndex        =   59
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox Txtvacaf 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3120
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   3
         Top             =   1365
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Txtvacai 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1000
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   2
         Top             =   1360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox Cmbturno 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Width           =   1455
      End
      Begin VB.PictureBox Txtvaca 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   7800
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox TxtFecCese 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1005
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   49
         Top             =   960
         Width           =   1095
      End
      Begin VB.PictureBox cbo_TipMotFinPer 
         BackColor       =   &H80000005&
         Height          =   315
         Left            =   4200
         ScaleHeight     =   255
         ScaleWidth      =   4515
         TabIndex        =   54
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox Txtcodpla 
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   0
         Top             =   120
         Width           =   1095
      End
      Begin VB.PictureBox LblCodTrabSunat 
         BackColor       =   &H80000010&
         Height          =   135
         Left            =   8400
         ScaleHeight     =   75
         ScaleWidth      =   315
         TabIndex        =   45
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000010&
         Height          =   375
         Left            =   5280
         TabIndex        =   79
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Fecha Periodo Vac."
         Height          =   195
         Left            =   6360
         TabIndex        =   17
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label LblFechaIngreso 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   5040
         TabIndex        =   76
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label LblPlanta 
         Height          =   255
         Left            =   6960
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LbLSemBolVac 
         Height          =   255
         Left            =   8400
         TabIndex        =   64
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000010&
         Caption         =   "Motivo Fin de Periodo"
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
         Left            =   2160
         TabIndex        =   53
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Label LblPensiones 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   52
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000010&
         Caption         =   "Sistema de Pensiones"
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
         Left            =   2160
         TabIndex        =   51
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Basico"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblAfpTipoComision 
         BackColor       =   &H80000010&
         Height          =   135
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   375
      End
      Begin VB.Label LblCese 
         BackColor       =   &H80000010&
         Caption         =   "F. Cese"
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
         TabIndex        =   47
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label LblID 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Height          =   75
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Lblbasico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   690
         TabIndex        =   28
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Lblnumafp 
         BackColor       =   &H80000010&
         Height          =   135
         Left            =   5640
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Lblcargo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6600
         TabIndex        =   26
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Lblvaca1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "F. Ret. Vac."
         Height          =   195
         Left            =   2160
         TabIndex        =   24
         Top             =   1360
         Width           =   855
      End
      Begin VB.Label Lblcodaux 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Height          =   75
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   165
         Width           =   420
      End
      Begin VB.Label Lblvaca2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "F. Ini. Vaca."
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Top             =   165
         Width           =   765
      End
      Begin VB.Label Lblcodafp 
         BackColor       =   &H80000010&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Lbltope 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5640
         TabIndex        =   25
         Top             =   240
         Width           =   45
      End
      Begin VB.Label LblFingreso 
         BackColor       =   &H80000010&
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   4200
         TabIndex        =   15
         Top             =   150
         Width           =   4575
      End
   End
   Begin VB.Frame frmImportar 
      Height          =   8775
      Left            =   9120
      TabIndex        =   32
      Top             =   120
      Width           =   7935
      Begin VB.ListBox LstObs 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   120
         TabIndex        =   44
         Top             =   7560
         Width           =   7605
      End
      Begin VB.CommandButton cmdImportar 
         Appearance      =   0  'Flat
         Caption         =   "Importar"
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
         Left            =   5280
         TabIndex        =   42
         Top             =   7080
         Width           =   2355
      End
      Begin VB.Frame Frame7 
         Caption         =   "Contenido del Archivo a importar"
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
         Height          =   5250
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   7530
         Begin VB.PictureBox DGrd 
            Height          =   4905
            Left            =   120
            ScaleHeight     =   4845
            ScaleWidth      =   7245
            TabIndex        =   41
            Top             =   240
            Width           =   7305
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Origen del Archivo a importar"
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
         Height          =   1155
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   7605
         Begin VB.TextBox Txtarchivos 
            BackColor       =   &H8000000A&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   705
            Width           =   5370
         End
         Begin VB.TextBox TxtRango 
            Height          =   315
            Left            =   3000
            TabIndex        =   34
            Text            =   "A6:U95"
            Top             =   330
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.PictureBox cmdVerArchivo 
            Height          =   495
            Left            =   6840
            ScaleHeight     =   435
            ScaleWidth      =   555
            TabIndex        =   36
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Ubicación"
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
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   39
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Rango de Datos Hoja de Excel"
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
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   2745
         End
         Begin VB.Label Label10 
            Caption         =   "Ejm: A1:G45"
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.PictureBox BarraImporta 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   4935
         TabIndex        =   43
         Top             =   7080
         Width           =   4965
      End
      Begin VB.PictureBox Box 
         Height          =   480
         Left            =   1800
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   80
         Top             =   7200
         Width           =   1200
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Descuentos Adicionales"
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
      ForeColor       =   &H00404040&
      Height          =   2640
      Left            =   4680
      TabIndex        =   11
      Top             =   6615
      Width           =   4335
      Begin VB.PictureBox DgrdDesAdic 
         Appearance      =   0  'Flat
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2265
         ScaleWidth      =   4065
         TabIndex        =   13
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pagos Adicionales"
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
      ForeColor       =   &H00404040&
      Height          =   4335
      Left            =   4680
      TabIndex        =   10
      Top             =   2280
      Width           =   4335
      Begin VB.PictureBox DgrdPagAdic 
         Appearance      =   0  'Flat
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4305
         ScaleWidth      =   4065
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "     Saldo Cta.Cte.                   S/."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   4680
         Width           =   2655
      End
      Begin VB.Label Lblctacte 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2580
         TabIndex        =   62
         Top             =   4680
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Centro de Costo"
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
      ForeColor       =   &H00404040&
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   7600
      Width           =   4335
      Begin VB.ListBox LstCcosto 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   435
         TabIndex        =   21
         Top             =   675
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox Dgrdccosto 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   4065
         TabIndex        =   12
         Top             =   240
         Width           =   4095
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
         TabIndex        =   18
         Top             =   1740
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Constantes de Calculo"
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
      ForeColor       =   &H00404040&
      Height          =   5295
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   4575
      Begin VB.PictureBox Dgrdhoras 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5505
         ScaleWidth      =   4185
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.PictureBox Panelprogress 
      BackColor       =   &H00C8D0D4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   1680
      ScaleHeight     =   675
      ScaleWidth      =   5595
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   5655
      Begin VB.PictureBox Barra 
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   5355
         TabIndex        =   30
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.PictureBox ProgressBar1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   4935
      TabIndex        =   31
      Top             =   3720
      Width           =   4965
   End
   Begin VB.PictureBox PnlPreView 
      BackColor       =   &H00C0C0C0&
      Height          =   7575
      Left            =   120
      ScaleHeight     =   7515
      ScaleWidth      =   8835
      TabIndex        =   55
      Top             =   1680
      Visible         =   0   'False
      Width           =   8895
      Begin VB.PictureBox LstIngresos 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   240
         ScaleHeight     =   6075
         ScaleWidth      =   4635
         TabIndex        =   56
         Top             =   240
         Width           =   4695
      End
      Begin VB.PictureBox LstDeducciones 
         BackColor       =   &H80000013&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   5040
         ScaleHeight     =   3555
         ScaleWidth      =   3555
         TabIndex        =   57
         Top             =   240
         Width           =   3615
      End
      Begin VB.PictureBox LstAportaciones 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   5040
         ScaleHeight     =   2355
         ScaleWidth      =   3555
         TabIndex        =   58
         Top             =   3960
         Width           =   3615
      End
      Begin VB.PictureBox SSCommand1 
         Height          =   420
         Left            =   240
         ScaleHeight     =   360
         ScaleWidth      =   8355
         TabIndex        =   61
         Top             =   6945
         Width           =   8415
      End
      Begin VB.Label LstNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   6480
         Width           =   8415
      End
   End
End
Attribute VB_Name = "FrmboletaPracticante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
