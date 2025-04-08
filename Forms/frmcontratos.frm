VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{92F3EE28-0DF4-4CB0-AAA5-E6578B12929B}#1.0#0"; "cbofacil.ocx"
Begin VB.Form frmcontratos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contratos"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   17730
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   0
      TabIndex        =   57
      Top             =   1320
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin Threed.SSPanel SpnlFechaContra 
      Height          =   1335
      Left            =   13200
      TabIndex        =   47
      Top             =   1320
      Visible         =   0   'False
      Width           =   4455
      _Version        =   65536
      _ExtentX        =   7858
      _ExtentY        =   2355
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
      Begin MSComCtl2.DTPicker DtFecContrato 
         Height          =   315
         Left            =   960
         TabIndex        =   48
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   12615808
         CalendarTitleForeColor=   16777215
         Format          =   121372673
         CurrentDate     =   37616
      End
      Begin MSForms.CommandButton CommandButton6 
         DragIcon        =   "frmcontratos.frx":0000
         Height          =   435
         Left            =   3960
         TabIndex        =   51
         Top             =   0
         Width           =   495
         ForeColor       =   16777215
         BackColor       =   0
         PicturePosition =   327683
         Size            =   "873;767"
         Picture         =   "frmcontratos.frx":058A
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Firma de Contrato"
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
         Left            =   480
         TabIndex        =   50
         Top             =   360
         Width           =   2370
      End
      Begin MSForms.CommandButton CommandButton5 
         DragIcon        =   "frmcontratos.frx":09DC
         Height          =   435
         Left            =   2640
         TabIndex        =   49
         Top             =   720
         Width           =   1575
         ForeColor       =   16777215
         BackColor       =   4210752
         Caption         =   "  Aceptar"
         PicturePosition =   327683
         Size            =   "2778;767"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.CheckBox ChkAll 
      Height          =   255
      Left            =   1320
      TabIndex        =   29
      Top             =   1560
      Width           =   255
   End
   Begin Threed.SSPanel SpnlModify 
      Height          =   3855
      Left            =   4440
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   9375
      _Version        =   65536
      _ExtentX        =   16536
      _ExtentY        =   6800
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
      BevelWidth      =   5
      BorderWidth     =   8
      BevelOuter      =   1
      BevelInner      =   2
      Begin VB.ComboBox CmbCargo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmcontratos.frx":0F66
         Left            =   1560
         List            =   "frmcontratos.frx":0F68
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1320
         Width           =   5835
      End
      Begin VB.TextBox TxtcodSuplencia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton OptFecha 
         BackColor       =   &H8000000C&
         Caption         =   "Fecha Final"
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
         Left            =   480
         TabIndex        =   42
         Top             =   2880
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton OptDias 
         BackColor       =   &H8000000C&
         Caption         =   "Dias"
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
         Left            =   480
         TabIndex        =   41
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox ChkAllDias 
         BackColor       =   &H8000000B&
         Caption         =   "Aplicar a Todos"
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
         Left            =   7200
         TabIndex        =   35
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox TxtDias 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   2760
         Visible         =   0   'False
         Width           =   615
      End
      Begin CboFacil.cbo_facil cbo_tipocont2 
         Height          =   315
         Left            =   1560
         TabIndex        =   20
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
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
      Begin MSComCtl2.DTPicker FecFin 
         Height          =   315
         Left            =   4440
         TabIndex        =   39
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   12615808
         CalendarTitleForeColor=   16777215
         Format          =   120848385
         CurrentDate     =   37616
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Trabjador a suplir "
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
         Left            =   600
         TabIndex        =   56
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label LblBasico 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   255
         Left            =   7560
         TabIndex        =   55
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   52
         Top             =   2055
         Width           =   7335
      End
      Begin MSForms.CommandButton CommandButton4 
         DragIcon        =   "frmcontratos.frx":0F6A
         Height          =   435
         Left            =   5880
         TabIndex        =   46
         Top             =   3240
         Width           =   1215
         ForeColor       =   0
         BackColor       =   8438015
         Caption         =   "  Calcular"
         PicturePosition =   327683
         Size            =   "2143;767"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label LblTiempo 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   600
         TabIndex        =   45
         Top             =   3240
         Width           =   5175
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
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
         Left            =   2880
         TabIndex        =   44
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label LblFecInicio 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   2805
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Termino"
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
         Left            =   4440
         TabIndex        =   40
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
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
         Left            =   600
         TabIndex        =   28
         Top             =   1365
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
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
         Left            =   2040
         TabIndex        =   26
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblCodPla 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   480
         Width           =   855
      End
      Begin VB.Label LblNombre 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   480
         Width           =   7335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basico"
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
         Left            =   7920
         TabIndex        =   23
         Top             =   990
         Width           =   585
      End
      Begin MSForms.CommandButton CommandButton1 
         DragIcon        =   "frmcontratos.frx":14F4
         Height          =   435
         Left            =   7320
         TabIndex        =   22
         Top             =   2760
         Width           =   1575
         ForeColor       =   16777215
         BackColor       =   4210752
         Caption         =   "  Aceptar"
         PicturePosition =   327683
         Size            =   "2778;767"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton2 
         DragIcon        =   "frmcontratos.frx":1A7E
         Height          =   435
         Left            =   8640
         TabIndex        =   21
         Top             =   120
         Width           =   615
         ForeColor       =   16777215
         BackColor       =   -2147483637
         PicturePosition =   327683
         Size            =   "1085;767"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   17295
      Begin VB.ComboBox CmbMotivo 
         Height          =   315
         ItemData        =   "frmcontratos.frx":2008
         Left            =   4920
         List            =   "frmcontratos.frx":2015
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton OptNuevos 
         BackColor       =   &H8000000B&
         Caption         =   "Nuevos - Ingresaron en "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   32
         Top             =   960
         Width           =   2640
      End
      Begin VB.ComboBox CmbCCosto 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   555
         Width           =   5415
      End
      Begin VB.Frame FrmPeriodo 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1160
         Left            =   9960
         TabIndex        =   7
         Top             =   80
         Width           =   2655
         Begin VB.TextBox txtano 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   660
            TabIndex        =   9
            Top             =   480
            Width           =   1005
         End
         Begin VB.ComboBox cbomes 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmcontratos.frx":2034
            Left            =   660
            List            =   "frmcontratos.frx":205C
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   120
            Width           =   1875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
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
            Height          =   210
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   330
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
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
            Height          =   210
            Left            =   120
            TabIndex        =   10
            Top             =   210
            Width           =   360
         End
      End
      Begin VB.OptionButton OptVence 
         BackColor       =   &H8000000B&
         Caption         =   "Con Vencimiento en"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   6
         Top             =   660
         Width           =   3240
      End
      Begin VB.OptionButton OptVigentes 
         BackColor       =   &H8000000B&
         Caption         =   "Vigentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   5
         Top             =   75
         Width           =   3240
      End
      Begin VB.OptionButton OptVencidos 
         BackColor       =   &H8000000B&
         Caption         =   "Vencidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         TabIndex        =   4
         Top             =   360
         Width           =   3240
      End
      Begin VB.ComboBox CmbTipoTrabajador 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   5415
      End
      Begin CboFacil.cbo_facil cbo_tipocont 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   900
         Width           =   5415
         _ExtentX        =   9551
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
      Begin MSForms.CommandButton CommandButton3 
         Height          =   1140
         Left            =   15840
         TabIndex        =   37
         Top             =   75
         Width           =   1335
         ForeColor       =   16777215
         BackColor       =   16761024
         Caption         =   "Exportar"
         PicturePosition =   327683
         Size            =   "2355;2011"
         Picture         =   "frmcontratos.frx":20C5
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
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
         Height          =   210
         Left            =   4320
         TabIndex        =   34
         Top             =   195
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Costo"
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
         Left            =   0
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Contrato"
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
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   945
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Trabajador "
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
         Height          =   210
         Left            =   60
         TabIndex        =   15
         Top             =   195
         Width           =   1335
      End
      Begin MSForms.CommandButton BtnGenera 
         Height          =   555
         Left            =   13920
         TabIndex        =   14
         Top             =   675
         Visible         =   0   'False
         Width           =   1815
         ForeColor       =   16777215
         BackColor       =   -2147483636
         Caption         =   "Generar Contratos"
         PicturePosition =   327683
         Size            =   "3201;979"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton BtnImprime 
         Height          =   555
         Left            =   12720
         TabIndex        =   13
         Top             =   0
         Width           =   3015
         ForeColor       =   16777215
         BackColor       =   -2147483636
         Caption         =   "  Imprimir Contratos"
         PicturePosition =   327683
         Size            =   "5318;979"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton btnCancel 
         Height          =   555
         Left            =   12720
         TabIndex        =   38
         Top             =   675
         Visible         =   0   'False
         Width           =   1215
         ForeColor       =   16777215
         BackColor       =   16761024
         Caption         =   "Cancelar"
         PicturePosition =   327683
         Size            =   "2143;979"
         Picture         =   "frmcontratos.frx":23DF
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton BtnOk 
         Height          =   555
         Left            =   13440
         TabIndex        =   36
         Top             =   675
         Visible         =   0   'False
         Width           =   1215
         ForeColor       =   16777215
         BackColor       =   16761024
         Caption         =   "Aplicar"
         PicturePosition =   327683
         Size            =   "2143;979"
         Picture         =   "frmcontratos.frx":2A59
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Frame frmcontratos 
      Caption         =   "          Todos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   17580
      Begin VB.CheckBox CheckSel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1200
         TabIndex        =   17
         Top             =   480
         Width           =   225
      End
      Begin MSDataGridLib.DataGrid DgrdContrato 
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   17385
         _ExtentX        =   30665
         _ExtentY        =   11245
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   -2147483624
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
         ColumnCount     =   24
         BeginProperty Column00 
            DataField       =   "placod"
            Caption         =   "Código"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Sel"
            Caption         =   "Sel"
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
         BeginProperty Column03 
            DataField       =   "Tipo_contra"
            Caption         =   "Tipo Contrato"
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
            DataField       =   "Dias"
            Caption         =   "Días"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "fec_ini"
            Caption         =   "F. Inicio"
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
            DataField       =   "fec_fin"
            Caption         =   "Fec. Termino"
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
            DataField       =   "Importe"
            Caption         =   "Basico"
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
            DataField       =   "Cargo"
            Caption         =   "Cargo"
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
            DataField       =   "cod_tip_contrato"
            Caption         =   "cod_tip_contrato"
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
            DataField       =   "codcargo"
            Caption         =   "codcargo"
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
            DataField       =   "num_contrato"
            Caption         =   "num_contrato"
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
            DataField       =   "contrato_ori"
            Caption         =   "contrato_ori"
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
            DataField       =   "cargo_Ori"
            Caption         =   "cargo_Ori"
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
            DataField       =   "empresa"
            Caption         =   "Empresa"
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
            DataField       =   "motivo"
            Caption         =   "Motivo"
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
            DataField       =   "fingreso"
            Caption         =   "fingreso"
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
         BeginProperty Column17 
            DataField       =   "planta"
            Caption         =   "planta"
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
         BeginProperty Column18 
            DataField       =   "fec_ininew"
            Caption         =   "Nuevo_Inicio"
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
         BeginProperty Column19 
            DataField       =   "Fec_Finnew"
            Caption         =   "Nuevo_Fin"
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
         BeginProperty Column20 
            DataField       =   "codcargoact"
            Caption         =   "cargoact"
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
         BeginProperty Column21 
            DataField       =   "ContratoInicialTrab_FecIni"
            Caption         =   "ContratoInicialTrab_FecIni"
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
         BeginProperty Column22 
            DataField       =   "ContratoInicialTrab_FecFin"
            Caption         =   "ContratoInicialTrab_FecFin"
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
         BeginProperty Column23 
            DataField       =   "TxtPeriodoContratoInicialTrab"
            Caption         =   "TxtPeriodoContratoInicialTrab"
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
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   420.095
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   3734.929
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   3795.024
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               Locked          =   -1  'True
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               Object.Visible         =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column08 
               Locked          =   -1  'True
               ColumnWidth     =   3044.977
            EndProperty
            BeginProperty Column09 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column10 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column11 
               Locked          =   -1  'True
               Object.Visible         =   0   'False
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
            BeginProperty Column17 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column20 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column21 
            EndProperty
            BeginProperty Column22 
            EndProperty
            BeginProperty Column23 
               Locked          =   -1  'True
               ColumnWidth     =   5864.882
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoContratos 
         Height          =   375
         Left            =   480
         Top             =   4680
         Visible         =   0   'False
         Width           =   16140
         _ExtentX        =   28469
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
         Caption         =   ""
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
Attribute VB_Name = "frmcontratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsContratos As New Recordset
Dim lCarga As Boolean
Dim lProcesa As Boolean
Dim lCia As String

Public Sub Eliminar_Contrato()

If Not OptVigentes.Value Then Exit Sub

If rsContratos.RecordCount <= 0 Then
    MsgBox "No hay Registros a Eliminar", vbCritical, Me.Caption
    Exit Sub
End If

Dim rsContratoTemp As New ADODB.Recordset
Set rsContratoTemp = rsContratos.Clone

rsContratoTemp.Filter = "SEL='S'"
If rsContratoTemp.RecordCount <= 0 Then
    MsgBox "Seleccione al menos un Contrato", vbCritical, Me.Caption
    Exit Sub
End If

Dim NoEliminados As String
Screen.MousePointer = vbArrowHourglass

rsContratoTemp.MoveFirst
Do While Not rsContratoTemp.EOF
    Sql = "usp_Pla_EliminarContrato '" & wcia & "','" & Trim(rsContratoTemp!num_contrato)
    Sql = Sql & "','" & Trim(rsContratoTemp!PlaCod) & "','" & wuser & "'"
    Dim rs2 As New ADODB.Recordset
    If (fAbrRst(rs2, Sql)) Then
        If rs2!Campo <> "ELIMINO" Then
            NoEliminados = NoEliminados & Chr(13) & Trim(rsContratoTemp!PlaCod) & " - " & Trim(rsContratoTemp!nombre)
        End If
    Else
       NoEliminados = NoEliminados & Chr(13) & Trim(rsContratoTemp!PlaCod) & " - " & Trim(rsContratoTemp!nombre)
    End If
    rs2.Close
    Set rs2 = Nothing
    rsContratoTemp.MoveNext
Loop

If NoEliminados <> "" Then
    NoEliminados = "No se eliminaron los contratos de los siguientes trabajadores " & NoEliminados
    MsgBox NoEliminados, vbCritical, Me.Caption
Else
    MsgBox "Se eliminaron Satisfactoriamente todos los contratos", vbInformation, Me.Caption
End If
Procesa
Screen.MousePointer = vbDefault
End Sub
Private Sub btnCancel_Click()
BtnGenera.Visible = False
btnCancel.Visible = False

DgrdContrato.BackColor = &H80000018

BtnOk.Enabled = True
BtnOk.Visible = True
BtnImprime.Visible = True
Call Procesa
End Sub

Private Sub BtnGenera_Click()
Dim rs2 As ADODB.Recordset
Dim f1 As String
Dim f2 As String
Dim lCount As Integer
On Error GoTo ErrorTrans
Dim NroTrans As Integer

NroTrans = 0
lCount = 0
If rsContratos.RecordCount > 0 Then rsContratos.MoveFirst
Do While Not rsContratos.EOF
   If Trim(rsContratos!SEL & "") = "S" Then
      f1 = rsContratos!fec_ini
      
      Sql$ = "set dateformat dmy "
      Sql$ = Sql$ & "select case  when fec_ini>='" & f1 + " 00:00:00" & "' then 'I' else 'F' end as Clave "
      Sql$ = Sql & "from placontrato where codcia='" & Trim(rsContratos!empresa) & "' and placod='" & Trim(rsContratos!PlaCod) & "' "
      Sql$ = Sql$ & "and status<>'*' and (fec_ini>='" & f1 + " 00:00:00" & "' or fec_fin>='" & f1 + " 00:00:00" & "')"
      If (fAbrRst(rs2, Sql$)) Then
         If rs2!Clave = "I" Then
            MsgBox "Existe un contrato con fecha inicial mayor o igual a nuevo contrato" & Chr(13) & Trim(rsContratos!nombre & ""), vbInformation
         Else
            MsgBox "Existe un contrato con fecha Final mayor o igual a nuevo contrato" & Chr(13) & Trim(rsContratos!nombre & ""), vbInformation
         End If
         rs2.Close
         Exit Sub
      End If
      lCount = lCount + 1
   End If
   rsContratos.MoveNext
Loop
If lCount <= 0 Then MsgBox "Debe seleccionar contratos", vbInformation: Exit Sub
rs2.Close

Screen.MousePointer = vbArrowHourglass
cn.BeginTrans
NroTrans = 1
cn.Execute "delete from placontratotmp"
If rsContratos.RecordCount > 0 Then rsContratos.MoveFirst
Do While Not rsContratos.EOF
   If Trim(rsContratos!SEL & "") = "S" Then
      Sql$ = "uSp_insertar_Contrato '" & Trim(rsContratos!empresa) & "','','" & Trim(rsContratos!PlaCod) & "','" & Trim(rsContratos!fec_ini) & "','" & Trim(rsContratos!fec_fin) & "','" & rsContratos!codcargo & "',"
      Sql$ = Sql$ & "'" & rsContratos!importe & "','" & wuser & "','','','" & rsContratos!cod_tip_contrato & "','" & rsContratos!placodsuple & "'"
      If (fAbrRst(rs2, Sql$)) Then
         If Trim(rs2(0) & "") = "" Then
            MsgBox "Contrato del trabajador => " & Trim(rsContratos!nombre & "") & Chr(13) & "No se Generó", vbInformation
         Else
            Sql$ = "uSp_Carga_ContratosTmp  '" & wcia & "','" & Trim(rs2(0) & "") & "','" & Trim(rsContratos!fec_ini) & "','" & wuser & "'"
            cn.Execute Sql$
            'If Trim(rsContratos!cod_tip_contrato & "") <> Trim(rsContratos!contrato_ori & "") Or Trim(rsContratos!codcargo & "") <> Trim(rsContratos!cargo_Ori & "") Then
            '   Sql$ = "update planillas set cargo='" & Trim(rsContratos!codcargo & "") & "',tipo_contrato='" & Trim(rsContratos!cod_tip_contrato & "") & "' where cia='" & wcia & "' and placod='" & Trim(rsContratos!PlaCod & "") & "' and status<>'*'"
            '   cn.Execute (Sql$)
            'End If
         End If
         
      End If
   End If
   rsContratos.MoveNext
Loop

cn.CommitTrans
NroTrans = 0
Contratos_Generados_Excel
BtnGenera.Visible = False
Screen.MousePointer = vbDefault
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

MsgBox Err.Description, vbCritical, Me.Caption

End Sub

Private Sub BtnImprime_Click()
Dim rsContratoTemp As New ADODB.Recordset

If lCia = "*" Then MsgBox "Selecciones Empresa", vbInformation: Exit Sub
If Cmbtipotrabajador.ListIndex < 0 Then MsgBox "Selecciones Tipo de Trabajador", vbInformation: Exit Sub
If cbo_tipocont.ListIndex < 0 Or cbo_tipocont.Text = "NINGUNO" Then MsgBox "Selecciones Tipo de Contrato", vbInformation: Exit Sub
'If CmbMotivo.Text = "TODOS" Then MsgBox "Selecciones Motivo", vbInformation: Exit Sub
Set rsContratoTemp = rsContratos.Clone

rsContratoTemp.Filter = "SEL='S'"
If rsContratoTemp.RecordCount <= 0 Then
    MsgBox "Seleccione al menos un Contrato", vbInformation
    rsContratos.Filter = adFilterNone
    Exit Sub
End If

If rsContratoTemp.RecordCount > 0 Then rsContratoTemp.MoveFirst

Dim lFi As String
lFi = ""
Do While Not rsContratoTemp.EOF
   If Trim(rsContratoTemp!SEL & "") = "S" Then
      If lFi = "" Then lFi = Trim(rsContratoTemp!fec_ini & "")
      If lFi <> Trim(rsContratoTemp!fec_ini & "") Then
         MsgBox "Solo puede seleccionar trabajadores con la misma fecha de inicio", vbInformation
         Exit Sub
      End If
   End If
   rsContratoTemp.MoveNext
Loop

Frame1.Enabled = False
frmcontratos.Enabled = False
DtFecContrato.Value = Format(DateAdd("d", -1, lFi), "dd/mm/yyyy")
SpnlFechaContra.Visible = True
End Sub
Private Sub Imprime()
Dim rsContratoTemp As New ADODB.Recordset
Set rsContratoTemp = rsContratos.Clone

cn.Execute "delete from placontratotmp"
If rsContratoTemp.RecordCount > 0 Then rsContratoTemp.MoveFirst

Do While Not rsContratoTemp.EOF
   If Trim(rsContratoTemp!SEL & "") = "S" Then
      'Sql$ = "uSp_Carga_ContratosTmp  '" & lCia & "','" & rsContratoTemp!num_contrato & "','" & Format(DtFecContrato.Value, "dd/mm/yyyy") & "'"
        Sql$ = "uSp_Carga_ContratosTmp  '" & lCia & "','" & rsContratoTemp!num_contrato & "','" & Format(DtFecContrato.Value, "dd/mm/yyyy") & "','" & wuser & "'"
      cn.Execute Sql$
   End If
   rsContratoTemp.MoveNext
Loop

rsContratoTemp.Filter = adFilterNone

Imprime_Contratos

End Sub
Private Sub BtnOk_Click()
If rsContratos.RecordCount <= 0 Then Exit Sub
rsContratos.MoveFirst
Dim f1 As String
Dim f2 As String
Dim rs2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim lTipContr As String
Dim rsContratoTemp As New ADODB.Recordset
Set rsContratoTemp = rsContratos.Clone
rsContratoTemp.Filter = "SEL='S'"
If rsContratoTemp.RecordCount <= 0 Then
    MsgBox "Seleccione al menos un contrato", vbCritical, Me.Caption
    rsContratoTemp.Close
    Set rsContratoTemp = Nothing
    Exit Sub
End If

rsContratoTemp.Close
Set rsContratoTemp = Nothing


Do While Not rsContratos.EOF
   If Trim(rsContratos!SEL & "") <> "S" Or Trim(rsContratos!fec_iniNew & "") = "" Then
      rsContratos.Delete
   Else
      rsContratos!fec_ini = rsContratos!fec_iniNew
      rsContratos!fec_fin = rsContratos!fec_finNew
      
      Sql$ = "select p.cargo as codcargo,(select descrip from maestros_31 where ciamaestro=p.cia+'055' and cod_maestro3=p.cargo) as Cargo,p.fingreso," _
             & "(select sum(Case Factor_horas When 8 then importe*30 When 48 then importe*4 when 240 then importe else 0 End) from plaremunbase where " _
             & "Cia=p.cia and concepto in('01','02') and placod=p.placod and status<>'*' ) as Importe " _
             & "from planillas p " _
             & "where p.cia='" & Trim(rsContratos!empresa & "") & "' and p.placod='" & Trim(rsContratos!PlaCod & "") & "' and p.status<>'*'"
      
      If (fAbrRst(rs2, Sql$)) Then
         rsContratos!importe = rs2!importe
         rsContratos!Cargo = Trim(rs2!Cargo & "")
         rsContratos!codcargo = Trim(rs2!codcargo & "")
         f2 = Format(Day(rs2!fIngreso), "00") & "/" & Format(Month(rs2!fIngreso), "00") & "/" & Format(Year(rs2!fIngreso), "0000")
         lTipContr = Tipo_contrato(Trim(rsContratos!empresa & ""), f2, Format(DateAdd("d", Val(rsContratos!Dias), Format(rsContratos!fec_ini, "dd/mm/yyyy")), "dd/mm/yyyy"))
         If lTipContr = "**" Then
            rsContratos!SEL = ""
            MsgBox "Trabajador => " & Trim(rsContratos!PlaCod & "") & " Excede los 5 años" & Chr(13) & "No se generará su contrato"
         Else
            'rsContratos!cod_tip_contrato = lTipContr
            'Sql$ = "select descrip from maestros_2 where ciamaestro='01144' and cod_maestro2='" & lTipContr & "'"
            'If (fAbrRst(RS3, Sql$)) Then
            '   rsContratos!Tipo_contra = Trim(RS3!DESCRIP & "")
            'Else
            '   rsContratos!Tipo_contra = ""
            'End If
            'RS3.Close: Set RS3 = Nothing
         End If
      Else
         MsgBox "No Hay Datos para trabajador => " & Trim(rsContratos!PlaCod & "")
      End If
      If rs2.State = 1 Then rs2.Close: Set rs2 = Nothing
   End If
   rsContratos.MoveNext
Loop

CheckSel.Visible = False
BtnGenera.Visible = True
btnCancel.Visible = True
DgrdContrato.BackColor = &HC0FFC0
BtnOk.Enabled = False
BtnOk.Visible = False
BtnImprime.Visible = False
End Sub

Private Sub BtnOk_Click_Dias()
If rsContratos.RecordCount <= 0 Then Exit Sub
rsContratos.MoveFirst
Dim f1 As String
Dim f2 As String
Dim rs2 As ADODB.Recordset
Dim RS3 As ADODB.Recordset
Dim lTipContr As String
Dim rsContratoTemp As New ADODB.Recordset
Set rsContratoTemp = rsContratos.Clone
rsContratoTemp.Filter = "SEL='S'"
If rsContratoTemp.RecordCount <= 0 Then
    MsgBox "Seleccione al menos un contrato", vbCritical, Me.Caption
    rsContratoTemp.Close
    Set rsContratoTemp = Nothing
    Exit Sub
End If

rsContratoTemp.Close
Set rsContratoTemp = Nothing


Do While Not rsContratos.EOF
   If Trim(rsContratos!SEL & "") <> "S" Then
      rsContratos.Delete
   Else
     If OptNuevos.Value Then
        f1 = rsContratos!fec_fin
        rsContratos!fec_ini = rsContratos!fec_fin
     Else
        f1 = Format(DateAdd("d", 1, Format(rsContratos!fec_fin, "dd/mm/yyyy")), "dd/mm/yyyy")
        rsContratos!fec_ini = Format(DateAdd("d", 1, Format(rsContratos!fec_fin, "dd/mm/yyyy")), "dd/mm/yyyy")
     End If
     
     rsContratos!fec_fin = Format(DateAdd("d", Val(rsContratos!Dias), Format(f1, "dd/mm/yyyy")), "dd/mm/yyyy")
     Sql$ = "select r.importe,p.cargo as codcargo,(select descrip from maestros_31 where ciamaestro=p.cia+'055' and cod_maestro3=p.cargo) as Cargo,p.fingreso " _
          & "from planillas p,plaremunbase r " _
          & "where p.cia='" & Trim(rsContratos!empresa & "") & "' and p.placod='" & Trim(rsContratos!PlaCod & "") & "' and r.concepto='01' and p.status<>'*' and r.status<>'*' and p.cia=r.cia and p.placod=r.placod"
     If (fAbrRst(rs2, Sql$)) Then
        rsContratos!importe = rs2!importe
        rsContratos!Cargo = Trim(rs2!Cargo & "")
        rsContratos!codcargo = Trim(rs2!codcargo & "")
        f2 = Format(Day(rs2!fIngreso), "00") & "/" & Format(Month(rs2!fIngreso), "00") & "/" & Format(Year(rs2!fIngreso), "0000")
        lTipContr = Tipo_contrato(Trim(rsContratos!empresa & ""), f2, Format(DateAdd("d", Val(rsContratos!Dias), Format(f1, "dd/mm/yyyy")), "dd/mm/yyyy"))
        If lTipContr = "**" Then
           rsContratos!SEL = ""
           MsgBox "Trabajador => " & Trim(rsContratos!PlaCod & "") & " Excede los 5 años" & Chr(13) & "No se generará su contrato"
        Else
           rsContratos!cod_tip_contrato = lTipContr
           Sql$ = "select descrip from maestros_2 where ciamaestro='01144' and cod_maestro2='" & lTipContr & "'"
           If (fAbrRst(RS3, Sql$)) Then
              rsContratos!Tipo_contra = Trim(RS3!DESCRIP & "")
           Else
              rsContratos!Tipo_contra = ""
           End If
           RS3.Close: Set RS3 = Nothing
        End If
     Else
        MsgBox "No Hay Datos para trabajador => " & Trim(rsContratos!PlaCod & "")
     End If
     If rs2.State = 1 Then rs2.Close: Set rs2 = Nothing
   End If
   rsContratos.MoveNext
Loop

CheckSel.Visible = False
BtnGenera.Visible = True
btnCancel.Visible = True
DgrdContrato.BackColor = &HC0FFC0
BtnOk.Enabled = False
BtnOk.Visible = False
BtnImprime.Visible = False
End Sub

Private Sub cbo_tipocont_Click()
If lCarga Then Procesa
End Sub

Private Sub cbomes_Click()
If lCarga Then Procesa
End Sub

Private Sub CheckSel_Click()
If BtnGenera.Visible Then Exit Sub
If rsContratos.EOF And rsContratos.RecordCount > 0 Then rsContratos.MoveLast
If rsContratos.RecordCount > 0 Then
If CheckSel.Value = 1 Then
   rsContratos!SEL = "S"
Else
   rsContratos!SEL = ""
End If
End If
End Sub
Private Sub ChkAll_Click()
Dim mSel As String
If ChkAll.Value = 1 Then
   mSel = "S"
   frmcontratos.Caption = "          Ninguno"
Else
   mSel = ""
   frmcontratos.Caption = "            Todos"
End If
If rsContratos.RecordCount > 0 Then rsContratos.MoveFirst
Do While Not rsContratos.EOF
   rsContratos!SEL = mSel
   rsContratos.MoveNext
Loop
CheckSel.Value = ChkAll.Value
End Sub

Private Sub Cmbccosto_Click()
If lCarga Then Procesa
End Sub

Private Sub CmbMotivo_Click()
If lCarga Then Procesa
End Sub

Private Sub Cmbtipotrabajador_Click()
If lCarga Then Procesa
End Sub

Private Sub CommandButton1_Click()
If cbo_tipocont2.Text = "NINGUNO" Or cbo_tipocont2.Text = "" Then
   MsgBox "Seleccione Tipo de Contrato", vbInformation
   Exit Sub
End If

Calcula_Tiempo
If Not IsNumeric(TxtDias.Text) Then
   MsgBox "Ingrese Correctamente los Días", vbInformation
   Exit Sub
End If
If Val(TxtDias.Text) < 15 Then
   MsgBox "Ingrese Correctamente los Días", vbInformation
   Exit Sub
End If
Dim TC As String
'Dim CG As String
TC = Format(cbo_tipocont2.ReturnCodigo, "00")
'CG = fc_CodigoComboBox(CmbCargo, 3)

'rsContratos!codcargo = CG
rsContratos!Cargo = Cmbcargo.Text
rsContratos.Update
'añadir la variable del codigo del trabajador por suplencia
rsContratos!placodsuple = Trim(TxtcodSuplencia.Text & "")

If ChkAllDias.Value = 1 Then
   If rsContratos.RecordCount > 0 Then rsContratos.MoveFirst
   Do While Not rsContratos.EOF
      rsContratos!Dias = TxtDias.Text
      rsContratos!fec_finNew = Format(FecFin.Value, "dd/mm/yyyy")
      If OptNuevos.Value Then
         rsContratos!fec_iniNew = rsContratos!fec_fin
      Else
         rsContratos!fec_iniNew = Format(DateAdd("d", 1, Format(rsContratos!fec_fin, "dd/mm/yyyy")), "dd/mm/yyyy")
      End If
      rsContratos!cod_tip_contrato = Format(cbo_tipocont2.ReturnCodigo, "00")
      rsContratos!Tipo_contra = Trim(cbo_tipocont2.Text & "")
      rsContratos.MoveNext
   Loop
Else
   rsContratos!Dias = TxtDias.Text
   rsContratos!fec_finNew = Format(FecFin.Value, "dd/mm/yyyy")
   If OptNuevos.Value Then
      rsContratos!fec_iniNew = rsContratos!fec_fin
   Else
      rsContratos!fec_iniNew = Format(DateAdd("d", 1, Format(rsContratos!fec_fin, "dd/mm/yyyy")), "dd/mm/yyyy")
   End If
   rsContratos!cod_tip_contrato = Format(cbo_tipocont2.ReturnCodigo, "00")
   rsContratos!Tipo_contra = Trim(cbo_tipocont2.Text & "")
End If

'rsContratos!cod_tip_contrato = TC
'rsContratos!Tipo_contra = cbo_tipocont2.Text
SpnlModify.Visible = False
frmcontratos.Enabled = True
Frame1.Enabled = True
End Sub

Private Sub CommandButton2_Click()
SpnlModify.Visible = False
frmcontratos.Enabled = True
Frame1.Enabled = True
End Sub

Private Sub CommandButton3_Click()

Reporte
Exit Sub

'If Cmbtipotrabajador.ListIndex < 0 Then MsgBox "Selecciones Tipo de Trabajador", vbInformation: Exit Sub
''Procesa
'Dim Sql As String
'
'Dim lAno As Integer
'Dim lMEs As Integer
'lAno = 0
'lMEs = 0
'If OptVence.Value = True Then
'   lAno = Txtano.Text: lMEs = cbomes.ListIndex + 1
'End If
'
'Sql = "Usp_Ultimos_Contratos '" & wcia & "','" & fc_CodigoComboBox(Cmbtipotrabajador, 2) & "'," & lAno & "," & lMEs & ""
'cn.CursorLocation = adUseClient
'Set AdoContratos.Recordset = cn.Execute(Sql, 64)
'Call Exportar_Excel(AdoContratos.Recordset)
End Sub

Private Sub Calcula_Tiempo()
If FecFin.Enabled = False Then
   FecFin.Value = Format(DateAdd("d", Val(TxtDias.Text), LblFecInicio.Caption), "dd/mm/yyyy")
End If

If TxtDias.Enabled = False Then
   TxtDias.Text = DateDiff("d", LblFecInicio.Caption, FecFin.Value)
End If

LblTiempo.Caption = ""
'Sql$ = "set dateformat dmy Select dbo.EntreFechasAnoMesDia('" & LblFecInicio.Caption & "','" & FecFin.Value & "') set dateformat mdy"
Sql$ = "set dateformat dmy Select dbo.EntreFechasAnoMesDia_2022('" & LblFecInicio.Caption & "','" & FecFin.Value & "') set dateformat mdy"

'Sql$ = "set dateformat dmy Select dbo.EntreFechasAnoMesDia_2022('" & Format(LblFecInicio.Caption, "mm/dd/yyyy") & "','" & Format(FecFin.Value, "mm/dd/yyyy") & "') set dateformat mdy"

Dim rs2 As ADODB.Recordset
If (fAbrRst(rs2, Sql$)) Then
   LblTiempo.Caption = Trim(rs2(0) & "")
End If
rs2.Close: Set rs2 = Nothing
End Sub

Private Sub CommandButton4_Click()
Calcula_Tiempo
End Sub

Private Sub CommandButton5_Click()
Imprime
SpnlFechaContra.Visible = False
Frame1.Enabled = True
frmcontratos.Enabled = True
End Sub

Private Sub CommandButton6_Click()
SpnlFechaContra.Visible = False
Frame1.Enabled = True
frmcontratos.Enabled = True
End Sub

Private Sub DgrdContrato_DblClick()
If rsContratos.EOF Then Exit Sub
If OptVigentes.Value Then Exit Sub
If OptVencidos.Value Then Exit Sub

   LblCodPla.Caption = rsContratos!PlaCod
   Lblnombre.Caption = rsContratos!nombre
   TxtDias.Text = rsContratos!Dias
   Lblbasico.Caption = Format(rsContratos!importe, "###,###,###.00")
   
   If OptNuevos.Value Then
      LblFecInicio.Caption = rsContratos!fec_ini
      FecFin.Value = rsContratos!fec_ini
   Else
      LblFecInicio.Caption = Format(DateAdd("d", 1, Format(rsContratos!fec_fin, "dd/mm/yyyy")), "dd/mm/yyyy")
      FecFin.Value = Format(DateAdd("d", 1, Format(rsContratos!fec_fin, "dd/mm/yyyy")), "dd/mm/yyyy")
   End If
   Call rUbiIndCmbBox(Cmbcargo, rsContratos!codcargo, "000")
   'cbo_tipocont2.SetIndice Trim(rsContratos!cod_tip_contrato & "")
   cbo_tipocont2.ListIndex = -1
   ChkAllDias.Value = 0
   If BtnOk.Visible = True And BtnGenera.Visible = False Then
      Call rUbiIndCmbBox(Cmbcargo, rsContratos!codcargoact, "000")
      
      ChkAllDias.Enabled = True
      CommandButton1.Enabled = True
      If OptVence.Value = True Then
            'CmbCargo.Enabled = True
      Else
         Cmbcargo.Enabled = False
      End If
   Else
        ChkAllDias.Enabled = False
        CommandButton1.Enabled = False
        Cmbcargo.Enabled = False
   End If
   
   
   cbo_tipocont2.NameTab = "maestros_2"
   cbo_tipocont2.NameCod = "cod_maestro2"
   cbo_tipocont2.NameDesc = "descrip"
   
   If OptNuevos.Value Then
      cbo_tipocont2.Filtro = "RIGHT(ciamaestro,3)='144' and cod_maestro2 in('02','03','04','07','08','09','10','11','12','13','14','15') and status!='*' and rtrim(isnull(codsunat,''))<>''"
   Else
      'cbo_tipocont2.Filtro = "RIGHT(ciamaestro,3)='144' and cod_maestro2='05' and status!='*' and rtrim(isnull(codsunat,''))<>''"
      cbo_tipocont2.Filtro = "RIGHT(ciamaestro,3)='144' and cod_maestro2 in('02','03','04','05','07','08','09','10','11','12','13','14','15') and status!='*' and rtrim(isnull(codsunat,''))<>''"
   End If
   
   cbo_tipocont2.conexion = cn
   cbo_tipocont2.Execute
   
   If OptNuevos.Value Then cbo_tipocont2.ListIndex = -1
   
   SpnlModify.Visible = True
   frmcontratos.Enabled = False
   Frame1.Enabled = False
   Label14.Caption = ""

End Sub

Private Sub DgrdContrato_KeyDown(KeyCode As Integer, Shift As Integer)
'If BtnGenera.Visible Then Exit Sub
If rsContratos.EOF Then Exit Sub
'If BtnOk.Visible = False Then Exit Sub
If KeyCode = 13 Then

   LblCodPla.Caption = rsContratos!PlaCod
   Lblnombre.Caption = rsContratos!nombre
   TxtDias.Text = rsContratos!Dias
   Lblbasico.Caption = Format(rsContratos!importe, "###,###,###.00")
   Call rUbiIndCmbBox(Cmbcargo, rsContratos!codcargo, "00")
   cbo_tipocont2.SetIndice Trim(rsContratos!cod_tip_contrato & "")
   ChkAllDias.Value = 0
   If BtnOk.Visible = True And BtnGenera.Visible = False Then
        ChkAllDias.Enabled = True
        CommandButton1.Enabled = True
        If OptVence.Value = True Then
            Cmbcargo.Enabled = True
        Else
            Cmbcargo.Enabled = False
        End If
   Else
        ChkAllDias.Enabled = False
        CommandButton1.Enabled = False
        Cmbcargo.Enabled = False
   End If
   SpnlModify.Visible = True
   frmcontratos.Enabled = False
   Frame1.Enabled = False
   
End If
End Sub

Private Sub DgrdContrato_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If BtnGenera.Visible Then CheckSel.Visible = False: Exit Sub
If rsContratos.AbsolutePage > 0 Then
   If UCase(rsContratos!SEL) = "S" Then CheckSel.Value = 1 Else CheckSel.Value = 0
   'CheckSel.Left = DgrdContrato.Left + DgrdConcepto.Columns(1).Left + 250
   CheckSel.Top = DgrdContrato.Top + DgrdContrato.RowTop(DgrdContrato.Row) + 5
   CheckSel.Visible = True
   CheckSel.ZOrder 0
End If

End Sub

Private Sub DgrdContrato_Scroll(Cancel As Integer)
CheckSel.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = 17820
Me.Height = 8730
Me.Left = 0
Me.Top = 0

Sql$ = "select codigo,descripcion From Pla_ccostos where status<>'*' union Select '00','..TODOS' order by descripcion"
Call rCarCbo(CmbCcosto, Sql$, "C", "00")

lCarga = False
Crea_Rs
CmbMotivo.ListIndex = 0
Procesa

OptVigentes.Value = True
' TIPO DE CONTRATOS
cbo_tipocont.NameTab = "maestros_2"
cbo_tipocont.NameCod = "cod_maestro2"
cbo_tipocont.NameDesc = "descrip"
cbo_tipocont.Filtro = "RIGHT(ciamaestro,3)='144' and status!='*' and rtrim(isnull(codsunat,''))<>''"
cbo_tipocont.conexion = cn
cbo_tipocont.Execute

Sql$ = "SELECT COD_MAESTRO3,DESCRIP FROM MAESTROS_31 WHERE CIAMAESTRO = '" & wcia & "055" & "' AND STATUS != '*' ORDER BY DESCRIP"
Call rCarCbo(Cmbcargo, Sql$, "C", "00")

Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador, True)
lCarga = True
End Sub

Private Sub OptDias_Click()
If OptDias.Value Then
   FecFin.Enabled = False
   TxtDias.Enabled = True
Else
   FecFin.Enabled = True
   TxtDias.Enabled = False
End If

End Sub

Private Sub OptFecha_Click()
If OptFecha.Value Then
   FecFin.Enabled = True
   TxtDias.Enabled = False
Else
   FecFin.Enabled = False
   TxtDias.Enabled = True
End If
End Sub

Private Sub OptNuevos_Click()
If OptNuevos.Value Then
   FrmPeriodo.Visible = True
   cbomes.ListIndex = Month(Date) - 1
   Txtano.Text = Year(Date)
   If lCarga Then Procesa
   BtnOk.Visible = True
   BtnGenera.Visible = False
   btnCancel.Visible = False
   BtnImprime.Visible = False
End If
End Sub

Private Sub OptVence_Click()
If OptVence.Value Then
   FrmPeriodo.Visible = True
   cbomes.ListIndex = Month(Date) - 1
   Txtano.Text = Year(Date)
   If lCarga Then Procesa
   BtnOk.Visible = True
   BtnGenera.Visible = False
   btnCancel.Visible = False
   BtnImprime.Visible = True
End If
End Sub

Private Sub OptVencidos_Click()
If OptVencidos.Value Then FrmPeriodo.Visible = False: Txtano.Text = "": cbomes.ListIndex = -1: BtnOk.Visible = False: BtnGenera.Visible = False: btnCancel.Visible = False: BtnImprime.Visible = True
If lCarga Then Procesa
End Sub

Private Sub OptVigentes_Click()
If OptVigentes.Value Then FrmPeriodo.Visible = False: Txtano.Text = "": cbomes.ListIndex = -1: BtnOk.Visible = False: BtnGenera.Visible = False: btnCancel.Visible = False: BtnImprime.Visible = True
If lCarga Then Procesa
End Sub

Private Sub Crea_Rs()
    If rsContratos.State = 1 Then rsContratos.Close
    rsContratos.Fields.Append "placod", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "Nombre", adChar, 50, adFldIsNullable
    rsContratos.Fields.Append "Tipo_contra", adChar, 150, adFldIsNullable
    rsContratos.Fields.Append "Dias", adChar, 5, adFldIsNullable
    rsContratos.Fields.Append "fec_ini", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "fec_fin", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "importe", adCurrency, 18, adFldIsNullable
    rsContratos.Fields.Append "Cargo", adChar, 100, adFldIsNullable
    rsContratos.Fields.Append "cod_tip_contrato", adChar, 2, adFldIsNullable
    rsContratos.Fields.Append "num_contrato", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "codcargo", adChar, 3, adFldIsNullable
    rsContratos.Fields.Append "Sel", adChar, 1, adFldIsNullable
    rsContratos.Fields.Append "cargo_Ori", adChar, 3, adFldIsNullable
    rsContratos.Fields.Append "contrato_ori", adChar, 2, adFldIsNullable
    rsContratos.Fields.Append "empresa", adChar, 2, adFldIsNullable
    rsContratos.Fields.Append "motivo", adChar, 1, adFldIsNullable
    rsContratos.Fields.Append "fingreso", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "Planta", adChar, 20, adFldIsNullable
    rsContratos.Fields.Append "fec_iniNew", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "fec_finNew", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "codcargoact", adChar, 3, adFldIsNullable
    'Añado el trabajador a quien se le va a suplir por un periodo
    rsContratos.Fields.Append "placodsuple", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "CenCosto", adChar, 80, adFldIsNullable
    
    'add jcms 291122
    rsContratos.Fields.Append "ContratoInicialTrab_FecIni", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "ContratoInicialTrab_FecFin", adChar, 10, adFldIsNullable
    rsContratos.Fields.Append "TxtPeriodoContratoInicialTrab", adVarChar, 100, adFldIsNullable
    
    rsContratos.Open
    Set DgrdContrato.DataSource = rsContratos
End Sub
Private Sub Procesa()
Dim mTipoTrab As String
Dim mTipoCont As String
Dim mVigentes As String
Dim mVencidos As String
Dim mVencen As String
Dim Mes As Integer
Dim ano As Integer
BtnGenera.Visible = False
DgrdContrato.BackColor = &H80000018
BtnOk.Enabled = True
BtnImprime.Visible = True
lCia = "*"
lCia = wcia
Dim lCcosto As String
lCcosto = fc_CodigoComboBox(CmbCcosto, 2)
If lCcosto = "" Then lCcosto = "00"
If Trim(Cmbtipotrabajador.Text & "") = "" Then mTipoTrab = "*" Else mTipoTrab = fc_CodigoComboBox(Cmbtipotrabajador, 2)
mTipoCont = IIf(cbo_tipocont.ReturnCodigo = -1, "*", Format(cbo_tipocont.ReturnCodigo, "00"))
If OptVigentes.Value Then mVigentes = "S" Else mVigentes = "*"
If OptVencidos.Value Then mVencidos = "S" Else mVencidos = "*"
mVencen = "*"
Mes = Month(Date)
ano = Year(Date)
If OptVence.Value Then mVencen = "S"
If OptVence.Value Or OptNuevos.Value Then
   Mes = cbomes.ListIndex + 1
   ano = Val(Txtano.Text)
End If
Dim lNuevo As String
lNuevo = "N"
If OptNuevos.Value Then lNuevo = "S"
If mTipoCont = "00" Then mTipoCont = "*"
'Sql$ = "uSp_Carga_Contratos '" & lCia & "','" & mTipoTrab & "','" & mTipoCont & "','" & mVigentes & "','" & mVencidos & "','" & mVencen & "'," & Mes & "," & ano & ",'" & lNuevo & "','" & lCcosto & "'"

Sql$ = "exec uSp_Carga_Contratos '" & lCia & "','" & mTipoTrab & "','" & mTipoCont & "','" & mVigentes & "','" & mVencidos & "','" & mVencen & "'," & Mes & "," & ano & ",'" & lNuevo & "','" & lCcosto & "'"


cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
'If rsContratos.RecordCount > 0 Then
'   rsContratos.MoveFirst
'   Do While Not rsContratos.EOF
'      rsContratos.Delete
'      rsContratos.MoveNext
'   Loop
'End If
Crea_Rs
If Not rs.RecordCount > 0 Then Exit Sub
Dim lMot As String
lMot = Mid(CmbMotivo.Text, 1, 1)
rs.MoveFirst
Dim lCod As String
lCod = ""
Do While Not rs.EOF
   If lCod <> Trim(rs!PlaCod & "") Then
      lCod = Trim(rs!PlaCod & "")
      If (lMot = Trim(rs!motivo & "")) Or lMot = "T" Then
          rsContratos.AddNew
          rsContratos!SEL = ""
          rsContratos!PlaCod = Trim(rs!PlaCod & "")
          rsContratos!nombre = Trim(rs!nombre & "")
          rsContratos!Tipo_contra = Trim(rs!Tipo_contra & "")
          rsContratos!Dias = rs!Dias
          rsContratos!fec_ini = Format(Day(rs!fec_ini), "00") & "/" & Format(Month(rs!fec_ini), "00") & "/" & Format(Year(rs!fec_ini), "0000")
          rsContratos!fec_fin = Format(Day(rs!fec_fin), "00") & "/" & Format(Month(rs!fec_fin), "00") & "/" & Format(Year(rs!fec_fin), "0000")
          rsContratos!importe = rs!importe
          rsContratos!Cargo = Trim(rs!Cargo & "")
          rsContratos!cod_tip_contrato = Trim(rs!cod_tip_contrato & "")
          rsContratos!num_contrato = Trim(rs!num_contrato & "")
          rsContratos!codcargo = Trim(rs!codcargo & "")
          rsContratos!SEL = ""
          rsContratos!cargo_Ori = Trim(rs!codcargo & "")
          rsContratos!contrato_ori = Trim(rs!cod_tip_contrato & "")
          rsContratos!empresa = Trim(rs!CODCIA & "")
          rsContratos!motivo = Trim(rs!motivo & "")
          rsContratos!fIngreso = Format(Day(rs!fIngreso), "00") & "/" & Format(Month(rs!fIngreso), "00") & "/" & Format(Year(rs!fIngreso), "0000")
          rsContratos!Planta = Trim(rs!Planta & "")
          rsContratos!codcargoact = Trim(rs!codcargoact & "")
          'añado el codigo del trabajador a suplir
          rsContratos!placodsuple = Trim(rs!placodsuple & "")
          rsContratos!CENCOSTO = Trim(rs!CENCOSTO & "")
          
          If Not IsNull(rs!FechaIni_ContratoInicial) Then
            rsContratos!ContratoInicialTrab_FecIni = Format(Day(rs!FechaIni_ContratoInicial), "00") & "/" & Format(Month(rs!FechaIni_ContratoInicial), "00") & "/" & Format(Year(rs!FechaIni_ContratoInicial), "0000")
            rsContratos!ContratoInicialTrab_FecFin = Format(Day(rs!FechaFin_ContratoInicial), "00") & "/" & Format(Month(rs!FechaFin_ContratoInicial), "00") & "/" & Format(Year(rs!FechaFin_ContratoInicial), "0000")
            
            rsContratos!TxtPeriodoContratoInicialTrab = Trim(rs!TxtPeriodoContratoInicialTrab & "")
            
          End If
          
          
          rsContratos.Update
      End If
   End If
   rs.MoveNext
Loop
If rsContratos.RecordCount > 0 Then rsContratos.MoveFirst
If rs.State = 1 Then rs.Close

End Sub


Private Sub Txtano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(Txtano.Text) Then BtnImprime.SetFocus
End Sub

Private Sub Txtano_LostFocus()
If lCarga Then Procesa
End Sub
Private Sub Imprime_Contratos()
'Dim xlApp02 As Excel.Application
'Dim xlApp01  As Excel.Application
'Dim xlBook01 As Excel.Workbook
'Dim xlSheet01 As Excel.Worksheet
'Dim xlBook02 As Excel.Workbook
'Dim xlSheet02 As Excel.Worksheet




Dim xlApp02 As Object
Dim xlApp01  As Object
Dim xlBook01 As Object
Dim xlSheet01 As Object
Dim xlBook02 As Object
Dim xlSheet02 As Object




Dim xlApp_1  As Object
Dim xlApp_2  As Object

Dim March01 As String
Dim MarchExcel01 As String
Dim MarchExcel02 As String
Dim sOrigen01 As String
Dim sOrigen02 As String
Dim sDestino01 As String
Dim sDestino02 As String
Dim NumeroExcel As Integer
Dim nFil01 As Integer
Dim nFil02 As Integer

On Error GoTo MsgError:

Dim rs2 As ADODB.Recordset
If (fAbrRst(rs2, "select * from placontratotmp")) Then
    rs2.MoveFirst
Else
    MsgBox "Seleccione los contratos", vbCritical, Me.Caption
    Exit Sub
End If

'Copiamos el Excel que va a contener la lista de contratos desde el servidor

March01 = "ListaCont01.xls"
'March01 = "ListaCont01.xlsb"
MarchExcel01 = March01

NumeroExcel = 0

Screen.MousePointer = vbHourglass

sOrigen01 = Path_Reports & "CONTR\" & March01
Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists(sOrigen01)) Then MsgBox "El archivo " & sOrigen01 & Chr(13) & "No se encuentra": Exit Sub
sDestino01 = App.Path & "\Reports\" & March01
CopyFile sOrigen01, sDestino01, FILE_NOTIFY_CHANGE_LAST_WRITE
NumeroExcel = 1

'If Trim(UCase(wuser)) = "sa" Then
    
    
'End If

'Set xlApp_1 = GetObject(App.Path & "\Reports\" & March01)
'Set xlApp01 = xlApp_1.Application
'For I = 1 To xlApp01.Workbooks.count
''    MsgBox "xlApp01.Workbooks(I).Name: " & xlApp01.Workbooks(I).Name & Chr(13) & "March01: " & March01, vbInformation, Me.Caption
'   If xlApp01.Workbooks(I).Name = March01 Then
'      Set xlBook01 = xlApp01.Workbooks(I)
'      xlApp01.Workbooks(I).Activate
'      numwindows01 = I
'      GoTo Continua
'   End If
'Next I
'
'For I = 1 To xlApp01.Workbooks.count
''    MsgBox "xlApp01.Workbooks(I).Name: " & xlApp01.Workbooks(I).Name & Chr(13) & "March01: " & March01, vbInformation, Me.Caption
'   If xlApp01.Workbooks(I).Name = March01 Then
'      Set xlBook01 = xlApp01.Workbooks(I)
'      xlApp01.Workbooks(I).Activate
'      numwindows01 = I
'      GoTo Continua
'   End If
'Next I


'Continua:
'xlApp01.Application.Visible = True
'quitar
'xlApp01.Parent.Windows(March01).Visible = True

Dim ls_IdTipoContrato As String
ls_IdTipoContrato = Format(cbo_tipocont.ReturnCodigo, "00")



    Set xlApp01 = CreateObject("Excel.Application")
    xlApp01.Workbooks.Add
    Set xlApp02 = xlApp01.Application
    Set xlBook = xlApp02.Workbooks(1)
'    Set xlSheet = xlApp2.Worksheets("HOJA1")
    
    
    
Set xlSheet01 = xlApp01.Worksheets("Hoja1")
xlSheet01.Activate

nFil01 = 1
   xlSheet01.Cells(nFil01, 1).Value = "PLACOD"
   xlSheet01.Cells(nFil01, 2).Value = "SEXO"
   xlSheet01.Cells(nFil01, 3).Value = "PLANOM"
   xlSheet01.Cells(nFil01, 4).Value = "FINGRESO"
   xlSheet01.Cells(nFil01, 5).Value = "AREA"
   xlSheet01.Cells(nFil01, 6).Value = "PLACAR"
   xlSheet01.Cells(nFil01, 7).Value = "PLADEPENDE"
   xlSheet01.Cells(nFil01, 8).Value = "PLADOCS"
   xlSheet01.Cells(nFil01, 9).Value = "PLADOM"
   xlSheet01.Cells(nFil01, 10).Value = "PLAING"
   xlSheet01.Cells(nFil01, 11).Value = "PLATINI"
   xlSheet01.Cells(nFil01, 12).Value = "PLATFIN"
   xlSheet01.Cells(nFil01, 13).Value = "PLABAS"
   xlSheet01.Cells(nFil01, 14).Value = "PLAFAM"
   'xlSheet01.Cells(nFil01, 15).NumberFormat = "@"
   xlSheet01.Cells(nFil01, 15).Value = "PLASUELDO"
   xlSheet01.Cells(nFil01, 16).Value = "letras"
   xlSheet01.Cells(nFil01, 17).Value = "TIEMPO_LETRAS"
   xlSheet01.Cells(nFil01, 18).Value = "FECHA_CONTRA"
   xlSheet01.Cells(nFil01, 19).Value = "FIRMANTE_NOMBRE"
   xlSheet01.Cells(nFil01, 20).Value = "FIRMANTE_CARGO"
   'xlSheet01.Cells(nFil01, 21).NumberFormat = "@"
   xlSheet01.Cells(nFil01, 21).Value = "FIRMANTE_DNI"
   xlSheet01.Cells(nFil01, 22).Value = "PLATINIANT"
   xlSheet01.Cells(nFil01, 23).Value = "PLATFINANT"
   xlSheet01.Cells(nFil01, 24).Value = "TIEMPO_LETRASANT"
   xlSheet01.Cells(nFil01, 25).Value = "FECHADIAANT"
   xlSheet01.Cells(nFil01, 26).Value = "NOMBRESUPLENCIA"
   xlSheet01.Cells(nFil01, 27).Value = "FECNACTRAB"
   xlSheet01.Cells(nFil01, 28).Value = "SEXOTRAB"
   xlSheet01.Cells(nFil01, 29).Value = "ECIVILTRAB"
   

nFil01 = nFil01 + 1

Do While Not rs2.EOF
   xlSheet01.Cells(nFil01, 1).Value = Trim(rs2!PlaCod & "")
   xlSheet01.Cells(nFil01, 2).Value = Trim(rs2!sexo & "")
   xlSheet01.Cells(nFil01, 3).Value = Trim(rs2!planom & "")
   xlSheet01.Cells(nFil01, 4).Value = Trim(rs2!fIngreso & "")
   xlSheet01.Cells(nFil01, 5).Value = Trim(rs2!Area & "")
   xlSheet01.Cells(nFil01, 6).Value = Trim(rs2!placar & "")
   xlSheet01.Cells(nFil01, 7).Value = Trim(rs2!PLADEPENDE & "")
   xlSheet01.Cells(nFil01, 8).Value = "'" & Trim(rs2!pladocs & "")
   xlSheet01.Cells(nFil01, 9).Value = Trim(rs2!pladom & "")
   xlSheet01.Cells(nFil01, 10).Value = Trim(rs2!plaing & "")
   xlSheet01.Cells(nFil01, 11).Value = Trim(rs2!platini & "")
   xlSheet01.Cells(nFil01, 12).Value = Trim(rs2!platfin & "")
   xlSheet01.Cells(nFil01, 13).Value = Trim(rs2!plabas & "")
   xlSheet01.Cells(nFil01, 14).Value = Trim(rs2!plaasigfam & "")
   xlSheet01.Cells(nFil01, 15).NumberFormat = "@"
   xlSheet01.Cells(nFil01, 15).Value = Trim(rs2!plasueldo & "")
   xlSheet01.Cells(nFil01, 16).Value = Trim(rs2!Letras & "") & " Soles"
   xlSheet01.Cells(nFil01, 17).Value = Trim(rs2!TIEMPO_LETRAS & "")
   xlSheet01.Cells(nFil01, 18).Value = Trim(rs2!fec_contra & "")
   xlSheet01.Cells(nFil01, 19).Value = Trim(rs2!NombreFirma & "")
   xlSheet01.Cells(nFil01, 20).Value = Trim(rs2!CargoFirma & "")
   xlSheet01.Cells(nFil01, 21).NumberFormat = "@"
   xlSheet01.Cells(nFil01, 21).Value = Trim(rs2!DniFirma & "")
   xlSheet01.Cells(nFil01, 22).Value = Trim(rs2!platiniant & "")
   xlSheet01.Cells(nFil01, 23).Value = Trim(rs2!platfinant & "")
   xlSheet01.Cells(nFil01, 24).Value = Trim(rs2!tiempo_letrasant & "")
   xlSheet01.Cells(nFil01, 25).Value = Trim(rs2!Fechadiaant & "")
   xlSheet01.Cells(nFil01, 26).Value = Trim(rs2!NOMBRESUPLENCIA & "")
   xlSheet01.Cells(nFil01, 27).Value = Trim(rs2!FECNACTRAB & "")
   xlSheet01.Cells(nFil01, 28).Value = Trim(rs2!SEXOTRAB & "")
   xlSheet01.Cells(nFil01, 29).Value = Trim(rs2!ECIVILTRAB & "")
   
   If ls_IdTipoContrato = "14" Then 'add jcms 291122 contrato x renovacion indicado por H.lizano+legal
        xlSheet01.Cells(nFil01, 25).Value = Trim(rs2!txtfecha_firma_contrato & "")
        xlSheet01.Cells(nFil01, 24).Value = Trim(rs2!TiempoLetra_ContratoInicial & "")
        xlSheet01.Cells(nFil01, 22).Value = Trim(rs2!FechaIni_ContratoInicial & "")
        xlSheet01.Cells(nFil01, 23).Value = Trim(rs2!FechaFin_ContratoInicial & "")
   End If
   
   nFil01 = nFil01 + 1
   rs2.MoveNext
Loop
rs2.Close: Set rs2 = Nothing
Dim ls_NameFile2 As String
ls_NameFile2 = App.Path & "\Reports\" & MarchExcel01
xlApp02.DisplayAlerts = False
xlApp02.ActiveWorkbook.SaveAs FileName:=ls_NameFile2, FileFormat:=xlNormal, Password:="", WriteRespassword:="", ReadOnlyRecommended:=True, CreateBackup:=False, AccessMode:=xlSaveChanges

'    xlApp02.Close
    xlApp02.Quit
'    xlApp01.DisplayAlerts = False
    xlApp01.Quit
    xlApp01.DisplayAlerts = True
    

'xlApp01.Parent.Windows(March01).Visible = False

'Copiamos el Word de contratos desde el servidor



March01 = "CONT" + wcia + Format(cbo_tipocont.ReturnCodigo, "00") + ".doc"
'fc_CodigoComboBox(Cmbtipotrabajador, 2) + ".doc"
sOrigen01 = Path_Reports & "CONTR\" & March01

Set fso = CreateObject("Scripting.FileSystemObject")
If Not (fso.FileExists(sOrigen01)) Then MsgBox "El archivo " & sOrigen01 & Chr(13) & "No se encuentra": Exit Sub
sDestino01 = App.Path & "\Reports\" & March01
CopyFile sOrigen01, sDestino01, FILE_NOTIFY_CHANGE_LAST_WRITE

strRutaArchivo = sDestino01
'Dim appWord01 As Word.Application
'Dim docWord01 As Word.Document

Dim appWord01 As Object
Dim docWord01 As Object


Set appWord01 = CreateObject("Word.Application")
Set docWord01 = appWord01.Documents.Open(strRutaArchivo)

docWord01.MailMerge.MainDocumentType = wdFormLetters

docWord01.MailMerge.OpenDataSource Name:=App.Path & "\Reports\" & MarchExcel01, _
Connection:="Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=Admin;Jet OLEDB:Engine Type=35;Jet OLE", _
SQLStatement:="SELECT * FROM `Hoja1$`", SQLStatement1:="", Subtype:=wdMergeSubTypeAccess

appWord01.Visible = True


With docWord01
   .MailMerge.Destination = wdSendToNewDocument
   .MailMerge.Execute True
'   .MailMerge.DataSource.FirstRecord = wdDefaultFirstRecord
'   .MailMerge.DataSource.LastRecord = wdDefaultLastRecord
'   .MailMerge.Execute Pause:=True
End With

docWord01.Close False
Set appWord01 = Nothing
Set docWord01 = Nothing

'    xlApp02.Close
    'xlApp01.Quit
'    xlApp01.DisplayAlerts = False
'    xlApp01.Quit
'    xlApp01.DisplayAlerts = True

    If Not xlApp02 Is Nothing Then Set xlApp02 = Nothing
    If Not xlApp01 Is Nothing Then Set xlApp101 = Nothing
    If Not xlBook01 Is Nothing Then Set xlBook01 = Nothing
    If Not xlSheet01 Is Nothing Then Set xlSheet01 = Nothing
    If Not xlBook02 Is Nothing Then Set xlBook02 = Nothing
    If Not xlSheet02 Is Nothing Then Set xlSheet02 = Nothing
    
Screen.MousePointer = vbDefault
Exit Sub

MsgError:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, Me.Caption
    Screen.MousePointer = vbDefault
    
End Sub
Private Function Tipo_contrato(lCia As String, lFecha_Ing As String, lFecha_Fin As String) As String
Tipo_contrato = "03"
Dim lDias As Integer
 lDias = DateDiff("D", lFecha_Ing, lFecha_Fin)
 
 If wGrupoPla = "01" And wcia = "03" And Trim(rsContratos!cod_tip_contrato & "") = "12" Then
    Tipo_contrato = Trim(rsContratos!cod_tip_contrato & "")
    Exit Function
 End If

 If lDias > 1825 Then 'Pasa los 5 años
    Tipo_contrato = "**"
    Exit Function
 End If
 If lCia = "03" Then
    Tipo_contrato = "12"
    Exit Function
 End If
 If lDias >= 1095 Then 'Pasa los 3 años
    Tipo_contrato = "04"
    Exit Function
 End If
End Function
Private Sub Contratos_Generados_Excel()
If rsContratos.RecordCount > 0 Then rsContratos.MoveFirst Else Exit Sub

'Dim nFil As Integer
'Dim xlApp2 As Excel.Application
'Dim xlApp1  As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet

Dim nFil As Integer
Dim xlApp2 As Object
Dim xlApp1  As Object
Dim xlBook As Object
Dim xlSheet As Object



Set xlApp1 = CreateObject("Excel.Application")
xlApp1.Workbooks.Add
Set xlApp2 = xlApp1.Application
Set xlBook = xlApp2.Workbooks(1)
Set xlSheet = xlApp2.Worksheets("HOJA1")
xlSheet.Name = "Contratos"

xlSheet.Cells(2, 1).Value = "RELACION DE CONTRATOS"

xlSheet.Range("A:A").ColumnWidth = 11
xlSheet.Range("B:B").ColumnWidth = 8
xlSheet.Range("C:C").ColumnWidth = 40
xlSheet.Range("D:E").ColumnWidth = 12
xlSheet.Range("F:F").ColumnWidth = 6
xlSheet.Range("F:F").HorizontalAlignment = xlCenter
xlSheet.Range("G:G").ColumnWidth = 12
xlSheet.Range("H:H").ColumnWidth = 9
xlSheet.Range("I:I").ColumnWidth = 20
xlSheet.Range("J:J").ColumnWidth = 45

xlSheet.Range("H:H").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
xlSheet.Range("D:E").NumberFormat = "@"
xlSheet.Range("G:G").NumberFormat = "@"

xlSheet.Cells(4, 1).Value = "PLANTA"
xlSheet.Cells(4, 2).Value = "CODIGO"
xlSheet.Cells(4, 3).Value = "NOMBRE"
xlSheet.Cells(4, 4).Value = "F.INGRESO"
xlSheet.Cells(4, 5).Value = "F.INICIO"
xlSheet.Cells(4, 6).Value = "DIAS"
xlSheet.Cells(4, 7).Value = "F.FINAL"
xlSheet.Cells(4, 8).Value = "BASICO"
xlSheet.Cells(4, 9).Value = "CARGO"
xlSheet.Cells(4, 10).Value = "TIPO_CONTRATO"


xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, 10)).Borders.LineStyle = xlContinuous
xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, 10)).HorizontalAlignment = xlCenter
xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, 10)).Interior.ColorIndex = 37
xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, 10)).Interior.Pattern = xlSolid
xlSheet.Range(xlSheet.Cells(4, 1), xlSheet.Cells(4, 10)).Font.Bold = True

nFil = 5
Do While Not rsContratos.EOF
   If Trim(rsContratos!SEL & "") = "S" Then
      xlSheet.Cells(nFil, 1).Value = Trim(rsContratos!Planta & "")
      xlSheet.Cells(nFil, 2).Value = Trim(rsContratos!PlaCod & "")
      xlSheet.Cells(nFil, 3).Value = Trim(rsContratos!nombre & "")
      xlSheet.Cells(nFil, 4).Value = Trim(rsContratos!fIngreso & "")
      xlSheet.Cells(nFil, 5).Value = Trim(rsContratos!fec_ini & "")
      xlSheet.Cells(nFil, 6).Value = rsContratos!Dias
      xlSheet.Cells(nFil, 7).Value = Trim(rsContratos!fec_fin & "")
      xlSheet.Cells(nFil, 8).Value = rsContratos!importe
      xlSheet.Cells(nFil, 9).Value = Trim(rsContratos!Cargo & "")
      xlSheet.Cells(nFil, 10).Value = Trim(rsContratos!Tipo_contra & "")
      nFil = nFil + 1
   End If
   rsContratos.MoveNext
Loop

xlApp2.Application.ActiveWindow.DisplayGridlines = False
'xlApp2.ActiveWindow.Zoom = 80
xlApp2.Application.Visible = True

If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xlBook Is Nothing Then Set xlBook = Nothing
If Not xlSheet Is Nothing Then Set xlSheet = Nothing

End Sub


Private Sub TxtcodSuplencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    Frmgrdpla.Show
    Frmgrdpla.ZOrder 0
End If
End Sub
Public Sub Reporte()
'On Error GoTo MsgErr:
If rsContratos.RecordCount = 0 Then
    MsgBox "No existen registros a exportar.", vbExclamation, Me.Caption
'    FecIni(0).SetFocus
    Me.Cmbtipotrabajador.SetFocus
    Exit Sub
End If
'Dim Rq As ADODB.Recordset
'Sql = "select c.fecha1 ,c.caja,(select descrip from maestros_2 where ciamaestro=c.cia+'048' and cod_maestro2=c.caja and status<>'*') as nom_caja"
'Sql = Sql & " ,c.nro_ffijo,d.ccosto"
'Sql = Sql & " ,(select cgdes from CONMASPCGE where cod_cia=c.cia and AYO=" & FecFin(1).Year & " AND cgcod=d.ccosto and status<>'*') as cc"
'Sql = Sql & " ,d.motivo,d.moneda,d.importe,d.cod_trab"
'Sql = Sql & " ,dbo.fc_Razsoc_Trabajador(d.cod_trab) as nom_trab"
'Sql = Sql & " from fondofijo01 c inner join fondofijo02 d on (c.cia=d.cia and c.caja=d.caja and c.nro_ffijo=d.nro_ffijo)"
'Sql = Sql & " where c.status<>'*' and d.status<>'*' and c.fecha1 between '" & Format(FecIni(0).Value, "mm/dd/yyyy") & " 12:00:00 am' and '" & Format(FecFin(1).Value, "mm/dd/yyyy") & " 11:59:59 pm'  and right(cuenta,4)='0218'"
'If Not fAbrRst(Rq, Sql) Then
'    MsgBox "No existen Registros para el criterio especificado", vbExclamation, Me.Caption
'    GoTo SALIR:
'End If


    Dim nFil As Integer
    Dim nCol As Integer
'    Dim xlApp2 As Excel.Application
'    Dim xlApp1 As Excel.Application
'    Dim xlBook As Excel.Workbook
'    Dim xlSheet As Excel.Worksheet
    
    Dim xlApp2 As Object
    Dim xlApp1 As Object
    Dim xlBook As Object
    Dim xlSheet As Object


Screen.MousePointer = vbHourglass

    Set xlApp1 = CreateObject("Excel.Application")
    xlApp1.Workbooks.Add
    Set xlApp2 = xlApp1.Application
    Set xlBook = xlApp2.Workbooks(1)
    Set xlSheet = xlApp2.Worksheets("HOJA1")
    
    With xlSheet
'        .Range("A:A").ColumnWidth = 22
'        '.Range("A:A").NumberFormat = "@"
'        .Range("B:B").ColumnWidth = 13
'        .Range("C:C").ColumnWidth = 11
'        .Range("D:D").ColumnWidth = 11


         nFil = 1
        .Cells(nFil, 1).Value = "COMACSA"
        .Cells(nFil, 1).Font.Size = 10
        .Cells(nFil, 1).Font.Bold = True
        
        
'        .cells(nFil, 4).Value = "AL " & Format(TxtFecha, "DD/MM/YYYY")
'        .cells(nFil, 4).Font.Size = 10
'        .cells(nFil, 4).Font.Bold = True
        
        Dim Columnas(11) As String
        Dim Ancho(11) As Double
        Columnas(0) = "Cod.Trabajador": Ancho(0) = 8
        Columnas(1) = "Trabajador": Ancho(1) = 25
        Columnas(2) = "Tipo Contrato": Ancho(2) = 25
        Columnas(3) = "Fecha Inicio": Ancho(3) = 12
        Columnas(4) = "Fecha Termino": Ancho(4) = 12
        Columnas(5) = "Importe S.Básico": Ancho(5) = 15
        Columnas(6) = "Cargo": Ancho(6) = 25
        Columnas(7) = "Nuevo Inicio": Ancho(7) = 15
        Columnas(8) = "Nuevo fin": Ancho(8) = 15
        Columnas(9) = "Fecha Ingreso": Ancho(9) = 12
        Columnas(10) = "Cen. Costo": Ancho(10) = 25
        
        Dim sTitulo As String
        
        
        If Me.OptVigentes.Value Then sTitulo = "VIGENTES AL " & Now
        If Me.OptVencidos.Value Then sTitulo = "VENCIDOS AL " & Now
        If Me.OptVence.Value Then sTitulo = "CON VENCIMIENTO EN " & cbomes.Text & " - " & Me.Txtano.Text
         If Me.OptNuevos.Value Then sTitulo = "NUEVOS INGRESARON EN " & cbomes.Text & " - " & Me.Txtano.Text
        

        
        
        nFil = nFil + 2
        
        '.cells(nFil, 1).Value = "DESPACHOS PATIO PATTERN Y ROMAN PATTER " & FecIni.Value & " AL " & FecFin.Value
        .Cells(nFil, 1).Value = "REPORTE CONTRATOS " & sTitulo
        .Cells(nFil, 1).Font.Size = 10
        .Cells(nFil, 1).Font.Bold = True
        
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Merge
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).HorizontalAlignment = xlCenter
        
        .Range(.Cells(nFil, 1), .Cells(nFil, UBound(Columnas))).Borders.LineStyle = xlContinuous
        
        
        nFil = nFil + 1
        
        Dim sTipoTrab As String
        sTipoTrab = "<TODOS>"
        If Cmbtipotrabajador.Text <> "" Then sTipoTrab = Cmbtipotrabajador.Text
        
        .Cells(nFil, 1).Value = "TIPO TRABAJADOR: " & sTipoTrab
        .Cells(nFil, 1).Font.Size = 10
        .Cells(nFil, 1).Font.Bold = True
        
          nFil = nFil + 1
        
        Dim sCenCosto As String
        sCenCosto = "<TODOS>"
        If Me.CmbCcosto.Text <> "" Then sCenCosto = CmbCcosto.Text
        
        .Cells(nFil, 1).Value = "CENTRO DE COSTO: " & sCenCosto
        .Cells(nFil, 1).Font.Size = 10
        .Cells(nFil, 1).Font.Bold = True
        
         Dim sTipoContrato As String
        sTipoContrato = "<TODOS>"
        If Me.cbo_tipocont.Text <> "" Then sTipoContrato = cbo_tipocont.Text
        
        .Cells(nFil, 1).Value = "TIPO CONTRATO: " & sTipoContrato
        .Cells(nFil, 1).Font.Size = 10
        .Cells(nFil, 1).Font.Bold = True
        
        
        
        
        
        

        
        nFil = nFil + 2
        
        
        
        Dim I As Integer
        nCol = 1
        For I = 1 To UBound(Columnas) + 1
                
            .Range(.Cells(nFil, nCol), .Cells(nFil, nCol)).ColumnWidth = Ancho(I - 1)
            .Cells(nFil, nCol).Value = Columnas(I - 1)
            .Cells(nFil, nCol).Font.Bold = True
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).Merge
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).HorizontalAlignment = xlCenter
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).VerticalAlignment = xlCenter
            .Range(.Cells(nFil, nCol), .Cells(nFil + 1, nCol)).WrapText = True
                            
            nCol = nCol + 1
        Next
        .Range(.Cells(nFil, 1), .Cells(nFil + 1, UBound(Columnas))).Borders.LineStyle = xlContinuous
        nFil = nFil + 1
        Dim J As Integer
        J = 0
        Dim F As Integer
        F = 0
        Dim xTip As String
        xTip = ""
        Barra.Value = 1
        Dim rsClon As New ADODB.Recordset
        Set rsClon = rsContratos.Clone
        If rsClon.RecordCount = 1 Then
            Barra.Min = 0
        Else
            Barra.Max = rsClon.RecordCount
            Barra.Min = 1
        End If
        Barra.Visible = True
       
     
        Dim xItem As Integer
        xItem = 1
        'With Rq
            If rsClon.RecordCount > 0 Then
                rsClon.MoveFirst
                Do While Not rsClon.EOF
                        Barra.Value = rsClon.AbsolutePosition
                        nFil = nFil + 1
                       
                        xItem = xItem + 1
                         .Cells(nFil, 1).Value = Trim(rsClon!PlaCod)
                         .Cells(nFil, 2).Value = Trim(rsClon!nombre)
                         .Cells(nFil, 3).Value = Trim(rsClon!Tipo_contra)
                         
                         .Cells(nFil, 4).Value = "'" & Format(Trim(rsClon!fec_ini), "dd/mm/yyyy")
                         .Cells(nFil, 5).Value = "'" & Format(Trim(rsClon!fec_fin), "dd/mm/yyyy")
                        
                         
                         .Cells(nFil, 6).Value = Trim(rsClon!importe)
                          .Range(.Cells(nFil, 6), .Cells(nFil, 5)).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
                          
                         
                         .Cells(nFil, 7).Value = Trim(rsClon!Cargo)
                         .Cells(nFil, 8).Value = "'" & Format(Trim(rsClon!fec_iniNew), "dd/mm/yyyy")
                         .Cells(nFil, 9).Value = "'" & Format(Trim(rsClon!fec_finNew), "dd/mm/yyyy")
                         
                         .Cells(nFil, 10).Value = "'" & Format(Trim(rsClon!fIngreso), "dd/mm/yyyy")
                         .Cells(nFil, 11).Value = "'" & Trim(rsClon!CENCOSTO)
                         
'                          rsContratos.Fields.Append "placod", adChar, 10, adFldIsNullable
'    rsContratos.Fields.Append "Nombre", adChar, 50, adFldIsNullable
'    rsContratos.Fields.Append "Tipo_contra", adChar, 150, adFldIsNullable
'    rsContratos.Fields.Append "Dias", adChar, 5, adFldIsNullable
'    rsContratos.Fields.Append "fec_ini", adChar, 10, adFldIsNullable
'    rsContratos.Fields.Append "fec_fin", adChar, 10, adFldIsNullable
'    rsContratos.Fields.Append "importe", adCurrency, 18, adFldIsNullable
'    rsContratos.Fields.Append "Cargo", adChar, 100, adFldIsNullable
'    rsContratos.Fields.Append "cod_tip_contrato", adChar, 2, adFldIsNullable
'    rsContratos.Fields.Append "num_contrato", adChar, 10, adFldIsNullable
'    rsContratos.Fields.Append "codcargo", adChar, 3, adFldIsNullable
'    rsContratos.Fields.Append "Sel", adChar, 1, adFldIsNullable
'    rsContratos.Fields.Append "cargo_Ori", adChar, 3, adFldIsNullable
'    rsContratos.Fields.Append "contrato_ori", adChar, 2, adFldIsNullable
'    rsContratos.Fields.Append "empresa", adChar, 2, adFldIsNullable
'    rsContratos.Fields.Append "motivo", adChar, 1, adFldIsNullable
'    rsContratos.Fields.Append "fingreso", adChar, 10, adFldIsNullable
'    rsContratos.Fields.Append "Planta", adChar, 20, adFldIsNullable
'    rsContratos.Fields.Append "fec_iniNew", adChar, 10, adFldIsNullable
'    rsContratos.Fields.Append "fec_finNew", adChar, 10, adFldIsNullable
'    rsContratos.Fields.Append "codcargoact", adChar, 3, adFldIsNullable
'    'Añado el trabajador a quien se le va a suplir por un periodo
'    rsContratos.Fields.Append "placodsuple", adChar, 10, adFldIsNullable
                                                  
                                                 
                    
                    rsClon.MoveNext
                Loop
            End If
            'Barra.Value = 0
            Barra.Visible = False
                            
                            
            nFil = nFil + 1
            .Cells(nFil, 1).Value = "'TOTAL REGISTROS: " & xItem
            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).Merge
            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).HorizontalAlignment = xlRight
            .Cells(nFil, 1).Font.Bold = True

'            .Cells(nFil, 4).FormulaR1C1 = "=SUM(R[-" & RsClon.RecordCount & "]C:R[-1]C)"
'            .Cells(nFil, 4).Font.Bold = True
''
'            .Range(.Cells(nFil, 8), .Cells(nFil, 8)).NumberFormat = "#,###,##0.000;[Red](#,###,##0.000)"
            
         nFil = nFil + 2
         
        '//*** Resumen***////
        
'        .Cells(nFil, 1).Value = "RESUMEN"
'        .Cells(nFil, 1).Font.Size = 10
'        .Cells(nFil, 1).Font.Bold = True
'
'        .Cells(nFil, 2).Value = "NOMBRE PRACTICANTE"
'        .Cells(nFil, 2).Font.Size = 10
'        .Cells(nFil, 2).Font.Bold = True
'
'        .Cells(nFil, 3).Value = "MON"
'        .Cells(nFil, 3).Font.Size = 10
'        .Cells(nFil, 3).Font.Bold = True
'
'        .Cells(nFil, 4).Value = "IMPORTE"
'        .Cells(nFil, 4).Font.Size = 10
'        .Cells(nFil, 4).Font.Bold = True
'
'        .Range(.Cells(nFil, 1), .Cells(nFil, 4)).Borders.LineStyle = xlContinuous
'
'        Sql = "select d.cod_trab,d.moneda,sum(d.importe) as importe,dbo.fc_Razsoc_Trabajador(d.cod_trab) as nom_trab"
'        Sql = Sql & " from fondofijo01 c inner join fondofijo02 d on (c.cia=d.cia and c.caja=d.caja and c.nro_ffijo=d.nro_ffijo)"
'        Sql = Sql & " where c.status<>'*' and d.status<>'*' and c.fecha1 between '" & Format(FecIni(0).Value, "mm/dd/yyyy") & " 12:00:00 am' and '" & Format(FecFin(1).Value, "mm/dd/yyyy") & " 11:59:59 pm'  and right(cuenta,4)='0218'"
'        Sql = Sql & " group by d.cod_trab,d.moneda order by nom_trab"
'        If Not fAbrRst(Rq, Sql) Then
'            MsgBox "No existen Registros para el criterio especificado", vbExclamation, Me.Caption
'            GoTo SALIR:
'        End If
'
'        xItem = 1
'        'With Rq
'            If RsClon.RecordCount > 0 Then
'                RsClon.MoveFirst
'                Do While Not RsClon.EOF
'                        Barra.Value = RsClon.AbsolutePosition
'                        nFil = nFil + 1
'
'                        xItem = xItem + 1
'                         .Cells(nFil, 1).Value = Trim(Rq!cod_trab)
'                         .Cells(nFil, 2).Value = Trim(Rq!nom_trab)
'                         .Cells(nFil, 3).Value = Trim(Rq!moneda)
'                         .Cells(nFil, 4).Value = Trim(Rq!importe)
'                         .Range(.Cells(nFil, 4), .Cells(nFil, 5)).NumberFormat = "#,###,##0.00;[Red](#,###,##0.00)"
'
'
'                    RsClon.MoveNext
'                Loop
'            End If
'            'Barra.Value = 0
'            Barra.Visible = False
'
'
'            nFil = nFil + 1
'            .Cells(nFil, 1).Value = "'TOTAL "
'            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).Merge
'            .Range(.Cells(nFil, 1), .Cells(nFil, 3)).HorizontalAlignment = xlRight
'            .Cells(nFil, 1).Font.Bold = True
'
'            .Cells(nFil, 4).FormulaR1C1 = "=SUM(R[-" & RsClon.RecordCount & "]C:R[-1]C)"
'            .Cells(nFil, 4).Font.Bold = True
'
'
'
        


         
     End With
     
'    xlSheet.Range(xlSheet.cells(1, 1), xlSheet.cells(nFil, 5)).Select
'    xlSheet.PageSetup.PrintArea = "$A$1:$E$" & nFil
'    With xlSheet.PageSetup
'        .PrintTitleRows = "$1:$9"
'        .PrintTitleColumns = ""
'    End With
'    xlSheet.PageSetup.PrintArea = "$A$1:$E$" & nFil
'    With xlSheet.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = ""
'        .RightHeader = "&P"
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = ""
'        .leftMargin = Application.InchesToPoints(0.590551181102362)
'        .rightMargin = Application.InchesToPoints(0)
'        .topMargin = Application.InchesToPoints(0)
'        .bottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0)
'        .FooterMargin = Application.InchesToPoints(0)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        '.PrintQuality = Array(120, 144)
'        .CenterHorizontally = False
'        .CenterVertically = False
'        .Orientation = xlPortrait
'        .Draft = False
'        .PaperSize = xlPaperLetter
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 100
'    End With
    'xlSheet.PrintPreview
    'xlSheet.SelectedSheets.PrintPreview
    'ActiveWindow.SelectedSheets.PrintPreview
    
'    xlSheet.cells(nFil, 1).Value = "Nro de Item(s): " & CStr(Nr)
'    xlSheet.cells(nFil, 1).Font.Bold = True
            
    'Barra.Visible = False
    Set xlSheet = xlApp2.Worksheets("Hoja1")
    xlApp2.Application.ActiveWindow.DisplayGridlines = False
    
    xlSheet.Range("A1:A1").Select
                
    
    xlApp1.ActiveWindow.Zoom = 80
    xlApp2.Application.Visible = True
    

    If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
    If Not xlBook Is Nothing Then Set xlBook = Nothing
    If Not xlSheet Is Nothing Then Set xlSheet = Nothing
Salir:
Screen.MousePointer = 0
rsClon.Close
Set Rq = Nothing
Exit Sub
MsgErr:
Screen.MousePointer = 0
MsgBox Err.Number & " - " & Err.Description, vbCritical, Me.Caption
End Sub
