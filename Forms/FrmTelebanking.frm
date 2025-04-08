VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmTelebanking 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivo Deposito al Banco"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14400
   Icon            =   "FrmTelebanking.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   14175
      Begin VB.OptionButton optCancelado 
         Caption         =   "Cancelado"
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
         Left            =   3120
         TabIndex        =   36
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optPendiente 
         Caption         =   "Pendiente"
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
         Left            =   1320
         TabIndex        =   35
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox cboCuenta 
         Height          =   315
         ItemData        =   "FrmTelebanking.frx":030A
         Left            =   10440
         List            =   "FrmTelebanking.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox CmbMonPago 
         Height          =   315
         ItemData        =   "FrmTelebanking.frx":030E
         Left            =   8760
         List            =   "FrmTelebanking.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CmbBcoPago 
         Height          =   315
         ItemData        =   "FrmTelebanking.frx":0312
         Left            =   3480
         List            =   "FrmTelebanking.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker dtpFechaPago 
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   51576833
         CurrentDate     =   37265
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Monto:"
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
         Left            =   9720
         TabIndex        =   33
         Top             =   650
         Width           =   600
      End
      Begin VB.Label lblImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10440
         TabIndex        =   30
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta:"
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
         Left            =   9720
         TabIndex        =   29
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
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
         TabIndex        =   27
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
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
         Left            =   2760
         TabIndex        =   25
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Pago: "
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
         TabIndex        =   22
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.CommandButton btnQuitar 
      Height          =   615
      Left            =   8040
      Picture         =   "FrmTelebanking.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Width           =   855
   End
   Begin VB.CheckBox chkTodos 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   13560
      TabIndex        =   18
      Top             =   3000
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   960
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   7815
      Begin VB.ComboBox Cmbtipotrabajador 
         Height          =   315
         ItemData        =   "FrmTelebanking.frx":0758
         Left            =   4800
         List            =   "FrmTelebanking.frx":075A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker Cmbfecha 
         Height          =   285
         Left            =   750
         TabIndex        =   2
         Top             =   585
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   51576833
         CurrentDate     =   37265
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   4000
         TabIndex        =   4
         Top             =   600
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
         Top             =   585
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51576833
         CurrentDate     =   37265
      End
      Begin MSComCtl2.DTPicker Cmbdel 
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   585
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   51576833
         CurrentDate     =   37267
      End
      Begin VB.TextBox Txtsemana 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         Top             =   585
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cmbtipo 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   6120
         TabIndex        =   16
         Top             =   585
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         Height          =   195
         Left            =   4440
         TabIndex        =   15
         Top             =   585
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana"
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   585
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Boleta"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14415
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   60
         Width           =   6450
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   11760
         TabIndex        =   32
         Top             =   75
         Width           =   2535
      End
      Begin VB.Label Lblfecha 
         Alignment       =   1  'Right Justify
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
         Left            =   6675
         TabIndex        =   11
         Top             =   120
         Width           =   1815
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
         TabIndex        =   8
         Top             =   120
         Width           =   825
      End
   End
   Begin TrueOleDBGrid70.TDBGrid GrdSolicitud 
      Height          =   4215
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7435
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "TipoRegistro"
      Columns(0).DataField=   "TipoRegistro"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Boleta"
      Columns(1).DataField=   "Boleta"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "pagotipcta"
      Columns(2).DataField=   "pagotipcta"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TipoCuenta"
      Columns(3).DataField=   "TipoCuenta"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Sucursal"
      Columns(4).DataField=   "Sucursal"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "NroCuenta"
      Columns(5).DataField=   "NroCuenta"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "tipo_doc"
      Columns(6).DataField=   "tipo_doc"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "PlaCod"
      Columns(7).DataField=   "PlaCod"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Trabajador"
      Columns(8).DataField=   "Trabajador"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Moneda"
      Columns(9).DataField=   "TipoMoneda"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Pago Neto"
      Columns(10).DataField=   "totneto"
      Columns(10).NumberFormat=   "#,##0.00"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   4
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   "Action"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "IdBoleta"
      Columns(12).DataField=   "IdBoleta"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=5292"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=5212"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(21)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=1244"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1164"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=3519"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3440"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=8196"
      Splits(0)._ColumnProps(42)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=1931"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1852"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=8196"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=5292"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=5212"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=8196"
      Splits(0)._ColumnProps(55)=   "Column(8).WrapText=1"
      Splits(0)._ColumnProps(56)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(57)=   "Column(9).Width=873"
      Splits(0)._ColumnProps(58)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(9)._WidthInPix=794"
      Splits(0)._ColumnProps(60)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(61)=   "Column(9)._ColStyle=8196"
      Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(63)=   "Column(10).Width=2302"
      Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=2223"
      Splits(0)._ColumnProps(66)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(67)=   "Column(10)._ColStyle=8706"
      Splits(0)._ColumnProps(68)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(69)=   "Column(11).Width=529"
      Splits(0)._ColumnProps(70)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(11)._WidthInPix=450"
      Splits(0)._ColumnProps(72)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(74)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(75)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(77)=   "Column(12)._EditAlways=0"
      Splits(0)._ColumnProps(78)=   "Column(12)._ColStyle=8196"
      Splits(0)._ColumnProps(79)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(80)=   "Column(12).Order=13"
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
      HeadLines       =   2
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
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&HC0C0C0&"
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
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=78,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.locked=-1"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(67)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(71)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.wraptext=-1,.locked=-1"
      _StyleDefs(72)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
      _StyleDefs(73)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
      _StyleDefs(74)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
      _StyleDefs(75)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.locked=-1"
      _StyleDefs(76)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(77)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(78)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(79)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=1,.locked=-1"
      _StyleDefs(80)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14,.alignment=2"
      _StyleDefs(81)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(82)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(83)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(84)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(85)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(86)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(87)  =   "Splits(0).Columns(12).Style:id=86,.parent=13,.locked=-1"
      _StyleDefs(88)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(89)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(90)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(91)  =   "Named:id=33:Normal"
      _StyleDefs(92)  =   ":id=33,.parent=0"
      _StyleDefs(93)  =   "Named:id=34:Heading"
      _StyleDefs(94)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(95)  =   ":id=34,.wraptext=-1"
      _StyleDefs(96)  =   "Named:id=35:Footing"
      _StyleDefs(97)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(98)  =   "Named:id=36:Selected"
      _StyleDefs(99)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(100) =   "Named:id=37:Caption"
      _StyleDefs(101) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(102) =   "Named:id=38:HighlightRow"
      _StyleDefs(103) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(104) =   "Named:id=39:EvenRow"
      _StyleDefs(105) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(106) =   "Named:id=40:OddRow"
      _StyleDefs(107) =   ":id=40,.parent=33"
      _StyleDefs(108) =   "Named:id=41:RecordSelector"
      _StyleDefs(109) =   ":id=41,.parent=34"
      _StyleDefs(110) =   "Named:id=42:FilterBar"
      _StyleDefs(111) =   ":id=42,.parent=33"
   End
   Begin VB.Label lblNro 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
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
      Height          =   255
      Left            =   11640
      TabIndex        =   34
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      Caption         =   "0"
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
      Height          =   255
      Left            =   13080
      TabIndex        =   31
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "FrmTelebanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VTipotrab As String
Dim VIdTelebanking As Integer
Dim VBcoPago As String
Dim VMoneda As String
Dim rsTelebanking As New ADODB.Recordset


Public Sub Grabar()
Dim NroTrans As Integer
Dim contar As Integer
Dim suma As Currency
Dim NroCtaTrabCia As Double
Dim rsTemp As New ADODB.Recordset
Dim IdBcoCta As Integer

On Error GoTo ErrorTrans
NroTrans = 0
If Me.CmbBcoPago.ListIndex < 0 Then
    MsgBox "Seleccione el Banco", vbCritical, Me.Caption
    Exit Sub
End If

If Me.CmbMonPago.ListIndex < 0 Then
    MsgBox "Seleccione la Moneda", vbCritical, Me.Caption
    Exit Sub
End If

If Me.cboCuenta.ListIndex < 0 Then
    MsgBox "Seleccione la Cuenta del Banco", vbCritical, Me.Caption
    Exit Sub
End If

If rsTelebanking.RecordCount <= 0 Then
    MsgBox "No hay Boletas a Pago a Depositar", vbCritical, Me.Caption
    Exit Sub
Else
    
    If rsTelebanking.EOF = False Then rsTelebanking.Update
    
    Set rsTemp = rsTelebanking.Clone
    rsTemp.MoveFirst
    rsTemp.Filter = "Action=1"
    If rsTemp.RecordCount <= 0 Then
        MsgBox "Debe checkear al menos a una boleta de pago a depositar", vbCritical, Me.Caption
        Exit Sub
    Else

        suma = 0
        NroCtaTrabCia = 0
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If rsTemp!Action = True Then
                suma = suma + rsTemp!totneto
                NroCtaTrabCia = NroCtaTrabCia + Val(rsTemp!NroCuenta)
            End If
            rsTemp.MoveNext
        Loop
        lblImporte.Caption = Format(suma, "#,##0.00")
        Me.lblNro.Caption = rsTemp.RecordCount
        If suma <= 0 Then
            MsgBox "El importe a depositar debe ser mayor a cero", vbCritical, Me.Caption
            Exit Sub
        End If
    End If

End If


cn.BeginTrans
NroTrans = 1
contar = 0

IdBcoCta = cboCuenta.ItemData(cboCuenta.ListIndex)

Sql$ = "select * from Plabcocta where idbcocta=" & Val(IdBcoCta)
If (fAbrRst(rs, Sql$)) Then
    NroCtaTrabCia = NroCtaTrabCia + Val(rs!cuentabco)
End If
rs.Close
Set rs = Nothing

Sql$ = "usp_Pla_PlaTelebanking " & Val(VIdTelebanking)
Sql$ = Sql$ & ",'" & wcia & "','" & Format(dtpFechaPago.Value, "mm/dd/yyyy") & "'," & Val(lblCodigo.Caption) & ",'" & Trim(VBcoPago) & "','" & Mid(Trim(Me.CmbMonPago.Text), 2, 3)
Sql$ = Sql$ & "'," & IdBcoCta & "," & Val(lblNro.Caption) & "," & suma & ",'" & VTipotrab & "','1','" & NroCtaTrabCia & "','"
Sql$ = Sql$ & "','" & wuser & "',"
If Val(VIdTelebanking) = 0 Then
    Sql$ = Sql$ & "1"
Else
    Sql$ = Sql$ & "2"
End If

If (fAbrRst(rs, Sql$)) Then
    VIdTelebanking = rs!Id
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If rsTemp!Action = True Then
            Sql$ = "usp_Pla_PlaTelebankingDetalle " & Val(VIdTelebanking)
            Sql$ = Sql$ & "," & Val(rsTemp!id_boleta) & ",'2','" & Trim(rsTemp!pagotipcta)
            Sql$ = Sql$ & "','" & Trim(rsTemp!Sucursal) & "','" & Trim(rsTemp!NroCuenta)
            Sql$ = Sql$ & "','" & Trim(rsTemp!TipoMoneda) & "'," & Val(rsTemp!totneto)
            Sql$ = Sql$ & ",'" & Trim(rsTemp!Boleta)
            Sql$ = Sql$ & "','" & IIf(Me.optPendiente.Value = True, "P", "") & "','" & wuser & "'"
            cn.Execute Sql$
        End If

        rsTemp.MoveNext
    Loop
    
End If

cn.CommitTrans

MsgBox "Se grabaron los datos correctamente", vbInformation, Me.Caption
Unload Me
Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

MsgBox ERR.Description, vbCritical, Me.Caption

End Sub
Public Sub Eliminar()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
If VIdTelebanking <= 0 Then
    MsgBox "Seleccione alguna Solicitud de Sindicato", vbCritical, Me.Caption
    Exit Sub
End If
If MsgBox("¿Esta seguro de eliminar los datos?", vbExclamation + vbYesNo, Me.Caption) = vbNo Then Exit Sub
cn.BeginTrans
NroTrans = 1

Sql$ = "usp_Pla_PlaSolicitudSindicato " & Val(VIdTelebanking)
Sql$ = Sql$ & ",'" & wcia & "','02','" & Trim(VTipo) & "','" & Trim(VTipotrab)
Sql$ = Sql$ & "'," & Year(Cmbfecha.Value) & ",'','" & Format(Cmbfecha.Value, "mm/dd/yyyy")
Sql$ = Sql$ & "','S/.'," & Val(txtMonto.Text) & ",'','" & wuser & "',"
Sql$ = Sql$ & "3"

cn.Execute Sql$
cn.CommitTrans
VIdTelebanking = 0
MsgBox "Se Eliminaron los datos correctamente", vbInformation, Me.Caption

Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

MsgBox ERR.Description, vbCritical, Me.Caption

End Sub

Private Sub btnQuitar_Click()
Dim rsTemp As New ADODB.Recordset
If Not rsTelebanking.EOF Then rsTelebanking.Update
Set rsTemp = rsTelebanking.Clone

If rsTemp.RecordCount > 0 Then
    rsTemp.Filter = "Action=0"
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If rsTemp!Action = False Then
                rsTemp.Delete
            End If
            rsTemp.MoveNext
        Loop
    End If
    rsTemp.Filter = adFilterNone

    Set rsTelebanking = rsTemp.Clone
End If
GrdSolicitud.Refresh
End Sub

Private Sub chkTodos_Click()
If rsTelebanking.RecordCount > 0 Then
    rsTelebanking.MoveFirst
    Do While Not rsTelebanking.EOF
        rsTelebanking!Action = CBool(chkTodos.Value)
        rsTelebanking.Update
        rsTelebanking.MoveNext
    Loop
    
    If chkTodos.Value = 1 Then
        lblNro.Caption = rsTelebanking.RecordCount
    Else
        lblNro.Caption = 0
    End If
End If
End Sub

Private Sub CmbBcoPago_Click()
Cargar_Cuenta_Banco
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))


Call fc_Descrip_Maestros2("01078", "", Cmbtipo)

If Cmbtipo.ListCount = 1 Then Cmbtipo.ListIndex = 0

   Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
   Procesa_Semana


End Sub

Private Sub Cmbfecha_Change()
'If Month(Cmbfecha.Value) = 1 And VTipo = "02" And Cmbtipotrabajador.ListIndex >= 0 Then Command1.Enabled = True Else Command1.Enabled = False
Procesa_Semana
End Sub

Private Sub Cargar_Cuenta_Banco()
VBcoPago = fc_CodigoComboBox(CmbBcoPago, 2)
VMoneda = Mid(Trim(Me.CmbMonPago.Text), 2, 3)
cboCuenta.Clear
Sql$ = "usp_pla_PlaBcoCta 0,'" & wcia & "','" & VBcoPago & "','" & Trim(VMoneda) & "','','','','" & wuser & "',5"
If fAbrRst(rs, Sql$) Then
    rs.MoveFirst
    Do While Not rs.EOF
        cboCuenta.AddItem Trim(rs!Sucursal) & Trim(rs!cuentabco)
        cboCuenta.ItemData(cboCuenta.NewIndex) = Val(rs!IdBcoCta)
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing

End Sub

Private Sub CmbMonPago_Click()
Call Cargar_Cuenta_Banco
End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
If rsTelebanking.RecordCount = 0 Then
    Txtsemana.Text = ""
Else
    Cmbtipotrabajador.Enabled = False
    Txtsemana.Enabled = False
    UpDown1.Enabled = False
End If
Cmbdel.Enabled = False
Cmbal.Enabled = False
Label5.Visible = False
Label6.Visible = False
Cmbdel.Visible = False
Cmbal.Visible = False


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
   Cmbdel.Enabled = True
   Cmbal.Enabled = True

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
End If
Cmbtipotrabajador_Click
Procesa_Semana
End Sub


Private Sub Cmbtipotrabajador_Click()
VTipotrab = fc_CodigoComboBox(Cmbtipotrabajador, 2)
Dim wciamae As String
Dim wBeginMonth As String

wciamae = Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where cod_maestro2='" & VTipotrab & "' and status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)

If VTipo = "01" Or VTipo = "05" Or VTipo = "11" Then
    If VTipotrab <> "" Then
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
                
                Sql$ = "select iniciomes from cia where cod_cia='" & wcia & "' and status<>'*'"
                If (fAbrRst(rs, Sql$)) Then
                   If IsNull(rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = rs!iniciomes
                End If
                rs.Close
                
                If Trim(wBeginMonth) = "" Then
                    MsgBox "Ingrese el Inicio Del Mes", vbInformation, ""
                Exit Sub
                End If
                
'                Cmbfecha.Month = Month(Date)
'                Cmbfecha.Year = Year(Date)
                If Trim(wBeginMonth) <> "1" Then
                   Cmbfecha.Day = Val(wBeginMonth) - 1
                Else
                   Cmbfecha.Day = Val(fMaxDay(Month(Date), Year(Date)))
                End If
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

If rs.State = 1 Then rs.Close
Procesa_Semana
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    'If Not TypeOf Screen.ActiveControl Is DataGrid Then
        SendKeys "{TAB}"
    'Else
        'Dgrdcabeza_DblClick
    'End If
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Call fc_Descrip_Maestros2("01007", "", CmbBcoPago, False, True)
Call fc_Descrip_Maestros2_Mon("01006", "", CmbMonPago)
For I = 0 To CmbMonPago.ListCount - 1
    If Right(Left(CmbMonPago.List(I), 4), 3) = wmoncont Then CmbMonPago.ListIndex = I: Exit For
Next



Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")

Cmbfecha.Year = Year(Date)
Cmbfecha.Month = Month(Date)
Cmbfecha.Day = Day(Date)
Me.dtpFechaPago.Value = Format(Now, "dd/mm/yyyy")

Cmbdel.Year = Year(Date)
Cmbdel.Month = Month(Date)
Cmbdel.Day = Day(Date)

Cmbal.Year = Year(Date)
Cmbal.Month = Month(Date)
Cmbal.Day = Day(Date)

Call Crea_Rs
End Sub

Private Sub GrdSolicitud_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 11 Then
    If Not rsTelebanking.EOF Then
        If rsTelebanking!Action = True Then
            lblNro.Caption = Val(lblNro.Caption) + 1
            lblImporte.Caption = Format((Val(Format(lblImporte.Caption, "###0.00")) + rsTelebanking!totneto), "#,##0.00")
            
        Else
            lblNro.Caption = Val(lblNro.Caption) - 1
            lblImporte.Caption = Format((Val(Format(lblImporte.Caption, "###0.00")) - rsTelebanking!totneto), "#,##0.00")
            
        End If
    End If
End If
End Sub

Private Sub Txtsemana_Change()
 Procesa_Semana
End Sub

Private Sub Txtsemana_KeyPress(KeyAscii As Integer)
Txtsemana.Text = Txtsemana.Text + fc_ValNumeros(KeyAscii)
End Sub

Private Sub UpDown1_DownClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "00"
If Txtsemana.Text > 0 Then Txtsemana.Text = Format(Val(Txtsemana.Text - 1), "00")
End Sub

Private Sub UpDown1_UpClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "00"
Txtsemana.Text = Format(Val(Txtsemana.Text + 1), "00")


End Sub
Private Sub Crea_Rs()

If rsTelebanking.State = 1 Then rsTelebanking.Close
    rsTelebanking.Fields.Append "TipoRegistro", adChar, 1, adFldIsNullable
    rsTelebanking.Fields.Append "IdBoleta", adChar, 2, adFldIsNullable
    rsTelebanking.Fields.Append "Boleta", adChar, 40, adFldIsNullable
    rsTelebanking.Fields.Append "pagotipcta", adChar, 2, adFldIsNullable
    rsTelebanking.Fields.Append "TipoCuenta", adVarChar, 100, adFldIsNullable
    rsTelebanking.Fields.Append "Sucursal", adVarChar, 5, adFldIsNullable
    rsTelebanking.Fields.Append "NroCuenta", adVarChar, 50, adFldIsNullable
    rsTelebanking.Fields.Append "PlaCod", adChar, 8, adFldIsNullable
    rsTelebanking.Fields.Append "Trabajador", adVarChar, 500, adFldIsNullable
    rsTelebanking.Fields.Append "TipoMoneda", adVarChar, 5, adFldIsNullable
    rsTelebanking.Fields.Append "totneto", adDouble, , adFldIsNullable
    rsTelebanking.Fields.Append "id_boleta", adBigInt, , adFldIsNullable
    rsTelebanking.Fields.Append "Action", adBoolean, , adFldIsNullable
    rsTelebanking.Open
    Set GrdSolicitud.DataSource = rsTelebanking
End Sub

Private Sub Procesa_Semana()
VIdTelebanking = 0
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

End Sub

Public Sub Nuevo()
Call Crea_Rs
dtpFechaPago.Value = Format(Now, "dd/mm/yyyy")

For I = 0 To CmbMonPago.ListCount - 1
    If Right(Left(CmbMonPago.List(I), 4), 3) = wmoncont Then CmbMonPago.ListIndex = I: Exit For
Next

Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
Cmbfecha.Year = Year(Date)
Cmbfecha.Month = Month(Date)
Cmbfecha.Day = Day(Date)
Cmbdel.Year = Year(Date)
Cmbdel.Month = Month(Date)
Cmbdel.Day = Day(Date)

Cmbal.Year = Year(Date)
Cmbal.Month = Month(Date)
Cmbal.Day = Day(Date)
lblCodigo.Caption = ""
VIdTelebanking = 0

Call Habilitar_Control(True)

End Sub

Private Sub Habilitar_Control(xBol As Boolean)
optPendiente.Enabled = Not xBol
optCancelado.Enabled = Not xBol
optPendiente.Value = xBol

CmbBcoPago.Enabled = xBol
dtpFechaPago.Enabled = xBol
CmbMonPago.Enabled = xBol
cboCuenta.Enabled = xBol
Cmbtipo.Enabled = xBol
Cmbtipotrabajador.Enabled = xBol
UpDown1.Enabled = xBol
Txtsemana.Enabled = xBol
btnQuitar.Enabled = xBol
chkTodos.Enabled = xBol

End Sub

Public Sub Cargar_Planilla(IdTeleBanking As Integer)
Dim rsTemp As New ADODB.Recordset
Dim I As Integer

Sql$ = "exec usp_Pla_PlaTelebanking @IdTelebanking=" & Val(VIdTelebanking)
Sql$ = Sql$ & ",@Tipo=4"
If (fAbrRst(rsTemp, Sql$)) Then
    dtpFechaPago.Value = Format(rsTemp!FechaProceso, "dd/mm/yyyy")
    Call rUbiIndCmbBox(CmbBcoPago, rsTemp!banco, "00")
    For I = 0 To CmbMonPago.ListCount - 1
        If Right(Left(CmbMonPago.List(I), 4), 3) = rsTemp!banco Then CmbMonPago.ListIndex = I: Exit For
    Next
    
    For I = 0 To Me.cboCuenta.ListCount - 1
        If Trim(cboCuenta.List(I)) = rsTemp!NroCta Then cboCuenta.ListIndex = I: Exit For
    Next
    
    Call rUbiIndCmbBox(Cmbtipotrabajador, rsTemp!TipoTrabajador, "00")
    
    lblImporte.Caption = Format(rsTemp!importe, "#,##0.00")
    
    
End If
rsTemp.Close
Set rsTemp = Nothing
End Sub

Public Sub Procesar_Boletas()

If Cmbtipo.ListIndex < 0 Then
    MsgBox "Seleccionar el Tipo de Boleta", vbCritical, Me.Caption
    Exit Sub
End If

If Cmbtipotrabajador.ListIndex < 0 Then
    MsgBox "Seleccionar el Tipo de Trabajador", vbCritical, Me.Caption
    Exit Sub
End If

If VTipotrab = "02" Then
    If Val(Txtsemana.Text) <= 0 Then
        MsgBox "Ingresar la Semana", vbCritical, Me.Caption
        Exit Sub
    End If
End If

If rsTelebanking.RecordCount > 0 Then
    If Not rsTelebanking.EOF Then rsTelebanking.Update
    GrdSolicitud.Update
    rsTelebanking.MoveFirst
    rsTelebanking.Filter = "IdBoleta='" & VTipo & "'"
    If rsTelebanking.RecordCount > 0 Then
        MsgBox "Ya se agregó el Tipo de Boleta seleccionada", vbCritical, Me.Caption
        Exit Sub
    End If
    rsTelebanking.Filter = adFilterNone
    rsTelebanking.Filter = "IdBoleta = '02' or IdBoleta='03'"
    
    If rsTelebanking.RecordCount > 0 Then
        rsTelebanking.Filter = adFilterNone
        MsgBox "No puede combinar la Boleta de Vacaciones y Gratificaciones con otros Tipos de Boletas", vbCritical, Me.Caption
        Exit Sub
    End If
    rsTelebanking.Filter = adFilterNone
    
    If VTipo = "02" Or VTipo = "03" Then
        MsgBox "No puede combinar la Boleta de Vacaciones y Gratificaciones con otros Tipos de Boletas", vbCritical, Me.Caption
        Exit Sub
    End If
End If

rsTelebanking.Filter = adFilterNone

Sql$ = "usp_Pla_Procesar_Telebanking_BoletaPago "
Sql$ = Sql$ & "'" & wcia
Sql$ = Sql$ & "','" & VTipo
Sql$ = Sql$ & "','" & VTipotrab
Sql$ = Sql$ & "'," & CStr(dtpFechaPago.Year)
Sql$ = Sql$ & "," & CStr(dtpFechaPago.Month)
If VTipotrab = "02" Then
    If VTipo <> "02" Then
        Sql$ = Sql$ & ",'" & Format(Trim(Me.Txtsemana.Text), "00")
    Else
        Sql$ = Sql$ & ",'"
    End If
Else
    Sql$ = Sql$ & ",'"
End If
Sql$ = Sql$ & "','" & Format(VBcoPago, "00")
Sql$ = Sql$ & "','" & Trim(VMoneda)
Sql$ = Sql$ & "','" & Format(dtpFechaPago.Value, "mm/dd/yyyy")
Sql$ = Sql$ & "','" & Format(Cmbdel.Value, "mm/dd/yyyy")
Sql$ = Sql$ & "','" & Format(Cmbal.Value, "mm/dd/yyyy") & "'"

If (fAbrRst(rs, Sql$)) Then
    rs.MoveFirst
    Do While Not rs.EOF
        rsTelebanking.AddNew
        With rsTelebanking
            !TipoRegistro = rs!TipoRegistro
            !IdBoleta = Format(VTipo, "00")
            !Boleta = rs!Boleta
            !pagotipcta = rs!pagotipcta
            !TipoCuenta = rs!TipoCuenta
            !Sucursal = rs!Sucursal
            !NroCuenta = rs!NroCuenta
            !TipoRegistro = rs!TipoRegistro
            !PlaCod = rs!PlaCod
            !Trabajador = rs!Trabajador
            !TipoMoneda = rs!TipoMoneda
            !totneto = rs!totneto
            !id_boleta = rs!id_boleta
            !Action = True
            .Update
    
        End With
        rs.MoveNext
    Loop
    rsTelebanking.MoveFirst
'    GrdSolicitud.Refresh

End If

Dim rsTemp As New ADODB.Recordset
Dim suma As Currency
Set rsTemp = rsTelebanking.Clone
rsTemp.Filter = "Action=1"
If rsTemp.RecordCount > 0 Then
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        suma = suma + rsTemp!totneto
        rsTemp.MoveNext
    Loop
End If
Me.lblNro.Caption = rsTemp.RecordCount
Me.lblTotal.Caption = rsTelebanking.RecordCount
Me.lblImporte.Caption = Format(suma, "#,##0.00")
rsTemp.Close
Set rsTemp = Nothing

End Sub

