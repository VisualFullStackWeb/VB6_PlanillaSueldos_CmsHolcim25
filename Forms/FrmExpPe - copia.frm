VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmExpPe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Exportar Datos Planilla Electrónica (PDT 601 v1.91  (09/05/2012) ) «"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   Icon            =   "FrmExpPe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optMasivas 
      Caption         =   "Bajas masivas"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   21
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   10560
         TabIndex        =   15
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox TxtAño 
         Enabled         =   0   'False
         Height          =   315
         Left            =   9705
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CmbMes 
         Height          =   315
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox CmbCia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo"
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
         Left            =   6720
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog Box 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraError 
      Caption         =   "Mensajes "
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
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   5520
      Width           =   11415
      Begin MSComctlLib.ListView LstError 
         Height          =   1455
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mensajes"
            Object.Width           =   18414
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cod"
            Object.Width           =   18
         EndProperty
      End
      Begin VB.CommandButton CmdOpen 
         Caption         =   "Abrir carpeta"
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
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1995
         Width           =   2055
      End
      Begin VB.Label LblRuta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1995
         Width           =   8895
      End
   End
   Begin VB.Frame FraGrd 
      Caption         =   "Exportar solo las elegidas"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   11415
      Begin TrueOleDBGrid70.TDBGrid Grd 
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   17
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Sel"
         Columns(0).DataField=   "add"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Estructura a importar"
         Columns(1).DataField=   "estructura"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   "fecfin"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=926"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=17674"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=17595"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(12)=   "Column(1).WrapText=1"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=370"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=291"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
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
         Caption         =   "Archivo a Exportar"
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
         _StyleDefs(14)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Arial"
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
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15,.alignment=3"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.wraptext=-1,.locked=-1"
         _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(51)  =   "Named:id=33:Normal"
         _StyleDefs(52)  =   ":id=33,.parent=0"
         _StyleDefs(53)  =   "Named:id=34:Heading"
         _StyleDefs(54)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   ":id=34,.wraptext=-1"
         _StyleDefs(56)  =   "Named:id=35:Footing"
         _StyleDefs(57)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(58)  =   "Named:id=36:Selected"
         _StyleDefs(59)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=37:Caption"
         _StyleDefs(61)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(62)  =   "Named:id=38:HighlightRow"
         _StyleDefs(63)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(64)  =   "Named:id=39:EvenRow"
         _StyleDefs(65)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(66)  =   "Named:id=40:OddRow"
         _StyleDefs(67)  =   ":id=40,.parent=33"
         _StyleDefs(68)  =   "Named:id=41:RecordSelector"
         _StyleDefs(69)  =   ":id=41,.parent=34"
         _StyleDefs(70)  =   "Named:id=42:FilterBar"
         _StyleDefs(71)  =   ":id=42,.parent=33"
      End
   End
   Begin VB.OptionButton optMasivas 
      Caption         =   "Altas masivas"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   20
      Top             =   1560
      Width           =   1335
   End
   Begin MSForms.Frame Frame4 
      Height          =   495
      Left            =   2880
      OleObjectBlob   =   "FrmExpPe.frx":030A
      TabIndex        =   19
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   11415
      Begin VB.CheckBox chkEsPlame 
         Caption         =   "Versión PLAME"
         ForeColor       =   &H80000002&
         Height          =   195
         Left            =   250
         TabIndex        =   18
         Top             =   810
         Width           =   1575
      End
      Begin ComctlLib.ProgressBar Barra 
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.CommandButton CmdExp 
         Caption         =   "Exportar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton OptElige 
         Caption         =   "Exportar solo las elegidas"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "Exportar todas las estructuras"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
End
Attribute VB_Name = "FrmExpPe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim dFechaSOL As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim xRucCia As String
Dim xRazSoc As String
Dim xRutaFile As String
Dim xPeriodoAño As String
Dim xPeriodoMes As String
Dim xUnTrabajador As String
Dim xFechaLimiteFinPeriodo As String
Dim RsConcep As New ADODB.Recordset

Const vSeparador = "|"

Private Sub chkEsPlame_Click()
 If chkEsPlame.Value = 0 Then optMasivas(0).Value = False: optMasivas(1).Value = False
End Sub

Private Sub Cmbmes_Click()
chkEsPlame.Value = IIf(Cmbmes.ListIndex + 1 >= 10 And Val(TxtAño.Text) >= 2012, 1, 0)
End Sub
Private Sub CmdExp_Click()
LstError.ListItems.Clear

xPeriodoAño = TxtAño.Text
xPeriodoMes = Format(Cmbmes.ListIndex + 1, "00")

xFechaLimiteFinPeriodo = Format(fMaxDay(Val(xPeriodoMes), Val(xPeriodoAño)) & "/" & xPeriodoMes & "/" & xPeriodoAño, "mm/dd/yyyy") & " 11:59:59pm"

If OptElige(1).Value = True Then
    Dim J As Integer
    J = 0
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If !Add = True Then J = J + 1
                .MoveNext
            Loop
        End If
    End With
    If J = 0 Then
        MsgBox "Elija por lo menos una estructura a exportar", vbExclamation, Me.Caption
        Me.Grd.SetFocus
        Exit Sub
    End If
End If

With rs
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                If OptAll(0).Value = True Or (rs!Add = True And Me.OptElige(1).Value = True) Then
                    Select Case rs!NRO
                        Case 1: Exp_Estructura01
                        Case 2: Exp_Estructura02: Exp_Estructura02_b (30)
                        Case 3: Exp_Estructura03
                        Case 4: Exp_Estructura04
                        Case 5: Exp_Estructura05
                        Case 6: Exp_Estructura06
                        Case 7: Exp_Estructura07
                        Case 8: Exp_Estructura08
                        Case 9: Exp_Estructura09
                        Case 10: Exp_Estructura10
                        Case 11: Exp_Estructura11
                        Case 12: Exp_Estructura12
                        Case 13: Exp_Estructura13
                        Case 14: Exp_Estructura14
                        Case 15: Exp_Estructura15
                        Case 16: Exp_Estructura16
                        Case 17: Exp_Estructura17
                        Case 18: Exp_Estructura18
                        Case 19: Exp_Estructura19
                        Case 20: Exp_Estructura20
                        Case 21: Exp_Estructura21
                        'Case 22: Exp_Estructura22
                        Case 23: Exp_Estructura23
                        Case 24: Exp_Estructura24
                        Case Else
                            MsgBox "Proceso Exportar a no implementado", vbExclamation, Me.Caption
                    End Select
                End If
            .MoveNext
        Loop
    End If
End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdOpen_Click()
Call ShellExecute(hWnd, "Open", App.Path & "\reports\", "", "", vbNormalFocus)
End Sub

Private Sub Form_Load()
'xUnTrabajador = "E0003"
'xUnTrabajador = "O3506"
xUnTrabajador = ""

Me.Top = 0
Me.Left = 0
Me.Width = 11730 '10650 ' 9135
Me.Height = 7900 '7830
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Crea_Rs
Dim Rq As ADODB.Recordset
Sql = "select ruc,razsoc from cia where cod_cia='" & wcia & "'"
If fAbrRst(Rq, Sql) Then
    xRucCia = Trim(Rq!ruc)
    xRazSoc = Trim(Rq!razsoc)
End If
Rq.Close
Set Rq = Nothing
Cmbmes.AddItem "ENERO"
Cmbmes.AddItem "FEBRERO"
Cmbmes.AddItem "MARZO"
Cmbmes.AddItem "ABRIL"
Cmbmes.AddItem "MAYO"
Cmbmes.AddItem "JUNIO"
Cmbmes.AddItem "JULIO"
Cmbmes.AddItem "AGOSTO"
Cmbmes.AddItem "SETIEMBRE"
Cmbmes.AddItem "OCTUBRE"
Cmbmes.AddItem "NOVIEMBRE"
Cmbmes.AddItem "DICIEMBRE"

If Month(Date) = 1 Then
    TxtAño.Text = CStr(Year(Date) - 1)
    Cmbmes.ListIndex = 11
Else
    TxtAño.Text = CStr(Year(Date))
    Cmbmes.ListIndex = Month(Date) - 2
End If


If Val(xPeriodoMes) = 1 Then
    xPeriodoAño = Val(TxtAño.Text) - 1
    xPeriodoMes = "12"
Else
    xPeriodoAño = TxtAño.Text
    xPeriodoMes = Format(Cmbmes.ListIndex + 1, "00")
End If

LblRuta.Caption = "Ubicación de Archivos Exportados: " & App.Path & "\Reports\"
End Sub
Public Sub Crea_Rs()
    
    
    
    If RsConcep.State = 1 Then RsConcep.Close
    RsConcep.Fields.Append "id", adChar, 5, adFldIsNullable
    RsConcep.Fields.Append "tipo", adChar, 2, adFldIsNullable
    RsConcep.Fields.Append "idconcepto", adChar, 3, adFldIsNullable
    RsConcep.Fields.Append "nomconcepto", adVarChar, 100, adFldIsNullable
    RsConcep.Fields.Append "codsunat", adChar, 4, adFldIsNullable
    RsConcep.Open
    
    Carga_Conceptos_Placonstantes
   'canteras
    If rs.State = 1 Then rs.Close
    rs.Fields.Append "add", adBoolean, 10, adFldIsNullable
    rs.Fields.Append "estructura", adVarChar, 200, adFldIsNullable
    rs.Fields.Append "nro", adInteger, 4, adFldIsNullable
    'Rs.Fields.Append "cod_suspension", adChar, 2, adFldIsNullable
    rs.Open
    Set Grd.DataSource = rs
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA1: Datos de Establecimientos Propios"
    rs!NRO = 1
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 2: Datos de empresas a quienes destaco o desplazo personal"
    rs!NRO = 2
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 3: Datos de empresas que me destacan o desplazan personal"
    rs!NRO = 3
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 4: Datos principales del trabajador, pensionista, prestador de servicios-cuarta categoría, prestador de servicios-modalidades formativas y personal de terceros"
    rs!NRO = 4
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 5: Datos del trabajador"
    rs!NRO = 5
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 6: Datos del pensionista"
    rs!NRO = 6
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 7: Datos del  prestador de servicios - cuarta categoría"
    rs!NRO = 7
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 8: Datos de suspensión de cuarta categoría"
    rs!NRO = 8
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 9: Datos del prestador de servicios -  modalidad formativa"
    rs!NRO = 9
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 10: Datos del personal de terceros"
    rs!NRO = 10
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 11: Datos de períodos"
    rs!NRO = 11
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 12: Datos de otros empleadores"
    rs!NRO = 12
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 13: Importar Datos de derechohabientes - ALTA"
    rs!NRO = 13
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 14: Importar Datos de la jornada laboral por trabajador"
    rs!NRO = 14
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 15: Datos de los días subsidiados del trabajador"
    rs!NRO = 15
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 16: Datos de los días no trabajados y no subsidiados del trabajador"
    rs!NRO = 16
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 17: Datos de los establecimientos donde labora el trabajador"
    rs!NRO = 17
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 18: Datos del detalle de la remuneración del trabajador"
    rs!NRO = 18
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 19: Datos del detalle de la remuneración del pensionista"
    rs!NRO = 19
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 20: Datos del detalle de comprobantes de prestadores de servicios - cuarta categoria"
    rs!NRO = 20

    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 21: Datos del detalle de comprobantes de los prestadores de servicios - modalidad formativa"
    rs!NRO = 21
    
    
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 23: Datos de los establecimientos - Prestador de servicios - Modalidad Formativa"
    rs!NRO = 23
    
    'ADD JCMS 240211
    rs.AddNew
    rs!ESTRUCTURA = "ESTRUCTURA 24: Importar Datos de derechohabientes - BAJA"
    rs!NRO = 24
    
'    Rs.AddNew
'    Rs!ESTRUCTURA = "ESTRUCTURA 22: Datos del detalle de personal de terceros - SCTR"
'    Rs!NRO = 22
    
End Sub

Private Sub Grd_AfterColUpdate(ByVal ColIndex As Integer)
If ColIndex = 0 Then Grd.Update
End Sub

Private Sub Grd_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
On Error Resume Next:
If Trim(Grd.Splits(Split).Columns(0).CellValue(Bookmark)) = True Then
    RowStyle.BackColor = &HEDDCDC          '&HC0FFFF
Else
    RowStyle.BackColor = vbWhite
End If
End Sub

Public Sub Exp_Estructura04()
'//**** ESTRUCTURA 4: "Datos principales del trabajador, pensionista, prestador de servicios-cuarta categoría, prestador de servicios-modalidades formativas y personal de terceros"

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
Dim blnDelPeriodo As Boolean
I = 0
xRutaFile = IIf(chkEsPlame.Value = 0, App.Path & "\reports\" & xRucCia & ".t00", App.Path & "\reports\" & "RP_" & xRucCia & ".ide")
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "',4," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
 If chkEsPlame.Value = 0 Then
    Do While Not Rq.EOF
        If Trim(Rq!categoria) = "4ta" Then 'solo para 4ta categoria
                        
            '---RODA Reg = "06" 'tipo doc Regimen unico del contribuyente
            '---RODA Reg = Reg & vSeparador & Trim(Left(Rq!RUC & "", 15))
            
            Reg = "01" 'tipo doc Regimen unico del contribuyente
            Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
            
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_pat & "", 40)))
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_mat & "", 40)))
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!nombres & "", 40)))
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador & Trim(Rq!sexo & "")
            Reg = Reg & vSeparador & Trim(Left(Rq!nacionalidad & "", 4))
            Reg = Reg & vSeparador & Trim(Left(Rq!telefono & "", 10))
            Reg = Reg & vSeparador & Trim(Left(Rq!Email & "", 50))
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador & IIf(Trim(Rq!domiciliado & "") = True, "1", "0")
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
        Else
            Reg = Trim(Left(Rq!tipo_doc & "", 2))
            Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_pat & "", 40)))
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_mat & "", 40)))
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!nombres & "", 40)))
            Reg = Reg & vSeparador & Trim(Format(Rq!fnacimiento & "", "dd/mm/yyyy"))
            Reg = Reg & vSeparador & Trim(Rq!sexo & "")
            Reg = Reg & vSeparador & Trim(Left(Rq!nacionalidad & "", 4))
            Reg = Reg & vSeparador & Trim(Left(Rq!telefono & "", 10))
            Reg = Reg & vSeparador & Trim(Left(Rq!Email & "", 50))
            Reg = Reg & vSeparador & Trim(Rq!essaludvida & "")
            Reg = Reg & vSeparador & IIf(Trim(Rq!domiciliado & "") = True, "1", "0")
            Reg = Reg & vSeparador '& Trim(Left(Rq!cod_via & "", 2))
            Reg = Reg & vSeparador '& Trim(Left(Rq!nomvia & "", 20))
            Reg = Reg & vSeparador '& Trim(Left(Rq!nrokmmza & "", 4))
            Reg = Reg & vSeparador '& Trim(Left(Rq!intdptolote & "", 4))
            Reg = Reg & vSeparador '& Trim(Left(Rq!cod_zona & "", 2))
            Reg = Reg & vSeparador '& Trim(Left(Rq!nomzona & "", 20))
            Reg = Reg & vSeparador '& Trim(Left(Rq!referencia & "", 40))
            Reg = Reg & vSeparador '& Trim(Left(Rq!ubigeo & "", 6))
            Reg = Reg & vSeparador
        
        End If
        
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
 Else
    'If optMasivas(1).Value = False Then
        Do While Not Rq.EOF
            If Trim(Rq!categoria) <> "4ta" Then 'solo trabajadores
                Reg = ""
                If optMasivas(0).Value = False Then
                 blnDelPeriodo = IIf(Format(Trim(Rq!fIngreso) & "", "yyyymm") = Format(xPeriodoAño, "0000") & Format(xPeriodoMes, "00") And Format(Trim(Rq!fcese) & "", "yyyymm") = "", 1, 0)
                Else
                    blnDelPeriodo = True
                End If
                
                If blnDelPeriodo Then
                    Reg = Trim(Left(Rq!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
                    Reg = Reg & vSeparador
                    Reg = Reg & vSeparador & Trim(Format(Rq!fnacimiento & "", "dd/mm/yyyy"))
                    Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_pat & "", 40)))
                    Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_mat & "", 40)))
                    Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!nombres & "", 40)))
                    Reg = Reg & vSeparador & Trim(Rq!sexo & "")
                    Reg = Reg & vSeparador & Trim(Left(Rq!nacionalidad & "", 4))
                    Reg = Reg & vSeparador 'Codigo larga distancia
                    Reg = Reg & vSeparador '& Trim(Left(Rq!telefono & "", 9))
                    Reg = Reg & vSeparador & Trim(Left(Rq!Email & "", 50))
                    Reg = Reg & vSeparador '& Trim(Left(Rq!cod_via & "", 2))
                    Reg = Reg & vSeparador '& Trim(Left(Rq!nomvia & "", 20))
                    Reg = Reg & vSeparador '& Trim(Left(Rq!nrokmmza & "", 4))
                    Reg = Reg & vSeparador 'Departamento
                    Reg = Reg & vSeparador 'Interior
                    Reg = Reg & vSeparador 'Mz
                    Reg = Reg & vSeparador 'Lt
                    Reg = Reg & vSeparador 'Km
                    Reg = Reg & vSeparador 'Block
                    Reg = Reg & vSeparador 'Etapa
                    Reg = Reg & vSeparador '& Trim(Left(Rq!cod_zona & "", 2))
                    Reg = Reg & vSeparador '& Trim(Left(Rq!nomzona & "", 20))
                    Reg = Reg & vSeparador '& Trim(Left(Rq!referencia & "", 40))
                    Reg = Reg & vSeparador '& Trim(Left(Rq!ubigeo & "", 6))
                    
                    Reg = Reg & vSeparador 'D2 & Trim(Left(Rq!cod_via & "", 2))
                    Reg = Reg & vSeparador 'D2 & Trim(Left(Rq!nomvia & "", 20))
                    Reg = Reg & vSeparador 'D2 & Trim(Left(Rq!nrokmmza & "", 4))
                    Reg = Reg & vSeparador 'D2 Departamento
                    Reg = Reg & vSeparador 'D2 Interior
                    Reg = Reg & vSeparador 'D2 Mz
                    Reg = Reg & vSeparador 'D2 Lt
                    Reg = Reg & vSeparador 'D2 Km
                    Reg = Reg & vSeparador 'D2 Block
                    Reg = Reg & vSeparador 'D2 Etapa
                    Reg = Reg & vSeparador 'D2 & Trim(Left(Rq!cod_zona & "", 2))
                    Reg = Reg & vSeparador 'D2 & Trim(Left(Rq!nomzona & "", 20))
                    Reg = Reg & vSeparador 'D2 & Trim(Left(Rq!referencia & "", 40))
                    Reg = Reg & vSeparador 'D2 & Trim(Left(Rq!ubigeo & "", 6))
                    Reg = Reg & vSeparador 'Indicador centro asistencial
                    Reg = Reg & vSeparador 'fin
                    Print #1, Reg
                End If
            
            End If

            I = I + 1
            Barra.Value = I
            Rq.MoveNext
        Loop
    'End If

 End If

    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0

End Sub

Sub DeleteAFile(especificaciondearchivo As String)
  Dim fso
  Set fso = CreateObject("Scripting.FileSystemObject")
  fso.DeleteFile (especificaciondearchivo)
End Sub






Public Sub Add_Mensaje(ByRef oBj As ListView, ByVal pMsj As String, ByVal pBold As Boolean, ByVal pColor As Long)
Dim itmX As ListItem
Set itmX = oBj.ListItems.Add(, , pMsj)
itmX.Bold = pBold
itmX.ForeColor = pColor
itmX.SubItems(1) = rs!NRO
End Sub
Public Sub Exp_Estructura05()
'/*** ESTRUCTURA 5: "Datos del trabajador"

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
Dim blnDelPeriodo As Boolean

I = 0
xRutaFile = IIf(chkEsPlame.Value = 0, App.Path & "\reports\" & xRucCia & ".t01", App.Path & "\reports\" & "RP_" & xRucCia & ".tra")
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "',5," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
 Barra.Max = Rq.RecordCount
 Barra.Min = 1 - 1
 If chkEsPlame.Value = 0 Then
    Do While Not Rq.EOF
    
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!tip_trabajador & "", 4))
        Reg = Reg & vSeparador & "1" 'PRIVADO
        Reg = Reg & vSeparador & Trim(Left(Rq!nivel_educativo & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!ocupacion & "", 6))
        Reg = Reg & vSeparador & IIf(Trim(Rq!discapacidad & "") = True, "1", "0")
        Reg = Reg & vSeparador & Trim(Left(Rq!reg_pensionario, 2))
        Reg = Reg & vSeparador & Trim(Format(Rq!afpfechaafil & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador & Trim(Left(Rq!NUMAFP & "", 12))
        Reg = Reg & vSeparador & Trim(Left(Rq!sctr_salud & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!sctr_pension & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!Tipo_contrato & "", 2))
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_reg_alternativo & "") = True, "1", "0")
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_jornada_trab_max & "") = True, "1", "0")
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_hor_nocturno & "") = True, "1", "0")
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_otros_ing_5ta & "") = True, "1", "0")
        Reg = Reg & vSeparador & Trim(Left(Rq!sindicalizado & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!trab_periodicidad_remuneracion & "", 1))
        Reg = Reg & vSeparador & IIf(Trim(Rq!afiliado_eps_serv & "") = True, "1", "0")
               
        If Rq!afiliado_eps_serv = False Then
            Reg = Reg & vSeparador
        Else
            Reg = Reg & vSeparador & Trim(Rq!codigo_eps & "")
        End If

        Reg = Reg & vSeparador & Trim(Left(Rq!situacion_eps & "", 2))
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_5ta_exonerada_inafecta & "") = True, "1", "0")
        Reg = Reg & vSeparador & Trim(Left(Rq!trab_situacion_especial & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_pago & "", 1))
        
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_afiliacion_asegura_tu_pension & "") = True, "1", "0")
        Reg = Reg & vSeparador
        Reg = Reg & vSeparador & Trim(Left(Rq!trab_evita_doble_tribu & "", 2))
        Reg = Reg & vSeparador
        
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
   
 Else
 
 If optMasivas(1).Value = False Then
 
    Do While Not Rq.EOF
        Reg = ""
        If optMasivas(0).Value = False Then
         blnDelPeriodo = IIf(Format(Trim(Rq!fIngreso) & "", "yyyymm") = Format(xPeriodoAño, "0000") & Format(xPeriodoMes, "00") And Format(Trim(Rq!fcese) & "", "yyyymm") = "", 1, 0)
        Else
            blnDelPeriodo = True
        End If
        
        If blnDelPeriodo Then

            Reg = Trim(Left(Rq!tipo_doc & "", 2))
            
            Reg = Reg & vSeparador
            Reg = Reg & Trim(Left(Rq!nro_doc & "", 15))
            Reg = Reg & vSeparador & "604" 'País emisor
            
            Reg = Reg & vSeparador & "01" 'PRIVADO
            Reg = Reg & vSeparador & Trim(Left(Rq!nivel_educativo & "", 2))
            Reg = Reg & vSeparador & Trim(Left(Rq!ocupacion & "", 6))
            Reg = Reg & vSeparador & IIf(Trim(Rq!discapacidad & "") = True, "1", "0")
            Reg = Reg & vSeparador & Trim(Left(LTrim(Rq!NUMAFP) & "", 12))
            Reg = Reg & vSeparador & Trim(Left(Rq!sctr_pension & "", 1))
            Reg = Reg & vSeparador & Trim(Left(Rq!Tipo_contrato & "", 2))
            Reg = Reg & vSeparador & IIf(Trim(Rq!trab_reg_alternativo & "") = True, "1", "0")
            Reg = Reg & vSeparador & IIf(Trim(Rq!trab_jornada_trab_max & "") = True, "1", "0")
            Reg = Reg & vSeparador & IIf(Trim(Rq!trab_hor_nocturno & "") = True, "1", "0")
            Reg = Reg & vSeparador & Trim(Left(Rq!sindicalizado & "", 1))
            Reg = Reg & vSeparador & Trim(Left(Rq!trab_periodicidad_remuneracion & "", 1))
            Reg = Reg & vSeparador & Format(Rq!basico_inicial, "#######.00")
            
            'T-REGISTRO
            Reg = Reg & vSeparador & Trim(Left(Rq!situacion_t & "", 1))
            
            Reg = Reg & vSeparador & IIf(Trim(Rq!trab_5ta_exonerada_inafecta & "") = True, "1", "0")
            Reg = Reg & vSeparador & Trim(Left(Rq!trab_situacion_especial & "", 1))
            Reg = Reg & vSeparador & Trim(Left(Rq!tipo_pago & "", 1))
            Reg = Reg & vSeparador & Trim(Rq!cat_ocupacional)  'Categoria trabajador
            Reg = Reg & vSeparador & IIf(Trim(Left(Rq!trab_evita_doble_tribu & "", 2)) = "", "0", Trim(Left(Rq!trab_evita_doble_tribu & "", 2)))
            Reg = Reg & vSeparador 'Ruc CAS
            Reg = Reg & vSeparador
            
    '        Reg = Reg & vSeparador & Trim(Left(Rq!reg_pensionario, 2))
    '        Reg = Reg & vSeparador & Trim(Format(Rq!afpfechaafil & "", "dd/mm/yyyy"))
    '        Reg = Reg & vSeparador & Trim(Left(Rq!sctr_salud & "", 1))
    '        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_otros_ing_5ta & "") = True, "1", "0")
    '        Reg = Reg & vSeparador & IIf(Trim(Rq!afiliado_eps_serv & "") = True, "1", "0")
    '        If Rq!afiliado_eps_serv = False Then
    '            Reg = Reg & vSeparador
    '        Else
    '            Reg = Reg & vSeparador & Trim(Rq!codigo_eps & "")
    '        End If
    '        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_afiliacion_asegura_tu_pension & "") = True, "1", "0")
            
            Print #1, Reg
        End If
        
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
  End If
  
 End If
   
 Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & "  ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub



Public Sub Exp_Estructura12()
'//*** ESTRUCTURA 12: 'Datos de otros empleadores"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".o00"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!ruc & "", 11))
        Reg = Reg & vSeparador & Trim(Left(Rq!razsoc & "", 100))
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub

Public Sub Exp_Estructura13_ANTES240211()
'//*** ESTRUCTURA 13: "Importar Datos de derechohabientes"

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".der"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_docdh & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!numero & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!ap_pat & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!ap_mat & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!nombres & "", 40))
        Reg = Reg & vSeparador & Trim(Format(Rq!fec_nacdh & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador & Trim(Rq!sexodh & "")
        Reg = Reg & vSeparador & Trim(Rq!vinculodh & "")
        
        If Trim(Rq!vinculodh & "") = "4" Then 'solo gestante
            Reg = Reg & vSeparador & Trim(Rq!tipdoc_acredita_paternidad & "")
            Reg = Reg & vSeparador & Trim(Rq!nrodoc_acredita_paternidad & "")
        Else
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
        End If
        Reg = Reg & vSeparador & Trim(Left(Rq!ACTIVO_BAJA & "", 2))
        'If Rq!nro_doc = "08341815" Then Stop
        If Trim(Left(Rq!ACTIVO_BAJA & "", 2)) = "10" Then 'alta
            Reg = Reg & vSeparador & Trim(Format(Rq!fecha_alta & "", "dd/mm/yyyy"))
        Else
            Reg = Reg & vSeparador
        End If
        If Trim(Left(Rq!ACTIVO_BAJA & "", 2)) = "11" Then 'baja
            Reg = Reg & vSeparador & Trim(Left(Rq!motivo_baja & "", 1))
            Reg = Reg & vSeparador & Trim(Format(Rq!fecha_baja & "", "dd/mm/yyyy"))
        Else
            Reg = Reg & vSeparador
            Reg = Reg & vSeparador
        End If
        If Trim(Rq!vinculodh & "") = "1" And Trim(Rq!nrocertificado) <> "" Then 'hijo
            Reg = Reg & vSeparador & Trim(Left(Rq!nrocertificado & "", 20))
        Else
            Reg = Reg & vSeparador
        End If
        Reg = Reg & vSeparador & Trim(Left(Rq!domicilio & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_via & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!NOM_VIA & "", 20))
        Reg = Reg & vSeparador & Trim(Left(Rq!NRO & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!Interior & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_zona & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!NOM_ZONA & "", 20))
        Reg = Reg & vSeparador & Trim(Left(Rq!referencia & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!ubigeo & "", 6))
        Reg = Reg & vSeparador
        
        
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub

Public Sub Exp_Estructura15()
'//*** ESTURCTURA 15: "Datos de los días subsidiados del trabajador"

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = IIf(chkEsPlame.Value = 0, App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".sub", App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".snl")
'xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".sub"

Dim Rq As ADODB.Recordset

If chkEsPlame.Value = 1 Then
    Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Else
    Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & "115" & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
End If

Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_subsidio & "", 2))
        
        If chkEsPlame.Value = 1 Then
            Reg = Reg & vSeparador & Rq!dias_suspension
        Else
            Reg = Reg & vSeparador & Trim(Left(Rq!nrocitt & "", 16))
            Reg = Reg & vSeparador & Trim(Format(Rq!FecIni & "", "dd/mm/yyyy"))
            Reg = Reg & vSeparador & Trim(Format(Rq!FecFin & "", "dd/mm/yyyy"))
        End If
        
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub

Private Sub LstError_ItemClick(ByVal Item As MSComctlLib.ListItem)
BuscarItems Item.SubItems(1)
End Sub

Private Sub OptAll_Click(index As Integer)
LstError.ListItems.Clear
FraGrd.Enabled = False
LimpiarCheck
End Sub

Private Sub OptElige_Click(index As Integer)
LstError.ListItems.Clear
FraGrd.Enabled = True
LimpiarCheck
End Sub

Private Sub TxtAño_Change()
    chkEsPlame.Value = IIf(Cmbmes.ListIndex + 1 >= 4 And Val(TxtAño.Text) >= 2012, 1, 0)
End Sub

Private Sub UpDown1_DownClick()
If Val(TxtAño.Text) > 2007 Then
    TxtAño.Text = Val(TxtAño.Text) - 1
End If
End Sub

Private Sub UpDown1_UpClick()
TxtAño.Text = Val(TxtAño.Text) + 1
End Sub

Public Sub BuscarItems(ByVal pId As String)
Me.Grd.SelBookmarks.Clear
If rs.RecordCount > 0 Then
    rs.MoveFirst
    rs.FIND "nro ='" & pId & "'", 0, 1, 1
     Me.Grd.SelBookmarks.Add Grd.Bookmark
     Grd.Refresh
End If


End Sub


Public Sub Exp_Estructura06()
'//*** ESTRUCTURA 6: "Datos del pensionista"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".t02"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_pensionista & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!reg_pensionario & "", 2))
        Reg = Reg & vSeparador & Trim(Format(Rq!afpfechaafil & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador & Trim(Left(Rq!NUMAFP & "", 12))
        Reg = Reg & vSeparador & Trim(Left(Rq!situacion_eps & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_pago & "", 1))
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Sub Exp_Estructura07()
'//*** ESTRUCTURA 7: "Datos del  prestador de servicios - cuarta categoría"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = IIf(chkEsPlame.Value = 0, App.Path & "\reports\" & xRucCia & ".t03", App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".ps4")
'---xRutaFile = App.Path & "\reports\" & xRucCia & ".t03"

Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        If chkEsPlame.Value = 1 Then
            Reg = "06" 'RUC proveedor 4ta cat. domiciliado
        Else
            Reg = Trim(Left(Rq!tipo_doc & "", 2))
            Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        End If
        Reg = Reg & vSeparador & Trim(Left(Rq!ruc & "", 11))
        If chkEsPlame.Value = 1 Then
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_pat & "", 40)))
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!ap_mat & "", 40)))
            Reg = Reg & vSeparador & fCadPrint(Trim(Left(Rq!nombres & "", 40)))
            Reg = Reg & vSeparador & IIf(Trim(Rq!domiciliado & "") = True, "1", "0")
        End If
        
        Reg = Reg & vSeparador & IIf(Trim(Rq!doble_tribu & "") <> "", Trim(Rq!doble_tribu & ""), "0")
        
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub





Public Sub Exp_Estructura08()
'//*** ESTRUCTURA 8: "Datos de suspensión de cuarta categoría"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".s00"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        'Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = "06" 'TIPO DOC RUC
        Reg = Reg & vSeparador & Trim(Left(Rq!ruc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!num_oper & "", 15))
        Reg = Reg & vSeparador & Format(Trim(Rq!fec_susp & ""), "dd/mm/yyyy")
        Reg = Reg & vSeparador & Trim(Left(Rq!ejercicio & "", 4))
        If Trim(Left(Rq!med_pres & "", 1)) = "I" Then 'internet
            Reg = Reg & vSeparador & "1"
        Else
            Reg = Reg & vSeparador & "2"
        End If
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub




Public Sub Exp_Estructura09()
'//*** ESTRUCTURA 9: "Datos del prestador de servicios -  modalidad formativa"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".t04"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        
        Reg = Reg & vSeparador & Trim(Left(Rq!seguro_medico & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!niveleducativo & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!ocupacion & "", 6))
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_madre_resp_familiar & "") = True, "1", "0")
        Reg = Reg & vSeparador & IIf(Trim(Rq!discapacidad & "") = True, "1", "0")
        Reg = Reg & vSeparador & Trim(Left(Rq!centro_formacion_profesional & "", 1))
        Reg = Reg & vSeparador & IIf(Trim(Rq!trab_hor_nocturno & "") = True, "1", "0")
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub



Public Sub Exp_Estructura10()
'//*** ESTRUCTURA 10: "Datos del personal de terceros"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".t05"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        
        Reg = Reg & vSeparador & Trim(Left(Rq!ruc & "", 11))
        Reg = Reg & vSeparador & Trim(Left(Rq!sctr_salud & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!sctr_pension & "", 1))
        
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub
Public Sub Exp_Estructura11()
'//*** ESTRUCTURA 11: "Datos de períodos"
On Error GoTo MsgErr:
Dim Reg As String, RegInicial As String
Dim I As Integer
Dim blnSoloActivo As Boolean
I = 0

xRutaFile = IIf(chkEsPlame.Value = 0, App.Path & "\reports\" & xRucCia & ".p00", App.Path & "\reports\" & "RP_" & xRucCia & ".per")

'---xRutaFile = App.Path & "\reports\" & xRucCia & ".p00"


Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    
    If chkEsPlame.Value = 0 Then
        'P.E V 1.91
        Do While Not Rq.EOF
            
            Reg = Trim(Left(Rq!tipo_doc & "", 2))
            Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
            
            Reg = Reg & vSeparador & Trim(Left(Val(Rq!cat_trab) & "", 1))
            
            Reg = Reg & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy")
            Reg = Reg & vSeparador & Format(Trim(Rq!fcese & ""), "dd/mm/yyyy")
            If Val(Rq!cat_trab) = 1 Or Val(Rq!cat_trab) = 2 Then ' trabajadores y pensionistas
                Reg = Reg & vSeparador & Trim(Left(Rq!mot_fin_periodo & "", 2))
            Else
                Reg = Reg & vSeparador
            End If
            If Val(Rq!cat_trab) = 5 Then ' modalidad formativa
                Reg = Reg & vSeparador & Trim(Left(Rq!modalidad_formativa & "", 2))
            Else
                Reg = Reg & vSeparador
            End If
            Reg = Reg & vSeparador
            
            Print #1, Reg
            I = I + 1
            Barra.Value = I
            Rq.MoveNext
        Loop
    Else
    'PLAME
        Do While Not Rq.EOF
            Reg = ""
            If optMasivas(0).Value = False And optMasivas(1).Value = False Then
              Reg = Trim(Left(Rq!tipo_doc & "", 2))
              Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
              Reg = Reg & vSeparador & "604"  'País emisor
              Reg = Reg & vSeparador & Trim(Left(Val(Rq!cat_trab) & "", 1))
            
              blnSoloActivo = True
              If Format(Trim(Rq!fIngreso) & "", "yyyymm") = Format(xPeriodoAño, "0000") & Format(xPeriodoMes, "00") And Format(Trim(Rq!fcese) & "", "yyyymm") = "" Then
                 'Altas
                  RegInicial = Reg
                  Reg = Reg & vSeparador & "1" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(4, vSeparador)
                  Print #1, Reg
                  Reg = RegInicial
                  'Tipo de trabajador
                  Reg = Reg & vSeparador & "2" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!tip_trabajador) & String(2, vSeparador)
                  Print #1, Reg
                  Reg = RegInicial
                  'Regimen de aseguramiento Essalud
                  Reg = Reg & vSeparador & "3" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!reg_aseguramiento_salud) & String(2, vSeparador)
                  
                  Print #1, Reg
                  Reg = RegInicial
                  'Regimen pensionario
                  Reg = Reg & vSeparador & "4" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!reg_pensionario) & String(2, vSeparador)
                  
                  Print #1, Reg
                  Reg = RegInicial
                  'SCTR Salud
                  Reg = Reg & vSeparador & "5" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!sctr_salud) & String(2, vSeparador)
                  Print #1, Reg
                  blnSoloActivo = False
              End If
              
              If Format(Trim(Rq!fcese) & "", "yyyymm") = Format(xPeriodoAño, "0000") & Format(xPeriodoMes, "00") And Format(Trim(Rq!fcese) & "", "yyyymm") <> Format(Trim(Rq!fIngreso) & "", "yyyymm") Then
                 'Bajas
                  Reg = Reg & vSeparador & "1" & vSeparador & vSeparador & Format(Trim(Rq!fcese & ""), "dd/mm/yyyy")
                  If Val(Rq!cat_trab) = 1 Or Val(Rq!cat_trab) = 2 Then ' trabajadores y pensionistas
                      Reg = Reg & vSeparador & Trim(Left(Rq!mot_fin_periodo & "", 2))
                  Else
                      Reg = Reg & vSeparador
                  End If
                  Reg = Reg & String(2, vSeparador)
                  
                  blnSoloActivo = False
              End If
              
              If blnSoloActivo Then
                Reg = Reg & String(4, vSeparador)
              End If
            Else
            'Cargas masivas
               If optMasivas(0).Value = True And Format(Trim(Rq!fcese) & "", "yyyymm") = "" Then
               'Altas
                Reg = Trim(Left(Rq!tipo_doc & "", 2))
                Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
                Reg = Reg & vSeparador & "604"  'País emisor
                Reg = Reg & vSeparador & Trim(Left(Val(Rq!cat_trab) & "", 1))
                
                RegInicial = Reg
                Reg = Reg & vSeparador & "1" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(4, vSeparador)
                Print #1, Reg
                Reg = RegInicial
                'Tipo de trabajador
                Reg = Reg & vSeparador & "2" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!tip_trabajador) & String(2, vSeparador)
                Print #1, Reg
                Reg = RegInicial
                'Regimen de aseguramiento Essalud
                Reg = Reg & vSeparador & "3" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!reg_aseguramiento_salud) & String(2, vSeparador)
                
                Print #1, Reg
                Reg = RegInicial
                'Regimen pensionario
                Reg = Reg & vSeparador & "4" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!reg_pensionario) & String(2, vSeparador)
                
                Print #1, Reg
                Reg = RegInicial
                'SCTR Salud
                Reg = Reg & vSeparador & "5" & vSeparador & Format(Trim(Rq!fIngreso & ""), "dd/mm/yyyy") & String(2, vSeparador) & Trim(Rq!sctr_salud) & String(2, vSeparador)
                Print #1, Reg
               End If
                
               If optMasivas(1).Value = True And Format(Trim(Rq!fcese) & "", "yyyymm") <> "" Then
                'Bajas
                Reg = Trim(Left(Rq!tipo_doc & "", 2))
                Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
                Reg = Reg & vSeparador & "604"  'País emisor
                Reg = Reg & vSeparador & Trim(Left(Val(Rq!cat_trab) & "", 1))
                
                Reg = Reg & vSeparador & "1" & vSeparador & vSeparador & Format(Trim(Rq!fcese & ""), "dd/mm/yyyy")
                If Val(Rq!cat_trab) = 1 Or Val(Rq!cat_trab) = 2 Then ' trabajadores y pensionistas
                    Reg = Reg & vSeparador & Trim(Left(Rq!mot_fin_periodo & "", 2))
                Else
                    Reg = Reg & vSeparador
                End If
                Reg = Reg & String(2, vSeparador)
                Print #1, Reg
               End If
            End If
            I = I + 1
            Barra.Value = I
            Rq.MoveNext
            
        Loop
    
    
    End If
    
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub



Public Sub Exp_Estructura14_ANTES_040808()
'//*** ESTRUCTURA 14: "Importar Datos de la jornada laboral por trabajador"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
'xUnTrabajador = "O5544"
xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".jor"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    
    Dim xMaxDay As Integer
    xMaxDay = fMaxDay(Val(xPeriodoMes), Val(xPeriodoAño))
    Do While Not Rq.EOF
    
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        'If CInt(Rq!nro_dias_trabajados & "") > 31 Then
        If CInt(Rq!nro_dias_trabajados & "") > xMaxDay Then
            'Reg = Reg & vSeparador & "31"
            Reg = Reg & vSeparador & CStr(xMaxDay)
        Else
            Reg = Reg & vSeparador & CInt(Rq!nro_dias_trabajados & "")
        End If
        Reg = Reg & vSeparador & CInt(Rq!nro_horas_ordinarias_trabajadas & "")
        Reg = Reg & vSeparador & CInt(Rq!H_EXTRAS & "")
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Sub Exp_Estructura01()
'//*** ESTRUCTURA1: "Datos de Establecimientos Propios"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".esp"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_establecimiento & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_establecimiento & "", 4))
        
        Reg = Reg & vSeparador & Trim(Left(Rq!nom_establecimiento & "", 40))
        Reg = Reg & vSeparador & IIf(Trim(Rq!indicador_centro_riesgo & "") = True, "1", "0")
        
        If Rq!indicador_centro_riesgo = False Then 'NO ES CENTRO DE RIESGO
            Reg = Reg & vSeparador
        Else
            Reg = Reg & vSeparador & Format(Rq!Tasa, "##0.00")
        End If
        
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub

Public Sub Exp_Estructura02()
'//*** ESTRUCTURA 2: "Datos de empresas a quienes destaco o desplazo personal"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".edd"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Reg = Trim(Left(Rq!ruc & "", 11))
        Reg = Reg & vSeparador & Trim(Left(Rq!razsoc & "", 100))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_serv & "", 6))
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub
Public Sub Exp_Estructura02_b(ByVal pNro As Integer)
'//*** ESTRUCTURA 2: "Datos de empresas a quienes destaco o desplazo personal - establecimientos"
'On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".sdd"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & pNro & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!ruc & "", 11))
        Reg = Reg & vSeparador & Trim(Left(Rq!nom_establecimiento & "", 40))
        Reg = Reg & vSeparador & IIf(Trim(Rq!indicador_centro_riesgo & "") = True, "1", "0")
        
        If Rq!indicador_centro_riesgo = False Then 'NO ES CENTRO DE RIESGO
            Reg = Reg & vSeparador
        Else
            Reg = Reg & vSeparador & Format(Rq!Tasa, "##0.00")
        End If
        
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & " (B): ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & " (B): no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub




Public Sub Exp_Estructura03()
'//*** ESTRUCTURA 3: "Datos de empresas que me destacan o desplazan personal"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\" & xRucCia & ".med"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Reg = Trim(Left(Rq!ruc & "", 11))
        Reg = Reg & vSeparador & Trim(Left(Rq!razsoc & "", 100))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_serv & "", 6))
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub





Public Sub Exp_Estructura16()
'//*** ESTURCTURA 15: "Datos de los días subsidiados del trabajador"

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".not"
Dim Rq As ADODB.Recordset

If chkEsPlame.Value = 1 Then
    Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Else
    Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & "116" & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
End If

Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_diasno_laborados & "", 2))
        Reg = Reg & vSeparador & Trim(Format(Rq!FecIni & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador & Trim(Format(Rq!FecFin & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Sub Exp_Estructura20()
'//*** ESTRUCTURA 20: "Datos del detalle de comprobantes de prestadores de servicios - cuarta categoria

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".4ta"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
    
        If chkEsPlame.Value = 1 Then
            Reg = "06" 'RUC proveedor 4ta cat. domiciliado
            Reg = Reg & vSeparador & Trim(Left(Rq!ruc & "", 11))
        Else
            Reg = Trim(Left(Rq!tipo_doc & "", 2))
            Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        End If
        Reg = Reg & vSeparador & "R" 'recibo honorarios
        Reg = Reg & vSeparador & IIf(chkEsPlame.Value = 1, Trim(Mid(Rq!serie_comprob & "", 2, 3)), Trim(Mid(Rq!serie_comprob & "", 1, 4)))
        Reg = Reg & vSeparador & Trim(Left(Val(Rq!nro_comprob & ""), 8))
        Reg = Reg & vSeparador & Format(Rq!Total, "#######.00")
               
        Reg = Reg & vSeparador & Trim(Format(Rq!fecha_doc & "", "dd/mm/yyyy"))
        
        '//** ojo para forzar que la cancelacion del doc. sea en el periodo
        'If Not (Year(CDate(Format(Rq!fecha_cancela & "", "dd/mm/yyyy"))) = Val(xPeriodoAño) And Month(CDate(Format(Rq!fecha_cancela & "", "dd/mm/yyyy"))) = Val(xPeriodoMes)) Then
        '    Reg = Reg & vSeparador & Trim(Format(xFechaLimiteFinPeriodo, "dd/mm/yyyy"))
        'Else
           Reg = Reg & vSeparador & Trim(Format(Rq!fecha_cancela & "", "dd/mm/yyyy"))
        'End If
        If Rq!impuesto <> 0 Then 'impuesto
            Reg = Reg & vSeparador & "1"
        Else
            Reg = Reg & vSeparador & "0"
        End If
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub
Public Sub Exp_Estructura17()
'//*** ESTRUCTURA 17: "Datos de los establecimientos donde labora el trabajador"

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
Dim blnDelPeriodo As Boolean

I = 0

xRutaFile = IIf(chkEsPlame.Value = 0, App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".tes", App.Path & "\reports\" & "RP_" & xRucCia & ".est")
'--xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".tes"

Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    
    If optMasivas(1).Value = False Then
        Do While Not Rq.EOF
            Reg = ""
            If optMasivas(0).Value = False And chkEsPlame.Value = 1 Then
                blnDelPeriodo = IIf(Format(Trim(Rq!fIngreso) & "", "yyyymm") = Format(xPeriodoAño, "0000") & Format(xPeriodoMes, "00") And Format(Trim(Rq!fcese) & "", "yyyymm") = "", 1, 0)
            Else
                blnDelPeriodo = True
            End If
            If blnDelPeriodo Then
                Reg = Trim(Left(Rq!tipo_doc & "", 2))
                Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
                If chkEsPlame.Value = 1 Then Reg = Reg & vSeparador & "604"  'País emisor
                Reg = Reg & vSeparador & Trim(Left(Rq!ruc & "", 11))
                Reg = Reg & vSeparador & Trim(Left(Rq!cod_establecimiento & "", 4))
                
                If chkEsPlame.Value = 0 Then
                    If Rq!PORC = 0 Then
                        Reg = Reg & vSeparador
                    Else
                        Reg = Reg & vSeparador & Format(Rq!PORC, "###.00")
                    End If
                End If
                Reg = Reg & vSeparador
                Print #1, Reg
            End If
            
            I = I + 1
            Barra.Value = I
            Rq.MoveNext
        Loop
    End If
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Sub Carga_Conceptos_Placonstantes()
Dim Sql As String
Sql = "Select tipomovimiento,codinterno,codsunat,descripcion from placonstante where cia='" & wcia & "' and tipomovimiento in ('02','03') and status<>'*'"
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
    With RsConcep
        Do While Not Rq.EOF
                .AddNew
                !Id = Trim(Rq!tipomovimiento) & Trim(Rq!codinterno)
                !tipo = Trim(Rq!tipomovimiento)
                !idconcepto = Trim(Rq!codinterno)
                !nomconcepto = Trim(Rq!Descripcion)
                !CODSUNAT = Trim(Rq!CODSUNAT)
            Rq.MoveNext
        Loop
    End With
End If

With RsConcep
        .AddNew
        !Id = "03111"
        !tipo = "03"
        !idconcepto = "111"
        !nomconcepto = "Aporte Afp"
        !CODSUNAT = "0608"
        
        .AddNew
        !Id = "03112"
        !tipo = "03"
        !idconcepto = "112"
        !nomconcepto = "Seguro Afp"
        !CODSUNAT = "0611"
        
        
        .AddNew
        !Id = "03114"
        !tipo = "03"
        !idconcepto = "114"
        !nomconcepto = "Comision Afp"
        !CODSUNAT = "0601"
        
End With
                
Rq.Close
Set Rq = Nothing
End Sub


Public Sub Exp_Estructura18_antes()
'//*** ESTRUCTURA 18: "Datos del detalle de la remuneración del trabajador"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".rem"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", "")
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            If Rq.Fields(intloop).Value <> 0 Then
                Reg = Trim(Left(Rq!tipo_doc & "", 2))
                Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
                'If Rq!nro_doc = "41838045" Then Stop
                If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                    xTipo = "02" 'remuneraciones
                Else
                    xTipo = "03" 'deducciones
                End If
                Reg = Reg & vSeparador & Buscar_CodSunat(RsConcep, xTipo & Right(Rq.Fields(intloop).Name, 2))
                Reg = Reg & vSeparador & Format(Rq.Fields(intloop).Value, "#######.00")
                Reg = Reg & vSeparador & Format(Rq.Fields(intloop).Value, "#######.00")
                Reg = Reg & vSeparador
                Print #1, Reg
            End If
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Function Buscar_CodSunat(ByVal pRs As ADODB.Recordset, ByVal pId As String) As String
Dim Rc As New ADODB.Recordset
Set Rc = pRs.Clone
With Rc
    If .RecordCount > 0 Then
          .MoveFirst
          .FIND "id='" & Trim(pId) & "'", 0, 1, 1
          If .EOF Then
            Buscar_CodSunat = ""
            Debug.Print "falta codigo de sunat=>" & Trim(pId)
          Else
            Buscar_CodSunat = Trim(.Fields("codsunat") & "")
            If Trim(.Fields("codsunat") & "") = "" Then
                Debug.Print "falta codigo de sunat=>" & Trim(pId) & ""
            End If
          End If
    End If
End With
Rc.Close
Set Rc = Nothing
End Function

Public Sub Exp_Estructura19_antes()
'//*** ESTRUCTURA 19: "Datos del detalle de la remuneración del pensionista"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".pen"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", "")
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            If Rq.Fields(intloop).Value <> 0 Then
                Reg = Trim(Left(Rq!tipo_doc & "", 2))
                Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
                If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                    xTipo = "02" 'remuneraciones
                Else
                    xTipo = "03" 'deducciones
                End If
                Reg = Reg & vSeparador & Buscar_CodSunat(RsConcep, xTipo & Right(Rq.Fields(intloop).Name, 2))
                Reg = Reg & vSeparador & Format(Rq.Fields(intloop).Value, "#######.00")
                Reg = Reg & vSeparador & Format(Rq.Fields(intloop).Value, "#######.00")
                Reg = Reg & vSeparador
                Print #1, Reg
            End If
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Sub Exp_Estructura21()
'//*** ESTRUCTURA 21: "Datos del detalle de comprobantes de los prestadores de servicios - modalidad formativa"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".for"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Format(Rq!importe, "#######.00")
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Sub LimpiarCheck()
  With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                !Add = False
                .MoveNext
            Loop
        End If
    End With
End Sub


Public Sub Exp_Estructura22()

'//*** ESTRUCTURA 22: "Datos del detalle de personal de terceros - SCTR"
'On Error GoTo MsgErr:


If Not Cuadro_Dialogo_Abrir("*.xls") Then
    Exit Sub
End If

Screen.MousePointer = 11
Dim Cnx As ADODB.Connection
Set Cnx = New ADODB.Connection
Dim RsEx As New ADODB.Recordset
Dim I As Integer
I = 0
With Cnx
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source=" & Box.FileName & ";" & _
    "Extended Properties=Excel 8.0;"
    .Open
End With
    Sql = "Select * from [Hoja1$]"
    RsEx.Open Sql, Cnx, adOpenDynamic, adLockOptimistic
    If RsEx.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existen registros a importar", vbExclamation, Me.Caption
        RsEx.Close
        Set RsEx = Nothing
        Cnx.Close
        Set Cnx = Nothing
        Exit Sub
    End If
    cn.BeginTrans
    Sql = "update platercerossctr set status='*',user_modi='" & Trim(wuser) & "',fec_modi=getdate() where cod_cia='" & wcia & "' and año=" & xPeriodoAño & " and mes=" & xPeriodoMes & " and status<>'*'"
    cn.Execute Sql, 64
    If Not RsEx.EOF Then
        
        RsEx.MoveFirst
        RsEx.MoveLast
        RsEx.MoveFirst
        'Barra.Max = RsEx.RecordCount
        'Barra.Min = 1 - 1
        Do While Not RsEx.EOF
                Sql = "insert into platercerossctr (cod_cia,año,mes,tipo_doc,nro_doc,cod_establecimiento,tasa_salud,scrt_salud,tasa_pension,scrt_pension,status,user_crea,fec_crea,user_modi,fec_modi)"
                Sql = Sql & " values('" & wcia & "'," & xPeriodoAño & "," & xPeriodoMes & ",'" & RsEx(0) & "','" & RsEx(1) & "','" & RsEx(2) & "'," & CCur(RsEx(3)) & "," & CCur(RsEx(4)) & "," & CCur(RsEx(5)) & "," & CCur(RsEx(6)) & ",'','" & Trim(wuser) & "',getdate(),null,null)"
                cn.Execute Sql, 64
                I = I + 1
               ' Barra.Value = i
            RsEx.MoveNext
        Loop
        'Barra.Value = 0
    End If
    RsEx.Close
    Set RsEx = Nothing
    cn.CommitTrans
    Cnx.Close
    Set Cnx = Nothing
    



Dim Reg As String


xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".sct"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", "")
Open xRutaFile For Output As #1

If fAbrRst(Rq, Sql) Then
    I = 0
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
    
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 3 To Rq.Fields.count - 1
            If intloop = 3 Or intloop = 5 Then
                If Rq.Fields(3).Value <> 0 Or Rq.Fields(5).Value <> 0 Then
                    Reg = Trim(Left(Rq!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Trim(Left(Rq!cod_establecimiento & "", 4))
                    Reg = Reg & vSeparador & Format(Rq.Fields(intloop).Value, "###.00")
                    Reg = Reg & vSeparador & Format(Rq.Fields(intloop + 1).Value, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                End If
            End If
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub



Public Function Cuadro_Dialogo_Abrir(pextension As String) As Boolean
 On Error GoTo ErrHandler
   ' Establece los filtros.
   Box.CancelError = True
   Select Case pextension
    Case "*.txt"
        Box.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt|"
    Case "*.dbf"
        Box.Filter = "All Files (*.*)|*.*|Tablas Files (*.dbf)|*.dbf|"
    Case "*.mdb"
        Box.Filter = "All Files (*.*)|*.*|BD Access (*.mdb)|*.mdb|"
    Case "*.csv"
        Box.Filter = "All Files (*.*)|*.*|Microsoft Excel (*.csv)|*.csv|"
    Case "*.xls"
        Box.Filter = "Excel files (*.xls)" '|*.xls|Microsoft Excel(*.xls)|*.xls)
        '"All Files (*.*)|*.*|Microsoft Excel 97/2000 (*.xls)|*.txt)"
   End Select
   ' Especifique el filtro predeterminado.
   Box.FilterIndex = 2
   'Box.FileName = "terceros_scrt.xls"
   Box.InitDir = App.Path
   ' Presenta el cuadro de diálogo Abrir.
   'Box.FileTitle = "Archivo a elegir terceros_scrt.xls"
   Box.ShowOpen
   ' Llamada al procedimiento para abrir archivo.
   Dim pos As String
   'CTA.CTE MN:
   'vNroBco
   
   
   
   Dim swExiste As Variant
   swExiste = InStr(1, UCase(Trim(Box.FileName)), UCase("terceros_scrt"), vbTextCompare)
   If swExiste = 0 Then
      MsgBox "Archivo Elegido no es el correcto" & Chr(13) & "El Correcto es terceros_scrt.xls", vbCritical, "Importacion"
      'salir = True
    Else
      Cuadro_Dialogo_Abrir = True
    End If
   Exit Function

ErrHandler:
   Cuadro_Dialogo_Abrir = False
   'El usuario hizo clic en el botón Cancelar.
   Exit Function
End Function


Public Sub Exp_Estructura18_antesv12_070808_modi_101208()
'//*** ESTRUCTURA 18: "Datos del detalle de la remuneración del trabajador"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
Dim xMsgUser As String

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".rem"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then

    Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    Dim xId As String
    Dim xCodSunat As String
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    Else
                        xTipo = "03" 'deducciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                End With
            End If
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    With rsRem
        If .RecordCount > 0 Then
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            .MoveFirst
            Do While Not .EOF
                    Reg = Trim(Left(!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                    I = I + 1
                    Barra.Value = I
                .MoveNext
            Loop
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1



Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
If Trim(xMsgUser) <> "" Then
    Add_Mensaje LstError, Trim(xMsgUser), True, vbRed
Else
    Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
End If
Screen.MousePointer = 0
End Sub

Public Sub Exp_Estructura19_ANTESV12_070809_MODI_101208()
'//*** ESTRUCTURA 19: "Datos del detalle de la remuneración del pensionista"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".pen"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then

    Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    Dim xId As String
    Dim xCodSunat As String
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then   'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
                 xId = "": xCodSunat = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    Else
                        xTipo = "03" 'deducciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                End With
            End If
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    With rsRem
        If .RecordCount > 0 Then
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            .MoveFirst
            Do While Not .EOF
                    Reg = Trim(Left(!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                    I = I + 1
                    Barra.Value = I
                .MoveNext
            Loop
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
        
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1



Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Private Function QuitarBasura(pCadena) As String
Dim POS1, POS2 As Integer
Dim Cadena As String

Cadena = ""
Dim I As Integer

For I = 1 To Len(pCadena)
    If Asc(Mid(pCadena, I, 1)) <> 34 And Not (Asc(Mid(pCadena, I, 1)) >= 128 And Asc(Mid(pCadena, I, 1)) <= 255) Then
        Cadena = Cadena & Mid(pCadena, I, 1)
    End If
Next I

QuitarBasura = Cadena
End Function

Public Function fCadPrint(ByRef pVal As String) As String
   Dim mPos As Integer
   Do
     mPos = InStr(pVal, Chr(165))
     If mPos > 0 Then pVal = Mid(pVal, 1, mPos - 1) & "Ñ" & Mid(pVal, mPos + 1)
   Loop Until (mPos = 0)
   fCadPrint = pVal
End Function



Public Sub Exp_Estructura14()
'//*** ESTRUCTURA 14: "Importar Datos de la jornada laboral por trabajador"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
'xUnTrabajador = "O5544"
xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".jor"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Dim xNroHrs_Ordinarias As Integer
Dim xNroHrs_Extras As Integer
Dim xNroHrs_Ordinarias_Decimal As Currency
Dim xNroHrs_Extras_Decimal As Currency

Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    
    Dim xMaxDay As Integer
    xMaxDay = fMaxDay(Val(xPeriodoMes), Val(xPeriodoAño))
    Do While Not Rq.EOF
        xNroHrs_Ordinarias = 0
        xNroHrs_Ordinarias_Decimal = 0
        
        xNroHrs_Extras = 0
        xNroHrs_Extras_Decimal = 0
    
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        'If CInt(Rq!nro_dias_trabajados & "") > 31 Then
'        If CInt(Rq!nro_dias_trabajados & "") > xMaxDay Then
'            'Reg = Reg & vSeparador & "31"
'            Reg = Reg & vSeparador & CStr(xMaxDay)
'        Else
'            Reg = Reg & vSeparador & CInt(Rq!nro_dias_trabajados & "")
'        End If
        'If Rq!nro_doc = "06777745" Then Stop
        xNroHrs_Ordinarias = Rq!nro_horas_ordinarias_trabajadas_entero + (Rq!nro_horas_ordinarias_trabajadas_decimal \ 60)
        xNroHrs_Ordinarias_Decimal = Rq!nro_horas_ordinarias_trabajadas_decimal Mod 60
                
        Reg = Reg & vSeparador & xNroHrs_Ordinarias
        Reg = Reg & vSeparador & xNroHrs_Ordinarias_Decimal
        
        
        xNroHrs_Extras = Rq!H_EXTRAS_entero + (Rq!H_EXTRAS_minutos_decimal \ 60)
        xNroHrs_Extras_Decimal = Rq!H_EXTRAS_minutos_decimal Mod 60
        
        Reg = Reg & vSeparador & xNroHrs_Extras
        Reg = Reg & vSeparador & xNroHrs_Extras_Decimal
        
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub



Public Sub Exp_Estructura23()
'add jcms 040808
'//*** ESTRUCTURA 23: "Datos de los establecimientos - Prestador de servicios - Modalidad Formativa"

On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".mfe"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!ruc & "", 11))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_establecimiento & "", 4))
        Reg = Reg & vSeparador
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


Public Function NombreConcepto_Plahistorico(ByVal pCodCia As String, ByVal pIdConcepto As String) As String
Dim Sql As String
Sql = "usp_pla_nombre_conceptos_boleta '" & pCodCia & "','" & Trim(pIdConcepto) & "'"
Dim Rq As ADODB.Recordset
If fAbrRst(Rq, Sql) Then
    NombreConcepto_Plahistorico = Trim(Rq(0)) & " - " & Trim(Rq(1))
Else
    MsgBox "Concepto " & pIdConcepto & "  de la tabla Plahistorico no encontrado", vbCritical, Me.Caption
    NombreConcepto_Plahistorico = ""
End If

Rq.Close
Set Rq = Nothing


End Function


Public Sub Exp_Estructura18_antes_050809()
'/*ULIMO CAMBIO ADD JCMS 10/12/08*/
'//*** ESTRUCTURA 18: "Datos del detalle de la remuneración del trabajador"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
Dim xMsgUser As String

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".rem"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then

    Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    Dim xId As String
    Dim xCodSunat As String
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    Else
                        xTipo = "03" 'deducciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Then
                        GoTo SaltarSunat:
                    End If
                    
                    
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                    
                End With
            End If
            
SaltarSunat:
            
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    With rsRem
        If .RecordCount > 0 Then
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            .MoveFirst
            Do While Not .EOF
                    Reg = Trim(Left(!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                    I = I + 1
                    Barra.Value = I
                .MoveNext
            Loop
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1



Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
If Trim(xMsgUser) <> "" Then
    Add_Mensaje LstError, Trim(xMsgUser), True, vbRed
Else
    Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
End If
Screen.MousePointer = 0
End Sub

Public Sub Exp_Estructura19()
'/*ULIMO CAMBIO ADD JCMS 10/12/08*/
'//*** ESTRUCTURA 19: "Datos del detalle de la remuneración del pensionista"
On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".pen"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then

    Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    Dim xId As String
    Dim xCodSunat As String
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then   'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
                 xId = "": xCodSunat = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    Else
                        xTipo = "03" 'deducciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Then
                        GoTo SaltarSunat:
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                End With
            End If

SaltarSunat:

        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    With rsRem
        If .RecordCount > 0 Then
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            .MoveFirst
            Do While Not .EOF
                    Reg = Trim(Left(!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                    I = I + 1
                    Barra.Value = I
                .MoveNext
            Loop
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
        
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1



Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub



Public Sub Exp_Estructura18_ANTES_090110()
'/*ULIMO CAMBIO ADD JCMS 05/08/09*/
'/*jcbr solicito acumular las boletas de gratificacion en el concepto de sunat 0406 excepto el codigo 0312*/

'//*** ESTRUCTURA 18: "Datos del detalle de la remuneración del trabajador"
'On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
Dim xMsgUser As String

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".rem"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "',0"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then

    Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    Dim xId As String
    Dim xCodSunat As String
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    Else
                        xTipo = "03" 'deducciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Then
                        GoTo SaltarSunat:
                    End If
                    
                    
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                    
                End With
            End If
            
SaltarSunat:
            
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    With rsRem
        If .RecordCount > 0 Then
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            .MoveFirst
            Do While Not .EOF
                    Reg = Trim(Left(!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                    I = I + 1
                    Barra.Value = I
                .MoveNext
            Loop
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0

'/*SOLO PARA GRATIFICACION*/

Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "',1"
If fAbrRst(Rq, Sql) Then
    'Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    'Dim xId As String
    'Dim xCodSunat As String
    
    xId = "": xCodSunat = ""
    
    Barra.Max = Rq.RecordCount
    Barra.Min = 0
    Do While Not Rq.EOF
        'Dim intloop  As Integer
        intloop = 0
        'Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            'If Rq.Fields(intloop).Name = "a20" Then Stop
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
            
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    ElseIf Left(Rq.Fields(intloop).Name, 1) = "d" Then
                        xTipo = "03" 'deducciones y aportaciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
                    'If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Or xTipo = "03" Then
                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Or (xTipo = "03" And Left(Rq.Fields(intloop).Name, 1) = "d") Then
                        GoTo SaltarSunat2:
                    End If
                    
                    
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                    
                End With
            End If
            
SaltarSunat2:
            
        Next intloop
        I = I + 1
        Barra.Value = 1
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    Dim xAcumGratificacion As Currency
    xAcumGratificacion = 0
    
    
    Dim stipdoc As String
    Dim sNumdoc As String
    stipdoc = ""
    sNumdoc = ""
    
    With rsRem
        If .RecordCount > 0 Then
            '***Print #1, "GRATI"
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            xAcumGratificacion = 0
            .MoveFirst
            Do While Not .EOF
                    If (stipdoc & sNumdoc <> Trim(!tipo_doc & "") & Trim(Left(!nro_doc & "", 15))) And Trim(stipdoc & sNumdoc) <> "" Then
                        
                            Reg = stipdoc
                            Reg = Reg & vSeparador & Trim(Left(sNumdoc, 15))
                            Reg = Reg & vSeparador & "0406" 'Right(!cod_sunat, 4)
                            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
                            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
                            Reg = Reg & vSeparador
                            Print #1, Reg ' & "***"
                            xAcumGratificacion = 0
                    End If
            
                    'If Right(!cod_sunat, 4) = "0312" Then '/*bonificacion extraordinaria*/
                    'add jcms 090110 se agrego scrt
                    If Right(!cod_sunat, 4) = "0312" Or Right(!cod_sunat, 4) = "0809" Or Right(!cod_sunat, 4) = "0810" Then '/*bonificacion extraordinaria,SCTR - S,SCTR - P */
                        Reg = Trim(Left(!tipo_doc & "", 2))
                        Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                        Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                        Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                        Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                        Reg = Reg & vSeparador
                        Print #1, Reg '& "***"
                    Else
                        xAcumGratificacion = xAcumGratificacion + CCur(!monto_devengado)
                    End If
                    I = I + 1
                    Barra.Value = I
                                                        
                    stipdoc = Trim(!tipo_doc & "")
                    sNumdoc = Trim(Left(!nro_doc & "", 15))
                    
                .MoveNext
            Loop
            Reg = stipdoc
            Reg = Reg & vSeparador & sNumdoc
            Reg = Reg & vSeparador & "0406" 'Right(!cod_sunat, 4)
            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
            Reg = Reg & vSeparador
            Print #1, Reg ' & "***"
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.(GRATIFICACION)", True, vbBlue
'Else
'    Add_Mensaje LstError, "Estructura " & Rs!NRO & ": no existen datos para exportar.", True, vbBlack
'
End If
Barra.Value = 0


Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1



Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
If Trim(xMsgUser) <> "" Then
    Add_Mensaje LstError, Trim(xMsgUser), True, vbRed
Else
    Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
End If
Screen.MousePointer = 0
End Sub




Public Sub Exp_Estructura18xxxx()
'/*ULIMO CAMBIO ADD JCMS 05/08/09*/
'/*jcbr solicito acumular las boletas de gratificacion en el concepto de sunat 0406 excepto el codigo 0312*/

'//*** ESTRUCTURA 18: "Datos del detalle de la remuneración del trabajador"
'On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
Dim xMsgUser As String

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".rem"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "',0"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then

    Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    Dim xId As String
    Dim xCodSunat As String
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        Dim sTipoBoleta As String
        sTipoBoleta = Trim(Rq!Proceso)
        If Trim(sTipoBoleta) = "03" Or Left(Trim(sTipoBoleta), 1) = "G" Then 'boleta de gratificaciones
            sTipoBoleta = "G" 'grati
        Else
            sTipoBoleta = "N" 'Normal o cualquier tipo
        End If
        
        'For intloop = 4 To Rq.Fields.count - 1
        For intloop = 5 To Rq.Fields.count - 1
            
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    Else
                        xTipo = "03" 'deducciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
'                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Then
'                        GoTo SaltarSunat:
'                    End If
                    
                    'add jcms 090110
                    If sTipoBoleta = "G" Then ' exclusiones gratificacion
                        If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Or (xTipo = "03" And Left(Rq.Fields(intloop).Name, 1) = "d") Then
                            GoTo SaltarSunat:
                        End If
                    Else
                        If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Then
                            GoTo SaltarSunat:
                        End If
                    End If
                    
                    
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                    
                End With
            End If
            
SaltarSunat:
            
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    With rsRem
        If .RecordCount > 0 Then
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            .MoveFirst
            Do While Not .EOF
                    Reg = Trim(Left(!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                    I = I + 1
                    Barra.Value = I
                .MoveNext
            Loop
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0
GoTo fin:
'/*SOLO PARA GRATIFICACION*/

Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "',1"
If fAbrRst(Rq, Sql) Then
    'Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    'Dim xId As String
    'Dim xCodSunat As String
    
    xId = "": xCodSunat = ""
    
    Barra.Max = Rq.RecordCount
    Barra.Min = 0
    Do While Not Rq.EOF
        'Dim intloop  As Integer
        intloop = 0
        'Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            'If Rq.Fields(intloop).Name = "a20" Then Stop
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
            
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    ElseIf Left(Rq.Fields(intloop).Name, 1) = "d" Then
                        xTipo = "03" 'deducciones y aportaciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
                    'If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Or xTipo = "03" Then
                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Or (xTipo = "03" And Left(Rq.Fields(intloop).Name, 1) = "d") Then
                        GoTo SaltarSunat2:
                    End If
                    
                    
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                    
                End With
            End If
            
SaltarSunat2:
            
        Next intloop
        I = I + 1
        Barra.Value = 1
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    Dim xAcumGratificacion As Currency
    xAcumGratificacion = 0
    
    
    Dim stipdoc As String
    Dim sNumdoc As String
    stipdoc = ""
    sNumdoc = ""
    
    With rsRem
        If .RecordCount > 0 Then
            '***Print #1, "GRATI"
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            xAcumGratificacion = 0
            .MoveFirst
            Do While Not .EOF
                    If (stipdoc & sNumdoc <> Trim(!tipo_doc & "") & Trim(Left(!nro_doc & "", 15))) And Trim(stipdoc & sNumdoc) <> "" Then
                        
                            Reg = stipdoc
                            Reg = Reg & vSeparador & Trim(Left(sNumdoc, 15))
                            Reg = Reg & vSeparador & "0406" 'Right(!cod_sunat, 4)
                            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
                            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
                            Reg = Reg & vSeparador
                            Print #1, Reg ' & "***"
                            xAcumGratificacion = 0
                    End If
            
                    'If Right(!cod_sunat, 4) = "0312" Then '/*bonificacion extraordinaria*/
                    'add jcms 090110 se agrego scrt
                    If Right(!cod_sunat, 4) = "0312" Or Right(!cod_sunat, 4) = "0809" Or Right(!cod_sunat, 4) = "0810" Then '/*bonificacion extraordinaria,SCTR - S,SCTR - P */
                        Reg = Trim(Left(!tipo_doc & "", 2))
                        Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                        Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                        Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                        Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                        Reg = Reg & vSeparador
                        Print #1, Reg '& "***"
                    Else
                        xAcumGratificacion = xAcumGratificacion + CCur(!monto_devengado)
                    End If
                    I = I + 1
                    Barra.Value = I
                                                        
                    stipdoc = Trim(!tipo_doc & "")
                    sNumdoc = Trim(Left(!nro_doc & "", 15))
                    
                .MoveNext
            Loop
            Reg = stipdoc
            Reg = Reg & vSeparador & sNumdoc
            Reg = Reg & vSeparador & "0406" 'Right(!cod_sunat, 4)
            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
            Reg = Reg & vSeparador
            Print #1, Reg ' & "***"
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.(GRATIFICACION)", True, vbBlue
'Else
'    Add_Mensaje LstError, "Estructura " & Rs!NRO & ": no existen datos para exportar.", True, vbBlack
'
End If
Barra.Value = 0

fin:
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1



Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
If Trim(xMsgUser) <> "" Then
    Add_Mensaje LstError, Trim(xMsgUser), True, vbRed
Else
    Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
End If
Screen.MousePointer = 0
End Sub

Public Sub Exp_Estructura18()
'/*ULIMO CAMBIO ADD JCMS 05/08/09*/
'/*jcbr solicito acumular las boletas de gratificacion en el concepto de sunat 0406 excepto el codigo 0312*/

'//*** ESTRUCTURA 18: "Datos del detalle de la remuneración del trabajador"
'On Error GoTo MsgErr:
Dim Reg As String
Dim I As Integer
I = 0
Dim xMsgUser As String

xRutaFile = App.Path & "\reports\" & "0601" & xPeriodoAño & xPeriodoMes & xRucCia & ".rem"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "',0"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then

    Dim rsRem As New ADODB.Recordset
    If rsRem.State = 1 Then rsRem.Close
    rsRem.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRem.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRem.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRem.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRem.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRem.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRem.Open

    Dim xId As String
    Dim xCodSunat As String
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        Dim intloop  As Integer
        intloop = 0
        Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            'If Rq.Fields(intloop).Name = "a03" Then Stop
            If Rq.Fields(intloop).Value <> 0 Or Rq.Fields(intloop).Name = "d13" Then ' Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRem
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    Else
                        xTipo = "03" 'deducciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    If Rq.Fields(intloop).Name = "i45" Then
                        Dim A
                        A = A
                    End If

                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Then
                        GoTo SaltarSunat:
                    End If
                    
                    
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                End With
            End If
            
SaltarSunat:
            
        Next intloop
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    
'    Barra.Max = rsRem.RecordCount
'    Barra.Min = 1 - 1
'    i = 0
'    With rsRem
'        If .RecordCount > 0 Then
'            .Sort = "tipo_doc,nro_doc,cod_sunat"
'            .MoveFirst
'            Do While Not .EOF
'                    Reg = Trim(Left(!tipo_doc & "", 2))
'                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
'                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
'                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
'                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
'                    Reg = Reg & vSeparador
'                    Print #1, Reg
'                    i = i + 1
'                    Barra.Value = i
'                .MoveNext
'            Loop
'        End If
'    End With
'    rsRem.Close
'    Set rsRem = Nothing
'    Add_Mensaje LstError, "Estructura " & Rs!NRO & ": ( " & CStr(i) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
    
End If
Barra.Value = 0

'/*SOLO PARA GRATIFICACION*/

Sql = "usp_pla_exporta_planilla_electronica '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "',1"
If fAbrRst(Rq, Sql) Then
    Dim rsRemG As New ADODB.Recordset
    If rsRemG.State = 1 Then rsRem.Close
    rsRemG.Fields.Append "id", adChar, 21, adFldIsNullable
    rsRemG.Fields.Append "tipo_doc", adChar, 2, adFldIsNullable
    rsRemG.Fields.Append "nro_doc", adChar, 15, adFldIsNullable
    rsRemG.Fields.Append "cod_sunat", adChar, 4, adFldIsNullable
    rsRemG.Fields.Append "tipo_rem", adChar, 2, adFldIsNullable
    rsRemG.Fields.Append "monto_devengado", adCurrency, , adFldIsNullable
    rsRemG.Fields.Append "monto_pagado", adCurrency, , adFldIsNullable
    rsRemG.Open

    'Dim xId As String
    'Dim xCodSunat As String
    
    xId = "": xCodSunat = ""
    
    Barra.Max = Rq.RecordCount
    Barra.Min = 0
    Do While Not Rq.EOF
        'Dim intloop  As Integer
        intloop = 0
        'Dim xTipo As String
        For intloop = 4 To Rq.Fields.count - 1
            'If Rq.Fields(intloop).Name = "a20" Then Stop
            If Rq.Fields(intloop).Value <> 0 Or (Rq.Fields(intloop).Name = "d13" Or Rq.Fields(intloop).Name = "a03") Then  'Rq.Fields(intloop).Name = "d04" Or Rq.Fields(intloop).Name = "d11"
            
                 xId = "": xCodSunat = "": xMsgUser = ""
                 With rsRemG
                    If Left(Rq.Fields(intloop).Name, 1) = "i" Then
                        xTipo = "02" 'remuneraciones
                    ElseIf Left(Rq.Fields(intloop).Name, 1) = "d" Then
                        xTipo = "03" 'deducciones y aportaciones
                    End If
                    If Len(Rq.Fields(intloop).Name) = 3 Then
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 2)))
                    Else
                        xCodSunat = Buscar_CodSunat(RsConcep, xTipo & Trim(Right(Rq.Fields(intloop).Name, 3)))
                    End If
                    
                    '/*modi jcms add 10/12/08 - jcbd no incluir ESSALUD VIDA(0604), ESSALUD (0804), SNP (0607) */
                    '/* segun cambio de sunat*/
                    'If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Or Or (xTipo = "03" And Left(Rq.Fields(intloop).Name, 1) = "d")  Then
                    If Trim(xCodSunat) = "0604" Or Trim(xCodSunat) = "0804" Or Trim(xCodSunat) = "0607" Then
                        GoTo SaltarSunat2:
                    End If
                    
                    
                    If Trim(xCodSunat) = "" Or Len(Trim(xCodSunat)) <> 4 Then
                        xMsgUser = "Error Se canceló exportación " & rs!ESTRUCTURA & Chr(13) & " Codigo de SUNAT NO SETEADO " & Chr(13) & "Verifique concepto " & NombreConcepto_Plahistorico(wcia, UCase(Rq.Fields(intloop).Name))
                        MsgBox xMsgUser, vbCritical, Me.Caption
                        GoTo Salir:
                        'Add_Mensaje LstError, "Error Se canceló exportación " & Rs!ESTRUCTURA & "Codigo de SUNAT no asignado verifique campo " & Rq.Fields(intloop).Name, True, vbRed
                    End If
                    
                    '//*** No se exporta el valor total de la afp, sino el disgregado comision,aporte y seguro del afp
                    If xTipo & Rq.Fields(intloop).Name <> "03d11" Then
    '                    If Trim(Rq!nro_doc & "") = "00327823" Then
    '                        Stop
    '                    End If
                        xId = Trim(Rq!tipo_doc & "") & Trim(Left(Rq!nro_doc & "", 15)) & Trim(xCodSunat)
                        If Not .EOF Then .MoveFirst
                        .FIND "id='" & xId & "'", 0, 1, 1
                        If .EOF Then
                            .AddNew
                            !Id = xId
                            !tipo_doc = Rq!tipo_doc & ""
                            !nro_doc = Rq!nro_doc & ""
                            !cod_sunat = xCodSunat
                            !monto_devengado = Rq.Fields(intloop).Value
                            !monto_pagado = Rq.Fields(intloop).Value
                        Else
                            !monto_devengado = !monto_devengado + Rq.Fields(intloop).Value
                            !monto_pagado = !monto_pagado + Rq.Fields(intloop).Value
                        End If
                    End If
                    
                End With
            End If
            
SaltarSunat2:
            
        Next intloop
        I = I + 1
        Barra.Value = 1
        Rq.MoveNext
    Loop
    
    Barra.Max = rsRemG.RecordCount
    Barra.Min = 1 - 1
    I = 0
    Dim xAcumGratificacion As Currency
    xAcumGratificacion = 0
    
    
    Dim stipdoc As String
    Dim sNumdoc As String
    stipdoc = ""
    sNumdoc = ""
    
    With rsRemG
        If .RecordCount > 0 Then
            '***Print #1, "GRATI"
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            xAcumGratificacion = 0
            .MoveFirst
            Do While Not .EOF
                    If (stipdoc & sNumdoc <> Trim(!tipo_doc & "") & Trim(Left(!nro_doc & "", 15))) And Trim(stipdoc & sNumdoc) <> "" Then
                        
'                            Reg = stipdoc
'                            Reg = Reg & vSeparador & Trim(Left(sNumdoc, 15))
'                            Reg = Reg & vSeparador & "0406" 'Right(!cod_sunat, 4)
'                            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
'                            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
'                            Reg = Reg & vSeparador
'                            Print #1, Reg ' & "***"

                        xId = Trim(!tipo_doc & "") & Trim(Left(!nro_doc & "", 15)) & "0406"
                        If Not rsRem.EOF Then rsRem.MoveFirst
                        rsRem.FIND "id='" & xId & "'", 0, 1, 1
                        If rsRem.EOF Then
                            rsRem.AddNew
                            rsRem!Id = xId
                            rsRem!tipo_doc = stipdoc
                            rsRem!nro_doc = sNumdoc
                            rsRem!cod_sunat = "0406"
                            rsRem!monto_devengado = Format(xAcumGratificacion, "#######.00")
                            rsRem!monto_pagado = Format(xAcumGratificacion, "#######.00")
                        Else
                            rsRem!monto_devengado = rsRem!monto_devengado + Format(xAcumGratificacion, "#######.00")
                            rsRem!monto_pagado = rsRem!monto_pagado + Format(xAcumGratificacion, "#######.00")
                        End If
                            
                            xAcumGratificacion = 0
                    End If
            
                    'If Right(!cod_sunat, 4) = "0312" Then '/*bonificacion extraordinaria*/
                    'add jcms 090110 se agrego scrt
                    If Right(!cod_sunat, 4) = "0312" Or Right(!cod_sunat, 4) = "0809" Or Right(!cod_sunat, 4) = "0810" Or Right(!cod_sunat, 4) = "0605" Then  '/*bonificacion extraordinaria,SCTR - S,SCTR - P  5TA CAT GRATI*/
'                        Reg = Trim(Left(!tipo_doc & "", 2))
'                        Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
'                        Reg = Reg & vSeparador & Right(!cod_sunat, 4)
'                        Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
'                        Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
'                        Reg = Reg & vSeparador
'                        Print #1, Reg '& "***"
                        
                        
                        xId = Trim(Left(!tipo_doc & "", 2)) & Trim(Left(!nro_doc & "", 15)) & Right(!cod_sunat, 4)
                        If Not rsRem.EOF Then rsRem.MoveFirst
                        rsRem.FIND "id='" & xId & "'", 0, 1, 1
                        If rsRem.EOF Then
                            rsRem.AddNew
                            rsRem!Id = xId
                            rsRem!tipo_doc = Trim(Left(!tipo_doc & "", 2))
                            rsRem!nro_doc = Trim(Left(!nro_doc & "", 15))
                            rsRem!cod_sunat = Right(!cod_sunat, 4)
                            rsRem!monto_devengado = Format(!monto_devengado, "#######.00")
                            rsRem!monto_pagado = Format(!monto_pagado, "#######.00")
                        Else
                            rsRem!monto_devengado = rsRem!monto_devengado + Format(!monto_devengado, "#######.00")
                            rsRem!monto_pagado = rsRem!monto_pagado + Format(!monto_pagado, "#######.00")
                        End If
            
                    Else
                        xAcumGratificacion = xAcumGratificacion + CCur(!monto_devengado)
                    End If
                    I = I + 1
                    Barra.Value = I
                                                        
                    stipdoc = Trim(!tipo_doc & "")
                    sNumdoc = Trim(Left(!nro_doc & "", 15))
                    
                .MoveNext
            Loop
'            Reg = stipdoc
'            Reg = Reg & vSeparador & sNumdoc
'            Reg = Reg & vSeparador & "0406" 'Right(!cod_sunat, 4)
'            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
'            Reg = Reg & vSeparador & Format(xAcumGratificacion, "#######.00")
'            Reg = Reg & vSeparador
'            Print #1, Reg ' & "***"
            
            rsRem.AddNew
            rsRem!Id = xId
            rsRem!tipo_doc = stipdoc
            rsRem!nro_doc = sNumdoc
            rsRem!cod_sunat = "0406"
            rsRem!monto_devengado = Format(xAcumGratificacion, "#######.00")
            rsRem!monto_pagado = Format(xAcumGratificacion, "#######.00")
            
                        
        End If
    End With
    
  
'Else
'    Add_Mensaje LstError, "Estructura " & Rs!NRO & ": no existen datos para exportar.", True, vbBlack
'
End If


  '/*GENERAR ARCHIVO FINAL*/
   Barra.Max = rsRem.RecordCount
    Barra.Min = 1 - 1
    I = 0
    With rsRem
        If .RecordCount > 0 Then
            .Sort = "tipo_doc,nro_doc,cod_sunat"
            .MoveFirst
            Do While Not .EOF
                    Reg = Trim(Left(!tipo_doc & "", 2))
                    Reg = Reg & vSeparador & Trim(Left(!nro_doc & "", 15))
                    Reg = Reg & vSeparador & Right(!cod_sunat, 4)
                    Reg = Reg & vSeparador & Format(!monto_devengado, "#######.00")
                    Reg = Reg & vSeparador & Format(!monto_pagado, "#######.00")
                    Reg = Reg & vSeparador
                    Print #1, Reg
                    I = I + 1
                    Barra.Value = I
                .MoveNext
            Loop
        End If
    End With
    rsRem.Close
    Set rsRem = Nothing
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue

   
   
Barra.Value = 0


Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1



Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
If Trim(xMsgUser) <> "" Then
    Add_Mensaje LstError, Trim(xMsgUser), True, vbRed
Else
    Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
End If
Screen.MousePointer = 0
End Sub


Public Sub Exp_Estructura13()
'//*** ESTRUCTURA 13: "Importar Datos de derechohabientes - ALTA"

'On Error GoTo MsgErr:
Dim sFechaMaxPeriodo As String
sFechaMaxPeriodo = Format(fMaxDay(Val(xPeriodoMes), Val(xPeriodoAño)), "00") & "/" & Format(xPeriodoMes, "00") & "/" & Format(xPeriodoAño, "0000")
dFechaSOL = InputBox("Ingrese fecha (dd/mm/yyyy) de registro a través de SOL", , dFechaSOL)
Do While Not IsDate(dFechaSOL) Or CDate(dFechaSOL) < CDate(sFechaMaxPeriodo)
   If Trim(dFechaSOL) = "" Then Exit Sub
   dFechaSOL = InputBox("Ingrese fecha (dd/mm/yyyy) de registro a través de SOL", , dFechaSOL)
Loop

Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\RD_" & xRucCia & "_" & Format(dFechaSOL, "DDMMYYYY") & "_ALTA.TXT"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica_derechohabientes '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_docdh & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!numero & "", 15))
        Reg = Reg & vSeparador & Format(Trim(Rq!cod_pais_emision & ""), "000")
        Reg = Reg & vSeparador & Trim(Format(Rq!fec_nacdh & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador & Trim(Left(Rq!ap_pat & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!ap_mat & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!nombres & "", 40))
        Reg = Reg & vSeparador & Trim(Rq!sexodh & "")
        Reg = Reg & vSeparador & Format(Trim(Rq!vinculodh & ""), "00")
        Reg = Reg & vSeparador & Format(Trim(Rq!tipdoc_acredita_vinculo & ""), "00")
        Reg = Reg & vSeparador & Trim(Rq!nrodoc_acredita_vinculo & "")
        If Format(Trim(Rq!vinculodh & ""), "00") = "04" Then 'solo gestante
            Reg = Reg & vSeparador & Trim(Rq!mes_concepcion & "")
        Else
            Reg = Reg & vSeparador
        End If
        'direccion 1
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_via & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!NOM_VIA & "", 20))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_via1 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!NRO & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!Interior & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_manzana1 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_lote1 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_kilometro1 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_block1 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_etapa1 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_zona & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!NOM_ZONA & "", 20))
        Reg = Reg & vSeparador & Trim(Left(Rq!referencia & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!ubigeo & "", 6))
        'direccion 2
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_via2 & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!NOM_VIA2 & "", 20))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_via1 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_departamento2 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_interior2 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_manzana2 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_lote2 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_kilometro2 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_block2 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_etapa2 & "", 4))
        Reg = Reg & vSeparador & Trim(Left(Rq!cod_zona2 & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!NOM_ZONA2 & "", 20))
        Reg = Reg & vSeparador & Trim(Left(Rq!referencia2 & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!ubigeo2 & "", 6))
        Reg = Reg & vSeparador & Trim(Left(Rq!indicador_centro_essalud & "", 1))
        Reg = Reg & vSeparador & Trim(Left(Rq!telef_cod_ciudad & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!telefono & "", 10))
        Reg = Reg & vSeparador & Trim(Left(Rq!Email & "", 50))
        Reg = Reg & vSeparador
        
        
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub

Public Sub Exp_Estructura24()
'//*** ESTRUCTURA 24: "Importar Datos de derechohabientes - BAJA"

On Error GoTo MsgErr:
Dim sFechaMaxPeriodo As String
sFechaMaxPeriodo = Format(fMaxDay(Val(xPeriodoMes), Val(xPeriodoAño)), "00") & "/" & Format(xPeriodoMes, "00") & "/" & Format(xPeriodoAño, "0000")
dFechaSOL = InputBox("Ingrese fecha (dd/mm/yyyy) de registro a través de SOL", , dFechaSOL)
Do While Not IsDate(dFechaSOL) Or CDate(dFechaSOL) < CDate(sFechaMaxPeriodo)
   If Trim(dFechaSOL) = "" Then Exit Sub
   dFechaSOL = InputBox("Ingrese fecha (dd/mm/yyyy) de registro a través de SOL", , dFechaSOL)
Loop

Dim Reg As String
Dim I As Integer
I = 0
xRutaFile = App.Path & "\reports\RD_" & xRucCia & "_" & Format(dFechaSOL, "DDMMYYYY") & "_BAJA.TXT"
Dim Rq As ADODB.Recordset
Sql = "usp_pla_exporta_planilla_electronica_derechohabientes '" & wcia & "'," & rs!NRO & "," & xPeriodoAño & "," & xPeriodoMes & IIf(Trim(xUnTrabajador) <> "", ",'" & xUnTrabajador & "'", ",'*'") & ",'" & xFechaLimiteFinPeriodo & "'"
Open xRutaFile For Output As #1
Screen.MousePointer = 11
If fAbrRst(Rq, Sql) Then
    Barra.Max = Rq.RecordCount
    Barra.Min = 1 - 1
    Do While Not Rq.EOF
        
        Reg = Trim(Left(Rq!tipo_doc & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!nro_doc & "", 15))
        Reg = Reg & vSeparador & Trim(Left(Rq!tipo_docdh & "", 2))
        Reg = Reg & vSeparador & Trim(Left(Rq!numero & "", 15))
        Reg = Reg & vSeparador & Format(Trim(Rq!cod_pais_emision & ""), "000")
        Reg = Reg & vSeparador & Trim(Format(Rq!fec_nacdh & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador & Trim(Left(Rq!ap_pat & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!ap_mat & "", 40))
        Reg = Reg & vSeparador & Trim(Left(Rq!nombres & "", 40))
        Reg = Reg & vSeparador & Format(Trim(Rq!vinculodh & ""), "00")
        Reg = Reg & vSeparador & Trim(Format(Rq!fecha_baja & "", "dd/mm/yyyy"))
        Reg = Reg & vSeparador & Format(Trim(Rq!motivo_baja & ""), "00")
        
        
        
        
'        Reg = Reg & vSeparador & Trim(Left(Rq!ACTIVO_BAJA & "", 2))
'        'If Rq!nro_doc = "08341815" Then Stop
'        If Trim(Left(Rq!ACTIVO_BAJA & "", 2)) = "10" Then 'alta
'            Reg = Reg & vSeparador & Trim(Format(Rq!fecha_alta & "", "dd/mm/yyyy"))
'        Else
'            Reg = Reg & vSeparador
'        End If
'        If Trim(Left(Rq!ACTIVO_BAJA & "", 2)) = "11" Then 'baja
'            Reg = Reg & vSeparador & Trim(Left(Rq!motivo_baja & "", 1))
'            Reg = Reg & vSeparador & Trim(Format(Rq!fecha_baja & "", "dd/mm/yyyy"))
'        Else
'            Reg = Reg & vSeparador
'            Reg = Reg & vSeparador
'        End If
'        If Trim(Rq!vinculodh & "") = "1" And Trim(Rq!nrocertificado) <> "" Then 'hijo
'            Reg = Reg & vSeparador & Trim(Left(Rq!nrocertificado & "", 20))
'        Else
'            Reg = Reg & vSeparador
'        End If
'        Reg = Reg & vSeparador & Trim(Left(Rq!domicilio & "", 1))
        
        Reg = Reg & vSeparador
        
        
        
        Print #1, Reg
        I = I + 1
        Barra.Value = I
        Rq.MoveNext
    Loop
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": ( " & CStr(I) & " ) registros exportados correctamente.", True, vbBlue
Else
    Add_Mensaje LstError, "Estructura " & rs!NRO & ": no existen datos para exportar.", True, vbBlack
End If
Barra.Value = 0
Rq.Close
Set Rq = Nothing
Screen.MousePointer = 0
Close #1
Exit Sub
MsgErr:
Close #1
MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
Salir:
On Error Resume Next:
Close #1
If Dir(xRutaFile) <> "" Then DeleteAFile xRutaFile
Add_Mensaje LstError, "Se canceló exportación " & rs!ESTRUCTURA, True, vbRed
Screen.MousePointer = 0
End Sub


