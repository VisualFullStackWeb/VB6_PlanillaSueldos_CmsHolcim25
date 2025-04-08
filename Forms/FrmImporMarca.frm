VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmImporMarca 
   Caption         =   "Importación de Marcaciones"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18165
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   18165
   Begin VB.Frame frmImportar 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18135
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   17895
         Begin VB.ComboBox Cmbcia 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   90
            Width           =   16650
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
            TabIndex        =   22
            Top             =   120
            Width           =   825
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   795
         Width           =   6975
         Begin VB.TextBox Txtsemana 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   14
            Top             =   705
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox Cmbtipotrabajador 
            Height          =   315
            ItemData        =   "FrmImporMarca.frx":0000
            Left            =   1440
            List            =   "FrmImporMarca.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   200
            Width           =   5415
         End
         Begin MSComCtl2.DTPicker Cmbfecha 
            Height          =   285
            Left            =   750
            TabIndex        =   10
            Top             =   705
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   57999361
            CurrentDate     =   37265
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   3270
            TabIndex        =   11
            Top             =   705
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker Cmbal 
            Height          =   285
            Left            =   5640
            TabIndex        =   12
            Top             =   705
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   57999361
            CurrentDate     =   37265
         End
         Begin MSComCtl2.DTPicker Cmbdel 
            Height          =   285
            Left            =   4080
            TabIndex        =   13
            Top             =   705
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   57999361
            CurrentDate     =   37267
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   705
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Semana"
            Height          =   195
            Left            =   2160
            TabIndex        =   18
            Top             =   705
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Del"
            Height          =   195
            Left            =   3720
            TabIndex        =   17
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Al"
            Height          =   195
            Left            =   5400
            TabIndex        =   16
            Top             =   705
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Trabajador"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1125
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
         Height          =   1095
         Left            =   7200
         TabIndex        =   3
         Top             =   795
         Width           =   10845
         Begin VB.TextBox Txtarchivos 
            BackColor       =   &H8000000A&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   345
            Width           =   8730
         End
         Begin Threed.SSCommand cmdVerArchivo 
            Height          =   495
            Left            =   10080
            TabIndex        =   5
            Top             =   240
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   873
            _StockProps     =   78
            Picture         =   "FrmImporMarca.frx":0004
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
            TabIndex        =   6
            Top             =   360
            Width           =   945
         End
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
         Height          =   6090
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   17970
         Begin TrueOleDBGrid70.TDBGrid DGrd 
            Height          =   5625
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   17745
            _ExtentX        =   31300
            _ExtentY        =   9922
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
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
            _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
         Left            =   15600
         TabIndex        =   1
         Top             =   8220
         Width           =   2355
      End
      Begin MSComctlLib.ProgressBar BarraImporta 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   8280
         Width           =   15285
         _ExtentX        =   26961
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog Box 
         Left            =   360
         Top             =   8280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "FrmImporMarca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsExport As New ADODB.Recordset
Dim VTipotrab As String


Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
End Sub


Private Sub Cmbtipotrabajador_Click()
VTipotrab = fc_CodigoComboBox(Cmbtipotrabajador, 2)
Dim wciamae As String
Dim wBeginMonth As String

VHorasBol = 0
VTipoPago = ""
wciamae = Determina_Maestro("01055")
Sql$ = "Select * from maestros_2 where cod_maestro2='" & VTipotrab & "' and status<>'*'"
Sql$ = Sql$ & wciamae
cn.CursorLocation = adUseClient
Set Rs = New ADODB.Recordset
Set Rs = cn.Execute(Sql$, 64)
If Rs.RecordCount > 0 Then
   Rs.MoveFirst
   VHorasBol = Val(Rs!flag2)
   VTipoPago = Left(Rs!flag1, 2)
End If

Select Case Left(Rs!flag1, 2)
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
            If (fAbrRst(Rs, Sql$)) Then
               If IsNull(Rs!iniciomes) Then wBeginMonth = "1" Else wBeginMonth = Rs!iniciomes
            End If
            Rs.Close
            
            If Trim(wBeginMonth) = "" Then
                MsgBox "Ingrese el Inicio Del Mes", vbInformation, ""
            Exit Sub
            End If
            
            If Trim(wBeginMonth) <> "1" Then
               Cmbfecha.Day = Val(wBeginMonth) - 1
            Else
               Cmbfecha.Day = Val(fMaxDay(Cmbfecha.Month, Cmbfecha.Year))
            End If
            Cmbfecha.Enabled = True
       Case Else
            Txtsemana.Visible = True
            UpDown1.Visible = True
            Label4.Visible = True
            Label5.Visible = True
            Label6.Visible = True
            Cmbdel.Visible = True
            Cmbal.Visible = True
            Cmbfecha.Enabled = False
End Select

If Rs.State = 1 Then Rs.Close

End Sub

Private Sub cmdImportar_Click()
If rsExport.State <> 1 Then Exit Sub
If rsExport.RecordCount > 0 Then rsExport.MoveFirst

BarraImporta.Max = rsExport.RecordCount
BarraImporta.Min = 0
Dim I As Integer
I = 1
Do While Not rsExport.EOF
Sql$ = "usp_Pla_Inserta_Impor_Marcaciones "
   Sql$ = Sql$ & "'" & wcia & "',"
   Sql$ = Sql$ & "'" & Trim(rsExport!Codigo & "") & "',"
   Sql$ = Sql$ & "'" & Trim(VTipotrab & "") & "',"
   Sql$ = Sql$ & "'" & rsExport!PLADIAST & "',"
   Sql$ = Sql$ & "'" & rsExport!PLAHORAS & "',"
   Sql$ = Sql$ & "'" & rsExport!PLADOMIN & "',"
   Sql$ = Sql$ & "'" & rsExport!PLAFERIA & "',"
   Sql$ = Sql$ & "'" & rsExport!HRSE_2PR & "',"
   Sql$ = Sql$ & "'" & rsExport!HRSE_3RA & "',"
   Sql$ = Sql$ & "'" & rsExport!HRSE_DFE & "',"
   Sql$ = Sql$ & "'" & rsExport!ENF_PAGA & "',"
   Sql$ = Sql$ & "'" & rsExport!ENF_NOPA & "',"
   Sql$ = Sql$ & "'" & rsExport!STASA2 & "',"
   Sql$ = Sql$ & "'" & rsExport!STASA3 & "',"
   Sql$ = Sql$ & "'" & rsExport!FALINJ & "',"
   Sql$ = Sql$ & "'" & Cmbfecha.Value & "',"
   Sql$ = Sql$ & "'" & Trim(Txtsemana.Text & "") & "',"
   Sql$ = Sql$ & "'" & Trim(wuser & "") & "',"
   If Trim(rsExport!REINTEGRo & "") = "" Then
      Sql$ = Sql$ & "0,"
   Else
      Sql$ = Sql$ & "'" & rsExport!REINTEGRo & "',"
   End If
   
   If Trim(rsExport!Canasta & "") = "" Then
      Sql$ = Sql$ & "0,"
   Else
      Sql$ = Sql$ & "'" & rsExport!Canasta & "',"
   End If
   
   If Trim(rsExport!OtrosDesc & "") = "" Then
      Sql$ = Sql$ & "0,"
   Else
      Sql$ = Sql$ & "'" & rsExport!OtrosDesc & "',"
   End If
   
   If Trim(rsExport!AsigEsc & "") = "" Then
      Sql$ = Sql$ & "0,"
   Else
      Sql$ = Sql$ & "'" & rsExport!AsigEsc & "',"
   End If
   
   If Trim(rsExport!Utilidades & "") = "" Then
      Sql$ = Sql$ & "0,"
   Else
      Sql$ = Sql$ & "'" & rsExport!Utilidades & "'"
   End If
   
   
   cn.Execute Sql$
   
   BarraImporta.Value = I
   I = I + 1
   rsExport.MoveNext
Loop

End Sub

Private Sub cmdVerArchivo_Click()
Set rsExport = Nothing
LimpiarRsT rsExport, DGrd
            
AbrirFile ("*.xls")
If Trim(Txtarchivos.Text) <> "" Then Importar_Excel
End Sub

Public Sub AbrirFile(pextension As String)
If Not Cuadro_Dialogo_Abrir(pextension) Then
    Txtarchivos.Text = ""
    Exit Sub
End If
Debug.Print Box.DefaultExt

If UCase(Right(Box.FileName, 3)) <> UCase(Right(pextension, 3)) Then
   MsgBox "La Extensión de archivo no concuerda con el formato elegido", vbCritical, "Archivo Inválido"
   Exit Sub
End If
Txtarchivos.Text = Box.FileName
Txtarchivos.ToolTipText = Box.FileName
End Sub

Public Sub Importar_Excel()
    'Referencia a la instancia de excel
    Dim xlApp2 As Excel.Application
    Dim xlApp1 As Excel.Application
    Dim xLibro As Excel.Workbook
        
    On Error Resume Next
        
    'Chequeamos si excel esta corriendo
        
    Set xlApp1 = GetObject(, "Excel.Application")
    If xlApp1 Is Nothing Then
        'Si excel no esta corriendo, creamos una nueva instancia.
        Set xlApp1 = CreateObject("Excel.Application")
    End If
        
    'ACP On Error GoTo 0
    
    'On Error GoTo ERR
        
    'Variable de tipo Aplicación de Excel
    
    Set xlApp2 = xlApp1.Application
    
    'Una variable de tipo Libro de Excel
    
    Dim Col As Integer, Fila As Integer
  
  
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
    
    Set xLibro = xlApp2.Workbooks.Open(Txtarchivos.Text)
  
    'Hacemos el Excel Invisible
    
    xlApp2.Visible = False
    
    Dim xTipoTrab As String
    xTipoTrab = vTipoTra
    
    'Eliminamos los objetos si ya no los usamos
    
    If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
    If Not xLibro Is Nothing Then Set xlBook = Nothing
    
    Dim conexion As ADODB.Connection, Rs As ADODB.Recordset
  
    Set conexion = New ADODB.Connection
       
    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Txtarchivos.Text & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""

    ' Nuevo recordset
    Set rsExport = New ADODB.Recordset
       
    With rsExport
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
         
         
Set rsPlaSeteo = Nothing
Sql$ = "select camposql, campoexcel from plaSeteoCampos_marcaciones where status !='*' order by tipo "
cn.CursorLocation = adUseClient
Set rsPlaSeteo = New ADODB.Recordset
Set rsPlaSeteo = cn.Execute(Sql$, 64)
         
If VTipotrab = "02" Then
   'CARGA SEMANAL
   If rsPlaSeteo.RecordCount > 0 Then
   Else
      MsgBox "Debe setear la columnas del archivo cargado con las columnas de la Boleta", vbCritical, Me.Caption
      GoTo Salir:
   End If
   rsExport.Open "SELECT * FROM [MSEMANA" & Trim(Txtsemana.Text) & "$] Where Codigo<>'' Order by Codigo", conexion, , , adCmdText
Else
   rsExport.Open "SELECT * FROM [2da quincena$] Where Codigo<>'' Order by Codigo", conexion, , , adCmdText
End If

If mPrefijo <> "" Then
    rsExport.Filter = "placod = '" & mPrefijo & "'"
End If
' Mostramos los datos en el datagrid
    
If rsExport.RecordCount <= 0 Then
    MsgBox "Codigo de trabajadores no corresponden a la compañia", vbCritical, Me.Caption
    GoTo Salir:
End If
    
Set DGrd.DataSource = rsExport
Dim I As Integer
For I = 0 To DGrd.Columns.count - 1
   DGrd.Columns(I).Visible = False
Next

If rsPlaSeteo.RecordCount > 0 Then rsPlaSeteo.MoveFirst
Do While Not rsPlaSeteo.EOF
   For I = 0 To DGrd.Columns.count - 1
      If UCase(Trim(DGrd.Columns(I).Caption & "")) = UCase(Trim(rsPlaSeteo!campoexcel & "")) Then DGrd.Columns(I).Visible = True: Exit For
   Next
   rsPlaSeteo.MoveNext
Loop
    
fc_SumaTotalesImportacion rsExport, DGrd
    
Salir:

xLibro.Close
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xLibro Is Nothing Then Set xlBook = Nothing
If rsPlaSeteo.State = 1 Then rsPlaSeteo.Close
Set rsPlaSeteo = Nothing
Exit Sub

Err:
    If rsPlaSeteo.State = 1 Then rsPlaSeteo.Close
    Set rsPlaSeteo = Nothing

    MsgBox Err.Number & "-" & Err.Description, vbCritical, Me.Caption
    Exit Sub
End Sub

Public Function Cuadro_Dialogo_Abrir(pextension As String) As Boolean
'IMPLEMENTACION GALLOS

 'On Error GoTo ErrHandler
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
        Box.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*"
        '"All Files (*.*)|*.*|Microsoft Excel 97/2000 (*.xls)|*.txt)"
   End Select
   ' Especifique el filtro predeterminado.
   Box.FilterIndex = 2
   'Box.FileName = "buenos.csv"
   Box.FileName = ""
   ' Presenta el cuadro de diálogo Abrir.
   
   Box.ShowOpen
   ' Llamada al procedimiento para abrir archivo.
   Dim pos As String
   'CTA.CTE MN:
   'vNroBco
   
   
   
   Dim swExiste As Variant
   swExiste = InStr(1, UCase(Trim(Box.FileName)), UCase(xFile), vbTextCompare)
   If swExiste = 0 Then
      MsgBox "Archivo Elegido no es el correcto" & Chr(13) & "El Correcto es " & xFile, vbCritical, "Importacion"
      'salir = True
      Txtarchivos = ""
    Else
      Cuadro_Dialogo_Abrir = True
    End If
   Exit Function

ErrHandler:
   Cuadro_Dialogo_Abrir = False
   'El usuario hizo clic en el botón Cancelar.
   Exit Function
End Function

Private Sub fc_SumaTotalesImportacion(ByRef pControl As ADODB.Recordset, ByRef Tdbgrid As TrueOleDBGrid70.Tdbgrid)

On Error GoTo ErrMsg:
Dim Rc As ADODB.Recordset
Set Rc = pControl.Clone

If mPrefijo <> "" Then
    Rc.Filter = "placod = '" & mPrefijo & "'"
End If
Dim Rt As New ADODB.Recordset
If Rt.State = 1 Then Rt.Close
Dim intloop  As Integer
intloop = 0
With Rc
        
        For intloop = 0 To .Fields.count - 1
            Debug.Print "campo " & .Fields(intloop).Name & "  tipo " & .Fields(intloop).Type
            Select Case .Fields(intloop).Type
               Case adCurrency, adNumeric, adDouble, adDecimal: xValue = 0#
                    Debug.Print "totalespor  " & .Fields(intloop).Name
                    Rt.Fields.Append .Fields(intloop).Name, adCurrency, , adFldIsNullable
            End Select
        Next
End With
Rt.Open
Rt.AddNew

intloop = 0
With Rc
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                intloop = 0
                For intloop = 0 To .Fields.count - 1
                   Select Case .Fields(intloop).Type
                    Case adCurrency, adNumeric, 5
                        Rt.Fields(.Fields(intloop).Name) = IIf(IsNull(Rt.Fields(.Fields(intloop).Name)), 0, Rt.Fields(.Fields(intloop).Name)) + IIf(IsNull(.Fields(intloop).Value), 0, .Fields(intloop).Value)
                    End Select
                Next
            .Update
            .MoveNext
        Loop
        .MoveFirst
    End If
End With


With Rt
    If .RecordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
                For intloop = 0 To .Fields.count - 1
                   Select Case .Fields(intloop).Type
                    Case adCurrency, adNumeric, 5
                        Tdbgrid.Columns(.Fields(intloop).Name).FooterText = Format(IIf(IsNull(.Fields(intloop).Value), 0, .Fields(intloop).Value), "###,##0.00")
                    End Select
                Next
            .MoveNext
        Loop
        .MoveFirst
    End If
End With

Rc.Close
Set Rc = Nothing
Rt.Close
Set Rt = Nothing
Exit Sub
ErrMsg:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub DGrd_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 18285: Me.Height = 9390

Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")

Cmbfecha.Year = Year(Date)
Cmbfecha.Month = Month(Date)
Cmbfecha.Day = Day(Date)


Cmbdel.Year = Year(Date)
Cmbdel.Month = Month(Date)
Cmbdel.Day = Day(Date)

Cmbal.Year = Year(Date)
Cmbal.Month = Month(Date)
Cmbal.Day = Day(Date)

End Sub

Private Sub Txtsemana_Change()

Set DGrd.DataSource = Nothing
Set rsExport = Nothing
Txtarchivos.Text = ""
Procesa_Semana
End Sub
Public Sub Procesa_Semana()
Dim mano As Integer
Dim mmes As Integer
On Error GoTo CORRIGE

If Trim(Txtsemana.Text) <> "" Then
   Sql$ = "select * from plasemanas where cia='" & wcia & "' and ano='" & Format(Cmbfecha.Year, "0000") & "' and semana='" & Format(Trim(Txtsemana.Text), "00") & "'  and status<>'*'"
   cn.CursorLocation = adUseClient
   Set Rs = New ADODB.Recordset
   
   Set Rs = cn.Execute(Sql$, 64)
   
   If Rs.RecordCount > 0 Then
      Cmbdel.Value = Format(Rs!fechai, "dd/mm/yyyy")
      Cmbal.Value = Format(Rs!fechaf, "dd/mm/yyyy")
      Cmbfecha.Value = Format(Rs!fechaf, "dd/mm/yyyy")
   End If
   
   If Rs.State = 1 Then Rs.Close
End If
CORRIGE:
End Sub

Private Sub UpDown1_DownClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "00"
If Txtsemana.Text > 0 Then Txtsemana.Text = Format(Val(Txtsemana.Text - 1), "00")
End Sub

Private Sub UpDown1_UpClick()
If Txtsemana.Text = "" Then Txtsemana.Text = "00"
Txtsemana.Text = Format(Val(Txtsemana.Text + 1), "00")


End Sub
