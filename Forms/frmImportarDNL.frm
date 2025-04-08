VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmImportarDNL 
   Caption         =   "Importar detalle de días no laborados"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8205
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
      TabIndex        =   9
      Top             =   7320
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
      Height          =   5850
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   7605
      Begin TrueOleDBGrid70.TDBGrid DGrd 
         Height          =   5385
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   9499
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
      TabIndex        =   0
      Top             =   0
      Width           =   7605
      Begin VB.TextBox TxtRango 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Text            =   "A1:I1"
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox Txtarchivos 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   705
         Width           =   5370
      End
      Begin Threed.SSCommand cmdVerArchivo 
         Height          =   495
         Left            =   6840
         TabIndex        =   3
         Top             =   600
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   873
         _StockProps     =   78
         Picture         =   "frmImportarDNL.frx":0000
      End
      Begin MSMask.MaskEdBox TxtPeriodo 
         Height          =   315
         Left            =   6480
         TabIndex        =   11
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   7
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo:"
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
         Index           =   0
         Left            =   5520
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Ejm: A1:G45"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   360
         Width           =   975
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
         TabIndex        =   5
         Top             =   360
         Width           =   2745
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
         TabIndex        =   4
         Top             =   720
         Width           =   945
      End
   End
   Begin MSComctlLib.ProgressBar BarraImporta 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog Box 
      Left            =   1800
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImportarDNL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsExport As New ADODB.Recordset
Dim rsRegUnicoTrabajador As New ADODB.Recordset

Dim rs As New ADODB.Recordset

Private Sub cmdImportar_Click()
    Exportar
End Sub

Private Sub Exportar()
With rsExport
If .RecordCount > 0 Then
    .MoveFirst
    FrmDiasNoTrab.TxtCodTrab.Text = IIf(Trim(.Fields("codigo").Value + "") <> "", Trim(.Fields("codigo").Value + ""), "")
    FrmDiasNoTrab.LblNomTrab.Caption = IIf(Trim(.Fields("trabajador").Value + "") <> "", Trim(.Fields("trabajador").Value + ""), "")
    
    Do While Not .EOF
        rs.AddNew
        rs!FecIni = Trim(.Fields("fecha").Value)
        rs!FecFin = rs!FecIni
        rs!nrocitt = ""
        rs!cod_suspension = IIf(Trim(.Fields("sunat").Value + "") <> "", Trim(.Fields("sunat").Value + ""), "")
        rs!PlaCod = Trim(.Fields("codigo").Value)
        rs!cod_cia = Trim(.Fields("codcia").Value)
        .MoveNext
    Loop
    FrmDiasNoTrab.TxtPeriodo.Text = TxtPeriodo.Text
    Call FrmDiasNoTrab.BorrarRegistro(rsRegUnicoTrabajador)
    Call FrmDiasNoTrab.GrabarRegistro(rs)
    Unload Me
End If

End With

End Sub


Private Sub cmdVerArchivo_Click()
'IMPLEMENTACION GALLOS

'    LstObs.Clear
    Set rsExport = Nothing
    LimpiarRsT rsExport, DGrd

    Txtarchivos.Text = AbrirFile("*.xls", Box)
    Txtarchivos.ToolTipText = Txtarchivos.Text
  
    If Trim(Txtarchivos.Text) <> "" Then Importar_Excel

End Sub
Private Sub Importar_Excel()

    Dim strMes As String
    
    strMes = Left(TxtPeriodo.Text, 2)
    
    'IMPLEMENTACION GALLOS

    'Referencia a la instancia de excel
    
    Dim xlApp2 As Excel.Application
    Dim xlApp1  As Excel.Application
    Dim xLibro  As Excel.Workbook
        
    On Error Resume Next
        
    'Chequeamos si excel esta corriendo
        
    Set xlApp1 = GetObject(, "Excel.Application")
    If xlApp1 Is Nothing Then
        'Si excel no esta corriendo, creamos una nueva instancia.
        Set xlApp1 = CreateObject("Excel.Application")
    End If
    
    'Variable de tipo Aplicación de Excel
    
    Set xlApp2 = xlApp1.Application
    
    Dim Col As Integer, Fila As Integer
  
  
    'Usamos el método open para abrir el archivo que está _
     en el directorio del programa llamado archivo.xls
 
    Set xLibro = xlApp2.Workbooks.Open(Txtarchivos.Text)
  
    'Hacemos el Excel Invisible
    
    xlApp2.Visible = False
    
'    Dim CadSem As String
'    CadSem = xLibro.Sheets(1).Cells(2, 1).Value
'    Dim pos As Integer
'    pos = InStr(1, CadSem, VSemana)
'    If pos = 0 Then
'        MsgBox "El contenido del archivo no pertences a la Semana " & VSemana, vbCritical, Me.Caption
'        GoTo Salir:
'    End If
    On Error GoTo ERR
    
    Dim conexion As ADODB.Connection, rs As ADODB.Recordset
  
    Set conexion = New ADODB.Connection
       
    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & Txtarchivos.Text & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
    
    ' Nuevo recordset
    
    Set rsRegUnicoTrabajador = New ADODB.Recordset
    With rsRegUnicoTrabajador
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
    
    
    Set rsExport = New ADODB.Recordset
    With rsExport
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
        
rsExport.Open "SELECT * FROM [DNL" & strMes & "$" & Trim(TxtRango.Text) & "]" & " ORDER BY codigo", conexion, , , adCmdText
rsRegUnicoTrabajador.Open "SELECT DISTINCT codigo As placod FROM [DNL" & strMes & "$" & Trim(TxtRango.Text) & "]", conexion, , , adCmdText

'rsExport.Filter = "placod LIKE '" & mPrefijo & "%'"
'
'' Mostramos los datos en el datagrid
'
'If rsExport.RecordCount <= 0 Then
'    MsgBox "Codigo de trabajadores no corresponden a la compañia", vbCritical, Me.Caption
'    GoTo Salir:
'End If
    

Set DGrd.DataSource = rsExport
    
Salir:

xLibro.Close
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xLibro Is Nothing Then Set xLibro = Nothing

Exit Sub

ERR:
    MsgBox ERR.Number & "-" & ERR.Description, vbCritical, Me.Caption
    Exit Sub

End Sub
Public Sub Crea_Rs()
    If rs.State = 1 Then rs.Close
    rs.Fields.Append "fecini", adChar, 10, adFldIsNullable
    rs.Fields.Append "fecfin", adChar, 10, adFldIsNullable
    rs.Fields.Append "nrocitt", adVarChar, 16, adFldIsNullable
    rs.Fields.Append "cod_suspension", adChar, 2, adFldIsNullable
    rs.Fields.Append "placod", adChar, 8, adFldIsNullable
    rs.Fields.Append "cod_cia", adChar, 2, adFldIsNullable
    rs.Open
End Sub

Private Sub Form_Load()
    Crea_Rs
    TxtPeriodo.Text = FrmDiasNoTrab.TxtPeriodo.Text
End Sub
