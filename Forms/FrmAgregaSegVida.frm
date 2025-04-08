VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAgregaSegVida 
   Caption         =   "Agrega Trabajadores al Seguro de Vida"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "FrmAgregaSegVida.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   8790
   Begin VB.Frame FrameKardex 
      Height          =   7815
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   8775
      Begin VB.TextBox Txtano 
         Enabled         =   0   'False
         Height          =   315
         Left            =   855
         MaxLength       =   4
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox CmbMes 
         Height          =   315
         ItemData        =   "FrmAgregaSegVida.frx":030A
         Left            =   1800
         List            =   "FrmAgregaSegVida.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
      Begin TrueOleDBGrid70.TDBGrid DgrdMovi 
         Height          =   6945
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   12250
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1508"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1429"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=12541"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=12462"
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
         _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
         _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2015
         Max             =   2030
         Min             =   2015
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Peiodo"
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
         TabIndex        =   7
         Top             =   285
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   60
         Width           =   6930
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
         TabIndex        =   2
         Top             =   120
         Width           =   825
      End
   End
End
Attribute VB_Name = "FrmAgregaSegVida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbMes_Click()
procesa_vida
End Sub
Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
Me.Width = 8910: Me.Height = 8760

Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Txtano.Text = Year(Date)
CmbMes.ListIndex = Month(Date) - 1
End Sub

Private Sub UpDown1_DownClick()
If Txtano.Text = "" Then Txtano.Text = "2015"
If Txtano.Text > 0 Then Txtano.Text = Format(Val(Txtano.Text - 1), "00")
procesa_vida
End Sub

Private Sub UpDown1_UpClick()
If Txtano.Text = "" Then Txtano.Text = "2015"
Txtano.Text = Format(Val(Txtano.Text + 1), "00")
procesa_vida
End Sub
Private Sub procesa_vida()
Dim Sql As String
Dim rsSaldos As New ADODB.Recordset

Sql = "select p.placod as Codigo,RTRIM(p.ap_pat)+ ' ' + RTRIM(p.ap_mat) + RTRIM(nom_1) + ' ' + RTRIM(nom_2) as Nombre " _
    & "from Pla_Agrega_Vida v, planillas p " _
    & "where v.cia=p.cia and v.placod=p.placod and v.status<>'*' and p.status<>'*' and v.ayo=" & Txtano.Text & " and v.mes=" & CmbMes.ListIndex + 1 & " " _
    & "order by p.placod"

cn.CursorLocation = adUseClient
Set rsSaldos = New ADODB.Recordset
Set rsSaldos = cn.Execute(Sql$, 64)

Set DgrdMovi.DataSource = rsSaldos


End Sub
Public Sub Agrega_Persona(Codigo)

Sql = "insert into Pla_Agrega_Vida values('" & wcia & "'," & Txtano.Text & "," & CmbMes.ListIndex + 1 & ",'" & Codigo & "','',GETDATE(),'" & wuser & "',GETDATE(),'')"
cn.Execute Sql, 64
procesa_vida

End Sub

Public Sub Elimina_Persona()
On Error GoTo CORRIGE
If Trim(DgrdMovi.Columns(0) & "") <> "" Then
   Sql = "update Pla_Agrega_Vida set status='*',user_modi='" & wuser & "',fec_modi=getdate() where  cia='" & wcia & "' and ayo=" & Txtano.Text & " and mes=" & CmbMes.ListIndex + 1 & " and placod= '" & Trim(DgrdMovi.Columns(0)) & "' and status<>'*'"
   cn.Execute Sql, 64
   procesa_vida
   
End If
CORRIGE:


End Sub
