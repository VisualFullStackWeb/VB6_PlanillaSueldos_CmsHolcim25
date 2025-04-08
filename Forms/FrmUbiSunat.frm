VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmUbiSunat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ubigeo Sunat"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "FrmUbiSunat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Limpiar"
      Height          =   255
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc AdoCon 
      Height          =   330
      Left            =   240
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
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
      Caption         =   "Adodc1"
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
   Begin TrueOleDBGrid70.TDBGrid GrdCon 
      Bindings        =   "FrmUbiSunat.frx":030A
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Departamento"
      Columns(0).DataField=   "departamento"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Provincia"
      Columns(1).DataField=   "provincia"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Ubigeo"
      Columns(2).DataField=   "nom_ubigeo"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Cod_Ubigeo"
      Columns(3).DataField=   "id_ubigeo"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4339"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5265"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5159"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=6350"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=6244"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2064"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1958"
      Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8196"
      Splits(0)._ColumnProps(21)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      Caption         =   "Ubigeos Sunat"
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Arial"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H800000&"
      _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000014&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=4,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HEFEFEF&,.appearance=0"
      _StyleDefs(14)  =   ":id=2,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=Arial"
      _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=975,.italic=0"
      _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=Arial"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=38"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000001&,.fgcolor=&H80000005&"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=33"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFFEBD7&"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=33,.bgcolor=&HEFEFEF&"
      _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42,.bgcolor=&HDDFFFF&,.bold=-1"
      _StyleDefs(27)  =   ":id=12,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(28)  =   ":id=12,.fontname=MS Sans Serif"
      _StyleDefs(29)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(30)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(31)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=33"
      _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=35"
      _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(54)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(57)  =   "Named:id=33:Normal"
      _StyleDefs(58)  =   ":id=33,.parent=0"
      _StyleDefs(59)  =   "Named:id=34:Heading"
      _StyleDefs(60)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   ":id=34,.wraptext=-1"
      _StyleDefs(62)  =   "Named:id=35:Footing"
      _StyleDefs(63)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   "Named:id=36:Selected"
      _StyleDefs(65)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(66)  =   "Named:id=37:Caption"
      _StyleDefs(67)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(68)  =   "Named:id=38:HighlightRow"
      _StyleDefs(69)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(70)  =   "Named:id=39:EvenRow"
      _StyleDefs(71)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(72)  =   "Named:id=40:OddRow"
      _StyleDefs(73)  =   ":id=40,.parent=33"
      _StyleDefs(74)  =   "Named:id=41:RecordSelector"
      _StyleDefs(75)  =   ":id=41,.parent=34"
      _StyleDefs(76)  =   "Named:id=42:FilterBar"
      _StyleDefs(77)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FrmUbiSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOpcion As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim RsConMov As New ADODB.Recordset

Private Sub Command1_Click()

End Sub

Private Sub CmdClear_Click()
    For Each Col In GrdCon.Columns
        Col.FilterText = ""
    Next Col
    AdoCon.Recordset.Filter = adFilterNone
    GrdCon.Refresh
End Sub

Private Sub Form_Activate()
SendKeys "{UP}"
SendKeys "{RIGHT}"
SendKeys "{RIGHT}"


    
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 9810
Me.Height = 4785
CargaConsulta

End Sub
Public Sub CargaConsulta()

On Error GoTo MsgError:
Dim Sql As String
Sql = "usp_pla_consulta_ubigeos '" & wcia & "'"

Screen.MousePointer = 11
If fAbrRst(RsConMov, Sql) Then
    Set AdoCon.Recordset = RsConMov
Else
    Set AdoCon.Recordset = cn.Execute("SELECT 'BLANCO'")
End If
Me.GrdCon.Refresh
Screen.MousePointer = 0
'Call fc_ActivarToolbar(StateTool(), 1, , 0, 0, 0, 0, 0, 1)
Exit Sub
MsgError:
    Screen.MousePointer = 0
    MsgBox ERR.Number & " - " & ERR.Description, vbCritical, Me.Caption
   ' ErrorLog "Se Canceló la Consulta", False, Sql, Err & Space(1) & Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
RsConMov.Close
Set RsConMov = Nothing
End Sub

Private Sub GrdCon_DblClick()
If AdoCon.Recordset.RecordCount = 1 Then
    AdoCon.Recordset.MoveFirst
Else
    If AdoCon.Recordset.AbsolutePosition = adPosUnknown Then
        MsgBox "Elija un ubigeo", vbExclamation, Me.Caption
        Exit Sub
    End If
    
End If
Dim Lugar As String
Dim Ubi As String
Lugar = Trim(AdoCon.Recordset!departamento) & " - " & Trim(AdoCon.Recordset!provincia) & " - " & Trim(AdoCon.Recordset!nom_ubigeo)
Ubi = Trim(AdoCon.Recordset!id_ubigeo)

  Select Case MDIplared.ActiveForm.Name
    
        Case "aFrmDH"
            If mOpcion = 1 Then
                aFrmDH.TxtNomUbicacion1.Text = Lugar
                aFrmDH.TxtNomUbicacion1.Tag = Ubi
            Else
                aFrmDH.TxtNomUbicacion2.Text = Lugar
                aFrmDH.TxtNomUbicacion2.Tag = Ubi
            End If
        Case "Frmpersona"
            If mOpcion = 0 Then
                Frmpersona.Text13.Text = Lugar
                Frmpersona.Text13.Tag = Ubi
            ElseIf mOpcion = 1 Then
                'Frmpersona.Text1.Text = Lugar
                'Frmpersona.Text1.Tag = Ubi
            End If
        Case "Frmcia"
              Frmcia.lbllugar = Lugar
              Frmcia.Ciacodubi = Ubi
              Frmcia.lbllugar.Tag = Ubi
        Case "Frmobras"
              Frmobras.LblUbica = Lugar
              Frmobras.Lblubigeo = Ubi
  End Select
  Unload Me
End Sub

Private Sub GrdCon_FilterChange()
On Error GoTo ErrHandler

Set Cols = GrdCon.Columns

Dim c As Integer

c = GrdCon.Col

GrdCon.HoldFields

AdoCon.Recordset.Filter = getFilter()

GrdCon.Col = c

GrdCon.EditActive = True

Exit Sub

 

ErrHandler:

    MsgBox ERR.Source & ":" & vbCrLf & ERR.Description

    Call CmdClear_Click

End Sub


Private Function getFilter() As String

    'Creates the SQL statement in adodc1.recordset.filter

    'and only filters text currently. It must be modified to
    'filter other data types.

    

    Dim tmp As String

    Dim n As Integer

    For Each Col In Cols

        If Trim(Col.FilterText) <> "" Then

            n = n + 1

            If n > 1 Then

                tmp = tmp & " AND "

            End If

            tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"

        End If

    Next Col

                

    getFilter = tmp

End Function

Private Sub GrdCon_KeyDown(KeyCode As Integer, Shift As Integer)
'If AdoCon.Recordset.AbsolutePosition > 1 Or (AdoCon.Recordset.AbsolutePosition = 1 And AdoCon.Recordset.RecordCount = 1) Then
'   If KeyCode = 13 Then GrdCon_DblClick
'End If
If KeyCode = 13 Then GrdCon_DblClick
End Sub

Public Property Get TipoCon() As Variant
TipoCon = mOpcion
End Property

Public Property Let TipoCon(ByVal vNewValue As Variant)
 mOpcion = vNewValue
End Property
