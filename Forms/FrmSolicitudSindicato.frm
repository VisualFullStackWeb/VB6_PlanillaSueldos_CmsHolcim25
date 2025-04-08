VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmSolicitudSindicato 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "» Solicitud de Sindicato «"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "FrmSolicitudSindicato.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkAll 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   2640
      Width           =   375
   End
   Begin VB.CheckBox chkTodos 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Height          =   1680
      Left            =   120
      TabIndex        =   10
      Top             =   495
      Width           =   9615
      Begin VB.ComboBox Cmbtipotrabajador 
         Height          =   315
         ItemData        =   "FrmSolicitudSindicato.frx":030A
         Left            =   7320
         List            =   "FrmSolicitudSindicato.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker Cmbfecha 
         Height          =   285
         Left            =   750
         TabIndex        =   2
         Top             =   705
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   37265
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   5790
         TabIndex        =   4
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
         Left            =   8160
         TabIndex        =   6
         Top             =   705
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16515073
         CurrentDate     =   37265
      End
      Begin MSComCtl2.DTPicker Cmbdel 
         Height          =   285
         Left            =   6600
         TabIndex        =   5
         Top             =   705
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16515073
         CurrentDate     =   37267
      End
      Begin VB.TextBox Txtsemana 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   705
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
      Begin MSMask.MaskEdBox txtMonto 
         Height          =   300
         Left            =   1680
         TabIndex        =   19
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Monto Solicitado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Trabajador"
         Height          =   195
         Left            =   6120
         TabIndex        =   17
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   7920
         TabIndex        =   16
         Top             =   705
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         Height          =   195
         Left            =   6240
         TabIndex        =   15
         Top             =   705
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana"
         Height          =   195
         Left            =   4680
         TabIndex        =   14
         Top             =   705
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
         Top             =   705
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
      Width           =   9855
      Begin VB.ComboBox Cmbcia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   60
         Width           =   8250
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
      Height          =   4455
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7858
      _LayoutType     =   4
      _RowHeight      =   17
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Sel"
      Columns(0).DataField=   "Action"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "PlaCod"
      Columns(1).DataField=   "PlaCod"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Trabajador"
      Columns(2).DataField=   "Trabajador"
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
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2117"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2037"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=8193"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=12647"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=12568"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=8192"
      Splits(0)._ColumnProps(18)=   "Column(2).WrapText=1"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=0,.wraptext=-1,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECCIONE A TRABAJADORES QUE NO SE LE DEBE DESCONTAR"
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
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   9615
   End
End
Attribute VB_Name = "FrmSolicitudSindicato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VTipo As String
Dim VTipotrab As String
Dim VIdSolSindicato As Integer
Dim rsSindicato As New ADODB.Recordset


Public Sub Grabar()
Dim NroTrans As Integer
Dim contar As Integer
On Error GoTo ErrorTrans
NroTrans = 0
If Me.Cmbtipo.ListIndex < 0 Then
    MsgBox "Seleccione el Tipo de Boleta", vbCritical, Me.Caption
    Exit Sub
End If

If Me.Cmbtipotrabajador.ListIndex < 0 Then
    MsgBox "Seleccione el Tipo de Trabajador", vbCritical, Me.Caption
    Exit Sub
End If

If VTipotrab = "02" Then
    If Val(Txtsemana.Text) <= 0 Then
        MsgBox "Ingrese una semana", vbCritical, Me.Caption
        Exit Sub
    Else
        Sql$ = "select * from plasemanas"
        Sql$ = Sql$ & " where cia='" & wcia & "' and semana='"
        Sql$ = Sql$ & Format(Trim(Txtsemana.Text), "00") & "' and ano=" & Year(Cmbfecha.Value) & " and status !='*'"
        If Not (fAbrRst(rs, Sql$)) Then
            MsgBox "Ingrese una semana valida", vbCritical, Me.Caption
            Exit Sub
        End If
    End If
Else
    Txtsemana.Text = ""
End If

If Val(TxtMonto.Text) <= 0 Then
    MsgBox "Ingrese el Importe", vbCritical, Me.Caption
    Exit Sub
End If

If rsSindicato.RecordCount <= 0 Then
    MsgBox "No hay Trabajadores afiliados al sindicato", vbCritical, Me.Caption
    Exit Sub
Else
    GrdSolicitud.Update
    GrdSolicitud.BatchUpdates = True
    If rsSindicato.EOF = False Then rsSindicato.Update
End If

cn.BeginTrans
NroTrans = 1
contar = 0
Sql$ = "usp_Pla_PlaSolicitudSindicato " & Val(VIdSolSindicato)
Sql$ = Sql$ & ",'" & wcia & "','01','" & Trim(VTipo) & "','" & Trim(VTipotrab)
Sql$ = Sql$ & "'," & Year(Cmbfecha.Value) & ",'" & IIf(VTipotrab = "02", Format(Trim(Me.Txtsemana.Text), "00"), Format(Trim(Str(Month(Cmbfecha.Value))), "00")) & "','" & Format(Cmbfecha.Value, "mm/dd/yyyy")
Sql$ = Sql$ & "','S/.'," & Val(TxtMonto.Text) & ",'','" & wuser & "',"
If Val(VIdSolSindicato) = 0 Then
    Sql$ = Sql$ & "1"
Else
    Sql$ = Sql$ & "2"
End If

If (fAbrRst(rs, Sql$)) Then
    VIdSolSindicato = rs!Id
    rsSindicato.MoveFirst
    
    
    Do While Not rsSindicato.EOF
        If rsSindicato!Action = True Then
            Sql$ = "usp_Pla_PlaSolicitudSindicatoDetalle " & Val(VIdSolSindicato)
            Sql$ = Sql$ & ",'" & wcia & "','" & Trim(rsSindicato!PlaCod) & "'," & Val(TxtMonto.Text)
            Sql$ = Sql$ & "," & Val(rsSindicato!importe) & ",'','" & wuser & "'"
            cn.Execute Sql$
            contar = contar + 1
        End If

        rsSindicato.MoveNext
    Loop
    
    
End If
        
'If contar <= 0 Then
'    cn.RollbackTrans
'    MsgBox "Debe Checkear al menos a un trabajador", vbInformation, Me.Caption
'    Exit Sub
'Else
'    cn.CommitTrans
'End If
    
cn.CommitTrans

MsgBox "Se grabaron los datos correctamente", vbInformation, Me.Caption

Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

MsgBox Err.Description, vbCritical, Me.Caption

End Sub
Public Sub Eliminar()
Dim NroTrans As Integer
On Error GoTo ErrorTrans
NroTrans = 0
If VIdSolSindicato <= 0 Then
    MsgBox "Seleccione alguna Solicitud de Sindicato", vbCritical, Me.Caption
    Exit Sub
End If
If MsgBox("¿Esta seguro de eliminar los datos?", vbExclamation + vbYesNo, Me.Caption) = vbNo Then Exit Sub
cn.BeginTrans
NroTrans = 1

Sql$ = "usp_Pla_PlaSolicitudSindicato " & Val(VIdSolSindicato)
Sql$ = Sql$ & ",'" & wcia & "','01','" & Trim(VTipo) & "','" & Trim(VTipotrab)
Sql$ = Sql$ & "'," & Year(Cmbfecha.Value) & ",'','" & Format(Cmbfecha.Value, "mm/dd/yyyy")
Sql$ = Sql$ & "','S/.'," & Val(TxtMonto.Text) & ",'','" & wuser & "',"
Sql$ = Sql$ & "3"

cn.Execute Sql$
cn.CommitTrans
VIdSolSindicato = 0
MsgBox "Se Eliminaron los datos correctamente", vbInformation, Me.Caption

Exit Sub

ErrorTrans:
If NroTrans = 1 Then
    cn.RollbackTrans
End If

MsgBox Err.Description, vbCritical, Me.Caption

End Sub

Private Sub ChkAll_Click()
Dim lChk As Boolean
If ChkAll.Value Then lChk = True Else lChk = False
If rsSindicato.RecordCount > 0 Then rsSindicato.MoveFirst
Do While Not rsSindicato.EOF
   rsSindicato!Action = lChk
   rsSindicato.MoveNext
Loop
End Sub

Private Sub chkTodos_Click()
If rsSindicato.RecordCount > 0 Then
    rsSindicato.MoveFirst
    Do While Not rsSindicato.EOF
        rsSindicato!Action = CBool(chkTodos.Value)
        rsSindicato.Update
        rsSindicato.MoveNext
    Loop
End If
End Sub

Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))


Call fc_Descrip_Maestros2("01078", "", Cmbtipo)

If Cmbtipo.ListCount = 1 Then Cmbtipo.ListIndex = 0

   Call fc_Descrip_Maestros2("01055", "", Cmbtipotrabajador)
   Procesa_Cabeza_Solicitud


End Sub

Private Sub Cmbfecha_Change()
'If Month(Cmbfecha.Value) = 1 And VTipo = "02" And Cmbtipotrabajador.ListIndex >= 0 Then Command1.Enabled = True Else Command1.Enabled = False
Procesa_Cabeza_Solicitud
End Sub

Private Sub CmbTipo_Click()
VTipo = fc_CodigoComboBox(Cmbtipo, 2)
Txtsemana.Text = ""
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
   If Month(Cmbfecha.Value) = 1 And Cmbtipotrabajador.ListIndex >= 0 Then Command1.Enabled = True Else Command1.Enabled = False
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
Procesa_Cabeza_Solicitud
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
Procesa_Cabeza_Solicitud
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

Lblfecha.Caption = Format(Date, "dd/mm/yyyy")
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
 Procesa_Cabeza_Solicitud
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

If rsSindicato.State = 1 Then rsSindicato.Close
    rsSindicato.Fields.Append "Action", adBoolean, , adFldIsNullable
    rsSindicato.Fields.Append "PlaCod", adChar, 8, adFldIsNullable
    rsSindicato.Fields.Append "Trabajador", adVarChar, 500, adFldIsNullable
    rsSindicato.Fields.Append "Importe", adCurrency, , adFldIsNullable
    
    rsSindicato.Open
    Set GrdSolicitud.DataSource = rsSindicato
End Sub

Private Sub Procesa_Cabeza_Solicitud()
VIdSolSindicato = 0
TxtMonto.Text = "0.00"
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

Call Crea_Rs

If Cmbtipo.ListIndex < 0 Then Exit Sub
If Cmbtipotrabajador.ListIndex < 0 Then Exit Sub
If VTipotrab = "02" Then If Val(Txtsemana.Text) = 0 Then Exit Sub
Sql$ = "usp_Pla_Consultar_SolicitudSindicato '" & wcia
Sql$ = Sql$ & "','01','" & VTipo & "','" & VTipotrab & "','"
If VTipotrab = "02" Then
    Sql$ = Sql$ & Format(Trim(Txtsemana.Text), "00")
Else
    Sql$ = Sql$ & Trim(Txtsemana.Text)
End If
Sql$ = Sql$ & "','" & Format(Me.Cmbfecha.Value, "mm/dd/yyyy") & "'"
If (fAbrRst(rs, Sql$)) Then
    VIdSolSindicato = rs!Id
    TxtMonto.Text = rs!importe
End If

Sql$ = "usp_Pla_Consultar_SolicitudSindicatoDetalle " & Val(VIdSolSindicato)
Sql$ = Sql$ & ",'" & wcia
Sql$ = Sql$ & "','" & VTipotrab & "'"
If (fAbrRst(rs, Sql$)) Then
    rs.MoveFirst
    Do While Not rs.EOF
        rsSindicato.AddNew
        With rsSindicato
            !Action = rs!Action
            !PlaCod = rs!PlaCod
            !Trabajador = rs!Trabajador
            !importe = rs!importe
            .Update
        End With
        rs.MoveNext
    Loop

GrdSolicitud.Refresh

End If

End Sub



