VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmCargaArchivo 
   Caption         =   "Carga de Archivos"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Carga de Archivos"
   ScaleHeight     =   5055
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CmbArchivo 
      Height          =   315
      ItemData        =   "FrmCargaArchivo.frx":0000
      Left            =   2280
      List            =   "FrmCargaArchivo.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog DlgAbrir 
      Left            =   6000
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton CmdBusqueda 
      Caption         =   "...."
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton CmdImporta 
      Caption         =   "Importar"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker FecIni 
      Height          =   315
      Left            =   2280
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   12615808
      CalendarTitleForeColor=   16777215
      Format          =   56360961
      CurrentDate     =   37616
   End
   Begin TrueOleDBGrid70.TDBGrid DGrd 
      Height          =   2820
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   4974
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
      AllowAddNew     =   -1  'True
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha del"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   705
   End
   Begin VB.Label TxtRuta 
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Ruta Archivo Origen"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione Tipo de Archivo"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "FrmCargaArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xlApp As Excel.Application
Dim mySheet As Excel.Worksheet
Dim rs As New ADODB.Recordset
Dim Codigo As String
Dim Dias As Integer
Dim Cadena As String
Dim narch As Integer
'Dim DlgAbrir As CommonDialog
Private Sub CmdBusqueda_Click()
    'TxtRuta.Caption = DLG_BrowseFolder(Me.hwnd, "Seleccione directorio")
    nombre_archivo = ""
    DlgAbrir.DialogTitle = "Abrir en carpeta"
    DlgAbrir.Filter = "Archivos Excel xls|*.xls"
    DlgAbrir.InitDir = "C:\"
    DlgAbrir.ShowOpen
    TxtRuta.Caption = DlgAbrir.FileName
    Me.Refresh
End Sub
Sub Valida()
    'If CmbArchivo.count > 0 Then CmbArchivo.ListIndex = 0
    Select Case CmbArchivo.Text
        Case "FALTAS EMPLEADOS"
            narch = 1
        Case "FALTAS"
            narch = 2
        Case "VACACIONES"
            narch = 3
        Case "DIVERSOS"
            narch = 4
        Case Else
            narch = 5
    End Select
 

   Sql$ = "select * from Pla_Dias_Importa where cia='01' and year(Fecha)='" & Format(FecIni.Year, "0000") & "' and month(fecha)='" & Format(FecIni.Month, "00") & "' and status<>'*' and tipo =" & narch
   cn.CursorLocation = adUseClient
   Set rs = New ADODB.Recordset
   Set rs = cn.Execute(Sql$, 64)
   If rs.RecordCount > 0 Then
      If MsgBox("La BD ya contiene informacion para este Archivo-Periodo Desea Eliminarla y volver a Importar", vbDefaultButton2 + vbYesNo + vbQuestion, "") = vbNo Then narch = 6: Exit Sub
      Sql = "update Pla_Dias_Importa set status='*' where cia='" & wcia & "' and year(Fecha)='" & Format(FecIni.Year, "0000") & "' and month(fecha)='" & Format(FecIni.Month, "00") & "'  and status<>'*' and tipo=" & narch
      cn.CursorLocation = adUseClient
      Set rs = New ADODB.Recordset
      Set rs = cn.Execute(Sql$, 64)
   End If
   If rs.State = 1 Then rs.Close
    
End Sub
Private Sub CmdImporta_Click()
    Dim Rproceso As Integer
    Call Valida
    
    If narch = 5 Then
        MsgBox "Plantilla de Archivo -- No seleccionada ", vbCritical, TitMsg
        Exit Sub
    End If
    
    If Trim(TxtRuta.Caption) = "" Then
      MsgBox "Plantilla de Archivo -- No seleccionada ", vbCritical, TitMsg
      Exit Sub
    End If
    
    If narch = 6 Then
        Exit Sub
    End If
        
        
    Dim xlApp2 As Excel.Application
    Dim xlApp1  As Excel.Application
    Dim xLibro  As Excel.Workbook
    
    'Chequeamos si excel esta corriendo
        
    'Set xlApp1 = GetObject(, "Excel.Application")
    If xlApp1 Is Nothing Then
        Set xlApp1 = CreateObject("Excel.Application")
    End If
    
    On Error GoTo ERR
    Set xlApp2 = xlApp1.Application
    Dim Col As Integer, Fila As Integer
       
    Set xLibro = xlApp2.Workbooks.Open(TxtRuta.Caption)
    
    xlApp2.Visible = False
    
    If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
    If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
    If Not xLibro Is Nothing Then Set xlBook = Nothing
    
    Dim conexion As ADODB.Connection, rs As ADODB.Recordset
  
    Set conexion = New ADODB.Connection
       
    conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & TxtRuta.Caption & _
                  ";Extended Properties=""Excel 8.0;HDR=Yes;"""
      
        Set rsExport = New ADODB.Recordset
       
    With rsExport
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
    End With
    
    
    rsExport.Open "SELECT * FROM [IMPORTA$]", conexion, , , adCmdText
    
    Set DGrd.DataSource = rsExport
    
    Dim Monto As Long
    Dim Total As Long
    Total = 0
    Rproceso = 0
    DGrd.MoveFirst
    For I = 0 To DGrd.ApproxCount - 1
    
        If Mid(DGrd.Columns(1), 1, 1) = "E" Or Mid(DGrd.Columns(1), 1, 1) = "O" Then
            If narch = 1 Then 'Empleados
                If Val(DGrd.Columns(3)) + Val(DGrd.Columns(4)) > 0 Then
                    Cadena = "SP_IMPORTA_ARCHIVO '01'," & narch & ",'" & Format(FecIni.Value, "mm/dd/yyyy") & "','" & Trim(DGrd.Columns(1)) & "', " & Val(DGrd.Columns(3)) & "," & Val(DGrd.Columns(4)) & ",0,0,0,0,0,0," & Val(DGrd.Columns(4)) + Val(DGrd.Columns(3)) & ",'','sa','sa'"
                    If Not EXEC_SQL(Cadena, cn) Then
                       Rproceso = 1
                    Else
                       Rproceso = 0
                    End If
                'MsgBox DGrd.Columns(1)
                End If
            End If
            If narch = 2 Then 'Obreros
                If Val(DGrd.Columns(4)) > 0 Then
                    Cadena = "SP_IMPORTA_ARCHIVO '01'," & narch & ",'" & Format(FecIni.Value, "mm/dd/yyyy") & "','" & Trim(DGrd.Columns(1)) & "', " & Val(DGrd.Columns(4)) & ",0,0,0,0,0,0,0," & Val(DGrd.Columns(4)) & ",'','sa','sa','" & DGrd.Columns(0) & "','" & DGrd.Columns(2) & "','" & DGrd.Columns(3) & "','" & DGrd.Columns(5) & "','" & DGrd.Columns(6) & "'"
                    'Cadena = "SP_IMPORTA_ARCHIVO '01'," & narch & ",'" & Format(FecIni.Value, "mm/dd/yyyy") & "','" & Trim(DGrd.Columns(1)) & "', " & Val(DGrd.Columns(4)) & ",0,0,0,0,0,0,0," & Val(DGrd.Columns(4)) & ",'','sa','sa'"
                    If Not EXEC_SQL(Cadena, cn) Then
                       Rproceso = 1
                    Else
                       Rproceso = 0
                    End If
                'MsgBox DGrd.Columns(1)
                End If
            End If
            If narch = 3 Then 'Vacaciones
                If Val(DGrd.Columns(3)) > 0 Then
                    Cadena = "SP_IMPORTA_ARCHIVO '01'," & narch & ",'" & Format(FecIni.Value, "mm/dd/yyyy") & "','" & Trim(DGrd.Columns(1)) & "',0,0, " & Val(DGrd.Columns(3)) & ",0,0,0,0,0," & Val(DGrd.Columns(3)) & ",'','sa','sa','" & DGrd.Columns(0) & "','" & DGrd.Columns(2) & "','" & DGrd.Columns(3) & "','" & DGrd.Columns(5) & "','" & DGrd.Columns(6) & "'"
                    If Not EXEC_SQL(Cadena, cn) Then
                       Rproceso = 1
                    Else
                       Rproceso = 0
                    End If
                'MsgBox DGrd.Columns(1)
                End If
            End If
            If narch = 4 Then 'Varios
                If Val(DGrd.Columns(3)) + Val(DGrd.Columns(4)) + Val(DGrd.Columns(5)) + Val(DGrd.Columns(6)) + Val(DGrd.Columns(7)) > 0 Then
                    Cadena = "SP_IMPORTA_ARCHIVO '01'," & narch & ",'" & Format(FecIni.Value, "mm/dd/yyyy") & "','" & Trim(DGrd.Columns(1)) & "',0,0,0, " & Val(DGrd.Columns(3)) & "," & Val(DGrd.Columns(4)) & "," & Val(DGrd.Columns(5)) & "," & Val(DGrd.Columns(6)) & "," & Val(DGrd.Columns(7)) & "," & Val(DGrd.Columns(3)) + Val(DGrd.Columns(4)) + Val(DGrd.Columns(5)) + Val(DGrd.Columns(6)) + Val(DGrd.Columns(7)) & ",'','sa','sa'"
                    If Not EXEC_SQL(Cadena, cn) Then
                       Rproceso = 1
                    Else
                       Rproceso = 0
                    End If
                'MsgBox DGrd.Columns(1)
                End If
            End If
        End If
        If I < DGrd.ApproxCount - 1 Then
          'DGrd.Row = DGrd.Row + 1
          DGrd.MoveNext
        End If
    Next I
    If Rproceso = 1 Then
       MsgBox "Error al Grabar el registro.", vbExclamation + vbOKOnly, Me.Caption
    Else
       MsgBox "Se grabó satisfactoriamente el registro.", vbInformation + vbOKOnly, Me.Caption
    End If
    'Set rsExport = Nothing
    'LimpiarRsT rsExport, DGrd
     
Salir:

xLibro.Close
If Not xlApp2 Is Nothing Then Set xlApp2 = Nothing
If Not xlApp1 Is Nothing Then Set xlApp1 = Nothing
If Not xLibro Is Nothing Then Set xlBook = Nothing
    Exit Sub
ERR:
    MsgBox ERR.Number & "-" & ERR.Description, vbCritical, Me.Caption
    Exit Sub
End Sub
Public Sub LimpiarRsT(ByRef pRs As ADODB.Recordset, ByRef pDgrd As TrueOleDBGrid70.Tdbgrid)
pDgrd.Refresh
Set pDgrd.DataSource = Nothing
If pRs.State = 1 Then
    If pRs.RecordCount > 0 Then
        pRs.MoveFirst
        Do While Not pRs.EOF
            pRs.Delete
            If Not pRs.EOF Then pRs.MoveNext
        Loop
    End If
End If
Set pDgrd.DataSource = pRs
pDgrd.Refresh
End Sub


Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
FecIni.Value = DateTime.Now
End Sub
