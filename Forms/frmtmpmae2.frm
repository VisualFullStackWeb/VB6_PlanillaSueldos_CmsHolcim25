VERSION 5.00
Begin VB.Form frmtmpmae2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmtmpmae2.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.Data datmae3 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4290
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5970
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Data datmae2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4170
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstsino 
      Appearance      =   0  'Flat
      Height          =   420
      ItemData        =   "frmtmpmae2.frx":030A
      Left            =   3480
      List            =   "frmtmpmae2.frx":0314
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   6480
      Width           =   6615
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   $"frmtmpmae2.frx":0320
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   360
         TabIndex        =   1
         Top             =   120
         Width           =   5835
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Data datmae 
      Caption         =   "DatMae"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmtmpmae2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'Dim COD_MAES As String
'Dim cod$, TIPO_MAE$, sql$
'
'Dim COL1                    As MSDBGrid.Column, COL2 As MSDBGrid.Column, COL3 As MSDBGrid.Column, COL5 As MSDBGrid.Column
'Dim COL6                    As MSDBGrid.Column, COL7 As MSDBGrid.Column, COL8 As MSDBGrid.Column, COL9 As MSDBGrid.Column
'Dim col4                    As MSDBGrid.Column, ColUpRow As MSDBGrid.Column, colCodMae As MSDBGrid.Column, colCodMae2 As MSDBGrid.Column
'Dim ColUpRow2               As MSDBGrid.Column, COL21 As MSDBGrid.Column, colCodMae21 As MSDBGrid.Column, colCodMae22 As MSDBGrid.Column
'Dim ColUpRow3               As MSDBGrid.Column, COL31 As MSDBGrid.Column, colCodMae31 As MSDBGrid.Column, colCodMae32 As MSDBGrid.Column
'
'Private StateTool(3)        As String
'
'Dim Procesar_Grabar         As Boolean
'
'Dim Eliminados              As Integer
'
'Dim nroGrids                As Byte
'Dim Eliminar                As Boolean
'Dim UpRows, cntelim         As Integer
'Private Reciclaje()         As String
'
'
'Dim eliminar2               As Boolean
'Dim UpRows2, cntelim2       As Integer
'Private Reciclaje2()        As String
'
'Dim eliminar3               As Boolean
'Dim UpRows3, I, cntelim3      As Integer
'Private Reciclaje3()        As String
'
'Dim Nuevo                   As Boolean
'Dim FRM_LOAD                As Boolean
'
'Dim Modificado_Grid2        As Boolean
'Dim Modificado_Grid3        As Boolean
'Dim Sqlquery                As String
'Dim Adosql                  As ADODB.Recordset
'Dim vNroreg2, vNroreg3     As Integer
'Private Sub Form_Activate()
'On Error Resume Next
'frmtmpmae2.ZOrder 0
'
'If FRM_LOAD = True Then
'    grdmae.Col = 0
'    grdmae.Row = 0
'    grdmae.SetFocus
'    UpRows = datmae.Recordset.RecordCount
'    If nroGrids = 2 Then
'     grdmae2.Col = 0
'     grdmae2.Row = 0
'     grdmae2.SetFocus
'     Select Case cod
'        Case "046"
'            COL1.Caption = "TIPOS " & colCodMae
'            grdmae2.Caption = "TIPOS DE CANCELACION " & COL1 & " 01"
'        Case "049"
'            COL1.Caption = "DOCUMENTO BANCARIO " & colCodMae
'            grdmae2.Caption = "DOCUMENTO BANCARIO " & COL1 & " 01"
'    End Select
'     UpRows2 = datmae2.Recordset.RecordCount
'    End If
'    If nroGrids = 3 Then
'     grdmae3.Col = 0
'     grdmae3.Row = 0
'     grdmae3.SetFocus
'     UpRows3 = datmae3.Recordset.RecordCount
'End If
'FRM_LOAD = False
'End If
'End Sub
'Sub CREATE_TABLE()
'On Error Resume Next
'Dim tdf0 As TableDef
'datmae.DatabaseName = nom_BD
'datmae.Refresh
' Set tdf0 = datmae.Database.CreateTableDef("TmpMae01")
'    With tdf0
'        .Fields.Append .CreateField("descrip", dbText)
'         .Fields(0).AllowZeroLength = True
'        .Fields.Append .CreateField("col1", dbText)
'         .Fields(1).AllowZeroLength = True
'        .Fields.Append .CreateField("col2", dbText)
'         .Fields(2).AllowZeroLength = True
'        .Fields.Append .CreateField("col3", dbText)
'         .Fields(3).AllowZeroLength = True
'        .Fields.Append .CreateField("uprow", dbText)
'         .Fields(4).AllowZeroLength = True
'        .Fields.Append .CreateField("cod_mae", dbText)
'         .Fields(5).AllowZeroLength = True
'        .Fields.Append .CreateField("cod_mae2", dbText)
'         .Fields(6).AllowZeroLength = True
'         .Fields.Append .CreateField("flag4", dbText)
'         .Fields(7).AllowZeroLength = True
'         .Fields.Append .CreateField("flag5", dbText)
'         .Fields(8).AllowZeroLength = True
'         .Fields.Append .CreateField("flag6", dbText)
'         .Fields(9).AllowZeroLength = True
'         .Fields.Append .CreateField("flag7", dbText)
'         .Fields(10).AllowZeroLength = True
'         .Fields.Append .CreateField("codsunat", dbText)
'         .Fields(11).AllowZeroLength = True
'        datmae.Database.TableDefs.Append tdf0
'     End With
'datmae.RecordSource = "TmpMae01"
'datmae.Refresh
'
'If nroGrids = 2 Or nroGrids = 3 Then
'datmae2.DatabaseName = nom_BD
'datmae2.Refresh
' Set tdf0 = datmae2.Database.CreateTableDef("TmpMae02")
'    With tdf0
'        .Fields.Append .CreateField("descrip", dbText)
'         .Fields(0).AllowZeroLength = True
'        .Fields.Append .CreateField("uprow", dbText)
'         .Fields(1).AllowZeroLength = True
'        .Fields.Append .CreateField("cod_mae", dbText)
'         .Fields(2).AllowZeroLength = True
'        .Fields.Append .CreateField("cod_mae2", dbText)
'         .Fields(3).AllowZeroLength = True
'        .Fields.Append .CreateField("flag1", dbText)
'         .Fields(4).AllowZeroLength = True
'         .Fields(4).DefaultValue = ""
'        .Fields.Append .CreateField("flag2", dbText)
'         .Fields(5).AllowZeroLength = True
'         .Fields(5).DefaultValue = ""
'        .Fields.Append .CreateField("flag3", dbText)
'         .Fields(6).AllowZeroLength = True
'         .Fields(6).DefaultValue = ""
'        datmae2.Database.TableDefs.Append tdf0
'     End With
'datmae2.RecordSource = "TmpMae02"
'datmae2.Refresh
'End If
'
'If nroGrids = 3 Then
'datmae3.DatabaseName = nom_BD
'datmae3.Refresh
' Set tdf0 = datmae3.Database.CreateTableDef("TmpMae03")
'    With tdf0
'        .Fields.Append .CreateField("descrip", dbText)
'         .Fields(0).AllowZeroLength = True
'        .Fields.Append .CreateField("uprow", dbText)
'         .Fields(1).AllowZeroLength = True
'        .Fields.Append .CreateField("cod_mae", dbText)
'         .Fields(2).AllowZeroLength = True
'        .Fields.Append .CreateField("cod_mae2", dbText)
'         .Fields(3).AllowZeroLength = True
'        datmae3.Database.TableDefs.Append tdf0
'     End With
'datmae3.RecordSource = "TmpMae03"
'datmae3.Refresh
'End If
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 27 And grdmae2.Visible Then
'           grdmae2.Visible = False
'           grdmae.Height = 6405
'        End If
'End Sub
'
'Private Sub Form_Load()
'Dim wciamae As String
'
'Me.KeyPreview = True
'Screen.MousePointer = 11
'vNroreg2 = 0: vNroreg3 = 0
'Set COL1 = grdmae.Columns(0)
'Set COL2 = grdmae.Columns(1)
'Set COL3 = grdmae.Columns(2)
'Set col4 = grdmae.Columns(3)
'Set COL5 = grdmae.Columns(7)
'Set COL6 = grdmae.Columns(8)
'Set COL7 = grdmae.Columns(9)
'Set COL8 = grdmae.Columns(10)
'Set COL9 = grdmae.Columns(11)
'
'Set ColUpRow = grdmae.Columns(4)
'Set colCodMae = grdmae.Columns(5)
'Set colCodMae2 = grdmae.Columns(6)
'
'Set COL21 = grdmae2.Columns(0)
'
'Set ColUpRow2 = grdmae2.Columns(1)
'Set colCodMae21 = grdmae2.Columns(2)
'Set colCodMae22 = grdmae2.Columns(3)
'
'Set COL31 = grdmae3.Columns(0)
'Set ColUpRow3 = grdmae3.Columns(1)
'Set colCodMae31 = grdmae3.Columns(2)
'Set colCodMae32 = grdmae3.Columns(3)
'
'grdmae3.AllowUpdate = True: grdmae3.AllowUpdate = True: grdmae3.AllowUpdate = True
'grdmae3.AllowDelete = True: grdmae3.AllowDelete = True: grdmae3.AllowDelete = True
'COD_MAES = Trim(frmtmpmae.lstvmae.ListItems(frmtmpmae.lstvmae.SelectedItem.Index))
'
'If COD_MAES = "026" Then grdmae3.AllowUpdate = False: grdmae.AllowUpdate = False: grdmae3.AllowDelete = False: grdmae.AllowDelete = False
'If COD_MAES = "027" Then grdmae2.AllowUpdate = False: grdmae3.AllowUpdate = False: grdmae2.AllowDelete = False: grdmae3.AllowDelete = False
'
'FRM_LOAD = True
'
'nroGrids = frmtmpmae.NRO_GRIDS
'CREATE_TABLE
'Me.Left = 4840
'Me.Top = 20
'Me.Height = 7650
'Me.Width = 6855
'Me.Caption = "Mantenimiento de " & frmtmpmae.nom_mae
'cod$ = frmtmpmae.COD_MAE
'TIPO_MAE$ = frmtmpmae.TIPO_MAE
'
'sql$ = "SELECT GENERAL FROM maestros where ciamaestro = '" & "01" + cod$ & "'  and status<>'*' "
'If (fAbrRst(rs, sql$)) Then
'   If rs!General = "S" Then
'      wciamae = " and right(ciamaestro,3)= '" & cod$ & "' ORDER BY cod_maestro2"
'   Else
'      wciamae = " and ciamaestro= '" & wcia + cod$ & "' ORDER BY cod_maestro2"
'   End If
'End If
'If rs.State = 1 Then rs.Close
'
'If frmtmpmae.NUEVO2 = True Then
'  Nuevo = True
'   cod = frmtmpmae.codMaestro
'
'       grdmae.HeadLines = 2
'       COL1.Width = 3000
'
'       If Not frmtmpmae.dat.Recordset.BOF Then frmtmpmae.dat.Recordset.MoveFirst
'       Do While Not frmtmpmae.dat.Recordset.EOF
'         Select Case frmtmpmae.dat.Recordset.AbsolutePosition
'         Case 0
'           COL1.Caption = frmtmpmae.dat.Recordset(1)
'         Case 1
'           COL2.Visible = True
'           COL2.Caption = frmtmpmae.dat.Recordset(1)
'         Case 2
'           COL3.Visible = True
'           COL3.Caption = frmtmpmae.dat.Recordset(1)
'         End Select
'       frmtmpmae.dat.Recordset.MoveNext
'       Loop
'   frmtmpmae.NUEVO2 = False
'   Screen.MousePointer = vbDefault
'   Exit Sub
'Else: Nuevo = False
'
'End If
'Eliminar = True
'eliminar2 = True
'eliminar3 = True
'Select Case TIPO_MAE$
' Case "M"
'   Select Case cod$
'    Case "006", "026"
'      'MONEDAS
'      COL2.Visible = True
'      Sqlquery = " SELECT cod_maestro2,descrip,flag1 FROM maestros_2 " _
'       & " where status=''"
'       Sqlquery = Sqlquery & wciamae
'        cn.CursorLocation = adUseClient
'        Set Adosql = New ADODB.Recordset
'        Set Adosql = cn.Execute(Sqlquery)
'                Do While Not Adosql.EOF
'                  datmae.Recordset.AddNew
'                  datmae.Recordset!descrip = Adosql!descrip
'                  datmae.Recordset!COL1 = Adosql!flag1
'                  datmae.Recordset!COD_MAE = Trim(Adosql!cod_maestro2)
'                  datmae.Recordset!cod_mae2 = Trim(Adosql!cod_maestro2)
'                  datmae.Recordset.Update
'                  vNroreg2 = Val(Adosql!cod_maestro2)
'                  Adosql.MoveNext
'                Loop
'    Case "017"
'      'COMPROBANTES DE PAGO
'      COL2.Visible = True
'      Sqlquery = " SELECT * FROM maestros_2 " _
'       & " where status=''"
'       Sqlquery = Sqlquery & wciamae
'        cn.CursorLocation = adUseClient
'        Set Adosql = New ADODB.Recordset
'        Set Adosql = cn.Execute(Sqlquery)
'
'      COL2.Caption = "Simbolo"
'      COL3.Caption = "Aux 1"
'      col4.Caption = "Aux 2"
'      COL5.Caption = "Aux 3"
'      COL6.Caption = "Aux 4"
'      COL7.Caption = "Aux 5"
'      COL8.Caption = "Aux 6"
'      COL9.Caption = "Cod Sunat"
'
'      COL2.Visible = True
'      COL3.Visible = True
'      col4.Visible = True
'      COL5.Visible = True
'      COL6.Visible = True
'      COL7.Visible = True
'      COL8.Visible = True
'      COL9.Visible = True
'
'                Do While Not Adosql.EOF
'                  datmae.Recordset.AddNew
'                  datmae.Recordset!descrip = Adosql!descrip
'                  datmae.Recordset!COL1 = Adosql!flag1 & ""
'                  datmae.Recordset!COL2 = Adosql!flag2 & ""
'                  datmae.Recordset!COL3 = Adosql!flag3 & ""
'                  datmae.Recordset!FLAG4 = Adosql!FLAG4 & ""
'                  datmae.Recordset!flag5 = Adosql!flag5 & ""
'                  datmae.Recordset!flag6 = Adosql!flag6 & ""
'                  datmae.Recordset!flag7 = Adosql!flag7 & ""
'                  datmae.Recordset!codsunat = Adosql!codsunat & ""
'
'
'                  datmae.Recordset!COD_MAE = Adosql!cod_maestro2
'                  datmae.Recordset!cod_mae2 = Adosql!cod_maestro2
'                  datmae.Recordset.Update
'                  vNroreg2 = Val(Adosql!cod_maestro2)
'                  Adosql.MoveNext
'                Loop
'
'    Case Else
'      Dim rs3 As ADODB.Recordset
'      'bUSCAR LOS INGRESADOS CON MAS DE 1 cOLUMNA
'      Sqlquery = " SELECT * FROM maestros_2 " _
'           & " where status=''"
'           Sqlquery = Sqlquery & wciamae
'     cn.CursorLocation = adUseClient
'     Set rs3 = New ADODB.Recordset
'     Set rs3 = cn.Execute(Sqlquery)
'      If Not rs3.EOF Then
'        If Not IsNull(rs3!flag6) And rs3!flag6 <> "" Then
'           grdmae.HeadLines = 2
'            COL2.Visible = True
'            COL2.Caption = "Columna 2"
'            COL3.Visible = True
'            COL3.Caption = "Columna 3"
'
'           Sqlquery = " SELECT cod_maestro2,descrip,flag6,flag7 FROM maestros_2 " _
'           & " where status=''"
'           Sqlquery = Sqlquery & wciamae
'            cn.CursorLocation = adUseClient
'            Set Adosql = New ADODB.Recordset
'            Set Adosql = cn.Execute(Sqlquery)
'                Do While Not Adosql.EOF
'                  datmae.Recordset.AddNew
'                  datmae.Recordset!descrip = Adosql!descrip
'                  datmae.Recordset!COL1 = Adosql!flag6
'                  datmae.Recordset!COL2 = Adosql!flag7
'                  datmae.Recordset!COD_MAE = Trim(Adosql!cod_maestro2)
'                  datmae.Recordset!cod_mae2 = Trim(Adosql!cod_maestro2)
'                  datmae.Recordset.Update
'                  vNroreg2 = Val(Adosql!cod_maestro2)
'                  Adosql.MoveNext
'                Loop
'
'
'        Else
'
'                Sqlquery = " SELECT cod_maestro2,descrip FROM maestros_2 " _
'                & " where status=''"
'                Sqlquery = Sqlquery & wciamae
'                cn.CursorLocation = adUseClient
'                Set Adosql = New ADODB.Recordset
'                Set Adosql = cn.Execute(Sqlquery)
'
'                Do While Not Adosql.EOF
'                  datmae.Recordset.AddNew
'                  datmae.Recordset!descrip = Adosql(1)
'                  datmae.Recordset!COD_MAE = Trim(Adosql!cod_maestro2)
'                  datmae.Recordset!cod_mae2 = Trim(Adosql!cod_maestro2)
'                  datmae.Recordset.Update
'                  vNroreg2 = Val(Adosql!cod_maestro2)
'                  Adosql.MoveNext
'                Loop
'
'        End If
'        End If
'  End Select
'
'
'End Select
'Screen.MousePointer = vbDefault
'frmtmpmae.NUEVO2 = False
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'Set grdmae.DataSource = Nothing
'datmae.RecordSource = ""
'datmae.Refresh
'datmae.Database.TableDefs.Delete "TmpMae01"
'datmae.Database.Close
'
'If nroGrids = 2 Or nroGrids = 3 Then
'Set grdmae2.DataSource = Nothing
'datmae2.RecordSource = ""
'datmae2.Refresh
'datmae2.Database.TableDefs.Delete "TmpMae02"
'datmae2.Database.Close
'End If
'
'If nroGrids = 3 Then
'Set grdmae3.DataSource = Nothing
'datmae3.RecordSource = ""
'datmae3.Refresh
'datmae3.Database.TableDefs.Delete "TmpMae03"
'datmae3.Database.Close
'End If
'
'End Sub
'
'Private Sub grdmae_AfterColEdit(ByVal ColIndex As Integer)
'If colCodMae <> colCodMae2 Then ColUpRow = "I" Else ColUpRow = "U"
'
'End Sub
'
'Private Sub grdmae_BeforeDelete(Cancel As Integer)
'If Eliminar = True Then
'
' UpRows = datmae.Recordset.RecordCount
'   If UpRows = 0 Then Eliminar = False
'    If datmae.Recordset.AbsolutePosition + 1 <= UpRows Then
'     cntelim = cntelim + 1
'     ReDim Preserve Reciclaje(cntelim)
'     Reciclaje(cntelim) = datmae.Recordset!COD_MAE
'
'     If UpRows > 0 Then UpRows = UpRows - 1
'    End If
'End If
'End Sub
'
'Private Sub grdmae_ButtonClick(ByVal ColIndex As Integer)
'Dim Y As Integer
'      Y = grdmae.Row
'      If Y < 15 Then
'       lstsino.Top = grdmae.Top + grdmae.RowTop(Y) + grdmae.RowHeight
'       lstsino.Left = grdmae.Left + COL3.Left
'      Else
'       lstsino.Top = grdmae.Top + grdmae.RowTop(Y) - lstsino.Height
'      End If
'      lstsino.Visible = True
'      lstsino.SetFocus
'End Sub
'Function VALIDAR() As Boolean
'datmae.Refresh
'If datmae.Recordset.RecordCount > 0 Then
'     If Not datmae.Recordset.BOF Then
'       datmae.Recordset.MoveFirst
'     End If
'      Do While Not datmae.Recordset.EOF
'         If ColUpRow <> "" Then
'           If Trim(COL1) = "" Or IsNull(COL1) Then
'             MsgBox "Debe haber una Descripción", vbCritical
'             grdmae.Col = 0
'             grdmae.SetFocus
'             GoTo salir
'           'ElseIf COL2.Visible = True And cod <> "005" And cod <> "014" And COL2 = "" Or IsNull(COL2) Then
'           '     MsgBox "En " & COL2.Caption & Chr(13) & "hay una celda en blanco", vbCritical
'           '     grdmae.Col = 0
'           '     grdmae.SetFocus
'           '     GoTo salir
'           'ElseIf COL3.Visible = True And cod <> "005" And COL3 = "" Or IsNull(COL3) Then
'           '    MsgBox "En " & COL3.Caption & Chr(13) & "hay una celda en blanco", vbCritical
'           '     grdmae.Col = 0
'           '     grdmae.SetFocus
'           '     GoTo salir
'           End If
'         End If
'        datmae.Recordset.MoveNext
'      Loop
'End If
'
'datmae2.Refresh
'      If grdmae2.Visible = True Then
'         If Not datmae2.Recordset.BOF Then datmae2.Recordset.MoveFirst
'
'         Do While Not datmae2.Recordset.EOF
'            If ColUpRow2 <> "" Then
'                If Trim(COL21) = "" Or IsNull(COL21) Then
'                    MsgBox "Debe haber una Descripción", vbCritical
'                    grdmae2.Col = 0
'                    grdmae2.SetFocus
'                    GoTo salir
'                End If
'         End If
'         datmae2.Recordset.MoveNext
'         Loop
'      End If
'
'    '  If grdmae3.Visible = True Then
'    '     If Not datmae3.Recordset.BOF Then datmae3.Recordset.MoveFirst
'    '
'    '     Do While Not datmae3.Recordset.EOF
'    '     If ColUpRow3 <> "" Then
'    '       If Trim(COL31) = "" Or IsNull(COL31) Then
'    '         MsgBox "Debe haber una Descripción", vbCritical
'    '         grdmae3.Col = 0
'    '         grdmae3.SetFocus
'    '         GoTo salir
'    '       End If
'     '    End If
'     '    datmae3.Recordset.MoveNext
'    '     Loop
'    '  End If
'   VALIDAR = True
'   Exit Function
'salir:
'   VALIDAR = False
'   Exit Function
'
'End Function
'
'Public Sub Grabar()
'
'Screen.MousePointer = 0
'
'If VALIDAR = False Then Exit Sub
'
'If datmae.Recordset.RecordCount = 0 Then Exit Sub
'If MsgBox("Desea Grabar los cambios", vbQuestion + vbYesNo) = vbNo Then Unload Me: Exit Sub
'
''On Error GoTo Err_Grabar
'
'  Sqlquery = wInicioTrans
'  cn.Execute Sqlquery, Options:=rdExecDirect
'
'  Screen.MousePointer = vbHourglass
'  '*********************************************************************************************
'       If eliminar3 = True And cntelim3 > 0 Then
'        If UBound(Reciclaje3) > 0 Then
'          For I = 1 To UBound(Reciclaje3)
'              Sqlquery = "DELETE carlinnat3 " _
'                   & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Trim(colCodMae) & "' and cod_nat='" & Trim(Reciclaje3(I)) & "' "
'              cn.Execute Sqlquery, Options:=rdExecDirect
'          Next I
'        End If
'       End If
'  '********************************************************************************************************************
'  '******************eLIMINAR DE GRID2***************************************************************************
'       If eliminar2 = True And cntelim2 > 0 Then
'        If UBound(Reciclaje2) > 0 Then
'          For I = 1 To UBound(Reciclaje2)
'            Select Case TIPO_MAE$
'             Case "M"
'               Sqlquery = "DELETE maestros_3  " _
'                   & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Right(grdmae2.Caption, 2) & "' and cod_maestro3='" & Trim(Reciclaje2(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'             Case "C"
'               Sqlquery = "DELETE carlinnat2 " _
'                   & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Trim(colCodMae21) & "' and cod_nat='" & Trim(Reciclaje2(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'             End Select
'          Next I
'        End If
'      End If
'      '********************************************************************************************************************
'     If Eliminar = True And cntelim > 0 Then
'       If UBound(Reciclaje) > 0 Then
'         For I = 1 To UBound(Reciclaje)
'            Select Case TIPO_MAE$
'             Case "M"
'               Sqlquery = "DELETE FROM maestros_2  " _
'                   & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Trim(Reciclaje(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'
'               Sqlquery = "DELETE FROM maestros_3  " _
'                   & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Trim(Reciclaje(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'
'             Case "C"
'               Sqlquery = "DELETE carlinnat1  " _
'                   & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Trim(Reciclaje(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'             Case "T01", "T02"
'                Sqlquery = "DELETE TABLAS  " _
'                     & " WHERE cod_tipo = '" & cod & "' AND cod_maestro2 ='" & Trim(Reciclaje(I)) & "' "
'                cn.Execute Sqlquery, Options:=rdExecDirect
'
'             Case "F"
'                 Sqlquery = "DELETE FLETES " _
'                      & " WHERE cod_cia = '" & wcia & "' AND cod_flete='" & Trim(Reciclaje(I)) & "' "
'                 cn.Execute Sqlquery, Options:=rdExecDirect
'
'            End Select
'         Next I
'       End If
'      End If
'
'       datmae.Recordset.MoveFirst
'       Do While Not datmae.Recordset.EOF
'        If Trim(ColUpRow) <> "" Then
'          Select Case TIPO_MAE$
'          Case "M"
'               Select Case cod$
'               Case "005"
'                  'CONDICION DE pAGOS
'                  If ColUpRow = "I" Then
'                        'Dim adoSQL As ADODB.rdoResultset
'                        Sqlquery = "SELECT * FROM maestros_2 WHERE ciamaestro='" & wcia & "005' ORDER BY cod_maestro2"
'                        cn.CursorLocation = adUseClient
'                        Set Adosql = New ADODB.Recordset
'                        Set Adosql = cn.Execute(Sqlquery)
'                        If Not Adosql.EOF Then
'                            colCodMae = "01"
'                        Else
'                            Adosql.MoveLast
'                            colCodMae = Format(Val(Adosql!cod_maestro2) + 1, "00")
'                        End If
'
'                    Sqlquery = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " INSERT INTO maestros_2 " _
'                         & " VALUES ('" & wcia & cod & "','" & Trim(colCodMae) & "','" & Trim(COL1) & "', " _
'                         & " NULL,'" & Trim(COL2) & "','" & Trim(COL3) & "','" & Trim(col4) & "',NULL,'','" & _
'                         wuser & "',NULL," & FechaSys & ",NULL,NULL,NULL) "
'                         '& " If '" & colCodMae2 & "' <> '' BEGIN DELETE maestros_2 SET status = '*',user_modi='" & WUSER & "',fec_modi=" & FechaSys & "() WHERE cialistp='" & wcialistp & "' AND cod_prod='" & Trim(ColCodProd) & "' AND status='' END"
'
'                    cn.Execute Sqlquery, Options:=rdExecDirect
'                  Else
'                     Sqlquery = "UPDATE maestros_2 SET descrip = '" & Trim(COL1) & "',flag2='" & Trim(COL2) & "',flag3='" & Trim(COL3) & "',flag4='" & Trim(col4) & "' " _
'                           & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Trim(colCodMae2) & "'"
'                     cn.Execute Sqlquery, Options:=rdExecDirect
'                  End If
'               Case "014"
'                  'TIPO DE pAGOS
'                  If ColUpRow = "I" Then
'                    Sqlquery = " SET DATEFORMAT " & Coneccion.FormatFechaSql & " INSERT INTO maestros_2 " _
'                    & " VALUES ('" & wcia & cod & "','" & Trim(colCodMae) & "','" & Trim(COL1) & "'," _
'                    & "'" & Trim(COL3) & "','" & Trim(COL2) & "',NULL,NULL,NULL,'','" & wuser & "',NULL," & FechaSys & ",NULL,NULL,NULL) "
'                    cn.Execute Sqlquery, Options:=rdExecDirect
'                  Else
'                    Sqlquery = "UPDATE maestros_2 SET descrip = '" & Trim(COL1) & "',FLAG1='" & Trim(COL3) & "',flag2='" & Trim(COL2) & "' " _
'                    & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Trim(colCodMae2) & "'"
'                    cn.Execute Sqlquery, Options:=rdExecDirect
'                  End If
'               Case Else
'                  If COL2.Visible = False Then
'                     If ColUpRow = "I" Then
'                       Sqlquery = " INSERT INTO maestros_2 " _
'                         & " VALUES ('" & wcia & cod & "','" & Trim(colCodMae) & "','" & Trim(COL1) & "'," _
'                         & "'" & Trim(COL2) & "','" & Trim(COL3) & "',NULL,NULL,NULL,'','" & wuser & "',NULL," & FechaSys & ",NULL,NULL,NULL,NULL,'','','','','S','','" & wcia & "')"
'
'                       cn.Execute Sqlquery, Options:=rdExecDirect
'                     Else
'                        Sqlquery = "UPDATE maestros_2 SET descrip = '" & Trim(COL1) & "',flag1='" & Trim(COL2) & "',flag2='" & Trim(COL3) & "' " _
'                          & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Trim(colCodMae2) & "'"
'                        cn.Execute Sqlquery, Options:=rdExecDirect
'                     End If
'                  Else
'                    If ColUpRow = "I" Then
'
'                       Sqlquery = " INSERT INTO maestros_2 (ciamaestro,cod_maestro2,descrip,flag1,flag2,flag3,flag4,flag5,status,user_crea,user_modi,fec_crea,fec_modi,flag6,flag7,codsunat)" _
'                         & " VALUES ('" & wcia & cod & "','" & Trim(colCodMae) & "','" & Trim(COL1) & "','" & Trim(COL2) & "','" & Trim(COL3) & "','" & Trim(col4) & "','" & Trim(COL5) & "','" & Trim(COL6) _
'                         & "','','" & wuser & "',NULL," & FechaSys & ",NULL,'" & Trim(COL7) & "','" & Trim(COL8) & "','" & Trim(COL9) & "')"
'                       cn.Execute Sqlquery, Options:=rdExecDirect
'                     Else
'                        Sqlquery = "UPDATE maestros_2 SET descrip = '" & Trim(COL1) & "',flag1='" & Trim(COL2) & "',FLAG2='" & Trim(COL3) & "',FLAG3='" & Trim(col4) & "',flag4='" & Trim(COL5) & "',FLAG5='" & Trim(COL6) _
'                         & "',flag6='" & Trim(COL7) & "',flag7='" & Trim(COL8) & "',codsunat='" & Trim(COL9) & "' WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Trim(colCodMae2) & "'"
'                        cn.Execute Sqlquery, Options:=rdExecDirect
'                     End If
'                  End If
'
'               End Select
'
'             End Select
'
'        ColUpRow = ""
'        colCodMae2 = colCodMae
'       End If
'       datmae.Recordset.MoveNext
'       Loop
' UpRows = datmae.Recordset.RecordCount
' cntelim = 0
'
'       '*****************************************GRID 2 */*********************************
'         If grdmae2.Visible = True Then Grabar_Grid2
'
'
'       '*****************************************GRID 3 */*********************************
'         If grdmae3.Visible = True Then Grabar_Grid3
'
'    Sqlquery = wFinTrans
'    cn.Execute Sqlquery, Options:=rdExecDirect
'    Screen.MousePointer = vbDefault
'
'    If grdmae.Visible = True And grdmae2.Visible = False Then Unload Me: Exit Sub
'    If grdmae.Visible = True And grdmae2.Visible = True Then frmtmpmae.lstvmae_DblClick
'
'Exit Sub
''Err_Grabar:
'
'     'Call UserSqlMsgGrabar("Maestros", 0, "Error")
'     'Exit Sub
'End Sub
'Sub Grabar_Grid2()
'Screen.MousePointer = 11
'
'   Modificado_Grid2 = False
'     If eliminar2 = True And cntelim2 > 0 Then
'        If UBound(Reciclaje2) > 0 Then
'          For I = 1 To UBound(Reciclaje2)
'            Select Case TIPO_MAE$
'             Case "M"
'               Sqlquery = "DELETE maestros_3 " _
'                    & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Right(grdmae2.Caption, 2) & "' and cod_maestro2='" & Trim(Reciclaje2(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'             Case "C"
'               Sqlquery = "DELETE carlinnat2 " _
'                   & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Right(grdmae2.Caption, 3) & "' and cod_nat='" & Trim(Reciclaje2(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'            End Select
'          Next I
'        End If
'      End If
'      datmae2.Refresh
'       If Not datmae2.Recordset.BOF Then datmae2.Recordset.MoveFirst
'       Do While Not datmae2.Recordset.EOF
'        If Trim(ColUpRow2) <> "" Then
'            Select Case TIPO_MAE$
'             Case "M"
'               'AUTOS
'               If ColUpRow2 = "I" Then
'
'                 Sqlquery = " INSERT INTO maestros_3 (ciamaestro,cod_maestro2,cod_maestro3,descrip,flag1,flag2,status,user_crea,user_modi,fec_crea,fec_modi,flag3)" _
'                      & " VALUES ('" & wcia & cod & "','" & Right(grdmae2.Caption, 2) & "','" & Trim(colCodMae21) & "','" & Trim(COL21) & "','" _
'                      & Trim(grdmae2.Columns(4)) & "','" & Trim(grdmae2.Columns(5)) & " ','','" & wuser & "',NULL," & FechaSys & ",NULL,'" & Trim(grdmae2.Columns(6)) & "') "
'                 cn.Execute Sqlquery, Options:=rdExecDirect
'               Else
'                  Sqlquery = "UPDATE maestros_3 SET descrip = '" & Trim(COL21) & "',FLAG1='" & Trim(grdmae2.Columns(4)) & "', FLAG2='" & Trim(grdmae2.Columns(5)) & "', FLAG3='" & Trim(grdmae2.Columns(6)) & _
'                  "' WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Right(grdmae2.Caption, 2) & "' and cod_maestro3='" & Trim(colCodMae21) & "'"
'                  cn.Execute Sqlquery, Options:=rdExecDirect
'               End If
'
'             Case "C"
'             'CARACT
'                If ColUpRow2 = "I" Then
'                  Sqlquery = " INSERT INTO carlinnat2 " _
'                       & " VALUES ('" & wcia & "','" & Right(grdmae2.Caption, 3) & "','" & Trim(colCodMae21) & "','" & Trim(COL21) & "', " _
'                       & " '','" & wuser & "',NULL,DEFAULT,NULL) "
'                  cn.Execute Sqlquery, Options:=rdExecDirect
'                Else
'                  Sqlquery = "UPDATE carlinnat2 SET descrip = '" & Trim(COL21) & "' " _
'                       & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Right(grdmae2.Caption, 3) & "' and cod_nat='" & colCodMae21 & "'"
'                  cn.Execute Sqlquery, Options:=rdExecDirect
'                End If
'            End Select
'
'         ColUpRow2 = ""
'         colCodMae22 = colCodMae21
'        End If
'        datmae2.Recordset.MoveNext
'      Loop
'
'UpRows2 = datmae2.Recordset.RecordCount
'cntelim2 = 0
'
'Screen.MousePointer = vbDefault
'
'
'End Sub
'Sub Grabar_Grid3()
'
'     Modificado_Grid3 = False
'      If eliminar3 = True And cntelim3 > 0 Then
'        If UBound(Reciclaje3) > 0 Then
'          For I = 1 To UBound(Reciclaje3)
'              Sqlquery = "DELETE carlinnat3 " _
'                  & " WHERE cod_cia = '" & wcia & "' " _
'                  & " AND cod_linea='" & Right(grdmae2.Caption, 3) & "' and cod_nat='" & Right(grdmae3.Caption, 3) & "' " _
'                  & " AND cod_car='" & Reciclaje3(cntelim3) & "'"
'              cn.Execute Sqlquery, Options:=rdExecDirect
'          Next I
'        End If
'      End If
'
'       If Not datmae3.Recordset.BOF Then datmae3.Recordset.MoveFirst
'       Do While Not datmae3.Recordset.EOF
'        If Trim(ColUpRow3) <> "" Then
'            If ColUpRow3 = "I" Then
'               Sqlquery = " INSERT INTO carlinnat3 " _
'                   & " VALUES ('" & wcia & "','" & Right(grdmae2.Caption, 3) & "','" & Right(grdmae3.Caption, 3) & "','" & Trim(colCodMae31) & "', " _
'                   & " '" & Trim(COL31) & "', " _
'                   & " '','" & wuser & "',NULL,DEFAULT,NULL) "
'               cn.Execute Sqlquery
'            Else
'               Sqlquery = "UPDATE carlinnat3 SET descrip = '" & Trim(COL31) & "' " _
'                    & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Right(grdmae2.Caption, 3) & "' and cod_nat='" & Right(grdmae3.Caption, 3) & "' " _
'                    & " AND cod_car='" & colCodMae31 & "'"
'               cn.Execute Sqlquery
'            End If
'            ColUpRow3 = ""
'         colCodMae32 = colCodMae31
'        End If
'        datmae3.Recordset.MoveNext
'      Loop
'
'       '*****************************************GRID 2 */*********************************
'UpRows3 = datmae3.Recordset.RecordCount
'cntelim3 = 0
'Screen.MousePointer = vbDefault
'
'End Sub
'Private Sub grdmae_DblClick()
'
'vNroreg3 = 0
'
'COL1.Caption = frmtmpmae.lstvmae.SelectedItem.SubItems(1) & " " & colCodMae
'grdmae2.Caption = frmtmpmae.lstvmae.SelectedItem.SubItems(1) & " " & COL1 & " " & colCodMae
'
'If Trim(colCodMae2) <> "" Then
'    grdmae.Height = 3600
'    grdmae2.Enabled = True
'    Screen.MousePointer = 11
'    Modificado_Grid2 = False
'    Sqlquery = "SELECT descrip,cod_maestro3,flag1,flag2,flag3 FROM maestros_3 " _
'    & " where status='' and ciamaestro= '" & wcia + cod$ & "' AND CAST(cod_maestro2 AS INTEGER)='" & Trim(colCodMae) & "' ORDER BY cod_maestro3"
'
'     cn.CursorLocation = adUseClient
'     Set Adosql = New ADODB.Recordset
'     Set Adosql = cn.Execute(Sqlquery)
'            Sqlquery = "DELETE * FROM TmpMae02"
'            datmae2.Database.Execute Sqlquery
'        If Adosql.RecordCount > 0 Then
'
'                Do While Not Adosql.EOF
'                  datmae2.Recordset.AddNew
'                  datmae2.Recordset!descrip = Adosql!descrip
'                  datmae2.Recordset!COD_MAE = Adosql!cod_maestro3
'                  datmae2.Recordset!cod_mae2 = Adosql!cod_maestro3
'                  datmae2.Recordset!flag1 = Adosql!flag1
'                  datmae2.Recordset!flag2 = Adosql!flag2
'                  datmae2.Recordset!flag3 = Adosql!flag3
'                  datmae2.Recordset.Update
'                  vNroreg3 = Val(Adosql!cod_maestro3)
'                  Adosql.MoveNext
'                Loop
'                datmae2.Refresh
'                grdmae2.Col = 0
'                grdmae2.Row = 0
'                If grdmae2.Visible = False Then grdmae2.Visible = True
'
'                grdmae2.SetFocus
'                Screen.MousePointer = 0
'                Exit Sub
'            Else
'                datmae2.Refresh
'                grdmae2.Visible = True
'                grdmae2.SetFocus
'            End If
'
'    Else
'            'Modificado_Grid2 = False
'            grdmae2.Visible = False
'            grdmae.Height = 7200
'    End If
'
''End Select
''grdmae2.BOR
'Screen.MousePointer = 0
'End Sub
'
'Private Sub grdmae_Error(ByVal DataError As Integer, Response As Integer)
'Response = 0
'datmae.Refresh
'End Sub
'
'Private Sub grdmae_KeyPress(KeyAscii As Integer)
'        If cod$ = "005" And UCase(Trim(COL1)) = "CONTADO" Then
'            KeyAscii = 0
'        End If
'
'Select Case grdmae.Col
'Case 0
'
'End Select
'End Sub
'
'Private Sub grdmae_OnAddNew()
''Select Case TIPO_MAE$
''    Case "T", "C": colCodMae = Format(datmae.Recordset.RecordCount + 2, "000")
''    Case Else: colCodMae = Format(datmae.Recordset.RecordCount + 1, "00")
''End Select
'vNroreg2 = vNroreg2 + 1
'colCodMae = Format(vNroreg2, "00")
''Select Case TIPO_MAE$
''    Case "T", "C": colCodMae = Format(datmae.Recordset.RecordCount + 2, "000")
''    Case Else: colCodMae = Generar_Correlativo(datmae)
''End Select
'
'
'grdmae2.Visible = False
'grdmae.Height = 7200
'End Sub
'
'Private Sub grdmae2_AfterColEdit(ByVal ColIndex As Integer)
'If colCodMae21 <> colCodMae22 Then ColUpRow2 = "I" Else ColUpRow2 = "U"
'
'End Sub
'
'Private Sub grdmae2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'grdmae.Enabled = False
'End Sub
'
'Private Sub grdmae2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'grdmae.Enabled = False
'End Sub
'
'Private Sub grdmae2_BeforeDelete(Cancel As Integer)
'If eliminar2 = True Then
'   grdmae.Enabled = False
'   UpRows2 = datmae2.Recordset.RecordCount
'   If UpRows2 = 0 Then eliminar2 = False
'    If datmae2.Recordset.AbsolutePosition + 1 <= UpRows2 Then
'     cntelim2 = cntelim2 + 1
'     ReDim Preserve Reciclaje2(cntelim2)
'     Reciclaje2(cntelim2) = datmae2.Recordset!COD_MAE
'     If UpRows2 > 0 Then UpRows2 = UpRows2 - 1
'    End If
'    If datmae2.Recordset.AbsolutePosition = UpRows2 Then
'     cntelim2 = cntelim2 + 1
'     ReDim Preserve Reciclaje2(cntelim2)
'     Reciclaje2(cntelim2) = datmae2.Recordset!COD_MAE
''     If UpRows2 > 0 Then UpRows2 = UpRows2 - 1
'    End If
'End If
'End Sub
'
'
'
'Private Sub grdmae2_KeyPress(KeyAscii As Integer)
'Modificado_Grid2 = True
'Select Case grdmae2.Col
'    Case 0: If Len(grdmae2.Columns(0)) > 60 And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
'    Case 4: If Len(grdmae2.Columns(4)) > 4 And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
'    Case 5: If Len(grdmae2.Columns(5)) > 1 And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
'    Case 6: If Len(grdmae2.Columns(6)) > 0 And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
'End Select
'End Sub
'
'Private Sub grdmae2_OnAddNew()
''Select Case TIPO_MAE$
''             Case "C": colCodMae21 = Format(datmae2.Recordset.RecordCount + 1, "000")
''             Case Else: colCodMae21 = Format(datmae2.Recordset.RecordCount + 1, "00")
''End Select
'
'vNroreg3 = vNroreg3 + 1
'colCodMae21 = Format(vNroreg3, "00")
'
'End Sub
'
'Private Sub grdmae3_AfterColEdit(ByVal ColIndex As Integer)
'If colCodMae31 <> colCodMae32 Then ColUpRow3 = "I" Else ColUpRow3 = "U"
'End Sub
'
'Private Sub grdmae3_BeforeDelete(Cancel As Integer)
' If eliminar3 = True Then
'   UpRows3 = datmae3.Recordset.RecordCount
'   If UpRows3 = 0 Then eliminar3 = False
'    If datmae3.Recordset.AbsolutePosition + 1 <= UpRows3 Then
'     cntelim3 = cntelim3 + 1
'     ReDim Preserve Reciclaje3(cntelim3)
'     Reciclaje3(cntelim3) = datmae3.Recordset!COD_MAE
'     If UpRows3 > 0 Then UpRows3 = UpRows3 - 1
'    End If
' End If
'
'End Sub
'
'Private Sub grdmae3_KeyPress(KeyAscii As Integer)
'Modificado_Grid3 = True
'End Sub
'
'Private Sub grdmae3_OnAddNew()
'Select Case TIPO_MAE$
'             Case "C": colCodMae31 = Format(datmae3.Recordset.RecordCount + 1, "000")
'
'End Select
'End Sub
'
'Private Sub lstsino_Click()
'COL3 = lstsino.ItemData(lstsino.ListIndex)
'lstsino.Visible = False
'If colCodMae <> colCodMae2 Then ColUpRow = "I" Else ColUpRow = "U"
'
'End Sub
'
'Private Sub lstsino_LostFocus()
'lstsino.Visible = False
'End Sub
'
'Sub Grabar_Grid2_DblClic()
'
'     If grdmae2.Visible = True Then
'
'        If Not datmae2.Recordset.BOF Then datmae2.Recordset.MoveFirst
'
'        Do While Not datmae2.Recordset.EOF
'         If ColUpRow2 <> "" Then
'           If Trim(COL21) = "" Or IsNull(COL21) Then
'             MsgBox "Debe haber una Descripción", vbCritical
'             grdmae2.Col = 0
'             grdmae2.SetFocus
'             Exit Sub
'           End If
'         End If
'        datmae2.Recordset.MoveNext
'        Loop
'
'     End If
'
''On Error GoTo Err_Grabar
'
'  Sqlquery = "BEGIN TRANSACTION"
'  cn.Execute Sqlquery, Options:=rdExecDirect
'
'
'Screen.MousePointer = 11
'
'   Modificado_Grid2 = False
'     If eliminar2 = True And cntelim2 > 0 Then
'        If UBound(Reciclaje2) > 0 Then
'          For I = 1 To UBound(Reciclaje2)
'            Select Case TIPO_MAE$
'             Case "M"
'               Sqlquery = "DELETE maestros_3 " _
'                    & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Right(grdmae2.Caption, 2) & "' and cod_maestro2='" & Trim(Reciclaje2(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'             Case "C"
'               Sqlquery = "DELETE carlinnat2 " _
'                   & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Right(grdmae2.Caption, 2) & "' and cod_nat='" & Trim(Reciclaje2(I)) & "'"
'               cn.Execute Sqlquery, Options:=rdExecDirect
'            End Select
'
'          Next I
'        End If
'      End If
'
'       If Not datmae2.Recordset.BOF Then datmae2.Recordset.MoveFirst
'       Do While Not datmae2.Recordset.EOF
'        If Trim(ColUpRow2) <> "" Then
'            Select Case TIPO_MAE$
'             Case "M"
'               'AUTOS
'               If ColUpRow2 = "I" Then
'                 Sqlquery = " INSERT INTO maestros_3 " _
'                      & " VALUES ('" & wcia & cod & "','" & Right(grdmae2.Caption, 2) & "','" & Trim(colCodMae21) & "','" & Trim(COL21) & "', " _
'                      & " NULL,NULL,'','" & wuser & "',NULL," & FechaSys & ",NULL,'1') "
'                 cn.Execute Sqlquery, Options:=rdExecDirect
'               Else
'                  Sqlquery = "UPDATE maestros_3 SET descrip = '" & Trim(COL21) & "' " _
'                      & " WHERE ciamaestro = '" & wcia & cod & "' and cod_maestro2='" & Right(grdmae2.Caption, 2) & "' and cod_maestro3='" & Trim(colCodMae21) & "'"
'                  cn.Execute Sqlquery, Options:=rdExecDirect
'               End If
'
'             Case "C"
'             'CARACT
'                If ColUpRow2 = "I" Then
'                  Sqlquery = " INSERT INTO carlinnat2 " _
'                       & " VALUES ('" & wcia & "','" & Right(grdmae2.Caption, 3) & "','" & Trim(colCodMae21) & "','" & Trim(COL21) & "', " _
'                       & " '','" & wuser & "',NULL,DEFAULT,NULL) "
'                  cn.Execute Sqlquery, Options:=rdExecDirect
'                Else
'                  Sqlquery = "UPDATE carlinnat2 SET descrip = '" & Trim(COL21) & "' " _
'                       & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Right(grdmae2.Caption, 3) & "' and cod_nat='" & colCodMae21 & "'"
'                  cn.Execute Sqlquery, Options:=rdExecDirect
'                End If
'            End Select
'
'         ColUpRow2 = ""
'         colCodMae22 = colCodMae21
'        End If
'        datmae2.Recordset.MoveNext
'      Loop
'
'UpRows2 = datmae2.Recordset.RecordCount
'cntelim2 = 0
'
'    Sqlquery = "COMMIT TRANSACTION"
'    cn.Execute Sqlquery, Options:=rdExecDirect
'    Screen.MousePointer = vbDefault
'
'Exit Sub
''Err_Grabar:
'
''     Call UserSqlMsgGrabar("Maestros", 0, "Error")
' '    Exit Sub
'
'End Sub
'Sub Grabar_Grid3_DblClic()
'
'     If grdmae3.Visible = True Then
'
'      If Not datmae3.Recordset.BOF Then datmae3.Recordset.MoveFirst
'
'        Do While Not datmae3.Recordset.EOF
'         If ColUpRow3 <> "" Then
'           If Trim(COL31) = "" Or IsNull(COL31) Then
'             MsgBox "Debe haber una Descripción", vbCritical
'             grdmae3.Col = 0
'             grdmae3.SetFocus
'             Exit Sub
'           End If
'         End If
'         datmae3.Recordset.MoveNext
'        Loop
'
'     End If
'
'On Error GoTo Err_Grabar
'
'  Sqlquery = "BEGIN TRANSACTION"
'  cn.Execute Sqlquery, Options:=rdExecDirect
'
'     Modificado_Grid3 = False
'      If eliminar3 = True And cntelim3 > 0 Then
'        If UBound(Reciclaje3) > 0 Then
'          For I = 1 To UBound(Reciclaje3)
'              Sqlquery = "DELETE carlinnat3 " _
'                  & " WHERE cod_cia = '" & wcia & "' " _
'                  & " AND cod_linea='" & Right(grdmae2.Caption, 3) & "' and cod_nat='" & Right(grdmae3.Caption, 3) & "' " _
'                  & " AND cod_car='" & colCodMae31 & "'"
'              cn.Execute Sqlquery, Options:=rdExecDirect
'          Next I
'        End If
'      End If
'
'       If Not datmae3.Recordset.BOF Then datmae3.Recordset.MoveFirst
'       Do While Not datmae3.Recordset.EOF
'        If Trim(ColUpRow3) <> "" Then
'            If ColUpRow3 = "I" Then
'               Sqlquery = " INSERT INTO carlinnat3 " _
'                   & " VALUES ('" & wcia & "','" & Right(grdmae2.Caption, 3) & "','" & Right(grdmae3.Caption, 3) & "','" & Trim(colCodMae31) & "', " _
'                   & " '" & Trim(COL31) & "', " _
'                   & " '','" & wuser & "',NULL,DEFAULT,NULL) "
'               cn.Execute Sqlquery
'            Else
'               Sqlquery = "UPDATE carlinnat3 SET descrip = '" & Trim(COL31) & "' " _
'                    & " WHERE cod_cia = '" & wcia & "' AND cod_linea='" & Right(grdmae2.Caption, 3) & "' and cod_nat='" & Right(grdmae3.Caption, 3) & "' " _
'                    & " AND cod_car='" & colCodMae31 & "'"
'               cn.Execute Sqlquery
'            End If
'            ColUpRow3 = ""
'         colCodMae32 = colCodMae31
'        End If
'        datmae3.Recordset.MoveNext
'      Loop
'
''*****************************************GRID 2 */*********************************
'UpRows3 = datmae3.Recordset.RecordCount
'cntelim3 = 0
'
'
'    Sqlquery = "COMMIT TRANSACTION"
'    cn.Execute Sqlquery, Options:=rdExecDirect
'    Screen.MousePointer = vbDefault
'
'Exit Sub
'Err_Grabar:
''     Call UserSqlMsgGrabar("Maestros", 0, "Error")
'     Exit Sub
'End Sub
'
'Function Generar_Correlativo(pdata As Data) As String
'Dim rs2 As DAO.Recordset
'Dim codigo As String
'Dim bd As DAO.Database
'    If pdata.Recordset.RecordCount > 0 Then
'        Set bd = OpenDatabase(nom_BD, , True)
'        Set rs2 = bd.OpenRecordset("select max(cod_mae2) as nro from TmpMae01")
'        If rs2.RecordCount > 0 Then
'            If Not IsNull(rs2!NRO) Then
'                Generar_Correlativo = Format(Val(rs2!NRO) + 1, "00")
'            Else
'                Generar_Correlativo = Format(datmae.Recordset.RecordCount + 1, "00")
'            End If
'        End If
'    Else
'        Generar_Correlativo = Format(datmae.Recordset.RecordCount + 1, "00")
'    End If
'    If rs.State = 1 Then rs.Close
'    'Set rs = Nothing
'    Set rs2 = Nothing
'End Function
'
'
