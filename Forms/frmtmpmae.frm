VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmtmpmae 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Tablas Maestras"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmtmpmae.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel PanelCol 
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   5530
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelInner      =   2
      Begin VB.Data dat 
         Caption         =   "Dta"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Aceptar"
      End
      Begin VB.CheckBox ChkDet1 
         BackColor       =   &H00800000&
         Caption         =   "Detalle 1"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox ChkDet2 
         BackColor       =   &H00800000&
         Caption         =   "Detalle 2"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Defina el titulo   de las columnas y el nivel de detalle   de la   descripción "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   3855
      End
   End
   Begin ComctlLib.ListView lstvmae 
      Height          =   6075
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10716
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cod.Maestro"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Tabla Maestra"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ComboBox cmbcia 
      Enabled         =   0   'False
      ForeColor       =   &H80000012&
      Height          =   315
      ItemData        =   "frmtmpmae.frx":030A
      Left            =   1290
      List            =   "frmtmpmae.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   3405
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   765
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   6480
      Width           =   4695
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2. Hacer Doble Click para mostrar los registros de la tabla seleccionada."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   4050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Seleccione la Compañía a la cual pertenece."
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   0
         Width           =   3375
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmtmpmae.frx":030E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label49 
      Caption         =   "Compañía"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmtmpmae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
''Dim Rs As rdoResultset
'Dim itmX As ComctlLib.ListItem
'Public COD_MAE As String, nom_mae As String, TIPO_MAE As String
'Private StateTool(3) As String
'Dim windex As Integer
''As ListItem
'Dim I As Integer
'Public NRO_GRIDS As Integer
'Public NUEVO2 As Boolean
'Public codMaestro As String
'Dim Sqlquery As String
'Dim AdoMaestro  As ADODB.Recordset
'
'Private Sub Cmbcia_Click()
'
'
'If cmbcia <> "" Then
'
' Dim AdoMaestro As New ADODB.Recordset
'
' wcia = Format(cmbcia.ItemData(cmbcia.ListIndex), "00")
'
'    Sqlquery = " SELECT id_maestro,descrip,sistema,ciamaestro,cod_cia,status,user_crea FROM maestros " _
'    & " where status<>'*' "
'    Screen.MousePointer = vbHourglass
'    ''''DEBUG.PRINT SqlQuery
'    cn.CursorLocation = adUseClient
'    Set AdoMaestro = New ADODB.Recordset
'    Set AdoMaestro = cn.Execute(Sqlquery)
'
'    lstvmae.ListItems.Clear
'    If AdoMaestro.EOF Then
'        Screen.MousePointer = vbDefault
'        MsgBox "No hay ninguna Tabla", vbExclamation
'    Else
'        I = 1
'        Do While Not AdoMaestro.EOF
'            If InStr(1, AdoMaestro!Sistema, "4") > 0 Then
'                Set itmX = lstvmae.ListItems.Add()
'                itmX.Icon = 1
'                itmX.SmallIcon = 1
'                itmX.Text = Trim(AdoMaestro!id_maestro)
'                itmX.SubItems(1) = Trim(AdoMaestro!descrip)
'                itmX.SubItems(2) = "M"
'                I = I + 1
'            End If
'            AdoMaestro.MoveNext
'        Loop
'    End If
'    Screen.MousePointer = vbDefault
'End If
'End Sub
'
'
'
'Private Sub Form_Activate()
'On Error Resume Next
' cmbcia.ListIndex = Val(wcia) - 1
'End Sub
'
'Private Sub Form_Load()
'Me.Left = 0
'Me.Top = 0
'Me.Height = 7680
'Me.Width = 4830
'
'windex = Val(wcia) - 1
'Call rCarCbo(cmbcia, Carga_Cia, "C", "00")
'
'Dim tdf0 As TableDef
'dat.DatabaseName = nom_BD
'dat.Refresh
' Set tdf0 = dat.Database.CreateTableDef("TmpMae")
'    With tdf0
'        .Fields.Append .CreateField("nro", dbText)
'        .Fields(0).AllowZeroLength = True
'        .Fields.Append .CreateField("descrip", dbText)
'        .Fields(1).AllowZeroLength = True
'      dat.Database.TableDefs.Append tdf0
'   End With
'dat.RecordSource = "TmpMae"
'dat.Refresh
''Set rstmat = datint.Recordset
'Set tdf0 = Nothing
'
'Me.Show
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
''Set grd.DataSource = Nothing
''dat.RecordSource = ""
''dat.Refresh
'''rstmat.Close
''dat.Database.TableDefs.Delete "TmpMae"
''dat.Database.Close
'End Sub
'
'
'Private Sub grd_BeforeInsert(Cancel As Integer)
'If dat.Recordset.RecordCount = 3 Then
'  MsgBox "El Nro. Máximo de Registros es 3", vbInformation
'  Cancel = True
'End If
'End Sub
'
'Private Sub grd_OnAddNew()
'grd.Columns(0) = dat.Recordset.RecordCount + 1
'
'End Sub
'
'
'Private Sub lstvmae_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'lstvmae.SortKey = ColumnHeader.Index - 1
'' Establece Sorted a True para ordenar la lista.
'lstvmae.Sorted = True
'
'End Sub
'Public Sub lstvmae_DblClick()
'Unload frmtmpmae2
'NRO_GRIDS = 0
' If lstvmae.ListItems.count > 0 Then
'   If lstvmae.ListItems(lstvmae.SelectedItem.Index).Selected = True Then
'    COD_MAE = Trim(lstvmae.ListItems(lstvmae.SelectedItem.Index))
'    nom_mae = Trim(lstvmae.ListItems(lstvmae.SelectedItem.Index).SubItems(1))
'    TIPO_MAE = Trim(lstvmae.ListItems(lstvmae.SelectedItem.Index).SubItems(2))
'
'    NRO_GRIDS = 2
'
'  If dat.Recordset.RecordCount = 0 And NUEVO2 = True Then
'    MsgBox "Deberá Ingresar la Descripción de la(s) Columna(s) " & Chr(13) & "para su Detalle de la Nueva Tabla", vbInformation
'    dat.Refresh
'     If dat.Recordset.RecordCount > 0 Then
'        Sqlquery = "DELETE * FROM TmpMae"
'        dat.Database.Execute Sqlquery
'     End If
'
'      SSCommand1.Visible = True
'      PanelCol.Visible = True
'      grd.Visible = True
'      grd.Col = 1
'      grd.SetFocus
'     Exit Sub
'   End If
'     Load frmtmpmae2
'     frmtmpmae2.Show
'     frmtmpmae2.datmae.Refresh
'   End If
'End If
'End Sub
'Public Sub Inicio()
''Dim AdoMaestro As rdoResultset
'
'Dim Nuevo$
'Dim itmX As ComctlLib.ListItem
' Screen.MousePointer = 0
' Nuevo = InputBox("Ingrese Una Nueva Tabla Maestra")
' Do While Nuevo = ""
'    If Nuevo = "" Then Exit Sub
'    Nuevo = UCase(InputBox("Ingrese Una Nueva Tabla Maestra"))
' Loop
'
'        Set itmX = lstvmae.ListItems.Add()
'        itmX.SmallIcon = 1
'        itmX.Text = Format(lstvmae.ListItems.count, "000")
'        itmX.SubItems(1) = Trim(UCase(Nuevo$))
'        itmX.SubItems(2) = "M" & codMaestro
'        Sqlquery = "SELECT MAX(ID_MAESTRO) FROM MAESTROS WHERE cod_cia='" & wcia & "'"
'        ''''DEBUG.PRINT SqlQuery
'        Screen.MousePointer = 11
'        cn.CursorLocation = adUseClient
'        Set AdoMaestro = New ADODB.Recordset
'        Set AdoMaestro = cn.Execute(Sqlquery)
'        codMaestro = Format(AdoMaestro(0) + 1, "000")
'        Set AdoMaestro = Nothing
'
'        Sqlquery = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
'        Sqlquery = Sqlquery & " INSERT INTO maestros(cod_cia,id_maestro,ciamaestro,descrip,status,user_crea,FEC_CREA,sistema) VALUES('" _
'        & wcia & "', " & " '" & codMaestro & "','" & wcia & codMaestro & "', " & " '" & Nuevo & "','','" & wuser & "',GETDATE(),'3')"
'
'        cn.Execute Sqlquery, Options:=rdExecDirect
'
'
'        Screen.MousePointer = 0
'
'NUEVO2 = True
'Cmbcia_Click
'Set lstvmae.SelectedItem = lstvmae.ListItems(lstvmae.ListItems.count)
'
'End Sub
'
'
'Private Sub SSCommand1_Click()
'dat.Refresh
'If dat.Recordset.RecordCount > 0 Then
'    If ChkDet2.Value = 1 Then NRO_GRIDS = 2
'    Load frmtmpmae2
'    frmtmpmae2.Show
'    SSCommand1.Visible = False
'    grd.Visible = False
'    PanelCol.Visible = False
'Else
' MsgBox "Deberá Ingresar la Descripción de la(s) Columna(s) " & Chr(13) & "para su Detalle de la Nueva Tabla", vbInformation
'End If
'
'End Sub
'
'Public Sub Eliminar_Maestro()
'Dim Preg As Integer
'
'        If lstvmae.ListItems.count > 0 Then
'            If lstvmae.ListItems(lstvmae.SelectedItem.Index).Selected = True Then
'
'                Preg = MsgBox(" Seguro de Eliminar el Concepto? ", vbQuestion + vbYesNo + vbDefaultButton2)
'                If Preg = vbYes Then
'
'                    Screen.MousePointer = vbHourglass
'                    Sqlquery = "BEGIN TRANSACTION"
'                    cn.Execute Sqlquery
'                    COD_MAE = Trim(lstvmae.ListItems(lstvmae.SelectedItem.Index))
'
'                    Sqlquery = "UPDATE maestros SET STATUS='*' WHERE ciamaestro='" & wcia & COD_MAE & "' and status=''"
'                    ''DEBUG.PRINT SqlQuery
'                    cn.Execute Sqlquery
'
'                    Sqlquery = "UPDATE maestros_2 SET STATUS='*' WHERE ciamaestro='" & wcia & COD_MAE & "' and status=''"
'                    ''DEBUG.PRINT SqlQuery
'                    cn.Execute Sqlquery
'
'                    Sqlquery = "UPDATE maestros_3 SET STATUS='*' WHERE ciamaestro='" & wcia & COD_MAE & "' and status=''"
'                    ''DEBUG.PRINT SqlQuery
'                    cn.Execute Sqlquery
'
'                    NUEVO2 = False
'
'                    Sqlquery = "COMMIT TRANSACTION"
'                    cn.Execute Sqlquery
'                    Unload frmtmpmae2
'                    Cmbcia_Click
'
'                    Screen.MousePointer = vbDefault
'
'            End If
'        End If
'    End If
'End Sub
'
