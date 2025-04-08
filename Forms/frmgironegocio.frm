VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form frmgironegocio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione el Giro de Negocio"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprocesa 
      Caption         =   "P r o c e s a r"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   5
      Top             =   720
      Width           =   7800
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdLib 
      Height          =   3465
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   6112
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin VB.TextBox txtbuscacampo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   5325
   End
   Begin VB.ComboBox cbo_campo 
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
      ItemData        =   "frmgironegocio.frx":0000
      Left            =   90
      List            =   "frmgironegocio.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2355
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   270
      Top             =   1710
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   2296
      Images          =   "frmgironegocio.frx":0056
      Version         =   131072
      KeyCount        =   2
      Keys            =   "SORTASCÿSORTDESC"
   End
   Begin vbalIml6.vbalImageList vbalImageList 
      Left            =   1575
      Top             =   1305
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   14924
      Images          =   "frmgironegocio.frx":096E
      Version         =   131072
      KeyCount        =   13
      Keys            =   "ÿÿÿÿÿÿÿÿÿÿÿÿ"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2520
      TabIndex        =   4
      Top             =   90
      Width           =   570
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Campo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   90
      Width           =   585
   End
End
Attribute VB_Name = "frmgironegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum Opcion
   GIROS_NEGOCIO = 1
   PROFESIONES = 2
End Enum
Private m_eOPCIONCARGA As Opcion


Public Property Let OpcionCarga(ByVal eOPCIONCARGA As Opcion)
    m_eOPCIONCARGA = eOPCIONCARGA
End Property

Private Sub setUpGrid()
   
   ' Set general options:
   With grdLib
      .HideGroupingBox = True
      .AllowGrouping = True
       
      .BackColor = RGB(255, 255, 245)
      '.AlternateRowBackColor = RGB(255, 255, 185)
      .GroupRowBackColor = RGB(180, 173, 176)
      .GroupingAreaBackColor = .BackColor
      .ForeColor = RGB(0, 0, 0)
      .GroupRowForeColor = .ForeColor
      .HighlightForeColor = vbWindowText
      .HighlightBackColor = RGB(180, 173, 176)
      .NoFocusHighlightBackColor = RGB(200, 200, 200)
      .SelectionAlphaBlend = True
      .SelectionOutline = True
      .DrawFocusRectangle = False
      .HighlightSelectedIcons = False
      .HotTrack = True
      .RowMode = True
      .MultiSelect = True
      .ImageList = vbalImageList
   
    If m_eOPCIONCARGA = GIROS_NEGOCIO Then
      ' Add the columns:
        .AddColumn "SECCION", "Seccion", , , 10, True, , , , , , CCLSortStringNoCase
        .AddColumn "DIVISION", "Division", , , 100, True, , , , , , CCLSortStringNoCase
        .AddColumn "GRUPO", "Grupo", , , 100, True, , , , , , CCLSortStringNoCase
        .AddColumn "CLASE", "Clase", ecgHdrTextALignLeft, , 100, True, , , , , , CCLSortStringNoCase
        .AddColumn "DESCRIP", "Descripcion", , , 350, True, , , , , , CCLSortStringNoCase
        .AddColumn "CODIGO", "Codigo", , , 50, True, , , , , , CCLSortStringNoCase
      Else
        .AddColumn "NIVEL1", "Nivel 1", , , 10, True, , , , , , CCLSortStringNoCase
        .AddColumn "NIVEL2", "Nivel2", , , 100, True, , , , , , CCLSortStringNoCase
        .AddColumn "NIVEL3", "Nivel 3", , , 100, True, , , , , , CCLSortStringNoCase
        .AddColumn "DESCRIP", "Descripcion", , , 350, True, , , , , , CCLSortStringNoCase
        .AddColumn "CODIGO", "Codigo", , , 50, True, , , , , , CCLSortStringNoCase
      End If
      
      .SetHeaders
      
      .HeaderImageList = ilsIcons
      
      .StretchLastColumnToFit = True
      
   End With
   
End Sub

Private Sub CargaData()

Dim sSQL As String, Generado As String
Dim CAMBIA As Boolean
Dim SECCION As String, SECCIONDESC As String, DIVISION As String, DIVISIONDESC As String, GRUPO As String, GRUPODESC As String, CLASE As String, CLASEDESC As String, SUBCLASE As String, SUBCLASEDESC As String
Dim NIVEL1 As String, NIVEL1DESC As String, NIVEL2 As String, NIVEL2DESC As String, NIVEL3 As String, NIVEL3DESC As String, NIVEL4 As String, NIVEL4DESC As String

Randomize
Generado = CStr(CLng(((99999 - 9999 + 1) * Rnd) + 9999))
If cbo_campo.Text = "" Then Exit Sub
Dim rs As ADODB.Recordset

If m_eOPCIONCARGA = GIROS_NEGOCIO Then
    If cbo_campo.Text = "(S) SECCION" Then
        grdLib.ColumnVisible("SECCION") = True
        grdLib.ColumnVisible("DIVISION") = True
        grdLib.ColumnVisible("GRUPO") = True
        grdLib.ColumnVisible("CLASE") = True
    ElseIf cbo_campo.Text = "(D) DIVISION" Then
        grdLib.ColumnVisible("SECCION") = False
        grdLib.ColumnVisible("DIVISION") = True
        grdLib.ColumnVisible("GRUPO") = True
        grdLib.ColumnVisible("CLASE") = True
    ElseIf cbo_campo.Text = "(G) GRUPO" Then
        grdLib.ColumnVisible("SECCION") = False
        grdLib.ColumnVisible("DIVISION") = False
        grdLib.ColumnVisible("GRUPO") = True
        grdLib.ColumnVisible("CLASE") = True
    ElseIf cbo_campo.Text = "(C) CLASE" Then
        grdLib.ColumnVisible("SECCION") = False
        grdLib.ColumnVisible("DIVISION") = False
        grdLib.ColumnVisible("GRUPO") = False
        grdLib.ColumnVisible("CLASE") = True
    ElseIf cbo_campo.Text = "(B) SUBCLASE" Then
        grdLib.ColumnVisible("SECCION") = False
        grdLib.ColumnVisible("DIVISION") = False
        grdLib.ColumnVisible("GRUPO") = False
        grdLib.ColumnVisible("CLASE") = False
    End If
Else
    If UCase(Trim(cbo_campo.Text)) = "(1) NIVEL" Then
        grdLib.ColumnVisible("NIVEL1") = False
        grdLib.ColumnVisible("NIVEL2") = False
        grdLib.ColumnVisible("NIVEL3") = False
        
'                grdLib.ColumnVisible("NIVEL1") = True
'        grdLib.ColumnVisible("NIVEL2") = True
'        grdLib.ColumnVisible("NIVEL3") = True
    ElseIf UCase(Trim(cbo_campo.Text)) = "(2) NIVEL" Then
        grdLib.ColumnVisible("NIVEL1") = False
        grdLib.ColumnVisible("NIVEL2") = True
        grdLib.ColumnVisible("NIVEL3") = True
    ElseIf UCase(Trim(cbo_campo.Text)) = "(3) NIVEL" Then
        grdLib.ColumnVisible("NIVEL1") = False
        grdLib.ColumnVisible("NIVEL2") = False
        grdLib.ColumnVisible("NIVEL3") = True
    ElseIf UCase(Trim(cbo_campo.Text)) = "(4) NIVEL " Then
        grdLib.ColumnVisible("NIVEL1") = False
        grdLib.ColumnVisible("NIVEL2") = False
        grdLib.ColumnVisible("NIVEL3") = False
    End If
End If

If m_eOPCIONCARGA = GIROS_NEGOCIO Then
    sSQL = "DELETE FROM TMPGIROS WHERE IDPROCESO=" & Generado
    cn.Execute sSQL
    
    sSQL = "EXEC sp_c_obtener_giros '" & Mid(Me.cbo_campo.Text, 2, 1) & "','%" & Trim(txtbuscacampo.Text) & "%',NULL," & Generado
    cn.Execute sSQL
    
    sSQL = "SELECT TIPO,DESCRIPCION,CODIGO FROM TMPGIROS WHERE IDPROCESO=" & Generado & " ORDER BY correla"
Else
    sSQL = "DELETE FROM TMPOCUPACIONES WHERE IDPROCESO=" & Generado
    cn.Execute sSQL
    
    sSQL = "EXEC sp_c_obtener_OCUPACIONES '" & Mid(cbo_campo.Text, 2, 1) & "','%" & Trim(txtbuscacampo.Text) & "%',NULL," & Generado
    cn.Execute sSQL
    
    sSQL = "SELECT TIPO,DESCRIPCION,CODIGO FROM TMPOCUPACIONES WHERE IDPROCESO=" & Generado & " ORDER BY Descripcion"
End If

Set rs = cn.Execute(sSQL)

If Not rs.EOF Then
With grdLib
    .Redraw = False
    
    If m_eOPCIONCARGA = GIROS_NEGOCIO Then
        grdLib.ColumnIsGrouped(1) = False
        grdLib.ColumnIsGrouped(2) = False
        grdLib.ColumnIsGrouped(3) = False
        grdLib.ColumnIsGrouped(4) = False
    Else
        grdLib.ColumnIsGrouped(1) = False
        grdLib.ColumnIsGrouped(2) = False
        grdLib.ColumnIsGrouped(3) = False
    End If
    
    .Clear
        Do While Not rs.EOF
            If m_eOPCIONCARGA = GIROS_NEGOCIO Then
                If rs!tipo = "S" Then SECCION = Trim(rs!codigo): SECCIONDESC = Trim(rs!Descripcion)
                If rs!tipo = "D" Then DIVISION = Trim(rs!codigo): DIVISIONDESC = Trim(rs!Descripcion)
                If rs!tipo = "G" Then GRUPO = Trim(rs!codigo): GRUPODESC = Trim(rs!Descripcion)
                If rs!tipo = "C" Then CLASE = Trim(rs!codigo): CLASEDESC = Trim(rs!Descripcion)
                If rs!tipo = "B" Then
                    If SUBCLASE <> Trim(rs!codigo) Then
                        .AddRow
                        .CellDetails .Rows, 1, SECCIONDESC
                        .CellDetails .Rows, 2, DIVISIONDESC
                        .CellDetails .Rows, 3, GRUPODESC
                        .CellDetails .Rows, 4, CLASEDESC
                        .CellDetails .Rows, 5, rs!Descripcion
                        .CellDetails .Rows, 6, rs!codigo
                    End If
                End If
            Else
                If rs!tipo = "1" Then NIVEL1 = Trim(rs!codigo): NIVEL1DESC = Trim(rs!Descripcion)
                If rs!tipo = "2" Then NIVEL2 = Trim(rs!codigo): NIVEL2DESC = Trim(rs!Descripcion)
                If rs!tipo = "3" Then NIVEL3 = Trim(rs!codigo): NIVEL3DESC = Trim(rs!Descripcion)
'                If rs!tipo = "4" Then
'                    .AddRow
'                    .CellDetails .Rows, 1, Trim(NIVEL1DESC)
'                    .CellDetails .Rows, 2, Trim(NIVEL2DESC)
'                    .CellDetails .Rows, 3, Trim(NIVEL3DESC)
'                    .CellDetails .Rows, 4, Trim(rs!Descripcion)
'                    .CellDetails .Rows, 5, Trim(rs!codigo)
'                End If
                If rs!tipo = "1" Then
                    .AddRow
                    .CellDetails .Rows, 1, Trim(NIVEL1DESC)
                    .CellDetails .Rows, 2, Trim(NIVEL2DESC)
                    .CellDetails .Rows, 3, Trim(NIVEL3DESC)
                    .CellDetails .Rows, 4, Trim(rs!Descripcion)
                    .CellDetails .Rows, 5, Trim(rs!codigo)
                End If
            End If
                    
                    
            rs.MoveNext
        Loop
End With

If m_eOPCIONCARGA = GIROS_NEGOCIO Then
    If cbo_campo.Text = "(S) SECCION" Then
        grdLib.ColumnIsGrouped(1) = True
        grdLib.ColumnIsGrouped(2) = True
        grdLib.ColumnIsGrouped(3) = True
        grdLib.ColumnIsGrouped(4) = True
        
    ElseIf cbo_campo.Text = "(D) DIVISION" Then
        grdLib.ColumnIsGrouped(2) = True
        grdLib.ColumnIsGrouped(3) = True
        grdLib.ColumnIsGrouped(4) = True
        
    ElseIf cbo_campo.Text = "(G) GRUPO" Then
        grdLib.ColumnIsGrouped(1) = False
        grdLib.ColumnIsGrouped(2) = False
        grdLib.ColumnIsGrouped(3) = True
        grdLib.ColumnIsGrouped(4) = True
        
    ElseIf cbo_campo.Text = "(C) CLASE" Then
        grdLib.ColumnIsGrouped(1) = False
        grdLib.ColumnIsGrouped(2) = False
        grdLib.ColumnIsGrouped(3) = False
        grdLib.ColumnIsGrouped(4) = True
    End If
Else
    If UCase(Trim(cbo_campo.Text)) = "(1) NIVEL" Then
'        grdLib.ColumnIsGrouped(1) = True
'        grdLib.ColumnIsGrouped(2) = True
'        grdLib.ColumnIsGrouped(3) = True
        
    ElseIf UCase(Trim(cbo_campo.Text)) = "(2) NIVEL" Then
        grdLib.ColumnIsGrouped(1) = False
        grdLib.ColumnIsGrouped(2) = True
        grdLib.ColumnIsGrouped(3) = True
        
    ElseIf UCase(Trim(cbo_campo.Text)) = "(3) NIVEL" Then
        grdLib.ColumnIsGrouped(1) = False
        grdLib.ColumnIsGrouped(2) = False
        grdLib.ColumnIsGrouped(3) = True
        
    End If
End If
    grdLib.Redraw = True
Else
    grdLib.Clear
    MsgBox "No Hay Informacion para Mostrar", vbExclamation, "Sistema de Planilla"
End If
Set rs = Nothing

End Sub

Private Sub cmdprocesa_Click()
    grdLib.Clear
    Screen.MousePointer = vbHourglass
    CargaData
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
setUpGrid

cbo_campo.Clear
If m_eOPCIONCARGA = PROFESIONES Then
    cbo_campo.AddItem "(1) Nivel "
'    cbo_campo.AddItem "(2) Nivel "
'    cbo_campo.AddItem "(3) Nivel "
'    cbo_campo.AddItem "(4) Nivel "
    cbo_campo.ListIndex = 0
    CargaData
Else
    cbo_campo.AddItem "(S) SECCION"
    cbo_campo.AddItem "(D) DIVISION"
    cbo_campo.AddItem "(G) GRUPO"
    cbo_campo.AddItem "(C) CLASE"
    cbo_campo.AddItem "(B) SUBCLASE"
End If

End Sub

Private Sub grdLib_DblClick(ByVal lRow As Long, ByVal lCol As Long)
If Not grdLib.RowIsGroup(lRow) Then
    If m_eOPCIONCARGA = GIROS_NEGOCIO Then
        Frmcia.txtgiro.Text = grdLib.CellText(lRow, 5)
        Frmcia.txtgiro.Tag = grdLib.CellText(lRow, 6)
    Else
        Frmpersona.txtprofesion.Text = grdLib.CellText(lRow, 4)
        Frmpersona.txtprofesion.Tag = grdLib.CellText(lRow, 5)
    End If
    Unload Me
End If
End Sub

Private Sub txtbuscacampo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmdprocesa_Click
End If
End Sub
