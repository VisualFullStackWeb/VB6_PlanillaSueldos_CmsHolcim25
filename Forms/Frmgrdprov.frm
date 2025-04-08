VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmgrdprov 
   Caption         =   "Consulta de Proveedores"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   Icon            =   "Frmgrdprov.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8535
      Begin VB.TextBox Txtruc 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Txtrazsoc 
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox Txtcod 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2400
         TabIndex        =   6
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
   End
   Begin MSDataGridLib.DataGrid Dtgrdprov 
      Bindings        =   "Frmgrdprov.frx":030A
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483628
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cod_prov"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "razsoc"
         Caption         =   "Razon Social / Apellidos Nombres"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ruc"
         Caption         =   "RUC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "per_nat"
         Caption         =   "Persona"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4454.929
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adoprov 
      Height          =   330
      Left            =   3360
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "                                    Maestro de Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   8535
   End
End
Attribute VB_Name = "Frmgrdprov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dtgrdprov_DblClick()
'On Error GoTo Err:
Select Case MDIplared.ActiveForm.Name

        Case Is = "FrmExamenes"
            FrmExamenes.txtCodProv.Text = Trim(Dtgrdprov.Columns(0).Text)
            FrmExamenes.txtProveedor.Text = Trim(Dtgrdprov.Columns(1).Text)
            FrmExamenes.txtRucProveedor.Text = Trim(Dtgrdprov.Columns(2).Text)
        End Select
        Unload Me
Err:
End Sub

Private Sub Dtgrdprov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Dtgrdprov.VisibleRows > 0 Then Dtgrdprov_DblClick
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
Me.Width = 8670
Me.Height = 6315
Me.Txtrazsoc.SetFocus
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Sql$ = "EXECUTE sp_proveedores_cia '" & wcia & "'"
cn.CursorLocation = adUseClient
Set Adoprov.Recordset = cn.Execute(Sql, 64)
If Adoprov.Recordset.EOF Then MsgBox ("No Existen Proveedores Para la Compañia")
Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
End Sub

Private Sub Txtcod_GotFocus()
Txtruc = ""
Txtrazsoc = ""
End Sub

Private Sub Txtcod_KeyPress(KeyAscii As Integer)
Dim vgCodProv As String
'Txtcod.Text = Txtcod.Text + fc_ValNumeros(KeyAscii)

If KeyAscii = 13 And Txtcod.Text <> "" Then

'vgCodProv = Right("00000000" & Trim(Txtcod.Text), 8)
vgCodProv = Trim(Txtcod.Text)
Screen.MousePointer = vbHourglass
    Sql$ = "SP_Proveedores_Codigo '" & wcia & "','" & vgCodProv & "'"
    Set Adoprov.Recordset = cn.Execute(Sql$)
    Screen.MousePointer = vbDefault
    If Adoprov.Recordset.RecordCount = 0 Then
       Txtcod.Text = vgCodProv
       Dtgrdprov.Refresh
       MsgBox "El codigo del Proveedor No existe", vbExclamation, "Verifique"
    End If
    Txtcod.SelStart = 0
    Txtcod.SelLength = Len(Txtcod)
    Txtcod.SetFocus
    Exit Sub
End If
Screen.MousePointer = vbDefault
Exit Sub

End Sub

Private Sub Txtrazsoc_Change()
Sql$ = "EXECUTE SP_PROVEEDORES_CIA_RAZSOC  '" & wcia & "','" & Trim(Txtrazsoc.Text) & "'"
cn.CursorLocation = adUseClient
Set Adoprov.Recordset = cn.Execute(Sql)
End Sub

Private Sub Txtrazsoc_GotFocus()
Txtruc = ""
Txtcod = ""
End Sub

Private Sub Txtrazsoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then Me.Dtgrdprov.SetFocus
End Sub

Private Sub Txtruc_GotFocus()
Txtcod = ""
Txtrazsoc = ""
End Sub

Private Sub Txtruc_KeyPress(KeyAscii As Integer)
Txtruc.Text = Txtruc.Text + fc_ValNumeros(KeyAscii)

Dim sw As Boolean

If KeyAscii = 13 And Len(Txtruc.Text) > 7 Then
sw = False
Screen.MousePointer = vbHourglass

Sql$ = "SP_Proveedores_Ruc '" & wcia & "','" & Trim(Txtruc.Text) & "'"
    Set Adoprov.Recordset = cn.Execute(Sql$)
    Screen.MousePointer = vbDefault
    If Adoprov.Recordset.RecordCount = 0 Then
       Dtgrdprov.Refresh
       MsgBox "El Numero de RUC No existe", vbExclamation, "Verifique"
    End If
    Txtruc.SelStart = 0
    Txtruc.SelLength = Len(Txtruc)
    Txtruc.SetFocus
    Exit Sub
    
End If
Screen.MousePointer = vbDefault
Exit Sub

End Sub
