VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmAyuda 
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   Icon            =   "FrmAyuda.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmAyuda.frx":030A
   ScaleHeight     =   5250
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adopla 
      Height          =   330
      Left            =   600
      Top             =   2520
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8415
      Begin VB.TextBox TxtBusca 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   7935
      End
   End
   Begin MSDataGridLib.DataGrid Dgrdplanilla 
      Bindings        =   "FrmAyuda.frx":0614
      Height          =   4650
      Left            =   45
      TabIndex        =   0
      Top             =   600
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8202
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "descripcion"
         Caption         =   "Descripcion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6494.74
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Maestro de Personal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   8295
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Busqueda As String
Public Tipoinst As Integer
Public Regimen As Integer



Private Sub Dgrdplanilla_DblClick()
Select Case MDIplared.ActiveForm.Name
        Case Is = "Frmpersona"
            If Busqueda = "OCUPACION" Then
               Frmpersona.txtprofesion.Text = Trim(Dgrdplanilla.Columns(1).Text)
               Frmpersona.txtprofesion.Tag = Trim(Dgrdplanilla.Columns(0).Text)
            End If
            If Busqueda = "UNIVERSIDAD" Then
               Frmpersona.TxtNombreUni.Text = Trim(Dgrdplanilla.Columns(1).Text)
               Frmpersona.TxtNombreUni.Tag = Trim(Dgrdplanilla.Columns(0).Text)
            End If
            If Busqueda = "CARRERA" Then
               Frmpersona.TxtCarrera.Text = Trim(Dgrdplanilla.Columns(1).Text)
               Frmpersona.TxtCarrera.Tag = Trim(Dgrdplanilla.Columns(0).Text)
            End If
            If Busqueda = "CANTERAS" Then
               Frmpersona.LblCodCantera.Caption = Trim(Dgrdplanilla.Columns(0).Text)
               Frmpersona.LblCantera.Caption = Trim(Dgrdplanilla.Columns(1).Text)
            End If
End Select
Unload Me
End Sub

Private Sub Form_Load()
Select Case Busqueda
       Case "OCUPACION"
           Me.Caption = "OCUPACIONES"
End Select
Carga_Data
End Sub
Private Sub Carga_Data()
Dim Sql As String
Dim xTipo As Integer
xTipo = 0
Select Case Busqueda
       Case "OCUPACION": xTipo = 1
       Case "UNIVERSIDAD": xTipo = 2
       Case "CARRERA": xTipo = 3
       Case "CANTERAS": xTipo = 0
End Select

Screen.MousePointer = vbHourglass
Sql$ = "Usp_Pla_Ayuda '" & wcia & "', " & xTipo & ", '" & Trim(TxtBusca.Text) & "'," & Regimen & "," & Tipoinst & ""
Set Adopla.Recordset = cn.Execute(Sql$, 64)
Screen.MousePointer = vbDefault

cn.CursorLocation = adUseClient

Screen.MousePointer = vbDefault

End Sub

Private Sub TxtBusca_Change()
Carga_Data
End Sub
