VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmgrdobra 
   BackColor       =   &H80000001&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Obras"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "Frmgrdobra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      Begin VB.ComboBox Cmbcia 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compañia"
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
         TabIndex        =   2
         Top             =   120
         Width           =   840
      End
   End
   Begin MSDataGridLib.DataGrid Dgdobra 
      Bindings        =   "Frmgrdobra.frx":030A
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9551
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cod_obra"
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
         DataField       =   "descrip"
         Caption         =   "Nombre de la Obra"
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
         BeginProperty Column00 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   5969.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adoobra 
      Height          =   330
      Left            =   1200
      Top             =   3960
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
End
Attribute VB_Name = "Frmgrdobra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmbcia_Click()
wcia = Trim(Right("00" & Cmbcia.ItemData(Cmbcia.ListIndex), 2))
Procesa_Obras
End Sub

Private Sub Dgdobra_DblClick()
If wobra = True Then
   If Dgdobra.Row < 0 Then Exit Sub
   Call Frmobras.Carga_Obra(Cmbcia.Text, Dgdobra.Columns(0))
Else
   Select Case NameForm
        Case Is = "Frmgrdtareo"
             Frmgrdtareo.Txtcodobra = Trim(Dgdobra.Columns(0).Text)
             Frmgrdtareo.Lblobra = Trim(Dgdobra.Columns(1).Text)
        Case Is = "Frmtareo"
             Frmtareo.Txtcodobra = Trim(Dgdobra.Columns(0).Text)
             Frmtareo.Lblobra = Trim(Dgdobra.Columns(1).Text)
        Case Is = "Frmpersona"
             Frmpersona.Txtcodobra.Text = Trim(Dgdobra.Columns(0).Text)
             Frmpersona.Lbldesobra.Caption = Trim(Dgdobra.Columns(1).Text)
        Case Is = "FrmCabezaBol"
             FrmCabezaBol.Txtcodobra.Text = Trim(Dgdobra.Columns(0).Text)
             FrmCabezaBol.Lblobra.Caption = Trim(Dgdobra.Columns(1).Text)
   End Select
   NameForm = ""
   Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Width = 7890
Me.Height = 6630
Call rCarCbo(Cmbcia, Carga_Cia, "C", "00")
Call rUbiIndCmbBox(Cmbcia, wcia, "00")
Procesa_Obras
End Sub
Public Sub Procesa_Obras()
SQL$ = "SELECT * from plaobras where cod_cia='" & wcia & "' and status<>'*'"
cn.CursorLocation = adUseClient
Set Adoobra.Recordset = cn.Execute(SQL$, 64)
If Adoobra.Recordset.RecordCount > 0 Then Adoobra.Recordset.MoveFirst
Dgdobra.Refresh
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
wobra = True
End Sub
